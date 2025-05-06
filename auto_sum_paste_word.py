import pyperclip
import time
import re
import subprocess
import sys # To check platform

# --- Configuration ---
CHECK_INTERVAL_SECONDS = 1.0 # How often to check clipboard
WORD_APP_NAME = "Microsoft Word" # Check this in Activity Monitor if unsure
#DECIMAL_PLACES = 2 # How many decimal places for the sum string
OUTPUT_DECIMAL_SEPARATOR = ',' # Use comma for the final output string
OUTPUT_THOUSANDS_SEPARATOR = '.' # Use dot for the final output string


# --- Helper Functions ---

def run_applescript(script):
    """Executes an AppleScript command and returns its output."""
    if sys.platform != 'darwin':
        print("AppleScript execution requires macOS.")
        return None
    try:
        process = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            check=True, # Raise exception on non-zero exit code
            timeout=5 # Prevent hanging indefinitely
        )
        return process.stdout.strip()
    except subprocess.CalledProcessError as e:
        print(f"AppleScript Error: {e.stderr}")
        # Might need permission in Privacy & Security > Automation
        return None
    except subprocess.TimeoutExpired:
        print("AppleScript command timed out.")
        return None
    except FileNotFoundError:
        print("Error: 'osascript' command not found. Is this macOS?")
        return None
    except Exception as e:
        print(f"An unexpected error occurred running AppleScript: {e}")
        return None

def get_frontmost_app():
    """Gets the name of the frontmost application using AppleScript."""
    script = '''
        tell application "System Events"
            try
                name of first application process whose frontmost is true
            on error errMsg number errorNum
                 # Common error if screen saver is active or no app has focus
                 # print "AppleScript Error getting frontmost app: " & errMsg & " (" & errorNum & ")"
                 return ""
            end try
        end tell
    '''
    return run_applescript(script)

def paste_string_via_applescript(text_to_paste):
    """ Puts text on clipboard and tells System Events to paste (Cmd+V).
        Assumes the target application (Word) is already frontmost.
    """
    try:
        # 1. Put the desired sum onto the clipboard
        pyperclip.copy(text_to_paste)
        print(f"Copied '{text_to_paste}' to clipboard for pasting.")
        time.sleep(0.1) # Small delay to ensure clipboard is updated

        # 2. Tell System Events to simulate Cmd+V
        script = '''
            tell application "System Events"
                keystroke "v" using command down
            end tell
        '''
        run_applescript(script) # We don't need the output here
        print("Sent paste command (Cmd+V).")
        return True
    except Exception as e:
        # Catch potential pyperclip errors too
        print(f"Error during paste process: {e}")
        return False

def show_notification(title, text):
     """Uses osascript to show a macOS notification (Fallback)."""
     if sys.platform != 'darwin':
        print(f"Notification (non-macOS): {title} - {text}")
        return
     try:
         safe_text = text.replace('"', '\\"')
         safe_title = title.replace('"', '\\"')
         script = f'display notification "{safe_text}" with title "{safe_title}"'
         subprocess.run(['osascript', '-e', script], check=True, timeout=5)
     except Exception as e:
         print(f"Error showing notification: {e}")
         print(f"NOTIFICATION: {title} - {text}")

def format_number_indonesian(number, decimal_places=0):
    """Formats a number into Indonesian string format (e.g., 1.234,56)."""
    try:
        # Format with standard US locale first to handle decimals correctly
        # Using f-string formatting for precision control
        formatted_us = f"{number:,.{decimal_places}f}"

        # Standard f-string formatting uses ',' for thousands and '.' for decimal.
        # We need to swap them.
        # Replace comma with a temporary placeholder
        temp_replace = formatted_us.replace(',', 'TEMP_THOUSANDS')
        # Replace dot with the desired decimal separator (comma)
        temp_replace = temp_replace.replace('.', OUTPUT_DECIMAL_SEPARATOR)
        # Replace the placeholder with the desired thousands separator (dot)
        final_formatted = temp_replace.replace('TEMP_THOUSANDS', OUTPUT_THOUSANDS_SEPARATOR)

        # If decimal_places was 0, the formatted_us string might end with ".0"
        # or ",0" after our swaps. We need to remove that specific ending.
        if decimal_places == 0:
           if final_formatted.endswith(f"{OUTPUT_DECIMAL_SEPARATOR}0"):
              final_formatted = final_formatted[:-2] # Remove the last 2 chars (e.g., ",0")

        # Handle potential negative sign placement if needed (usually correct)
        return final_formatted
    except Exception as e:
        print(f"Error formatting number: {e}")
        return str(number) # Fallback to simple string conversion


# --- Main Monitoring Logic ---

print("--- Automatic Sum (Indonesian Format) & Paste to Word Monitor ---")
print(f"Monitoring clipboard. Will paste sum into '{WORD_APP_NAME}' if active.")
print("Press Ctrl+C to stop.")

previous_clipboard_content = ""
try:
    # Get initial clipboard content to avoid immediate trigger on start
    previous_clipboard_content = pyperclip.paste()
except Exception as e:
    print(f"Warning: Could not read initial clipboard content. {e}")


try:
    while True:
        current_clipboard_content = None
        try:
            current_clipboard_content = pyperclip.paste()
        except Exception as e:
            # Handle potential errors accessing clipboard (e.g., if copied by some protected apps)
            print(f"Error reading clipboard: {e}. Skipping this check.")
            # If clipboard is inaccessible, wait and try again
            time.sleep(CHECK_INTERVAL_SECONDS * 2) # Wait a bit longer
            continue # Skip the rest of this loop iteration

        # Proceed only if clipboard content was read successfully and changed
        if current_clipboard_content is not None and current_clipboard_content != previous_clipboard_content:
            print("\nClipboard changed.")
            # Store the *new* content as the "previous" for the *next* check *before* modification
            original_new_content = current_clipboard_content
            previous_clipboard_content = current_clipboard_content # Update tracking

            # Regex to find potential numbers (including ., and sign)
            # It's broad, we will clean and validate later
            # Regex to find sequences potentially representing numbers (digits, dots, commas)
            potential_numbers_str = re.findall(r"[-+]?[\d.,]+", current_clipboard_content)
            print(f"Menemukan string angka potensial: {potential_numbers_str}")
            total_sum = 0.0
            valid_numbers = []

            for num_str in potential_numbers_str:
                # *** START INDONESIAN FORMAT CLEANING ***
                # 1. Remove thousands separators (dots)
                cleaned_str = num_str.replace('.', '').replace(',', '.')
                # *** END INDONESIAN FORMAT CLEANING ***

                try:
                    # Attempt conversion *after* cleaning
                    number = float(cleaned_str)
                    # Check if the original string actually contained digits
                    # (prevents converting just "." or "," which become "" or ".")
                    if any(char.isdigit() for char in num_str):
                         total_sum += number
                         valid_numbers.append(number)
                         # print(f"  Berhasil konversi '{num_str}' -> '{cleaned_str}' -> {number}") # Debug
                    # else: # Debug
                    #    print(f"  (Mengabaikan '{num_str}' -> '{cleaned_str}' - tidak ada digit)")

                except ValueError:
                    # This string wasn't a valid number even after cleaning
                    # (e.g., "1.2.3,4,5" or "..." or ",,")
                    # print(f"  (Mengabaikan '{num_str}' -> '{cleaned_str}' - ValueError)") # Debug
                    pass # Silently ignore invalid conversions

            if valid_numbers:
                num_count = len(valid_numbers)

                # --- ROUNDING STEP ---
                rounded_total_sum = round(total_sum)
                # --- END ROUNDING STEP ---

                # Format the ROUNDED integer sum into Indonesian string format
                # Pass the rounded integer and explicitly request 0 decimal places for formatting

                sum_string_formatted = format_number_indonesian(rounded_total_sum, decimal_places=0)

                print(f"Angka yang berhasil di-parse (setelah dibersihkan): {valid_numbers}")
                print(f"JUMLAH ASLI = {total_sum}") # Show original sum for comparison
                print(f"JUMLAH DIBULATKAN = {rounded_total_sum} (String diformat: {sum_string_formatted})")

                # Check the frontmost application
                front_app = get_frontmost_app()
                print(f"Frontmost application: '{front_app}'")

                if front_app == WORD_APP_NAME:
                    print(f"'{WORD_APP_NAME}' is active. Attempting to paste sum...")
                    # --- PASTE ACTION ---
                    if paste_string_via_applescript(sum_string_formatted):
                       print(f"Successfully pasted '{sum_string_formatted}' into Word.")
                       # IMPORTANT: The clipboard now holds sum_string. Update
                       # previous_clipboard_content again to prevent the script
                       # immediately re-calculating the sum from the pasted value.
                       previous_clipboard_content = sum_string_formatted
                    else:
                       print("Paste attempt failed. Showing notification instead.")
                       # Fallback to notification if paste fails
                       show_notification("Clipboard Sum (Paste Failed)", f"Sum = {sum_string_formatted}")
                       # Decide if clipboard should contain the sum or original content after failed paste
                       # Let's leave the sum on the clipboard for now.
                       pyperclip.copy(sum_string_formatted)
                       previous_clipboard_content = sum_string_formatted # Track clipboard holds sum

                else:
                    print(f"'{WORD_APP_NAME}' is not active. Putting sum on clipboard and notifying.")
                    # --- NOTIFICATION ACTION (Word not active) ---
                    # Put the sum on the clipboard anyway, user might want it
                    pyperclip.copy(sum_string_formatted)
                    # Update tracking since we modified the clipboard
                    previous_clipboard_content = sum_string_formatted
                    show_notification("Clipboard Sum Calculated", f"Sum = {sum_string_formatted} (Copied to clipboard)")

            else:
                 print("No valid numbers found in new clipboard content.")
                 # Do nothing further if no numbers found

        # Wait before the next check
        time.sleep(CHECK_INTERVAL_SECONDS)

except KeyboardInterrupt:
    print("\n--- Monitor stopped by user ---")
except Exception as e:
    # Catch any other unexpected errors in the main loop
    print(f"\nAn critical error occurred in the main loop: {e}")
finally:
    print("Exiting.")