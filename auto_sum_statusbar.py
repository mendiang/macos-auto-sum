import pyperclip
import time
import re
import subprocess
import sys
import rumps # <-- Import library rumps

# --- Konfigurasi (tetap sama) ---
CHECK_INTERVAL_SECONDS = 1.0
WORD_APP_NAME = "Microsoft Word"
OUTPUT_DECIMAL_SEPARATOR = ','
OUTPUT_THOUSANDS_SEPARATOR = '.'

# --- Fungsi Helper (tetap sama, tidak perlu diubah) ---
def run_applescript(script):
    if sys.platform != 'darwin': return None
    try:
        process = subprocess.run(
            ['osascript', '-e', script], capture_output=True, text=True, check=True, timeout=5
        )
        return process.stdout.strip()
    except subprocess.CalledProcessError as e:
        print(f"AppleScript Error: {e.stderr}")
        return None
    except Exception as e:
        print(f"Error running AppleScript: {e}")
        return None

def get_frontmost_app():
    script = '''
        tell application "System Events"
            try
                name of first application process whose frontmost is true
            on error
                 return ""
            end try
        end tell
    '''
    return run_applescript(script)

def paste_string_via_applescript(text_to_paste):
    try:
        pyperclip.copy(text_to_paste)
        print(f"Copied '{text_to_paste}' to clipboard for pasting.")
        time.sleep(0.2)
        script = '''
            tell application "System Events"
                keystroke "v" using command down
            end tell
        '''
        run_applescript(script)
        print("Sent paste command (Cmd+V).")
        return True
    except Exception as e:
        print(f"Error during paste process: {e}")
        return False

def show_rumps_notification(title, subtitle, message):
    """Gunakan notifikasi bawaan rumps."""
    try:
        rumps.notification(title=title, subtitle=subtitle, message=message)
    except Exception as e:
        print(f"Error showing rumps notification: {e}")
        # Fallback ke print jika notifikasi gagal
        print(f"NOTIFICATION: {title} - {subtitle} - {message}")


def format_number_indonesian(number, decimal_places=0):
    try:
        formatted_us = f"{number:,.{max(0, decimal_places)}f}"
        temp_replace = formatted_us.replace(',', 'TEMP_THOUSANDS')
        temp_replace = temp_replace.replace('.', OUTPUT_DECIMAL_SEPARATOR)
        final_formatted = temp_replace.replace('TEMP_THOUSANDS', OUTPUT_THOUSANDS_SEPARATOR)
        if decimal_places == 0 and final_formatted.endswith(f"{OUTPUT_DECIMAL_SEPARATOR}0"):
            final_formatted = final_formatted[:-2]
        return final_formatted
    except Exception as e:
        print(f"Error formatting number: {e}")
        return str(int(round(number)) if isinstance(number, float) else number)

# --- Kelas Aplikasi Status Bar ---
class AutoSumApp(rumps.App):
    def __init__(self):
        # Judul awal aplikasi (ikon saja)
        super(AutoSumApp, self).__init__("∑", title=None, quit_button=None) # "∑" adalah ikon simbol Sum

        # State aplikasi
        self.monitoring_active = False
        self.clipboard_timer = None
        self.previous_clipboard_content = ""

        # Definisi item menu
        self.menu_start = rumps.MenuItem("Mulai Monitoring", callback=self.start_monitoring)
        self.menu_stop = rumps.MenuItem("Hentikan Monitoring", callback=self.stop_monitoring)
        self.menu_stop.set_callback(self.stop_monitoring) # Pastikan callback terpasang
        self.menu = [self.menu_start, self.menu_stop]

        # Nonaktifkan menu "Stop" di awal
        self.menu_stop.set_callback(None) # Hapus callback sementara agar tidak bisa diklik


    def start_monitoring(self, sender):
        """Dipanggil saat menu 'Mulai Monitoring' diklik."""
        if not self.monitoring_active:
            print("Starting monitoring...")
            self.monitoring_active = True
            self.title = "∑•" # Ubah ikon/judul untuk indikasi aktif

            # Update state menu
            self.menu_start.set_callback(None) # Nonaktifkan Start
            self.menu_stop.set_callback(self.stop_monitoring) # Aktifkan Stop

            # Baca clipboard awal
            try:
                self.previous_clipboard_content = pyperclip.paste()
            except Exception as e:
                print(f"Warning: Could not read initial clipboard. {e}")
                self.previous_clipboard_content = ""

            # Buat dan mulai timer untuk check_clipboard
            self.clipboard_timer = rumps.Timer(self.check_clipboard, CHECK_INTERVAL_SECONDS)
            self.clipboard_timer.start()
            print("Monitoring started.")
            show_rumps_notification("Auto Sum", "Status", "Monitoring Clipboard Dimulai")


    def stop_monitoring(self, sender):
        """Dipanggil saat menu 'Hentikan Monitoring' diklik."""
        if self.monitoring_active:
            print("Stopping monitoring...")
            self.monitoring_active = False
            self.title = None # Kembali ke ikon saja

            # Hentikan timer jika ada
            if self.clipboard_timer is not None:
                self.clipboard_timer.stop()
                self.clipboard_timer = None

            # Update state menu
            self.menu_start.set_callback(self.start_monitoring) # Aktifkan Start
            self.menu_stop.set_callback(None) # Nonaktifkan Stop
            print("Monitoring stopped.")
            show_rumps_notification("Auto Sum", "Status", "Monitoring Clipboard Dihentikan")


    # @rumps.timer(CHECK_INTERVAL_SECONDS) # Alternatif: bisa pakai decorator timer
    def check_clipboard(self, sender=None): # sender diisi oleh rumps.Timer
        """Fungsi inti yang memeriksa clipboard secara berkala."""
        if not self.monitoring_active: # Pengaman ekstra
            return

        current_clipboard_content = None
        try:
            current_clipboard_content = pyperclip.paste()
        except Exception as e:
            print(f"Error reading clipboard: {e}. Skipping check.")
            # Di lingkungan status bar, mungkin lebih baik diam daripada error terus menerus
            # Pertimbangkan untuk menghentikan monitoring jika error berulang?
            return

        if current_clipboard_content is not None and current_clipboard_content != self.previous_clipboard_content:
            print("\nClipboard changed.")
            original_new_content = current_clipboard_content
            self.previous_clipboard_content = current_clipboard_content

            # Regex & Logika Penjumlahan (Sama seperti sebelumnya)
            potential_numbers_str = re.findall(r"[-+]?[\d.,]+", current_clipboard_content)
            print(f"Found potential numbers: {potential_numbers_str}")
            total_sum = 0.0
            valid_numbers = []

            for num_str in potential_numbers_str:
                cleaned_str = num_str.replace('.', '').replace(',', '.')
                try:
                    number = float(cleaned_str)
                    if any(char.isdigit() for char in num_str):
                        total_sum += number
                        valid_numbers.append(number)
                except ValueError:
                    pass

            if valid_numbers:
                rounded_total_sum = round(total_sum)
                sum_string_formatted = format_number_indonesian(rounded_total_sum, decimal_places=0)

                print(f"Parsed numbers: {valid_numbers}")
                print(f"Original Sum = {total_sum}")
                print(f"Rounded Sum = {rounded_total_sum} (Formatted: {sum_string_formatted})")

                front_app = get_frontmost_app()
                print(f"Frontmost app: '{front_app}'")

                if front_app == WORD_APP_NAME:
                    print(f"'{WORD_APP_NAME}' active. Attempting paste...")
                    if paste_string_via_applescript(sum_string_formatted):
                        print(f"Pasted '{sum_string_formatted}' into Word.")
                        # Update previous content agar tidak re-trigger dari hasil paste
                        self.previous_clipboard_content = sum_string_formatted
                        show_rumps_notification("Auto Sum", f"Pasted to {WORD_APP_NAME}", f"Jumlah = {sum_string_formatted}")
                    else:
                        print("Paste failed. Showing notification.")
                        show_rumps_notification("Auto Sum", "Paste Gagal", f"Jumlah = {sum_string_formatted}")
                        pyperclip.copy(sum_string_formatted)
                        self.previous_clipboard_content = sum_string_formatted
                else:
                    print(f"'{WORD_APP_NAME}' not active. Copying sum and notifying.")
                    pyperclip.copy(sum_string_formatted)
                    self.previous_clipboard_content = sum_string_formatted
                    show_rumps_notification("Auto Sum", "Jumlah Dihitung", f"Jumlah = {sum_string_formatted} (Disalin)")
            else:
                print("No valid numbers found.")

    @rumps.clicked("Quit") # Menambahkan menu Quit standar
    def quit_app(self, sender):
        """Dipanggil saat menu 'Quit' diklik."""
        print("Quit clicked.")
        self.stop_monitoring(None) # Pastikan timer berhenti sebelum keluar
        rumps.quit_application()


# --- Menjalankan Aplikasi ---
if __name__ == "__main__":
    # Set locale jika perlu (biasanya tidak untuk format manual ini)
    # import locale
    # try:
    #    locale.setlocale(locale.LC_ALL, 'id_ID.UTF-8') # Atau locale sistem Anda
    # except locale.Error:
    #    print("Warning: Indonesian locale not found, using default.")

    print("Starting Auto Sum Status Bar App...")
    app = AutoSumApp()
    app.run()
    print("Application has quit.")