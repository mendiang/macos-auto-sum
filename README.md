# Auto Sum Status Bar App for macOS

A simple macOS utility that runs in the status bar, monitors the clipboard for numbers (using Indonesian number format), calculates their sum, and optionally pastes the result into Microsoft Word.

## Description

This Python application provides a convenient way to quickly sum numbers copied from various sources on your Mac. It's designed with Indonesian number formatting in mind (using `.` as a thousands separator and `,` as the decimal separator, e.g., `1.234,56`). The final sum is rounded to the nearest integer before being processed.

The app features a status bar icon (∑) with menu controls to start/stop monitoring and manually process the current clipboard content. If Microsoft Word is the active application when numbers are processed, the sum is automatically pasted. Otherwise, a notification is shown, and the sum is copied to the clipboard.

## Features

*   Runs discreetly in the macOS status bar.
*   Monitors clipboard changes automatically when active.
*   Extracts numbers formatted according to Indonesian locale (e.g., `1.500.000`, `-123,45`).
*   Calculates the total sum of extracted numbers.
*   Rounds the final sum to the nearest integer (no decimal places in the output).
*   Copies the formatted integer sum to the clipboard.
*   **Auto-Paste:** Automatically pastes the sum into Microsoft Word if it is the frontmost application.
*   **Manual Trigger:** Option to manually process the current clipboard content via the status bar menu.
*   Provides system notifications for status updates and results.
*   Simple menu controls: Start Monitoring, Stop Monitoring, Process Clipboard Now, Quit.

## Prerequisites

*   **macOS:** This application relies on macOS-specific features (AppleScript, `rumps` library).
*   **Python 3:** Ensure you have Python 3 installed. You can check with `python3 --version`.
*   **Microsoft Word:** Required for the automatic pasting feature.
*   **Git:** (Optional) If you are cloning this repository.

## Installation

1.  **Clone the Repository (or download the files):**
    ```bash
    git clone https://github.com/your-username/your-repository-name.git
    # Ganti URL dengan URL repositori Anda
    cd your-repository-name
    ```

2.  **Install Required Python Packages:**
    Navigate to the project directory in your Terminal and run:
    ```bash
    pip3 install rumps pyperclip
    ```
    *Note: `py2app` is only needed if you plan to build the standalone `.app` bundle (see below).*

## Running the Application

There are two main ways to run the application:

**Method 1: Directly from Terminal (for testing/development)**

1.  Navigate to the project directory in Terminal.
2.  Run the main script:
    ```bash
    python3 auto_sum_statusbar.py
    ```
3.  The `∑` icon will appear in your status bar.
4.  **Important:** The Terminal window where you ran the command must remain open while the application is running.
5.  To stop the application, either use the "Quit" menu from the status bar icon or press `Control + C` in the Terminal window.

**Method 2: Building a Standalone `.app` Bundle (Recommended for regular use)**

This packages the script and its dependencies into a standard macOS application that you can run without keeping a Terminal window open.

1.  **Install `py2app`:**
    ```bash
    pip3 install py2app
    ```
2.  **Build the Application:**
    Make sure you are in the project directory (where `setup.py` is located) and run:
    ```bash
    python3 setup.py py2app
    ```
3.  If the build is successful, a `dist` folder will be created. Inside `dist`, you will find your application bundle (e.g., `AutoSumApp.app`).
4.  **Run:** Move `AutoSumApp.app` to your Applications folder (or anywhere you like) and double-click it to run. The `∑` icon will appear in the status bar. The app runs in the background.
5.  **Stop:** Click the status bar icon and select "Quit" from the menu.

## How to Use

1.  **Run the application** using one of the methods above.
2.  Click the `∑` icon in your status bar.
3.  Select **"Mulai Monitoring"**. The icon might change slightly (e.g., `∑•`) to indicate it's active.
4.  Copy any text containing numbers (e.g., from a document, spreadsheet, or webpage). Numbers should ideally use `.` for thousands and `,` for decimals, but the script attempts cleaning.
5.  **Automatic Action:**
    *   If the clipboard content *changes* and contains valid numbers:
        *   The sum will be calculated (and rounded).
        *   If **Microsoft Word** is the active application, the sum will be automatically pasted into your document.
        *   If Word is *not* active, a system notification will appear showing the calculated sum, and the sum will be copied to your clipboard for manual pasting.
6.  **Manual Action:**
    *   If you copy text and the app doesn't react (because the content hasn't changed since the last check), or if you just want to re-process the current clipboard, click the `∑` icon and select **"Proses Clipboard Sekarang"**. This will perform the calculation and paste/notify action regardless of whether the clipboard content has changed recently.
7.  Select **"Hentikan Monitoring"** to stop the app from watching the clipboard. The icon will revert to normal.
8.  Select **"Quit"** to stop monitoring and close the application completely.

## macOS Permissions

The first time the application tries to interact with Microsoft Word or simulate pasting, macOS will likely prompt you for permissions. You **must grant** these permissions for the app to function correctly.

*   Check **System Settings > Privacy & Security > Automation**: Ensure your app (or Terminal/Python if running directly) is allowed to control "Microsoft Word" and "System Events".
*   Check **System Settings > Privacy & Security > Accessibility**: Sometimes needed for controlling keystrokes ("System Events").

---

*Selamat menggunakan! Feel free to contribute or report issues.*