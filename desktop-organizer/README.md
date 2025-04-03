# Windows Desktop Organizer PowerShell Script

A PowerShell script to automatically organize files on the Windows Desktop into categorized folders based on file extension.

## Features

* Scans the user's Desktop for files.
* Moves files into a timestamped root organization folder (e.g., `Desktop Organization 2025-04-02 2345`).
* Categorizes files based on common extensions (Documents, Images, Videos, Archives, etc.). Customizable via `$fileTypeMappings` in the script.
* Creates category folders if they don't exist.
* Handles files with unknown extensions or no extension.
* Renames files automatically if a file with the same name already exists in the destination.
* Generates a detailed log file (`Desktop_Organization_Log_YYYY-MM-DD_HHmmss.log`) inside the organization folder.
* Skips the script itself and its log file.

## Prerequisites

* Windows Operating System
* PowerShell (usually built-in)

## How to Use

1.  **Download:** Download the `Organize-Desktop.ps1` script from this repository.
2.  **Customize (Optional):** Open the script in a text editor and modify the `$fileTypeMappings` variable near the top to change categories or add/remove file types.
3.  **Execution Policy:** You may need to adjust your PowerShell Execution Policy. Open PowerShell as **Administrator** and run:
    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
    ```
    You can change `RemoteSigned` to `Unrestricted` if needed, but be aware of the security implications. Answer 'Y' or 'A' when prompted.
4.  **Run:**
    * Right-click the `Organize-Desktop.ps1` file and select "Run with PowerShell".
    * OR Open PowerShell, navigate (`cd`) to the script's directory, and run: `.\Organize-Desktop.ps1`
5.  **Check Results:** A new `Desktop Organization [Date Time]` folder will appear on your Desktop containing the organized files and the log file.

## Disclaimer

Use this script at your own risk. While designed to be safe, it's always recommended to **back up your Desktop** before running it for the first time, especially if you have critical files there.

## License

MIT License
