<#
.SYNOPSIS
    Organizes files on the user's Desktop by moving them into folders based on their file extension.
    Creates a log file detailing all move operations.

.DESCRIPTION
    This script scans the current user's Desktop for files (excluding directories).
    It determines the file type based on the extension and uses a predefined mapping
    to decide which destination folder (within a main 'Desktop Organization' folder)
    the file should be moved to.
    If the destination folder doesn't exist, it's created.
    All actions (moves and errors) are logged to a timestamped text file on the Desktop.
    The script will skip itself and its own log file.

.NOTES
    Version:       1.0
    Prerequisites: PowerShell
    Execution:     Save as .ps1 file (e.g., Organize-Desktop.ps1). Right-click -> Run with PowerShell,
                   or run from a PowerShell terminal: .\Organize-Desktop.ps1
    Important:     Review and potentially customize the $fileTypeMappings variable before running.
                   Ensure your PowerShell Execution Policy allows running local scripts.
                   Consider backing up your Desktop before the first run.

.EXAMPLE
    .\start-desktoporganizer.ps1
    (Runs the script with default settings)
#>

# --- Script Configuration ---

# Get the path to the current user's Desktop
$desktopPath = [Environment]::GetFolderPath('Desktop')

# Define the name of the main folder on the Desktop where organized folders will be created
$organizationRootFolderName = "Desktop Organization $(Get-Date -Format 'yyyy-MM-dd HHmm')" # Adds timestamp to avoid conflicts if run multiple times
$organizationRootPath = Join-Path -Path $desktopPath -ChildPath $organizationRootFolderName

# Define the name for the log file
# It will be placed *inside* the $organizationRootPath once that's created
$logFileName = "Desktop_Organization_Log_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
$logFilePath = Join-Path -Path $organizationRootPath -ChildPath $logFileName # Initial log path definition

# Define the mapping between file extensions (lowercase) and target folder names
# Add or modify categories and extensions as needed.
$fileTypeMappings = @{
    # Documents
    ".txt"  = "Documents"
    ".doc"  = "Documents"
    ".docx" = "Documents"
    ".pdf"  = "Documents"
    ".rtf"  = "Documents"
    ".odt"  = "Documents"
    ".wpd"  = "Documents"
    ".log"  = "Logs and Text" # Separate category for logs if desired

    # Spreadsheets
    ".xls"  = "Spreadsheets"
    ".xlsx" = "Spreadsheets"
    ".csv"  = "Spreadsheets"
    ".ods"  = "Spreadsheets"

    # Presentations
    ".ppt"  = "Presentations"
    ".pptx" = "Presentations"
    ".odp"  = "Presentations"

    # Images
    ".jpg"  = "Images"
    ".jpeg" = "Images"
    ".png"  = "Images"
    ".gif"  = "Images"
    ".bmp"  = "Images"
    ".tif"  = "Images"
    ".tiff" = "Images"
    ".svg"  = "Images"
    ".webp" = "Images"
    ".heic" = "Images"
    ".psd"  = "Images - Photoshop" # Specific category example

    # Audio
    ".mp3"  = "Audio"
    ".wav"  = "Audio"
    ".aac"  = "Audio"
    ".flac" = "Audio"
    ".ogg"  = "Audio"
    ".wma"  = "Audio"
    ".m4a"  = "Audio"

    # Video
    ".mp4"  = "Videos"
    ".avi"  = "Videos"
    ".mkv"  = "Videos"
    ".mov"  = "Videos"
    ".wmv"  = "Videos"
    ".flv"  = "Videos"
    ".webm" = "Videos"

    # Archives
    ".zip"  = "Archives"
    ".rar"  = "Archives"
    ".7z"   = "Archives"
    ".tar"  = "Archives"
    ".gz"   = "Archives"

    # Code / Scripts
    ".ps1"  = "Scripts"
    ".bat"  = "Scripts"
    ".sh"   = "Scripts"
    ".py"   = "Scripts"
    ".js"   = "Scripts"
    ".html" = "Web Files"
    ".css"  = "Web Files"
    ".xml"  = "Code Files"
    ".json" = "Code Files"
    ".java" = "Code Files"
    ".cs"   = "Code Files"
    ".cpp"  = "Code Files"
    ".c"    = "Code Files"
    ".h"    = "Code Files"

    # Executables / Installers
    ".exe"  = "Executables"
    ".msi"  = "Installers"
    ".app"  = "Applications" # For macOS compatibility if cross-platform context matters, or general apps
    ".jar"  = "Java Apps"

    # Shortcuts
    ".lnk"  = "Shortcuts"
    ".url"  = "Shortcuts"

    # Other / Specific
    ".iso"  = "Disk Images"
    ".vhd"  = "Virtual Disks"
    ".vhdx" = "Virtual Disks"
    ".torrent" = "Torrents"

    # Add more mappings as needed...
}

# Folder name for files with extensions not found in the map
$defaultFolderName = "Other Files"

# --- Initialization ---

Write-Host "Starting Desktop Organization..."
Write-Host "Desktop Path: $desktopPath"
Write-Host "Organization Root Folder: $organizationRootFolderName"

# Create the main organization folder if it doesn't exist
if (-not (Test-Path -Path $organizationRootPath -PathType Container)) {
    try {
        New-Item -Path $organizationRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Created main organization folder: $organizationRootPath"
    } catch {
        Write-Error "FATAL: Could not create the main organization folder '$organizationRootPath'. Error: $($_.Exception.Message)"
        # Optional: Exit if the main folder cannot be created
        # exit 1
        # Or, attempt to log to the desktop directly if the org folder fails
        $logFilePath = Join-Path -Path $desktopPath -ChildPath $logFileName
        Write-Warning "Attempting to write log directly to Desktop."
    }
} else {
     Write-Host "Main organization folder already exists: $organizationRootPath"
     # Update log file path to be inside the existing folder
     $logFilePath = Join-Path -Path $organizationRootPath -ChildPath $logFileName
}


# Initialize the log file
$logHeader = @"
-----------------------------------------
Desktop Organization Log
Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Organization Root: $organizationRootPath
-----------------------------------------
"@
try {
    $logHeader | Out-File -FilePath $logFilePath -Encoding UTF8 -ErrorAction Stop
    Write-Host "Log file created at: $logFilePath"
} catch {
    Write-Error "FATAL: Could not create or write to the log file '$logFilePath'. Error: $($_.Exception.Message)"
    Write-Warning "Logging will not be available for this session."
    # Consider exiting if logging is critical
    # exit 1
}

# Function to simplify logging
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    # Only attempt to log if the log file path seems valid (was created or existed)
    if ($logFilePath -and (Test-Path -Path (Split-Path $logFilePath -Parent))) {
         try {
            "$(Get-Date -Format 'HH:mm:ss') - $Message" | Out-File -FilePath $logFilePath -Encoding UTF8 -Append -ErrorAction Stop
         } catch {
             Write-Warning "Failed to write to log file '$logFilePath'. Error: $($_.Exception.Message)"
         }
    } else {
        Write-Warning "Log file path is not valid. Cannot log message: $Message"
    }
}

# --- Main Processing Logic ---

Write-Log "Scanning Desktop for files..."
Write-Host "Scanning Desktop for files..."

# Get all *files* directly on the Desktop (not in subdirectories)
# Exclude the script itself and the log file's eventual location (though log isn't created yet, be safe)
# Also exclude the main organization folder itself
$scriptName = $MyInvocation.MyCommand.Name
$filesToProcess = Get-ChildItem -Path $desktopPath -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne $scriptName -and $_.Name -ne $organizationRootFolderName -and $_.Name -ne $logFileName }

# Counter for processed files
$processedCount = 0
$movedCount = 0
$errorCount = 0
$totalFiles = ($filesToProcess | Measure-Object).Count

Write-Host "Found $totalFiles files to potentially process."
Write-Log "Found $totalFiles files to potentially process."

# Loop through each file found
foreach ($file in $filesToProcess) {
    $processedCount++
    Write-Host "Processing ($processedCount/$totalFiles): $($file.Name)"

    # Get the file extension (lowercase for case-insensitive matching)
    $extension = $file.Extension.ToLower()

    # Determine the target subfolder name
    $targetFolderName = $null
    if ($fileTypeMappings.ContainsKey($extension)) {
        $targetFolderName = $fileTypeMappings[$extension]
    } elseif ([string]::IsNullOrEmpty($extension)) {
         # Handle files with no extension
         $targetFolderName = "Files Without Extension"
         Write-Log "INFO: File '$($file.Name)' has no extension. Assigning to '$targetFolderName'."
    }
    else {
        # Use the default folder name for unmapped extensions
        $targetFolderName = $defaultFolderName
        Write-Log "INFO: Extension '$extension' not found in mappings for file '$($file.Name)'. Assigning to default folder '$targetFolderName'."
    }

    # Construct the full path for the target folder (inside the organization root)
    $destinationFolderPath = Join-Path -Path $organizationRootPath -ChildPath $targetFolderName

    # Create the target subfolder if it doesn't exist
    if (-not (Test-Path -Path $destinationFolderPath -PathType Container)) {
        try {
            Write-Host "  Creating folder: $destinationFolderPath"
            New-Item -Path $destinationFolderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Log "CREATED FOLDER: '$destinationFolderPath'"
        } catch {
            Write-Warning "  ERROR: Could not create folder '$destinationFolderPath'. Skipping file '$($file.Name)'. Error: $($_.Exception.Message)"
            Write-Log "ERROR: Failed to create folder '$destinationFolderPath' for file '$($file.Name)'. Error: $($_.Exception.Message)"
            $errorCount++
            continue # Skip to the next file
        }
    }

    # Construct the full destination path for the file
    $destinationFilePath = Join-Path -Path $destinationFolderPath -ChildPath $file.Name

    # Check if a file with the same name already exists in the destination
    $counter = 1
    $originalFileNameWithoutExtension = $file.BaseName
    $originalExtension = $file.Extension # Keep original extension including case if needed for renaming
    while (Test-Path -Path $destinationFilePath -PathType Leaf) {
        Write-Warning "  WARNING: File '$($file.Name)' already exists in '$targetFolderName'. Renaming."
        $newName = "$($originalFileNameWithoutExtension)_$($counter)$($originalExtension)"
        $destinationFilePath = Join-Path -Path $destinationFolderPath -ChildPath $newName
        Write-Log "WARNING: File '$($file.Name)' already exists in '$targetFolderName'. Will attempt to move as '$newName'."
        $counter++
        # Safety break to prevent infinite loops in unlikely scenarios
        if ($counter -gt 100) {
             Write-Error "  ERROR: Could not find a unique name for '$($file.Name)' in '$targetFolderName' after 100 attempts. Skipping."
             Write-Log "ERROR: Could not find unique name for '$($file.Name)' in '$targetFolderName'. Skipped."
             $errorCount++
             continue 2 # Continue to the next file in the outer loop
        }
    }


    # Move the file
    try {
        Write-Host "  Moving '$($file.Name)' to '$targetFolderName'..."
        Move-Item -Path $file.FullName -Destination $destinationFilePath -Force -ErrorAction Stop
        Write-Log "MOVED: '$($file.FullName)' => '$destinationFilePath'"
        $movedCount++
    } catch {
        Write-Error "  ERROR: Failed to move '$($file.Name)' to '$destinationFolderPath'. Error: $($_.Exception.Message)"
        Write-Log "ERROR: Failed to move '$($file.FullName)'. Error: $($_.Exception.Message)"
        $errorCount++
    }
}

# --- Completion ---

$endTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$summary = @"

-----------------------------------------
Organization Complete
End Time: $endTime
-----------------------------------------
Total files scanned on Desktop: $totalFiles
Files successfully moved:     $movedCount
Files skipped due to errors:  $errorCount
Organization Root Folder:     $organizationRootPath
Log file location:            $logFilePath
-----------------------------------------
"@

Write-Host $summary -ForegroundColor Green
Write-Log $summary

Write-Host "Script finished."