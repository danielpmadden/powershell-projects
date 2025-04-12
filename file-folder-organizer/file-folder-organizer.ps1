<#
.SYNOPSIS
    Organizes files into a hierarchical folder structure based on file types.

.DESCRIPTION
    This script creates a hierarchical organization system where files are sorted 
    first into major categories (Documents, Media, etc.) and then into subcategories
    based on specific file types (.docx, .pdf, etc.).
#>

# ===== USER CONFIGURATION =====
# Set these values according to your needs
$sourceFolderPath = "C:\Users\username\Downloads"  # Change this to your folder path
$destinationPath = "C:\Users\username\Downloads\Organized"  # Change this to your destination
$copyInsteadOfMove = $false  # Set to $true to copy files instead of moving them
$recurseSubfolders = $false  # Set to $true to include files in subfolders
# =============================

# Validate source folder
if (-not (Test-Path -Path $sourceFolderPath -PathType Container)) {
    Write-Error "Source folder '$sourceFolderPath' does not exist or is not accessible."
    exit 1
}

# Create destination folder if it doesn't exist
if (-not (Test-Path -Path $destinationPath -PathType Container)) {
    try {
        New-Item -Path $destinationPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Created destination folder: $destinationPath"
    } catch {
        Write-Error "Could not create destination folder '$destinationPath'. Error: $($_.Exception.Message)"
        exit 1
    }
}

# Define the name for the log file
$logFileName = "Organization_Log_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
$logFilePath = Join-Path -Path $destinationPath -ChildPath $logFileName

# Define hierarchical file type mappings
# Format: Extension = @("Main Category", "Subcategory")
$fileHierarchy = @{
    # Document files
    ".doc"  = @("Documents", "Microsoft Word Documents")
    ".docx" = @("Documents", "Microsoft Word Documents")
    ".rtf"  = @("Documents", "Rich Text Documents")
    ".txt"  = @("Documents", "Text Files")
    ".pdf"  = @("Documents", "PDFs")
    ".odt"  = @("Documents", "OpenDocument Text")
    ".epub" = @("Documents", "eBooks")
    ".mobi" = @("Documents", "eBooks")
    ".md"   = @("Documents", "Markdown Files")
    ".log"  = @("Documents", "Log Files")
    ".tex"  = @("Documents", "LaTeX Files")
    ".json" = @("Documents", "Data Files")
    ".xml"  = @("Documents", "Data Files")
    ".yml"  = @("Documents", "Data Files")
    ".yaml" = @("Documents", "Data Files")

    # Spreadsheet files
    ".xls"  = @("Documents", "Microsoft Excel Spreadsheets")
    ".xlsx" = @("Documents", "Microsoft Excel Spreadsheets")
    ".csv"  = @("Documents", "CSV Files")
    ".tsv"  = @("Documents", "CSV Files")
    ".ods"  = @("Documents", "OpenDocument Spreadsheets")

    # Presentation files
    ".ppt"  = @("Documents", "Microsoft PowerPoint Presentations")
    ".pptx" = @("Documents", "Microsoft PowerPoint Presentations")
    ".odp"  = @("Documents", "OpenDocument Presentations")

    # Image files
    ".jpg"  = @("Media", "Images")
    ".jpeg" = @("Media", "Images")
    ".png"  = @("Media", "Images")
    ".gif"  = @("Media", "Images")
    ".bmp"  = @("Media", "Images")
    ".tiff" = @("Media", "Images")
    ".webp" = @("Media", "Images")
    ".heic" = @("Media", "Images")
    ".ico"  = @("Media", "Icons")
    ".svg"  = @("Media", "Vector Images")
    ".psd"  = @("Media", "Photoshop Files")
    ".ai"   = @("Media", "Illustrator Files")

    # Audio files
    ".mp3"  = @("Media", "Audio")
    ".wav"  = @("Media", "Audio")
    ".flac" = @("Media", "Audio")
    ".aac"  = @("Media", "Audio")
    ".ogg"  = @("Media", "Audio")
    ".wma"  = @("Media", "Audio")
    ".m4a"  = @("Media", "Audio")
    ".alac" = @("Media", "Audio")
    ".aiff" = @("Media", "Audio")
    ".opus" = @("Media", "Audio")

    # Video files
    ".mp4"  = @("Media", "Videos")
    ".avi"  = @("Media", "Videos")
    ".mkv"  = @("Media", "Videos")
    ".mov"  = @("Media", "Videos")
    ".wmv"  = @("Media", "Videos")
    ".flv"  = @("Media", "Videos")
    ".webm" = @("Media", "Videos")
    ".m4v"  = @("Media", "Videos")
    ".3gp"  = @("Media", "Videos")

    # Code files
    ".html"   = @("Code", "Web Files")
    ".css"    = @("Code", "Web Files")
    ".asp"    = @("Code", "Web Files")
    ".aspx"   = @("Code", "Web Files")
    ".jsp"    = @("Code", "Web Files")
    ".htaccess" = @("Code", "Web Files")
    ".js"     = @("Code", "JavaScript")
    ".jsx"    = @("Code", "JavaScript")
    ".ts"     = @("Code", "TypeScript")
    ".tsx"    = @("Code", "TypeScript")
    ".py"     = @("Code", "Python")
    ".java"   = @("Code", "Java")
    ".c"      = @("Code", "C/C++")
    ".cpp"    = @("Code", "C/C++")
    ".cs"     = @("Code", "C#")
    ".php"    = @("Code", "PHP")
    ".sql"    = @("Code", "SQL")
    ".ps1"    = @("Code", "PowerShell")
    ".sh"     = @("Code", "Shell Scripts")
    ".bat"    = @("Code", "Batch Files")
    ".go"     = @("Code", "Go")
    ".rb"     = @("Code", "Ruby")
    ".r"      = @("Code", "R")
    ".swift"  = @("Code", "Swift")
    ".kt"     = @("Code", "Kotlin")
    ".vue"    = @("Code", "Vue.js")
    ".rs"     = @("Code", "Rust")
    ".ipynb"  = @("Code", "Jupyter Notebooks")

    # Archives
    ".zip"  = @("Archives", "ZIP Files")
    ".rar"  = @("Archives", "RAR Files")
    ".7z"   = @("Archives", "7-Zip Files")
    ".tar"  = @("Archives", "TAR Files")
    ".gz"   = @("Archives", "GZ Files")
    ".bz2"  = @("Archives", "Bzip2 Files")
    ".xz"   = @("Archives", "XZ Files")
    ".cab"  = @("Archives", "CAB Files")
    ".dmg"  = @("Archives", "Mac Disk Images")

    # Executables
    ".exe"  = @("Programs", "Executables")
    ".msi"  = @("Programs", "Installers")
    ".app"  = @("Programs", "Mac Applications")
    ".apk"  = @("Programs", "Android Packages")
    ".deb"  = @("Programs", "Linux Packages")
    ".rpm"  = @("Programs", "Linux Packages")
    ".jar"  = @("Programs", "Java Archives")
    ".bin"  = @("Programs", "Binary Files")
    ".run"  = @("Programs", "Linux Installers")

    # Disk Images / Virtualization
    ".iso"  = @("Disk Images", "ISO Files")
    ".vhd"  = @("Disk Images", "Virtual Hard Disks")
    ".vmdk" = @("Disk Images", "Virtual Hard Disks")
    ".ova"  = @("Disk Images", "VM Packages")
    ".ovf"  = @("Disk Images", "VM Metadata")

    # Config / Registry
    ".cfg"  = @("Configuration", "Config Files")
    ".ini"  = @("Configuration", "Config Files")
    ".reg"  = @("Configuration", "Registry Files")

    # Misc
    ".torrent" = @("Downloads", "Torrent Files")
}


# Default category and subcategory for unknown file types
$defaultMainCategory = "Other Files"
$defaultSubCategory = "Miscellaneous"

# --- Start Processing ---
$startTime = Get-Date
Write-Host "Starting hierarchical file organization..." -ForegroundColor Cyan
Write-Host "Source: $sourceFolderPath"
Write-Host "Destination: $destinationPath"
Write-Host "Mode: $(if ($copyInsteadOfMove) {"Copy"} else {"Move"})"

# Log file initialization
$logHeader = "Hierarchical File Organization Log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`nSource: $sourceFolderPath`nDestination: $destinationPath`n"
$logHeader | Out-File -FilePath $logFilePath -Encoding UTF8

# Process files
$filesToProcess = Get-ChildItem -Path $sourceFolderPath -File -Recurse:$recurseSubfolders -ErrorAction SilentlyContinue
$totalFiles = ($filesToProcess | Measure-Object).Count

"Found $totalFiles files to process." | Out-File -FilePath $logFilePath -Append
Write-Host "Found $totalFiles files to process." -ForegroundColor Yellow

# Initialize counters
$processedCount = 0
$successCount = 0
$errorCount = 0

# Track created main categories and subcategories for reporting
$createdCategories = @{}

foreach ($file in $filesToProcess) {
    $processedCount++
    $statusMsg = "[$processedCount of $totalFiles] Processing: $($file.Name)"
    Write-Host $statusMsg
    $statusMsg | Out-File -FilePath $logFilePath -Append
    
    # Get file extension and determine hierarchy placement
    $extension = $file.Extension.ToLower()
    
    if ($fileHierarchy.ContainsKey($extension)) {
        $mainCategory = $fileHierarchy[$extension][0]
        $subCategory = $fileHierarchy[$extension][1]
    } elseif ([string]::IsNullOrEmpty($extension)) {
        $mainCategory = "No Extension"
        $subCategory = "Files Without Extension"
    } else {
        $mainCategory = $defaultMainCategory
        $subCategory = $defaultSubCategory
    }
    
    # Create the main category folder if needed
    $mainCategoryPath = Join-Path -Path $destinationPath -ChildPath $mainCategory
    if (-not (Test-Path -Path $mainCategoryPath -PathType Container)) {
        New-Item -Path $mainCategoryPath -ItemType Directory -Force | Out-Null
        "Created main category folder: $mainCategoryPath" | Out-File -FilePath $logFilePath -Append
        $createdCategories[$mainCategory] = @()
    }
    
    # Create the subcategory folder if needed
    $subCategoryPath = Join-Path -Path $mainCategoryPath -ChildPath $subCategory
    if (-not (Test-Path -Path $subCategoryPath -PathType Container)) {
        New-Item -Path $subCategoryPath -ItemType Directory -Force | Out-Null
        "Created subcategory folder: $subCategoryPath" | Out-File -FilePath $logFilePath -Append
        
        # Track for reporting
        if (-not $createdCategories.ContainsKey($mainCategory)) {
            $createdCategories[$mainCategory] = @()
        }
        if ($createdCategories[$mainCategory] -notcontains $subCategory) {
            $createdCategories[$mainCategory] += $subCategory
        }
    }
    
    # Handle file name conflicts
    $destFile = Join-Path -Path $subCategoryPath -ChildPath $file.Name
    $counter = 1
    $newName = $file.Name
    
    while (Test-Path -Path $destFile -PathType Leaf) {
        $newName = "$($file.BaseName)_$counter$($file.Extension)"
        $destFile = Join-Path -Path $subCategoryPath -ChildPath $newName
        $counter++
    }
    
    # Copy or move the file
    try {
        if ($copyInsteadOfMove) {
            Copy-Item -Path $file.FullName -Destination $destFile -Force
            $actionText = "Copied"
        } else {
            Move-Item -Path $file.FullName -Destination $destFile -Force
            $actionText = "Moved"
        }
        
        $resultMsg = "${actionTaken}: $($file.FullName) -> $destFile"
        $resultMsg | Out-File -FilePath $logFilePath -Append
        $successCount++
    } catch {
        $errorMsg = "ERROR: Failed to process $($file.FullName). $($_.Exception.Message)"
        Write-Host $errorMsg -ForegroundColor Red
        $errorMsg | Out-File -FilePath $logFilePath -Append
        $errorCount++
    }
}

# --- Report Created Folder Structure ---
$structureReport = "`nCreated Folder Structure:"
foreach ($mainCat in $createdCategories.Keys | Sort-Object) {
    $structureReport += "`n+ $mainCat"
    foreach ($subCat in $createdCategories[$mainCat] | Sort-Object) {
        $structureReport += "`n  - $subCat"
    }
}

$structureReport | Out-File -FilePath $logFilePath -Append
Write-Host $structureReport -ForegroundColor Cyan

# --- Summary ---
$duration = (Get-Date) - $startTime
$summaryMsg = @"

Organization Complete!
--------------------------------------------
Total files processed: $totalFiles
Successfully processed: $successCount
Errors: $errorCount
Duration: $($duration.TotalSeconds.ToString("0.00")) seconds
--------------------------------------------
"@

Write-Host $summaryMsg -ForegroundColor Green
$summaryMsg | Out-File -FilePath $logFilePath -Append

Write-Host "Log file created at: $logFilePath"
