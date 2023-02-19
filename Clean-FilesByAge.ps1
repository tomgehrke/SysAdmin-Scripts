# Clean-FilesByAge.ps1
#
# Arguments
# ===================
# -Path <path>         | Root folder path that the script will operate on
# -ReportFolder <path> | Path where the script will save a report of its actions
# -Age <n>             | Maximum age in days for a file or folder for it to be considered for deletion
# -Recurse             | Cause the script to recursively search subfolders for files and folders to delete
# -Hidden              | Cause the script to delete hidden files and folders
# -KeepEmptyFolders    | Cause the script to keep empty folders even if they are older than the specified age
# -TestMode            | Cause the script to only report on the files and folders it would have deleted, but not actually delete them

Param( 
  [Parameter(Mandatory = $false)][string]$Path, 
  [Parameter(Mandatory = $false)][string]$ReportPath = ".",
  [Parameter(Mandatory = $false)][int]$Age = 30,
  [Parameter(Mandatory = $false)][switch]$Recurse,
  [Parameter(Mandatory = $false)][switch]$Hidden,
  [Parameter(Mandatory = $false)][switch]$KeepEmptyFolders,
  [Parameter(Mandatory = $false)][switch]$TestMode,
  [Parameter(Mandatory = $false)][switch]$ShowDetail
)

# A path was not provided on the command line. Let the user select one.
if ($Path -eq "") {
  $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
  $FolderDialog.Description = "Select the root folder"
  $FolderDialog.RootFolder = "MyComputer"
  $FolderDialog.SelectedPath = $(Get-Location).Path
  $FolderDialog.ShowNewFolderButton = $false
  $ShowForm = $FolderDialog.ShowDialog()
  if ($ShowForm -eq "OK") {
    $Path = $FolderDialog.SelectedPath
  }
  else {
    Write-Output "[ERROR!] A target path must be selected!"
    Exit
  }
}

# Make sure the path exists. Exit if it does not.
if (!(Test-Path $Path -PathType Container)) {
  Write-Output "[ERROR] ""$Path"" does not exist!"
  Exit
}

# If provided, check for the report folder and create if it does not exist.
if ($ReportPath -and !(Test-Path -Path $ReportPath)) {
  Write-Host "- Backup folder does not exist. Creating it..."
  $null = New-Item -Path $ReportPath -ItemType Directory -Force
}

# Initialize Global variables
$global:Report = @()
$global:FoldersRemoved = 0
$global:Action = if ($TestMode) { "Report" } else { "Delete" }

# Initialize other variables
$Now = Get-Date -Format "yyyyMMddHHmm"
$ReportFile = $Path.TrimEnd('\\') -Replace '\\','-' -Replace ' ','_' -Replace '[^\w-_]',''
$RecoveredSpace = 0

# Function takes a size in bytes and returns a string with the size in a 
# more human-readable format (e.g. "5.5 MB").
function Format-Size {
  Param (
    [Parameter(Position = 0, Mandatory = $true)][int64]$SizeInBytes
  )

  if ($SizeInBytes -ge 1GB) { 
    return "{0:N2} GB" -f ($SizeInBytes / 1GB) 
  } elseif ($SizeInBytes -ge 1MB) { 
    return "{0:N2} MB" -f ($SizeInBytes / 1MB) 
  } elseif ($SizeInBytes -ge 1KB) { 
    return "{0:N2} KB" -f ($SizeInBytes / 1KB) 
  } else { 
    return "{0} Bytes" -f $SizeInBytes 
  }
}

# Function handles deleting what should be an empty folder
function Remove-Folder {
  Param( 
    [Parameter(Mandatory=$true)]$Folder,
    [Parameter(Mandatory=$false)][switch]$ShowOutput
  )  

  $TargetFolder = Get-Item -LiteralPath "$Folder" -Force:$Hidden

  # Grab these so we can report on them after the fact
  $CreatedDate = (Get-Item "$Folder").CreationTime
  $AccessedDate = (Get-Item "$Folder").LastAccessTime
  $ModifiedDate = (Get-Item "$Folder").LastWriteTime

  try {
    Remove-Item -LiteralPath "$($TargetFolder.FullName)" -Force:$Hidden -WhatIf:$TestMode -Confirm:$false
  } catch {
    $global:Report += [PSCustomObject]@{
      "Type"     = "Error"
      "Path"     = "$Folder"
      "Size"     = ""
      "SizeInBytes" = ""
      "Created"  = ""
      "Accessed" = ""
      "Modified" = ""
      "Action"   = "Error"
      "Detail"   = $_.Exception.Message
    }
    return
  }

  # Update our deleted folder count
  $global:FoldersRemoved ++

  $global:Report += [PSCustomObject]@{
    "Type"        = "Folder"
    "Path"        = "$Folder"
    "Size"        = ""
    "SizeInBytes" = ""
    "Created"     = $CreatedDate
    "Accessed"    = $AccessedDate
    "Modified"    = ModifiedDate
    "Action"      = $global:Action
    "Detail"      = ""
  }
}

# Function determines if a folder can be deleted by checking whether it has any children (folders or files)
# and if it passes the age requirement
function Test-Folder {
  Param (
    [Parameter(Mandatory=$true)][string]$Folder,
    [Parameter(Mandatory=$false)][switch]$Recurse,
    [Parameter(Mandatory=$false)][switch]$DeleteEmpty
  )

  # Grab these before we start messing in sub-folders
  $ModifiedDate = (Get-Item "$Folder").LastWriteTime

  # If the script was run with the -Recurse switch, we will drill down into all the child folders,
  # test them for "deletability" and then delete them if they pass.
  if ($Recurse) {
    $Children = Get-ChildItem -LiteralPath "$Folder" -Force:$Hidden | Where-Object { $_.PSIsContainer -eq $true } | Select-Object -ExpandProperty FullName
    foreach ($Child in $Children) {
      Test-Folder -Folder $Child -Recurse:$Recurse -DeleteEmpty:$DeleteEmpty
    }
  }

  if ($ShowDetail) {Write-Host $Folder}

  # If there are any children (files or folders), the test fails
  if ($(Get-ChildItem -LiteralPath "$Folder" -Force)) {
    if ($ShowDetail) {Write-Host "- Has children!"}
    return
  }

  # If the folder is not old enough, the test fails
  if ($ModifiedDate -ge $ExpirationDate) {
    if ($ShowDetail) {Write-Host "- Not old enough!"}
    return
  }

  # If you've reached this point, all tests must have passed!
  if ($DeleteEmpty) {
    if ($ShowDetail) {Write-Host "- REMOVING FOLDER!"}
    Remove-Folder -Folder "$Folder"
  }
}

# If run against a local drive or drive mapping, some additional information is available.
# We will capture that here.
if ($Path -match "^[aA-zZ]:\\") {
  $DeviceID = $Path.Substring(0, 2)
  $TargetDrive = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $DeviceID }
  $TotalVolumeSize = $TargetDrive.Size
  $FreeSpace = $TargetDrive.FreeSpace

  $global:Report += [PSCustomObject]@{
    "Type"      = "Info"
    "Path"        = $DeviceID
    "Size"        = ""
    "SizeInBytes" = ""
    "Created"     = ""
    "Accessed"    = ""
    "Modified"    = ""
    "Action"      = "Report"
    "Detail"      = "{0} free of {1} total ({2:N1}%)" -f (Format-Size $FreeSpace), (Format-Size $TotalVolumeSize), (($FreeSpace / $TotalVolumeSize) * 100)
  }
}

# Calculate the target date based on -Age
$ExpirationDate = (Get-Date).AddDays(-$Age)

# Create a list of all the files that exceed the target age
$OldFiles = Get-ChildItem -Path $Path -Recurse:$Recurse -Force:$Hidden | Where-Object { $_.LastWriteTime -lt $ExpirationDate -and $_.PSIsContainer -eq $false }

# Start iterating through all of the files
foreach ($OldFile in $OldFiles) {
  try {
    Remove-Item -LiteralPath $OldFile.FullName -WhatIf:$TestMode -Confirm:$false
  } catch {
    # Hit an error during the delete. Capture it and immediately loop on to the next file
    $global:Report += [PSCustomObject]@{
      "Type"        = "Error"
      "Path"        = $OldFile.FullName
      "Size"        = ""
      "SizeInBytes" = ""
      "Created"     = ""
      "Accessed"    = ""
      "Modified"    = ""
      "Action"      = "Error"
      "Detail"      = $_.Exception.Message
    }
    continue
  }

  # Deletion must have been successful. Add it to the log.
  $global:Report += [PSCustomObject]@{
    "Type"        = "File"
    "Path"        = $OldFile.FullName
    "Size"        = Format-Size $OldFile.Length
    "SizeInBytes" = $OldFile.Length
    "Created"     = $OldFile.CreationTime
    "Accessed"    = $OldFile.LastAccessTime
    "Modified"    = $OldFile.LastWriteTime
    "Action"      = $global:Action
    "Detail"      = ""
  }

  # Update our total reclaimed space 
  $RecoveredSpace += $OldFile.Length
}

# All appropriate files deleted so let's report how we did.
$global:Report += [PSCustomObject]@{
  "Type"        = "Info"
  "Path"        = ""
  "Size"        = ""
  "SizeInBytes" = ""
  "Created"     = ""
  "Accessed"    = ""
  "Modified"    = ""
  "Action"      = "Report"
  "Detail"      = "{0} recovered ({1:N1}% of the total volume size)" -f (Format-Size $RecoveredSpace), (($RecoveredSpace / $TotalVolumeSize) * 100)
}

# If we deleting empty folders...
if (!$KeepEmptyFolders) {
   # Get a list of all of the first-level folders. The Test-Folder function will handle folders
  # further down if the -Recurse switch was passed
  $Folders = Get-ChildItem -Path "$Path" -Force:$Hidden | Where-Object { $_.PSIsContainer -eq $true } | Select-Object -ExpandProperty FullName
  foreach ($Folder in $Folders) {
    Test-Folder -Folder "$Folder" -DeleteEmpty:$(!$KeepEmptyFolders) -Recurse:$Recurse
  }

  # Folder cleanup has completed. Let's report how we did.
  $global:Report += [PSCustomObject]@{
    "Type"        = "Info"
    "Path"        = ""
    "Size"        = ""
    "SizeInBytes" = ""
    "Created"     = ""
    "Accessed"    = ""
    "Modified"    = ""
    "Action"      = "Report"
    "Detail"      = "{0} empty folders were deleted" -f ($FoldersRemoved)
  }
}

# Generate the log file
$global:Report | Export-Csv ("{0}\cfba_{1}-{2}.csv" -f ($ReportPath), ($ReportFile.ToLower()), ($Now)) -NoTypeInformation