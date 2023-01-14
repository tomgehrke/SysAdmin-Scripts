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
  [Parameter(Mandatory = $false)][switch]$TestMode
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
    [Parameter(Mandatory = $true)]$EmptyFolder
  )  

  try {
    Remove-Item -LiteralPath $EmptyFolder.FullName -Force:$Hidden -WhatIf:$TestMode -Confirm:$false
  } catch {
    $Report += [PSCustomObject]@{
      "Type"   = "Error"
      "Path"     = $EmptyFolder.FullName
      "Size"     = ""
      "Created"  = ""
      "Accessed" = ""
      "Modified" = ""
      "Action"   = "Error"
      "Detail"   = $Error[0].Exception.Message
    }
    return $false
  }

  $Report += [PSCustomObject]@{
    "Type"   = "Folder"
    "Path"     = $EmptyFolder.FullName
    "Size"     = ""
    "Created"  = $EmptyFolder.CreationTime
    "Accessed" = $EmptyFolder.LastAccessTime
    "Modified" = $EmptyFolder.LastWriteTime
    "Action"   = if ($TestMode) { "Report" } else { "Delete" }
  }

  return $true
}

# Function determines if a folder can be deleted by checking whether it has any children (folders or files)
# and if it passes the age requirement
function Test-Folder {
  Param (
    [Parameter(Mandatory = $true)]$Folder
  )

  # If the script was run with the -Recurse switch, we will drill down into all the child folders,
  # test them for "deletability" and then delete them if they pass.
  if ($Recurse) {
    $Children = Get-ChildItem -LiteralPath $Folder.FullName -Force:$Hidden | Where-Object { $_.PSIsContainer -eq $true }
    foreach ($Child in $Children) {
      if ($(Test-Folder -Folder $Child).value) {
        Remove-Folder -Folder $Child
      }
    }
  }

  # If there are any children (files or folders), the test fails
  if ($(Get-ChildItem -LiteralPath $Folder.FullName -Force)) {
    return $false
  }

  # If the folder is not old enough, the test fails
  if ($Folder.LastWriteTime -ge $ExpirationDate) {
    return $false
  }

  # If you've reached this point, all tests must have passed!
  return $true
}

$Now = Get-Date -Format "yyyyMMddHHmm"
$ReportFile = $Path.TrimEnd('\\') -Replace '\\','-' -Replace ' ','_' -Replace '[^\w-_]',''
$Report = @()

# If run against a local drive or drive mapping, some additional information is available.
# We will capture that here.
if ($Path -match "^[aA-zZ]:\\") {
  $DeviceID = $Path.Substring(0, 2)
  $TargetDrive = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $DeviceID }
  $TotalVolumeSize = $TargetDrive.Size
  $FreeSpace = $TargetDrive.FreeSpace

  $Report += [PSCustomObject]@{
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

# Initialize the variable to store the amount of space successfully reclaimed
$RecoveredSpace = 0

# Start iterating through all of the files
foreach ($OldFile in $OldFiles) {
  try {
    Remove-Item -LiteralPath $OldFile.FullName -WhatIf:$TestMode -Confirm:$false
  } catch {
    # Hit an error during the delete. Capture it and immediately loop on to the next file
    $Report += [PSCustomObject]@{
      "Type"      = "Error"
      "Path"        = $OldFile.FullName
      "Size"        = ""
      "SizeInBytes" = ""
      "Created"     = ""
      "Accessed"    = ""
      "Modified"    = ""
      "Action"      = "Error"
      "Detail"      = $Error[0].Exception.Message
    }
    continue
  }

  # Deletion must have been successful. Add it to the log.
  $Report += [PSCustomObject]@{
    "Type"      = "File"
    "Path"        = $OldFile.FullName
    "Size"        = Format-Size $OldFile.Length
    "SizeInBytes" = $OldFile.Length
    "Created"     = $OldFile.CreationTime
    "Accessed"    = $OldFile.LastAccessTime
    "Modified"    = $OldFile.LastWriteTime
    "Action"      = if ($TestMode) { "Report" } else { "Delete" }
    "Detail"      = ""
  }

  # Update our total reclaimed space 
  $RecoveredSpace += $OldFile.Length
}

# All appropriate files deleted so let's report how we did.
$Report += [PSCustomObject]@{
  "Type"      = "Info"
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
  # Initialize variable to keep track of how many empty folders we end up deleting
  $FoldersRemoved = 0

  # Get a list of all of the first-level folders. The Test-Folder function will handle folders
  # further down if the -Recurse switch was passed
  $Folders = Get-ChildItem -Path $Path -Force:$Hidden | Where-Object { $_.PSIsContainer -eq $true }
  foreach ($Folder in $Folders) {
    if (Test-Folder -Folder $Folder) {
      # If the folder passed the tests, delete it. If successfully deleted, increment our counter.
      if (Remove-Folder -EmptyFolder $Folder) { $FoldersRemoved ++ }
    }
  }

  # Folder cleanup has completed. Let's report how we did.
  $Report += [PSCustomObject]@{
    "Type"      = "Info"
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
$Report | Export-Csv ("{0}\cfba_{1}-{2}.csv" -f ($ReportPath), ($ReportFile.ToLower()), ($Now)) -NoTypeInformation