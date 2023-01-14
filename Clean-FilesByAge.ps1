# Clean-FilesByAge.ps1
#
# Arguments
# ===================
# -Path <path>         | 
# -ReportFolder <path> |
# -Age <n>             | 
# -Recurse             |
# -Hidden              |
# -KeepEmptyFolders    |
# -TestMode            | 

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

# Check for backup folder if provided and create if it doesn't exist
if ($ReportPath -and !(Test-Path -Path $ReportPath)) {
  Write-Host "- Backup folder does not exist. Creating it..."
  $null = New-Item -Path $ReportPath -ItemType Directory -Force
}

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

function Test-Folder {
  Param (
    [Parameter(Mandatory = $true)]$Folder
  )

  if ($Recurse) {
    $Children = Get-ChildItem -LiteralPath $Folder.FullName -Force:$Hidden | Where-Object { $_.PSIsContainer -eq $true }
    foreach ($Child in $Children) {
      if ($(Test-Folder -Folder $Child).value) {
        Remove-Folder -Folder $Child
      }
    }
  }

  if ($(Get-ChildItem -LiteralPath $Folder.FullName -Force)) {
    return $false
  }

  if ($Folder.LastWriteTime -ge $ExpirationDate) {
    return $false
  }

  return $true
}

$Now = Get-Date -Format "yyyyMMddHHmm"
$ReportFile = $Path.TrimEnd('\\') -Replace '\\','-' -Replace ' ','_' -Replace '[^\w-_]',''
$Report = @()

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

$ExpirationDate = (Get-Date).AddDays(-$Age)
$OldFiles = Get-ChildItem -Path $Path -Recurse:$Recurse -Force:$Hidden | Where-Object { $_.LastWriteTime -lt $ExpirationDate -and $_.PSIsContainer -eq $false }
$RecoveredSpace = 0

foreach ($OldFile in $OldFiles) {
  try {
    Remove-Item -LiteralPath $OldFile.FullName -WhatIf:$TestMode -Confirm:$false
  } catch {
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

  $RecoveredSpace += $OldFile.Length
}

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

if (!$KeepEmptyFolders) {
  $FoldersRemoved = 0

  $Folders = Get-ChildItem -Path $Path -Force:$Hidden | Where-Object { $_.PSIsContainer -eq $true }
  foreach ($Folder in $Folders) {
    if (Test-Folder -Folder $Folder) {
      if (Remove-Folder -EmptyFolder $Folder) { $FoldersRemoved ++ }
    }
  }

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

$Report | Export-Csv ("{0}\cfba_{1}-{2}.csv" -f ($ReportPath), ($ReportFile.ToLower()), ($Now)) -NoTypeInformation