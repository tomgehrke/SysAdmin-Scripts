<# 
Organize-FilesByDate.ps1

Arguments
===================
-Path <path>         | Source path that the script will operate on
-Destination <path>  | Destination root path for files
-Filemask <mask>     | Filemask to filter files to be moved
-Recurse             | Cause the script to recursively search subfolders for files
-TestMode            | Cause the script to only report on the files and folders it would have moved
-ShowDetail          | Displays processing notes in console
#>

Param( 
  [Parameter(Mandatory = $false)][string]$Path = $(Get-Location).Path, 
  [Parameter(Mandatory = $false)][string]$Destination = $(Get-Location).Path, 
  [Parameter(Mandatory = $false)][string]$Filemask = "*.*", 
  [Parameter(Mandatory = $false)][switch]$Recurse,
  [Parameter(Mandatory = $false)][switch]$TestMode
)

# Make sure the source path exists. Exit if it does not.
if ($Path -and !(Test-Path $Path -PathType Container)) {
    Write-Output "[ERROR] ""$Path"" does not exist!"
    Exit
  }

# Test for destinate path and create if it does not exist.
if ($Destination -and !(Test-Path -Path $Destination)) {
    Write-Host "- Destination folder does not exist. Creating it..."
    $null = New-Item -Path $Destination -ItemType Directory -Force
  }

# Get all files in the source folder
$Files = Get-ChildItem $Filemask -Path $Path -File -Recurse:$Recurse

# Iterate through each file
foreach ($File in $Files) {
    # Get the creation date of the file
    $ModifiedDate = $File.LastWriteTime

    # Create the year and month folders if they don't exist
    $YearFolder = Join-Path -Path $Destination -ChildPath $ModifiedDate.Year.ToString()
    $MonthFolder = Join-Path -Path $YearFolder -ChildPath ("{0:D2}" -f $ModifiedDate.Month)

    if (-not (Test-Path $YearFolder)) {
        $null = New-Item -ItemType Directory -Path $YearFolder
    }

    if (-not (Test-Path $MonthFolder)) {
        $null = New-Item -ItemType Directory -Path $MonthFolder
    }

    # Move the file to the destination folder
    $DestinationPath = Join-Path -Path $MonthFolder -ChildPath $File.Name

    Write-Host "Moving ""$File"" to $DestinationPath..."
    Move-Item -Path $File.FullName -Destination $DestinationPath -Force -WhatIf:$TestMode
}

Write-Host "Files have been organized into folders based on their modified dates."
