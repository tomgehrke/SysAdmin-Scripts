# Get-DiskUsage.ps1
#
# Arguments
# ===================
# -Path       | 

Param( 
    [Parameter(Position=0,Mandatory=$false)][string]$StartingPath
)

Add-Type -AssemblyName System.Windows.Forms

# A path was not provided on the command line. Let the user select one.
if ($StartingPath -eq "") {
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = "Select the root folder"
    $FolderDialog.RootFolder = "MyComputer"
    $FolderDialog.SelectedPath = $(Get-Location).StartingPath
    $FolderDialog.ShowNewFolderButton = $false
    $ShowForm = $FolderDialog.ShowDialog()
    if ($ShowForm -eq "OK") {
        $StartingPath = $FolderDialog.SelectedPath
    } else {
        Write-Output "[ERROR!] A target path must be selected!"
        Exit
    }
}

# Make sure the path exists. Exit if it does not.
if (!(Test-Path $StartingPath -PathType Container)) {
    Write-Output "[ERROR] ""$StartingPath"" does not exist!"
    Exit
}

$Report = @()

# Recursively populate the collection of folders in the target folder's tree
Write-Output "- Loading directory tree..."
$Folders = Get-ChildItem $StartingPath -Directory
$FolderCount = $Folders.Count
$CurrentFolder = 0

# Loop through all of the folders calling our Backup-Acl function for each one
foreach ($Folder in $Folders) {
    $CurrentFolder++
    Write-Progress -Id 0 -Activity "Scanning Folders" -PercentComplete ($CurrentFolder/$FolderCount*100) -Status "Scanning $($Folder.FullName)"
   
    $Report += [PSCustomObject]@{
        "Folder" = $Folder.FullName
        "Size" = $(Get-ChildItem $Folder.FullName -Force -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    }
}

$SelectedFolder = $Report | Sort-Object -Property Size -Descending | Select-Object Folder, @{n='Size';e={if ($_.Size -ge 1GB) {"{0:N2} GB" -f ($_.Size/1GB)} elseif ($_.Size -ge 1MB) {"{0:N2} MB" -f ($_.Size/1MB)} elseif ($_.Size -ge 1KB) {"{0:N2} KB" -f ($_.Size/1KB)} else {"{0}" -f $_.Size}}} | Out-GridView -Title "Tom's Disk Usage" -PassThru
if ($SelectedFolder) {
    .\Get-DiskUsage -StartingPath $SelectedFolder.Folder
}