# Export-DNS
#
# Arguments
# ===================
# -Server        | Target DNS or Domain Controller [REQUIRED]
# -ExportFolder  | Folder for storing dns export files

Param( 
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$false)][string]$ExportFolder="."
    )

# Check for backup folder if provided and create if it doesn't exist
if ($ExportFolder -ne "" -and !(Test-Path -Path $ExportFolder)) {
    Write-Host "Export folder does not exist. Creating it..."
    $null = New-Item -Path $ExportFolder -ItemType Directory -Force
}

$Zones = Get-DnsServerZone -ComputerName $Server
$TotalZones = $Zones.Count
$Now = Get-Date -Format "yyyyMMddHHmm"

$CurrentZone = 0

# Go through the list of zones
foreach ($Zone in $Zones) {
    $CurrentZone++;
    Write-Progress -Id 0 -Activity "Exporting Zones" -PercentComplete ($CurrentZone/$TotalZones*100) -Status "Exporting $($Zone.ZoneName)"

    if ($Zone.ZoneName -eq "") {
        Continue
    }

    $ExportPath = "$ExportFolder\$($Zone.ZoneName -Replace ('\W',''))-$Now.dbak"
    Get-DnsServerResourceRecord -ComputerName $Server -ZoneName $Zone.ZoneName | Select-Object -ExpandProperty RecordData -Property HostName, RecordType | Export-Csv $ExportPath -NoTypeInformation
}
