# Get-UsbDeviceStatus

[CmdletBinding()]
param (
    [Parameter()][string]$Name="",
    [Parameter()][string]$DeviceId="",
    [Parameter()][double]$ScanInterval=60,
    [Parameter()][switch]$AlarmOnUnknown,
    [Parameter()][double]$AlarmInterval=10
)

Add-Type -AssemblyName System.Windows.Forms 

$NotificationIcon = New-Object System.Windows.Forms.NotifyIcon
$ProcessPath = (Get-Process -id $pid).Path
$NotificationIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ProcessPath) 

$Devices = Get-PnpDevice | Where-Object {$_.InstanceId -match '^USB' -and $_.Name -match "$Name" -and $_.DeviceID -match "$($DeviceId.Replace("\", "\\"))"}

if ($Devices) {
    if ($Devices.Count -gt 1) {
        $Devices | Select-Object Status, Name, Class, DeviceID, Present, ProblemDescription | Sort-Object Name | Format-Table
    } else {
        $LastStatus = ""
        
        while ($true) {
            if ($LastStatus -ne $Devices.Status -or ($AlarmOnUnknown -and $LastStatus -eq "Unknown") ) {
                $LastStatus = $Devices.Status

                if ($AlarmOnUnknown -and $LastStatus -eq "Unknown") {
                    $NotificationIcon.BalloonTipTitle = "USB Device Unplugged!" 
                    $NotificationIcon.BalloonTipText = """$($Devices.Name)"" has been unplugged!"
                    $NotificationIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Error
                } else {
                    $NotificationIcon.BalloonTipTitle = "USB Device Status Change!" 
                    $NotificationIcon.BalloonTipText = "Status of ""$($Devices.Name)"" is now ""$($Devices.Status)""!"
                    $NotificationIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
                }

                $NotificationIcon.Visible = $true 
                $NotificationIcon.ShowBalloonTip(5000)
            }
            if ($AlarmOnUnknown -and $LastStatus -eq "Unknown") {
                Start-Sleep -Seconds $AlarmInterval
            } else {
                Start-Sleep -Seconds $ScanInterval
            }
            $Devices = Get-PnpDevice | Where-Object {$_.InstanceId -match '^USB' -and $_.Name -match "$Name" -and $_.DeviceID -match "$($DeviceId.Replace("\", "\\"))"}
        }
    }
} else {
    Write-Output "Device ""$Name"" not found!"
}
