# Fix-UnquotedServicePaths

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)][string]$Computer="."
)

$Services = Get-WmiObject -ComputerName $Computer -Class Win32_Service | Where-Object { $_.PathName -ne $null -and $_.PathName -match " " -and $_.PathName -inotmatch "`"" -and $_.PathName -inotmatch " -" -and $_.PathName -inotmatch " \\" -and $_.PathName -inotmatch " /" };
$Services | Sort-Object -Property Name | Format-Table -Property ProcessId, Name, DisplayName, PathName, StartName, StartMode, State;

foreach ($Service in $Services) {
    $CommandLine = "sc.exe \\$Computer config ""$($Service.Name)"" binpath= ""\""$($Service.PathName)\"""""
    cmd /c $CommandLine
}