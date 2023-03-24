<#
Connect-RemoteDesktop.ps1

Arguments
===========================================================
-ComputerName | PC name or IP address of target [REQUIRED]
-ViewOnly     | Switch that will request read-only access to target
-NoConsent    | Will access target without asking/notifying user (if allowed by policy)
-Width        | Sets the Width of the session window
-Height       | Sets the Height of the session window

#>

Param( 
    [Parameter(Mandatory=$true)][string]$ComputerName, 
    [Parameter(Mandatory=$false)][switch]$ViewOnly,
    [Parameter(Mandatory=$false)][switch]$NoConsent,
    [Parameter(Mandatory=$false)][int64]$Width, 
    [Parameter(Mandatory=$false)][int64]$Height
    )

function Get-UserSessions {
    return (query user /server:$ComputerName) -split "\n" -replace '\s\s+', ';' | ConvertFrom-Csv -Delimiter ';'
}

$Sessions = Get-UserSessions

$TargetSession = $Sessions | Out-GridView -Title "Select target session" -PassThru

$CommandLine = "mstsc.exe /v:$ComputerName /shadow:$($TargetSession.ID) "

if (!$ViewOnly) {$CommandLine += "/control "}
if ($NoConsent) {$CommandLine += "/noConsentPrompt "}
if ($Width) {$CommandLine += "/w:$Width "}
if ($Height) {$CommandLine += "/h:$Height "}

Write-Host $CommandLine

cmd /c "$CommandLine"
