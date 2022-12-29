Param( 
    [Parameter(Mandatory=$false)]
    [int]
    $DaysAgo=30
    )

$TargetDate = [DateTime]::Today.AddDays(-$DaysAgo)
Get-ADComputer -Filter 'WhenCreated -ge $TargetDate' -Properties whenCreated 