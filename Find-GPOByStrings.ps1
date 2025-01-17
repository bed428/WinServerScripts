$SearchStrings = ("Outbound", "636", "tcp", "ldap")
$SearchPattern = $SearchStrings -join "|"

$GPOs = Get-GPO -All

$SearchResults = @()
foreach ($GPO in $GPOs){
    $GPOReport = Get-GPOReport -Guid $GPO.id -reporttype xml | out-string
    $Matches = ($GPOReport -split "`n") | Where-Object { $_ -match $SearchPattern }

    foreach($Match in $Matches){
        $SearchResults += [PSCustomObject]@{
            GPOName = $GPO.DisplayName
            GPOGuid = $gpo.Id
            SearchMatch = $Match
        }
        Remove-Variable Match -ErrorAction SilentlyContinue #prevents loops from registering twice if one is null? 
    }
}
