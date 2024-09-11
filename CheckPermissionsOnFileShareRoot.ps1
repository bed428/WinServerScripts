
$Servername = "servername"
$shares = net view \\$Servername | ForEach-Object { $_.Trim() } | Where-Object { $_ -match 'Disk' } | ForEach-Object {
	$_ -split '\s{2,}' | Select-Object -First 1
}


$Allowed = @()
$Denied = @()
foreach ($share in $shares) {
	$sharePath = "\\$Servername\$share"
	try {
    	Test-Path $sharePath -ErrorAction Stop | Out-Null
    	$Allowed += $sharePath    
	} catch {
    	$Denied += $sharePath
	}
}

Write-Host -ForegroundColor Green "##### ALLOWED #####"
foreach($allow in $Allowed){Write-Host -ForegroundColor Green $allow}
Write-Host -ForegroundColor Red `n "##### DENIED #####"
foreach($deny in $denied){Write-Host -ForegroundColor Red $deny}
