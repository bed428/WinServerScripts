$Servers = Get-Content "path"

$AllServerVolumes = @()
$AllMountPointVolumes = @()

# Simple function to retrieve mount-pointed NTFS volumes
function Get-MountPointVolumes {
	param (
    	[string]$Server
	)

	if ($Server -ne $env:COMPUTERNAME) {
    	return Invoke-Command -ComputerName $Server {
        	Get-CimInstance Win32_Volume | Where-Object {
            	-not $_.DriveLetter -and           	# Exclude volumes with drive letters
            	$_.FileSystem -eq 'NTFS' -and      	# Include only NTFS file systems
            	$_.Label -ne "System Reserved" -and	# Exclude "System Reserved" volumes
            	$_.Name -notmatch '^\\\\\?\\Volume\{[A-F0-9\-]+\}\\$' # Exclude Volume GUID paths
        	} | Select-Object `
            	@{Name='PSComputerName'; Expression={$env:COMPUTERNAME}},  # Add server name
            	@{Name='MountPath'; Expression={$_.Name}},             	# Mount path
            	@{Name='Size'; Expression={$_.Capacity}},              	# Total size in bytes
            	@{Name='SizeRemaining'; Expression={$_.FreeSpace}}      	# Free space in bytes
    	}
	}
	else {
    	$volumes = Get-CimInstance Win32_Volume | Where-Object {
        	-not $_.DriveLetter -and
        	$_.FileSystem -eq 'NTFS' -and
        	$_.Label -ne "System Reserved" -and
        	$_.Name -notmatch '^\\\\\?\\Volume\{[A-F0-9\-]+\}\\$'
    	}
    	$volumes | Select-Object `
        	@{Name='PSComputerName'; Expression={$env:COMPUTERNAME}},
        	@{Name='MountPath'; Expression={$_.Name}},
        	@{Name='Size'; Expression={$_.Capacity}},
        	@{Name='SizeRemaining'; Expression={$_.FreeSpace}}
	}
}

foreach ($server in $servers) {
	$Volumes = @()

	if ($server -ne $env:COMPUTERNAME) {
    	$Volumes = Invoke-Command -ComputerName $server {
        	Get-Volume | Where-Object {
            	($_.DriveLetter) -and
            	($_.FileSystem -eq 'NTFS') -and
            	($_.FileSystemLabel -notlike "*Page*") -and
            	($_.FileSystemLabel -notlike "*Mounts*")
        	}
    	}
	}
	else {
    	$Volumes = Get-Volume | Where-Object {
        	($_.DriveLetter) -and
        	($_.FileSystem -eq 'NTFS') -and
        	($_.FileSystemLabel -notlike "*Page*") -and
        	($_.FileSystemLabel -notlike "*Mounts*")
    	}
    	$Volumes | ForEach-Object { $_ | Add-Member -NotePropertyName PSComputerName -NotePropertyValue $env:COMPUTERNAME -Force }
	}

	$AllServerVolumes += $Volumes

	# Retrieve mount-pointed NTFS volumes for the current server
	$mountVolumes = Get-MountPointVolumes -Server $server
	$AllMountPointVolumes += $mountVolumes
}

# Function to calculate PercentUsed and determine classes
function Process-Volumes {
	param (
    	[array]$Volumes,
    	[string]$Type  # "MountPoint" or "Volume"
	)

	foreach ($vol in $Volumes) {
    	# Calculate PercentUsed
    	$PercentUsed = (($vol.Size - $vol.SizeRemaining) / $vol.Size) * 100
    	$vol | Add-Member -NotePropertyName PercentUsed -NotePropertyValue $PercentUsed -Force

    	# Determine PercentUsed class
    	if ($PercentUsed -gt 90) {
        	$percentClass = 'percent-red'
    	}
    	elseif ($PercentUsed -gt 80) {
        	$percentClass = 'percent-yellow'
    	}
    	else {
        	$percentClass = 'percent-green'
    	}

    	$vol | Add-Member -NotePropertyName PercentClass -NotePropertyValue $percentClass -Force

    	# Determine SizeRemaining class
    	if ($vol.SizeRemaining -lt 5GB) {
        	$sizeRemainingClass = 'low-space'
    	}
    	elseif ($vol.SizeRemaining -lt 10GB) {
        	$sizeRemainingClass = 'medium-space'
    	}
    	else {
        	$sizeRemainingClass = ''
    	}

    	$vol | Add-Member -NotePropertyName SizeRemainingClass -NotePropertyValue $sizeRemainingClass -Force
	}

	return $Volumes
}

# Process and sort Mount Point Volumes
$AllMountPointVolumes = Process-Volumes -Volumes $AllMountPointVolumes -Type "MountPoint"

$SortedMountPointVolumes = $AllMountPointVolumes | Sort-Object -Property `
	@{Expression = {
    	switch ($_.PercentClass) {
        	'percent-red' { 1 }
        	'percent-yellow' { 2 }
        	'percent-green' { 3 }
        	default { 4 }
    	}
	}}, PSComputerName

# Process and sort Server Volumes
$AllServerVolumes = Process-Volumes -Volumes $AllServerVolumes -Type "Volume"

$SortedServerVolumes = $AllServerVolumes | Sort-Object -Property `
	@{Expression = {
    	switch ($_.PercentClass) {
        	'percent-red' { 1 }
        	'percent-yellow' { 2 }
        	'percent-green' { 3 }
        	default { 4 }
    	}
	}}, PSComputerName

# Start building the HTML report
$HtmlReport = @"
<html>
<head>
	<style>
    	table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    	th, td { padding: 8px; text-align: left; border: 1px solid black; }
    	th { background-color: #f2f2f2; }
    	.healthy { background-color: green; color: white; }
    	.unhealthy { background-color: red; color: white; }
    	.low-space { background-color: red; color: white; }
    	.medium-space { background-color: yellow; color: black; }
    	.percent-green { background-color: green; color: white; }
    	.percent-yellow { background-color: yellow; color: black; }
    	.percent-red { background-color: red; color: white; }
	</style>
</head>
<body>
	<h2>Mount Points Report</h2>
	<table>
    	<tr>
        	<th>PSComputerName</th>
        	<th>MountPath</th>
        	<th>Size (GB)</th>
        	<th>SizeRemaining (GB)</th>
        	<th>PercentUsed</th>
    	</tr>
"@

foreach ($mount in $SortedMountPointVolumes) {
	# Append row to the mount points table
	$HtmlReport += @"
    	<tr>
        	<td>$($mount.PSComputerName)</td>
        	<td>$($mount.MountPath)</td>
        	<td>$([math]::Round($mount.Size / 1GB, 2))</td>
        	<td class='$($mount.SizeRemainingClass)'>$([math]::Round($mount.SizeRemaining / 1GB, 2))</td>
        	<td class='$($mount.PercentClass)'>$([math]::Round($mount.PercentUsed, 2))%</td>
    	</tr>
"@
}

# Close the mount points table
$HtmlReport += @"
	</table>

	<h2>Volume Report</h2>
	<table>
    	<tr>
        	<th>PSComputerName</th>
        	<th>HealthStatus</th>
        	<th>DriveLetter</th>
        	<th>FileSystemLabel</th>
        	<th>Size (GB)</th>
        	<th>SizeRemaining (GB)</th>
        	<th>PercentUsed</th>
    	</tr>
"@

foreach ($volume in $SortedServerVolumes) {
	# Determine HealthStatus class
	$healthClass = if ($volume.HealthStatus -eq 'Healthy') { 'healthy' } else { 'unhealthy' }

	# Append row to the volume report table
	$HtmlReport += @"
    	<tr>
        	<td>$($volume.PSComputerName)</td>
        	<td class='$healthClass'>$($volume.HealthStatus)</td>
        	<td>$($volume.DriveLetter)</td>
        	<td>$($volume.FileSystemLabel)</td>
        	<td>$([math]::Round($volume.Size / 1GB, 2))</td>
        	<td class='$($volume.SizeRemainingClass)'>$([math]::Round($volume.SizeRemaining / 1GB, 2))</td>
        	<td class='$($volume.PercentClass)'>$([math]::Round($volume.PercentUsed, 2))%</td>
    	</tr>
"@
}

# Close the volume report table and HTML tags
$HtmlReport += @"
	</table>
</body>
</html>
"@

# Email with HTML Report as body.
#$HtmlReport | Out-File "E:\ADM\Dupy\Scripts\VolumeReport.html" -Encoding utf8

$emailFrom = "NoReply@domain.com"
$emailTo = @(
	"first.last@domain.com"#,
	#"first.last2@domain.com"
)
$smtpServer = "mxrelay.doe.gov"
Send-MailMessage -From $emailFrom -To $emailTo -Subject "Disk Report" -Body $HtmlReport -SmtpServer $smtpServer -BodyAsHtml
