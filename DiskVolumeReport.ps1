$Servers = Get-Content ".txt file of servers 1 per line"

$AllServerVolumes = @()

foreach($server in $servers){
    $Volumes = @()

    if($Server -ne $env:COMPUTERNAME){
        $Volumes = Invoke-Command -ComputerName $server {
            $Volumes = Get-Volume | where {$_.DriveLetter -and $_.FileSystem -eq 'NTFS'}
            $Volumes
        }
    }

    else {
        $Volumes = Get-Volume | where {$_.DriveLetter -and $_.FileSystem -eq 'NTFS'}
        $Volumes | ForEach-Object { $_ | Add-Member -NotePropertyName PSComputerName -NotePropertyValue $env:COMPUTERNAME -Force }
    } #invoke-command doesn't work on the executing server if not admin, so if it's checking itself - have to do it without that + add the pscomputername. 

    $allservervolumes += $Volumes

}
 

# Start building the HTML report
$HtmlReport = @"
<html>
<head>
    <style>
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 8px; text-align: left; border: 1px solid black; }
        th { background-color: #f2f2f2; }
        .healthy { background-color: green; color: white; }
        .unhealthy { background-color: red; color: white; }
        .low-space { background-color: red; color: white; }
        .percent-green { background-color: green; color: white; }
        .percent-yellow { background-color: yellow; color: black; }
        .percent-red { background-color: red; color: white; }
    </style>
</head>
<body>
    <h2>Volume Report</h2>
    <table>
        <tr>
            <th>PSComputerName</th>
            <th>HealthStatus</th>
            <th>DriveLetter</th>
            <th>FileSystemLabel</th>
            <th>Size (GB)</th>
            <th>SizeRemaining (GB)</th>
            <th>PercentRemaining</th>
        </tr>
"@


foreach($volume in $AllServerVolumes){
    # Calculate percentage remaining
    $PercentRemaining = (($volume.SizeRemaining / $volume.Size) * 100)


    # Determine HealthStatus class
    $healthClass = if ($volume.HealthStatus -eq 'Healthy') { 'healthy' } else { 'unhealthy' }


    # Determine PercentRemaining class
    if ($PercentRemaining -lt 80) {
        $percentClass = 'percent-green'
    } elseif ($PercentRemaining -ge 80 -and $PercentRemaining -le 90) {
        $percentClass = 'percent-yellow'
    } else {
        $percentClass = 'percent-red'
    }


    # Determine SizeRemaining class (if below 5GB)
    $sizeRemainingClass = if ($volume.SizeRemaining -lt 5GB) { 'low-space' } else { '' }


    # Append row to the report
    $HtmlReport += @"
        <tr>
            <td>$($volume.PSComputerName)</td>
            <td class='$healthClass'>$($volume.HealthStatus)</td>
            <td>$($volume.DriveLetter)</td>
            <td>$($volume.FileSystemLabel)</td>
            <td>$([math]::Round($volume.Size / 1GB, 2))</td>
            <td class='$sizeRemainingClass'>$([math]::Round($volume.SizeRemaining / 1GB, 2))</td>
            <td class='$percentClass'>$([math]::Round($PercentRemaining, 2))%</td>
        </tr>
"@
}

# Close the HTML tags
$HtmlReport += @"
    </table>
</body>
</html>
"@

# Email with HTML Report as body.
    $emailFrom = "fromemail@domain.com"
    $emailTo = @(
        "toemail1domain.com",
        "toemail2@domain.com"
    )
    $smtpServer = "smtpserver.domain.com"
    Send-MailMessage -From $emailFrom -To $emailTo -Subject "Disk Volume Report" -Body $HtmlReport -SmtpServer $smtpServer -BodyAsHtml


