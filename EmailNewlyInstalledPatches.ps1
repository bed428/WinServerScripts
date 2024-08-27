$Servers = Get-Content E:\ScheduledTasks\FindNewPatches\TargetServers.txt
$DaysToCheck = -1

$NewEvents = @()
$NewHotfixes = @()

foreach($Server in $Servers){
	$ApplicationLog = @()
	$SetupLog = @()
	if(Test-Path \\$server\C$){
    	Write-Host -ForegroundColor Yellow "Checking $server for Hotfixes and events."
    	$NewHotfixes += Get-HotFix -ComputerName $Server | where {$_.InstalledOn -gt (Get-Date).AddDays($DaysToCheck)}
    	$NumberOfSuccessfulInstalls = $NewHotfixes.Count

    	$ApplicationLog = Get-EventLog -LogName Application -After (Get-Date).AddDays($DaysToCheck) -ComputerName $Server
        	if($ApplicationLog -ne $Null){    
            	$NewEvents += $ApplicationLog | where {($_.InstanceID -eq 1022) -and ($_.Source -eq "MsiInstaller")}
            	$NewEvents += $ApplicationLog | where {($_.InstanceID -eq 1) -and ($_.Message -like "*Initiating an Server/OS update")}
            	$NewEvents += $ApplicationLog | where {($_.InstanceID -eq 2) -and ($_.Message -like "*Server/OS Update changed to Installed State*")}
        	} else {Write-Host -ForegroundColor Gray "No relevant application logs found."}

    	$SetupLog = Get-EventLog -LogName Setup -After (Get-Date).AddDays($DaysToCheck) -ComputerName $Server
        	if($SetupLog -ne $Null){
            	$NewEvents += $SetupLog | where {($_.InstanceId -eq 1022) -and ($_.Message -like "*Application/Product Update Installed*")}
        	} else {Write-Host -ForegroundColor Gray "No relevant SetupLogs found."}

    	$SystemLog = Get-EventLog -LogName System -After (Get-Date).AddDays($DaysToCheck) -ComputerName $Server
        	if($SystemLog -ne $Null){
            	$NewEvents += $SystemLog | where {$_.Source -eq "Microsoft-Windows-WindowsUpdateClient"}
            	$NewEvents += $SystemLog | where {$_.Source -eq "User32"} #Detects reboots

            	$UniqueInstallsStarted = $SystemLog | where {$_.Message -like "*Installation Started*"} | Select-Object -ExpandProperty Message -Unique
            	$NumberOfUniqueInstallsStarted = $UniqueInstallsStarted.Count

        	} else {Write-Host -ForegroundColor Gray "No relevant System Logs found."}
    
	} else {Write-Host -ForegroundColor Red "$server offline!"}
}


$Table = @()

foreach($Hotfix in $NewHotfixes){
	$Table += [PSCustomObject]@{
	ServerName = $Hotfix.PSComputerName
	Description = $Hotfix.Description + " " + $hotfix.HotFixID + " " + $hotfix.InstalledBy + " " + $hotfix.Caption
	ID = $Hotfix.HotFixID
	InstalledOn = $Hotfix.InstalledOn
	DetectionType = "Get-Hotfix"
	}
}

foreach($Event in $NewEvents){
	$Table += [PSCustomObject]@{
	ServerName = $event.MachineName.Split('.')[0]
	Description = $Event.Message
	ID = "MsiInstaller"
	InstalledOn = $event.TimeGenerated
	DetectionType = "Get-Eventlog"
	}
}

#CSV attached to email output + archiving ability.
if ($Table -ne $Null){
$Body = @"
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; color: #333; }
        h2 { color: #2E4053; }
        table { width: 50%; border-collapse: collapse; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #4CAF50; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <h2>Here is a summary of the server status:</h2>
    <table>
        <tr>
            <th>Server</th>
            <th>Started Installations</th>
            <th>Finished Installations</th>
        </tr>
"@

foreach ($Server in $Servers) {
    $Body += "<tr>"
    $Body += "<td>$Server</td>"
    $Body += "<td>$NumberOfUniqueInstallsStarted</td>"
    $Body += "<td>$NumberOfSuccessfulInstalls</td>"
    $Body += "</tr>"
}

$Body += @"
    </table>
    <p>Please find the detailed report attached.</p>
</body>
</html>
"@

	# Convert the table to CSV format
	$CsvFilePath = "E:\ScheduledTasks\FindNewPatches\Reports\NewHotfixesReport_" + (Get-Date -Format yyyy-MM-dd_hhmm) + ".csv"
	$Table | Select-Object ServerName, Description, DetectionType, InstalledOn | Sort-Object -Property ServerName | Export-Csv -Path $CsvFilePath -NoTypeInformation

	# Email parameters
	$MailParams = @{
    	SmtpServer = 'mail.domain.com'
    	From = 'email@domain.com'
    	To = 'email@domain.com'
    	Subject = 'New Hotfixes Installed on Server(s)'
    	Body = $Body
    	Attachments = $CsvFilePath
	}
	Send-MailMessage @MailParams
}
