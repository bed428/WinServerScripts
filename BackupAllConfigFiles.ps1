#Backs up all server's web.config files to a network share.
	$RunningJobTemp = "E:\WebConfigBackups\Backups\RunningJob"
	Mkdir $RunningJobTemp

#Archive Directory to store the final backup zip.
	$archivedir = "E:\WebConfigBackups\Backups"

#Server list of servers to be checked. 1 server per line no indents or formatting.
	$servers = Get-Content "C:\Path\WebConfigBackupTargets.txt"

#Backs up web.configs from every server in $servers.
	foreach($server in $servers){
    	Invoke-Command -ComputerName $server -AsJob -ScriptBlock {
        	$Server = $env:COMPUTERNAME
        	$RunningJobTemp = "\\SERVER\E$\WebConfigBackups\Backups\RunningJob"
        	$testpathC = Test-Path "C:\"
        	$testpathE = Test-Path "E:\"
        	$AllConfigs = @()

        	if($testpathC -eq "True"){
            	mkdir $RunningJobTemp\$server -ErrorAction SilentlyContinue
            	mkdir $RunningJobTemp\$server\C
            	$AllConfigs += Get-ChildItem -Path "C:\" -Filter "*.config" -Recurse -erroraction SilentlyContinue
       	 
        	}
        	if($testpathE -eq "True"){
            	mkdir $RunningJobTemp\$server -ErrorAction SilentlyContinue
            	mkdir $RunningJobTemp\$server\E
            	$AllConfigs += Get-ChildItem -Path "E:\" -Filter "*.config" -Recurse -erroraction SilentlyContinue
            	#Copy-Item -Path "E:\" -Filter "*.config" -Container -Destination $RunningJobTemp\$server\E -Recurse -Force
        	}

        	#CopyTheItems
            	foreach($Config in $AllConfigs){
                	$destinationdir = "$RunningJobTemp\$server\$($config.PSDrive.Name)$($config.DirectoryName.Split(':')[1])"
                	$newDirParams = @{
                    	Path = $destinationdir
                    	ItemType = 'Directory'
                    	Force = $TRUE
                	}
                	New-Item @NewDirParams -ErrorAction SilentlyContinue

                	Copy-Item -Path ($Config.FullName) -Destination $destinationdir -Force
            	}
    	}
	}

	Get-Job | Wait-Job | Receive-Job #| Remove-Job
	Write-Host -ForegroundColor Green "ALL JOBS FINISHED. STARTING COMPRESSION"

#Compresses the archive to save storage space and for organization.
	$date = get-date -format MM_dd_yyyy_HHmm
	Compress-Archive -Path $RunningJobTemp\* -DestinationPath $archivedir\$date.zip -CompressionLevel Fastest

#Remove non-archive files.
	Get-Item -Path "$RunningJobTemp" | Remove-Item -Recurse -Force -Confirm:$false

#cleans up backups >60 days old.
	#Get-ChildItem -Path $archivedir -Recurse | Where-Object CreationTime -lt (Get-Date).AddDays(-60) | Remove-Item
