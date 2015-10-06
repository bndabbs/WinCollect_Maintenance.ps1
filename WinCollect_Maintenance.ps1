Add-Type -assembly "system.io.compression.filesystem" # This .NET assembly is required for the commands on lines 23 & 25
$Computers = Import-Csv C:\users\bdabbs\Desktop\hostToGather.txt -Header Name

ForEach ($Computer IN $Computers)
	{
	$Computer = $Computer.Name
	$Root = "\\$Computer\C$\IBM\WinCollect"
	$ServiceName = "WinCollect"
	$ProcesesName = "WinCollect*"
	$Service = $(Get-Service -ComputerName $Computer -Name $ServiceName -ErrorAction SilentlyContinue)
	
	# Kill WinCollect to prevent file locking errors
	$Processes = (Get-WmiObject Win32_Process -ComputerName $Computer | Where { $_.ProcessName -match $ProcesesName })
		ForEach ($Process IN $Processes)
			{
			$Process.Terminate() | Out-Null 
			}
			
	# Gather and compress remote files
	# File extension is .piz because Exchange will strip .zip files
	New-PSDrive -Name P -PSProvider Filesystem -Root $Root -Scope Global
		If(Test-path P:\config.piz) {Remove-item P:\config.piz}
		[io.compression.zipfile]::CreateFromDirectory("$Root\config", "$Root\config.piz")
		If(Test-path P:\logs.piz) {Remove-item P:\logs.piz}
		[io.compression.zipfile]::CreateFromDirectory("$Root\logs", "$Root\logs.piz")
	$File1 = "P:\config\install_config.txt" # This is already in the config directory that we are compressing, but the requester wants it sent this way
	$File2 = "P:\logs.piz" 
	$File3 = "P:\config.piz"
	
	# Send results
	Send-MailMessage `
	-Attachments $File1, $File2, $File3 `
	-Body "Please see attached WinCollect diagnostic information" `
	-From "you@example.com" `
	-To "someone@example.com" `
	-Cc "someone2@example.com" `
	-Subject "$Computer Wincollect Dump" `
	-SmtpServer "addyour.smtpserver.here"
	
	# Cleanup
	Remove-Item P:\logs.piz -Force
	Remove-Item P:\config.piz -Force
	$Service.Start()
	Remove-PSDrive -Name P
	}