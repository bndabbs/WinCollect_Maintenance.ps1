Function CheckProcess
	{
	$Service = $(Get-Service -ComputerName $Computer -Name $ServiceName -ErrorAction SilentlyContinue)
		IF ($Service -eq $null)
			{Write-Host "No service named" $ServiceName "on" $Computer "or the computer could not be reached."}
		ELSE {Write-Host $Service "is" $Service.Status "on" $Computer.ToUpper()}
		
		IF ($Service.Status -eq "Stopped") {StartService} # Call Start function if service is stopped
		ELSEIF	($Service.Status -eq "StopPending")	{KillProcess}	# Forcefully kill the process
		ELSEIF	($Service.Status -eq "Running" -and $count -lt 1)	{KillProcess}
	}
Function StartService
	{
	Write-Host "Starting" $Service.Name "service on" $Computer.ToUpper()
	$Service.Start() 
	Start-Sleep 5
	CheckProcess
	}
Function KillProcess 
	{
	$Processes = (Get-WmiObject Win32_Process -ComputerName $Computer | Where { $_.ProcessName -match $ProcesesName })
	ForEach ($Process IN $Processes)
		{
		Write-Host "Killing" $Process.ProcessName "on" $Computer.ToUpper()
		$Process.Terminate() | Out-Null 
		Start-Sleep 5 
		}
	$count++
	CheckProcess
	}

## Start of script ##
$ServiceName = "WinCollect"
$ProcesesName = "WinCollect*"
$Computers = Import-Csv C:\users\bdabbs\Desktop\00001687.csv -Header Name
ForEach ($Computer IN $Computers)
	{
	$Computer = $Computer.Name
	$count = 0
	CheckProcess
	}