
	# [System.Console]::Beep(196,1200);
	# [System.Console]::Beep(196,300);
	# [System.Console]::Beep(220,400);

    #Write-Host -NoNewLine 'This application adds the H2 printer directly to this device. Would you like to continue?'
	#$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
	param ($DRIVERLOCATION)
	param ($DELETEOLD)

	$DRIVERNAME="KONICA MINOLTA C360iSeriesPCL"

	if ($DRIVERLOCATION -eq $null) {
		throw "No driver location was provided" 
	}




	Write-Host "Do not close window, window will close itself."
	try {
		Remove-Printer -Name "H2 Direct"
		Remove-PrinterPort -Name "h2direct10.10.10.5"
	}
	catch {}
	try {
				if ( $DELETEOLD -match 'y') {
			Remove-Printer -Name "\\print.westpoint.edu\WPMC6021STH2"
		}
	}
	catch {}
	
	Invoke-Command {PNPUtil.exe /add-driver $DRIVERLOCATION /install}
	Add-PrinterDriver -Name $DRIVERNAME

	$SLEEPY=4
	Add-PrinterPort -Name "h2direct10.10.10.5" -PrinterHostAddress "10.10.10.5"
	Start-Sleep $SLEEPY
	Add-Printer -Name "H2 Direct" -PortName "h2direct10.10.10.5" -DriverName $DRIVERNAME

	Write-Output $DRIVERLOCATION
	Write-Output "DONE!"


