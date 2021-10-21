
	# [System.Console]::Beep(196,1200);
	# [System.Console]::Beep(196,300);
	# [System.Console]::Beep(220,400);
	$DRIVERLOCATION="driver\KOAXCJ__.inf"

    Write-Host -NoNewLine 'This application adds the H2 printer directly to this device. Would you like to continue?'
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

	Write-Host "Do not close window, window will close itself."
	try {
		Remove-Printer -Name "H2 Direct"
		Remove-PrinterPort -Name "h2direct10.10.10.5"
	}
	catch {}
	
	
	Invoke-Command {pnputil.exe -a $DRIVERLOCATION }	
	Add-PrinterDriver -Name "KONICA MINOLTA C360iSeriesPCL"

	$SLEEPY=10
	try {Add-PrinterPort -Name "h2direct10.10.10.5" -PrinterHostAddress "10.10.10.5"}
	catch {
		"Port already added, proceeding"
		$SLEEPY=0
	}
	Add-Printer -Name "H2 Direct" -PortName "h2direct10.10.10.5" -DriverName "KONICA MINOLTA C360iSeriesPCL"

	Write-Output "DONE!"


