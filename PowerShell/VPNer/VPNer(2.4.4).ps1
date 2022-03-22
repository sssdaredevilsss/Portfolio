############## Check internet ##############
do {
	#Check using DNS
	ping google.com -n 1 >> ".\VPNer.log"
	$DNSPingError = $?
	if ($DNSPingError -eq $True) {
		break
	}
	elseif ($DNSPingError -ne $True) {
		do {
			#Check using IP
			ping 8.8.8.8 -n 1 >> ".\VPNer.log"
			$IPPingError = $?
			if ($IPPingError -eq $True) {
				Write-Host "FQDN not resolving. Check DNS." -ForegroundColor blue -BackgroundColor red
				Write-Host 'Press any key to try again...'
				Write-Host ' '
				$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
			}
			elseif ($IPPingError -ne $True) {
				Write-Host "No internet connection. You must establish internet connection first." -ForegroundColor blue -BackgroundColor red
				Write-Host 'Press any key to try again...'
				Write-Host ' '
				$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
			}
		}until ($IPPingError -eq $False)
	}
}until (($DNSPingError -eq $True) -and ($IPPingError -eq $True))
############## Version Check #############
$ErrorActionPreference = "SilentlyContinue"
$LocalVersion = "2.4.0.4"
$VersionSource = 'https://bitbucket.org/daredevil739/version/downloads/vpner.version'
$VersionDestination = "$env:appdata\vpner.version"
$ImprovmentsSource = 'https://bitbucket.org/daredevil739/version/downloads/improvment.list'
$ImprovmentsDestination = "$env:appdata\improvment.list"
$user = 'bitbucket@email.com'
$pass = 'JMF83*#hf2hfU72egt33'  # may be enqrypted if needed
$pair = "$($user):$($pass)"
$encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
$basicAuthValue = "Basic $encodedCreds"
$Headers = @{
	Authorization = $basicAuthValue
}
Invoke-WebRequest -Uri $VersionSource -OutFile $VersionDestination -Headers $Headers
$ServerVersion = Get-Content $env:appdata\vpner.version
$Improvments = Get-Content $env:appdata\improvment.list
Remove-Item $env:appdata\vpner.version -Force
if ([version]$ServerVersion -gt [version]$LocalVersion) {
	$Shell = New-Object -ComObject "WScript.Shell"
	$Button = $Shell.Popup("You are using old version of VPNer ($LocalVersion)`nDownload the new one ($ServerVersion) from the repo", 0, "Update", 0)
}
else {
	Write-Host $null
}
$ErrorActionPreference = "Continue"
############## Check UAC ##############
Import-Module "$env:appdata\SwitchUACLevel.psm1"
Get-UACLevel | Findstr "Never notIfy" | Out-Null
if ($? -eq $True) {
	Set-UACLevel 2 | Out-Null
	Write-Host "UAC level was changed, restart the Setup" -ForegroundColor blue -BackgroundColor red
	Start-Sleep -s 10
	exit
}
############## Check PBK folder path ################
$ErrorActionPreference = "SilentlyContinue"
foreach ($UserFolder in Get-ChildItem -Name -Path "C:\Users" -Exclude "Public") { 
	$PbkLocation = Test-Path "C:\Users\$UserFolder\AppData\Roaming\Microsoft\Network\Connections\Pbk"
	if ($PbkLocation -eq $False) {
		New-Item -path "C:\Users\$UserFolder\AppData\Roaming\Microsoft\Network\Connections" -Type Get-ChildItemectory -Force
	}
}
$ErrorActionPreference = "Continue"
######## Creating PBK ###########
$ErrorActionPreference = "SilentlyContinue"
foreach ($UserFolder in Get-ChildItem -Name -Path "C:\Users" -Exclude "Public") { 
	$PbkFileExisting = Test-Path "C:\Users\$UserFolder\AppData\Roaming\Microsoft\Network\Connections\Pbk\rasphone.pbk"
	if ($PbkFileExisting -eq $True) {
		Get-Content "$env:appdata\rasphone.pbk" | Out-File -Encoding "UTF8" -Append "C:\Users\$UserFolder\AppData\Roaming\Microsoft\Network\Connections\Pbk\rasphone.pbk" 
	}
	else {
		Copy-Item "$env:appdata\rasphone.pbk" -Destination "C:\Users\$UserFolder\AppData\Roaming\Microsoft\Network\Connections\Pbk"  -Recurse -Force
	}
}
$ErrorActionPreference = "Continue"
############## Pick Domain ##############
do {
	Write-Host "`n====================== Pick the Domain ========================="
	Write-Host "`t'1' `tONE `tone.email.com"
	Write-Host "`t'2' `tTWO `ttwo.email.com"
	Write-Host "================================================================"
	$choice = Read-Host "`nEnter Choice"
} until (($choice -eq '1') -or ($choice -eq '2') )
switch ($choice) {
	'1' {
		Write-Host "`nYou have selected ONE domain"
		$domain = "one.email.com"
	}
	'2' {
		Write-Host "`nYou have selected TWO domain"
		$domain = "two.email.com"
	}
}
############## Installing Certs ##############
#####CA#####
do {
	Import-Certificate -FilePath "$env:appdata\CA.crt" -CertStoreLocation 'Cert:\LocalMachine\Root' >> ".\VPNer.log"
	$CAImportError = $?
	if ($CAImportError -eq $False) {
		Write-Host "There was an error during CA2 install, check the log file (.\VPNer.log)" -ForegroundColor blue -BackgroundColor red
		Write-Host 'Press any key to try again...'
		Write-Host ' '
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
	}
}until ($CAImportError -eq $True)
#####UserCrt#####
do {
	$user_cert = Get-ChildItem -Path ".\usr_crt" -Name
	if ($user_cert.count -gt 1) {
		Write-Host "It is more that one file in usr_crt folder. That should be only user cert! Not CA.." -ForegroundColor blue -BackgroundColor red	
		Write-Host 'Press any key to try again...'
		Write-Host ' '
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
		$UserCertImportError = $False
	}
	elseif ($user_cert.count -eq 1) {
		$UserCertImportError = $True
	}
	else {
		Write-Host "The usr_crt folder, looks empty, copy user cert into this folder.." -ForegroundColor blue -BackgroundColor red	
		Write-Host 'Press any key to try again...'
		Write-Host ' '
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
		$UserCertImportError = $False
	}
	$tp = Test-Path ".\usr_crt"
	if ($tp -eq $False) {
		Write-Host "Do not see usr_crt folder, copy it to the folder root and try again.." -ForegroundColor blue -BackgroundColor red	
		Write-Host 'Press any key to try again...'
		Write-Host ' '
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
		$UserCertImportError = $False
	}
}until (($tp -eq $True) -and ($UserCertImportError -eq $True))
#####Wrong password loop/Installing cert######
do {
	$secure_passwd = Read-Host -Prompt 'Enter cert password' -AsSecureString
	$ErrorActionPreference = "SilentlyContinue"
	$Thumbpr = (Get-PfxData -Password $secure_passwd -FilePath .\usr_crt\$user_cert).EndEntityCertificates.Thumbprint
	$ImportPFXOutput = try { Import-PfxCertificate -FilePath ".\usr_crt\$user_cert" -Password $secure_passwd -CertStoreLocation Cert:\LocalMachine\My *>&1 }catch { $_ }
	$DetectWrongPassword = $ImportPFXOutput | Select-String "import requires either a different password" -Quiet
	$DetectPreviousImport = $ImportPFXOutput | Select-String "$Thumbpr" -Quiet
	$ErrorActionPreference = "Continue"
	if ($DetectWrongPassword -eq $True) {
		Write-Host "Wrong password, try again" -ForegroundColor blue -BackgroundColor red
		Write-Host ' '
		$ImportPFXError = $False
	}
	elseif ($DetectPreviousImport -eq $True) {
		Write-Host "Cert installed" -ForegroundColor blue -BackgroundColor green
		break
	}
	elseif ($DetectPreviousImport -eq $False) {
		$ImportPFXOutput >> ".\VPNer.log"
		$ImportPFXError = $False
		Write-Host "There was an error during user cert install, check the log file (.\VPNer.log)" -ForegroundColor blue -BackgroundColor red
		Write-Host 'Press any key to try again...'
		Write-Host ' '
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
	}
}until ($ImportPFXError -eq $True)
############## Running VPN and writing routes ##############
do {
	rasdial "VPN IKE" >> ".\VPNer.log"
	$RasdialError = $?
	if ($RasdialError -eq $True) {
		netsh interface ipv4 add route 192.168.0.0/29 "VPN IKE" >> ".\VPNer.log"
		Set-VpnConnection -Name "VPN IKE" -MachineCertificateIssuerFilter "$env:appdata\CA.crt" >> ".\VPNer.log"
		rasdial "VPN IKE" /disconnect >> ".\VPNer.log"
	}
	elseif ($RasdialError -ne $True) {
		Write-Host "There is problem with runing VPN interface, check the log file (.\VPNer.log)" -ForegroundColor blue -BackgroundColor red
		Write-Host 'Press any key to try again...'
		Write-Host ' '
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
	}
}until ($RasdialError -eq $True)
############## RDP shortcut choice ##############
do {
	Write-Host "`n================== Pick the RDP shortcut set ==================="
	Write-Host "`t'1' `tOnly PC"
	Write-Host "`t'2' `tOnly Terminal"
	Write-Host "`t'3' `tBoth sets"
	Write-Host "================================================================"
	$RDPSet = Read-Host "`nEnter Choice"
} until (($RDPSet -eq '1') -or ($RDPSet -eq '2') -or ($RDPSet -eq '3') )
############## Making RDP shortcut ##############
$ErrorActionPreference = "SilentlyContinue"
switch ($RDPSet) {
	'1' {
		Write-Host "`nYou have selected: Only PC"
		$UserNameRDP = Read-Host -Prompt '(without domain) Enter User name'
		$PCNameRDP = Read-Host -Prompt '(without domain) Enter PC name'
		$PCTemplateRDP = (Get-Content $env:appdata\PCTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain"
		$PCTemplateRDP = $PCTemplateRDP -replace "full address:s:", "full address:s:$PCNameRDP.$domain"
		foreach ($UserFolder in Get-ChildItem -Name -Path "C:\Users" -Exclude "Public") {
			$PCTemplateRDP | Set-Content -path "C:\Users\$UserFolder\Desktop\PC.rdp" -Force
		}
	}
	'2' {
		Write-Host "`nYou have selected: Only Terminal"
		$UserNameRDP = Read-Host -Prompt '(without domain) Enter User name'
		$TerminalTemplateRDP = (Get-Content $env:appdata\TerminalTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain"
		foreach ($UserFolder in Get-ChildItem -Name -Path "C:\Users" -Exclude "Public") {
			$TerminalTemplateRDP | Set-Content -path "C:\Users\$UserFolder\Desktop\Terminal.rdp" -Force
		}
	}
	'3' {
		Write-Host "`nYou have selected: Both sets"
		$UserNameRDP = Read-Host -Prompt '(without domain) Enter User name'
		$PCNameRDP = Read-Host -Prompt '(without domain) Enter PC name'
		$PCTemplateRDP = (Get-Content $env:appdata\PCTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain"
		$PCTemplateRDP = $PCTemplateRDP -replace "full address:s:", "full address:s:$PCNameRDP.$domain"
		$TerminalTemplateRDP = (Get-Content $env:appdata\TerminalTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain"
		foreach ($UserFolder in Get-ChildItem -Name -Path "C:\Users" -Exclude "Public") {
			$PCTemplateRDP | Set-Content -path "C:\Users\$UserFolder\Desktop\PC.rdp" -Force
			$TerminalTemplateRDP | Set-Content -path "C:\Users\$UserFolder\Desktop\Terminal.rdp" -Force
		}
	}
}
$ErrorActionPreference = "Continue"
############## Succes button ##############
$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("Success. Press ok to exit.", 0, "Done", 0)