############### Translit Function ###############
function global:Translit {
    param([string]$inString)
    $Translit = @{
        [char]'а' = "a"
        [char]'А' = "A"
        [char]'б' = "b"
        [char]'Б' = "B"
        [char]'в' = "v"
        [char]'В' = "V"
        [char]'г' = "h"
        [char]'Г' = "H"
        [char]'ґ' = "g"
        [char]'Ґ' = "G"
        [char]'д' = "d"
        [char]'Д' = "D"
        [char]'е' = "e"
        [char]'Е' = "E"
        [char]'є' = "ie"
        [char]'Є' = "Ye"
        [char]'ж' = "zh"
        [char]'Ж' = "Zh"
        [char]'з' = "z"
        [char]'З' = "Z"
        [char]'и' = "y"
        [char]'И' = "Y"
        [char]'і' = "i"
        [char]'І' = "I"
        [char]'ї' = "i"
        [char]'Ї' = "Yi"
        [char]'й' = "i"
        [char]'Й' = "Y"
        [char]'к' = "k"
        [char]'К' = "K"
        [char]'л' = "l"
        [char]'Л' = "L"
        [char]'м' = "m"
        [char]'М' = "M"
        [char]'н' = "n"
        [char]'Н' = "N"
        [char]'о' = "o"
        [char]'О' = "O"
        [char]'п' = "p"
        [char]'П' = "P"
        [char]'р' = "r"
        [char]'Р' = "R"
        [char]'с' = "s"
        [char]'С' = "S"
        [char]'т' = "t"
        [char]'Т' = "T"
        [char]'у' = "u"
        [char]'У' = "U"
        [char]'ф' = "f"
        [char]'Ф' = "F"
        [char]'х' = "kh"
        [char]'Х' = "Kh"
        [char]'ц' = "ts"
        [char]'Ц' = "Ts"
        [char]'ч' = "ch"
        [char]'Ч' = "Ch"
        [char]'ш' = "sh"
        [char]'Ш' = "Sh"
        [char]'щ' = "shch"
        [char]'Щ' = "Shch"
        [char]'ь' = ""
        [char]'Ь' = ""
        [char]'ю' = "iu"
        [char]'Ю' = "Yu"
        [char]'я' = "ia"
        [char]'Я' = "Ya"
        [char]' ' = ""
        [char]'`' = ""
        [char]"'" = ""
    }
    $outCHR = ""
    foreach ($CHR in $inCHR = $inString.ToCharArray()) {
        if ($Null -cne $Translit[$CHR] )
        { $outCHR += $Translit[$CHR] }
        else
        { $outCHR += $CHR }
    }
    Write-Output $outCHR
}
#Get domain
$domain = (Get-ADDomain).dnsroot
############### Password Generating Function ###############
function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs = ""
    return [String]$characters[$random]
}
function Scramble-String([string]$inputString) {     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}
############### Pick Email Function ###############
Function PickMail {
    $CurrentAdmin = (whoami).Split('\')[1]
    $CurrentAdminMail = (Get-ADUser -Identity "$CurrentAdmin" -Properties mail).mail
    if (($null -eq $CurrentAdminMail) -and ($ActionChoice -in 4..6)) { $CurrentAdminMail = '(It is no email in ADMIN profile in AD. Chouse variant 3 and dont forget to add email in profile, lazy ass.)' }
    do {
        Write-Host "`n=================== Choose email address ======================="
        Write-Host "`t'1' `tSend to current user mail: $gMailAddr"
        Write-Host "`t'2' `tSend to current ADMIN mail: $CurrentAdminMail"
        Write-Host "`t'3' `tSend to specific mail address"
        Write-Host "`t'0' `tCancel"
        Write-Host "================================================================="
        $EmailAddrChoice = Read-Host "`nEnter Choice"
    } until (($EmailAddrChoice -in 1..3) -or ($EmailAddrChoice -eq '0'))
    switch ($EmailAddrChoice) {
        '1' {
            Write-Host "`nEmail will send to: $gMailAddr"
        }
        '2' {
            Write-Host "`nEmail will send to: $CurrentAdminMail"
            $Global:gMailAddr = $CurrentAdminMail
        }
        '3' {
            Write-Host "`nEmail will send to specify email"
            $Global:gMailAddr = Read-Host -Prompt "Enter email address to send Certs"
        }
        '0' {
            Clear-Host
            Continue RevokeLoop
        }
    }
}
############### Resend Cert Function ###############
function ResendCert {
    [int]$PlinkErrorCounter = 0
    $ResendVpnCertScript = "{:global gWorkingMode ""existing""; :global gExistingCertificate ""$global:gExistingCertificate""; :global gMailAddr ""$gMailAddr""; :global gExportPass ""$gExportPass"";[/system script run CreateNewCert];}"
    $ResendVpnCertScript | Out-File "$HOME\Documents\Ganter\ResendVpnCert.sh" -Encoding utf8
    do{
    Write-Host "Working on.."
    $PlinkOutput = plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -m "$HOME\Documents\Ganter\ResendVpnCert.sh"
    $PlinkOutput
    if(([int][int]$PlinkErrorCounter -gt '0') -and ($PlinkOutput | Select-String -Pattern 'expected end of command')){
        Write-Host "Ooops, ssh work failed $PlinkErrorCounter time. Try again, and call Bala if it wont run correctly in few times." -ForegroundColor blue -BackgroundColor red
        $PlinkRestart = Read-Host -Prompt "Press ENTER to try again"
    }
    if (($PlinkOutput | Select-String -Pattern 'expected end of command') -and ([int]$PlinkErrorCounter -eq '0')){
        Write-Host 'Ooops, ssh work failed. Trying again...' -ForegroundColor blue -BackgroundColor red
        [int]$PlinkErrorCounter++
        Start-Sleep 5
    }
    }until(-Not ($PlinkOutput | Select-String -Pattern 'expected end of command'))
    Remove-Item $HOME\Documents\Ganter\ResendVpnCert.sh
}
############### VPNer generate function ##############
function GenerateVPNEr {
    Start-Sleep 1
    if (($ActionChoice -eq '1') -or ($ActionChoice -in 3..4)){
        #Compiling VPNer
        $PlinkOutputFiltred = $PlinkOutput | Select-String -Pattern 'Cert Name:'
        $CertName = $PlinkOutputFiltred -replace 'Cert Name:', ''
        $CertName = $CertName + '.p12'
        $DownloadVpnCertScript = "get ""$CertName"" ""$env:appdata\Ganter\VPNerExtras\$CertName"""
        $DownloadVpnCertScript | Set-Content -Path "$env:appdata\Ganter\DownloadVpnCertScript.sh"
        $ErrorActionPreference = "SilentlyContinue"
        psftp -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -b "$env:appdata\Ganter\DownloadVpnCertScript.sh" | out-null
        $ErrorActionPreference = "Continue"
        if ($ActionChoice -eq '3'){$ResourseChoice = '3'}
        $DomainChoice = "'" + (Get-ADDomain).dnsroot + "'"
        $domain = (Get-ADDomain).dnsroot
        $UserSamNameInQuotes = "'" + $UserSamName + '@' + $domain +"'"
        $PCNameInQuotes = "'" + $PCName + '.' + $domain + "'"
        $VPNerTemplate = Get-Content -Path "$env:appdata\Ganter\VPNerBuilding\VPNerTemplate(2.4.4).ps1" 
        $VPNerTemplate -replace '!!!ChoiceData', "$DomainChoice" -replace '!!!RDPSetData', "$ResourseChoice" -replace '!!!UserNameRDPData', "$UserSamNameInQuotes" -replace '!!!PCNameRDPData', "$PCNameInQuotes" | Set-Content -Path "$env:appdata\Ganter\VPNerBuilding\VPNerConfiguredTemplate(2.4.4).ps1"
        & ".\PS2EXE\Ps1_To_Exe.exe" /ps1 "$env:appdata\Ganter\VPNerBuilding\VPNerConfiguredTemplate(2.4.4).ps1" `
        /exe "$env:appdata\Ganter\VPNerBuilding\VPNer\VPNer($UserSamName).exe" `
        /icon "$env:appdata\Ganter\VPNerBuilding\VPNer.ico" `
        /include "$env:appdata\Ganter\VPNerExtras" `
        /x64 `
        /workdir 0 `
        /extractdir 2 `
        /uac-admin `
        /deleteonexit `
        /overwrite `
        /fileversion 2.4.0.4 `
        /productversion 2.4.0.4 `
        /productname VPNer `
        /originalfilename VPNer `
        /internalfilename VPNer `
        /company EUGROUP `
        /trademarks EUGROUP `
        /copyright "Bala Nazarii"
        Wait-Process -Name 'Ps1_To_Exe'
        Start-Sleep 3
        Write-Host 'Archivating data...'
    & ".\7-Zip\7z.exe" a -tzip -mx0 -r0 -p1111 "$env:appdata\Ganter\VPNerBuilding\VPNer($UserSamName).zip" "$env:appdata\Ganter\VPNerBuilding\VPNer\*" | out-null
    & ".\7-Zip\7z.exe" a -tzip -mx0 -r0 -p1111 "$env:appdata\Ganter\VPNerBuilding\VPNer($UserSamName)q.zip" "$env:appdata\Ganter\VPNerBuilding\VPNer($UserSamName).zip" | out-null
    #Extra .rdp files personalization
$ErrorActionPreference = "SilentlyContinue"
switch ($ResourseChoice) {
	'1' {
		Write-Host "`nYou have selected: Only PC"
		$UserNameRDP = $UserSamName
		$PCNameRDP = $PCName
		$PCTemplateRDP = (Get-Content $env:appdata\Ganter\PCTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain" -replace "full address:s:", "full address:s:$PCNameRDP.$domain"
		$PCTemplateRDP | Set-Content -path "$env:appdata\Ganter\PC($UserSamName).rdp" -Force
	}
	'2' {
		Write-Host "`nYou have selected: Only Terminal"
		$UserNameRDP = $UserSamName
		$TerminalTemplateRDP = (Get-Content $env:appdata\Ganter\TerminalTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain"
		$TerminalTemplateRDP | Set-Content -path "$env:appdata\Ganter\Terminal($UserSamName).rdp" -Force
	}
	'3' {
		Write-Host "`nYou have selected: Both sets"
		$UserNameRDP = $UserSamName
		$PCNameRDP = $PCName
		$PCTemplateRDP = (Get-Content $env:appdata\Ganter\PCTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain" -replace "full address:s:", "full address:s:$PCNameRDP.$domain"
		$TerminalTemplateRDP = (Get-Content $env:appdata\Ganter\TerminalTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain"
			$PCTemplateRDP | Set-Content -path "$env:appdata\Ganter\PC($UserSamName).rdp" -Force
			$TerminalTemplateRDP | Set-Content -path "$env:appdata\Ganter\Terminal($UserSamName).rdp" -Force
	}
}
$ErrorActionPreference = "Continue"
    }
}
############### Ganter mail function ###############
function GanterMail {
    #Calling VPNer gen function
    GenerateVPNEr
    #Getting mail credentials
    $MailLogin = 'Ganter@email.com'
    #$GetMailPasswordHash = (get-credential).password | ConvertFrom-SecureString
    $MailPasswordHash = '76492d1116743f0423413b16050a5345MgB8AFAATQBkAGkASgB3AAZQAwADMAYwAxAGMANwBkAGMAZAAxAGQAZgBjADMAYwAzADMAMwBhADAAYgAzADgAMgA2AGUAYQA2AGUAMwA3AGIANgA4ADgAMAAxADUAMABmADMANwBjADAANwA4ADYANAAyAGIANwAwADAANQBhADUAOABhADkA'
    $MailPassword = $MailPasswordHash | ConvertTo-SecureString -Key (Get-Content $env:appdata\Ganter\Ganter_aes.key)
    $MailCredentials = New-Object System.Management.Automation.PsCredential($MailLogin, $MailPassword)
    #Send mail with user and PC names
    $From = 'Ganter@email.com'
    $To = "$gMailAddr"
    $Subject = "4"
    $htmlhead = "<html>
  <meta charset=utf-8>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				</style>
				<body>
                <p><strong>Username:</strong> $UserSamName</p>"
    $htmltail = "<p><strong>PC name:</strong> $PCName</p><p><strong>Archive password:</strong> 1111</p></body></html>"
    $body = $htmlhead + $htmltail
    $SMTPServer = 'mx1.email.com'
    $SMTPPort = '587'
    if($ResourseChoice -eq '1'){
        $Attachment = “$env:appdata\Ganter\VPNerBuilding\VPNer($UserSamName)q.zip”, “$env:appdata\Ganter\PC($UserSamName).rdp”
    }elseif($ResourseChoice -eq '2'){
        $Attachment = “$env:appdata\Ganter\VPNerBuilding\VPNer($UserSamName)q.zip”, “$env:appdata\Ganter\Terminal($UserSamName).rdp”
    }elseif(($ResourseChoice -eq '3') -or ($ActionChoice -eq '3')){
            $UserNameRDP = $UserSamName
            $PCNameRDP = $PCName
            $PCTemplateRDP = (Get-Content $env:appdata\Ganter\PCTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain" -replace "full address:s:", "full address:s:$PCNameRDP.$domain"
            $TerminalTemplateRDP = (Get-Content $env:appdata\Ganter\TerminalTemplate.rdp) -replace "username:s:", "username:s:$UserNameRDP@$domain"
                $PCTemplateRDP | Set-Content -path "$env:appdata\Ganter\PC($UserSamName).rdp" -Force
                $TerminalTemplateRDP | Set-Content -path "$env:appdata\Ganter\Terminal($UserSamName).rdp" -Force
    $Attachment = “$env:appdata\Ganter\VPNerBuilding\VPNer($UserSamName)q.zip”, “$env:appdata\Ganter\PC($UserSamName).rdp”, “$env:appdata\Ganter\Terminal($UserSamName).rdp”
    }elseif(($ActionChoice -eq '4') -or ($ActionChoice -eq '6')){
        $Attachment = “$env:appdata\Ganter\VPNer(Universal)q.zip”, “$env:appdata\Ganter\PCTemplate.rdp”, “$env:appdata\Ganter\TerminalTemplate.rdp”
    }
    Write-Host "Send email-4"
    Start-Sleep -Seconds 1
    Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Attachments $Attachment -Port $SMTPPort -UseSsl -Credential $MailCredentials
    #Remove cert from Mikrotik export storage
    $RemoveVpnCertFileScript = "rm ""$CertName"""
    $RemoveVpnCertFileScript | Set-Content -Path "$env:appdata\Ganter\RemoveVpnCertFileScript.sh"
    $ErrorActionPreference = "SilentlyContinue"
    psftp -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -b "$env:appdata\Ganter\RemoveVpnCertFileScript.sh" | out-null
    $ErrorActionPreference = "Continue"
}
############### Selection organization function ###############
function SelectOrg {
    $DomainName = (Get-ADDomain).name
    $Global:gOrgName = switch ($DomainName) {
        "one" { "Org1" }
        "two" { "Org2" }
    }
    if ($DomainName -eq "two") {
        do {
            Write-Host "`n================ Choose your organization ====================="
            Write-Host "`t'1' `tOrg2"
            Write-Host "`t'2' `tSubOrg"
            Write-Host "`t'0' `tCancel"
            Write-Host "=========================================================="
            $Global:OrganizationChoice = Read-Host "`nEnter Choice"
        } until (($Global:OrganizationChoice -eq '1') -or ($Global:OrganizationChoice -eq '2') )
        switch ($Global:OrganizationChoice) {
            '1' {
                Write-Host "`nYou have selected Org2"
                $Global:gOrgName = "Org2"
            }
            '2' {
                Write-Host "`nYou have selected SubOrg"
                $Global:gOrgName = "SubOrg"
            }
            '0' {
                Clear-Host
                Continue RevokeLoop
            }
        }
    }
}
############### Dev type ###############
function SelectDevType {
    do {
        Write-Host "`n=================== Choose device type ======================="
        Write-Host "`t'1' `tWorkstation"
        Write-Host "`t'2' `tNotebook"
        Write-Host "`t'3' `tMobile"
        Write-Host "`t'0' `tCancel"
        Write-Host "==============================================================="
        $Global:DeviceTypeChoice = Read-Host "`nEnter Choice"
    } until (($Global:DeviceTypeChoice -eq '0') -or ($Global:DeviceTypeChoice -eq '1') -or ($Global:DeviceTypeChoice -eq '2') -or ($Global:DeviceTypeChoice -eq '3') )
    switch ($Global:DeviceTypeChoice) {
        '1' {
            Write-Host "`nYou have selected: Device type Workstation"
            [string] $Global:gDevType = "Workstation"
        }
        '2' {
            Write-Host "`nYou have selected: Device type Notebook"
            [string] $Global:gDevType = "Notebook"
        }
        '3' {
            Write-Host "`nYou have selected: Device type Mobile"
            [string] $Global:gDevType = "Mobile"
        }
        '0' {
            Clear-Host
            continue RevokeLoop
        }
    }
}
############### Mobile OS ###############
function SelectMobileOS {
    if ($DeviceTypeChoice -eq '3') {
        do {
            Write-Host "`n=================== Choose mobile OS ======================="
            Write-Host "`t'1' `tAndroid"
            Write-Host "`t'2' `tIOS"
            Write-Host "`t'0' `tCancel"
            Write-Host "=========================================================="
            $Global:MobileOSChoice = Read-Host "`nEnter Choice"
        } until (($Global:MobileOSChoice -eq '0') -or ($Global:MobileOSChoice -eq '1') -or ($Global:MobileOSChoice -eq '2'))
        switch ($Global:MobileOSChoice) {
            '1' {
                Write-Host "`nYou have selected: Device OS Android"
                [string] $Global:gMobileType = "Android"
            }
            '2' {
                Write-Host "`nYou have selected: Device OS iOS"
                [string] $Global:gMobileType = "IOS"
            }
            '0' {
                Clear-Host
                continue RevokeLoop
            }
        }
    }
}
############### Cert existing check-function for device based ###############
function ExistingCheckDeviсeBased {
    plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server "/certificate print" | Out-File "$HOME\Documents\Ganter\Ganter.log" -Append  #dumb command to clear cert output
    $SSHError = $?
    if ($SSHError -eq $False) {
        do {
            Write-Host "There is some problem with SSH communucation. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
            $Error >> "$HOME\Documents\Ganter\Ganter.log"
            Write-Host ' '
            Write-Host "Press Enter to try again..."
            Read-Host 
        }until($SSHError -eq $True)
    }
    $CertNamesList = @()
    $CertNamesList += plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server "/system script run OutputCertNames" | Select-String -Pattern "CA2" -NotMatch
    $Global:FilteredCertNamesList = $CertNamesList | Select-String -Pattern "$global:DeviceSerialNumber" -AllMatches
    if (@($Global:FilteredCertNamesList).count -eq 1) {
        do {
            Write-Host "We found an existing Certificate for this device:"
            Write-Host "$Global:FilteredCertNamesList"
            $CreateConfirmation = Read-Host "Create the new one or resend current? [c - create new/r - resend/n - Cancel]"
            if ($CreateConfirmation -eq 'n') {
                Clear-Host
                continue RevokeLoop
            }
            elseif ($CreateConfirmation -eq 'c') {
                Write-Host "Creating..."
            }
            elseif ($CreateConfirmation -eq 'r') {
                $Global:gExistingCertificate = $FilteredCertNamesList
                PickMail
                ResendCert
                GanterMail
                $UserSamName >> "$HOME\Documents\Ganter\Ganter.log"
                $Shell = New-Object -ComObject "WScript.Shell"
                $Button = $Shell.Popup("Success. Press ok to exit.", 0, "Done", 0)
                $EndLoop = 0
                exit
            }
        }until (($CreateConfirmation -eq "c") -or ($CreateConfirmation -eq "r") -or ($CreateConfirmation -eq "n"))
    }
    elseif (@($FilteredCertNamesList).count -gt 1) {
        do {
            Write-Host "We found few Certificates for this device:"
            $FilteredCertNamesList
            $CreateConfirmation = Read-Host "Create the new one or select one to resend? [c - create new/r - resend/n - Cancel]"
            if ($CreateConfirmation -eq 'n') {
                Clear-Host
                continue RevokeLoop
            }
            elseif ($CreateConfirmation -eq 'c') {
                Write-Host "Creating..."
            }
            elseif ($CreateConfirmation -eq 'r') {
                $global:ans = $null
                $global:selection = $null
                While ($ans -lt 1 -or $ans -gt $FilteredCertNamesList.count) {
                    $mhead
                    Write-Host # empty line
                    Write-Host "Existing cert list."
                    Write-Host # empty line
                    $menu = @{}
                    for ($i = 1; $i -le $FilteredCertNamesList.count; $i++) {
                        if ($FilteredCertNamesList.count -gt 1) {
                            Write-Host -fore Cyan "  $i." $($FilteredCertNamesList[$i - 1]) 
                            $menu.Add($i, ($FilteredCertNamesList[$i - 1]))
                        }
                        else {
                            Write-Host -fore Cyan "  $i." $FilteredCertNamesList
                            $menu.Add($i, $FilteredCertNamesList)
                        }
                    }
                    Write-Host # empty line
                    [int]$global:ans = Read-Host ' Choose cert to resend'
                }
                $global:selection = $menu.Item([int]$ans)
                $Global:gExistingCertificate = $global:selection
                PickMail
                ResendCert
                GanterMail
                $UserSamName >> "$HOME\Documents\Ganter\Ganter.log"
                $Shell = New-Object -ComObject "WScript.Shell"
                $Button = $Shell.Popup("Success. Press ok to exit.", 0, "Done", 0)
                $EndLoop = 0
                exit
            }
        }until (($CreateConfirmation -eq "c") -or ($CreateConfirmation -eq "r") -or ($CreateConfirmation -eq "n"))
    }
}
############### Device based cert search #############
function DeviceBasedCertSearch {
    do {
        $global:DeviceSerialNumber = Read-Host 'Enter device serial or other unique ID'
        if ($null -eq $global:DeviceSerialNumber) {
            Write-Host "Enter value!"
        }
    }until($null -ne $global:DeviceSerialNumber)
    plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server "/certificate print" | out-null  #dumb command to clear cert output on mikrotik
    $SSHError = $?
    if ($SSHError -eq $False) {
        do {
            Write-Host "There is some problem with SSH communucation. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
            $Error >> "$HOME\Documents\Ganter\Ganter.log"
            Write-Host ' '
            Write-Host "Press Enter to try again..."
            Read-Host 
        }until($SSHError -eq $True)
    }
    $CertNamesList = @()
    $CertNamesList += plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server "/system script run OutputCertNames" | Select-String -Pattern "CA2" -NotMatch
    $FilteredCertNamesList = $CertNamesList | Select-String -Pattern "$DeviceSerialNumber" -AllMatches
    if (!$FilteredCertNamesList) {
        Write-Host "No such cert." -ForegroundColor blue -BackgroundColor red
        Write-Host ' '
        Write-Host "Press Enter to start over..."
        Read-Host 
        Clear-Host
        continue RevokeLoop
    }
    elseif (@($FilteredCertNamesList).count -eq 1) {
        $global:gExistingCertificate = $FilteredCertNamesList
        Write-Host "==========================>Certificate $global:gExistingCertificate found"
    }
    elseif (@($FilteredCertNamesList).count -gt 1) {
        $global:ans = $null
        $global:selection = $null
        While ($ans -lt 1 -or $ans -gt $FilteredCertNamesList.count) {
            $mhead
            Write-Host # empty line
            Write-Host "We found few certs with this search query."
            Write-Host # empty line
            $menu = @{}
            for ($i = 1; $i -le $FilteredCertNamesList.count; $i++) {
                if ($FilteredCertNamesList.count -gt 1) {
                    Write-Host -fore Cyan "  $i." $($FilteredCertNamesList[$i - 1]) 
                    $menu.Add($i, ($FilteredCertNamesList[$i - 1]))
                }
                else {
                    Write-Host -fore Cyan "  $i." $FilteredCertNamesList
                    $menu.Add($i, $FilteredCertNamesList)
                }
            }
            Write-Host # empty line
            [int]$global:ans = Read-Host ' Choose the right one'
        }
        $global:selection = $menu.Item([int]$ans)
        $global:gExistingCertificate = $global:selection
    }
}
############### Version Check ###############
$ErrorActionPreference = "SilentlyContinue"
$LocalVersion = "3.0.0.3"
.\curl\bin\curl.exe -s -u bitbucket@Org1.com.ua:JMF83*#hf2hfUH#*"&" https://api.bitbucket.org/2.0/repositories/test/grversion/src/master/Ganter.version -o $env:temp\Ganter.version
$ServerVersion = Get-Content $env:temp\Ganter.version
Remove-Item $env:temp\Ganter.version -Force
if ([version]$ServerVersion -gt [version]$LocalVersion) {
    $Shell = New-Object -ComObject "WScript.Shell"
    $Button = $Shell.Popup("You are using old version of Ganter ($LocalVersion)`nVersion ($ServerVersion) will be installed", 0, "Update", 0)
}
else {
    Write-Host $null
}
$ErrorActionPreference = "Continue"
############### Making log file location ###############
$LogPath = "$HOME\Documents\Ganter"
If (!(test-path $LogPath)) {
    New-Item -ItemType Directory -Force -Path $LogPath >> $ErrorVar
    Get-date >> "$HOME\Documents\Ganter\Ganter.log"
}
else {
    Get-date >> "$HOME\Documents\Ganter\Ganter.log"
    $ErrorVar >> "$HOME\Documents\Ganter\Ganter.log"
}
#Label for script restart if revoke-confirmation eq 'No'
:RevokeLoop While ($EndLoop -ne '0') {
    ############### Get random data for password Gen ###############
    $TempPassword = Get-RandomCharacters -length 5 -characters 'abcdefghiklmnoprstuvwxyz'
    $TempPassword += Get-RandomCharacters -length 5 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
    $TempPassword += Get-RandomCharacters -length 5 -characters '1234567890'
    $TempPassword += Get-RandomCharacters -length 5 -characters '!$%&/()=?}][{@#*+'
    ############### Action Choice Menu ###############
    do {
        Write-Host "`n================ How can I help you? ====================="
        Write-Host "`t'1' `tAllow remote access"
        Write-Host "`t'2' `tDenny remote access"
        Write-Host "`t'3' `tResend user cert"
        Write-Host "`n==================== Extra tasks ========================="
        Write-Host "`t'4' `tCreate device based cert (without user)"
        Write-Host "`t'5' `tRevoke device based cert (without user)"
        Write-Host "`t'6' `tResend device based cert (without user)"
        Write-Host "`t'7' `tOnly add USER and PC to remote access groups"
        Write-Host "`t'8' `tOnly remove USER and PC from remote access groups"
        Write-Host "=========================================================="
        $ActionChoice = Read-Host "`nEnter Choice"
        if ($ActionChoice -notin 1..8) {
            Clear-Host
        }
    } until ($ActionChoice -in 1..8)
    switch ($ActionChoice) {
        '1' {
            Write-Host "`nYou have selected: Allow remote access"
            $Action = "Add"
        }
        '2' {
            Write-Host "`nYou have selected: Denny remote access"
            $Action = "Remove"
        }
        '3' {
            Write-Host "`nYou have selected: Resend user cert"
        }
        '4' {
            Write-Host "`nYou have selected: Create device based cert"
        }
        '5' {
            Write-Host "`nYou have selected: Revoke device based cert"
        }
        '6' {
            Write-Host "`nYou have selected: Resend device based cert"
        }
        '7' {
            Write-Host "`nYou have selected: Only add user to remote access groups"
            $Action = "Add"
        }
        '8' {
            Write-Host "`nYou have selected: Only remove user from remote access groups"
            $Action = "Remove"
        }
    }
    ############### Resourses Choice Menu ###############
    #Will ask only if Allowing access
    if (($ActionChoice -eq '1') -or ($ActionChoice -eq '7')) {
        do {
            Write-Host "`n================ What remote resource? ====================="
            Write-Host "`t'1' `tOnly PC"
            Write-Host "`t'2' `tOnly Terminal"
            Write-Host "`t'3' `tPC and Terminal"
            Write-Host "`t'0' `tCancel"
            Write-Host "============================================================"
            $ResourseChoice = Read-Host "`nEnter Choice"
        } until (($ResourseChoice -eq '0') -or ($ResourseChoice -eq '1') -or ($ResourseChoice -eq '2') -or ($ResourseChoice -eq '3') )
        switch ($ResourseChoice) {
            '1' {
                Write-Host "`nYou have selected: Only PC access"
            }
            '2' {
                Write-Host "`nYou have selected: Only Terminal access"
            }
            '3' {
                Write-Host "`nYou have selected: PC and Terminal access"
            }
            '0' {
                Clear-Host
                continue RevokeLoop
            }
        }
    }
    ############### Checking and searching user in AD ###############
    if ($ActionChoice -notin 4..6) {
        do {
            do {
                $UNAME = Read-Host -Prompt "Enter User name"
                if ($Null -eq $UNAME) {
                    write-host "Enter value!"
                }
            }until ($null -ne $UNAME)
            do {
                $FullUserAdData = Get-ADUser -Filter "(Name -Like '*$UNAME*') -and (Enabled -eq 'True')"
                $UserSamName = $FullUserAdData.SamAccountName
                $ErrorVar = $?
                if ($ErrorVar -eq $False) {
                    Write-Host "There is some problem with AD communucation. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
                    $Error >> "$HOME\Documents\Ganter\Ganter.log"
                    Write-Host ' '
                    Write-Host "Press Enter to continue..."
                    Read-Host 
                }
            }until ($ErrorVar -eq $True)
            $i = 0
            $ChoiceData = @()
            $UserSamName | Foreach-Object {
                $ChoiceData += "$_"
            }
            if (!$ChoiceData) {
                Write-Host "No such User. Try to search in another way or check if it exist or enabled" -ForegroundColor blue -BackgroundColor red
            }
            elseif ($ChoiceData -eq $UserSamName) {
                Write-Host "==========================>User $UserSamName found"
                break
            }
            elseif ($ChoiceData.count -gt 1) {
                $global:ans = $null
                $global:selection = $null
                While ($ans -lt 1 -or $ans -gt $ChoiceData.count) {
                    $mhead
                    Write-Host # empty line
                    Write-Host "We found few users with this search query."
                    Write-Host # empty line
                    $menu = @{}
                    for ($i = 1; $i -le $ChoiceData.count; $i++) {
                        if ($ChoiceData.count -gt 1) {
                            Write-Host -fore Cyan "  $i." $($ChoiceData[$i - 1]) 
                            $menu.Add($i, ($ChoiceData[$i - 1]))
                        }
                        else {
                            Write-Host -fore Cyan "  $i." $ChoiceData
                            $menu.Add($i, $ChoiceData)
                        }
                    }
                    Write-Host # empty line
                    [int]$global:ans = Read-Host ' Choose the right one (Or press 0 to cancel)'
                    if ($global:ans -eq 0) {
                        Clear-Host
                        Continue RevokeLoop
                    }
                }
                $global:selection = $menu.Item([int]$ans)
                $UserSamName = $global:selection
                $FullUserAdData = Get-ADUser -Filter "(SamAccountName -Like '*$UserSamName*') -and (Enabled -eq 'True')"
            }
        } While (!$UserSamName)
        ############### Checking and searching PC in AD ###############
        if ($ResourseChoice -ne '2') {
            do {
                $PCDescriptionName = $FullUserAdData.GivenName + " " + $FullUserAdData.Surname
                $PCDescriptionNameMirror = $FullUserAdData.Surname + " " + $FullUserAdData.GivenName
                do {
                    $PCFullData = Get-ADComputer -Filter "(Description -Like '*$PCDescriptionName*') -and (Enabled -eq 'True')" -Properties Name, Description
                    $PCName = $PCFullData.Name
                    if ($null -eq $PCName) {
                        $PCFullData = Get-ADComputer -Filter "(Description -Like '*$PCDescriptionNameMirror*') -and (Enabled -eq 'True')" -Properties Name, Description
                        $PCName = $PCFullData.Name
                    }
                    if ($null -eq $PCName) {
                        $UserSurname = $FullUserAdData.Surname
                        $PCFullData = Get-ADComputer -Filter "(Description -Like '*$UserSurname*') -and (Enabled -eq 'True')" -Properties Name, Description
                        $PCName = $PCFullData.Name
                    }
                    $ErrorVar = $?
                    if ($ErrorVar -eq $False) {
                        Write-Host "There is some problem with AD communucation. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
                        $Error >> "$HOME\Documents\Ganter\Ganter.log"
                        Write-Host ' '
                        Write-Host "Press Enter to continue..."
                        Read-Host 
                    }
                }until ($ErrorVar -eq $True)
                $i = 0
                $ChoiceData = @()
                $PCName | Foreach-Object {
                    $ChoiceData += "$_"
                }
                if (!$ChoiceData) {
                    Write-Host "Cant find PC with that desctiption." -ForegroundColor blue -BackgroundColor red
                    Write-Host "That could happen if it is no description in PC object or it`s different with search query." -ForegroundColor blue -BackgroundColor red
                    Write-Host "Enter PC-name manualy and check AD information for mistakes" -ForegroundColor blue -BackgroundColor red
                    $PCName = Read-Host -Prompt "Enter PC Name"
                }
                elseif ($ChoiceData -eq $PCName) {
                    $PCDescription = "(" + $PCFullData.description + ")"
                    Write-Host "==========================>PC: $PCName $PCDescription found"
                    break
                }
                elseif ($ChoiceData.count -gt 1) {
                    $global:ans = $null
                    $global:selection = $null
                    While ($ans -lt 1 -or $ans -gt $ChoiceData.count) {
                        $mhead
                        Write-Host # empty line
                        Write-Host "We found few PC with this description."
                        Write-Host # empty line
                        $menu = @{}
                        for ($i = 1; $i -le $ChoiceData.count; $i++) {
                            if ($ChoiceData.count -gt 1) {
                                Write-Host -fore Cyan "  $i." $($ChoiceData[$i - 1]) 
                                $menu.Add($i, ($ChoiceData[$i - 1]))
                            }
                            else {
                                Write-Host -fore Cyan "  $i." $ChoiceData
                                $menu.Add($i, $ChoiceData)
                            }
                        }
                        Write-Host # empty line
                        [int]$global:ans = Read-Host ' Choose the right one (Or press 0 to cancel)'
                        if ($global:ans -eq 0) {
                            Clear-Host
                            Continue RevokeLoop
                        }
                    }
                    $global:selection = $menu.Item([int]$ans)
                    $PCName = $global:selection
                }
            }while (!$PCName)
        }#if not only user site or only rdp access
    }#if not device based
    ############### Addind or Removing groups ###############
    #Only if Allowing or Removing access
    if (($ActionChoice -in 1..2) -or ($ActionChoice -in 7..8)) {
        Do {
            #Only for PC or both resourse allowing or removing
            if (($ResourseChoice -eq '1') -or ($ResourseChoice -eq '3') -or ($ActionChoice -eq '8') -or ($ActionChoice -eq '2') -or (($ActionChoice -in 7..8) -and ($ResourseChoice -eq '1')) -or (($ActionChoice -in 7..8) -and ($ResourseChoice -eq '3'))) {
                $GroupManage = "$Action-ADGroupMember -Identity 'RDGW PC Users' -Members $UserSamName"
                Invoke-Expression $GroupManage
                    $GroupManage = "$Action-ADGroupMember -Identity 'RDGW resources' -Members $PCName$"
                    Invoke-Expression $GroupManage
            }
            #Only for Terminal or both resourse allowing or removing
            if (($ResourseChoice -eq '2') -or ($ResourseChoice -eq '3') -or ($ActionChoice -eq '8') -or ($ActionChoice -eq '2') -or (($ActionChoice -eq '7') -and ($ResourseChoice -eq '2')) -or (($ActionChoice -eq '7') -and ($ResourseChoice -eq '3'))) {
                $GroupManage = "$Action-ADGroupMember -Identity 'RDGW terminal Users' -Members $UserSamName"
                Invoke-Expression $GroupManage
            }
            $ErrorVar = $?
            if ($ErrorVar -eq $False) {
                Write-Host "There is some problem with operation of adding/removing AD group. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
                $Error >> "$HOME\Documents\Ganter\Ganter.log"
                Write-Host ' '
                Write-Host "Press Enter to continue..."
                Read-Host 
            }
        }until ($ErrorVar -eq $True)
    }
    ############### End if no cert task running ###############
    if ($ActionChoice -in 7..8) {
        $UserSamName >> "$HOME\Documents\Ganter\Ganter.log"
        $Shell = New-Object -ComObject "WScript.Shell"
        $Button = $Shell.Popup("Success. Press ok to exit.", 0, "Done", 0)
        $EndLoop = 0
        exit
    }
    #
    #This block temporaty commented, couse no winRM services allowed
    ############### Addind or Removing local group on PC remotly ###############
    #    Do{
    #Import-Module AdmPwd.PS
    #$LocalAdmin = '☻dmin'
    #$Password = (Get-AdmPwdPassword -ComputerName $PCName).Password
    #$ErrorVarLaps = $?
    #$SPassword = ConvertTo-SecureString -AsPlainText $Password -Force
    #$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $LocalAdmin,$SPassword
    #$GroupManage = "$Action-LocalGroupMember -Group 'Remote Desktop Users' -Member '$domain\$UserSamName'"
    #Invoke-Command -ComputerName $PCName -Credential $Cred -ScriptBlock {$GroupManage}
    #$ErrorVar = $?
    #    if ($ErrorVar -eq $False){
    #        Write-Host "There is some problem with operation remote communication with PC. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
    #        $Error >> "C:\ProgramData\Ganter\logfile.txt"
    #        Write-Host ' '
    #        Write-Host "Press Enter to continue..."
    #        $TrashInput = Read-Host 
    #    }elseif ($ErrorVarLaps -eq $False) {
    #        Write-Host "There is some problem with getting Password from LAPS. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
    #        $Error >> "C:\ProgramData\Ganter\logfile.txt"
    #        Write-Host ' '
    #        Write-Host "Press Enter to continue..."
    #        $TrashInput = Read-Host 
    #    }
    #}until ($ErrorVar -eq $True)
    #
    ############### Writing data to variables for script building ###############
    if ($ActionChoice -notin 4..6) {
        $ADEndUserData = Get-ADUser -Filter "(SamAccountName -like '$UserSamName') -and (Enabled -eq 'True')" -Properties Surname, GivenName, mail
        $gUserFirstName = Translit($ADEndUserData.GivenName)
        $gUserSecondName = Translit($ADEndUserData.Surname)
        $gMailAddr = $ADEndUserData.mail
    }
    if ($ActionChoice -in 4..6) {
        $gMailAddr = '(No current mail avaible when device based operating)'
    }
    elseif (($null -eq $gMailAddr) -and ($ActionChoice -notin 4..6)) { $gMailAddr = '(It is no email in user profile in AD. Chouse variant 3 and dont forget to add email in profile, lazy ass.)' }
    $gExportPass = Scramble-String $TempPassword
    ############################## SSH ##############################
    #Check fingerprint
    $ErrorActionPreference = "SilentlyContinue"
    $ec2FingerPrint = 'vpn.server ssh-rsa AAAAB3NzaC1yc2EAAAABAwAAAQEAoBQcYs204eILe0UxZIQ4v6vLrGsm0BLwR4qpdVYeSAzcTxwzgzddd8dR7bCeySTLtBmfTZxDaVJsYTqQJq4eSvZYEdpOq2/+lby+/oMNqSebhlAuiToNHIgAQfl7pdDpoIbRk9/VXlw=='
    Get-ItemProperty -Path "HKCU:\\Software\SimonTatham\PuTTY\SshHostKeys" -Name "rsa2@22:vpn.server" | out-null
    $FingerPrintError = $?
    if ($FingerPrintError -eq $True) {
        $null
    }
    else {
        Set-ItemProperty -Path "HKCU:\\Software\SimonTatham\PuTTY\SshHostKeys" -Name "rsa2@22:vpn.server" -value "$ec2FingerPrint"  | out-null
    }
    $ErrorActionPreference = "Continue"
    #Getting ppk passphrase
    $PPKPassphrase = "Snn388SAfij3uiikSJsgfujt3giSKJFG3i9923WQJ"
    ############### Cert existing check ###############
    #only if Creating
    if ($ActionChoice -eq '1') {
        plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server "/certificate print" | Out-File "$HOME\Documents\Ganter\Ganter.log" -Append  #dumb command to clear cert output
        $SSHError = $?
        if ($SSHError -eq $False) {
            do {
                Write-Host "There is some problem with SSH communucation. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
                $Error >> "$HOME\Documents\Ganter\Ganter.log"
                Write-Host ' '
                Write-Host "Press Enter to try again..."
                Read-Host 
            }until($SSHError -eq $True)
        }
        $CertNamesList = @()
        $CertNamesList += plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server "/system script run OutputCertNames" | Select-String -Pattern "CA2" -NotMatch
        $FilteredCertNamesList = $CertNamesList | Select-String -Pattern "$gUserSecondName $gUserFirstName" -AllMatches
        if (@($FilteredCertNamesList).count -eq 1) {
            do {
                Write-Host "We found an existing Certificate for this user:"
                Write-Host "$FilteredCertNamesList"
                $CreateConfirmation = Read-Host "Create the new one or resend current? [c - create new/r - resend/n - Cancel]"
                if ($CreateConfirmation -eq 'n') {
                    Clear-Host
                    continue RevokeLoop
                }
                elseif ($CreateConfirmation -eq 'c') {
                    Write-Host "Creating..."
                }
                elseif ($CreateConfirmation -eq 'r') {
                    $global:gExistingCertificate = $FilteredCertNamesList
                    PickMail
                    ResendCert
                    GanterMail
                    $UserSamName >> "$HOME\Documents\Ganter\Ganter.log"
                    $Shell = New-Object -ComObject "WScript.Shell"
                    $Button = $Shell.Popup("Success. Press ok to exit.", 0, "Done", 0)
                    $EndLoop = 0
                    exit
                }
            }until (($CreateConfirmation -eq "c") -or ($CreateConfirmation -eq "r") -or ($CreateConfirmation -eq "n"))
        }
        elseif (@($FilteredCertNamesList).count -gt 1) {
            do {
                Write-Host "We found few Certificates for this user:"
                $FilteredCertNamesList
                $CreateConfirmation = Read-Host "Create the new one or select one to resend? [c - create new/r - resend/n - Cancel]"
                if ($CreateConfirmation -eq 'n') {
                    Clear-Host
                    continue RevokeLoop
                }
                elseif ($CreateConfirmation -eq 'c') {
                    Write-Host "Creating..."
                }
                elseif ($CreateConfirmation -eq 'r') {
                    $global:ans = $null
                    $global:selection = $null
                    While ($ans -lt 1 -or $ans -gt $FilteredCertNamesList.count) {
                        $mhead
                        Write-Host # empty line
                        Write-Host "Existing cert list."
                        Write-Host # empty line
                        $menu = @{}
                        for ($i = 1; $i -le $FilteredCertNamesList.count; $i++) {
                            if ($FilteredCertNamesList.count -gt 1) {
                                Write-Host -fore Cyan "  $i." $($FilteredCertNamesList[$i - 1]) 
                                $menu.Add($i, ($FilteredCertNamesList[$i - 1]))
                            }
                            else {
                                Write-Host -fore Cyan "  $i." $FilteredCertNamesList
                                $menu.Add($i, $FilteredCertNamesList)
                            }
                        }
                        Write-Host # empty line
                        [int]$global:ans = Read-Host ' Choose cert to resend'
                    }
                    $global:selection = $menu.Item([int]$ans)
                    $global:gExistingCertificate = $global:selection
                    PickMail
                    ResendCert
                    GanterMail
                    $UserSamName >> "$HOME\Documents\Ganter\Ganter.log"
                    $Shell = New-Object -ComObject "WScript.Shell"
                    $Button = $Shell.Popup("Success. Press ok to exit.", 0, "Done", 0)
                    $EndLoop = 0
                    exit
                }
            }until (($CreateConfirmation -eq "c") -or ($CreateConfirmation -eq "r") -or ($CreateConfirmation -eq "n"))
        }
        ############### Selection creating data ###############
        SelectOrg # Selection organization
        SelectDevType # Dev type
        SelectMobileOS # Mobile OS
    } #end of Creating statement
    ############### Generating creating script ###############
    if ($ActionChoice -eq '1') {
        PickMail
        [int]$PlinkErrorCounter = '0'
        $CreateVpnCertScript = "{:global gWorkingMode ""yes""; :global gUserFirstName ""$gUserFirstName""; :global gUserSecondName ""$gUserSecondName""; :global gState ""Lviv""; :global gLocality ""Home""; :global gOrgName ""$global:gOrgName""; :global gDevType ""$gDevType""; :global gDaysValid ""365""; :global gUnit ""nounit""; :global gMailAddr ""$gMailAddr""; :global gExportPass ""$gExportPass"";[/system script run CreateNewCert];}"
        $CreateVpnCertScript | Out-File "$HOME\Documents\Ganter\CreateVpnCert.sh" -Encoding utf8
        do{
        Write-Host "Working on.."
        $PlinkOutput = plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -m "$HOME\Documents\Ganter\CreateVpnCert.sh"
        $PlinkOutput
        if(([int]$PlinkErrorCounter -gt '0') -and ($PlinkOutput | Select-String -Pattern 'expected end of command')){
            Write-Host "Ooops, ssh work failed $PlinkErrorCounter time. Try again, and call Bala if it wont run correctly in few times." -ForegroundColor blue -BackgroundColor red
            $PlinkRestart = Read-Host -Prompt "Press ENTER to try again"
        }
        if (($PlinkOutput | Select-String -Pattern 'expected end of command') -and ([int]$PlinkErrorCounter -eq '0')){
            Write-Host 'Ooops, ssh work failed. Trying again...' -ForegroundColor blue -BackgroundColor red
            [int]$PlinkErrorCounter++
            Start-Sleep 5
        }
    }until(-Not ($PlinkOutput | Select-String -Pattern 'expected end of command'))
        Remove-Item $HOME\Documents\Ganter\CreateVpnCert.sh
    }
    elseif (($ActionChoice -eq '2') -or ($ActionChoice -eq '3')) {
        ############### Cert search ###############
        plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server "/certificate print" | out-null  #dumb command to clear cert output on mikrotik
        $SSHError = $?
        if ($SSHError -eq $False) {
            do {
                Write-Host "There is some problem with SSH communucation. Full log you can see ($HOME\Documents\Ganter\Ganter.log)" -ForegroundColor blue -BackgroundColor red 
                $Error >> "$HOME\Documents\Ganter\Ganter.log"
                Write-Host ' '
                Write-Host "Press Enter to try again..."
                Read-Host 
            }until($SSHError -eq $True)
        }
        $CertNamesList = @()
        $CertNamesList += plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server "/system script run OutputCertNames" | Select-String -Pattern "CA2" -NotMatch
        $FilteredCertNamesList = $CertNamesList | Select-String -Pattern "$gUserSecondName $gUserFirstName" -AllMatches
        if (!$FilteredCertNamesList) {
            $ManualCertSearchQuery = Read-Host -Prompt "Not found any cert by Name-Surname match. Try manual search (Use latina!)"
            $FilteredCertNamesList = $CertNamesList | Select-String -Pattern "$ManualCertSearchQuery" -AllMatches
        }
        if (!$FilteredCertNamesList) {
            Write-Host "No such cert." -ForegroundColor blue -BackgroundColor red
            Write-Host ' '
            Write-Host "Press Enter to start over..."
            Read-Host 
            Clear-Host
            continue RevokeLoop
        }
        elseif (@($FilteredCertNamesList).count -eq 1) {
            $global:gExistingCertificate = $FilteredCertNamesList
            Write-Host "==========================>Certificate $global:gExistingCertificate found"
        }
        elseif (@($FilteredCertNamesList).count -gt 1) {
            $global:ans = $null
            $global:selection = $null
            While ($ans -lt 1 -or $ans -gt $FilteredCertNamesList.count) {
                $mhead
                Write-Host # empty line
                Write-Host "We found few certs with this search query."
                Write-Host # empty line
                $menu = @{}
                for ($i = 1; $i -le $FilteredCertNamesList.count; $i++) {
                    if ($FilteredCertNamesList.count -gt 1) {
                        Write-Host -fore Cyan "  $i." $($FilteredCertNamesList[$i - 1]) 
                        $menu.Add($i, ($FilteredCertNamesList[$i - 1]))
                    }
                    else {
                        Write-Host -fore Cyan "  $i." $FilteredCertNamesList
                        $menu.Add($i, $FilteredCertNamesList)
                    }
                }
                Write-Host # empty line
                [int]$global:ans = Read-Host ' Choose the right one'
            }
            $global:selection = $menu.Item([int]$ans)
            $global:gExistingCertificate = $global:selection
        }
        ############### Revoke Cert ###############
        if ($ActionChoice -eq '2') {
            do {
                $RevokeConfirmation = Read-Host "Do you really want to revoke certificate $global:gExistingCertificate ? [y/n]"
                if ($RevokeConfirmation -eq 'n') {
                    Clear-Host
                    continue RevokeLoop
                }
            }while ($RevokeConfirmation -ne "y")
            if ($RevokeConfirmation -eq "y") {
                Write-Host "If you say so ..."
                $RevokeVpnCertScript = "{:global gExistingCertificate ""$global:gExistingCertificate"";[/system script run RevokeCert];}"
                $RevokeVpnCertScript | Out-File "$HOME\Documents\Ganter\RevokeVpnCert.sh" -Encoding utf8
                plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -m "$HOME\Documents\Ganter\RevokeVpnCert.sh"
            }
        }
        elseif ($ActionChoice -eq '3') {
            ###############  Resend cert ############### 
            #Generating resend script
            PickMail
            [int]$PlinkErrorCounter = '0'
            $ResendVpnCertScript = "{:global gWorkingMode ""existing""; :global gExistingCertificate ""$global:gExistingCertificate""; :global gMailAddr ""$gMailAddr""; :global gExportPass ""$gExportPass"";[/system script run CreateNewCert];}"
            $ResendVpnCertScript | Out-File "$HOME\Documents\Ganter\ResendVpnCert.sh" -Encoding utf8
            do{
            Write-Host "Working on.."
            $PlinkOutput = plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -m "$HOME\Documents\Ganter\ResendVpnCert.sh"
            $PlinkOutput
            if(([int]$PlinkErrorCounter -gt '0') -and ($PlinkOutput | Select-String -Pattern 'expected end of command')){
                Write-Host "Ooops, ssh work failed $PlinkErrorCounter time. Try again, and call Bala if it wont run correctly in few times." -ForegroundColor blue -BackgroundColor red
                $PlinkRestart = Read-Host -Prompt "Press ENTER to try again"
            }
            if (($PlinkOutput | Select-String -Pattern 'expected end of command') -and ([int]$PlinkErrorCounter -eq '0')){
                Write-Host 'Ooops, ssh work failed. Trying again...' -ForegroundColor blue -BackgroundColor red
                [int]$PlinkErrorCounter++
                Start-Sleep 5
            }
              }until(-Not ($PlinkOutput | Select-String -Pattern 'expected end of command'))
            Remove-Item $HOME\Documents\Ganter\ResendVpnCert.sh
        }
    }#and Revoke/Resend statement
    ############### Manual cert operations ###############
    #device based cert creation
    if ($ActionChoice -eq '4') {
        do {
            $DeviceSerialNumber = Read-Host 'Enter device serial or other unique ID'
            if ($null -eq $DeviceSerialNumber) {
                Write-Host "Enter value!"
            }
        }until ($null -ne $DeviceSerialNumber)
        ExistingCheckDeviсeBased
        SelectOrg
        SelectDevType
        SelectMobileOS
        PickMail
        [int]$PlinkErrorCounter = '0'
        $CreateVpnCertScript = "{:global gWorkingMode ""yes""; :global gUserFirstName ""$DeviceSerialNumber""; :global gUserSecondName ""$gDevType""; :global gState ""Lviv""; :global gLocality ""Home""; :global gOrgName ""$global:gOrgName""; :global gDevType ""$gDevType""; :global gDaysValid ""365""; :global gUnit ""nounit""; :global gMailAddr ""$gMailAddr""; :global gExportPass ""$gExportPass"";[/system script run CreateNewCert];}"
        $CreateVpnCertScript | Out-File "$HOME\Documents\Ganter\CreateVpnCert.sh" -Encoding utf8
        do{
        Write-Host "Working on.."
        $PlinkOutput = plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -m "$HOME\Documents\Ganter\CreateVpnCert.sh"
        $PlinkOutput
        if(([int]$PlinkErrorCounter -gt '0') -and ($PlinkOutput | Select-String -Pattern 'expected end of command')){
            Write-Host "Ooops, ssh work failed $PlinkErrorCounter time. Try again, and call Bala if it wont run correctly in few times." -ForegroundColor blue -BackgroundColor red
            $PlinkRestart = Read-Host -Prompt "Press ENTER to try again"
        }
        if (($PlinkOutput | Select-String -Pattern 'expected end of command') -and ([int]$PlinkErrorCounter -eq '0')){
            Write-Host 'Ooops, ssh work failed. Trying again...' -ForegroundColor blue -BackgroundColor red
            [int]$PlinkErrorCounter++
            Start-Sleep 5
        }
        }until(-Not ($PlinkOutput | Select-String -Pattern 'expected end of command'))
        GanterMail
        Remove-Item $HOME\Documents\Ganter\CreateVpnCert.sh
    }
    #Revoke Device based cert
    if ($ActionChoice -eq '5') {
        DeviceBasedCertSearch
        do {
            $RevokeConfirmation = Read-Host "Do you really want to revoke certificate $global:gExistingCertificate ? [y/n]"
            if ($RevokeConfirmation -eq 'n') {
                Clear-Host
                continue RevokeLoop
            }
        }while ($RevokeConfirmation -ne "y")
        if ($RevokeConfirmation -eq "y") {
            Write-Host "If you say so ..."
            $RevokeVpnCertScript = "{:global gExistingCertificate ""$global:gExistingCertificate"";[/system script run RevokeCert];}"
            $RevokeVpnCertScript | Out-File "$HOME\Documents\Ganter\RevokeVpnCert.sh" -Encoding utf8
            plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -m "$HOME\Documents\Ganter\RevokeVpnCert.sh"
        }
    }
    #Resend device based cert
    if ($ActionChoice -eq '6') {
        DeviceBasedCertSearch
        PickMail
        [int]$PlinkErrorCounter = '0'
        $ResendVpnCertScript = "{:global gWorkingMode ""existing""; :global gExistingCertificate ""$global:gExistingCertificate""; :global gMailAddr ""$gMailAddr""; :global gExportPass ""$gExportPass"";[/system script run CreateNewCert];}"
        $ResendVpnCertScript | Out-File "$HOME\Documents\Ganter\ResendVpnCert.sh" -Encoding utf8
        do{
        Write-Host "Working on.."
        $PlinkOutput = plink -pw $PPKPassphrase -batch -i $env:appdata\Ganter\private_p.ppk script@vpn.server -m "$HOME\Documents\Ganter\ResendVpnCert.sh"
        $PlinkOutput
        if(([int]$PlinkErrorCounter -gt '0') -and ($PlinkOutput | Select-String -Pattern 'expected end of command')){
            Write-Host "Ooops, ssh work failed $PlinkErrorCounter time. Try again, and call Bala Nazarii if it wont run correctly in few times." -ForegroundColor blue -BackgroundColor red
            $PlinkRestart = Read-Host -Prompt "Press ENTER to try again"
        }
        if (($PlinkOutput | Select-String -Pattern 'expected end of command') -and ([int]$PlinkErrorCounter -eq '0')){
            Write-Host 'Ooops, ssh work failed. Trying again...' -ForegroundColor blue -BackgroundColor red
            [int]$PlinkErrorCounter++
            Start-Sleep 5
        }
        }until(-Not ($PlinkOutput | Select-String -Pattern 'expected end of command'))
        GanterMail
        Remove-Item $HOME\Documents\Ganter\ResendVpnCert.sh
    }
    ############### Extra Email ###############
    if (($ActionChoice -eq '1') -or ($ActionChoice -eq '3')) {
        GanterMail
    }
    Start-Sleep 1
    $UserSamName >> "$HOME\Documents\Ganter\Ganter.log"
    $Shell = New-Object -ComObject "WScript.Shell"
    $Button = $Shell.Popup("Success. Press ok to exit.", 0, "Done", 0)
    Remove-Item "$env:appdata\Ganter" -Force -Recurse -ErrorAction SilentlyContinue
    if ([version]$ServerVersion -gt [version]$LocalVersion) {
        & .\Updater.exe
    }
    $EndLoop = 0
}