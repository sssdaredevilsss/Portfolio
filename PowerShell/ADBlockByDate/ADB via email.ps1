#Gettind mail credentials
$MailLogin = 'notify@email.com' #login
#$GetMailPasswordHash = (get-credential).password | ConvertFrom-SecureString  ## Use this command to Get password hash with first run (key based on your system)
$MailPasswordHash = '01000000d08c9ddf0115d005435b753a6703c267f3c8d81f832ae2eff39757034e418b9ef80f19a23f57d9a90a0891f7b79a685d193f915fb4fd5f1edd7c1542b488d87f13821dd25590e58e822c666e11fc39a3330e4e01c231984dc49d9530a5eb45c39fcb936c81463be140000008793b728e4975a4cd886b3949d185674fd52c3fd' #some hash
$MailPassword = $MailPasswordHash | ConvertTo-SecureString 
$MailCredentials = New-Object System.Management.Automation.PsCredential($MailLogin,$MailPassword) #Secure Password
#Getting data from AD
$BannedList = Get-ADUser -Filter {Enabled -eq $TRUE} -SearchBase 'OU=Users,OU=Example,OU=Organization,DC=example,DC=com' `
-Properties Name,SamAccountName,LastLogonDate | Where {($_.LastLogonDate -lt (Get-Date).AddDays(-30)) -and ($_.LastLogonDate -ne $NULL)} | Sort | Select SamAccountName,LastLogonDate
$BannedCounter = @($BannedList).count
if (($BannedList.count -ne 0) -and ($BannedList.count -lt 2) -and ($BannedList.count -eq $null)) {$BannedCounter = 1}
$CurrentDate = Get-Date -Format 'dd.MM.yyyy HH:mm'
#Disable users and add description
$BannedList | Foreach-Object {
	Set-ADUser -Identity $_.SamAccountName -Replace @{Description="Blocked by script at ($CurrentDate)"}
    Disable-ADAccount -Identity $_.SamAccountName
}
#Send mail
$From = 'notify@email.com'
$To = 'alert@email.com'
$Subject = "Bloked users report ($CurrentDate)"
$htmlhead = "<html>
  <meta charset=utf-8>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif; charset=utf-8;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif; charset=utf-8;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif; charset=utf-8;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt; charset=utf-8;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000; charset=utf-8;}
				TD{border: 1px solid #969595; padding: 5px; charset=utf-8; }
				td.pass{background: #B7EB83; charset=utf-8;}
				td.warn{background: #FFF275; charset=utf-8;}
				td.fail{background: #FF2626; color: #ffffff; charset=utf-8;}
				td.info{background: #85D4FF; charset=utf-8;}
				</style>
				<body>
                <p>List of bloked users by lastlogon attribute:</p>"
$htmltail = "<p>Total count:$BannedCounter</p></body></html>"
$html = $BannedList | ConvertTo-Html -Fragment
$body = $htmlhead + $html + $htmltail
$SMTPServer = 'mail.server.com'
$SMTPPort = '465'
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $MailCredentials