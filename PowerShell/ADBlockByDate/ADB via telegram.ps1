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
Set-Content -Path $env:TEMP\blocked.txt -Value $BannedList
$ConvertedList = Get-Content -Path $env:TEMP\blocked.txt
$FilteredConvertedList = $ConvertedList -replace '@{SamAccountName=',"`nUsername: " -replace ';','' -replace 'LastLogonDate=',"`n`LogonDate: " -replace '}',"`n"
#Telegram Notify
$MyToken = "16118356:AAErEnsfnK3COX09_f5oA11yJUE" #Your token
$chatID = 2701361234 # Bot`s chatID
$date = Get-Date -Format "dd.MM.yyyy HH:mm"
$Message = "Task: Block by Last Logon `
Server: dc0.brainsetter.com `
Target: ActiveDirectory `
Date: $date `
________________________________`
List:   (Total: $BannedCounter)`
$FilteredConvertedList" #Messege ends
$Response = Invoke-RestMethod -Uri "https://api.telegram.org/bot$($MyToken)/sendMessage?chat_id=$($chatID)&text=$($Message)"