#Save cred block, for new generation, if it will be changed
###$ADCredential = $host.ui.PromptForCredential("AcriveDirectory", "Please enter AD service account username and password.", "", "")
###$ADCredential.UserName | set-content ".\ID.adc"
###$ADCredential.Password | ConvertFrom-SecureString | set-content ".\IDP.adc"
###$SQLCredential = $host.ui.PromptForCredential("SQL", "Please enter SQL username and password.", "", "")
###$ADCredential.UserName | set-content ".\ID.sqlc"
###$ADCredential.Password | ConvertFrom-SecureString | set-content ".\IDP.sqlc"
#
#Getting encrypted domain credentials
$ADUserName = Get-Content "$WorkFolder\ID.adc"
$ADPassword = Get-Content "$WorkFolder\IDP.adc" | ConvertTo-SecureString
$Credentials = New-Object System.Management.Automation.PsCredential("$ADUserName", $ADPassword)
#Getting SQL encrypted credentials
$SqlLogin = Get-Content -Path ".\ID.sqlc"
$SQLPasswordHash = Get-Content -Path ".\IDP.sqlc"
$SqlPassw = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString $SQLPassword)))
#SQL connection block
$SqlServer = "Sql-Server-Name";
$SqlCatalog = "DatabaseName";
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server=$SqlServer; Database=$SqlCatalog; User ID=$SqlLogin; Password=$SqlPassw;"
$SqlConnection.Open()
#Getting data from AD and forwarding to SQL command
Get-ADUser -Filter "(Name -Like '*') -and (Enabled -eq 'True')" -Properties Surname,GivenName,MobilePhone,Title -Credential $Credentials | ForEach-Object {
#Temporaty variables for sql command
$GivenName = $_.GivenName
$Surname = $_.Surname
$MobilePhone = $_.MobilePhone
$JobTitle = $_.Title
#SQL command execution
$SqlCmd = $SqlConnection.CreateCommand()
$SqlCmd.CommandText = "UPDATE <TableName> `
SET Status = '$JobTitle' `
WHERE Surname = '$GivenName' `
AND Name = '$Surname' `
UPDATE <TableName> `
SET WorkPhone = '$MobilePhone' `
WHERE Surname = '$GivenName' `
AND Name = '$Surname'"
$objReader = $SqlCmd.ExecuteReader()
$objReader.close()
}
#Closing SQL connection
$SqlConnection.Close()