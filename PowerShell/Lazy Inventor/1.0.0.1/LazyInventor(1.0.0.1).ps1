##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjofG5iZk2UXvWGk5UuGUrriry4C47NbkvijQWtodSlt5hRX1Cl24V+YdR+xbvdIeNQ==
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4NI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWI0g==
##OsfOAYaPHGbQvbyVvnQmqxugEiZ6Dg==
##LNzNAIWJGmPcoKHc7Do3uAu/DDhlPovK2Q==
##LNzNAIWJGnvYv7eVvnRb5FH3AkEleMCVrbm1pA==
##M9zLA5mED3nfu77Q7TV64AuzAkQqdNzbkLixwY+o8PiM
##NcDWAYKED3nfu77Q7TV64AuzAkQCDg==
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8pXZ7DZ26WfhRWEoDg==
##KNzDAJWHD2fS8u+VxDh450ribmcsZ8b7
##P8HSHYKDCX3N8u+V1TVk/EeuamkvdMSXttY=
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2UXvWGk5UuGeqr2zy5GA2+/6vinWWZcRR0BLuijmHQuUV+QXW+Eapu09VAczb9sF9LfeD+i7DfNEwq0vJbTH6LcxEDo=
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
#Main loop
:MainLoop do{
#Getting location and filenames
$WorkingLocation = Get-Location
$ActualInventoryFilename = Get-ChildItem -Name -Path "$WorkingLocation\ActualInventory"
$InventoryFilename = Get-ChildItem -Name -Path "$WorkingLocation\Inventory"
#Clearing log file
" " | Out-File -FilePath "$WorkingLocation\Inventory.log"
#Create items arrays
$FoundItems = @()
$NotFoundItems = @()
#Files count check
if (($ActualInventoryFilename.count -gt 1) -or ($InventoryFilename.count -gt 1)) {
    Write-Host ' '
    Write-Host "There is more than one file in work folders. It must be only one document in each work folder" -ForegroundColor blue -BackgroundColor red
    Write-Host ' '
    Write-Host "Press Enter to continue..."
    Read-Host 
    Clear-Host
    continue MainLoop
}
#Start logging
#Start-Transcript -Path "$WorkingLocation\Inventory.log"
#Installing modules
if (Get-Module -ListAvailable -Name ImportExcel) {
    Write-Host " "
}else{
    Write-Host "You need to install Excel module"
    Install-Module ImportExcel
}
#Import excel
$ActualInventory = Open-ExcelPackage -Path "$WorkingLocation\ActualInventory\$ActualInventoryFilename"
$ActualInventoryWorksheet = $ActualInventory.Workbook.Worksheets['Sheet1']
$Inventory = Open-ExcelPackage -Path "$WorkingLocation\Inventory\$InventoryFilename"
$InventoryWorksheet = $Inventory.Workbook.Worksheets['Sheet1']
#ActualInventory cell counter
$ActualInventoryCount = 1
#Getting curent date
$CurrentDate = Get-Date -Format "dd/MM/yyyy"
#Search start and writing data to Inventory file when in match.(Loop goes to start point if found matching data)
:StartLoop do{
$ActualInventoryCell = $ActualInventoryWorksheet.Cells["A$ActualInventoryCount"].Value
$ActualInventoryCount++
$InventoryCount = 2
do{
    $InventoryCell = $InventoryWorksheet.Cells["B$InventoryCount"].Value
    if ($ActualInventoryCell -ne $InventoryCell){
        $NotFoundItems += "$ActualInventoryCell"
        $InventoryCount++
        }elseif (($ActualInventoryCell -eq $InventoryCell) -and ($null -ne $InventoryCell)){
            $MatchCount++
            $FoundItems += "$ActualInventoryCell"
            $InventoryWorksheet.Cells["I$InventoryCount"].Value = "$CurrentDate"
            $InventoryWorksheet.Cells["H$InventoryCount"].Value = 'Vasyl Hadzalo'
            continue StartLoop
        }
}until ($null -eq $InventoryCell)
}until ($null -eq $ActualInventoryCell)
#Closing exchel process
Close-ExcelPackage $ActualInventory
Close-ExcelPackage $Inventory
#Stop logging
#Stop-Transcript
#Appeding log with count of found matches 
" " | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
" " | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
"*************************************************" | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
" " | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
"We found $MatchCount matches during current job" | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
"List of founds:" | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
$FoundItems | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
" " | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
"*************************************************" | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
#Appeding log with count of not found matches 
" " | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
" " | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
"*************************************************" | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
" " | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
$UniqeNotFoundItems = $NotFoundItems | Sort-Object -Unique | Where-Object {$_ -notin $FoundItems}
$UniqeNotFoundItemsCount = $UniqeNotFoundItems.count - 1
"Not found $UniqeNotFoundItemsCount" | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
"List:" | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
$UniqeNotFoundItems | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
" " | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
"*************************************************" | Out-File -FilePath "$WorkingLocation\Inventory.log" -Append
}until(($ActualInventoryFilename.count -eq 1) -and ($InventoryFilename.count -eq 1))
$Shell = New-Object -ComObject "WScript.Shell"
$Shell.Popup("Success. Press ok to exit.", 0, "Done", 0)