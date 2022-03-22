#Substring func, couse no Fucking methon invokation avaible in constr. mode
#########################Collecting Statistics #########################
Start-Job -Name "stat" -ScriptBlock {
    
function substr ([string] $String, [int] $Start = 0, [int] $Length = -1) {
    if ($Length -eq -1 -or $Start + $Length -ge $String.Length) {
      $String.Substring($Start)
    }
    else {
      $String.Substring($Start, $Length)
    }
  }
    do {
        $Statistic = @()
        $ServerList = @('s3rds1', 's3rds2', 's3rds3', 's3rds4')
        $ServerList | ForEach-Object {
            $CurrentServer = $_
            #########################UserCounters #########################
            $RDSTotalUserCount = (query user /server:$CurrentServer | Select-String -Pattern "Active", "Disc" ).count
            $RDSActiveUserCount = (query user /server:$CurrentServer | Select-String "Active" ).count
            $RDSDisconectedCount = (query user /server:$CurrentServer | Select-String "Disc" ).count
            #########################CPU #########################
            $CPUAveragePerformance = (GET-COUNTER -Counter "\Processor(_Total)\% Processor Time" -ComputerName $CurrentServer | Select-Object -ExpandProperty countersamples | Select-Object -ExpandProperty cookedvalue | Measure-Object -Average).average
            #$CPUAveragePerformanceRounded = [math]::Round($CPUAveragePerformance, 2)
            $CPUAveragePerformanceRounded = substr "$CPUAveragePerformance" -Length 4
            #########################RAM #########################
            $TotalPhysicalMemory = 68719476736 / 1MB # yes, hardcode, but no need to elevating mode
            $AvaibleMemory = (Get-Counter -Counter "\Memory\Available MBytes" -ComputerName $CurrentServer).CounterSamples.CookedValue
            $UserMemoryMB = (($TotalPhysicalMemory - $AvaibleMemory) / $TotalPhysicalMemory) * 100
            #$UserMemoryGB = [math]::Round($UserMemoryMB, 2)
            $UserMemoryGB = substr "$UserMemoryMB" -Length 4
            #########################HashTable #########################
            $CurrentDate = Get-Date -Format "HH:mm:ss"
            $CurrentHash = New-Object PSObject -property @{Server = "$CurrentServer"; Users = "$RDSTotalUserCount"; Active = "$RDSActiveUserCount"; Disc = "$RDSDisconectedCount"; CPU = "$CPUAveragePerformanceRounded"; RAM = "$UserMemoryGB"; "%" = '%'; "Collection Time" = "`0`0`0`0`0`0`0$CurrentDate" }
            $Statistic += $CurrentHash
        }
        Start-Sleep -Seconds 60
        $Statistic | Format-Table -Property Server, Users, Active, Disc, CPU, %, RAM, %, "Collection Time" -AutoSize | Out-File $env:appdata\Shadow.stat
    }while (1 -eq 1)
} | Out-Null
#########################Renew dataset#########################
#$ErrorActionPreference = "SilentlyContinue"
#$LastCollectionDate = (Get-Item -Path "$env:appdata\Shadow.stat").LastWriteTime
#if ($LastCollectionDate.ToString("yyyy:MM:dd") -lt ((Get-Date).ToString("yyyy:MM:dd"))) {
#    Remove-Item -Path "$env:appdata\Shadow.stat" -Force
#}
#$ErrorActionPreference = "Continue"
$TotalSecs = 75
$VerbosePreference = "Continue"
while (-not (Test-Path -Path "$env:appdata\Shadow.stat")) {
    $TotalSecs--
    Write-Verbose -Message "Its first run! Wait for recollecting data, please be patient [$TotalSecs] seconds..."
    Start-Sleep -Seconds 1
    Clear-Host
}
$VerbosePreference = "SilentlyContinue"
#########################Loop#########################
:MainLoop While (2 -gt 1) {
    $GoSearch = $null
    #########################Server Choice Menu#########################
    do {
        Get-Content $env:appdata\Shadow.stat
        Write-Host "`n================ Whose turn to Suffer? ====================="
        Write-Host "`t'0' `tSpecific user"
        Write-Host "`t'1' `tS3RDS1"
        Write-Host "`t'2' `tS3RDS2"
        Write-Host "`t'3' `tS3RDS3"
        Write-Host "`t'4' `tS3RDS4"
        Write-Host "`t'5' `tAll Terminals"
        Write-Host "============================================================"
        $choice = Read-Host "`nEnter Choice"
        Clear-Host
    } until (($choice -in 1..5) -or ($choice -eq 0))
    switch ($choice) {
        '0' {
            Write-Host "`nYou have selected Specific user"
            $GoSearch = 0
        }
        '1' {
            Write-Host "`nYou have selected S3RDS1"
            $ServerList = "S3RDS1"
        }
        '2' {
            Write-Host "`nYou have selected S3RDS2"
            $ServerList = "S3RDS2"
        }
        '3' {
            Write-Host "`nYou have selected S3RDS3"
            $ServerList = "S3RDS3"
        }
        '4' {
            Write-Host "`nYou have selected S3RDS4"
            $ServerList = "S3RDS4"
        }
        '5' {
            Write-Host "`nYou have selected All Terminals"
            $ServerList = @('s3rds1', 's3rds2', 's3rds3', 's3rds4') 
        }
    }
    #########################################################
    if ($GoSearch -eq 0) {
        #########################Checking and searching user in AD#########################
        $DomainList = @('one.domain', 'two.domain') 
        $Counter = 0
        $id = $null
        do {
            do {
                do {
                    $UNAME = Read-Host -Prompt "Enter User name"
                    if ($UNAME.Length -eq 0) {
                        Clear-Host
                    }
                }while ($UNAME.Length -eq 0)
                $ChoiceData = @()
                $DomainList | ForEach-Object {
                    $Counter++
                    $UserSamName = (Get-ADUser -server $_ -Filter "(Name -Like '*$UNAME*') -and (Enabled -eq 'True')").SamAccountName
                    if ($UserSamName -ne $null) {
                        $UserSamName | Foreach-Object {
                            $ChoiceData += "$_"
                        }
                    }
                }
#                $ErrorActionPreference = "SilentlyContinue"
#                $UserSamName = $(dsquery * OU=Org, OU=Organization, DC=olddomain2003 -filter "(&(objectcategory=person)(objectclass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)(name=*$UNAME*))" -limit 0 -attr SamAccountName | Select-Object -Skip 1).Trim() 
#                $ErrorActionPreference = "Continue"
#                $UserSamName | Foreach-Object {
#                    $ChoiceData += "$_"
#                }
                if ($ChoiceData.count -eq 1) {
                    $UserSamName = $ChoiceData
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
                        [int]$global:ans = Read-Host ' (Select "0" to search again) Choose the right one'
                        if ($global:ans -eq 0){
                            Clear-Host
                            Continue MainLoop
                        }
                    }
                    $global:selection = $menu.Item([int]$ans)
                    $UserSamName = $global:selection
                }
            } While (!$ChoiceData -and $Counter -eq $DomainList.Count)
            if (!$ChoiceData) {
                Clear-Host
                Write-Host "No such User. Try to search in another way or check if it exist or enabled" -ForegroundColor blue -BackgroundColor red
            }
        }Until ($null -ne $ChoiceData)
        ####Getting session id and logoff########################
        $Counter = 1
        $Header = "Sesion", "UserName", "ID", "Status"
        $ServerList = @('s3rds1', 's3rds2', 's3rds3', 's3rds4') 
        $ServerList | ForEach-Object {
            $SiteCount = $Counter++
            $Server = $_
            $(query session /server:$_) -replace "^[\s>]" , "" -replace "\s+" , "," | Select-String $UserSamName | ConvertFrom-Csv -Header $Header | ForEach-Object {
                $id = $_.ID
                $Status = $_.Status
                if ($id -ne $null) {
                    Clear-Host
                    logoff $id /server:$Server
                    if ($? -eq $True) {
                        Write-Host "User $UserSamName logged out successfully" -ForegroundColor blue -BackgroundColor green
                        Start-Sleep 3
                    }
                }
            }
        }
        if (($null -eq $id) -and ($SiteCount -eq $ServerList.Count)) {
            Clear-Host
            Write-Host # empty line
            Write-Host "User is not connected." -ForegroundColor blue -BackgroundColor red
            Write-Host # empty line 
        }
    }
    else {
        #########################################################
        #########################Filter Choice Menu#########################
        do {
            Write-Host "`n================ Which state of users? ====================="
            Write-Host "`t'1' `tOnly Disconnected"
            Write-Host "`t'2' `tAll users"
            Write-Host "============================================================"
            $choice = Read-Host "`nEnter Choice"
        } until ($choice -in 1..2)
        switch ($choice) {
            '1' {
                Write-Host "`nYou have selected Only Disconnected"
                $StateChoise = "0"
            }
            '2' {
                Write-Host "`nYou have selected All users"
                $StateChoise = "1"
            }
        }
        #########################Logoff#########################
        if ($StateChoise -eq 0) {
            do {
                $Confirmation = Read-Host "Do you really want to logoff all disconnected users frorm $ServerList ? [y/n]"
                if ($Confirmation -eq 'n') {
                    Clear-Host
                    continue MainLoop
                }
            }while ($Confirmation -ne "y")
            $Header = "Sesion", "UserName", "ID", "Status"
            $ServerList | ForEach-Object {
                $Server = $_
                $Query = ($(query session /server:$Server) -replace "^[\s>]" , "" -replace "\s+" , "," | Select-String -NotMatch "$env:username" | Select-String -Pattern "Disc" | ConvertFrom-Csv -Header $Header)
                $TotalTimes = $Query.Count
                $i = 0
                $Query.ID | ForEach-Object {
                    $i++
                    logoff $_ /server:$Server
                    $percentComplete = ($i / $totalTimes) * 100
                    Write-Progress -Activity "Total user count: $TotalTimes" -Status "Progres: $i" -PercentComplete $percentComplete -CurrentOperation "Server: $Server | State: Disconected users"
                }
            }
        }
        elseif ($StateChoise -eq 1) {
            do {
                $Confirmation = Read-Host "Do you really want to logoff all users frorm $ServerList ? [y/n]"
                if ($Confirmation -eq 'n') {
                    Clear-Host
                    continue MainLoop
                }
            }while ($Confirmation -ne "y")
            $Header = "Sesion", "UserName", "ID", "Status"
            $ServerList | ForEach-Object {
                $Server = $_
                $Query = ($(query session /server:$Server) -replace "^[\s>]" , "" -replace "\s+" , "," | Select-String -NotMatch "$env:username" | ConvertFrom-Csv -Header $Header)
                $TotalTimes = $Query.Count
                $i = 0
                $Query.ID | ForEach-Object {
                    $i++
                    logoff $_ /server:$Server
                    $percentComplete = ($i / $totalTimes) * 100
                    Write-Progress -Activity "Total user count: $TotalTimes" -Status "Progres: $i" -PercentComplete $percentComplete -CurrentOperation "Server: $Server | State: All users"
                }
            }
        }
    }
    Start-Sleep -Seconds 3
    Clear-Host
} #end of main loop