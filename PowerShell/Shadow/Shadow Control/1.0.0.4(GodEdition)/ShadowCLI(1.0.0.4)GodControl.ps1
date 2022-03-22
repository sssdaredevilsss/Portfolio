#########################Loop#########################
:MainLoop While (2 -gt 1){
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
            ################# OLD Domain #################
            #       $ErrorActionPreference = "SilentlyContinue"
            #        $UserSamName = $(dsquery * OU=org,OU=Organization,DC=olddomain2003 -filter "(&(objectcategory=person)(objectclass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)(name=*$UNAME*))" -limit 0 -attr SamAccountName | Select-Object -Skip 1).Trim() 
            #        $ErrorActionPreference = "Continue"
            #        $UserSamName | Foreach-Object {
            #        $ChoiceData += "$_"
            #    }
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
    ####Getting session id########################
    $Counter = 1
    $Header = "Sesion", "UserName", "ID", "Status"
    $ServerList = @('s3rds1', 's3rds2', 's3rds3', 's3rds4') 
    $ServerList | ForEach-Object {
        $SiteCount = $Counter++
        $Server = $_
        $(query session /server:$_) -replace "^[\s>]" , "" -replace "\s+" , "," | Select-String $UserSamName | ConvertFrom-Csv -Header $Header | ForEach-Object {
            $id = $_.ID
            $Status = $_.Status
            if ($null -ne $id) {
                Clear-Host
                mstsc /shadow:$id /v:$Server /control /noConsentPrompt
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