# Author: Moenks, Dominik
# Version: 2.5.2 (07.08.2017)
# Intention: Create security groups in AD to manage administrative rights for specified users on
# specified computers according to an Excel workbook, which are then applied to the clients by
# group policy.

$outputWidth = 75
$entries = @{}

# Open the Excel workbook for access
$excel = New-Object -ComObject Excel.Application
$excel.Workbooks.Open("$PSScriptRoot\AdministratorRights.xlsx") | Out-Null
# Check if opening the workbook was successful
Write-Host 'Checking if the configuration file is available:'.PadRight($outputWidth) -NoNewline
if (($sheet = $excel.Workbooks["AdministratorRights.xlsx"].Worksheets["AdministratorRights"]) -ne $null)
{
    Write-Host 'Succeeded' -ForegroundColor Green
    $rows = 2
    # Check all entries up to the first empty row
    Write-Host 'Checking configuration entries:'.PadRight($outputWidth) -NoNewline
    while ($sheet.Cells($rows, 1).FormulaR1C1Local -ne "")
    {
        # Read/parse values from list
        $computerDomain = $sheet.Cells($rows, 1).Text.Trim()
        $computerName = $sheet.Cells($rows, 2).Text.Trim()
        $userDomain = $sheet.Cells($rows, 3).Text.Trim()
        $userName = $sheet.Cells($rows, 4).Text.Trim()
        # Check if the start date's value is actually a date
        $dateStart = [datetime]::MaxValue
        if (-not [datetime]::TryParse($sheet.Cells($rows, 5).Text, [ref]$dateStart))
        {
            # True, so check if the value is empty
            if ($sheet.Cells($rows, 5).Text -eq "")
            {
                # True, so set the start date to the minimum available value to enable it
                $dateStart = [datetime]::MinValue
            }
            else
            {
                # True, so set the start date to the maximum available value to disable it
                $dateStart = [datetime]::MaxValue
            }
        }
        # Check if the end date's value is actually a date
        $dateEnd = [datetime]::MinValue
        if (-not [datetime]::TryParse($sheet.Cells($rows, 6).Text, [ref]$dateEnd))
        {
            # True, so check if the value is empty
            if ($sheet.Cells($rows, 6).Text -eq "")
            {
                # True, so set the end date to the maximum available value to enable it
                $dateEnd = [datetime]::MaxValue
            }
            else
            {
                # True, so set the end date to the minimum available value to disable it
                $dateEnd = [datetime]::MinValue
            }
        }
        # Check if there's already an entry for the current computer/user combination and either create or update it
        if (-not $entries.ContainsKey($computerDomain))
        {
            $entries.Add($computerDomain, @{})
        }
        if (-not $entries.$computerDomain.ContainsKey($computerName))
        {
            $entries.$computerDomain.Add($computerName, @{})
        }
        if (-not $entries.$computerDomain.$computerName.ContainsKey($userDomain))
        {
            $entries.$computerDomain.$computerName.Add($userDomain, @{})
        }
        if (-not $entries.$computerDomain.$computerName.$userDomain.ContainsKey($userName))
        {
            $entries.$computerDomain.$computerName.$userDomain.Add($userName, @{})
        }
        if (-not $entries.$computerDomain.$computerName.$userDomain.$userName.ContainsKey("Start"))
        {
            $entries.$computerDomain.$computerName.$userDomain.$userName.Add("Start", $dateStart)
        }
        elseif ($entries.$computerDomain.$computerName.$userDomain.$userName.Start -lt $dateStart)
        {
            $entries.$computerDomain.$computerName.$userDomain.$userName.Start = $dateStart
        }
        if (-not $entries.$computerDomain.$computerName.$userDomain.$userName.ContainsKey("End"))
        {
            $entries.$computerDomain.$computerName.$userDomain.$userName.Add("End", $dateEnd)
        }
        elseif ($entries.$computerDomain.$computerName.$userDomain.$userName.End -lt $dateEnd)
        {
            $entries.$computerDomain.$computerName.$userDomain.$userName.End = $dateEnd
        }
        $rows++
    }
    Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
    Write-Host ", found $($rows - 1) entries"
}
else
{
    Write-Host 'Failed' -ForegroundColor Red
}
# Close the Excel workbook
$excel.DisplayAlerts = $false
$excel.Workbooks.Close()
$excel.Quit()

foreach ($computerDomain in $entries.Keys)
{
    $computerADSI = [adsi]::new("LDAP://$([string]::Join(",", ($computerDomain.Split(".") | ForEach-Object{"DC=$_"})))")
    $groupADSI = [adsi]::new("LDAP://OU=AdministratorRights,OU=Policy,OU=Special,$([string]::Join(",", ($computerDomain.Split(".") | ForEach-Object{"DC=$_"})))")
    foreach ($computerName in $entries.$computerDomain.Keys)
    {
        # Check if the computer exists in the target domain
        Write-Host "Searching for computer $computerDomain\${computerName}:".PadRight($outputWidth) -NoNewline
        $computerADSIsearcher = [adsisearcher]::new($computerADSI, "(&(objectClass=computer)(sAMAccountName=$computerName$))")
        if ($computerADSIsearcher.FindOne() -ne $null)
        {
            Write-Host 'Succeeded' -ForegroundColor Green
            $groupName = "AdministratorRights-$($computerName.ToUpper())"
            foreach ($userDomain in $entries.$computerDomain.$computerName.Keys)
            {
                $userADSI = [adsi]::new("LDAP://DC=$([string]::Join(",DC=", $userDomain.Split(".")))")
                foreach ($userName in $entries.$computerDomain.$computerName.$userDomain.Keys)
                {
                    # Check if the user exists in the target domain
                    Write-Host "Searching for user $userDomain\${userName}:".PadRight($outputWidth) -NoNewline
                    $userADSIsearcher = [adsisearcher]::new($userADSI, "(&(objectClass=user)(sAMAccountName=$userName))")
                    if (($user = $userADSIsearcher.FindOne()) -ne $null)
                    {
                        Write-Host 'Succeeded' -ForegroundColor Green
                        # Check if the current computer/user combination is valid for today
                        Write-Host 'Checking if today is in configured timespan:'.PadRight($outputWidth) -NoNewline
                        $groupADSIsearcher = [adsisearcher]::new($groupADSI, "(&(objectClass=group)(sAMAccountName=$groupName))", @("member"))
                        if ($entries.$computerDomain.$computerName.$userDomain.$userName.Start -le [datetime]::Today -and [datetime]::Today -le $entries.$computerDomain.$computerName.$userDomain.$userName.End)
                        {
                            Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
                            Write-Host ', creating/updating the group'
                            # True, so either create or update the matching AD group
                            $groupADSIsearcher = [adsisearcher]::new($groupADSI, "(&(objectClass=group)(sAMAccountName=$groupName))", @("member"))
                            if (($group = $groupADSIsearcher.FindOne()) -eq $null)
                            {
                                $group = $groupADSI.Children.Add("CN=$groupName", "group")
                                $group.Put("groupType", 0x80000008)
                                $group.Put(“sAMAccountName”, $groupName)
                                $group.SetInfo()

                            }
                            if ($group.GetType() -eq [System.DirectoryServices.SearchResult])
                            {
                                $group = $group.GetDirectoryEntry()
                            }
                            $group.Properties["member"].Add($user.Properties["distinguishedname"][0]) | Out-Null
                            $group.Properties["description"].Value = "Used for provisioning local administrative rights"
                            $group.Properties["info"].Value = "Automatically updated on $([datetime]::Now.ToString('yyyy-MM-dd'))"
                            $group.CommitChanges()
                        }
                        else
                        {
                            Write-Host 'Failed' -ForegroundColor Red -NoNewline
                            Write-Host ', updating/deleting the group'
                            # False, so either update or delete the matching AD group
                            if (($group = $groupADSIsearcher.FindOne()) -ne $null)
                            {
                                $group = $group.GetDirectoryEntry()
                                $group.Properties["member"].Remove($user.Properties["distinguishedname"][0])
                                $group.Properties["description"].Value = "Used for provisioning local administrative rights"
                                $group.Properties["info"].Value = "Automatically updated on $([datetime]::Now.ToString('yyyy-MM-dd'))"
                                $group.CommitChanges()
                                if ($group.Properties["member"].Count -eq 0)
                                {
                                    $groupADSI.Children.Remove($group)
                                }
                            }
                        }
                    }
                    else
                    {
                        Write-Host 'Failed' -ForegroundColor Red
                    }
                }
            }
        }
        else
        {
            Write-Host 'Failed' -ForegroundColor Red
        }
    }
}
