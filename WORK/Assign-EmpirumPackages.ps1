<#
.SYNOPSIS
This script is intended to manage computer and software assignments in matrix42 Empirum.
.DESCRIPTION
This script manages computer and software assignments in matrix42 Empirum.
.PARAMETER ComputerNames
.PARAMETER AssignCurrentSoftware
.PARAMETER AssignAdditionalSoftware
.PARAMETER AssignAdditionalSoftwareRegister
.PARAMETER AssignmentRoot
.PARAMETER ComputerGroupBlockSize
.PARAMETER RolloutBlockSize
.PARAMETER RolloutInventoryAge
.PARAMETER UseExistingAssignmentGroup
.PARAMETER EMPserver
.PARAMETER SQLserver
.PARAMETER SQLdatabase
.PARAMETER Username
.PARAMETER Password
.EXAMPLE
EmpAgent.ps1 -DHCPOptionNumber 128 -FallbackServer 'EmpirumMaster' -WeeklyErrorThreshold 10 -EmpAgentBatch 'EmpirumAgent.bat' -EmpInventoryBatch 'EmpirumInventory.bat'
.NOTES
Version:    3.1
Author:     Mönks, Dominik
#>

param([Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerNames,
        [switch]$AssignCurrentSoftware,
        [ValidateNotNullOrEmpty()]
        [string[]]$AssignAdditionalSoftware,
        [ValidateNotNullOrEmpty()]
        [string[]]$AssignAdditionalSoftwareRegister,
        [ValidateNotNullOrEmpty()]
        [string]$AssignmentRoot,
        [ValidateNotNullOrEmpty()]
        [int]$ComputerGroupBlockSize = 500,
        [ValidateNotNullOrEmpty()]
        [int]$RolloutBlockSize = 10,
        [ValidateNotNullOrEmpty()]
        [int]$RolloutInventoryAge = 7,
        [switch]$UseExistingAssignmentGroup,
        [ValidateNotNullOrEmpty()]
        [string]$EMPserver,
        [ValidateNotNullOrEmpty()]
        [string]$SQLserver,
        [ValidateNotNullOrEmpty()]
        [string]$SQLdatabase,
        [ValidateNotNullOrEmpty()]
        [string]$Username,
        [ValidateNotNullOrEmpty()]
        [SecureString]$Password)

$outputWidth = 75

$timeout = [timespan]::FromHours(1)
$starttime = [datetime]::Now
$averagetime = [timespan]::Zero

#region: Imports
$definition = @"
[DllImport("advapi32.dll", SetLastError = true)]
public static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);
[DllImport("advapi32.dll", SetLastError = true)]
public static extern bool ImpersonateLoggedOnUser(IntPtr phToken);
[DllImport("advapi32.dll", SetLastError = true)]
public static extern bool RevertToSelf();
"@
Add-Type -MemberDefinition $definition -name "DLL" -namespace "Import" -ErrorAction SilentlyContinue
#endregion

#region: Check for installed Empirum SDK
try
{
    Write-Host "Checking for installed Empirum SDK:".PadRight($outputWidth) -NoNewline
    Get-Command "Open-Matrix42ServiceConnection" -ErrorAction Stop | Out-Null
    $SDK = $true
    Write-Host "Succeeded" -ForegroundColor Green
}
catch [Management.Automation.CommandNotFoundException]
{
    $SDK = $false
    Write-Host "Failed" -ForegroundColor Red -NoNewline
}
#endregion
if ($SDK)
{
    #region: Create sessions
    try
    {
        Write-Host "Checking credentials:".PadRight($outputWidth) -NoNewline
        if ($Username -eq "" -or $Password -eq "")
        {
            $credentials = (Get-Credential "$env:USERDOMAIN\$env:USERNAME").GetNetworkCredential()
            Write-Host "Succeeded" -ForegroundColor Green -NoNewline
            Write-Host ", credentials were provided by GUI"
        }
        else
        {
            $credentials = [Net.NetworkCredential]::new($Username.Split("\")[1], $Password, $Username.Split("\")[0])
            Write-Host "Succeeded" -ForegroundColor Green -NoNewline
            Write-Host ", credentials were provided as parameters"
        }
        Write-Host "Authenticating with Empirum server:".PadRight($outputWidth) -NoNewline
        try
        {
            $EMPsession = Open-Matrix42ServiceConnection -ServerName $EMPserver -Port 9200 -UserName "$($credentials.Domain)\$($credentials.UserName)" -Password $credentials.Password -ErrorAction Stop
            Write-Host "Succeeded" -ForegroundColor Green
        }
        catch [Matrix42.SDK.ServiceContracts.Matrix42ServiceException]
        {
            Write-Host "Failed" -ForegroundColor Red
        }
        Write-Host "Authenticating with Empirum database:".PadRight($outputWidth) -NoNewline
        try
        {
            $token = [IntPtr]::Zero
            [Import.DLL]::LogonUser($($credentials.UserName),$($credentials.Domain),$($credentials.Password),9,0,[ref]$token) | Out-Null
            [Import.DLL]::ImpersonateLoggedOnUser($token) | Out-Null
            $SQLsession = [Data.SqlClient.SqlConnection]::new("Server=$SQLserver;Database=$SQLdatabase;Integrated Security=SSPI")
            $SQLsession.Open()
            [Import.DLL]::RevertToSelf() | Out-Null
            Write-Host "Succeeded" -ForegroundColor Green
        }
        catch [InvalidOperationException], [Data.SqlClient.SqlException]
        {
            Write-Host "Failed" -ForegroundColor Red
        }
    }
    catch [Management.Automation.ParameterBindingException]
    {
        Write-Host "Failed" -ForegroundColor Red -NoNewline
        Write-Host ", no credentials were provided"
    }
    #endregion
    if ($EMPsession -ne $null -and $SQLsession.State -eq [Data.ConnectionState]::Open)
    {
        if ($ComputerNames.Count -eq 1 -and $ComputerNames[0].Contains(","))
        {
            $ComputerNames = $ComputerNames -split ","
        }
        $RolloutCounter = 1
        Write-Host "Updating Empirum computers:".PadRight($outputWidth) -NoNewline
        $EMPcomputers = $EMPsession.Computers
        Write-Host "Succeeded" -ForegroundColor Green
        Write-Host "Updating Empirum groups:".PadRight($outputWidth) -NoNewline
        $EMPgroups = $EMPsession.Groups
        Write-Host "Succeeded" -ForegroundColor Green
        Write-Host "Updating Empirum packages:".PadRight($outputWidth) -NoNewline
        $EMPpackages = $EMPsession.Packages
        Write-Host "Succeeded" -ForegroundColor Green
        foreach ($ComputerName in $ComputerNames)
        {
            $RolloutStatus = $true
            Write-Host ""
            Write-Host "Checking rollout status:".PadRight($outputWidth) -NoNewline
            if ($RolloutCounter -le $RolloutBlockSize)
            {
                Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                Write-Host ", counter within limits ($RolloutCounter/$RolloutBlockSize)"
                Write-Host "Checking set '$($ComputerName.ToUpper())':".PadRight($outputWidth) -NoNewline
                if (($match = [regex]::Match($ComputerName, "([-a-z]+\d+)(?:\s*>\s*([-a-z]+\d+))?", [Text.RegularExpressions.RegexOptions]::IgnoreCase)).Success)
                {
                    #region: Check computer name
                    Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                    Write-Host ", found the following values: $(($match.Groups[1..2].Value.ToUpper() | Where-Object{$_ -ne ''}) -join ", ")"
                    if ($match.Groups[2].Success)
                    {
                        Write-Host "Searching for source computer $($match.Groups[1].Value.ToUpper()):".PadRight($outputWidth) -NoNewline
                        $SourceComputer = $EMPcomputers | Where-Object{$_.Name -eq $match.Groups[1].Value}
                        if ($SourceComputer -ne $null)
                        {
                            Write-Host "Succeeded" -ForegroundColor Green
                        }
                        else
                        {
                            Write-Host "Failed" -ForegroundColor Red
                        }
                        Write-Host "Searching for target computer $($match.Groups[2].Value.ToUpper()):".PadRight($outputWidth) -NoNewline
                        $TargetComputer = $EMPcomputers | Where-Object{$_.Name -eq $match.Groups[2].Value} | Sort-Object ID | Select-Object -Last 1
                        if ($TargetComputer -ne $null)
                        {
                            Write-Host "Succeeded" -ForegroundColor Green
                        }
                        else
                        {
                            Write-Host "Failed" -ForegroundColor Red
                        }
                    }
                    else
                    {
                        Write-Host "Searching for computer $($match.Groups[1].Value.ToUpper()):".PadRight($outputWidth) -NoNewline
                        $SourceComputer = $TargetComputer = $EMPcomputers | Where-Object{$_.Name -eq $match.Groups[1].Value} | Sort-Object ID | Select-Object -Last 1
                        if ($SourceComputer -ne $null -and $TargetComputer -ne $null)
                        {
                            Write-Host "Succeeded" -ForegroundColor Green
                        }
                        else
                        {
                            Write-Host "Failed" -ForegroundColor Red
                        }
                    }
                    #endregion
                    if ($SourceComputer -ne $null -and $TargetComputer -ne $null)
                    {
                        #region: Assign computer
                        $computerNode = $null                        
                        if ($UseExistingAssignmentGroup)
                        {
                            Write-Host "Searching for assignment group (computer level, current structure):".PadRight($outputWidth) -NoNewline
                            if (($computerNodes = $EMPgroups | Where-Object{$_.Name -eq $TargetComputer.Name -and $_.GroupType -eq [Matrix42.SDK.Contracts.Models.GroupType]::AssignmentGroup}) -eq $null)
                            {
                                Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                                Write-Host ", new structure will be created"
                            }
                            else
                            {
                                Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                                Write-Host ", using following structure:"
                                foreach ($computerNode in $computerNodes)
                                {
                                    $structure = [Collections.Stack]::new()
                                    $structure.Push($computerNode)
                                    while ($structure.Peek().ParentGroupId -ne $null)
                                    {
                                        $structure.Push(($EMPgroups | Where-Object{$_.Id -eq $structure.Peek().ParentGroupId}))
                                    }
                                    if ($structure.Name.Contains($TargetComputer.DomainName))
                                    {
                                        break
                                    }
                                }
                                $structureLevels = $structure.Count
                                while ($structure.Count -gt 0)
                                {
                                    Write-Host "$(''.PadLeft($outputWidth).PadRight($outputWidth + $structureLevels - $structure.Count + 1, '-')) $($structure.Pop().Name)"
                                }
                            }
                        }
                        if ($computerNode -eq $null)
                        {
                            Write-Host "Searching for assignment group (root level):".PadRight($outputWidth) -NoNewline
                            if (($rootNode = $EMPgroups | Where-Object{$_.Name -eq $AssignmentRoot}) -eq $null)
                            {
                                Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                                Write-Host ", creating new assignment group"
                                $rootNode = New-EmpirumGroup -Name $AssignmentRoot -GroupType AssignmentGroup -Session $EMPsession
                            }
                            else
                            {
                                Write-Host "Succeeded" -ForegroundColor Green
                            }
                            Write-Host "Searching for assignment group (domain level):".PadRight($outputWidth) -NoNewline
                            if (($domainNode = $EMPgroups | Where-Object{$_.ParentGroupId -eq $rootNode.Id -and $_.Name -like $TargetComputer.DomainName}) -eq $null)
                            {
                                Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                                Write-Host ", creating new assignment group"
                                $domainNode = New-EmpirumGroup -Name $TargetComputer.DomainName.ToUpper() -ParentGroup $rootNode -GroupType AssignmentGroup -Session $EMPsession
                            }
                            else
                            {
                                Write-Host "Succeeded" -ForegroundColor Green
                            }
                            Write-Host "Searching for assignment group (block level):".PadRight($outputWidth) -NoNewline
                            if (($match = [regex]::Match($TargetComputer.Name, "(\D+)(\d+)")).groups.count -eq 3)
                            {
                                $blockPrefix = $match.Groups[1].Value.ToUpper()
                                $blockLow = [int]::Parse($match.Groups[2].Value) - [int]::Parse($match.Groups[2].Value) % $ComputerGroupBlockSize
                                $blockHigh = $blockLow + $ComputerGroupBlockSize - 1
                                $blockName =  "$blockPrefix$($blockLow.ToString().PadLeft(4, "0"))-$($blockHigh.ToString().PadLeft(4, "0"))"
                            }
                            else
                            {
                                $blockName = "Other"
                            }
                            if (($blockNode = $EMPgroups | Where-Object{$_.ParentGroupId -eq $domainNode.Id -and $_.Name -like $blockName}) -eq $null)
                            {
                                Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                                Write-Host ", creating new assignment group"
                                $blockNode = New-EmpirumGroup -Name $blockName -ParentGroup $domainNode -GroupType AssignmentGroup -Session $EMPsession
                            }
                            else
                            {
                                Write-Host "Succeeded" -ForegroundColor Green
                            }
                            Write-Host "Searching for assignment group (computer level):".PadRight($outputWidth) -NoNewline
                            if (($computerNode = $EMPgroups | Where-Object{$_.ParentGroupId -eq $blockNode.Id -and $_.Name -like $TargetComputer.Name}) -eq $null)
                            {
                                Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                                Write-Host ", creating new assignment group"
                                $computerNode = New-EmpirumGroup -Name $TargetComputer.Name.ToUpper() -ParentGroup $blockNode -GroupType AssignmentGroup -Session $EMPsession
                            }
                            else
                            {
                                Write-Host "Succeeded" -ForegroundColor Green
                            }
                            $EMPgroups = $EMPsession.Groups
                        }
                        if ((Find-EmpirumGroup -HasMember $TargetComputer -Session $EMPsession).ParentGroupId -notcontains $computerNode.Id)
                        {
                            Add-EmpirumComputerToGroup -Computer $TargetComputer -Group $computerNode -Session $EMPsession
                        }
                        #endregion
                        #region: Assign currently installed software
                        if ($AssignCurrentSoftware)
                        {
                            $SQLquery_installedpackages = [Data.SqlClient.SqlCommand]::new("SELECT DISTINCT SW_Relations_new.SoftwareID
                                                                                                    FROM SW_Relations_new, SW_Relations_Update_new, (SELECT SW_Relations_Group_new.LatestSoftwareID AS ID
                                                                                                                                                    FROM InvSoftware, Clients, SW_Relations_new, SW_Relations_Group_new
							                                                                                                                        WHERE InvSoftware.SwType = 'Empirum'
							                                                                                                                        AND InvSoftware.Developer <> 'matrix42'
							                                                                                                                        AND InvSoftware.Client_id = '$($SourceComputer.Id)'
							                                                                                                                        AND InvSoftware.SoftwareID = SW_Relations_new.SoftwareID
							                                                                                                                        AND InvSoftware.ProductNameShort = SW_Relations_new.ProductnameDetailed
							                                                                                                                        AND InvSoftware.Version = SW_Relations_new.Version
							                                                                                                                        AND SW_Relations_new.GroupID = SW_Relations_Group_new.GroupID
							                                                                                                                        AND SW_Relations_Group_new.ShowGroupOnList = 1
							                                                                                                                        AND SW_Relations_Group_new.IsSupported = 1) LatestSoftware
                                                                                                    WHERE SW_Relations_Update_new.SoftwareID = LatestSoftware.ID
                                                                                                    AND SW_Relations_Update_new.UpdatePackageID = SW_Relations_new.SoftwareID
                                                                                                    OR LatestSoftware.ID = SW_Relations_new.SoftwareID", $SQLsession)
                            $SQLreader = $SQLquery_installedpackages.ExecuteReader()
                            if ($SQLreader.HasRows)
                            {
                                Write-Host "Assigning installed packages..."
                                while ($SQLreader.Read())
                                {
                                    if (($package = $EMPpackages | Where-Object{$_.Id -eq $SQLreader["softwareid"]}) -eq $null)
                                    {
                                        if (($package = $EMPpackages | Where-Object{$_.Name -like "*$($SQLreader["productname"])*"} | Sort-Object -Property Version -Descending | Select-Object -First 1) -eq $null)
                                        {
                                            Write-Host "- $($SQLreader["productname"]):".PadRight($outputWidth) -NoNewline
                                            Write-Host "Failed" -ForegroundColor Red -NoNewline
                                            Write-Host ", please assign manually"
                                        }
                                    }
                                    if ($package -ne $null)
                                    {
                                        Write-Host "- $($package.Vendor) $($package.Name) $($package.Version):".PadRight($outputWidth) -NoNewline
                                        Add-EmpirumPackageToGroup -Package $package -Group $computerNode -Session $EMPsession
                                        Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                                        $state = Get-EmpirumPackageState -Package $package -TargetObject $TargetComputer -Session $EMPsession
                                        Write-Host " (current installation status: " -NoNewline
                                        switch ($state.ExecutionStatus)
                                        {
                                            Finished
                                            {
                                                Write-Host "Finished" -ForegroundColor Green -NoNewline
                                                Write-Host ", $($state.FinishedTime.ToShortDateString())" -NoNewline
                                            }
                                            Running
                                            {
                                                Write-Host "Running" -ForegroundColor Green -NoNewline
                                            }
                                            Unknown
                                            {
                                                Write-Host "Unknown" -ForegroundColor Yellow -NoNewline
                                            }
                                            Cancelled
                                            {
                                                Write-Host "Cancelled" -ForegroundColor Red -NoNewline
                                            }
                                            Failed
                                            {
                                                Write-Host "Failed" -ForegroundColor Red -NoNewline
                                                Write-Host ", $($state.ErrorText))" -NoNewline
                                            }
                                            default
                                            {
                                                Write-Host "Unknown" -ForegroundColor Yellow -NoNewline
                                            }
                                        }
                                        Write-Host ")"
                                        if ($state.ExecutionStatus -ne "Finished")
                                        {
                                            $RolloutStatus = $RolloutStatus -and $false
                                        }
                                    }

                                }
                            }
                            $SQLreader.Close()
                        }
                        #endregion
                        #region: Assign additional software
                        if ($AssignAdditionalSoftware -ne $null)
                        {
                            Write-Host "Assigning additional packages..."
                            foreach ($packagename in $AssignAdditionalSoftware)
                            {
                                switch -Wildcard ($packagename)
                                {
                                    "D:*"
                                    {
                                        $delete = $true
                                        $reinstall = $false
                                        $uninstall = $false
                                        $packagename = $packagename.Substring(2)
                                    }
                                    "R:*"
                                    {
                                        $delete = $false
                                        $reinstall = $true
                                        $uninstall = $false
                                        $packagename = $packagename.Substring(2)
                                    }
                                    "U:*"
                                    {
                                        $delete = $false
                                        $reinstall = $false
                                        $uninstall = $true
                                        $packagename = $packagename.Substring(2)
                                    }
                                    default
                                    {
                                        $delete = $false
                                        $reinstall = $false
                                        $uninstall = $false
                                    }
                                }
                                if (($package = $EMPpackages | Where-Object{$_.Id -eq $packagename}) -eq $null)
                                {
                                    if (($package = $EMPpackages | Where-Object{(@($_.Vendor, $_.Name, $_.Version) -join " ") -like "*$packagename*"} | Sort-Object -Property Version -Descending | Select-Object -First 1) -eq $null)
                                    {
                                        Write-Host "- {$packagename}:".PadRight($outputWidth) -NoNewline
                                        Write-Host "Failed" -ForegroundColor Red -NoNewline
                                        Write-Host ", please assign manually"
                                    }
                                }
                                if ($package -ne $null)
                                {
                                    Write-Host "- $($package.Vendor) $($package.Name) $($package.Version):".PadRight($outputWidth) -NoNewline
                                    if ($delete)
                                    {
                                        try
                                        {
                                            Remove-EmpirumPackageFromGroup -Package $package -Group $computerNode -Session $EMPsession
                                            Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                                            Write-Host ", deleted assignment (was to be deleted" -NoNewline
                                        }
                                        catch [Matrix42.SDK.Contracts.Matrix42SdkException]
                                        {
                                            Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                                            Write-Host ", found no assignment (was to be deleted" -NoNewline
                                        }
                                    }
                                    else
                                    {
                                        Add-EmpirumPackageToGroup -Package $package -Group $computerNode -Session $EMPsession
                                        if ($reinstall)
                                        {
                                            if ((Find-EmpirumGroup -HasMember $TargetComputer -Session $EMPsession).ParentGroupId -notcontains $computerNode.Id)
                                            {
                                                Invoke-EmpirumPackageReinstallation -Group $computerNode -Package $package -Computer $TargetComputer -Pull -Session $EMPsession
                                                Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                                                Write-Host ", invoked reinstallation (was to be reinstalled" -NoNewline
                                                $EMPgroups = $EMPsession.Groups
                                            }
                                            else
                                            {
                                                Set-EmpirumDistributionCommands -Group $computerNode -Package $package -Install -Update -Revoke 3 -Session $EMPsession
                                                Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                                                Write-Host ", invoked installation (was to be reinstalled" -NoNewline
                                            }
                                        }
                                        elseif ($uninstall)
                                        {
                                            Set-EmpirumDistributionCommands -Group $computerNode -Package $package -Uninstall -Session $EMPsession
                                            Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                                            Write-Host ", invoked uninstallation (was to be uninstalled" -NoNewline
                                        }
                                        else
                                        {
                                            Set-EmpirumDistributionCommands -Group $computerNode -Package $package -Install -Update -Revoke 3 -Session $EMPsession
                                            Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                                            Write-Host ", invoked installation (was to be installed" -NoNewline
                                        }
                                    }
                                    $state = Get-EmpirumPackageState -Package $package -TargetObject $TargetComputer -Session $EMPsession
                                    Write-Host ", current installation status: " -NoNewline
                                    switch ($state.ExecutionStatus)
                                    {
                                        Finished
                                        {
                                            Write-Host "Finished" -ForegroundColor Green -NoNewline
                                            Write-Host ", $($state.FinishedTime.ToShortDateString())" -NoNewline
                                        }
                                        Running
                                        {
                                            Write-Host "Running" -ForegroundColor Green -NoNewline
                                        }
                                        Unknown
                                        {
                                            Write-Host "Unknown" -ForegroundColor Yellow -NoNewline
                                        }
                                        Cancelled
                                        {
                                            Write-Host "Cancelled" -ForegroundColor Red -NoNewline
                                        }
                                        Failed
                                        {
                                            Write-Host "Failed" -ForegroundColor Red -NoNewline
                                            Write-Host ", $($state.ErrorText))" -NoNewline
                                        }
                                        default
                                        {
                                            Write-Host "Unknown" -ForegroundColor Yellow -NoNewline
                                        }
                                    }
                                    Write-Host ")"
                                    if (-not $delete -and $state.ExecutionStatus -ne "Finished" -and $TargetComputer.LastInventory -gt [datetime]::Now.AddDays(-$RolloutInventoryAge))
                                    {
                                        $RolloutStatus = $RolloutStatus -and $false
                                    }
                                }
                            }
                        }
                        #endregion
                        #region: Assign additional software registers
                        if ($AssignAdditionalSoftwareRegister -ne $null)
                        {
                            $SQLquery_registerpackages = [Data.SqlClient.SqlCommand]::new("SELECT softwareid, softwarename
                                                                                                FROM software
                                                                                                INNER JOIN softwaredepotregister
                                                                                                ON software.parentid = softwaredepotregister.idpk
                                                                                                WHERE software.type = 'app'
                                                                                                AND softwaredepotregister.name IN ($(($AssignAdditionalSoftwareRegister | ForEach-Object{"'" + $_.Trim() + "'"}) -join ","))", $SQLsession)
                            $SQLreader = $SQLquery_registerpackages.ExecuteReader()
                            if ($SQLreader.HasRows)
                            {
                                Write-Host "Assigning default packages..."
                                while ($SQLreader.Read())
                                {
                                    if (($package = $EMPpackages | Where-Object{$_.Id -eq $SQLreader["softwareid"]}) -eq $null)
                                    {
                                        if (($package = $EMPpackages | Where-Object{$_.Name -like "*$($SQLreader["softwarename"])*"} | Sort-Object -Property Version -Descending | Select-Object -First 1) -eq $null)
                                        {
                                            Write-Host "- $($SQLreader["softwarename"]):".PadRight($outputWidth) -NoNewline
                                            Write-Host "Failed" -ForegroundColor Red -NoNewline
                                            Write-Host ", please assign manually"
                                        }
                                    }
                                    if ($package -ne $null)
                                    {
                                        Write-Host "- $($package.Vendor) $($package.Name) $($package.Version):".PadRight($outputWidth) -NoNewline
                                        Add-EmpirumPackageToGroup -Package $package -Group $computerNode -Session $EMPsession
                                        Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                                        $state = Get-EmpirumPackageState -Package $package -TargetObject $TargetComputer -Session $EMPsession
                                        Write-Host " (current installation status: " -NoNewline
                                        switch ($state.ExecutionStatus)
                                        {
                                            Finished
                                            {
                                                Write-Host "Finished" -ForegroundColor Green -NoNewline
                                                Write-Host ", $($state.FinishedTime.ToShortDateString())" -NoNewline
                                            }
                                            Running
                                            {
                                                Write-Host "Running" -ForegroundColor Green -NoNewline
                                            }
                                            Unknown
                                            {
                                                Write-Host "Unknown" -ForegroundColor Yellow -NoNewline
                                            }
                                            Cancelled
                                            {
                                                Write-Host "Cancelled" -ForegroundColor Red -NoNewline
                                            }
                                            Failed
                                            {
                                                Write-Host "Failed" -ForegroundColor Red -NoNewline
                                                Write-Host ", $($state.ErrorText))" -NoNewline
                                            }
                                            default
                                            {
                                                Write-Host "Unknown" -ForegroundColor Yellow -NoNewline
                                            }
                                        }
                                        Write-Host ")"
                                        if ($state.ExecutionStatus -ne "Finished")
                                        {
                                            $RolloutStatus = $RolloutStatus -and $false
                                        }
                                    }
                                }
                            }
                            $SQLreader.Close()
                        }
                        #endregion
                        Write-Host "Invoking activation:".PadRight($outputWidth) -NoNewline
                        Invoke-EmpirumComputerActivation -Computer $TargetComputer -Flags Software -Session $EMPsession
                        Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                        if ($RolloutStatus -eq $false)
                        {
                            $RolloutCounter++
                            Write-Host ", rollout counter was increased"
                        }
                        else
                        {
                            Write-Host ", rollout counter wasn't increased"
                        }
                    }
                }
                else
                {
                    Write-Host "Failed" -ForegroundColor Red
                }
            }
            else
            {
                Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                Write-Host ", counter off limits ($RolloutCounter/$RolloutBlockSize, skipping set '$($ComputerName.ToUpper())')"
            }
        }
        #region: Cleanup
        Write-Host ""
        Write-Host "Closing sessions:".PadRight($outputWidth) -NoNewline
        $SQLsession.Close()
        Close-Matrix42ServiceConnection $EMPsession
        Write-Host "Succeeded" -ForegroundColor Green
        #endregion
    }
}
