$outputWidth = 40

Write-Host "Checking domain controller:".PadRight($outputWidth) -NoNewline
if ($null -ne (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue))
{
    if ($null -ne ($dnsResult = Resolve-DnsName $env:LOGONSERVER.Trim('\')))
    {
        $domainController = $dnsResult.Name
        Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
        Write-Host ", using server $domainController"
    }
    else
    {
        Write-Host 'Failed' -ForegroundColor Yellow
        Write-Host "Checking alternate domain controller:".PadRight($outputWidth) -NoNewline
        if ($null -ne ($dnsResult = Resolve-DnsName (Get-WmiObject Win32_ComputerSystem).Domain))
        {
            $domainController = $dnsResult.Name
            Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
            Write-Host ", using server $domainController"
        }
        else
        {
            $domainController = ''
            Write-Host 'Failed' -ForegroundColor Red -NoNewline
	        Write-Host ', please contact IT department'
        }
    }
}
else
{
    $domainController = $env:LOGONSERVER.Trim('\')
    Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
	Write-Host ", using server $domainController"
}
if ($domainController -ne '')
{
    Write-Host "Checking network connectivity:".PadRight($outputWidth) -NoNewline
    if (Test-Connection -ComputerName $domainController -Count 5 -Quiet)
    {
        Write-Host 'Succeeded' -ForegroundColor Green
        Write-Host "Checking domain connectivity:".PadRight($outputWidth) -NoNewline
        try
        {
            if (Test-ComputerSecureChannel -Server $domainController -ErrorAction SilentlyContinue)
            {
                Write-Host 'Succeeded' -ForegroundColor Green
            }
            else
            {
                Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
                Write-Host ', which might cause issues'
            }
        }
        catch [System.Management.Automation.MethodInvocationException]
        {
            Write-Host 'Failed' -ForegroundColor Red -NoNewline
            Write-Host ', which probably causes issues'
        }
        Write-Host 'Checking current drive mappings:'.PadRight($outputWidth) -NoNewline
        $networkDrivesPre = (Get-WmiObject -Query 'SELECT DeviceID FROM Win32_LogicalDisk WHERE DriveType = 4').DeviceID
        Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
        switch ($networkDrivesPre.Count)
        {
            0
            {
                Write-Host ", found no drives"
            }
            1
            {
                Write-Host ", found drive $networkDrivesPre"
            }
            default
            {
                Write-Host ", found drives $($networkDrivesPre -join ', ')"
            }
        }
        Write-Host 'Checking drive mapping configuration:'.PadRight($outputWidth) -NoNewline
        try
        {
            $domain = New-Object adsi "LDAP://$env:USERDOMAIN"
            $user = (New-Object adsisearcher $domain, "(&(ObjectClass=user)(SamAccountName=$env:USERNAME))").FindOne()
            $script = $user.Properties.scriptpath
            Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
            if ($null -eq $script -or 'GPO.cmd' -eq $script)
            {
                Write-Host ', probably mapped by group policy'
                Write-Host 'Updating group policies:'.PadRight($outputWidth) -NoNewline
                Start-Process $env:ComSpec -ArgumentList '/C', 'ECHO n | %SystemRoot%\System32\gpupdate.exe /Target:User /Force' -WindowStyle Hidden -Wait
                Write-Host 'Succeeded' -ForegroundColor Green
            }
            else
            {
                Write-Host ', probably mapped by script'
                Write-Host "Running script $($script):".PadRight($outputWidth) -NoNewline
                Start-Process $env:ComSpec -ArgumentList '/C', "\\$domainController\NETLOGON\$script" -WindowStyle Hidden -Wait
                Write-Host 'Succeeded' -ForegroundColor Green
            }
            $updateSucceeded = $true
        }
        catch
        {
            $updateSucceeded = $false
        }
        if ($updateSucceeded)
        {
            Write-Host 'Checking updated drive mappings:'.PadRight($outputWidth) -NoNewline
            $networkDrivesPost = (Get-WmiObject -Query 'SELECT DeviceID FROM Win32_LogicalDisk WHERE DriveType = 4').DeviceID
            Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
            switch ($networkDrivesPost.Count)
            {
                0
                {
                    Write-Host ", found no drives"
                }
                1
                {
                    Write-Host ", found drive $networkDrivesPost"
                }
                default
                {
                    Write-Host ", found drives $($networkDrivesPost -join ', ')"
                }
            }
        }
        else
        {
            Write-Host 'Failed' -ForegroundColor Red -NoNewline
	        Write-Host ', please contact IT department'
        }
    }
    else
    {
	    Write-Host 'Failed' -ForegroundColor Red -NoNewline
	    Write-Host ', please check network connectivity'
    }
    Start-Sleep -Seconds 5
}
