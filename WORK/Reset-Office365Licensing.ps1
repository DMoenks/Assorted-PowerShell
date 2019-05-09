param([Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Computers)

$outputWidth = 70

foreach ($Computer in $Computers)
{
    Write-Host "Connecting to computer $($Computer.ToUpper()):".PadRight($outputWidth) -NoNewline
    if (Test-Connection $Computer -Quiet)
    {
        Write-Host 'Succeeded' -ForegroundColor Green
        Write-Host 'Searching for product keys:'.PadRight($outputWidth) -NoNewline
        $OSVersion = [version](Get-WmiObject -Query 'SELECT Version FROM Win32_OperatingSystem' -ComputerName $Computer).Version
        if ($OSVersion.Major -gt 6 -or ($OSVersion.Major -eq 6 -and $OSVersion.Minor -gt 1))
        {
            $Products = @(Get-WmiObject -Query 'SELECT ID, PartialProductKey FROM SoftwareLicensingProduct WHERE Name LIKE "%O365ProPlus%" AND PartialProductKey <> NULL' -ComputerName $Computer)
        }
        else
        {
            $Products = @(Get-WmiObject -Query 'SELECT ID, PartialProductKey FROM OfficeSoftwareProtectionProduct WHERE Name LIKE "%O365ProPlus%" AND PartialProductKey <> NULL' -ComputerName $Computer)
        }
        Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
        Write-Host ", found $($Products.Count) product keys"
        foreach ($Product in $Products)
        {
            Write-Host "Uninstalling product key $($Product.PartialProductKey):".PadRight($outputWidth) -NoNewline
            try
            {
                $Product.UninstallProductKey() | Out-Null
                Write-Host 'Succeeded' -ForegroundColor Green
            }
            catch [System.Management.Automation.MethodInvocationException]
            {
                Write-Host 'Failed' -ForegroundColor Red
            }
        }
        Write-Host 'Accessing registry:'.PadRight($outputWidth) -NoNewline
        try
        {
            $RHKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $Computer, [Microsoft.Win32.RegistryView]::Default)
            $RHKU = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::Users, $Computer, [Microsoft.Win32.RegistryView]::Default)
            $SIDs = $RHKU.GetSubKeyNames() -match '^S-1-5-21(-\d{10}){3}-\d{4,}$'
            Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
            Write-Host ", found $($SIDs.Count) SIDs"
        }
        catch [System.Management.Automation.MethodInvocationException]
        {
            Write-Host 'Failed' -ForegroundColor Red
        }
        if ($null -ne $RHKLM.Handle -and $null -ne $RHKU.Handle)
        {
            Write-Host 'Clearing computer:'.PadRight($outputWidth) -NoNewline
            if ($null -ne $RHKLM.OpenSubKey('Software\Microsoft\Office'))
            {
                $RHKLM.OpenSubKey('Software\Microsoft\Office\16.0', $true).CreateSubKey('User Settings\ResetIdentity\Delete\Software\Microsoft\Office\16.0\Common\Identity') | Out-Null
                $RHKLM.OpenSubKey('Software\Microsoft\Office\16.0', $true).CreateSubKey('User Settings\ResetUserLicense\Delete\Software\Microsoft\Office\16.0\Common\Licensing') | Out-Null
                $RHKLM.OpenSubKey('Software\Microsoft\Office\16.0', $true).CreateSubKey('User Settings\ResetUserRegistration\Delete\Software\Microsoft\Office\16.0\Registration') | Out-Null
                $RHKLM.OpenSubKey('Software\Microsoft\Office\ClickToRun\Configuration', $true).DeleteValue('O365ProPlusRetail.EmailAddress', $false)
                $RHKLM.OpenSubKey('Software\Microsoft\Office\ClickToRun\Configuration', $true).DeleteValue('O365ProPlusRetail.TenantId', $false)
                $RHKLM.OpenSubKey('Software\Microsoft\Office\ClickToRun\Configuration', $true).DeleteValue('ProductKeys', $false)
            }
            Write-Host 'Succeeded' -ForegroundColor Green
            foreach ($SID in $SIDs)
            {
                Write-Host "Clearing user ${SID}:".PadRight($outputWidth) -NoNewline
                if ($null -ne $RHKU.OpenSubKey("$SID\Software\Microsoft\Office"))
                {
                    $RHKU.OpenSubKey("$SID\Software\Microsoft\Office\16.0\Common", $true).DeleteSubKeyTree('Identity', $false)
                    $RHKU.OpenSubKey("$SID\Software\Microsoft\Office\16.0\Common", $true).DeleteSubKeyTree('Licensing', $false)
                    $RHKU.OpenSubKey("$SID\Software\Microsoft\Office\16.0", $true).DeleteSubKeyTree('Registration', $false)
                    $Profile = $RHKLM.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$SID", $false).GetValue('ProfileImagePath').Replace(':', '$')
                    Remove-Item "\\$Computer\$Profile\AppData\Local\Microsoft\Office\16.0\Licensing\*" -Recurse -Force -ErrorAction SilentlyContinue
                }
                Write-Host 'Succeeded' -ForegroundColor Green
            }
        }
        else
        {

        }
    }
    else
    {
        Write-Host 'Failed' -ForegroundColor Red
    }
}
