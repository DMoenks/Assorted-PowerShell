param([switch]$FixOutlook,
        [switch]$FixSkype,
        [switch]$FixTeams,
        [switch]$ResetAuthentication,
        [switch]$ResetLicensing)

$outputWidth = 30

function Close-Process([string[]]$processNames)
{
    foreach ($processName in $processNames)
    {
        if ($null -ne ($processes = Get-Process | Where-Object{$_.Name -like "*$processName*" -or $_.Description -like "*$processName*" -or $_.Product -like "*$processName*"}))
        {
            foreach ($process in $processes)
            {
                Write-Host "Closing '$($process.Name)':".PadRight($outputWidth) -NoNewline
                $process | Stop-Process -Force | Wait-Process $process
                Write-Host 'Succeeded' -ForegroundColor Green
            }
        }
    }
}

$response = Read-Host 'Multiple Office applications will be closed, please confirm (Y/N)'
Write-Host ''
if ($response -like 'y*')
{
    # Log current time
    $timestamp = [datetime]::Now
    # Remove common data
    Write-Host 'Removing common data:'.PadRight($outputWidth) -NoNewline
    Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities\*' -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host 'Succeeded' -ForegroundColor Green
    # Remove Outlook data
    if ($FixOutlook)
    {
        Close-Process 'Outlook', 'Skype', 'Teams'
        Write-Host 'Removing Outlook data:'.PadRight($outputWidth) -NoNewline
        Remove-Item 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\*' -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item 'HKCU:\Software\Microsoft\Office\Outlook' -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Outlook' -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Outlook" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Roaming\Microsoft\Outlook" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host 'Succeeded' -ForegroundColor Green
    }
    # Remove Skype data
    if ($FixSkype)
    {
        Close-Process 'Outlook', 'Skype', 'Teams'
        Write-Host 'Removing Skype data:'.PadRight($outputWidth) -NoNewline
        Remove-Item 'HKCU:\Software\Microsoft\Office\Lync' -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Lync' -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Lync" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Roaming\Microsoft\Office\16.0\Lync" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host 'Succeeded' -ForegroundColor Green
    }
    # Remove Teams data
    if ($FixTeams)
    {
        Close-Process 'Outlook', 'Skype', 'Teams'
        Write-Host 'Removing Teams data:'.PadRight($outputWidth) -NoNewline
        Remove-Item 'HKCU:\Software\Microsoft\Office\Teams' -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Roaming\Microsoft\Teams" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host 'Succeeded' -ForegroundColor Green
    }
    if ($ResetAuthentication)
    {
        Close-Process 'Office 2016'
        Write-Host 'Removing authentication data:'.PadRight($outputWidth) -NoNewline
        Remove-Item 'HKCU:\Software\Microsoft\Protected Storage System Provider\*' -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Credentials\*" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Roaming\Microsoft\Credentials\*" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Roaming\Microsoft\Protect\*" -Recurse -Force -ErrorAction SilentlyContinue
        $tmpFile = "$env:TEMP\$($timestamp.ToString('yyyyMMddHHmmss'))-FixO365-Auth.tmp"
        New-Item $tmpFile -ItemType File -Force | Out-Null
        Start-Process 'CMDKEY' -ArgumentList '/list' -RedirectStandardOutput $tmpFile -NoNewWindow -Wait
        $credentials = [regex]::Matches((Get-Content $tmpFile), 'LegacyGeneric:target=MicrosoftOffice\d+_Data:\S*')
        foreach ($credential in $credentials)
        {
            $target = $credential.Value
            Start-Process 'CMDKEY' -ArgumentList "/delete:$target" -NoNewWindow -PassThru -Wait | Out-Null
        }
        $rebootNeeded = $true
        Write-Host 'Succeeded' -ForegroundColor Green
    }
    if ($ResetLicensing)
    {
        Close-Process 'Office 2016'
        Write-Host 'Removing licensing data:'.PadRight($outputWidth) -NoNewline
        Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\*' -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\AppData\Local\Microsoft\Office\16.0\Licensing\*" -Recurse -Force -ErrorAction SilentlyContinue
        if (([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
        {
            $tmpFile = "$env:TEMP\$($timestamp.ToString('yyyyMMddHHmmss'))-FixO365-Lic.tmp"
            New-Item $tmpFile -ItemType File -Force | Out-Null
            Start-Process 'CSCRIPT' -ArgumentList """${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"" /dstatus" -RedirectStandardOutput $tmpFile -NoNewWindow -Wait
            $keys = [regex]::Matches((Get-Content $tmpFile), 'Last 5 characters of installed product key: (\w{5})')
            foreach ($key in $keys)
            {
                $target = $key.Groups[1].Value
                $tmpFile = "$env:TEMP\$($timestamp.ToString('yyyyMMddHHmmss'))-FixO365-Lic$target.tmp"
                New-Item $tmpFile -ItemType File -Force | Out-Null
                Start-Process 'CSCRIPT' -ArgumentList """${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"" /unpkey:$target" -RedirectStandardOutput $tmpFile -NoNewWindow -Wait
            }
            $rebootNeeded = $true
            Write-Host 'Succeeded' -ForegroundColor Green
        }
        else
        {
            Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
            Write-Host ', administrative privileges needed'
        }
    }
    Write-Host ''
    if ($rebootNeeded)
    {
        Write-Host 'Cleanup finished, please reboot device...'
    }
    else
    {
        Write-Host 'Cleanup finished...'
    }
    Read-Host
}
