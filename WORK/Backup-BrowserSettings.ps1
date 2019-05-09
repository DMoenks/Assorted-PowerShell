param([ValidateNotNullOrEmpty()]
        [string]$Path,
        [switch]$Restore)

$outputWidth = 50
$browsers = @{'Internet Explorer' = "DIR|$env:USERPROFILE\Favorites";
            'Chrome' = "FILE|$env:USERPROFILE\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"}

Write-Host 'Checking for script mode:'.PadRight($outputWidth) -NoNewline
if ($Restore)
{
    Write-Host 'Restore'
    Write-Host 'Checking for existing backups:'.PadRight($outputWidth) -NoNewline
    if (Test-Path $Path)
    {
        Write-Host 'Succeeded' -ForegroundColor Green
        foreach ($browser in $browsers.Keys)
        {
            Write-Host "Checking for $browser backup:".PadRight($outputWidth) -NoNewline
            if (Test-Path "$Path\$browser")
            {
                Write-Host 'Succeeded' -ForegroundColor Green
                Write-Host "Restoring $browser backup:".PadRight($outputWidth) -NoNewline
                try
                {
                    switch -Wildcard ($browsers[$browser])
                    {
                        'DIR|*'
                        {
                            Remove-Item "$($browsers[$browser].Split('|')[1])\*" -Recurse -Force
                            Copy-Item "$Path\$browser\*" "$($browsers[$browser].Split('|')[1])\" -Recurse -Force | Out-Null
                        }
                        'FILE|*'
                        {
                            Remove-Item "$($browsers[$browser].Split('|')[1])" -Force
                            Copy-Item "$Path\$browser\$(Split-Path $browsers[$browser].Split('|')[1] -Leaf)" "$(Split-Path $browsers[$browser].Split('|')[1] -Parent)\" -Force | Out-Null
                        }
                        default
                        {

                        }
                    }
                    Write-Host 'Succeeded' -ForegroundColor Green
                }
                catch [System.Management.Automation.ItemNotFoundException]
                {
                    Write-Host 'Failed' -ForegroundColor Red
                }
            }
            else
            {
                Write-Host 'Failed' -ForegroundColor Red
            }
        }
        Write-Host 'Removing existing backups:'.PadRight($outputWidth) -NoNewline
        Remove-Item $Path -Recurse -Force
        Write-Host 'Succeeded' -ForegroundColor Green
    }
    else
    {
        Write-Host 'Failed' -ForegroundColor Red
    }
}
else
{
    Write-Host 'Backup'
    Write-Host 'Checking for backup drive:'.PadRight($outputWidth) -NoNewline
    if (Test-Path (Split-Path $Path -Qualifier))
    {
        Write-Host 'Succeeded' -ForegroundColor Green
        Write-Host 'Checking for existing backups:'.PadRight($outputWidth) -NoNewline
        if (Test-Path $Path)
        {
            Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
            Write-Host ', backups already exist'
            Write-Host 'Checking for backup date'.PadRight($outputWidth) -NoNewline
            if ((Get-Item L:\BrowserBackup -Force).CreationTime -lt [System.Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem).InstallDate))
            {
                Write-Host 'Succeeded' -ForegroundColor Green -NoNewline
                Write-Host ', keeping existing backups'
            }
            else
            {
                Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
                Write-Host ', removing existing backups'
                Remove-Item $Path -Recurse -Force
            }
        }
        else
        {
            Write-Host 'Succeeded' -ForegroundColor Green
        }
        (New-Item $Path -ItemType Directory).Attributes += [System.IO.FileAttributes]::Hidden
        foreach ($browser in $browsers.Keys)
        {
            Write-Host "Checking for $browser files to backup:".PadRight($outputWidth) -NoNewline
            if (Test-Path $browsers[$browser].Split('|')[1])
            {
                Write-Host 'Succeeded' -ForegroundColor Green
                Write-Host "Backing up ${browser}:".PadRight($outputWidth) -NoNewline
                New-Item "$Path\$browser" -ItemType Directory | Out-Null
                switch -Wildcard ($browsers[$browser])
                {
                    'DIR|*'
                    {
                        Copy-Item "$($browsers[$browser].Split('|')[1])\*" "$Path\$browser\" -Recurse -Force | Out-Null
                    }
                    'FILE|*'
                    {
                        Copy-Item "$($browsers[$browser].Split('|')[1])" "$Path\$browser\" -Force | Out-Null
                    }
                    default
                    {

                    }
                }
                Write-Host 'Succeeded' -ForegroundColor Green
            }
            else
            {
                Write-Host 'Failed' -ForegroundColor Yellow
            }
        }
    }
    else
    {
        Write-Host 'Failed' -ForegroundColor Red
    }
}
