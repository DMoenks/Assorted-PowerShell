param([string]$SourceComputer,
        [Parameter(Mandatory=$true)]
        [string[]]$TargetComputers,
        [Parameter(Mandatory=$true)]
        [string[]]$FileExtensions)

$regexSID = "^S(-\d+){4,}$"
$regexFileExt = "^\.?[-0-9a-zA-Z]+$"
$outputWidth = 100

function copyStruct ([Microsoft.Win32.RegistryKey]$source,
                    [Microsoft.Win32.RegistryKey]$target,
                    [string]$key)
{
    # Create key structure in target's registry, if not existant
    if ($target.GetSubKeyNames() -notcontains $key)
    {
        $target.CreateSubKey($key, $true) | Out-Null
    }
    # Copy all of this key's values
    foreach ($value in $source.OpenSubKey($key).GetValueNames())
    {
        # Create value while maintaining name, data type and content
        $target.OpenSubKey($key, $true).SetValue($value, $source.OpenSubKey($key).GetValue($value), $source.OpenSubKey($key).GetValueKind($value))
        # When reaching the OpenWithProgids key copy also copy all ProgID structures referenced by any value
        if ($key -eq "OpenWithProgids" -and $value -ne "")
        {
            copyStruct $HKCR_source $HKCR_target $value
        }
    }
    # Recurse for any sub key
    foreach ($subkey in $source.OpenSubKey($key).GetSubKeyNames())
    {
        copyStruct $source.OpenSubKey($key) $target.OpenSubKey($key, $true) $subkey
    }
}

if ($SourceComputer -ne $null)
{
    Write-Host "Trying to reach source system $($SourceComputer.ToUpper()):".PadRight($outputWidth) -NoNewline
    if (Test-Connection $SourceComputer -Count 1 -Quiet)
    {
        Write-Host "Succeeded" -ForegroundColor Green
        Write-Host "Checking source's RemoteRegistry service for status:".PadRight($outputWidth) -NoNewline
        try
        {
            if (($service = Get-Service RemoteRegistry -ComputerName $SourceComputer -ErrorAction Stop).Status -eq [ServiceProcess.ServiceControllerStatus]::Running)
            {
                Write-Host "Running" -ForegroundColor Green
            }
            else
            {
                if ($service.StartType -notin @([ServiceProcess.ServiceStartMode]::Automatic, [ServiceProcess.ServiceStartMode]::Manual))
                {
                    Set-Service RemoteRegistry -ComputerName $SourceComputer -StartupType Manual -ErrorAction Stop
                }
                Set-Service RemoteRegistry -ComputerName $SourceComputer -Status Running -ErrorAction Stop
                Write-Host "Stopped" -ForegroundColor Yellow -NoNewline
                Write-Host ", service was restarted"
            }
            $HKCR_source = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("ClassesRoot", $SourceComputer)
        }
        catch [Microsoft.PowerShell.Commands.ServiceCommandException]
        {
            Write-Host "Failed" -ForegroundColor Red -NoNewline
            Write-Host ", service was inaccessible, so settings won't be copied"
        }
    }
    else
    {
        Write-Host "Failed" -ForegroundColor Red -NoNewline
        Write-Host ", source system wasn't reachable, so settings won't be copied"
    }
}
foreach ($TargetComputer in $TargetComputers)
{
    Write-Host "Trying to reach target system $($TargetComputer.ToUpper()):".PadRight($outputWidth) -NoNewline
    if (Test-Connection $TargetComputer -Count 1 -Quiet)
    {
        Write-Host "Succeeded" -ForegroundColor Green
        Write-Host "Checking target's RemoteRegistry service for status:".PadRight($outputWidth) -NoNewline
         try
        {
            if (($service = Get-Service RemoteRegistry -ComputerName $TargetComputer -ErrorAction Stop).Status -eq [ServiceProcess.ServiceControllerStatus]::Running)
            {
                Write-Host "Running" -ForegroundColor Green
            }
            else
            {
                if ($service.StartType -notin @([ServiceProcess.ServiceStartMode]::Automatic, [ServiceProcess.ServiceStartMode]::Manual))
                {
                    Set-Service RemoteRegistry -ComputerName $TargetComputer -StartupType Manual -ErrorAction Stop
                }
                Set-Service RemoteRegistry -ComputerName $TargetComputer -Status Running -ErrorAction Stop
                Write-Host "Stopped" -ForegroundColor Yellow -NoNewline
                Write-Host ", service was restarted"
            }
            $HKU = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Users", $TargetComputer)
            foreach ($fileext in $FileExtensions)
            {
                if (($match = [regex]::Match($fileext, $regexFileExt)).Success)
                {
                    $tmpFileExt = $match.Value.ToLower()
                    foreach ($sid in $HKU.GetSubKeyNames())
                    {
                        if (($match = [regex]::Match($sid, $regexSID)).Success)
                        {
                            $tmpSID = $match.Value.ToUpper()
                            Write-Host "Deleting custom settings for *.$tmpFileExt (SID:$tmpSID):".PadRight($outputWidth) -NoNewline
                            try
                            {
                                $HKU.DeleteSubKey("$($tmpSID.Value)\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.$tmpFileExt\UserChoice")
                                Write-Host "Succeeded" -ForegroundColor Green -NoNewline
                                Write-Host ", deleted custom settings for extension *.$tmpFileExt"
                            }
                            catch [System.Management.Automation.MethodInvocationException]
                            {
                                Write-Host "Failed" -ForegroundColor Yellow -NoNewline
                                Write-Host ", found no custom settings for extension *.$tmpFileExt"                            
                            }
                        }
                    }
                    if ($HKCR_source -ne $null)
                    {
                        Write-Host "Copying settings for *.$tmpFileExt from $($SourceComputer.ToUpper()) to $($TargetComputer.ToUpper())".PadRight($outputWidth) -NoNewline
                        if ($HKCR_source.GetSubKeyNames() -contains ".$tmpFileExt")
                        {
                            $HKCR_target = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("ClassesRoot", $TargetComputer)
                            copyStruct $HKCR_source $HKCR_target ".$tmpFileExt"
                            Write-Host "Succeeded" -ForegroundColor Green
                            foreach ($MIMEtype in $HKCR_source.OpenSubKey("MIME\Database\Content Type", $false).GetSubKeyNames())
                            {
                                if ($HKCR_source.OpenSubKey("MIME\Database\Content Type\$MIMEtype", $false).GetValue("Extension") -eq ".$tmpFileExt")
                                {
                                    # Copy MIME type
                                    copyStruct $HKCR_source.OpenSubKey("MIME\Database\Content Type", $false) $HKCR_target.OpenSubKey("MIME\Database\Content Type", $true) $MIMEtype
                                    # Copy CLSID tree referenced in MIME type
                                    $CLSID = $HKCR_source.OpenSubKey("MIME\Database\Content Type\$MIMEtype", $false).GetValue("CLSID")
                                    copyStruct $HKCR_source.OpenSubKey("CLSID", $false) $HKCR_target.OpenSubKey("CLSID", $true) $CLSID
                                    # Copy AppID tree referenced in CLSID
                                    $AppID = $HKCR_source.OpenSubKey("CLSID\$CLSID", $false).GetValue("AppID")
                                    copyStruct $HKCR_source.OpenSubKey("AppID", $false) $HKCR_target.OpenSubKey("AppID", $true) $AppID
                                    # Copy CLSID trees for ProgIDs referenced in CLSID
                                    foreach ($ver in $HKCR_source.OpenSubKey("CLSID\$CLSID\ProgID", $false).GetValueNames())
                                    {
                                        $ProgIDVer = $HKCR_source.OpenSubKey("CLSID\$CLSID\ProgID", $false).GetValue($ver)
                                        copyStruct $HKCR_source $HKCR_target $ProgIDVer
                                    }
                                    # Copy CLSID trees for most recent ProgID referenced in CLSID
                                    $ProgID = $HKCR_source.OpenSubKey("CLSID\$CLSID\VersionIndependentProgID", $false).GetValue("")
                                    copyStruct $HKCR_source $HKCR_target $ProgID
                                }
                            }
                        }
                        else
                        {
                            Write-Host "Failed" -ForegroundColor Red -NoNewline
                            Write-Host ", found no settings for extension *.$tmpFileExt on $($SourceComputer.ToUpper())"
                        }
                    }
                }
            }
        }
        catch [Microsoft.PowerShell.Commands.ServiceCommandException]
        {
            Write-Host "Failed" -ForegroundColor Red -NoNewline
            Write-Host ", service was inaccessible, so settings won't be copied"
        }
    }
    else
    {
        Write-Host "Failed" -ForegroundColor Red -NoNewline
        Write-Host ", target system wasn't reachable"
    }
}
