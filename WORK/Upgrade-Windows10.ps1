<#
.SYNOPSIS
This script is intended to either install, repair or reconfigure the 'Empirum Remote Installation Service'.
.DESCRIPTION
This script either installs, repairs or reconfigures the 'Empirum Remote Installation Service' by checking the following parameters:
- Was the system provisioned recently?
    If so, refrain from applying any changes as the script might be running during initial provisioning of the device.
- Is the service installed?
    If not, install it.
- Is the service working as expected (i.e. are there error-level entries in EventLog)?
    If not, reinstall it.
- Is the service configured correctly (i.e. is the startup type configured to 'Automatic (Delayed Start)')?
    If not, reconfigure it.
    (Please note: If the startup type is configured to 'Disabled' this setting will be retained.)
.PARAMETER DHCPOptionNumber
If the local environment variable EmpirumServer is missing or empty the script will try to get the needed value from a DHCP server.
Provide the DHCP option number used to inform clients about the correct Empirum depot server in your infrastructure.
.PARAMETER FallbackServer
If the depot server configured in the local environment variable EmpirumServer is unavailable the script will use this fallback server instead.
.PARAMETER WeeklyErrorThreshold
If the amount of error-level entries in EventLog from 'Service Control Manager' naming the 'Empirum Remote Installation Service' as source exceeds this value a reinstallation will be invoked.
.PARAMETER EmpAgentBatch
If a (re-)installation of the server needs to invoked the script will run this batch file from '\\%EmpirumServer%\User\'.
.PARAMETER EmpInventoryBatch
If an inventory of the client needs to invoked the script will run this batch file from '\\%EmpirumServer%\User\'.
.PARAMETER EmpAgentLog
The script automatically writes debug information to this location.
Provide a path to the desired log file if the default location is inacceptable.
.EXAMPLE
EmpAgent.ps1 -DHCPOptionNumber 128 -FallbackServer 'EmpirumMaster' -WeeklyErrorThreshold 10 -EmpAgentBatch 'EmpirumAgent.bat' -EmpInventoryBatch 'EmpirumInventory.bat'
.NOTES
Version:    3.1.2
Author:     Mönks, Dominik
#>

param([ValidateNotNullOrEmpty()]
        [string]$TMP_FOLDER = "$PSScriptRoot\Upgrade-Windows10",
        [ValidateNotNullOrEmpty()]
        [string]$LOG_FOLDER = "$TMP_FOLDER\Log",
        [switch]$Force)

$logtime = [datetime]::Now.ToString("yyyyMMddHHmmss")

#region:Functions
function log([string]$message)
{
    if (-not (Test-Path $LOG_FOLDER))
    {
        New-Item -Path $LOG_FOLDER -ItemType Directory | Out-Null
    }
    "$([datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")) | $message" | Out-File  "$LOG_FOLDER\$logname.$logtime.information.log" -Append
}

function docStart()
{
    $xml.WriteStartDocument()
}

function docEnd()
{
    $xml.WriteEndDocument()
}

function eleStart([string]$name, [string]$namespace, [string]$prefix)
{
    if ($namespace -ne "")
    {
        $script:defaultnamespace = $namespace
        if ($prefix -ne "")
        {
            $xml.WriteStartElement($prefix, $name, $namespace)
        }
        else
        {
            $xml.WriteStartElement($name, $namespace)
        }
    }
    elseif ($defaultnamespace -ne $null)
    {
        $xml.WriteStartElement($name, $defaultnamespace)
    }
    else
    {
        $xml.WriteStartElement($name)
    }
}

function eleEnd()
{
    $xml.WriteEndElement()
}

function att([string]$name, [string]$value)
{
    $xml.WriteAttributeString($name, $value)
}
#endregion

# Check Windows architecture
switch ((Get-CimInstance Win32_OperatingSystem).OSArchitecture)
{
    "32-bit"
    {
        $OSArchitecture = "x86"
    }
    "64-bit"
    {
        $OSArchitecture = "x64"
    }
    default
    {
        $OSArchitecture = "x64"
    }
}
# Check Windows base language
$OSLanguage = [cultureinfo]::GetCultureInfo([int](Get-CimInstance Win32_OperatingSystem).OSLanguage).Name
# Select matching installation source
$ISO_WIN = Get-ChildItem -Filter "*$OSArchitecture*$OSLanguage*.iso" -Recurse | Sort-Object CreationTime -Descending | Select-Object -First 1
if ($ISO_WIN -ne $null)
{
    $logname = $ISO_WIN.BaseName
    log "Detected operating system architecture $OSArchitecture and base language $OSLanguage, found matching installation source"
    log "Using '$($ISO_WIN.FullName)' as installation source"
    if (($WIN_DRIVE = Mount-DiskImage $ISO_WIN.FullName -PassThru | Get-Volume) -ne $null)
    {
        log "Mounted installation source as drive $($WIN_DRIVE.DriveLetter):"
    }
    else
    {
        log "Failed to mount installation source"    
    }    
}
else
{
    $logname = "Unknown"
    log "Detected operating system architecture $OSArchitecture and base language $OSLanguage, couldn't find matching installation source"
}
# Select matching language pack source
$ISO_WINLIP = Get-ChildItem -Filter "*LIP*.iso" -Recurse | Sort-Object CreationTime -Descending | Select-Object -First 1
if ($ISO_WINLIP -ne $null)
{
    log "Using '$($ISO_WINLIP.FullName)' as language pack source"
    if (($WINLIP_DRIVE = Mount-DiskImage $ISO_WINLIP.FullName -PassThru | Get-Volume) -ne $null)
    {
        log "Mounted language pack source as drive $($WINLIP_DRIVE.DriveLetter):"
        # Enumerate installed language packs
        $MUILanguages = (Get-CimInstance Win32_OperatingSystem).MUILanguages + "en-US" -notlike $OSLanguage
        log "Found the following additional languages: $($MUILanguages -join ", ")"
        # Copy language packs to local drive
        if (-not (Test-Path "$TMP_FOLDER\LIPs"))
        {
            New-Item "$TMP_FOLDER\LIPs" -ItemType Directory | Out-Null
        }
        $WINLIP_FOLDER = Get-Item "$TMP_FOLDER\LIPs"
        foreach ($language in $MUILanguages)
        {
            if (-not (Test-Path "$WINLIP_FOLDER\$language"))
            {
                New-Item "$WINLIP_FOLDER\$language" -ItemType Directory | Out-Null
            }
            foreach($LIP in Get-ChildItem "$($WINLIP_DRIVE.DriveLetter):\$OSArchitecture\*" -Filter "*.cab" -Recurse | Where-Object{$_.FullName -like "*$language*"})
            {
                Copy-Item $LIP.FullName "$WINLIP_FOLDER\$language\lp.cab" -Force
                log "Successfully copied $language language pack to '$WINLIP_FOLDER\$language\'"
            }
        }
    }
    else
    {
        log "Failed to mount language pack source"
    }
}
else
{
    log "Detected operating system architecture $OSArchitecture, couldn't find matching language pack source"
}
if ($WIN_DRIVE -ne $null -and $WINLIP_DRIVE -ne $null)
{
    # Create post-installation script
    $POCMD_FILE = "$TMP_FOLDER\setupcomplete.cmd"
    $POCMD_CONTENT = [Collections.Generic.List[string]]::new()
    $encoding = [Text.UTF8Encoding]::new($false)
    $xmlcfg = [Xml.XmlWriterSettings]::new()
    $xmlcfg.CloseOutput = $true
    $xmlcfg.Encoding = $encoding
    $xmlcfg.Indent = $true
    # Create globalization settings
    $GSXML_FILE = "$TMP_FOLDER\GlobalizationServices.xml"
    $xml = [Xml.XmlWriter]::Create($GSXML_FILE, $xmlcfg)
    docStart
        eleStart "GlobalizationServices" "urn:longhornGlobalizationUnattend" "gs"
            eleStart "UserList"
                eleStart "User"
                    att "UserID" "Current"
                    att "CopySettingsToDefaultUserAcct" "true"
                eleEnd
            eleEnd
            eleStart "LocationPreferences"
                eleStart "GeoID"
                    att "Value" (Get-WinHomeLocation).GeoId
                eleEnd
            eleEnd
            eleStart "MUILanguagePreferences"
                eleStart "MUILanguage"
                    att "Value" (Get-WinUserLanguageList)[0].LanguageTag
                eleEnd
                eleStart "MUIFallback"
                    att "Value" "en-US"
                eleEnd
            eleEnd
            eleStart "SystemLocale"
                att "Name" (Get-WinSystemLocale).Name
            eleEnd
            eleStart "InputPreferences"
                eleStart "InputLanguageID"
                    att "Action" "add"
                    att "ID" (Get-WinUserLanguageList)[0].InputMethodTips[0]
                    att "Default" "true"
                eleEnd
            eleEnd
            eleStart "UserLocale"
                eleStart "Locale"
                    att "Name" (Get-Culture).Name
                    att "SetAsCurrent" "true"
                    att "ResetAllSettings" "true"
                eleEnd
            eleEnd
        eleEnd
    docEnd
    $xml.Close()
    $POCMD_CONTENT.Add("control intl.cpl,, /f:""$GSXML_FILE""")
    # Reactivate local Administrator account
    $POCMD_CONTENT.Add("wmic USERACCOUNT WHERE (LocalAccount=""TRUE"" AND SID LIKE ""%%500"") SET Disabled=FALSE")
    # Uninstall incompatible software
    $applications = @{
                        "SafeSign"=
                        @{
                            "x86"=@("SafeSign.msi","REBOOT=REALLYSUPPRESS ARPSYSTEMCOMPONENT=1 ALLUSERS=1 ADDLOCAL=""Locales,CSP,CSP_Library,Common_Dialogs,KSP,PKCS11,x32,CSP,KSP,User,TAU,Task_Manager"" /QB-!");
                            "x64"=@("SafeSign 64-bits.msi","REBOOT=REALLYSUPPRESS ARPSYSTEMCOMPONENT=1 ALLUSERS=1 ADDLOCAL=""Locales,CSP,CSP_Library,Common_Dialogs,KSP,PKCS11,x64,x64CSP,x64KSP,x64User,x64TAU,x64Task"" /QB-!")
                        }
                    }
    foreach ($application in $applications.Keys)
    {
        foreach ($entity in @(Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE name LIKE '%$application%'"))
        {
            if ($entity.Uninstall().ReturnValue -eq 0)
            {
                log "Application $application was uninstalled successfully"
                if (Test-Path "$TMP_FOLDER\Software\$application\$OSArchitecture\$($applications[$application][$OSArchitecture][0])")
                {
                    $POCMD_CONTENT.Add("msiexec.exe /I ""$TMP_FOLDER\Software\$application\$OSArchitecture\$($applications[$application][$OSArchitecture][0])"" $($applications[$application][$OSArchitecture][1])")
                }
                else
                {
                    log "Couldn't find installation package for application $application and operating system architecture $OSArchitecture"
                }
            }
            else
            {
                log "Application $application wasn't uninstalled"
            }
        }
    }
    # Write script to be run after setup completes
    [IO.File]::WriteAllLines($POCMD_FILE, $POCMD_CONTENT, $encoding)
    # Prepare system
    Get-ChildItem "$env:SystemRoot\SoftwareDistribution\" | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    Start-Process "$TMP_FOLDER\Software\Tools\psexec.exe" -ArgumentList "-s $env:SystemRoot\System32\rundll32.exe pnpclean.dll,RunDLL_PnpClean /DRIVERS /MAXCLEAN" -Wait
    # Start the upgrade process
    $setupargs = "/Auto Upgrade /Compat IgnoreWarning /DynamicUpdate Disable /InstallLangPacks ""$WINLIP_FOLDER"" /ReflectDrivers ""$TMP_FOLDER\Drivers\$OSArchitecture"" /PostOOBE ""$POCMD_FILE"" /CopyLogs ""$LOG_FOLDER"" /Quiet /NoReboot"
    log "Starting setup using '$setupargs'"
    $setup = Start-Process "$($WIN_DRIVE.DriveLetter):\setup.exe" -ArgumentList $setupargs -RedirectStandardOutput "$LOG_FOLDER\$logname.$logtime.output.log" -RedirectStandardError "$LOG_FOLDER\$logname.$logtime.error.log" -PassThru -Wait
    log "Setup completed"
    # Clean up
    Dismount-DiskImage $ISO_WIN.FullName
    log "Dismounted installation source"
    Dismount-DiskImage $ISO_WINLIP.FullName
    log "Dismounted language pack source"
    Remove-Item $WINLIP_FOLDER -Recurse -Force
    log "Deleted copied language packs"
    # Check for result
    switch ($setup.ExitCode)
    {
        0
        {
            log "Installation was successful"
            log "Finishing with exit code 0"
        }
        default
        {
            log "Installation wasn't successful"
            # Check for incompatible apps
            $incompatibleApps = [Collections.Generic.List[string]]::new()
            foreach ($compatLog in Get-ChildItem $LOG_FOLDER -Filter "CompatData*.xml")
            {
                foreach ($result in [regex]::Matches((Get-Content $compatLog.FullName), '<Program Name="(.+?)".+?BlockingType="Hard".+?<\/Program>'))
                {
                    $incompatibleApps.Add($result.Groups[1].Value)
                }
            }
            if ($incompatibleApps.Count -gt 0)
            {
                log "Found the following incompatible apps: $($incompatibleApps -join ", ")"
            }
            log "Finishing with exit code $($setup.ExitCode)"
        }
    }
}
else
{
    log "Finishing with exit code 1612"
}
