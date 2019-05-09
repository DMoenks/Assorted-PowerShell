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
.PARAMETER EmpAgentBatch
If a (re-)installation of the server needs to invoked the script will run this batch file from '\\%EmpirumServer%\User\'.
.PARAMETER EmpInventoryBatch
If an inventory of the client needs to invoked the script will run this batch file from '\\%EmpirumServer%\User\'.
.PARAMETER ErrorThreshold
If the amount of Error-level entries in EventLog from 'Service Control Manager' naming the 'Empirum Remote Installation Service' as source exceeds this value a reinstallation will be invoked.
.PARAMETER IntervalDays
This interval is used for every check made by this script, e.g. the system's provisioning date, the service's error count, etc.
.PARAMETER EmpAgentLog
The script automatically writes debug information to this location.
Provide a path to the desired log file if the default location is inacceptable.
.PARAMETER WhatIf
If the script would end up executing either of the files set by EmpInventoryBatch or EmpAgentBatch this switch just logs this result without execution.
.EXAMPLE
EmpAgent.ps1 -DHCPOptionNumber 128 -FallbackServer 'EmpirumMaster' -EmpAgentBatch 'EmpirumAgent.bat' -EmpInventoryBatch 'EmpirumInventory.bat' -ErrorThreshold 10
.NOTES
Version:    2.3.2
Author:     Mönks, Dominik

Parts of this script are based on Chris Dent's article 'DHCP Discovery' available at http://www.indented.co.uk/2010/02/17/dhcp-discovery/.
The license statement for the script included in the aforementioned article is as follows as of 2017-04-18:

(c) 2008-2014 Chris Dent.
Permission to use, copy, modify, and/or distribute this software for any purpose with or without fee is hereby granted, provided that the above copyright notice and this permission notice appear in all copies.
THE SOFTWARE IS PROVIDED “AS IS” AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#>

param([Parameter(Mandatory=$true)]
        [ValidateRange(128,254)]
        [int]$DHCPOptionNumber,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$FallbackServer,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$EmpAgentBatch,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$EmpInventoryBatch,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [int]$ErrorThreshold,
        [int]$IntervalDays=7,
        [string]$EmpAgentLog='C:\EmpAgent.log',
        [switch]$Force,
        [switch]$WhatIf)

function New-DHCPINFORM([int[]]$optioncodes)
{
    $netconf = (Get-WmiObject -Query 'SELECT IPAddress, MACAddress FROM Win32_NetworkAdapterConfiguration WHERE DHCPServer IS NOT NULL')
    
    # Create the Byte Array
    $packet = New-Object Byte[] (243 + 3 * $optioncodes.Count)
    # op: BOOTREQUEST
    $packet[0] = 1
    # htype: 10Mbit Ethernet
    $packet[1] = 1
    # hlen: 10Mbit Ethernet
    $packet[2] = 6
    # xid: Random transaction ID
    $xid = New-Object Byte[] 4
    (New-Object Random).NextBytes($xid)
    [array]::Copy($xid, 0, $packet, 4, 4)
    # flags: Broadcast
    $packet[10] = 128
    # ciaddr: IP address
    $ipaddress = @($netconf.IPAddress | ForEach-Object{[ipaddress]::Parse($_)} | Where-Object{$_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork})[0]
    $packet[12] = $ipaddress.GetAddressBytes()[0]
    $packet[13] = $ipaddress.GetAddressBytes()[1]
    $packet[14] = $ipaddress.GetAddressBytes()[2]
    $packet[15] = $ipaddress.GetAddressBytes()[3]
    # chaddr: MAC address
    $MACstring = $netconf.MACAddress.Replace(':', '')
    $MACbytes = [System.BitConverter]::GetBytes(([uint64]::Parse($MACstring, [System.Globalization.NumberStyles]::HexNumber)))
    [array]::Reverse($MACbytes)
    [array]::Copy($MACbytes, 2, $packet, 28, 6)
    # options: Magic cookie
    $packet[236] = 99 
    $packet[237] = 130
    $packet[238] = 83
    $packet[239] = 99
    # options/DHCP Message Type: DHCPINFORM
    $packet[240] = 53
    $packet[241] = 1
    $packet[242] = 8
    # options/Parameter Request List: DHCPINFORM
    for ($i = 0; $i -lt $optioncodes.Count; $i++)
    {
        $option = $i * 3
        $packet[242 + $option + 1] = 55
        $packet[242 + $option + 2] = 1
        $packet[242 + $option + 3] = $optioncodes[$i]
    }
    return $packet
}

function Read-DHCP([byte[]]$packet, [int[]]$optioncodes)
{
    $DHCPresponse = New-Object Object
    $reader = New-Object IO.BinaryReader(New-Object IO.MemoryStream(@(,$packet)))
    # Skip package header
    $reader.ReadBytes(4) | Out-Null
    # Read transaction ID for comparison
    $DHCPresponse | Add-Member NoteProperty 'XID' $reader.ReadUInt32()
    # Skip package metadata
    $reader.ReadBytes(232) | Out-Null
    # Read options
    $DHCPresponse | Add-Member NoteProperty 'Options' @()
    while ($reader.BaseStream.Position -lt $reader.BaseStream.Length)
    {
        $optioncode = $reader.ReadByte()
        if ($optioncode -ne 0 -and $optioncode -ne 255)
        {
            $option = New-Object Object
            $option | Add-Member NoteProperty 'OptionCode' $optioncode
            $option | Add-Member NoteProperty 'Length' ($reader.ReadByte())
            $option | Add-Member NoteProperty 'OptionValue' (New-Object Byte[] $option.Length)
            [void]$reader.Read($option.OptionValue, 0, $option.Length)
            # Was the option requested?
            if ($optioncodes -contains $optioncode)
            {
                $DHCPresponse.Options += $option
            }
        }
    }
    return $DHCPresponse
}

function Write-Log([string]$Message, [System.ConsoleColor]$Color = [System.ConsoleColor]::White, [switch]$NoNewLine)
{
    if ($NoNewLine)
    {
        Write-Host $Message -ForegroundColor $Color -NoNewline
        [System.IO.File]::AppendAllText($EmpAgentLog, $Message)
    }
    else
    {
        Write-Host $Message -ForegroundColor $Color
        [System.IO.File]::AppendAllText($EmpAgentLog, $Message + [System.Environment]::NewLine)
    }
}

$outputWidth = 100

# Check if the script was run in an administrative context.
if (([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
{
    # Collect previous log files.
    $logFiles = Get-ChildItem (Split-Path $EmpAgentLog -Parent) -Filter (((Split-Path $EmpAgentLog -Leaf).Split('.') | Select-Object -First 1 -Last 1) -join '*.')
    # Remove all but the last 5 log files.
    $logFiles | Sort-Object -Property 'LastWriteTime' -Descending | Select-Object -Skip 5 | Remove-Item -Force
    # Create actual log file name from given file name.
    $EmpAgentLog = [regex]::Replace($EmpAgentLog, '\w+$', "$([datetime]::Now.ToString('yyyyMMddHHmmss')).$&")
    # Was the operating system installed recently?
    Write-Log -Message 'Checking if system was provisioned recently:'.PadRight($outputWidth) -NoNewLine
    if (($installDate = [System.Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject 'Win32_OperatingSystem').InstallDate)) -lt [datetime]::Now.AddDays(-$IntervalDays))
    {
        Write-Log -Message 'Succeeded' -Color Green
        # Is the service installed?
        Write-Log -Message 'Checking for ERIS service:'.PadRight($outputWidth) -NoNewLine
        if ((Get-Service ERIS -ErrorAction SilentlyContinue) -ne $null)
        {
            Write-Log -Message 'Succeeded' -Color Green
            # Is installation enforced?
            Write-Log -Message 'Checking if installation is enforced:'.PadRight($outputWidth) -NoNewLine
            if ($Force)
            {
                Write-Log -Message 'Succeeded' -Color Green
                $install = $true
            }
            else
            {
                Write-Log -Message 'Failed' -Color Yellow
                # How often and when was the script run recently?
                if ($logFiles.Count -gt 1)
                {
                    # Calculate intervals
                    $intervals = New-Object 'System.Collections.Generic.List[long]'
                    for ($i = 1; $i -lt $logFiles.Count; $i++)
                    {
                        $intervals.Add(($logFiles[$i].LastWriteTime - $logFiles[$i - 1].LastWriteTime).Ticks)
                    }
                    $intervals = $intervals | Sort-Object
                    # Calculate the median interval
                    if ($intervals.Count % 2 -eq 0)
                    {
                        $median = ($intervals[$intervals.Count / 2 - 1] + $intervals[$intervals.Count / 2]) / 2
                    }
                    else
                    {
                        $median = $intervals[($intervals.Count - 1) / 2]
                    }
                    # Calculate the average interval
                    $average = 0
                    $intervals | ForEach-Object{$average += $_}
                    $average = $average / $intervals.Count
                }
                # Was the script run recently?
                Write-Log -Message 'Checking if script was run recently:'.PadRight($outputWidth) -NoNewline
                if (($median -le $average) -or (($logFiles | Select-Object -First 1).LastWriteTime -le [datetime]::Now.AddDays(-$IntervalDays)))
                {
                    Write-Log -Message 'Succeeded' -Color Green
                    Write-Log -Message 'Checking service health:'.PadRight($outputWidth) -NoNewline
                    # Is the service disabled?
                    if ((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\ERIS').Start -eq 4)
                    {
                        Write-Log -Message 'Failed' -Color Yellow -NoNewline
                        Write-Log ', service is disabled'
                        Write-EventLog -LogName Application -Source 'Application' -EntryType Warning -EventId 10002 -Category 4 -Message 'ERIS service has warnings, service is disabled'
                        Remove-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\ERIS' -Name DelayedAutostart
                        $install = $false
                    }
                    else
                    {
                        # Make sure the service is configured for delayed automatic start
                        Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\ERIS' -Name Start -Value 2 -Type DWORD
                        Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\ERIS' -Name DelayedAutostart -Value 1 -Type DWORD
                        # Does the service's error count exceed the threshold?
                        if (($ErrorCount = @(Get-EventLog -LogName System -EntryType Error -Source 'Service Control Manager' -Message '*Empirum Remote Installation Service*' -After ([datetime]::Now.AddDays(-$IntervalDays)) -ErrorAction SilentlyContinue).Count) -gt $ErrorThreshold)
                        {
                            Write-Log -Message 'Failed' -Color Red -NoNewline
                            Write-Log ", initiating reinstallation (EventLog errors: $ErrorCount/$ErrorThreshold)"
                            Write-EventLog -LogName Application -Source 'Application' -EntryType Error -EventId 10003 -Category 4 -Message 'ERIS service has errors, initiating reinstallation...'
                            $install = $true
                        }
                        # Does the agent config contain settings?
                        elseif ((Get-Item "$env:SystemRoot\System32\Empirum\AgentConfig.xml" -ErrorAction SilentlyContinue).Length -lt 1024)
                        {
                            Write-Log -Message 'Failed' -Color Red -NoNewline
                            Write-Log ', initiating reinstallation (AgentConfig.xml empty or missing)'
                            Write-EventLog -LogName Application -Source 'Application' -EntryType Error -EventId 10003 -Category 4 -Message 'ERIS service has errors, initiating reinstallation...'
                            $install = $true
                        }
                        # Is the service's log up-to-date?
                        elseif ((Get-Item "$env:SystemRoot\System32\Empirum\ERIS.log" -ErrorAction SilentlyContinue).LastWriteTime -lt [System.Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject 'Win32_OperatingSystem').LastBootUpTime).AddDays(-$IntervalDays))
                        {
                            Write-Log -Message 'Failed' -Color Red -NoNewline
                            Write-Log ', initiating reinstallation (ERIS.log outdated)'
                            Write-EventLog -LogName Application -Source 'Application' -EntryType Error -EventId 10003 -Category 4 -Message 'ERIS service has errors, initiating reinstallation...'
                            $install = $true
                        }
                        else
                        {
                            Write-Log -Message 'Succeeded' -Color Green -NoNewLine
                            Write-Log -Message " (EventLog errors: $ErrorCount/$ErrorThreshold)"
                            Write-EventLog -LogName Application -Source 'Application' -EntryType Information -EventId 10001 -Category 4 -Message 'ERIS service is configured correctly'
                            $install = $false
                        }
                    }
                }
                else
                {
                    Write-Log -Message 'Failed' -Color Yellow -NoNewline
                    Write-Log ", script was last run on $(($logFiles | Sort-Object -Property 'LastWriteTime' -Descending | Select-Object -First 1).LastWriteTime.ToString('yyyy%-MM%-dd'))"
                    $install = $false
                }
            }
        }
        else
        {
            Write-Log -Message 'Failed' -Color Yellow -NoNewline
            Write-Log -Message ', initiating installation'
            Write-EventLog -LogName Application -Source 'Application' -EntryType Error -EventId 10003 -Category 4 -Message 'ERIS service is missing, initiating installation...'
            $install = $true
        }
        # Does the service need to be (re-)installed?
        if ($install)
        {    
            # Does the environment variable EmpirumServer contain a value and, if so, is it different from the fallback server?
            Write-Log -Message 'Checking for EmpirumServer variable:'.PadRight($outputWidth) -NoNewline
            if ($env:EmpirumServer -ne $null -and $env:EmpirumServer -ne '' -and $env:EmpirumServer -notlike "$FallbackServer*")
            {
                Write-Log -Message 'Succeeded' -Color Green
                $EmpirumServer = $env:EmpirumServer
            }
            else
            {
                Write-Log -Message 'Failed' -Color Yellow -NoNewline
                Write-Log -Message ', reading value from DHCP options'
                # Read DHCP settings from registry and process them into readable format.
                $settings = (Get-WmiObject -Query 'SELECT SettingID FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = "True" AND DHCPEnabled = "True"').SettingID
                foreach ($setting in $settings)
                {
                    $DHCPInterfaceOptions = (Get-ItemProperty -Path ("HKLM:\SYSTEM\CurrentControlSet\services\Tcpip\Parameters\Interfaces\$setting")).DhcpInterfaceOptions
                    $DHCPOptions = @{}
                    $position = 0
                    while ($position -lt $DHCPInterfaceOptions.Length) 
                    {
                        $DHCPOptionCode = $DHCPInterfaceOptions[$position]
                        $position = $position + 8

                        $DHCPOptionLength = $DHCPInterfaceOptions[$position]
                        if (($DHCPOptionLength % 4) -eq 0)
                        {
                            $DHCPOptionBytes = ($DHCPOptionLength - ($DHCPOptionLength % 4))
                        }
                        else
                        {
                            $DHCPOptionBytes = ($DHCPOptionLength - ($DHCPOptionLength % 4) + 4)
                        }
                        $position = $position + 12

                        $DHCPOptionValue = New-Object Byte[] $DHCPOptionBytes
                        for ($i = 0; $i -lt $DHCPOptionLength; $i++)
                        {
                            $DHCPOptionValue[$i] = $DHCPInterfaceOptions[$position + $i]
                        }
                        $position = $position + $DHCPOptionBytes

                        if (-not $DHCPOptions.ContainsKey($DHCPOptionCode))
                        {
                            $DHCPOptions.Add($DHCPOptionCode, $DHCPOptionValue)
                        }
                    }
                }
                # Do the DHCP settings from registry contain the expected option number?
                Write-Log -Message "Checking if DHCP options contain option ${DHCPOptionNumber}:".PadRight($outputWidth) -NoNewline
                if ($DHCPOptions.ContainsKey([byte]$DHCPOptionNumber))
                {
                    Write-Log -Message 'Succeeded' -Color Green
                    $EmpirumServer = [string]::Join('', ($DHCPOptions[[byte]$DHCPOptionNumber] | ForEach-Object{[char]$_})).ToUpper()
                }
                else
                {
                    Write-Log -Message 'Failed' -Color Yellow -NoNewline
                    Write-Log -Message ', requesting value from DHCP server'
                    # Broadcast a DHCP request to check for the expected option number.
                    $DHCPINFORMpacket = New-DHCPINFORM $DHCPOptionNumber
                    $socket = New-Object Net.Sockets.Socket([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.SocketType]::Dgram, [System.Net.Sockets.ProtocolType]::Udp)
                    $socket.EnableBroadcast = $true
                    $socket.ExclusiveAddressUse = $false
                    $socket.SendTimeOut = [timespan]::FromSeconds(30).TotalMilliseconds
                    $socket.ReceiveTimeOut = [timespan]::FromSeconds(30).TotalMilliseconds
                    $socket.Bind((New-Object Net.IPEndPoint([ipaddress]::Any, 68)))
                    $socket.SendTo($DHCPINFORMpacket, (New-Object Net.IPEndPoint([ipaddress]::Broadcast, 67))) | Out-Null
                    # Wait for a response...
                    $timeout = $false
                    $begin = Get-Date
                    while (-not $timeout)
                    {
                        $BytesReceived = 0
                        try
                        {
                            $endpoint = [Net.EndPoint](New-Object Net.IPEndPoint([ipaddress]::Any, 0))
                            $buffer = New-Object Byte[] 1024
                            $BytesReceived = $socket.ReceiveFrom($buffer, [Ref]$endpoint)
                        }
                        catch [Net.Sockets.SocketException]
                        {
                            $timeout = $true
                        }
                        # When receiving a response process it into readable format.
                        if ($BytesReceived -gt 0)
                        {
                            $result = Read-DHCP $buffer[0..$BytesReceived] $DHCPOptionNumber
                            if ($result.XID -eq [System.BitConverter]::ToUInt32($DHCPINFORMpacket[4..7],0))
                            {
                                $timeout = $true
                            }
                        }
                        if ((Get-Date) -gt $begin.AddSeconds(60))
                        {
                            $timeout = $true
                        }
                    }
                    $socket.Shutdown('Both')
                    $socket.Close()
                    # Does the response contain the expected option number?
                    Write-Log -Message "Checking if DHCP response contains option ${DHCPOptionNumber}:".PadRight($outputWidth) -NoNewline
                    if (($result.Options | Where-Object{$_.OptionCode -eq $DHCPOptionNumber}) -ne $null)
                    {
                        Write-Log -Message 'Succeeded' -Color Green
                        # Yes, so use it.
                        $EmpirumServer = [string]::Join('', (($result.Options | Where-Object{$_.OptionCode -eq $DHCPOptionNumber}).OptionValue | ForEach-Object{[char]$_})).ToUpper()
                    }
                    else
                    {
                        Write-Log -Message 'Failed' -Color Red -NoNewline
                        Write-Log -Message ', using fallback server'
                        # No, so use the fallback server.
                        $EmpirumServer = $FallbackServer
                    }
                }
            }
            Write-Log -Message 'Trying to resolve server name:'.PadRight($outputWidth) -NoNewline
            try
            {
                # Resolve the server's name into FQDN and configure the environment variable accordingly.
                $EmpirumServer = ([System.Net.Dns]::GetHostEntry($EmpirumServer)).HostName
                Write-Log -Message 'Succeeded' -Color Green
                Write-Log -Message "Setting variable EmpirumServer to value ${EmpirumServer}:".PadRight($outputWidth) -NoNewline
                [System.Environment]::SetEnvironmentVariable('EmpirumServer', $EmpirumServer, [System.EnvironmentVariableTarget]::Machine)
                $env:EmpirumServer = $EmpirumServer
                Write-Log -Message 'Succeeded' -Color Green
                # Is the server reachable?
                Write-Log -Message 'Checking if server is reachable:'.PadRight($outputWidth) -NoNewline
                if (Test-Connection $env:EmpirumServer -Quiet)
                {
                    Write-Log -Message 'Succeeded' -Color Green
                    # Does the server hold a configuration file for the client?
                    Write-Log -Message "Checking for $env:COMPUTERNAME.ini on server:".PadRight($outputWidth) -NoNewline
                    if (@(Get-ChildItem "\\$env:EmpirumServer\Configurator$\Values\MachineValues\" -Filter "$env:COMPUTERNAME.ini" -Recurse).Count -gt 0)
                    {
                        Write-Log -Message 'Succeeded' -Color Green
                        # Yes, so invoke a (re-)installation.
                        Write-Log -Message 'Trying to install EmpAgent:'.PadRight($outputWidth) -NoNewline
                        if (Test-Path "\\$env:EmpirumServer\Configurator$\User\$EmpAgentBatch")
                        {
                            if ($WhatIf)
                            {
                                Write-Log -Message 'Succeeded' -Color Yellow -NoNewLine
                                Write-Log -Message ', at least would have...'
                            }
                            else
                            {
                                $result = Start-Process -FilePath "\\$env:EmpirumServer\Configurator$\User\$EmpAgentBatch" -Verb 'runas' -Wait -PassThru
                                Write-Log -Message 'Succeeded' -Color Green -NoNewLine
                                Write-Log -Message ", return code was $($result.ExitCode)"
                            }
                        }
                        else
                        {
                            Write-Log -Message 'Failed' -Color Red -NoNewline
                            Write-Log -Message ", couldn't find file $EmpAgentBatch"
                        }
                    }
                    else
                    {
                        Write-Log -Message 'Failed' -Color Yellow
                        # No, so invoke an inventory.
                        Write-Log -Message 'Trying to run EmpInventory:'.PadRight($outputWidth) -NoNewline
                        if (Test-Path "\\$env:EmpirumServer\Configurator$\User\$EmpInventoryBatch")
                        {
                            if ($WhatIf)
                            {
                                Write-Log -Message 'Succeeded' -Color Yellow -NoNewLine
                                Write-Log -Message ', at least would have...'
                            }
                            else
                            {
                                $result = Start-Process -FilePath "\\$env:EmpirumServer\Configurator$\User\$EmpInventoryBatch" -Verb 'runas' -Wait -PassThru
                                Write-Log -Message 'Succeeded' -Color Green -NoNewLine
                                Write-Log -Message ", return code was $($result.ExitCode)"
                            }
                        }
                        else
                        {
                            Write-Log -Message 'Failed' -Color Red -NoNewLine
                            Write-Log -Message ", couldn't find file $EmpInventoryBatch"
                        }
                    }
                }
                else
                {
                    Write-Log -Message 'Failed' -Color Red
                }
            }
            catch [Net.Sockets.SocketException]
            {
                Write-Log -Message 'Failed' -Color Red
            }
        }
    }
    else
    {
        Write-Log -Message 'Failed' -Color Red -NoNewLine
        Write-Log -Message ", system was provisioned on $($installDate.ToString('yyyy%-MM%-dd'))"
    }
}
else
{
    Write-Host 'The script was run without administrative rights. Please try again.'
}
