# Author: Moenks, Dominik
# Version: 1.2 (11.08.2016)
# Intention: Requests DHCP option values from the server
# Based on: http://www.indented.co.uk/2010/02/17/dhcp-discovery/

param([Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [int[]]$DHCPOptions)

function Create-DHCPINFORM([int[]]$options)
{
    $netconf = (Get-WmiObject -Query 'SELECT IPAddress, MACAddress FROM Win32_NetworkAdapterConfiguration WHERE DHCPServer IS NOT NULL')
    
    # Create the Byte Array
    $packet = New-Object Byte[] (243 + 3 * $options.Count)
    
    # op: BOOTREQUEST
    $packet[0] = 1
    # htype: 10Mbit Ethernet
    $packet[1] = 1
    # hlen: 10Mbit Ethernet
    $packet[2] = 6
    # xid: Random transaction ID
    $xid = New-Object Byte[] 4
    [Random]::new().NextBytes($xid)
    [Array]::Copy($xid, 0, $packet, 4, 4)
    # flags: Broadcast
    $packet[10] = 128
    # ciaddr: IP address
    $packet[12] = $netconf.IPAddress.Split(".")[0]
    $packet[13] = $netconf.IPAddress.Split(".")[1]
    $packet[14] = $netconf.IPAddress.Split(".")[2]
    $packet[15] = $netconf.IPAddress.Split(".")[3]
    # chaddr: MAC address
    $MACstring = $netconf.MACAddress.Replace(":", "")
    $MACbytes = [BitConverter]::GetBytes(([UInt64]::Parse($MACstring, [Globalization.NumberStyles]::HexNumber)))
    [Array]::Reverse($MACbytes)
    [Array]::Copy($MACbytes, 2, $packet, 28, 6)
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
    for ($i = 0; $i -lt $options.Count; $i++)
    {
        $option = $i * 3
        $packet[242 + $option + 1] = 55
        $packet[242 + $option + 2] = 1
        $packet[242 + $option + 3] = $options[$i]
    }
    
    return $packet
}

function Read-DHCP( [byte[]]$Packet )
{
    $Reader = New-Object IO.BinaryReader(New-Object IO.MemoryStream(@(,$Packet)))

    $DhcpResponse = New-Object Object

    $Reader.ReadBytes(4) | Out-Null
    $DhcpResponse | Add-Member NoteProperty XID $Reader.ReadUInt32()
    $Reader.ReadBytes(232) | Out-Null

    # Start reading Options
    $DhcpResponse | Add-Member NoteProperty Options @()
    While ($Reader.BaseStream.Position -lt $Reader.BaseStream.Length)
    {
        $optioncode = $Reader.ReadByte()
        If ($optioncode -ne 0 -and $optioncode -ne 255)
        {
            $Option = New-Object Object
            $Option | Add-Member NoteProperty OptionCode $optioncode
            $Option | Add-Member NoteProperty Length 0
            $Option | Add-Member NoteProperty OptionValue ""
            $Option.Length = $Reader.ReadByte()
            $Buffer = New-Object Byte[] $Option.Length
            
            [Void]$Reader.Read($Buffer, 0, $Option.Length)
            $Option.OptionValue = $Buffer
            if ($optioncode -in $DHCPOptions)
            {
                # Override the ToString method
                $Option | Add-Member ScriptMethod ToString `
                { Return "$($this.OptionName) ($($this.OptionValue))" } -Force

                $DhcpResponse.Options += $Option
            }
        }
    }
    Return $DhcpResponse
}

$DHCPINFORMpacket = Create-DHCPINFORM $DHCPOptions

$socket = [Net.Sockets.Socket]::new([Net.Sockets.AddressFamily]::InterNetwork, [Net.Sockets.SocketType]::Dgram, [Net.Sockets.ProtocolType]::Udp)
$socket.EnableBroadcast = $true
$socket.ExclusiveAddressUse = $false
$socket.SendTimeOut = [timespan]::FromSeconds(30).TotalMilliseconds
$socket.ReceiveTimeOut = [timespan]::FromSeconds(30).TotalMilliseconds

$socket.Bind([Net.IPEndPoint]::new([ipaddress]::Any, 68))
$socket.SendTo($DHCPINFORMpacket, [Net.IPEndPoint]::new([ipaddress]::Broadcast, 67)) | Out-Null

$timeout = $false
$begin = Get-Date
while (-not $cancel)
{
    $BytesReceived = 0
    Try
    {
        $endpoint = [Net.EndPoint]([Net.IPEndPoint]::new([ipaddress]::Any, 0))
        $buffer = New-Object Byte[] 1024
        $BytesReceived = $socket.ReceiveFrom($buffer, [Ref]$endpoint)
    }
    catch [Net.Sockets.SocketException]
    {
        $cancel = $true
    }
    if ($BytesReceived -gt 0)
    {
        $result = Read-DHCP $buffer[0..$BytesReceived]
        if ($result.XID -eq [BitConverter]::ToUInt32($DHCPINFORMpacket[4..7],0))
        {
            $cancel = $true
        }
    }
    if ((Get-Date) -gt $begin.AddSeconds(60))
    {
        $cancel = $true
    }
}

foreach ($option in $result.Options)
{
    Write-Host "$($option.OptionCode):".PadRight(10) -NoNewline
    Write-Host $([string]::Join('', ($option.OptionValue | %{[char]$_})).ToUpper())
}

$socket.Shutdown("Both")
$socket.Close()
