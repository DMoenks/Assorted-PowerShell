# Author: Mönks, Dominik
# Version: 1.1 (07.05.2015)
# Purpose: Checks the local system for the current status of antivirus signature updates, operating system updates and system services that shouldn't be running on a client system
     
# Check antivirus signature updates
Write-Host "Last antivirus signature update: " -NoNewline
Write-Host ([DateTime]::Parse((Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName "AntiVirusProduct" -Filter "NOT displayName LIKE '%Windows%'").timestamp)).ToShortDateString()

# Check operating system updates
Write-Host "Last operating system update: " -NoNewline
Write-Host ((New-Object -ComObject "Microsoft.Update.Session").CreateUpdateSearcher().QueryHistory(0,1) | Select-Object -ExpandProperty Date).ToShortDateString()

# Check state of DHCP server service
Write-Host "State of DHCP server service: " -NoNewline
if ((Get-CimInstance -ClassName "Win32_Service" -Filter "Name LIKE 'DHCPServer'") -eq $null) {Write-Host "Not available"} else {(Get-CimInstance -ClassName "Win32_Service" -Filter "Name LIKE 'DHCPServer'").State}

# Check state of DNS server service
Write-Host "State of DNS server service: " -NoNewline
if ((Get-CimInstance -ClassName "Win32_Service" -Filter "Name LIKE 'DNS'") -eq $null) {Write-Host "Not available"} else {(Get-CimInstance -ClassName "Win32_Service" -Filter "Name LIKE 'DNS'").State}

# Check for possible DHCP services
Write-Host "Found the following DHCP or DNS services:" -NoNewline
Get-CimInstance -ClassName "Win32_Service" -Filter "Name LIKE '%dhcp%' OR Name LIKE '%dns%' OR Caption LIKE '%dhcp%' OR Caption LIKE '%dns%' OR Description LIKE '%dhcp%' OR Description LIKE '%dns%' OR DisplayName LIKE '%dhcp%' OR DisplayName LIKE '%dns%'" | ft -Property @{Expression={$_.Caption};Label="Name"}, @{Expression={$_.State};Label="Status"}, @{Expression={$_.StartMode};Label="Startup Type"}