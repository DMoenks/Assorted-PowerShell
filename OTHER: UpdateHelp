# Author: Moenks, Dominik
# Version: 1.0 (08.01.2014)
# Intention: Updates the PowerShell help when behind a proxy server

$wc = New-Object System.Net.WebClient
$wc.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
Update-Help
