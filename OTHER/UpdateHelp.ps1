# Author: Moenks, Dominik
# Version: 4.0 (08.01.2014)
# Intention: Install or updates modules and updates help files

param([ValidateNotNullOrEmpty()]
        [string[]]$Modules,
        [ValidateNotNullOrEmpty()]
        [switch]$Force,
        [switch]$Verbose)

$outputwidth = 60

Write-Host 'Checking for administrative privileges:'.PadRight($outputwidth) -NoNewline
if (([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
{
    Write-Host 'Succeeded' -ForegroundColor Green
    Write-Host 'Authenticating with proxy:'.PadRight($outputwidth) -NoNewline
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    Write-Host 'Succeeded' -ForegroundColor Green
    Write-Host 'Checking for package provider:'.PadRight($outputwidth) -NoNewline
    if ($null -eq (Get-PackageProvider -Name 'NuGet' -ErrorAction SilentlyContinue | Where-Object{$_.Version -ge [version]::Parse('2.8.5.201')}))
    {
        Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
        Write-Host ', installing needed provider'
        Install-PackageProvider -Name 'NuGet' -MinimumVersion 2.8.5.201 -Verbose:$Verbose.IsPresent
    }
    else
    {
        Write-Host 'Succeeded' -ForegroundColor Green
    }
    Write-Host 'Checking for package repository:'.PadRight($outputwidth) -NoNewline
    if ($null -eq (Get-PSRepository 'PSGallery' -ErrorAction SilentlyContinue))
    {
        Write-Host 'Failed' -ForegroundColor Yellow -NoNewline
        Write-Host ', registering default repository'
        Register-PSRepository -Default -Verbose:$Verbose.IsPresent
    }
    else
    {
        Set-PSRepository 'PSGallery' -InstallationPolicy Trusted
        Write-Host 'Succeeded' -ForegroundColor Green
    }
    if ($null -ne $Modules)
    {
        foreach ($entry in $Modules)
        {
            Write-Host "Searching for module '${entry}':".PadRight($outputwidth) -NoNewline
            if ($null -ne (Find-Module $entry -ErrorAction SilentlyContinue))
            {
                Write-Host 'Succeeded' -ForegroundColor Green
                if ($null -ne (Get-Module $entry -ListAvailable -ErrorAction SilentlyContinue))
                {
                    Write-Host "Updating module '${entry}':".PadRight($outputwidth) -NoNewline
                    Update-Module -Name $entry -Force:$Force.IsPresent -Verbose:$Verbose.IsPresent
                    Write-Host 'Succeeded' -ForegroundColor Green
                }
                else
                {
                    Write-Host "Installing module '${entry}':".PadRight($outputwidth) -NoNewline
                    Install-Module -Name $entry -AllowClobber -Force:$Force.IsPresent -Verbose:$Verbose.IsPresent
                    Write-Host 'Succeeded' -ForegroundColor Green
                }
            }
            else
            {
                Write-Host 'Failed' -ForegroundColor Yellow
            }
        }
    }
    else
    {
        Write-Host "Updating modules:".PadRight($outputwidth) -NoNewline
        Update-Module -Force:$Force.IsPresent -Verbose:$Verbose.IsPresent
        Write-Host 'Succeeded' -ForegroundColor Green
    }
    Write-Host "Updating help:".PadRight($outputwidth) -NoNewline
    Update-Help -Force:$Force.IsPresent -Verbose:$Verbose.IsPresent -ErrorAction SilentlyContinue
    Write-Host 'Succeeded' -ForegroundColor Green
}
else
{
    Write-Host 'Failed' -ForegroundColor Red
}
