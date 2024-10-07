# Author: Elbert Beverdam

# Synopsis:
# This PowerShell script registers an Azure AD application for PnP PowerShell using Microsoft Graph. 
# It checks if an application with the specified name already exists and gives the user the option to skip or recreate the app.

param (
    [string]$tenant = ""
)
# Check if the script is running in PowerShell 7
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or higher. Please run or install PowerShell 7 and try again." -ForegroundColor Red
    Write-Host "You can download and install PowerShell 7 from the following link:" -ForegroundColor Yellow
    Write-Host "https://aka.ms/powershell-release?tag=stable" -ForegroundColor Cyan
    exit 1
}
if (-not $tenant) {
    Write-Host -ForegroundColor Red "Error: Tenant is not configured. Use parameter -tenant ""tenantname"" without .sharepoint.com suffix or update script."
    exit 1
}

# Scope
$WorkingDir = "C:\ProgramData\PnP-PowerShell"
$clientFile = "$WorkingDir\PnP-PowerShell-$tenant.txt"

# Ensure the Microsoft.Graph module is installed and imported
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
    Install-Module -Name Microsoft.Graph.Applications -Force -Scope CurrentUser
}
Import-Module Microsoft.Graph.Applications

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome

# Check if the application already exists
$existingApp = Get-MgApplication -Filter "displayName eq 'PnP-PowerShell-$tenant'"
if ($existingApp) {
    Write-Host "Application 'PnP-PowerShell-$tenant' already exists."
    $choice = Read-Host "Do you want to recreate the app? (y/n)"
    if ($choice -eq "n") {
        Write-Host "Skipping app creation."
        exit
    } else {
        Write-Host "Deleting existing app..."
        Remove-MgApplication -ApplicationId $existingApp.Id
        Start-Sleep -Seconds 10
        Write-Host "Application 'PnP-PowerShell-$tenant' deleted. "
    }
}

# Register the application
try {

        # Register
        Set-Location -Path $WorkingDir
        $password = ConvertTo-SecureString -String "password" -AsPlainText -Force
        $reg = Register-PnPAzureADApp -ApplicationName "PnP-PowerShell-$tenant" -Tenant "$tenant.onmicrosoft.com" -CertificatePassword $password -Interactive
        mkdir $WorkingDir -ErrorAction SilentlyContinue
        $reg."AzureAppId/ClientId" | Out-File $clientFile
    
    Write-Host "Successfully registered application 'PnP-PowerShell-$tenant' with AppId: $($reg."AzureAppId/ClientId")"
} catch {
    Write-Host -ForegroundColor Red "Failed to register the application: $_"
}
