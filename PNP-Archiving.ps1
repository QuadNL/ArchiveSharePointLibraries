<# Author: Elbert Beverdam 2024

# Synopsis:
# This PowerShell script, inspired by Salaudeen Rajack's work on connecting to SharePoint Online using Azure AD App ID from PowerShell, automates the process of downloading files and folders from SharePoint Online document libraries. 
The script ensures that the necessary Azure AD app, `PnP-PowerShell-$tenant`, is registered. 
For more information on setting up the required Azure resources, refer to other scripts like `PnP-Register.ps1`.

# Key Features:
# 1. Folder and File Download:
#    - Recursively downloads files and folders from specified SharePoint Online document libraries.
#    - Skips files that already exist locally with the same size to avoid redundant downloads.

# 2. Logging and Error Handling:
#    - Logs the number of files downloaded, folders created, and files skipped to a CSV file.
#    - Handles errors gracefully and stops processing further URLs if an error occurs.

# 3. Archived File Creation:
#    - Creates an `.Archived` file in the root folder of each document library after successful download to mark it as processed.

# Usage:
# - Ensure the `PnP-PowerShell-$tenant` Azure AD app is registered.
# - Run `PnP-Register.ps1` to create the necessary Azure resources if not already done.
# - Update the script parameters such as `$tenant`, `$WorkingDir`, `$clientFile`, `$CertificatePath`, `$siteUrlFilePath`, and `$DownloadPath` as per your environment.
# - Run 'PnP-Archiving.ps1 with parameters  -tenant <sharepointtenantname> -SiteUrlPath "<location to txt file>" e.g. -tenant contoso
# - Execute the script to start downloading files and folders from the specified SharePoint Online sites.

# Source: Salaudeen Rajack, Connect to SharePoint Online using Azure AD App ID from PowerShell (https://www.sharepointdiary.com/2022/10/connect-to-sharepoint-online-using-azure-ad-app-id-from-powershell.html)

# KNOWN BUGS
### Filecomparisation not working

#>

param (

    [string]$tenant = "",
    [string]$DownloadPath = ""
)

$ErrorActionPreference = 'Stop'

# Check if the script is running in PowerShell 7
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or higher. Please run or install PowerShell 7 and try again." -ForegroundColor Red
    Write-Host "You can download and install PowerShell 7 from the following link:" -ForegroundColor Yellow
    Write-Host "https://aka.ms/powershell-release?tag=stable" -ForegroundColor Cyan
    exit 1
}
if (-not $DownloadPath) {
    Write-Host -ForegroundColor Red  "Error: DownloadPath is not configured. Use parameter -DownloadPath ""\\server\share"" or update script."
    exit 1
}
if (-not $tenant) {
    Write-Host -ForegroundColor Red "Error: Tenant is not configured. Use parameter -tenant ""tenantname"" without .sharepoint.com suffix or update script."
    exit 1
}
Function Download-SPOFolder([Microsoft.SharePoint.Client.Folder]$Folder, $DestinationFolder, [ref]$FileCounter, [ref]$FolderCounter, [ref]$SkippedCounter)
{
    # Get the Folder's Site Relative URL
    $FolderURL = $Folder.ServerRelativeUrl.Substring($Folder.Context.Web.ServerRelativeUrl.Length)
    $LocalFolder = $DestinationFolder + ($FolderURL -replace "/","\")
    
    # Create Local Folder, if it doesn't exist
    If (!(Test-Path -Path $LocalFolder)) {
        New-Item -ItemType Directory -Path $LocalFolder | Out-Null
        $FolderCounter.Value++
    }
    
    # Get all Files from the folder
    $FilesColl = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -ItemType File

    # Iterate through each file and download
    Foreach($File in $FilesColl)
    {
        $LocalFilePath = Join-Path -Path $LocalFolder -ChildPath $File.Name
        
        # Check if the file exists locally
        if (Test-Path -Path $LocalFilePath) {
            # Compare file sizes to verify the local copy
            $LocalFileSize = (Get-Item $LocalFilePath).Length
            $OnlineFileSize = $File.Length
            
            if ($LocalFileSize -eq $OnlineFileSize) {
                $SkippedCounter.Value++  # Increment skipped counter
                continue
            } else {
                Write-Host -ForegroundColor Yellow "`tFile '$($File.Name)' exists but differs in size. Downloading again."
            }
        }
        
        # Download the file
        Get-PnPFile -ServerRelativeUrl $File.ServerRelativeUrl -Path $LocalFolder -FileName $File.Name -AsFile -Force
        $FileCounter.Value++  # Increment downloaded files counter
    }
    
    # Get Subfolders of the Folder and call the function recursively
    $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderURL -ItemType Folder
    Foreach ($SubFolder in $SubFolders | Where-Object {$_.Name -ne "Forms"})
    {
        Download-SPOFolder $SubFolder $DestinationFolder $FileCounter $FolderCounter $SkippedCounter  # No need for [ref] here again
    }
}


# Scope and input parameters
$WorkingDir = "C:\ProgramData\PnP-PowerShell"
$siteUrlFilePath = "$WorkingDir\$tenant-input.txt"
$clientFile = "$WorkingDir\PnP-PowerShell-$tenant.txt"
$CertificatePath = "$WorkingDir\PnP-PowerShell-$tenant.pfx"
$csvLogPath = Join-Path -Path $WorkingDir -ChildPath "SharePointDownloadLog.csv"  # CSV log path
$password = "password" # Certificate Password generated with PNP-Connect.ps1 script.
$secPassword = $password | ConvertTo-SecureString -AsPlainText -Force

# Initialize a flag to track if errors occur
$hasError = $false

# Initialize CSV log file with headers
If (!(Test-Path $csvLogPath)) {
    "URL,FilesDownloaded,FoldersCreated,FilesSkipped" | Out-File $csvLogPath
}

# Get Client ID from file
$clientId = Get-Content $clientFile

# Check if the file exists
if (-Not (Test-Path -Path $siteUrlFilePath)) {
    Write-Host "The file '$tenant-input.txt' does not exist." -ForegroundColor Red
    exit 1
}

# Check if the file is empty
if ((Get-Content -Path $siteUrlFilePath).Length -eq 0) {
    Write-Host "The file '$tenant-input.txt' is empty." -ForegroundColor Red
    exit 1
}

# Read URLs from the text file
$siteUrls = Get-Content -Path $siteUrlFilePath

Write-Host -ForegroundColor Cyan "Following SharePoint Site(s) found:" 
$siteUrls | Format-Table | Out-String | Write-Host -ForegroundColor Yellow
Write-Host -ForegroundColor Yellow "Starting processing the archiving process."

# Begin processing each site URL
foreach ($SiteURL in $siteUrls) {
    try {
        # Connect to SharePoint Online using certificate-based authentication
        Connect-PnPOnline -Url $SiteURL -ClientId $clientId -Tenant "$tenant.onmicrosoft.com" -CertificatePath $CertificatePath -CertificatePassword $secPassword

        # Get all document libraries in the site
        $documentLibraries = Get-PnPList | Where-Object { 
            $_.BaseTemplate -eq 101 -and 
            $_.RootFolder.ServerRelativeUrl -notmatch "/FormServerTemplates|/SiteAssets|/Style Library|/SiteCollectionDocuments"
        }

        # Get the root web of the site
        $web = Get-PnPWeb
        $ProjectName = Join-Path -Path $DownloadPath -ChildPath "$($web.Title)"

        # Iterate through each document library and download its contents
        foreach ($library in $documentLibraries) {
            $FolderServerRelativeURL = $library.RootFolder.ServerRelativeUrl

            # Check if .Archived file exists in the document library
            $archivedFile = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderServerRelativeURL -ItemType File | Where-Object { $_.Name -eq ".Archived" }
            if ($archivedFile) {
                Write-Host -ForegroundColor Yellow "Skipping '$($library.Title)' as it already contains an .Archived file."
                continue
            }

            # Initialize counters for the current library
            $FileCounter = [ref] 0
            $FolderCounter = [ref] 0
            $SkippedCounter = [ref] 0

            # Get the folder and start downloading
            $Folder = Get-PnPFolder -Url $FolderServerRelativeURL
            Write-Host -ForegroundColor Yellow "Archiving '$($FolderServerRelativeURL)' to $ProjectName."
            Download-SPOFolder $Folder $ProjectName $FileCounter $FolderCounter $SkippedCounter

            # Create and place .Archived file in the root folder of the current document library
            $archivedContent = "Archived on $(Get-Date)"
            $archivedFilePath = Join-Path -Path $WorkingDir -ChildPath ".Archived"
            Set-Content -Path $archivedFilePath -Value $archivedContent

            try {
                Add-PnPFile -Path $archivedFilePath -Folder $FolderServerRelativeURL
                Write-Host -ForegroundColor Green "Placed .Archived file in the root folder of the library '$($library.Title)'."
            } catch {
                Write-Host -ForegroundColor Red "Failed to place .Archived file in the root folder of the library '$($library.Title)'."
                $hasError = $true
                break
            }

            # Log summary for the current document library
            Write-Host -ForegroundColor Cyan "Summary for '$FolderServerRelativeURL':"
            Write-Host "    Files downloaded: $($FileCounter.Value)"
            Write-Host "    Folders created: $($FolderCounter.Value)"
            Write-Host "    Files skipped: $($SkippedCounter.Value)"
            
            # Write the summary to the CSV log file
            "$FolderServerRelativeURL,$($FileCounter.Value),$($FolderCounter.Value),$($SkippedCounter.Value)" | Out-File -Append -FilePath $csvLogPath
        }

    } catch {
        Write-Host -ForegroundColor Red "An error occurred while processing '$SiteURL'. Error details: $_"
        $hasError = $true
        break  # Stop processing further URLs if an error occurs
    }
}

# If no errors, clear the input file to prevent future downloads
if (-not $hasError) {
    Clear-Content $siteUrlFilePath
    Write-Host -ForegroundColor Green "All URLs processed successfully. Input file cleared."
} else {
    Write-Host -ForegroundColor Red "An error occurred. Input file NOT cleared."
}
