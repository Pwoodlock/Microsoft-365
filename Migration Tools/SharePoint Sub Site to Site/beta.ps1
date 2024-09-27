<#
.SYNOPSIS
Executes tasks related to SharePoint document library migration including directory creation, logging, and content export using app-only authentication.

.DESCRIPTION
This script:
1. Checks if a specified directory exists and creates it if it doesn't.
2. Updates LogSettings to reflect the new directory path.
3. Provides a function for formatted log message output.
4. Connects to a source tenant using app-only authentication.
5. Retrieves all document libraries from the specified site.
6. Exports the content of each document library.

.PARAMETER None

.EXAMPLE
PS C:\> .\MigrateSharePointContent.ps1
Executes the script to perform the described tasks.

.NOTES
- Replace placeholders with actual values for $clientId, $clientSecret, $tenantId, $siteUrl, and $certificatePath.
- Requires the SharePointPnPPowerShellOnline module.

.LINK
https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/?view=sharepoint-ps
#>

# Required module check
if (-not (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline)) {
    Install-Module -Name SharePointPnPPowerShellOnline -Force
}

# Directory setup
$directoryPath = "C:\MigrationData"
$LocalDirectoryPath = Join-Path -Path $directoryPath -ChildPath "SiteContent"
$LogFilePath = Join-Path -Path $directoryPath -ChildPath "Log.txt"

# Ensure directory existence
$directoryPath, $LocalDirectoryPath | ForEach-Object {
    if (-not (Test-Path -Path $_)) {
        New-Item -ItemType Directory -Path $_ | Out-Null
    }
}

# Log message function
function Log-Message {
    param (
        [string]$MessageKey,
        [string[]]$Arguments
    )

    $messageTemplate = @{
        Connecting = "Initiating connection to {0}..."
        Connected  = "Successfully connected to {0}."
        Exporting  = "Starting export of {0} from {1} to local CSV files..."
        Exported   = "Export of {0} from {1} completed successfully."
    }

    $messageText = $messageTemplate[$MessageKey] -f $Arguments
    if ($MessageKey -in 'Connecting', 'Exporting') {
        $color = 'Yellow'
    } else {
        $color = 'Green'
    }

    Write-Host $messageText -ForegroundColor $color
    Add-Content -Path $LogFilePath -Value $messageText
}

# SharePoint connection
$clientId = "<your_client_id>"
$clientSecret = "<your_client_secret>"
$tenantId = "<your_tenant_id>"
$siteUrl = "<your_site_url>"
$certificatePath = "<path_to_certificate>"

Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $certificatePath -Tenant $tenantId
Log-Message -MessageKey 'Connected' -Arguments @($siteUrl)

# Document library processing
$DocumentLibraries = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }

foreach ($Library in $DocumentLibraries) {
    $LibraryPath = Join-Path -Path $LocalDirectoryPath -ChildPath $Library.Title
    if (-not (Test-Path -Path $LibraryPath)) {
        New-Item -ItemType Directory -Path $LibraryPath | Out-Null
    }

    Log-Message -MessageKey 'Exporting' -Arguments @($Library.Title, $siteUrl)
    
    $Files = Get-PnPFile -FolderSiteRelativeUrl $Library.RootFolder.ServerRelativeUrl -AsListItem
    foreach ($File in $Files) {
        $FileServerRelativeUrl = $File["FileRef"]
        $FileName = [System.IO.Path]::GetFileName($FileServerRelativeUrl)
        $LocalFilePath = Join-Path -Path $LibraryPath -ChildPath $FileName
        
        Get-PnPFile -Url $FileServerRelativeUrl -Path $LocalFilePath -AsFile -Force
    }
    
    Log-Message -MessageKey 'Exported' -Arguments @($Library.Title, $siteUrl)
}

# Finalization
Write-Host "Migration completed successfully." -ForegroundColor Green
