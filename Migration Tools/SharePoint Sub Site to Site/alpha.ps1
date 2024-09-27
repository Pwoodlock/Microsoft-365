#===================================================================================================================================================================================
#                                                          ROCKFIELD IT QuickStart USER TEMPLATE FOR SHAREPOINT & PnP by "Patrick Woodlock"
#===================================================================================================================================================================================

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host -ForegroundColor Red "Please run this script as an administrator."
    exit
}
# Check PowerShell version
$psVersion = $PSVersionTable.PSVersion.Major

if ($psVersion -lt 7) {
    Write-Host -ForegroundColor Red "PowerShell version $psVersion is not supported. Please download and install PowerShell 7 or later."
    exit
}

Write-Host "PowerShell version: $psVersion" -ForegroundColor Green

# Check and install PnP PowerShell module
$requiredModules = @("PnP.PowerShell")
$missingModules = $requiredModules | Where-Object { -not (Get-Module $_ -ListAvailable) }

if ($missingModules.Count -gt 0) {
    Write-Host -ForegroundColor Red "The following required modules are missing: $($missingModules -join ', ')"
    $installModules = Read-Host "Do you want to install the required modules? (Y/N)"
    
    if ($installModules -eq "Y") {
        $missingModules | ForEach-Object {
            Install-Module -Name $_ -Scope CurrentUser -Force
        }
        Write-Host "Required modules installed." -ForegroundColor Green
    } else {
        Write-Host -ForegroundColor Red "Script cannot continue without the required modules."
        exit
    }
}
#===================================================================================================================================================================================
#                                                                   END OF ROCKFIELD IT USER TEMPLATE FOR SHAREPOINT & PnP 
#===================================================================================================================================================================================
# Created for Rockfield IT Services
# Initial draft of final script to migrate data from one tenant to another tenant. All the mapping, downloading, user * group creation and bascially everything I can think of has been tested here on the Rockfield Side and works!
# This script is a work in progress and is not yet complete, and im fucking tired so I'm going to bed now. Will clean up more later. Damn you Shiels and not getting back to us sooner!!!
# version: 20231017

#
#
# Source Tenant
$SourceSiteURL = "https://rockfieldit.sharepoint.com/sites/ShielsDemo"
Connect-PnPOnline -Url $SourceSiteURL -Interactive

#
#
# Now lets download the user data from both tenants and combine them into one CSV file to get ready for the export to the target tenant.
# 
#

# Define variables
$TargetTenantURL = "SHIELS MASTER ON.MICROSOFT DOMAIN HERE"
$MigrationDataPath = "C:\MigrationData"
$UsersCSVPath = Join-Path -Path $MigrationDataPath -ChildPath "Users.csv"

# Step 1: Authentication
# Connect to the target tenant
Connect-PnPOnline -Url $TargetTenantURL -Interactive

# Step 2: Fetching Users and Groups from Target Tenant
$users = Get-PnPUser
$usersData = @()

foreach ($user in $users) {
    $userData = @{
        "Email" = $user.Email
        "Title" = $user.Title
        "Groups" = (Get-PnPTeamsTab -Team $user.Title | Select-Object -ExpandProperty DisplayName) -join ';'
    }
    $usersData += New-Object PSObject -Property $userData
}

# Step 3: Storing the Information
$usersData | Export-Csv -Path $UsersCSVPath -NoTypeInformation


# Define the local directory where content will be exported
$LocalDirectoryPath = "C:\MigrationData"


# Check if the directory exists and create it if it doesn't
$directoryPath = "C:\MigrationData"
if (-not (Test-Path -Path $directoryPath)) {
    New-Item -ItemType Directory -Path $directoryPath | Out-Null
}

# Now proceed with exporting the users to the CSV file
$SourceUsers = Get-PnPUser
$csvPath = Join-Path -Path $directoryPath -ChildPath "Users.csv"
$SourceUsers | Export-Csv -Path $csvPath -NoTypeInformation


# For importing users to the target tenant, 
# it's assumed that user accounts are pre-created in the target tenant.
#$TargetUsers = Import-Csv -Path "C:\MigrationData\Users.csv"
#foreach ($User in $TargetUsers) {
#    # The actual cmdlet to add users to groups or assign permissions may vary
#    # This is a placeholder to indicate where you would add users to the target tenant
#    # Actual implementation would depend on your specific environment and requirements
#}

# Similarly for groups:
# Export groups from source tenant
$SourceGroups = Get-PnPGroup
$SourceGroups | Export-Csv -Path "C:\MigrationData\Groups.csv" -NoTypeInformation

# Import groups to target tenant
#$TargetGroups = Import-Csv -Path "C:\MigrationData\Groups.csv"
#foreach ($Group in $TargetGroups) {
    # Similarly, this is a placeholder to indicate where you would create groups in the target tenant
    # The actual cmdlet to create groups may vary
    # Actual implementation would depend on your specific environment and requirements
#}


# Export site structure from source tenant
$SiteTemplateFile = "C:\MigrationData\SiteTemplate.xml"
Get-PnPSiteTemplate -Out $SiteTemplateFile -IncludeAllClientSidePages -IncludeSiteGroups

# Get all document libraries in the source site
$DocumentLibraries = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false }

# Iterate through each document library and export its content
foreach ($Library in $DocumentLibraries) {
    $LibraryPath = Join-Path -Path $LocalDirectoryPath -ChildPath $Library.Title
    New-Item -ItemType Directory -Path $LibraryPath | Out-Null  # Create a directory for each library
    
    $Files = Get-PnPListItem -List $Library.Title | Where-Object { $_.FileSystemObjectType -eq "File" }
    foreach ($File in $Files) {
        $FileServerRelativeUrl = $File.FieldValues.FileRef
        $FileName = $File.FieldValues.FileLeafRef
        $LocalFilePath = Join-Path -Path $LibraryPath -ChildPath $FileName
        Get-PnPFile -Url $FileServerRelativeUrl -FileName $FileName -Path $LibraryPath -AsFile
    }
}

# At this point, all document libraries and their files should be exported to the local directory "C:\MigrationData"
# Disconnect from the Source Tenant
Disconnect-PnPOnline
#
#
#
#
#
#  Now it's time to check everyting in the C:\MigrationData folder first, Check the data in the users and Groups CSV files, and then check the XML file.
#  Once you are happy with the data, you can proceed with the import. 
#  Also, make your life easier, and disable the security defaults on the target tenant for this migration and don't forget to turn it back on again.
#
#
#
#
#
#Import required module
Import-Module -Name PnP.PowerShell

# Define more variables
$TargetTenantURL = "https://<YourCustomerDomain>.sharepoint.com" # I need to check if their domain is is actually registered with Microsoft. If not, I'll need to register it and sort out the DNS records.
$TargetSiteURL = "$TargetTenantURL/sites/NewSite" # again, I will need to change this to the correct name.
$LocalDirectoryPath = "C:\MigrationData"

# Step 1: Authentication
# Connect to the source tenant with their Global Admin account!!
# Connect to the target tenant
Connect-PnPOnline -Url $TargetTenantURL -Interactive

# Step 2: Site Creation
# Create a new SharePoint site
New-PnPTeamSite -Url $TargetSiteURL -Title "New Site" -Description "Migrated Site for Shiels & Co" -IsPublic -Alias "NewSite" # I need to ask claire if she wants it public or private and also what exact name she wants it to be called.

# Step 3: User and Group Migration
# (Assuming you have a CSV file with user info and you probably want to cross-check it with the source tenant to make sure the users exist)
# Actually, I think I will download  the users from both tenants and then compare them, basically naming Claire as the owner with full privilages.
$UsersCSV = Import-Csv "C:\MigrationData\Users.csv"
foreach ($User in $UsersCSV) {
    # Create user or add to group as needed
    # Example: Add-PnPTeamsTeamUser -Team "NewSite" -User $User.Email
}

# Step 4: Content Migration
# Iterate through each directory in the local directory, creating a document library for each
$LocalLibraries = Get-ChildItem -Path $LocalDirectoryPath -Directory
foreach ($LocalLibrary in $LocalLibraries) {
    $LibraryName = $LocalLibrary.Name
    New-PnPList -Title $LibraryName -Template DocumentLibrary -Url $LibraryName

    # Upload files to the new document library. I looked in the directory and it's about 1gig in size for the video content.
    $Files = Get-ChildItem -Path $LocalLibrary.FullName -File
    foreach ($File in $Files) {
        $LocalFilePath = $File.FullName
        Add-PnPFile -Path $LocalFilePath -Folder $LibraryName
    }
}

# Disconnect from the target tenant
Disconnect-PnPOnline

