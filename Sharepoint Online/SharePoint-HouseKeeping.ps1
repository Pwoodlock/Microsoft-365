#===================================================================================================================================================================================
#                                                             ROCKFIELD IT USER TEMPLATE FOR SHAREPOINT & PnP 2023
#
#       Date: 18/12/2023 v0.91
#
#   ++ Added Second layer of indexing for SharePoint sites that are over 5000 items.  This is to help even further with the throttling that occurs with SharePoint Online with messy
#      sites.
#   ++ Rescoped Permissions according to to MS365 best practices 
#
#
#   TODO:
#   && Need to add user input for line: 162 $pagesize, $batchsize, $sleeptime
#   && Display logged in User Permissions for the site
#
#                            This is for testing only so I don't  have to keep inserting the site and list name each time I run the script.
#$SiteURL = "https://#####.sharepoint.com/sites/#####/"
#$ListName = "Documents"
#
#
#===================================================================================================================================================================================
# Get the console window size
$windowWidth = $Host.UI.RawUI.WindowSize.Width
$windowHeight = $Host.UI.RawUI.WindowSize.Height
# Function to write centered text
function Write-Centered {
    param (
        [string]$Text,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::White
    )
    
    $xPosition = [math]::Round(($windowWidth - $Text.Length) / 2)
    $yPosition = $Host.UI.RawUI.CursorPosition.Y

    $Host.UI.RawUI.CursorPosition = New-Object -TypeName System.Management.Automation.Host.Coordinates -ArgumentList ($xPosition, $yPosition)
    Write-Host $Text -ForegroundColor $ForegroundColor -NoNewline
    Write-Host ("") # Reset to left margin
}

#    ----------------------------------    EDIT THE BELOW TO SUIT YOUR NEEDS REGARDING VERSION NO. ETC.   --------------------------------


Write-Centered "*** Rockfield IT Services ***" -ForegroundColor Green
Write-Host ("`n" * 1)

Write-Centered "Sharepoint ==> Versioning Tools & Bin Purge" -ForegroundColor Green
Write-Host ("`n" * 1)
Write-Centered "Bugs or issues email: patrick.woodlock@rockfieldit.com" -ForegroundColor Green


#    ---------------------------    CHECKING PERMISSIONS AND IF YOU HAVE THE REQUIRED MODULES INSTALLED FOR THE SCRIPT   --------------------------------
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
#========================================================================================================================
#  END OF ROCKFIELD IT USER TEMPLATE FOR SHAREPOINT & PnP 
#========================================================================================================================


# Prompt for SharePoint site and List Name
$SiteURL = (Read-Host "Please enter the SharePoint Site URL (e.g., https://Tenant.sharepoint.com/sites/KinderSurprise)").Trim()
$ListName = (Read-Host "Please enter the SharePoint List Name (e.g., Documents)").Trim()

# Store the site and list names for the current session goinog forward
$sessionSettings = @{
    SiteURL  = $SiteURL
    ListName = $ListName
}
###############################################################################################################################


# Connect to PnP Online with the latest version of the PnP PowerShell module
Connect-PnPOnline -Url $SiteURL -Interactive

# User-defined Rate Limiting to help with SharePoint Online throttling which could occur with large lists!!!!
$batchSize = Read-Host "Enter the batch size for operations before a delay (e.g., 10)"
$sleepTime = Read-Host "Enter the delay time in seconds after the above operations has been completed (e.g., 5)"

# Validate input
if (![int]::TryParse($batchSize, [ref]$null) -or ![int]::TryParse($sleepTime, [ref]$null)) {
    Write-Host "Invalid input. Batch size and delay time should be integers." -ForegroundColor Red
    exit
}

# Get the Document Library
$List = Get-PnPList -Identity $ListName

# Display Versioning Settings
if ($List.EnableVersioning) {
    Write-Host "Versioning Enabled: $($List.EnableVersioning)" -ForegroundColor Green
    Write-Host "Number of Major Versions to Keep: $($List.MajorVersionLimit)" -ForegroundColor Green
} else {
    Write-Host "Versioning Enabled: $($List.EnableVersioning)" -ForegroundColor Red
}

# Prompt to set Major Version Limit
$newLimit = Read-Host "Do you want to change the Major Version Limit? (Enter new value or press Enter to keep)"
if (![string]::IsNullOrWhiteSpace($newLimit)) {
    Set-PnPList -Identity $ListName -MajorVersions $newLimit
    Write-Host "Major Version Limit updated to: $newLimit" -ForegroundColor Green
}

# Display Minor Versioning Settings
if ($List.EnableMinorVersions) {
    Write-Host "Minor Versioning Enabled: $($List.EnableMinorVersions)" -ForegroundColor Green

    # Ask the user if they want to disable minor versioning
    $disableMinorVersion = Read-Host "Do you want to disable Minor Versioning? (Y/N)"
    
    if ($disableMinorVersion -match '^(Y|y)$') {
        # Disable Minor Versioning
        Set-PnPList -Identity $ListName -EnableMinorVersions $false
        Write-Host "Minor Versioning has been disabled." -ForegroundColor Green
    } elseif ($disableMinorVersion -match '^(N|n)$') {
        Write-Host "Minor Versioning remains enabled." -ForegroundColor Yellow
    } else {
        Write-Host "Invalid input. No changes made to Minor Versioning settings." -ForegroundColor Red
    }

} else {
    Write-Host "Minor Versioning Enabled: $($List.EnableMinorVersions)" -ForegroundColor Red

    # Ask the user if they want to enable minor versioning
    $enableMinorVersion = Read-Host "Do you want to enable Minor Versioning? (Y/N)"
    
    if ($enableMinorVersion -match '^(Y|y)$') {
        # Enable Minor Versioning
        Set-PnPList -Identity $ListName -EnableMinorVersions $true
        Write-Host "Minor Versioning has been enabled." -ForegroundColor Green
    } elseif ($enableMinorVersion -match '^(N|n)$') {
        Write-Host "Minor Versioning remains disabled." -ForegroundColor Yellow
    } else {
        Write-Host "Invalid input. No changes made to Minor Versioning settings." -ForegroundColor Red
    }
}

Write-Host "Please read the above information carefully before continuing." -ForegroundColor Yellow

Write-Host "Press any key to continue ..."
$null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")


### Delete old versions of all files in the Document Library ###


$ListName = "Documents"
$pageSize = 2000 # Adjust as needed, keeping it below the threshold
$allItems = @()

# Retrieve items in pages
do {
    $pagedItems = Get-PnPListItem -List $ListName -PageSize $pageSize
    $allItems += $pagedItems
} while ($pagedItems -ne $null -and $pagedItems.Count -eq $pageSize)


# Now $allItems contains all items in batches
foreach ($item in $allItems) {
    Write-Host ("Checking file " + $item["FileRef"]) -ForegroundColor Green
    
    try {
        # Fetch all versions of the file
        $versions = Get-PnPFileVersion -Url $item["FileRef"]
        $versionCount = $versions.Count

        if ($versionCount -gt $newLimit) {
            Write-Host "Found $versionCount versions, reducing to latest $newLimit versions." -ForegroundColor Green

            # Skip the latest $newLimit versions and delete the older ones
            $versions | Select-Object -Skip $newLimit | ForEach-Object {
                Write-Host "Deleting version $($_.VersionLabel)" -ForegroundColor Green
                Remove-PnPFileVersion -Url $item["FileRef"] -Identity $_.VersionLabel -Force

                # Increment the counter after each operation
                $counter++

                # Check if the counter reached the batchSize
                if ($counter -eq $batchSize) {
                    Write-Host "Pausing for $sleepTime seconds to avoid throttling..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $sleepTime

                    # Reset the counter
                    $counter = 0
                }
            }
            Write-Host "Old versions removed based on new major version limit." -ForegroundColor Green
        } else {
            Write-Host "Found $versionCount versions, no action needed." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
    }
}

# Warning message to the user about the impact of emptying the SharePoint Recycle Bin
Write-Host "WARNING: Emptying the SharePoint Recycle Bin is a non-reversible operation. All items will be permanently deleted. Please review the Recycle Bin contents before proceeding." -ForegroundColor Red

# Ask for user confirmation
$confirmation = Read-Host "Are you sure you want to proceed? (Y/N)"

# Validate the user input
if ($confirmation -match '^(Y|y|N|n)$') {
    switch ($confirmation) {
        { $_ -match '^(Y|y)$' } {
            # Ask the user about emptying the SharePoint Recycle Bin
            $actionChoice = Read-Host "Do you want to empty the SharePoint Recycle Bin? (Y/N)"
            
            if ($actionChoice -match '^(Y|y|N|n)$') {
                switch ($actionChoice) {
                    { $_ -match '^(Y|y)$' } {
                        Write-Host "Emptying SharePoint Recycle Bin..." -ForegroundColor Green
                        # Empty the Recycle Bin
                        Clear-PnPRecycleBinItem -All -Force
                    }
                    { $_ -match '^(N|n)$' } {
                        Write-Host "No action taken on the Recycle Bin." -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "Invalid choice. Exiting..." -ForegroundColor Red
            }
        }
        { $_ -match '^(N|n)$' } {
            Write-Host "Operation cancelled by the user." -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "Invalid choice. Exiting..." -ForegroundColor Red
}

Write-Host "That's it. The script has finished and given the nature of the Sharepoint upates, you might not see the results for at least 3 days!!!  Any questions please contact patrick.woodlock@rockfieldit.com" -ForegroundColor Cyan


#   The below code is for verification purposes only for when working with small Sharepoint sites.  Uncomment to use but be aware the process is
#   basically doubled in time.  This is because the script is first deleting the versions and then verifying the versions !

# Now lets add an Verification step to the mix !!
#Write-Host "Verifying version deletions..." -ForegroundColor Green
#foreach ($item in $allItems) {
#    Write-Host ("Verifying file " + $item["FileRef"]) -ForegroundColor Green
#    try {
#        $verifyVersions = Get-PnPFileVersion -Url $item["FileRef"]
#        $verifyCount = $verifyVersions.Count#

#        if ($verifyCount -le $newLimit) {
#            Write-Host "Successfully reduced to $verifyCount versions." -ForegroundColor Cyan
#        } else {
#            Write-Host "$verifyCount versions found, verification failed." -ForegroundColor Red
#        }
#    } catch {
#        Write-Host "An error occurred during verification: $_" -ForegroundColor Red
#    }
#}



# Disconnect from PnP Online When we are done! Just uncomment the below line to use!
#     Disconnect-PnPOnline