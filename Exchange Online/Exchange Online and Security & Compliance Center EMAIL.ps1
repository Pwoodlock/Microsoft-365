# Loop through the array and display each message
foreach ($message in $messages) {
    Write-Centered $message -ForegroundColor Green
}

# Ensure the script is running as an administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "Please run this script as an administrator." -ForegroundColor Red
    exit
}

# Check PowerShell version
$psVersion = $PSVersionTable.PSVersion.Major
if ($psVersion -lt 7) {
    Write-Host -ForegroundColor Red "PowerShell version $psVersion is not supported. Please download and install PowerShell 5 or later."
    exit
}

# Import the Exchange Online Management Module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online and Security & Compliance Center
Connect-ExchangeOnline -ShowBanner:$false
Connect-IPPSSession

# Get all mailboxes
$mailboxes = Get-ExoMailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress

# Display mailboxes and ask user to select one
$mailboxes | ForEach-Object {
    Write-Host ("[$index] " + $_.DisplayName + " (" + $_.PrimarySmtpAddress + ")")
    $index++
}

$selected = Read-Host "Please select a mailbox by entering the corresponding number"
$selectedMailbox = $mailboxes[$selected - 1].PrimarySmtpAddress

# Hard-coded date range
$startDate = "07/17/2014"
$endDate = "06/08/2023"

# Construct the search query with date range
$searchQuery = "(sent>=$startDate AND sent<=$endDate) OR (received>=$startDate AND received<=$endDate) AND (from:$selectedMailbox OR to:$selectedMailbox)"

# Use a unique name for the Compliance Search
$searchName = "SearchForUserEmails_" + (Get-Date -Format "yyyyMMddHHmmss")

# Create the compliance search
New-ComplianceSearch -Name $searchName -ExchangeLocation $selectedMailbox -ContentMatchQuery $searchQuery

# Start the compliance search
Start-ComplianceSearch -Identity $searchName

# Wait for the search to complete
do {
    Start-Sleep -Seconds 30
    $searchStatus = (Get-ComplianceSearch -Identity $searchName).Status
} while ($searchStatus -eq "NotStarted" -or $searchStatus -eq "Running")

# Display the number of items found
$itemCount = (Get-ComplianceSearch -Identity $searchName).Items
Write-Host "Number of items found: $itemCount"

# Provide options for actions
$options = @("Purge", "Soft Delete", "Hard Delete", "Exit")
$options | ForEach-Object { Write-Host ("[" + ($options.IndexOf($_) + 1) + "] " + $_) }
$action = Read-Host "Please select an action by entering the corresponding number"

# Take the selected action
switch ($action) {
    1 { # Purge
        New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete
        Write-Host "Purge action initiated. Items will be moved to the recoverable items folder."
    }
    2 { # Soft Delete
        New-ComplianceSearchAction -SearchName $searchName -SoftDelete
        Write-Host "Soft Delete action initiated. Items will be moved to the recoverable items folder."
    }
    3 { # Hard Delete
        New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete
        Write-Host "Hard Delete action initiated. Items will be permanently removed, bypassing the recoverable items folder."
    }
    4 { # Exit
        Write-Host "Exiting without taking any action."
    }
    default {
        Write-Host "Invalid option selected."
    }
}

# Retrieve the status of the most recent compliance search action from the portal
$recentSearchAction = Get-ComplianceSearchAction | Sort-Object -Property CreatedTime -Descending | Select-Object -First 1

# Display the status and other relevant details
$recentSearchAction | Format-List Name, Status, JobStartTime, JobEndTime, Errors