<#
.SYNOPSIS
    Executes a process to enable Auto Expanding Archive for a selected mailbox in Exchange Online. While also offering the user to enable it globally for the tenant as of 05/04/2024

.DESCRIPTION
    This script retrieves a list of mailboxes from Exchange Online and displays them with corresponding numbers. 
    It prompts the user to select a mailbox by entering the corresponding number. 
    If a valid selection is made, it enables In-Place Archive for the selected mailbox if it is not already enabled. 
    Then, it enables Auto Expanding Archive for the selected mailbox. 
    The script provides feedback on the status of each step and handles any errors that occur.

.INPUTS
    None

.OUTPUTS
    None

.EXAMPLE
    .\EmailArchiveAutoExpand.ps1
    This example demonstrates how to execute the script to enable Auto Expanding Archive for a selected mailbox.

.NOTES
    Author: Patrick Woodlock
    Date: 05/04/2024
    Version: 1.1

#>

# Ensure the script is running as an administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "Please run this script as an administrator." -ForegroundColor Red
    exit
}

# Install required modules if they are not present
if (-not (Get-Module -ListAvailable -Name "ExchangeOnlineManagement")) {
    Install-Module -Name ExchangeOnlineManagement -Force -SkipPublisherCheck
}
# Connect to Exchange Online
Connect-ExchangeOnline
# Again the below seems to be causing issues auto reverting to to the statment of Write-Host "Failed to connect to Exchange Online. Please check your credentials and internet connection." -ForegroundColor Red
# I think the issue is lying where it's not waiting for the user to perform the MFA in their browser so a timer or a better check function here should fix the issue.  But no time today.  

#try {
#    $connection = Connect-ExchangeOnline -ShowProgress $true
#    if (-not $connection) {
#        Write-Host "Failed to connect to Exchange Online. Please check your credentials and internet connection." -ForegroundColor Red
#        exit
#    } else {
#        Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
#    }
#} catch {
#    Write-Host "Error connecting to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
#    exit
#}

# Function to check and enable Auto Archive globally
function CheckAndEnableGlobalAutoArchive {
    try {
        $autoExpandEnabled = Get-OrganizationConfig | Select-Object -ExpandProperty AutoExpandingArchiveEnabled

        if (-not $autoExpandEnabled) {
            $enableGlobal = Read-Host "Auto Expanding Archive is not enabled globally. Would you like to enable it? (y/n)"
            if ($enableGlobal -ieq 'y') {
                Set-OrganizationConfig -AutoExpandingArchiveEnabled $true
                Write-Host "Auto Expanding Archive has been enabled globally." -ForegroundColor Green
            } else {
                Write-Host "Auto Expanding Archive remains disabled globally." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Auto Expanding Archive is already enabled globally." -ForegroundColor Green
        }
    } catch {
        Write-Host "An error occurred while checking or enabling global Auto Expanding Archive: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Call the function to check and enable Auto Archive globally
CheckAndEnableGlobalAutoArchive

function RunProcess {
    try {
        $mailboxes = Get-ExoMailbox -ResultSize Unlimited |
                     Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' } |
                     Select-Object DisplayName, ArchiveDatabase, ArchiveStatus

        if (-not $mailboxes) {
            Write-Host "No mailboxes found or unable to retrieve mailboxes." -ForegroundColor Red
            return
        }

        Write-Host ("`n" * 2) -NoNewline
        Write-Host "List of users with Exchange Licenses:" -ForegroundColor Cyan
        $count = 0
        $mailboxes | ForEach-Object {
            $count++
            Write-Host ("{0}. {1} - {2}" -f $count, $_.DisplayName, $_.ArchiveStatus) -ForegroundColor Yellow
        }

        $selectedNumber = Read-Host "`nEnter the number corresponding to the mailbox you'd like to manage"
        if ($selectedNumber -lt 1 -or $selectedNumber -gt $mailboxes.Count) {
            Write-Host "Invalid selection. Exiting." -ForegroundColor Red
            exit
        }

        $selectedMailbox = $mailboxes[$selectedNumber - 1]
        Write-Host "`nYou've selected $($selectedMailbox.DisplayName) for enabling Auto Expanding Archive." -ForegroundColor Green

        $confirmation = Read-Host "`nWould you like to enable Auto Expanding Archive for the selected mailbox? (y/n)"
        if ($confirmation -ieq 'y') {
            if ($selectedMailbox.ArchiveDatabase -eq $null) {
                Write-Host "Enabling In-Place Archive for $($selectedMailbox.DisplayName)..." -ForegroundColor Yellow
                Enable-Mailbox -Identity $selectedMailbox.DisplayName -Archive
                Write-Host "In-Place Archive has been enabled." -ForegroundColor Green
            } else {
                Write-Host "In-Place Archive is already enabled for $($selectedMailbox.DisplayName)." -ForegroundColor Green
            }

            Enable-Mailbox -Identity $selectedMailbox.DisplayName -AutoExpandingArchive
            Write-Host "Auto Expanding Archive has been enabled for $($selectedMailbox.DisplayName)." -ForegroundColor Green
        } else {
            Write-Host "Auto Expanding Archive was not enabled." -ForegroundColor Red
        }
    } catch {
        Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
    }
}

do {
    RunProcess
    $userInput = Read-Host "`nWould you like to pick and activate another user? (y/n)"
} while ($userInput -ieq 'y')

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
