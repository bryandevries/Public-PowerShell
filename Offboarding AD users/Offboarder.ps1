# Author: https://github.com/bryandevries/

# Version 1.1
# Date initial version: 07-02-2023

# Changelog
# 13-11-2023 - Improved Group Export, added parameters to fucntions & added try/catch to FindADUser


# =============================================================================================================================================


# Function that searches for the user account in AD and asks for validation
function FindADUser {
    param (
        [parameter(Mandatory=$true)]
        [string]$UPN,
        [parameter(Mandatory=$true)]
        [string]$Wnummer
    )

    Clear-Host
    Write-host -ForegroundColor Green "Searching for user account $UPN"
    
    try {
        Get-ADUser -Identity $UPN -Properties * | Select-Object Name, UserPrincipalName, EmailAddress, Title, Department, Company | Format-List -ErrorAction Stop
    }
    catch {
        Write-Host -ForegroundColor Yellow -BackgroundColor DarkRed "ERROR: Could not find the AD user specified"
        Write-Host -ForegroundColor Yellow -BackgroundColor DarkRed "Try using the PRE-2000 logon name instead"
        Pause
        StartMenu
    }

    Write-Host -ForegroundColor Green "Is this the correct user account?"
    Write-Host -ForegroundColor Green "Type 1 and press ENTER to continue and disable the account"
    Write-Host -ForegroundColor Green "Type 2 and press ENTER to Start Over" 
    Write-Host -ForegroundColor Green "Type 3 and press ENTER to Quit"   
    
    # use the switch statement to choose an option
    switch (Read-Host) {
        1 { WriteGroupLog -UPN $UPN -Wnummer $Wnummer }
        2 { StartMenu }
        3 { exit }
    default { Write-Host "Error: You did not choose a valid option, restarting..." ; Start-Sleep -Seconds 1  ; StartMenu }
    }

}

# Function to write a log of the current ad groups before removing them
function WriteGroupLog {
    param (
        [parameter(Mandatory=$true)]
        [string]$UPN,
        [parameter(Mandatory=$true)]
        [string]$Wnummer
    )

    $LogPath = "C:\Beheer\Scripts\Offboarding\GroupLog\$UPN,$(Get-Date -format dd-MM-yyyy).txt"

    Write-Output "========= START OF NEW ENTRY =========" | Out-File $LogPath -Append

    Write-Output "On $(Get-Date) $UPN has been disabled." | Out-File $LogPath -Append

    Get-ADPrincipalGroupMembership -Identity $UPN | Select-Object name | Out-File $LogPath -Append

    Start-Transcript -Path $LogPath -Append

    Write-Host "`n$(Get-Date)"
    Write-Host "User disabled: $UPN"
    Write-Host "Change number: $Wnummer"
    
    Stop-Transcript

    DisableUser -UPN $UPN -Wnummer $Wnummer

}


# Function that disables the user account
function DisableUser {
    param (
        [parameter(Mandatory=$true)]
        [string]$UPN,
        [parameter(Mandatory=$true)]
        [string]$Wnummer
    )

    Clear-Host
    # this will disable the user, remove groups and fill in description
    Disable-ADaccount -identity $UPN
    Write-host -ForegroundColor Green "The account $UPN has been disabled"

    $date = Get-Date -format dd-MM-yyyy
    Set-ADUser -Identity $UPN -Description "$Wnummer, $date"
    Write-host -ForegroundColor Green "The description has been changed to $date, $Wnummer"

    # remove ad groups
    $Groups = (Get-ADUser -Identity $UPN -Properties memberOf).memberOf
    ForEach ($Group In $Groups) {
        Remove-ADGroupMember -Identity $Group -Members $UPN -Confirm:$False
    }

    Write-host -ForegroundColor Green "The AD groups have been removed"

    MakeOnPremShared -UPN $UPN -Wnummer $Wnummer
    
}


# Function to make on prem mailbox shared
function MakeOnPremShared {
    param (
        [parameter(Mandatory=$true)]
        [string]$UPN,
        [parameter(Mandatory=$true)]
        [string]$Wnummer
    )

    Clear-Host
    Write-host -ForegroundColor Green "Connecting to on-prem mailserver"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://FQDN/PowerShell/ -Authentication Kerberos 
    Import-PSSession $Session -DisableNameChecking

    Set-RemoteMailbox -Identity "$UPN" -Type Shared

    Set-RemoteMailbox -Identity "$UPN" -HiddenFromAddressListsEnabled $true

    Write-host -ForegroundColor Green "The on-prem remote mailbox has been made shared and hidden from address lists"

    Pause

}


# =============================================================================================================================================
# Main Code
# =============================================================================================================================================


function StartMenu {
    Clear-Host
    Write-host -ForegroundColor Green "============================================="
    Write-host -ForegroundColor Green "This script will do the following"
    Write-host -ForegroundColor Green "Disable the AD account and set description"
    Write-host -ForegroundColor Green "Remove ALL AD groups from the user account"
    Write-host -ForegroundColor Green "Set the on-prem and 365 mailbox shared"
    Write-host -ForegroundColor Green "============================================="
    Write-host -ForegroundColor Green "`nLets disable the account"
    
    $UPN = Read-Host -Prompt "Enter the username"
    Write-Host "`n(Example W2305 323)"
    $Wnummer = Read-Host -Prompt "Enter the Topdesk change number"

# Prompt user to validate
    Write-host "`n======================================================="
    Write-host -ForegroundColor Green " !!! Item Validation !!!"
    Write-host "======================================================="
    Write-Host "Username       : $UPN"
    Write-Host "Topdesk change : $Wnummer"

    Write-Host -ForegroundColor Green "Is this correct?"
    Write-Host -ForegroundColor Green "Type 1 and press ENTER to continue"
    Write-Host -ForegroundColor Green "Type 2 and press ENTER to start over" 
    Write-Host -ForegroundColor Green "Type 3 and press ENTER to Quit"         


switch (Read-Host) {
    1 { FindADUser -UPN $UPN -Wnummer $Wnummer }
    2 { StartMenu }
    3 { exit }
    default { Write-Host "Error: You did not choose a valid option, restarting..." ; Start-Sleep -Seconds 1  ; StartMenu }
}

}


# Invoking the start menu (first step in the script)
StartMenu


