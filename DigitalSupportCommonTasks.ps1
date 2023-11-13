function Show-Menu {
    Write-Host ""
    Write-Warning "Ensure you have the right permissions to run these commands"
    Write-Host ""
    Write-Host "  Entra (formally Azure AD)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    1.  Activate Entra PIM Role(s)" #
    Write-Host "    2.  Entra All Users report" #
    Write-Host "    3.  Entra Nested Security Groups report" #
    Write-Host "    4.  Check User Aho Assignment Status" #
    Write-Host "    5.  Change Username and Email Address of User" #
    Write-Host "    6.  Add/Remove domain from external access (TBC)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Exchange Online" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    7.  Create a new shared mailbox" 
    Write-Host "    8.  Get mailbox access report for user" 
    Write-Host "    9.  Set calendar delegate access for user (TBC)" -ForegroundColor Red
    Write-Host "    10. Update approved senders on distribution lists"
    Write-Host "    11. Remove calendar Events for user"
    Write-Host "    12. Distribution list members report"
    Write-Host ""
    Write-Host "  SharePoint Online" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    13. Get List Item Report"
    Write-Host "    14. Generate Basic Site Report"
    Write-Host "    15. Move files between (TBC)" -ForegroundColor Red
    Write-Host "" 
    Write-Host "  Intune Reporting" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    16. Get All Apps and Group Assignments"
    Write-Host "    17. Generate All Discovered Apps Report"

    Write-Host ""    

    $option = Read-Host "Enter your number choice (or Q to exit)"
    return $option
}
function AppAssignments {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Intune Administrator"

    start-sleep 3

    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Intune/Reporting/Report_App_Assignments.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}
function DiscoveredApps {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Intune Administrator"

    start-sleep 3

    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Intune/Reporting/Report_Discovered_Apps.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}
function New-PIMSession {
    Clear-Host
    Write-Warning "Please ensure you have been granted access to your selected role before attempting to activate"

    start-sleep 3

    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Enable_PIM_Assignments.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}
function Export-AllUserReport {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: User Administrator"

    start-sleep 3
  
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Reports/Entra_All_Users.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}
function Export-NestedGroupReport {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: User Administrator & Exchange Recipient Administrator"

    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Reports/Entra_Nested_Security_Group_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}
function Get-EmployeeAssignment {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: User Administrator"

    start-sleep 3
  
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Get-AhoEmployeeAssignment.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}
function Update-AllowedDomains {
    Write-Host ""
    Write-host "This function is a work in progress" -ForegroundColor Red
    Write-Host ""
}
function Update-UserNameAndEmail {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: User Administrator & Exchange Recipient Administrator"

    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Update-Username.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)

}
function New-SharedMailbox {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: User Administrator & Exchange Recipient Administrator"

    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/New-SharedMailbox.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}
function Get-UserMailboxAccess {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"

    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Check_Mailbox_Access.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}
function Update-CalendarDelegate {
    Write-Host ""
    Write-host "This function is a work in progress" -ForegroundColor Red
    Write-Host ""
}
function Edit-DLApprovedSenders {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"

    start-sleep 3

    $DL = Read-Host "Please Enter Distribution List Name"

    $Action = Read-Host "Please specify 'Add', 'Remove', 'Review' [Default is 'Review']"
    
    if (-not $Action) {
        $Action = 'Review'
    }   
   
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Update_DL_Approved_Senders.ps1"
    $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    $scriptBlock = [scriptblock]::Create($scriptContent)
    Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $DL, $Action

}
function Remove-CalendarEventsForUser {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"

    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Remove-CalenderEvents.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)   
}
function Get-DLGroupMember {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Exchange Recipient Administrator"

    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Exchange/Reports/Distribution_Group_Member_Export.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
}

function Get-SPOListItemReport {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: SharePoint Administrator"
    Write-Warning "Alternatively you must have full permissions to the respective site to run this"

    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/SharePoint%20Online/Get-SPListItemReport.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  

}

function Get-SPOBasicSiteReport {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: SharePoint Administrator"
    Write-Warning "Alternatively you must have full permissions to the respective site to run this"

    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/SharePoint%20Online/Get-BasicSiteReport.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
    
}
#Select task based on Show-Menu function
do
{
    Clear-Host
    Write-Host ""
    Write-Host '## HUD Digital Support Tool ##' -ForegroundColor Yellow
    $selection = Show-Menu
    switch ($selection) {
                 '1'  {New-PIMSession}
                 '2'  {Export-AllUserReport}
                 '3'  {Export-NestedGroupReport}
                 '4'  {Get-EmployeeAssignment}
                 '5'  {Update-UserNameAndEmail}
                 '6'  {Update-AllowedDomains}
                 '7'  {New-SharedMailbox}
                 '8'  {Get-UserMailboxAccess}
                 '9'  {Update-CalendarDelegate}
                 '10' {Edit-DLApprovedSenders}
                 '11' {Remove-CalendarEventsForUser}
                 '12' {Get-DLGroupMember}
                 '13' {Get-SPOListItemReport}
                 '14' {Get-SPOBasicSiteReport}
                 '15' {}
                 '16' {AppAssignments}
                 '17' {DiscoveredApps}
                 'q'  {return}
                 }
        pause
        }
until ($selection -eq 'q')


<#
 #New-HUDUser
 #New-SharedMailbox
 Search-MailboxAccess
 Add-PhoneNumber

 #>