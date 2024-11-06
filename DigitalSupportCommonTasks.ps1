function Show-Menu {
    # Define the menu items for each column
    $column1Items = @(
        "ENTRA ID (AAD)"
        ""
        "  1.  ACTION: Activate Entra PIM Role(s)"
        "  2.  REPORT: M365 User Accounts"
        "  24.  REPORT: M365 Guest Accounts"
        "  3.  REPORT: Nested Security Groups"
        "  4.  ACTION: Check User AHO Assignment"
        "  5.  ACTION: Update User UPN and Email"
        ""
        "EXCHANGE ONLINE (EXO)"
        ""
        "  6.  ACTION: Create a new shared mailbox" 
        "  7.  REPORT: User Mailbox Access" 
        "  8.  ACTION: Modify calendar permission/delegate access"
        "  9.  ACTION: Update approved senders on distribution lists"
        "  10. ACTION: Remove calendar Events for user"
        "  11. REPORT: Distribution Group Members"
        ""
        "SHAREPOINT ONLINE (SPO)"
        ""
        "  12. REPORT: List Item Report"
        "  13. REPORT: Site Report"
        "  14. ACTION: Move between sites/libraries"
        "  15. ACTION: Bulk Delete Files"
        ""
    )
    $column2Items = @(
        "ONE DRIVE (OD)"
        ""
        "  16. REPORT: User File Types Report"
        "  17. REPORT: Bulk Delete OD Files"
        ""
        "INTUNE (MAM)"
        ""
        "  18. REPORT: Get All Apps and Group Assignments"
        "  19. REPORT: Generate All Discovered Apps"
        ""
        "TEAMS (MS TEAMS)"
        ""
        "  20. REPORT: Get All Teams Owner and Members"
        "  21. REPORT: Get Users Teams Access"
        ""
        "COMPLIANCE (PURVIEW)"
        ""
        "  22. REPORT: User Activity"
        "  23. REPORT: SPO Activity"
    )
    # Define a fixed width for the first column, enough to accommodate the longest line
    $column1Width = 60

    # Write the header
    Write-Host "## HUD Digital Support Tool ##" -ForegroundColor Green
    Write-Host ""
    Write-Warning "Ensure you have the right permissions to run these commands"
    Write-Host ""

    # Print the menu items side by side
    for ($i = 0; $i -lt [Math]::Max($column1Items.Length, $column2Items.Length); $i++) {
        $column1Text = $column1Items[$i] -replace "`t", "    " # replace tabs with spaces if needed
        $column2Text = $column2Items[$i] -replace "`t", "    " # replace tabs with spaces if needed

        # Check if we have an item for the current index in each column
        if ($null -eq $column1Text) { $column1Text = "" }
        if ($null -eq $column2Text) { $column2Text = "" }

        # Determine if the item is a heading
        $isColumn1Heading = $column1Text -match "^\D+$" # Matches text that doesn't contain numbers
        $isColumn2Heading = $column2Text -match "^\D+$" # Matches text that doesn't contain numbers

        # Print the items with padding to align the columns
        $formattedColumn1Text = $column1Text.PadRight($column1Width)
        
        # Apply color to headings
        if ($isColumn1Heading) {
            Write-Host $formattedColumn1Text -NoNewline -ForegroundColor Cyan
        } else {
            Write-Host $formattedColumn1Text -NoNewline
        }
        
        if ($isColumn2Heading) {
            Write-Host $column2Text -ForegroundColor Cyan
        } else {
            Write-Host $column2Text
        }
    }
    Write-Host ""
    $option = Read-Host "Enter your number choice (or Q to exit)"
    return $option
}

# Define the functions
function New-PIMSession {
    Clear-Host
    Write-Warning "Please ensure you have been granted access to your selected role before attempting to activate"
    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Entra/Enable_PIM_Assignments.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

# Ensure the script is running with PowerShell 5
if ($PSVersionTable.PSVersion.Major -ne 5) {
    Write-Error "This script requires PowerShell version 5."
    exit
}

function Export-AllUserReport {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/001_ENTRA_User_Report.ps1"

    try {
        $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
        $tempFile      = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName() + ".ps1")
        Set-Content -Path $tempFile -Value $scriptContent

        Invoke-Expression "& `"$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe`" -File `"$tempFile`""
    } catch {
        Write-Error "Failed to download or execute the script from $scriptUrl. Error: $_"
    }
}

function Export-NestedGroupReport {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/002_ENTRA_Nested_Group_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Get-EmployeeAssignment {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/003_ENTRA_Update_Aho_Assignment.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Update-UserNameAndEmail {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/004_ENTRA_Update_Username.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Export-AllGuestsReport {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/022_ENTRA_Guest_Account_Report.ps1"
    try {
        $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
        $tempFile      = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName() + ".ps1")
        Set-Content -Path $tempFile -Value $scriptContent

        Invoke-Expression "& `"$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe`" -File `"$tempFile`""
    } catch {
        Write-Error "Failed to download or execute the script from $scriptUrl. Error: $_"
    }
}

function New-SharedMailbox {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/005_EXO_New_Shared_Mailbox.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Get-UserMailboxAccess {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/006_EXO_User_Access_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Get-DelegateAccess {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/018_EXO_Update_Calendar_Delegates.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Edit-DLApprovedSenders {
    Clear-Host
    $DL     = Read-Host "Please Enter Distribution List Name"
    $Action = Read-Host "Please specify 'Add', 'Remove', 'Review' [Default is 'Review']"
    if (-not $Action) {
        $Action = 'Review'
    }   
    $scriptUrl     = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/019_EXO_Update_DL_Approved_Senders.ps1"
    $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    $scriptBlock   = [scriptblock]::Create($scriptContent)
    Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $DL, $Action
}

function Remove-CalendarEventsForUser {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/020_EXO_Remove_Calendar_Events.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)   
}

function Get-DLGroupMember {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/021_EXO_DL_Member_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
}

function Get-SPOListItemReport {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/008_SPO_List_Item_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
}

function Get-SPOBasicSiteReport {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/007_SPO_Basic_Site_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
}

function Move-SPOFolders {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/009_SPO_Move_Folders_Between_Sites.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
}

function Remove-SPOItems {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/010_SPO_Bulk_Delete_Files.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
}

function Get-ODFileTypes {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/017_OD_User_File_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)  
}

function Remove-ODItems {
    param (
        [switch]$IncludeOneDriveSites
    )
    Clear-Host
    $scriptUrl     = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/010_SPO_Bulk_Delete_Files.ps1"
    $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    $scriptBlock   = [scriptblock]::Create($scriptContent)
    Invoke-Command -ScriptBlock $scriptBlock -ArgumentList $IncludeOneDriveSites
}

function Get-AppAssignments {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/011_INTUNE_App_Assignment_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Get-DiscoveredApps {
    Clear-Host
    $scriptUrl     = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/012_INTUNE_App_Discovery_Report.ps1"
    $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
    $scriptContent | Out-File -FilePath "C:\HUD\00_Staging\Report_Discovered_Apps.ps1" -Encoding UTF8
    $Platform = Read-Host "Enter the Platform value (Windows, AndroidWorkProfile, iOS)"
    & C: \HUD\00_Staging\Report_Discovered_Apps.ps1 -Platform $Platform
}

function Get-AllTeamMembersAndOwners {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/013_TEAMS_All_Team_Owners_And_Members.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Get-TeamAccessReportForUser {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/_Projects/HUD%20Digital%20Support/Scripts/014_TEAMS_User_Access_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Get-UserActivityAuditReport {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Compliance Administrator"
    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Purview/Audit_User_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

function Get-SPOActivityAuditReport {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Compliance Administrator"
    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Purview/SPO_Activity_Report.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

# Select task based on Show-Menu function
do {
    Clear-Host
    $selection = Show-Menu
    switch ($selection) {
        '1'  { New-PIMSession }
        '2'  { Export-AllUserReport }
        '3'  { Export-NestedGroupReport }
        '4'  { Get-EmployeeAssignment }
        '5'  { Update-UserNameAndEmail }
        '6'  { New-SharedMailbox }
        '7'  { Get-UserMailboxAccess }
        '8'  { Get-DelegateAccess }
        '9'  { Edit-DLApprovedSenders }
        '10' { Remove-CalendarEventsForUser }
        '11' { Get-DLGroupMember }
        '12' { Get-SPOListItemReport }
        '13' { Get-SPOBasicSiteReport }
        '14' { Move-SPOFolders }
        '15' { Remove-SPOItems }
        '16' { Get-ODFileTypes }
        '17' { Remove-ODItems -IncludeOneDriveSites}
        '18' { Get-AppAssignments }
        '19' { Get-DiscoveredApps }
        '20' { Get-AllTeamMembersAndOwners }
        '21' { Get-TeamAccessReportForUser }
        '22' { Get-UserActivityAuditReport }
        '23' { Get-SPOActivityAuditReport }
        '24' { Export-AllGuestsReport }
        'q'  { return }
    }
    pause
} until ($selection -eq 'q')