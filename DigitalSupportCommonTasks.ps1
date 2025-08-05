function Show-Menu {
    # Define the menu items for each column
    $column1Items = @(
        "ENTRA ID (AAD)"
        ""
        "  1.  ACTION: Activate Entra PIM Role(s)"
        "  2.  REPORT: M365 User Accounts"
        "  3.  REPORT: M365 Guest Accounts"
        "  4.  REPORT: Nested Security Groups"
        "  5.  ACTION: Check User AHO Assignment"
        "  6.  ACTION: Update User UPN and Email"
        ""
        "EXCHANGE ONLINE (EXO)"
        ""
        "  7.  ACTION: Create a new shared mailbox" 
        "  8.  REPORT: User Mailbox Access" 
        "  9.  ACTION: Modify calendar permission/delegate access"
        "  10. ACTION: Update approved senders on distribution lists"
        "  11. ACTION: Remove calendar Events for user"
        "  12. REPORT: Distribution Group Members"
        ""
        "SHAREPOINT ONLINE (SPO)"
        ""
        "  13. REPORT: List Item Report"
        "  14. REPORT: Site Report"
        "  15. ACTION: Move between sites/libraries"
        "  16. ACTION: Bulk Delete Files"
        "  17. REPORT: Site / Library / Folder Report"
        ""
    )
    $column2Items = @(
        "ONE DRIVE (OD)"
        ""
        "  18. REPORT: User File Types Report"
        "  19. REPORT: Bulk Delete OD Files"
        ""
        "INTUNE (MAM)"
        ""
        "  20. REPORT: Get All Apps and Group Assignments"
        "  21. REPORT: Generate All Discovered Apps"
        ""
        "TEAMS (MS TEAMS)"
        ""
        "  22. REPORT: Get All Teams Owner and Members"
        "  23. REPORT: Get Users Teams Access"
        ""
        "COMPLIANCE (PURVIEW)"
        ""
        "  24. REPORT: User Activity"
        "  25. REPORT: SPO Activity"
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
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Entra/refs/heads/main/hud.govt.nz/Project/Digital%20Support%20Tool%20Scripts/Admin_PIM_Role_Activation.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
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
    & C:\HUD\00_Staging\Report_Discovered_Apps.ps1 -Platform $Platform
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
    try {
        $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
        $tempFile      = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName() + ".ps1")
        Set-Content -Path $tempFile -Value $scriptContent

        Invoke-Expression "& `"$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe`" -File `"$tempFile`""
    } catch {
        Write-Error "Failed to download or execute the script from $scriptUrl. Error: $_"
    }

}

function Get-SPOActivityAuditReport {
    Clear-Host
    Write-Warning "Minimum Entra PIM role required to run this report: Compliance Administrator"
    start-sleep 3
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/main/Purview/SPO_Activity_Report.ps1"
    try {
        $scriptContent = (New-Object System.Net.WebClient).DownloadString($scriptUrl)
        $tempFile      = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName() + ".ps1")
        Set-Content -Path $tempFile -Value $scriptContent

        Invoke-Expression "& `"$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe`" -File `"$tempFile`""
    } catch {
        Write-Error "Failed to download or execute the script from $scriptUrl. Error: $_"
    }
}


function Generate-SPOSiteHierarchyView {
    Clear-Host
    $scriptUrl = "https://raw.githubusercontent.com/hud-govt-nz/Microsoft-365-and-Azure/d73d38d9092d0fc6e1f0713bf9c5dc35d4502327/_Projects/HUD%20Digital%20Support/Scripts/023_SPO_Site_Hierarchy_View.ps1"
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString($scriptUrl)
}

# Select task based on Show-Menu function
do {
    Clear-Host
    $selection = Show-Menu
    switch ($selection) {
        '1'  { New-PIMSession }
        '2'  { Export-AllUserReport }
        '3'  { Export-AllGuestsReport }
        '4'  { Export-NestedGroupReport }
        '5'  { Get-EmployeeAssignment }
        '6'  { Update-UserNameAndEmail }
        '7'  { New-SharedMailbox }
        '8'  { Get-UserMailboxAccess }
        '9'  { Get-DelegateAccess }
        '10' { Edit-DLApprovedSenders }
        '11' { Remove-CalendarEventsForUser }
        '12' { Get-DLGroupMember }
        '13' { Get-SPOListItemReport }
        '14' { Get-SPOBasicSiteReport }
        '15' { Move-SPOFolders }
        '16' { Remove-SPOItems }
        '17' { Generate-SPOSiteHierarchyView }
        '18' { Get-ODFileTypes }
        '19' { Remove-ODItems -IncludeOneDriveSites }
        '20' { Get-AppAssignments }
        '21' { Get-DiscoveredApps }
        '22' { Get-AllTeamMembersAndOwners }
        '23' { Get-TeamAccessReportForUser }
        '24' { Get-UserActivityAuditReport }
        '25' { Get-SPOActivityAuditReport }

        'q'  { return }
    }
    pause
} until ($selection -eq 'q')