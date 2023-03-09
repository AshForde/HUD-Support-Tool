function Show-Menu {
    Write-Host "1. Elevate PIM roles"
    Write-Host "2. Create a new user in Azure AD"
    Write-Host "3. Create a new shared mailbox in Exchange Online"
    Write-Host "4. Check users mailbox access in Exchange Online"
    Write-Host "5. Add a DDI Phone Number to a user"
    Write-Host "6. Add/Change Aho 'Employee Category' to existing user"
    Write-Host "7. Change users email address in Exchange Online"
    Write-Host "8. Run Basic Reports"
    Write-Host "Q. Exit"
    Write-Host ""
    $option = Read-Host "Enter your choice (1-8 or Q to exit)"
    return $option
}
# 1: PIM elevate
function New-PIMSession {
    Clear-Host
    Write-Host ''
    Write-Host '## Elevate PIM roles ##' -ForegroundColor Yellow
    Write-Host '1. Global Administrator' -ForegroundColor Green
    Write-Host '2. Cloud Operations' -ForegroundColor Green
    Write-Host '3. Digital Workplace Support' -ForegroundColor Green
    Write-Host '4. Sharepoint Support' -ForegroundColor Green

    $PIM = Read-Host 'Please select your role based on the options Above (1-4)'
    
    # Connect to Azure AD
    $upn = whoami /upn
    Connect-AzureAD -AccountId $upn

    #PIM selection
    Switch ($PIM) {
        "1" {       
            $Reason = Read-host "Please provide a reason for activating the Global Administrator role"
            Enable-DCAzureADPIMRole -RolesToActivate 'Global Administrator' `
            -UseMaximumTimeAllowed -Reason $Reason}
        "2"{
            $Reason = Read-host "Please provide a reason for activating your Cloud Operations roles"
            Enable-DCAzureADPIMRole `
            -RolesToActivate `
                'Billing Administrator',`
                'Exchange Administrator',`
                'Global Reader',`
                'Groups Administrator',`
                'Intune Administrator',`
                'Security Reader',`
                'SharePoint Administrator',`
                'Teams Administrator',`
                'User Administrator' -UseMaximumTimeAllowed -Reason $Reason
        }
        "3"{            
            $Reason = Read-host "Please provide a reason for activating your Digital Workplace Support roles"
            Enable-DCAzureADPIMRole `
            -RolesToActivate `
                'Intune Administrator',`
                'HelpDesk Administrator',`
                'User Administrator',`
                'Teams Communications Administrator',`
                'Exchange Recipient Administrator',`
                'Global Reader',`
                'Security Reader' `
            -UseMaximumTimeAllowed -Reason $Reason
        }
        "4"{
            $Reason = Read-host "Please provide a reason for activating your Sharepoint Support roles"
            Enable-DCAzureADPIMRole -RolesToActivate 'SharePoint Administrator','Teams Administrator', 'User Administrator' `
            -UseMaximumTimeAllowed -Reason $Reason
        }
    }
}
# 2: Create new user
function New-HUDUser {
    Clear-Host
    Write-Host ''
    Write-Host '## Create a new user in Azure AD ##' -ForegroundColor Yellow
    # Connect to MgGraph and define scope for user account creation
    Connect-MgGraph -Scopes "Directory.Read.All", "Directory.ReadWrite.All", "User.Read.All", "User.ReadWrite.All" | Out-Null

    # Set Graph Schema (v1.0 or Beta)
    Select-MgProfile -Name "v1.0" 

    # Gather Details
    $givenName = Read-Host "Please enter the users first name"
    $surname = Read-Host "Please enter the users last name"
    $department = Read-Host "Please enter the users department"
    $jobTitle = Read-Host "Please enter the users job title"
    $manager = Read-Host "Please enter the users manager (firstname.lastname@hud.govt.nz)"

    #Employee Status
    Write-Host ''
    $EmpCategory = Read-Host "Enter an employee category (Permanent, Contractor, Vendor, FixedTerm, Secondment)"
        Switch ($EmpCategory) {
            "Permanent" {$EmployeeCategory = "HUD_SUBSTANTIVE_POSITION"}
            "Contractor"{$EmployeeCategory = "ORA_HRX_CONTRACTOR"}
            "Vendor"{$EmployeeCategory = "ORA_HRX_CONSULTANT"}
            "FixedTerm"{$EmployeeCategory = "HUD_FIX_TERM"}
            "Secondment"{$EmployeeCategory = "HUD_EXTERNAL_SECONDMENT"}
        }

    # Start Date
    Write-Host ''
    $startDate = Read-Host "Please enter the users start date (dd/MM/yyyy)"
    
    # Location info
    Write-Host ''
    Write-Host "Please select the location of the user" -ForegroundColor Green
    Write-Host ""
    Write-Host "    1. Level 6,7 Waterloo Quay,Pipitea,Wellington"
    Write-Host "    2. Level 7,7 Waterloo Quay,Pipitea,Wellington"
    Write-Host "    3. Level 8,7 Waterloo Quay,Pipitea,Wellington"
    Write-Host "    4. Level 9,7 Waterloo Quay,Pipitea,Wellington"
    Write-Host "    5. APO, Level 6, 45 Queen Street, Auckland"
    Write-Host "    6. APO, Level 7, 45 Queen Street, Auckland"
    Write-Host ""
    $Location = Read-Host "Location Choice"

    Switch ($Location) {
            "1" {
                $StreetAddress = "Level 6,7 Waterloo Quay,Pipitea,Wellington"
                $City = "Wellington"
                $PostalCode = "6011"
                $OfficeLocation = "Wellington - 7WQ - Level 6"
            }
            "2" {
                $StreetAddress = "Level 6,7 Waterloo Quay,Pipitea,Wellington"
                $City = "Wellington"
                $PostalCode = "6011"
                $OfficeLocation = "Wellington - 7WQ - Level 7"
            }
            "3" {
                $StreetAddress = "Level 6,7 Waterloo Quay,Pipitea,Wellington"
                $City = "Wellington"
                $PostalCode = "6011"
                $OfficeLocation = "Wellington - 7WQ - Level 8"
            }
            "4" {
                $StreetAddress = "Level 6,7 Waterloo Quay,Pipitea,Wellington"
                $City = "Wellington"
                $PostalCode = "6011"
                $OfficeLocation = "Wellington - 7WQ - Level 9"
            }
            "5" {
                $StreetAddress = "APO, Level 6, 45 Queen Street, Auckland"
                $City = "Auckland"
                $PostalCode = "1010"
                $OfficeLocation = "Auckland - APO - Level 6"
            }
            "6" {
                $StreetAddress = "APO, Level 7, 45 Queen Street, Auckland"
                $City = "Auckland"
                $PostalCode = "1010"
                $OfficeLocation = "Auckland - APO - Level 6"
            }
        }

    # formatting
    $displayName = "$givenName $surname"
    $mailNickName = "$GivenName$Surname"
    $UserPrincipalName = "$givenName.$Surname" + "@hud.govt.nz"

    # Password Profile
    Add-Type -AssemblyName 'System.Web'
    $NewPassword = [System.Web.Security.Membership]::GeneratePassword(10, 3)
    $NewPasswordProfile = @{}
    $NewPasswordProfile["Password"]= $NewPassword
    $NewPasswordProfile["ForceChangePasswordNextSignIn"] = $True

    # Wrap output into parameters table
    $NewUserParams = @{
        'DisplayName' = $displayName
        'PasswordProfile' = $NewPasswordProfile
        'AccountEnabled' = $true
        'MailNickName' = $mailNickName
        'UserPrincipalName' = $UserPrincipalName
        'GivenName' = $givenName
        'Surname' = $surname
        'CompanyName' = "Ministry of Housing and Urban Development"
        'Department' = $department
        'JobTitle' = $jobTitle
        'StreetAddress' = $StreetAddress
        'state' = "NZ"
        'City' = $City
        'Country' = "New Zealand"
        'PostalCode' = $PostalCode
        'OfficeLocation' = $OfficeLocation
        'preferredLanguage' = "en-NZ"

    }

    # Create New User Object
    New-MgUser -BodyParameter $NewUserParams
    $id = (Get-MgUser -UserId $UserPrincipalName).Id

    # Assign Custom Attributes
    $Attributes =@{
        'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $EmployeeCategory
        'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate'  = $startDate
    }
    Update-MgUser -UserId $UserPrincipalName -AdditionalProperties $Attributes

    # Assign Manager
    $Manager = @{
        '@odata.id' = "https://graph.microsoft.com/v1.0/users/$manager"
    } 
    Set-MgUserManagerByRef -UserId $UserPrincipalName -BodyParameter $Manager

    Write-Host "New User: $($displayName) has been created with the following details" -ForegroundColor Cyan
    Write-Host ''
    Write-Host "Username: $($UserPrincipalName)" -ForegroundColor Green
    Write-Host "Password: $($NewPasswordProfile.Password)" -ForegroundColor Green
    Write-Host "Manager: $manager" -ForegroundColor Green
    Write-Host "Location: $StreetAddress" -ForegroundColor Green
    Write-Host "Employee Category: $EmployeeCategory" -ForegroundColor Green
    Write-Host "Start Date: $startDate" -ForegroundColor Green
    Write-Host "ID: $id" -ForegroundColor Green
    Write-Host ''
    Write-Host 'The users mailbox and standard groups will be applied during the next onboarding automation sync' -ForegroundColor Cyan
}
# 3: Create Shared Mailbox in Exchange Online
function New-SharedMailbox {
    Clear-Host
    Write-Host ''
    Write-Host '## Create Shared Mailbox in Exchange Online ##' -ForegroundColor Yellow

    # Connect to Exchange Online
    $UserPrincipalName = Whoami /UPN
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false
        
    # Create Mailbox
    $SharedMBName = Read-Host "Please enter the name of the mailbox"
    $Alias = $SharedMBName -replace ' ',""
    $EmailAddress = "$Alias" + "@hud.govt.nz"
    $Mailbox = New-mailbox -Shared -Name $SharedMBName -Alias $Alias -PrimarySmtpAddress $EmailAddress

    Write-Host "Shared Mailbox: $(($Mailbox).Name) has been created with the following details" -ForegroundColor Cyan
    Write-Host ''
    Write-Host "Primary Email Address: $(($Mailbox).PrimarySmtpAddress)" -ForegroundColor Green
    Write-Host "Shared Mailbox: $(($Mailbox).IsShared)" -ForegroundColor Green
    Write-Host "ID: $(($Mailbox).Id)" -ForegroundColor Green
    Write-Host "Exchange GUID: $(($Mailbox.ExchangeGuid).Guid)" -ForegroundColor Green
    Write-Host ''

    $title = "Add Users to Mailbox"
    $sendasmsg = "Would you like to add users to this mailbox?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    do {
        $result = $host.ui.PromptForChoice($title, $sendasmsg, $options, 0) 
        switch ($result) {
            '0' {
                $UserMB = Read-Host "Please enter users email address"
                $FullAccess = Add-MailboxPermission -Identity ($Mailbox).Id -User $UserMB -AccessRights FullAccess -InheritanceType All -AutoMapping $true
                $Sendas = Add-RecipientPermission ($Mailbox).Id -AccessRights SendAs -Trustee $UserMB -Confirm:$false
                Write-Host "$($UserMB) has been granted 'Full' and 'Send As' permissions to mailbox $(($Mailbox).Id)" -foregroundcolor cyan
            } 
            '1' {
                break # Exit the loop if user selects "No"
            }
        }
    } while ($result -ne '1')

    # Disconnect Exchange Online Session
    Disconnect-ExchangeOnline -Confirm:$false
}
# 4: Check users mailbox access in Exchange Online
function Search-MailboxAccess {
    Clear-Host
    Write-Host ''
    Write-Host '## Check users mailbox access in Exchange Online ##' -ForegroundColor Yellow

    # Connect to Exchange Online
    $UserPrincipalName = Whoami /UPN
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false
    
    # Obtain User ID
    $UserEmailAddress = Read-Host "Please enter the user's email address of the user you want to check"

    # Get the total number of mailboxes to process
    $totalMailboxes = (Get-Mailbox).Count
    Write-Host "Total number of mailboxes to process: $totalMailboxes"
    Write-Host "Searching through Exchange Online will take some time, please be patient..." -ForegroundColor Cyan

    # Gather Output Array for results
    $Output = @()

    # Use a counter to keep track of the number of mailboxes processed
    $counter = 0

    Get-Mailbox | ForEach-Object {
        # Increment the counter for each mailbox processed
        $counter++
    
        # Calculate the percentage of mailboxes processed
        $percentComplete = ($counter / $totalMailboxes) * 100
        
        # Get the mailbox permissions and filter for the specified user
        $recipientPermissions = $_ | Get-RecipientPermission | Where-Object {($_.Trustee -like "*$UserEmailAddress*") -and ($_.Trustee -notmatch "S-1-5-21") -and ($_.Trustee -notmatch "NT AUTHORITY\SELF")} -ErrorAction SilentlyContinue
        $mailboxPermissions = $_ | Get-MailboxPermission | Where-Object {($_.User -like "*$UserEmailAddress*") -and ($_.User -notmatch "S-1-5-21") -and ($_.User -notlike "NT AUTHORITY\SELF")} -ErrorAction SilentlyContinue
        
        # Add the permissions to the output array
        foreach ($permission in $recipientPermissions) {
            $Output += [PSCustomObject]@{
                Identity = $permission.Identity
                User = $permission.Trustee
                AccessRights = $permission.AccessRights -join ', '    
            }
        }
        foreach ($permission in $mailboxPermissions) {
            $Output += [PSCustomObject]@{
                Identity = $permission.Identity
                User = $permission.User
                AccessRights = $permission.AccessRights -join ', '
            }
        }
        # Create the progress bar
        Write-Progress -Activity "Processing mailboxes..." -PercentComplete $percentComplete -Status "Processing mailbox $counter of $totalMailboxes"    
    }

    # Clear the progress bar
    Write-Progress -Activity "Processing mailboxes..." -Completed

    # Display the list of mailboxes the user has access to
    if ($Output){
        Clear-Host
        Write-Host "The user $($UserEmailAddress) has access to the following mailboxes:" -ForegroundColor Green
        $Output | Format-Table -AutoSize -Wrap
    } else {
        Write-Host "The user $($UserEmailAddress) does not have access to any mailboxes." -ForegroundColor Green
    }
    
}

# 5. Add a DDI Phone Number to a user
function Add-PhoneNumber{
    Clear-Host
    Write-Host ''
    Write-Host '## Add a DDI Phone Number to a user ##' -ForegroundColor Yellow

    # Connect to MgGraph and Microsoft Teams
    Connect-MgGraph -Scopes "Directory.Read.All", "Directory.ReadWrite.All", "User.Read.All", "User.ReadWrite.All" | Out-Null
    Connect-MicrosoftTeams
    
    # Obtain User ID
    $User = Read-Host "Enter the User Principal Name of the user"
    $DDI = Read-Host "Enter the DDI phone number to be assigned. Copy/Paste from DDI Master List or use format +64 X XXX XXXX"   

    # Assign Number in Azure AD
    Update-MgUser -UserId $User -BusinessPhones $DDI

    # Obtain users location
    $Location = (Get-mguser -UserId $User).OfficeLocation
    $DisplayName = (Get-mguser -UserId $User).DisplayName
    $UserID = (Get-mguser -UserId $User).id
    
    #Format DDI for Teams application "+64XXXXXXXX"
    $DDI = [string]$DDI -replace " ",''

    #Assign number in Microsoft Teams 
    Set-CsPhoneNumberAssignment -Identity $UserID -PhoneNumber $DDI -PhoneNumberType DirectRouting    
    #Set-CsOnlineVoicemailUserSettings -Identity $UserID -VoicemailEnabled $true       

    if ($Location -match 'Wellington') {
        Grant-CsTenantDialPlan -Identity $UserID  -PolicyName "DP-04Region"
        Grant-CsOnlineVoiceRoutingPolicy -Identity $UserID  -PolicyName Tag:VP-Unrestricted
    } else {
        Grant-CsTenantDialPlan -Identity $UserID -PolicyName "DP-09Region"
        Grant-CsOnlineVoiceRoutingPolicy -Identity $UserID -PolicyName Tag:VP-Unrestricted
    }

    # Result
    Write-Host "$($DisplayName) has access to DDI: $($DDI) " -ForegroundColor Green
}

# 6. Add Aho 'Employee Category' to existing user
function Add-EmployeeCategory {
    Clear-Host
    Write-Host ''
    Write-Host '## Create a new user in Azure AD ##' -ForegroundColor Yellow
    
    # Connect to MgGraph and define scope for user account creation
    Connect-MgGraph -Scopes "Directory.Read.All", "Directory.ReadWrite.All", "User.Read.All", "User.ReadWrite.All" | Out-Null
    
    # Set Graph Schema (v1.0 or Beta)
    Select-MgProfile -Name "v1.0" 
    
    #Employee Status
    Write-Host ''
    $UserPrincipalName = Read-Host "Enter the User Principal Name of the user"

    $EmpCategory = Read-Host "Enter an employee category (Permanent, Contractor, Vendor, FixedTerm, Secondment)"
    Switch ($EmpCategory) {
        "Permanent" {$EmployeeCategory = "HUD_SUBSTANTIVE_POSITION"}
        "Contractor"{$EmployeeCategory = "ORA_HRX_CONTRACTOR"}
        "Vendor"{$EmployeeCategory = "ORA_HRX_CONSULTANT"}
        "FixedTerm"{$EmployeeCategory = "HUD_FIX_TERM"}
        "Secondment"{$EmployeeCategory = "HUD_EXTERNAL_SECONDMENT"}
    }
    
    # Start Date
    Write-Host ''
    $startDate = Read-Host "Please enter the users start date (dd/MM/yyyy). Hit enter if you want to leave blank."
    
    # Assign Custom Attributes
    $Attributes =@{
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $EmployeeCategory
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate'  = $startDate
    }
    if ($startDate -eq $null){
        $Attributes =@{
            'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $EmployeeCategory
            }
        Update-MgUser -UserId $UserPrincipalName -AdditionalProperties $Attributes
    } else {
        $Attributes =@{
            'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $EmployeeCategory
            'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate'  = $startDate
            }
        Update-MgUser -UserId $UserPrincipalName -AdditionalProperties $Attributes
    }

    # Obtain users Object ID
    $id = (Get-MgUser -UserId $UserPrincipalName).Id
    
    # Output
    Write-Host "Employee: $($UserPrincipalName) has been assigned the following Employee Category and Start Date" -ForegroundColor Cyan
    Write-Host "Employee Category: $EmployeeCategory" -ForegroundColor Green
    Write-Host "Start Date: $startDate" -ForegroundColor Green
    Write-Host "ID: $id" -ForegroundColor Green
    Write-Host ''

}

# 7. change users name and email in Azure AD and Exchange Online
function Update-UserName {
    Clear-Host
    Write-Host ''
    Write-Host '## change users email address in Exchange Online ##' -ForegroundColor Yellow

    # Connect to MgGraph and define scope for user account modification
    Connect-MgGraph -Scopes "Directory.Read.All", "Directory.ReadWrite.All", "User.Read.All", "User.ReadWrite.All" | Out-Null

    # Connect to Exchange Online
    $UserPrincipalName = Whoami /UPN
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false
    
    # Obtain User ID
    $UserEmailAddress = Read-Host "Please enter the users current UserPrincipalName or email"
    $User = (Get-MgUser -UserId $UserEmailAddress).DisplayName

    # Cofirm change required
    $title = "Update User Email"
    $sendasmsg = "Would you like to update the email and display name address for $($User)?"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.ui.PromptForChoice($title, $sendasmsg, $options, 0) 
        switch ($result) {
            '0' {
                Write-Host ''
                Write-Host "Provide new details below" -ForegroundColor Yellow
                Write-Host ''
                $givenName = Read-Host "Enter first name"
                $surname = Read-Host "Enter last name"

                $displayName = "$givenName $surname"
                $mailNickName = "$givenName$surname"
                $mailAlias = "$givenName.$surname"
                $NewUserPrincipalName = "$givenName.$surname" + "@hud.govt.nz"

                # Update Azure AD details
                Update-MgUser -UserId $UserEmailAddress -GivenName $givenName -Surname $surname -UserPrincipalName $NewUserPrincipalName -DisplayName $displayName -MailNickname $mailNickName

                # Update Exchange Online Mailbox Primary SMTP
                start-sleep 5
                Set-Mailbox -Identity $NewUserPrincipalName -EmailAddresses "SMTP:$($NewUserPrincipalName)","smtp:$($UserEmailAddress)" -Name $displayName -Alias $mailAlias -WindowsEmailAddress $NewUserPrincipalName

                Write-Host "Primary mail address $($UserEmailAddress) has been updated to $($NewUserPrincipalName)" -foregroundcolor cyan
            } 
            '1' {
                break # Exit the loop if user selects "No"
            }
        }

}

# 8. Run Basic Report Extracts
function Export-Reports {
    # Location info
    Clear-host
    Write-Host ''
    Write-Host "Please select report" -ForegroundColor Green
    Write-Host ""
    Write-Host "    1. Export All Azure AD users"
    Write-Host "    2. Export DL - ALL STAFF distribution members"
    Write-Host "    3. Export DL - Wellington distribution members"
    Write-Host "    4. Export DL - Auckland HUD Only distribution members"
    Write-Host "    5. Select distribution list to export"

    Write-Host ""
    $Location = Read-Host "Please select Report"

    Switch ($Location) {
            "1" {
                Connect-MgGraph -Scopes "Directory.Read.All", "Directory.ReadWrite.All", "User.Read.All", "User.ReadWrite.All","AuditLog.Read.All", "Organization.Read.All" | Out-Null

                # Gather Users IDs and export details into custom table
                $results = Get-MgUser -All | Select-Object -Property `
                  Id, CreatedDateTime, AccountEnabled, UserType, DisplayName, GivenName, Surname, UserPrincipalName, Mail, UsageLocation,
                  Department, JobTitle, CompanyName, StreetAddress, City, PostalCode, State, Country, SecurityIdentifier, MobilePhone, 
                  @{Name='BusinessPhone';Expression={[string]$_.BusinessPhones -replace "{",'' }},
                  @{Name='passwordPolicies';Expression={[string]$_.passwordPolicies}},
                  @{Name='LastSignInDateTime';Expression={($_.SignInActivity | Select-Object -ExpandProperty SignInActivity).LastSignInDateTime}},
                  @{Name='StartDate';Expression={$_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate']}},
                  @{Name='EmployeeCategory';Expression={$_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory']}},
                  @{Name='M365E3';Expression={if ($_.assignedLicenses.skuid -eq "05e9a617-0261-4cee-bb44-138d3ef5d965"){$true}else{$false}}},
                  @{Name='M365E5';Expression={if ($_.assignedLicenses.skuid -eq "06ebc4ee-1bb5-47dd-8120-11324bc54e06"){$true}else{$false}}},
                  @{Name='NoLicense';Expression={($_.assignedLicenses.count -eq 0)}},
                  @{Name='RoomMailbox';Expression={$_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox']}},
                  @{Name='SharedMailbox';Expression={$_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox']}}

                  $counter = 0 
                  $TotalCount = $Results.Count
  
                  $Results | ForEach-Object {
    
                    $counter++
                    $percentComplete = ($counter / $TotalCount) * 100
                    Write-Progress -Activity "Processing mailboxes..." -PercentComplete $percentComplete -Status "Processing mailbox $counter of $($results.Count)"
                    #$_
                  }

                # Output the results to the console
                $Folder = "$($env:homedrive)\HUD\06_Reporting"

                if(Test-Path -Path $Folder) {
                    "06_Reporting Folder exists..."
                } else {
                    New-Item -Path C:\HUD\ -Name 06_Reporting -ItemType Directory -Force -Confirm:$false
                }

                $Date = Get-Date -f yyyyMMddhhmm
                $FileName = "Full_AAD_Report_$Date.CSV"
                $Results | Export-CSV "$Folder\$FileName" -NoTypeInformation -Encoding UTF8

                Write-Host "The report $FileName has been saved in C:\HUD\06_Reports\" -ForegroundColor Green

            }
            "2" {
                # Connect to Exchange Online
                $UserPrincipalName = Whoami /UPN
                Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false

                # Gather all staff Dynamic Distribution Group Members
                $Results = Get-DynamicDistributionGroupMember -Identity "DL - All Users" | Select DisplayName,PrimarySMTPAddress
                
                # Output the results to the console
                $Folder = "$($env:homedrive)\HUD\06_Reporting"

                if(Test-Path -Path $Folder) {
                    "06_Reporting Folder exists..."
                } else {
                    New-Item -Path C:\HUD\ -Name 06_Reporting -ItemType Directory -Force -Confirm:$false
                }

                $Date = Get-Date -f yyyyMMddhhmm
                $FileName = "DL - All Users_$Date.CSV"
                $Results | Export-CSV "$Folder\$FileName" -NoTypeInformation -Encoding UTF8

                Write-Host "The report $FileName has been saved in C:\HUD\06_Reports\" -ForegroundColor Green

                Disconnect-ExchangeOnline -Confirm:$false

            }
            "3" {
                # Connect to Exchange Online
                $UserPrincipalName = Whoami /UPN
                Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false

                # Gather all staff Dynamic Distribution Group Members
                $Results = Get-DynamicDistributionGroupMember -Identity "DL - Wellington" | Select DisplayName,PrimarySMTPAddress
                
                # Output the results to the console
                $Folder = "$($env:homedrive)\HUD\06_Reporting"

                if(Test-Path -Path $Folder) {
                    "06_Reporting Folder exists..."
                } else {
                    New-Item -Path C:\HUD\ -Name 06_Reporting -ItemType Directory -Force -Confirm:$false
                }

                $Date = Get-Date -f yyyyMMddhhmm
                $FileName = "DL - Wellington_$Date.CSV"
                $Results | Export-CSV "$Folder\$FileName" -NoTypeInformation -Encoding UTF8

                Write-Host "The report $FileName has been saved in C:\HUD\06_Reports\" -ForegroundColor Green

                Disconnect-ExchangeOnline -Confirm:$false

            }
            "4" {
                # Connect to Exchange Online
                $UserPrincipalName = Whoami /UPN
                Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false

                # Gather all staff Dynamic Distribution Group Members
                $Results = Get-DynamicDistributionGroupMember -Identity "DL - Auckland HUD Only" | Select DisplayName,PrimarySMTPAddress
                
                # Output the results to the console
                $Folder = "$($env:homedrive)\HUD\06_Reporting"

                if(Test-Path -Path $Folder) {
                    "06_Reporting Folder exists..."
                } else {
                    New-Item -Path C:\HUD\ -Name 06_Reporting -ItemType Directory -Force -Confirm:$false
                }

                $Date = Get-Date -f yyyyMMddhhmm
                $FileName = "DL - Auckland HUD Only_$Date.CSV"
                $Results | Export-CSV "$Folder\$FileName" -NoTypeInformation -Encoding UTF8

                Write-Host "The report $FileName has been saved in C:\HUD\06_Reports\" -ForegroundColor Green

                Disconnect-ExchangeOnline -Confirm:$false
       
            }
            "5" {
             
            }
            "6" {
            
            }
        }




}


#Select task based on Show-Menu function
do
{
    Clear-Host
    Write-Host ""
    Write-Host '## Digital Support Common Tasks ##' -ForegroundColor Yellow
    $selection = Show-Menu
    switch ($selection) {
                 '1' {New-PIMSession}
                 '2' {New-HUDUser}
                 '3' {New-SharedMailbox}
                 '4' {Search-MailboxAccess}
                 '5' {Add-PhoneNumber}
                 '6' {Add-EmployeeCategory}
                 '7' {Update-UserName}
                 '8' {Export-Reports}
                 'q' {return}
                 }
        pause
        }
until ($selection -eq 'q')
