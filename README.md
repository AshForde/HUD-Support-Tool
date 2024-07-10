# HUD-Support-Tool
A simple tool leveraging PowerShell and Graph SDK to assist Digital Support. 

## Comment
This is an actively updated tool so the instructions/steps below may change over time however the process will remain the same for accessing the tool.

## Requirements
Administrators running this script will require the necessary PowerShell Modules already be installed on their local machine.

  - AzureAD
  - AzureADPreview
  - ExchangeOnlineManagement
  - Microsoft.Graph
  - Microsoft.Online.SharePoint.PowerShel
  - MicrosoftTeams
  - MSOnline
  - PnP.PowerShell
  - SharePointPnPPowerShellOnline
  - DCToolbox

Administrators will also need to confirm that they have the necessary RBAC roles to action tasks within M365 and Azure AD

## Running the Script
Specified HUD staff will be granted permissions within Azure AD to utilize this tool however it can be used on any tenant as long as the user has the necessary permissions to run the commands.

For Digital Support staff a direct shortcut link to the tool can be deployed for install via the Company Portal, see below:

- This requires being added to the PowerShell Administrators group in Azure AD 

![image](https://user-images.githubusercontent.com/108703933/219210571-6349c14f-812e-4ce9-909b-794424b22904.png)

## Operation
Once successfully deployed administrators can run the script by clicking on the desktop shortcut

![image](https://user-images.githubusercontent.com/108703933/219211239-a49c4310-d7f7-474f-a16e-210f61cdf0fa.png)

The administrator will be presented with a landing page displaying all the functions that can be taken 

![image](https://user-images.githubusercontent.com/108703933/219214012-2e5dc2d7-ba0d-49c4-bd21-ca4cc87ccd91.png)

Each option performs a common task processed by Digital Support. My recommendation is to run the Elevating PIM roles task first to ensure you have the right access before running further tasks. Below is an example of running option 1. 

### Elevating PIM Roles

1. Selecting option 1

![image](https://user-images.githubusercontent.com/108703933/219219320-966f15c6-7fcb-4f36-9563-a87da83f2a33.png)

2. Selecting Digital Workplace Support - as the name suggests Digital Support will have pre-defined roles assigned to them in Azure AD PIM. Selecting this option will ask the administrator to provide a business justification for elevating i.e. "Performing business duties". It automatically sets the PIM activation for the longest time permitted on the respective role. See below:

![image](https://user-images.githubusercontent.com/108703933/219219998-d08a181c-84b5-401c-9b97-0836e2cc7776.png)

3. Providing a reason for elevating

![image](https://user-images.githubusercontent.com/108703933/219220192-a49512b9-5db6-4416-ae96-f838404ef554.png)

4. PIM'ing up to each of the pre-approved roles within the Digital Workplace Support choice made at step 2.

![image](https://user-images.githubusercontent.com/108703933/219220542-5212eae0-cb96-434c-9051-8654aecf85a3.png)

5. Press Enter to continue.... this will take you back to the landing page for the app. 






























