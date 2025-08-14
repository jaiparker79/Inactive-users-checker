# Inactive-users-checker
This PowerShell script is used to identify inactive users from a Library database by programatically comparing them to a list of Inactive Users exported from the Alma library management system. It was made by Jai Parker, Information
Access Librarian at the Queensland University of Technology with help from Microsoft Copilot.  As per the [license](./LICENSE), caveat emptor.

This script requires using the ImportExcel PowerShell module.  To install this run PowerShell as Administrator and enter: 

`Install-Module -Name ImportExcel -Scope CurrentUser`

### Input
1. Generate a list of Inactive Users in Alma Analytics and export their records, ensuring the Email Address.  The output file must be named Inactive_users.csv and contain a Row heading titled Email
2. Export a full list of users from the relevant Library database. The file must be called SubscriptionDataExport.xlsx and contain a Row heading with the user's email entitled USERNAME
3. Ensure both Inactive_users.csv and SubscriptionDataExport.xlsx are in the same folder as User_comparison.ps1
4. Right click User_comparison.ps1 and select Run with PowerShell

### Output
The script will display a list of users in the PowerShell window entitled Matching emails: and it also generates this list in a file MatchingUsers.txt

