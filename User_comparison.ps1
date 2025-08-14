Write-Host "##################################" -ForegroundColor DarkYellow
Write-Host "User matching script (Version 1.1)" -ForegroundColor DarkYellow
Write-Host "##################################" -ForegroundColor DarkYellow
Write-Host ""
Write-Host Comparing user lists... -ForegroundColor Green

# Ensure ImportExcel module is installed
Import-Module ImportExcel

# Load data from CSV and Excel
$inactiveUsers = Import-Csv -Path "Inactive_users.csv"
$subscriptionData = Import-Excel -Path "SubscriptionDataExport.xlsx"

# Extract and normalize email columns
$inactiveEmails = $inactiveUsers | Where-Object { $_.Email -and $_.Email.Trim() -ne "" } |
    ForEach-Object { $_.Email.Trim().ToLower() }

$subscriptionEmails = $subscriptionData | Where-Object { $_.USERNAME -and $_.USERNAME.Trim() -ne "" } |
    ForEach-Object { $_.USERNAME.Trim().ToLower() }

# Find matching emails
$matchingEmails = $inactiveEmails | Where-Object { $subscriptionEmails -contains $_ }

# Output results to console
Write-Host ""
Write-Output "Matching Emails:"
$matchingEmails

# Export results to a text file
$matchingEmails | Out-File -FilePath "MatchingUsers.txt" -Encoding UTF8

# Keep the window open
Write-Host ""
Write-Host Open MatchingUsers.txt -ForegroundColor Green
Write-Host ""
Read-Host -Prompt "Press Enter to close" 
