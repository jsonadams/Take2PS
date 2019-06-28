

$Import = import-csv 'C:\Users\lscadmin\Desktop\Users.csv'
#First thing’s first, import the csv as a variable.  Use this variable for all subsequent tasks.
Write-Host '1. C:\Users\lscadmin\Desktop\Users.csv imported as variable $Import' -ForegroundColor Green

#How many users are there?
Write-Host '2. There are' $Import.Count 'users.' -ForegroundColor Green

#What is the total size of all mailboxes?
$totalMBsize = $Import | Measure-Object -Property MailboxSizeGB -Sum
Write-Host '3. Total size of all mailboxes :' $totalMBsize.Sum 'GB' -ForegroundColor Green

#How many accounts exist with non-identical EmailAddress/UserPrincipalName? Be mindful of case sensitivity.
$diffCHAR = Compare-Object -ReferenceObject $Import.EmailAddress -DifferenceObject $Import.UserPrincipalName -PassThru
$diff = Compare-Object -ReferenceObject $Import.EmailAddress -DifferenceObject $Import.UserPrincipalName -PassThru -CaseSensitive
$diffCS = Compare-Object -ReferenceObject $diffCHAR -DifferenceObject $diff -PassThru
Write-Host '4. There are' ($diff.count/2) 'accounts that have Email/samName mismatches' -ForegroundColor Green
Write-Host '     Accounts mismatched by string value :' -ForegroundColor Yellow
Write-host '    '($diffCHAR | Sort-Object ) -ForegroundColor DarkYellow
Write-Host '     Accounts mismatched by case sensitivity :' -ForegroundColor Yellow
Write-host '    '($diffCS | Sort-Object ) -ForegroundColor DarkYellow

#Same as question 3, but limited only to Site: NYC
$ImportNYC = $Import | Where-Object {$_.Site -eq 'NYC'} 
$NYCmbSize = $ImportNYC | Measure-Object -Property MailboxSizeGB -Sum 
Write-Host '5. Total size of all NYC site mailboxes :' $NYCmbSize.Sum 'GB' -ForegroundColor Green

#How many Employees (AccountType: Employee) have mailboxes larger than 10 GB?  (remember MailboxSizeGB is already in GB.)
$gtTEN = $Import.Where({ [int]$_.MailboxSizeGB -gt 10; })
$gtTENEMP = $gtTEN | Where-Object {$_.AccountType -eq 'Employee'} 
Write-Host '6. There are' $gtTENEMP.count 'Employees (AccountType: Employee) with mailboxes larger than 10 GB' -ForegroundColor Green

#Provide a list of the top 10 users with EmailAddress @domain2.com in Site: NYC by mailbox size, descending. 
$dom2 =$Import | Where-Object{$_.EmailAddress -match "domain2.com"}
$top10 = $dom2  | Where-Object {$_.Site -eq 'NYC'} | Sort-Object{[int]$_.MailboxSizeGB} -Descending | Select-Object -First 10
Write-Host '7a. The top 10 users with EmailAddress @domain2.com in Site: NYC by mailbox size, descending are:' -ForegroundColor Green
$top10 | Format-Table

#The boss already knows that they’re @domain2.com; he wants to only know their usernames, that is, the part of the EmailAddress before the “@” symbol.  
#There is suspicion that IT Admins managing domain2.com are a quirky bunch and are encoding hidden messages in their directory via email addresses.  
#Parse out these usernames (in the expected order) and place them in a single string, separated by spaces – should look like: “user1 user2 … user10”

$trimd = ($top10.EmailAddress).Replace("@domain2.com","")
Write-Host '7b.' $trimd -ForegroundColor Green

#Create a new CSV file that summarizes Sites, using the following headers: Site, TotalUserCount, EmployeeCount, ContractorCount, TotalMailboxSizeGB, AverageMailboxSizeGB
Add-Content -Path 'C:\Users\lscadmin\Desktop\Employees.csv'  -Value '"Site","TotalUserCount","EmployeeCount","ContractorCount","TotalMailboxSizeGB","AverageMailboxSizeGB"'
$Sites = $Import.Site | Select -Unique
ForEach ($site in $Sites){

$tsUserCount = ($Import | Where-Object {$_.Site -eq $site}).count
$tsEMPCount = ($Import | Where-Object {($_.Site -eq $site) -and ($_.AccountType -eq "Employee")}).count
$tsCTRCount = ($Import | Where-Object {($_.Site -eq $site) -and ($_.AccountType -eq "Contractor")}).count
$tsTOTMB = $Import | Where-Object {$_.Site -eq $site} | Measure-Object -Property MailboxSizeGB -Sum
$tsMBavg = $Import | Where-Object {$_.Site -eq $site} | Measure-Object -Property MailboxSizeGB -Average
$avemb=  [math]::Round($tsMBavg.Average,1)
$nfoDrop = "{0},{1},{2},{3},{4},{5}" -f $site,$tsUserCount,$tsEMPCount,$tsCTRCount,$tsTOTMB.sum,$avemb
$nfoDrop | foreach { Add-Content -Path 'C:\Users\lscadmin\Desktop\Employees.csv' -Value $_ }
}
Write-Host '8. Employees.csv created and saved at C:\Users\lscadmin\Desktop\Employees.csv' -ForegroundColor Green
