# Q1 - import the csv as a variable
# For Windows - make sure Users.csv is in C:\temp\ directory - please create if need be!!
# This was compliled on macOS and worked OK, I've hashed out the macOS directory commands.
# Replace with Windows > C:\temp\ and hope it still works there too!

#$emailUsers = import-CSV /Users/stevenslaviero/Desktop/Powershell/Users.csv | select EmailAddress,UserPrincipalName,Site,@{Name="MailboxSizeGB";Expression={[math]::Round($_.MailboxSizeGB, 2)}},AccountType

$emailUsers = import-CSV C:\temp\Users.csv  | select EmailAddress,UserPrincipalName,Site,@{Name="MailboxSizeGB";Expression={[math]::Round($_.MailboxSizeGB, 2)}},AccountType

# Q2 - How many users are there?

$emailUsers.count

# Q3 - What is the total size of all mailboxes?

$total = $emailUsers | Measure-Object 'MailboxSizeGB' -Sum
$total.sum

# Q4 - How many accounts exist with non-identical EmailAddress/UserPrincipalName? Be mindful of case sensitivity.

$objects = @{
  ReferenceObject = $emailUsers.EmailAddress
  DifferenceObject = $emailUsers.UserPrincipalName
}
Compare-Object @objects -CaseSensitive

# ANSWER: =  Shows the differences both ways.

# Q5 - Same as question 3, but limited only to Site: NYC

$nycObjects = $emailUsers | Where { $_.Site -eq 'NYC'}
$Objects2 = @{
  ReferenceObject = $nycObjects.EmailAddress
  DifferenceObject = $nycObjects.UserPrincipalName
}
Compare-Object @Objects2 -CaseSensitive
Compare-Object @Objects2 -IncludeEqual

# ANSWER = They are all equal!!

# Q6 - How many Employees (AccountType: Employee) have mailboxes larger than 10 GB?  (remember MailboxSizeGB is already in GB.)

$employees = $emailUsers | Where { $_.AccountType -eq 'Employee'}
$mailboxGT10 = $employees.MailboxSizeGB -gt 10
$mailboxGT10.count

# Q7 - Provide a list of the top 10 users with EmailAddress @domain2.com in Site: NYC by mailbox size, descending.
# The boss already knows that they’re @domain2.com; he wants to only know their usernames, that is, the part of the EmailAddress before the “@” symbol.  There is suspicion that IT Admins managing domain2.com are a quirky bunch and are encoding hidden messages in their directory via email addresses.  Parse out these usernames (in the expected order) and place them in a single string, separated by spaces – should look like: “user1 user2 … user10”

$nycObjects = $emailUsers | Where { $_.Site -eq 'NYC'} | Sort-Object MailboxSizeGB -Descending
$output = $nycObjects.EmailAddress -replace ".{12}$"
Write-Output "$($output)"

# Q8 - Create a new CSV file that summarizes Sites, using the following headers: Site, TotalUserCount, EmployeeCount, ContractorCount, TotalMailboxSizeGB, AverageMailboxSizeGB
# Create this CSV file based off of the original Users.csv.  Note that the boss is picky when it comes to formatting – make sure that AverageMailboxSizeGB is formatted to the nearest tenth of a GB (e.g. 50.124124 is formatted as 50.1).  You must use PowerShell to format this because Excel is down for maintenance.

# Very long winded script, I'm sure there is a better way to do this - but it gets the job done!

#BOS
$bosUsers = $emailUsers | Where { $_.Site -eq 'BOS'}
$bosEmpl = $bosUsers | Where { $_.AccountType -eq 'Employee'}
$bosEmplSum = $bosEmpl.count

$bosCont = $bosUsers | Where { $_.AccountType -eq 'Contractor'}
$bosContSum = $bosCont.count

$bosTMBS = $bosUsers | Measure-Object 'MailboxSizeGB' -Sum

$bosAvgMBS = $bosUsers | Measure-Object 'MailboxSizeGB' -Average 
$bosAvgMBSDP = [MATH]::Round($bosAvgMBS.average,1)


#BRZ
$brzUsers = $emailUsers | Where { $_.Site -eq 'BRZ'}
$brzEmpl = $brzUsers | Where { $_.AccountType -eq 'Employee'}
$brzEmplSum = $brzEmpl.count

$brzCont = $brzUsers | Where { $_.AccountType -eq 'Contractor'}
$brzContSum = $brzCont.count

$brzTMBS = $brzUsers | Measure-Object 'MailboxSizeGB' -Sum

$brzAvgMBS = $brzUsers | Measure-Object 'MailboxSizeGB' -Average 
$brzAvgMBSDP = [MATH]::Round($brzAvgMBS.average,1)


#LAS
$lasUsers = $emailUsers | Where { $_.Site -eq 'LAS'}
$lasEmpl = $lasUsers | Where { $_.AccountType -eq 'Employee'}
$lasEmplSum = $lasEmpl.count

$lasCont = $lasUsers | Where { $_.AccountType -eq 'Contractor'}
$lasContSum = $lasCont.count

$lasTMBS = $lasUsers | Measure-Object 'MailboxSizeGB' -Sum

$lasAvgMBS = $lasUsers | Measure-Object 'MailboxSizeGB' -Average 
$lasAvgMBSDP = [MATH]::Round($lasAvgMBS.average,1)


#NYC
$nycUsers = $emailUsers | Where { $_.Site -eq 'NYC'}
$nycEmpl = $nycUsers | Where { $_.AccountType -eq 'Employee'}
$nycEmplSum = $nycEmpl.count

$nycCont = $nycUsers | Where { $_.AccountType -eq 'Contractor'}
$nycContSum = $nycCont.count

$nycTMBS = $nycUsers | Measure-Object 'MailboxSizeGB' -Sum

$nycAvgMBS = $nycUsers | Measure-Object 'MailboxSizeGB' -Average 
$nycAvgMBSDP = [MATH]::Round($nycAvgMBS.average,1)


#RIO
$rioUsers = $emailUsers | Where { $_.Site -eq 'RIO'}
$rioEmpl = $rioUsers | Where { $_.AccountType -eq 'Employee'}
$rioEmplSum = $rioEmpl.count

$rioCont = $rioUsers | Where { $_.AccountType -eq 'Contractor'}
$rioContSum = $rioCont.count

$rioTMBS = $rioUsers | Measure-Object 'MailboxSizeGB' -Sum

$rioAvgMBS = $rioUsers | Measure-Object 'MailboxSizeGB' -Average 
$rioAvgMBSDP = [MATH]::Round($rioAvgMBS.average,1)


#SEA
$seaUsers = $emailUsers | Where { $_.Site -eq 'SEA'}
$seaEmpl = $seaUsers | Where { $_.AccountType -eq 'Employee'}
$seaEmplSum = $seaEmpl.count

$seaCont = $seaUsers | Where { $_.AccountType -eq 'Contractor'}
$seaContSum = $seaCont.count

$seaTMBS = $seaUsers | Measure-Object 'MailboxSizeGB' -Sum

$seaAvgMBS = $seaUsers | Measure-Object 'MailboxSizeGB' -Average 
$seaAvgMBSDP = [MATH]::Round($seaAvgMBS.average,1)


#TOR
$torUsers = $emailUsers | Where { $_.Site -eq 'TOR'}
$torEmpl = $torUsers | Where { $_.AccountType -eq 'Employee'}
$torEmplSum = $torEmpl.count

$torCont = $torUsers | Where { $_.AccountType -eq 'Contractor'}
$torContSum = $torCont.count

$torTMBS = $torUsers | Measure-Object 'MailboxSizeGB' -Sum

$torAvgMBS = $torUsers | Measure-Object 'MailboxSizeGB' -Average 
$torAvgMBSDP = [MATH]::Round($torAvgMBS.average,1)


$headers = "Site", "EmployeeCount", "ContractorCount", "TotalMailboxSizeGB", "AverageMailboxSizeGB"
$psObject = New-Object psobject
foreach($header in $headers)
{
 Add-Member -InputObject $psobject -MemberType noteproperty -Name $header -Value ""
}
#$psObject | Export-Csv /Users/stevenslaviero/Desktop/Powershell/Users1.csv -NoTypeInformation
$psObject | Export-Csv C:\temp\Users1.csv -NoTypeInformation

#BOS
$bbosSite = $bosUsers.site[0]
$bbosTotalMailboxSizeGB = $bosTMBS.sum
$bosdata = "$bbosSite, $bosEmplSum, $bosContSum, $bbosTotalMailboxSizeGB, $bosAvgMBSDP"
#$bosdata | Out-File /Users/stevenslaviero/Desktop/Powershell/Users1.csv -append
$bosdata | Out-File C:\temp\Users1.csv -append

#BRZ
$bbrzSite = $brzUsers.site[0]
$bbrzTotalMailboxSizeGB = $brzTMBS.sum
$brzdata = "$bbrzSite, $brzEmplSum, $brzContSum, $bbrzTotalMailboxSizeGB, $brzAvgMBSDP"
#$brzdata | Out-File /Users/stevenslaviero/Desktop/Powershell/Users1.csv -append
$brzdata | Out-File C:\temp\Users1.csv -append

#LAS
$blasSite = $lasUsers.site[0]
$blasTotalMailboxSizeGB = $lasTMBS.sum
$lasdata = "$blasSite, $lasEmplSum, $lasContSum, $blasTotalMailboxSizeGB, $lasAvgMBSDP"
#$lasdata | Out-File /Users/stevenslaviero/Desktop/Powershell/Users1.csv -append
$lasdata | Out-File C:\temp\Users1.csv -append

#NYC
$bNYCSite = $nycUsers.site[0]
$bNYCTotalMailboxSizeGB = $nycTMBS.sum
$NYCdata = "$bNYCSite, $nycEmplSum, $nycContSum, $bNYCTotalMailboxSizeGB, $nycAvgMBSDP"
#$NYCdata | Out-File /Users/stevenslaviero/Desktop/Powershell/Users1.csv -append
$NYCdata | Out-File C:\temp\Users1.csv -append

#RIO
$brioSite = $rioUsers.site[0]
$brioTotalMailboxSizeGB = $rioTMBS.sum
$riodata = "$brioSite, $rioEmplSum, $rioContSum, $brioTotalMailboxSizeGB, $rioAvgMBSDP"
#$riodata | Out-File /Users/stevenslaviero/Desktop/Powershell/Users1.csv -append
$riodata | Out-File C:\temp\Users1.csv -append

#SEA
$bseaSite = $seaUsers.site[0]
$bseaTotalMailboxSizeGB = $seaTMBS.sum
$seadata = "$bseaSite, $seaEmplSum, $seaContSum, $bseaTotalMailboxSizeGB, $seaAvgMBSDP"
#$seadata | Out-File /Users/stevenslaviero/Desktop/Powershell/Users1.csv -append
$seadata | Out-File C:\temp\Users1.csv -append

#TOR
$btorSite = $torUsers.site[0]
$btorTotalMailboxSizeGB = $torTMBS.sum
$tordata = "$btorSite, $torEmplSum, $torContSum, $btorTotalMailboxSizeGB, $torAvgMBSDP"
#$tordata | Out-File /Users/stevenslaviero/Desktop/Powershell/Users1.csv -append
$tordata | Out-File C:\temp\Users1.csv -append
