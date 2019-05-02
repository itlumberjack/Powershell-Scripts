##########################
#### ITLUMBERJACK.COM ####
### UPDATED ON: 5/2/19 ###
##########################

import-module activedirectory  

########################
### General Settings ###
########################

# Enter in the organization name between the quotation marks
$Organization_Name = "Organization Name"

# Enter in the number of days a user needs to be inactive. 
$Users_DaysInactive = 365

# Enter in the number of days a computer needs to be inactive.
$Computers_DaysInactive = 365

# Email Inactive Users List | 0 = NO | 1 = YES
$Inactive_Users_HTML = 1

# Email Inactive Computers List | 0 = NO | 1 = YES
$Inactive_Computers_HTML = 1

# Export Inactive Users List to CSV | 0 = NO | 1 = YES
$Inactive_Users_CSV = 1

# Export Inactive Computers List to CSV | 0 = NO | 1 = YES
$Inactive_Computers_CSV = 1

# Directory to Save CSV Exports
$Export_Directory = "C:\Users\username\Desktop"


#####################
### SMTP Settings ###
#####################

$SMTP_Username = "Username"
$SMTP_Password = ConvertTo-SecureString "Password" -AsPlainText -Force
$Creds = New-Object System.Management.Automation.PSCredential ($SMTP_Username, $SMTP_Password)
$From = ""
$To = ""
$SMTPServer = ""
$SMTPPort = ""

###################
### CSS Styling ###
################### 

$css = @"
<style>
div {overflow-x:auto;}
table {border-collapse: collapse;width: 100%;}
th, td {text-align: left;padding: 8px;font-family:arial;font-size: 10pt;}
tr:nth-child(even){background-color: #f2f2f2}
tr:hover {background-color: #86ff65;}
th {background-color: #c90022;text-shadow: 1px 1px 4px black;color: white;}
</style>
"@

#############################
### Inactive Users Script ###
#############################

$time = (Get-Date).Adddays(-($Users_DaysInactive)) 

$Inactive_Users = Get-ADUser -Filter {LastLogonDate -lt $time} -Properties PasswordLastSet,Created,LastLogonDate | 

select-object @{Name="User Name"; Expression={$_.Name}}, @{Name="Password Last Changed Date"; Expression={$_.PasswordLastSet}}, @{Name="Account Created Date"; Expression={$_.Created}}, @{Name="Last Logon Date"; Expression={$_.LastLogonDate}}, @{Name="Account Enabled Status"; Expression={$_.Enabled}}

if ($Inactive_Users_CSV -eq 1) {$Inactive_Users | Export-Csv $Export_Directory\InactiveUser-$(get-date -f yyyy-MM-dd).csv}

if ($Inactive_Users_HTML -eq 1) {
$Inactive_Users_HTML = $Inactive_Users | ConvertTo-Html -Head $css
$Subject = $Organization_Name +  " AD Report | Inactive Users | " + (Get-Date).ToString()
$body=@" 
$Inactive_Users_HTML
"@
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Creds –DeliveryNotificationOption OnSuccess -BodyAsHtml
}

#################################
### Inactive Computers Script ###
#################################

$time = (Get-Date).Adddays(-($Computers_DaysInactive)) 

$Inactive_Computers = Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -Properties LastLogonTimeStamp | 

select-object @{Name="Computer Name"; Expression={$_.Name}},@{Name="Distinguished Name"; Expression={$_.DistinguishedName}},@{Name="Last Logon Date"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}},@{Name="Account Enabled Status"; Expression={$_.Enabled}}

if ($Inactive_Computers_HTML -eq 1) {
$Inactive_Computers_HTML = $Inactive_Computers | ConvertTo-Html -Head $css
$Subject = $Organization_Name + " AD Report | Inactive Computers | " + (Get-Date).ToString()
$body=@" 
$Inactive_Computers_HTML
"@
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Creds –DeliveryNotificationOption OnSuccess -BodyAsHtml
}

if ($Inactive_Computers_CSV -eq 1) {$Inactive_Computers | Export-Csv $Export_Directory\InactiveComputer-$(get-date -f yyyy-MM-dd).csv}

