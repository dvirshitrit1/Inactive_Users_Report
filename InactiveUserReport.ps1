# Load Active Directory module
Import-Module ActiveDirectory

# Define parameters
$DaysInactive = 90
$OutputPath = "C:\temp\InactiveUsers.csv"
$smtpServer = "smtp.office365.com"
$smtpFrom = "reports@company.com"
$smtpTo = "admin@company.com"
$smtpPort = 587
$smtpUsername = "your_email@company.com"
$smtpPassword = "your_password"

# Calculate the date threshold
$Time = (Get-Date).AddDays(-$DaysInactive)

# Get inactive AD users and select properties
$inactiveUsers = Get-ADUser -Filter { LastLogonTimeStamp -lt $Time -and Enabled -eq $true } -Properties Name, LastLogonDate |
    Select-Object Name, LastLogonDate

# Export results to CSV
$inactiveUsers | Export-Csv -Path $OutputPath -Encoding UTF8 -NoTypeInformation

# Send Email
$message = New-Object Net.Mail.MailMessage
$message.From = $smtpFrom
$message.To.Add($smtpTo)
$message.Subject = "Inactive Users Report"
$message.Body = "Attached is the monthly inactive users report."

# Attach the CSV file to the email
$attachment = New-Object System.Net.Mail.Attachment($OutputPath)
$message.Attachments.Add($attachment)

# Setup SMTP client with authentication
$smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtp.EnableSsl = $true
$smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUsername, $smtpPassword)

# Send the email
$smtp.Send($message)

# Clean up - delete the CSV file
Remove-Item -Path $OutputPath -Force
