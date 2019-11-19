# E-mail notification settings
$EmailFrom = "it@example.com"
$EmailTo = "admin@example.com"
$Subject = "New Users"
$body = "New users in domain: `n`n"

# Common options
$UserList=IMPORT-CSV  -Encoding default users.csv
$Domain="@example.com"
$domain_short="example\"

# OU where new users will be crated
$OU="OU=new,DC=example,DC=com"

# SMTP Server 
$SmtpServer = “smtp.example.com”
$smtp_port=25

# Exchange server remote powershell connection url & mail database name where mailbox will be crated
$exch_database="Mailbox Database Name"
$exch_server="http://exch.example.com/PowerShell/"

# Skype server remote powershell connection url & server pool where sip account will be crated
$sip_domain="example.com"
$skype_server="https://skype.example.com/ocspowershell/"
$sip_pool="pool.example.com"
