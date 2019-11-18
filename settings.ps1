#Параметры отправки почты
$EmailFrom = "it@example.com"
$EmailTo = "admin@example.com"
$Subject = "Новые пользователи"
$body = "Новые пользователи в домене: `n`n"

#Основные параметры
$UserList=IMPORT-CSV  -Encoding default users.csv
$Domain="@example.com"
$domain_short="example\"
# OU куда добавлять
$OU="OU=new,DC=example,DC=com"
#Параметры почтового сервера для отправки сообщения
$SmtpServer = “smtp.example.com”
$smtp_port=25

#Параметры Exchange
$exch_database="Mailbox Database Name"
$exch_server="http://exch.example.com/PowerShell/"

#Skype
$sip_domain="example.com"
$skype_server="https://skype.example.com/ocspowershell/"
$sip_pool="pool.example.com"
