#��������� �������� �����
$EmailFrom = "it@example.com"
$EmailTo = "admin@example.com"
$Subject = "����� ������������"
$body = "����� ������������ � ������: `n`n"

#�������� ���������
$UserList=IMPORT-CSV  -Encoding default users.csv
$Domain="@example.com"
$domain_short="example\"
# OU ���� ���������
$OU="OU=new,DC=example,DC=com"
#��������� ��������� ������� ��� �������� ���������
$SmtpServer = �smtp.example.com�
$smtp_port=25

#��������� Exchange
$exch_database="Mailbox Database Name"
$exch_server="http://exch.example.com/PowerShell/"

#Skype
$sip_domain="example.com"
$skype_server="https://skype.example.com/ocspowershell/"
$sip_pool="pool.example.com"
