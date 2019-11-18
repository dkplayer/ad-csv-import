Add-Type -assembly System.Windows.Forms
# Запрос учетных данных администратора. УЗ должна обладать правами для создания пользователей в домене, создания почтовых ящиков в Exchange и Skype for business
$UserCredential = Get-Credential

function Get-RandomCharacters1($length, $characters) { 
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    return [String]$characters[$random]
}

# чтение настроек
. ./settings.ps1

# Создание формы
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Import CSV Users'
$main_form.Width = 210
$main_form.Height = 100
$main_form.AutoSize = $true

#добавление чекбокса на создание почтового ящика в exchange
$checkbox1 = new-object System.Windows.Forms.checkbox
$checkbox1.Location = new-object System.Drawing.Size(10,10)
$checkbox1.Size = new-object System.Drawing.Size(60,20)
$checkbox1.Text = "email"
$checkbox1.Checked = $false
$main_form.Controls.Add($checkbox1) 

#добавление чекбокса на создание уз в Skype for Business
$checkbox2 = new-object System.Windows.Forms.checkbox
$checkbox2.Location = new-object System.Drawing.Size(80,10)
$checkbox2.Size = new-object System.Drawing.Size(60,20)
$checkbox2.Text = "Skype"
$checkbox2.Checked = $false
$main_form.Controls.Add($checkbox2) 

#добавление чекбокса отправки уведомления на email по умолчанию включен
$checkbox3 = new-object System.Windows.Forms.checkbox
$checkbox3.Location = new-object System.Drawing.Size(150,10)
$checkbox3.Size = new-object System.Drawing.Size(120,20)
$checkbox3.Text = "Уведомление"
$checkbox3.Checked = $true
$main_form.Controls.Add($checkbox3) 

function TranslitRU2LAT ($inString) {
    $Translit = @{
        [char]'а' = "a";[char]'А' = "a";
        [char]'б' = "b";[char]'Б' = "b";
        [char]'в' = "v";[char]'В' = "v";
        [char]'г' = "g";[char]'Г' = "g";
        [char]'д' = "d";[char]'Д' = "d";
        [char]'е' = "e";[char]'Е' = "e";
        [char]'ё' = "e";[char]'Ё' = "e";
        [char]'ж' = "zh";[char]'Ж' = "zh";
        [char]'з' = "z";[char]'З' = "z";
        [char]'и' = "i";[char]'И' = "i";
        [char]'й' = "y";[char]'Й' = "y";
        [char]'к' = "k";[char]'К' = "k";
        [char]'л' = "l";[char]'Л' = "l";
        [char]'м' = "m";[char]'М' = "m";
        [char]'н' = "n";[char]'Н' = "n";
        [char]'о' = "o";[char]'О' = "o";
        [char]'п' = "p";[char]'П' = "p";
        [char]'р' = "r";[char]'Р' = "r";
        [char]'с' = "s";[char]'С' = "s";
        [char]'т' = "t";[char]'Т' = "t";
        [char]'у' = "u";[char]'У' = "u";
        [char]'ф' = "f";[char]'Ф' = "f";
        [char]'х' = "kh";[char]'Х' = "kh";
        [char]'ц' = "ts";[char]'Ц' = "ts";
        [char]'ч' = "ch";[char]'Ч' = "ch";
        [char]'ш' = "sh";[char]'Ш' = "sh";
        [char]'щ' = "shch";[char]'Щ' = "shch";
        [char]'ъ' = "";[char]'Ъ' = "";
        [char]'ы' = "y";[char]'Ы' = "y";
        [char]'ь' = "";[char]'Ь' = "";
        [char]'э' = "e";[char]'Э' = "e";
        [char]'ю' = "yu";[char]'Ю' = "yu";
        [char]'я' = "ya";[char]'Я' = "ya"
    }
    $outChars = ""
    foreach ($c in $inChars = $inString.ToCharArray())
        {
        if ($Translit[$c] -cne $Null )
            {$outChars += $Translit[$c]}
        else
            {$outChars += $c}
        }

	Write-Output $outChars

 }
#

$i=1
#Задание массивов данных пользователей
$transname=@()
$FIO=@()
$FN=@()
$SN=@()
$City=@()
$Company=@()
$Departament=@()
$Title=@()
$Phone=@()
$old=@()
$Password=@()
$f_i_o=@()

FOREACH ($Person in $UserList) {
    # разделение ФИО и транслит
    $f_i_o=$Person.FIO.Split(" ")
    $firstletter=$f_i_o[1][0]
    $transname1=$firstletter+$f_i_o[0]
    $transname+=TranslitRU2LAT $transname1
    # заполение переменных
    $FIO+=$Person.FIO
    $FN+=$f_i_o[1]
    $SN+=$f_i_o[0]
    $City+=$Person.City
    $Company+=$Person.Company
    $Departament+=$Person.Departament
    $Title+=$Person.Title
    $Phone+=$Person.Phone
    $old+=$Person.uSERNAME
    # генерация рандомного пароля
    $tmppw1=Get-RandomCharacters1 -length 4 -characters 'abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ'
    $tmppw2=Get-RandomCharacters1 -length 4 -characters '1234567890@&!'
    $tmppw3=Get-RandomCharacters1 -length 4 -characters 'abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ'
    $Password+=$tmppw1+$tmppw2+$tmppw3
    # заполнение формы
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $FIO[($i-1)]
    $Label.Location  = New-Object System.Drawing.Point(10,($i*20+20))
    $Label.AutoSize = $true
    $main_form.Controls.Add($Label)
    $TextBox = New-Object System.Windows.Forms.TextBox
    $textBox.Name = "TextBox"+$i;
    $TextBox.Location  = New-Object System.Drawing.Point(256,($i*20+20))
    $TextBox.Text = $transname[($i-1)]
    $main_form.Controls.Add($TextBox)
    $i=$i+1
}
#добавление кнопок на форму
$button = New-Object System.Windows.Forms.Button
$button.Text = 'Add'
$button.Location = New-Object System.Drawing.Point(10,($i*20+20))
$main_form.Controls.Add($button)


$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location  = New-Object System.Drawing.Point(10,($i*20+40))
$Label2.AutoSize = $true
$main_form.Controls.Add($Label2)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location  = New-Object System.Drawing.Point(10,($i*20+60))
$main_form.Controls.add($ProgressBar)

#клик по кнопке добавить
$button.Add_Click({
$progress=(100/$i)
$ProgressBar.Value = $progress

$controls = $main_form.controls
For ($j=1; $j -lt $i; $j++)  {
$text1=$controls['TextBox'+$j].text
$Label2.Text = "Adding "+$text1+"..."
#Add AD Users
$UPN=$text1+$Domain

# Отключение проверки отзыва сертификата при подключении к удаленному PowerShell
$sessionOption = New-PSSessionOption -SkipRevocationCheck

#Добавление пользователя в AD
Write-host "Создаю Учетную запись в AD"
New-ADUser -Name $FIO[($j-1)] –GivenName $FN[($j-1)] –Surname $SN[($j-1)] –DisplayName $FIO[($j-1)]  –SamAccountName $text1 –UserPrincipalName $UPN -City $City[($j-1)] -Company $Company[($j-1)] -Department $Departament[($j-1)] -Title $Title[($j-1)] -OfficePhone $Phone[($j-1)] -Path $OU
Set-ADAccountPassword -Identity $text1 -NewPassword (ConvertTo-SecureString -AsPlainText $Password[($j-1)] -Force)
Enable-ADAccount -Identity $text1
Start-sleep -s 10



#Создание почтового ящика
if ($checkbox1.Checked -eq $true)
{
    Write-host "Создаю почтовый ящик"
    $body=$body+"Создан почтовый ящик "+$text1+$Domain+"`n`n"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exch_server -Authentication Kerberos -Credential $UserCredential -SessionOption $sessionOption 
    Import-PSSession $Session
    Enable-Mailbox -Identity $domain_short$text1 -Database $exch_database
    Remove-PSSession $Session
    Start-sleep -s 3
  }
  #/создание ящика

#Добавление на Lync User
if ($checkbox2.Checked -eq $true)
{
    Write-host "Создаю Skype"
    $body=$body+"Создан Skype "+$text1+$Domain+"`n`n"
    $Session = New-PSSession -ConnectionUri $skype_server -Credential $UserCredential -SessionOption $sessionOption 
    Import-PSSession $Session
    Enable-CsUser -Identity $domain_short$text1 -RegistrarPool $sip_pool -SipAddressType UserPrincipalName  -SipDomain $sip_domain
    Remove-PSSession $Session
}
#/Добавление на Lync User

# Увеличение прогресс-бара
$ProgressBar.Value = $ProgressBar.Value+$progress

#Добавление записи в текст письма
$body=$body+" "+$Departament[($j-1)]+"`n`n ФИО: "+$FIO[($j-1)]+"`n`n Логин: "+$text1+"`n`n Пароль: "+$Password[($j-1)]+"`n`n`n`n" 
}

#Отправка почты
if ($checkbox3.Checked -eq $true)
{
    $smtp = New-Object net.mail.smtpclient($SmtpServer)
    $smtp.Port=$smtp_port
    $smtp.Send($EmailFrom, $EmailTo, $Subject, $body)

}
#/отправка почты

    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Operation Completed",0,"Done",0x1)
    $ProgressBar.Value = 0
    })#/клик по кнопке добавить



$main_form.ShowDialog()

