Add-Type -assembly System.Windows.Forms
# Request administrator credentials. KM must have the rights to create users in the domain, create mailboxes in Exchange and Skype for business
$UserCredential = Get-Credential

function Get-RandomCharacters1($length, $characters) { 
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    return [String]$characters[$random]
}

# settings reading
. ./settings.ps1

# Making main form
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Import CSV Users'
$main_form.Width = 210
$main_form.Height = 100
$main_form.AutoSize = $true

# adding a checkbox to create a mailbox in exchange
$checkbox1 = new-object System.Windows.Forms.checkbox
$checkbox1.Location = new-object System.Drawing.Size(10,10)
$checkbox1.Size = new-object System.Drawing.Size(60,20)
$checkbox1.Text = "email"
$checkbox1.Checked = $false
$main_form.Controls.Add($checkbox1) 

# adding a checkbox for creating ties in Skype for Business
$checkbox2 = new-object System.Windows.Forms.checkbox
$checkbox2.Location = new-object System.Drawing.Size(80,10)
$checkbox2.Size = new-object System.Drawing.Size(60,20)
$checkbox2.Text = "Skype"
$checkbox2.Checked = $false
$main_form.Controls.Add($checkbox2) 

# Adding a checkbox for sending email notifications is enabled by default
$checkbox3 = new-object System.Windows.Forms.checkbox
$checkbox3.Location = new-object System.Drawing.Size(150,10)
$checkbox3.Size = new-object System.Drawing.Size(120,20)
$checkbox3.Text = "Óâåäîìëåíèå"
$checkbox3.Checked = $true
$main_form.Controls.Add($checkbox3) 

# Cyrilyc translit table 
function TranslitRU2LAT ($inString) {
    $Translit = @{
        [char]'à' = "a";[char]'À' = "a";
        [char]'á' = "b";[char]'Á' = "b";
        [char]'â' = "v";[char]'Â' = "v";
        [char]'ã' = "g";[char]'Ã' = "g";
        [char]'ä' = "d";[char]'Ä' = "d";
        [char]'å' = "e";[char]'Å' = "e";
        [char]'¸' = "e";[char]'¨' = "e";
        [char]'æ' = "zh";[char]'Æ' = "zh";
        [char]'ç' = "z";[char]'Ç' = "z";
        [char]'è' = "i";[char]'È' = "i";
        [char]'é' = "y";[char]'É' = "y";
        [char]'ê' = "k";[char]'Ê' = "k";
        [char]'ë' = "l";[char]'Ë' = "l";
        [char]'ì' = "m";[char]'Ì' = "m";
        [char]'í' = "n";[char]'Í' = "n";
        [char]'î' = "o";[char]'Î' = "o";
        [char]'ï' = "p";[char]'Ï' = "p";
        [char]'ð' = "r";[char]'Ð' = "r";
        [char]'ñ' = "s";[char]'Ñ' = "s";
        [char]'ò' = "t";[char]'Ò' = "t";
        [char]'ó' = "u";[char]'Ó' = "u";
        [char]'ô' = "f";[char]'Ô' = "f";
        [char]'õ' = "kh";[char]'Õ' = "kh";
        [char]'ö' = "ts";[char]'Ö' = "ts";
        [char]'÷' = "ch";[char]'×' = "ch";
        [char]'ø' = "sh";[char]'Ø' = "sh";
        [char]'ù' = "shch";[char]'Ù' = "shch";
        [char]'ú' = "";[char]'Ú' = "";
        [char]'û' = "y";[char]'Û' = "y";
        [char]'ü' = "";[char]'Ü' = "";
        [char]'ý' = "e";[char]'Ý' = "e";
        [char]'þ' = "yu";[char]'Þ' = "yu";
        [char]'ÿ' = "ya";[char]'ß' = "ya"
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
# Defining user data arrays
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
    # FIO split
    $f_i_o=$Person.FIO.Split(" ")
    $firstletter=$f_i_o[1][0]
    $transname1=$firstletter+$f_i_o[0]
    $transname+=TranslitRU2LAT $transname1
    $FIO+=$Person.FIO
    $FN+=$f_i_o[1]
    $SN+=$f_i_o[0]
    $City+=$Person.City
    $Company+=$Person.Company
    $Departament+=$Person.Departament
    $Title+=$Person.Title
    $Phone+=$Person.Phone
    $old+=$Person.uSERNAME
    # password generation
    $tmppw1=Get-RandomCharacters1 -length 4 -characters 'abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ'
    $tmppw2=Get-RandomCharacters1 -length 4 -characters '1234567890@&!'
    $tmppw3=Get-RandomCharacters1 -length 4 -characters 'abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ'
    $Password+=$tmppw1+$tmppw2+$tmppw3
    # adding fields on form
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
# adding buttons & labels to form
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

# button click procedure
$button.Add_Click({
    $progress=(100/$i)
    $ProgressBar.Value = $progress

    $controls = $main_form.controls
    For ($j=1; $j -lt $i; $j++)  {
        $text1=$controls['TextBox'+$j].text
        $Label2.Text = "Adding "+$text1+"..."
        #Add AD Users
        $UPN=$text1+$Domain

        # disable certificate validation on remote PowerShell connect
        $sessionOption = New-PSSessionOption -SkipRevocationCheck

        # creating user account in AD
        Write-host "Creating AD user account..."
        New-ADUser -Name $FIO[($j-1)] –GivenName $FN[($j-1)] –Surname $SN[($j-1)] –DisplayName $FIO[($j-1)]  –SamAccountName $text1 –UserPrincipalName $UPN -City $City[($j-1)] -Company $Company[($j-1)] -Department $Departament[($j-1)] -Title $Title[($j-1)] -OfficePhone $Phone[($j-1)] -Path $OU
        Set-ADAccountPassword -Identity $text1 -NewPassword (ConvertTo-SecureString -AsPlainText $Password[($j-1)] -Force)
        Enable-ADAccount -Identity $text1
        #/ crating user account

        Start-sleep -s 10

        # crating mailbox if checkbox selected
        if ($checkbox1.Checked -eq $true)
            {
                Write-host "Creating Exchange mailbox"
                $body=$body+"Mailbox created "+$text1+$Domain+"`n`n"
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exch_server -Authentication Kerberos -Credential $UserCredential -SessionOption $sessionOption 
                Import-PSSession $Session
                Enable-Mailbox -Identity $domain_short$text1 -Database $exch_database
                Remove-PSSession $Session
                Start-sleep -s 3
            }
        #/ creating mailbox

        # Creating Skype sip account
        if ($checkbox2.Checked -eq $true)
            {
                Write-host "Creating Skype account"
                $body=$body+"Skype account created "+$text1+$Domain+"`n`n"
                $Session = New-PSSession -ConnectionUri $skype_server -Credential $UserCredential -SessionOption $sessionOption 
                Import-PSSession $Session
                Enable-CsUser -Identity $domain_short$text1 -RegistrarPool $sip_pool -SipAddressType UserPrincipalName  -SipDomain $sip_domain
                Remove-PSSession $Session
            }
        #/Creating Skype sip account

        # increase progress bar
        $ProgressBar.Value = $ProgressBar.Value+$progress

        # adding record to notification email body
        $body=$body+" "+$Departament[($j-1)]+"`n`n ÔÈÎ: "+$FIO[($j-1)]+"`n`n Ëîãèí: "+$text1+"`n`n Ïàðîëü: "+$Password[($j-1)]+"`n`n`n`n" 
    }
    #/for

    # sending email if selected checkbox
    if ($checkbox3.Checked -eq $true)
        {
        $smtp = New-Object net.mail.smtpclient($SmtpServer)
        $smtp.Port=$smtp_port
        $smtp.Send($EmailFrom, $EmailTo, $Subject, $body)
        }
    #/ sending email

    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Operation Completed",0,"Done",0x1)
    $ProgressBar.Value = 0
})
#/ button click procedure

$main_form.ShowDialog()

