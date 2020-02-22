### load settings.txt

Get-Content "C:\Users\Administrator\Documents\Skripte\settings.txt" | foreach-object -begin {$h=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $h.Add($k[0], $k[1]) } }
# to get an item $h.Get_Item("MySetting1")


### get computername
$Computername = $env:computername | Select-Object
$date =  Get-Date

$mailBody = "Neustart von `"" + $Computername + "`" wurde neu gestartet.`n`nAm " + $date + "."

# to force TLS11 or higher
[System.Net.ServicePointManager]::SecurityProtocol = 'TLS11,TLS12'

    
$PSEmailServer = $h.Get_Item("mailSmtpURL")
$pw = Get-Content $h.Get_Item("mailPw") | ConvertTo-SecureString

$to = $h.Get_Item("mailTo")
$from = $h.Get_Item("mailFrom")
$cc = $h.Get_Item("mailCC")
$cred = New-Object System.Management.Automation.PSCredential $from, $pw



$mailSubject = "Start von " + $Computername + " durchgeführt."


Send-MailMessage -Credential $cred -from $from -to $to -CC $cc -Subject $mailSubject -body $mailBody  -encoding ([System.Text.Encoding]::UTF8) -UseSSL