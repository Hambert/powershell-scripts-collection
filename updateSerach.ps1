## This scipt installs new Windows Defender signatures and search for new Windows updates
## It send an e-mail when the work is done
### load settings.txt

Get-Content "C:\Users\Administrator\Documents\Skripte\settings.txt" | foreach-object -begin {$h=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $h.Add($k[0], $k[1]) } }
# to get an item $h.Get_Item("MySetting1")


### Windows Defender
$Versions = Get-MpComputerStatus
$oldVersion = $Versions | select -ExpandProperty AntispywareSignatureVersion

Update-MpSignature

$Versions = Get-MpComputerStatus
$newVersion = $Versions | select -ExpandProperty AntispywareSignatureVersion
$newDate = $Versions | select -ExpandProperty AntispywareSignatureLastUpdated
$fScan = $Versions | select -ExpandProperty FullScanStartTime
$qScan = $Versions | select -ExpandProperty QuickScanStartTime
$mailBody = ""

### Windows update
$Computername = $env:COMPUTERNAME
$updatesession =  [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$Computername))
$UpdateSearcher = $updatesession.CreateUpdateSearcher()
$searchresult = $updatesearcher.Search("IsInstalled=0")
$count  = $searchresult.Updates.Count 


if ( $count -gt 0) {

    # create Mail Body
    $mailBody = $mailBody + "###################################  Windows Update  #######################################`n"
    $mailBody = "There are " + $count + " new update(s) for " + $Computername + " available:`n`n"

     For ($i=0; $i -lt $Count; $i++) {
        $Update  = $searchresult.Updates.Item($i)
        $mailBody = $mailBody + "Title: " + $Update.Title + "`n"
        $mailBody = $mailBody + "KB: " + $Update.KBArticleIDs + "`n"
        $mailBody = $mailBody + "Security bulletin IDs: " + $Update.SecurityBulletinIDs + "`n"
        $mailBody = $mailBody + "Msrc severity: " + $Update.MsrcSeverity + "`n"
        $mailBody = $mailBody + "Is downloaded: " +$Update.IsDownloaded + "`n"
        $mailBody = $mailBody + "Info url: " + $Update.MoreInfoUrls + "`n`n"
     }

} else {
    $mailBody = $mailBody + "###################################  Windows Update  #######################################`n"
    $mailBody = $mailBody + "No updates available.`n`n"
}

if( $oldVersion -ne $newVersion ){

    $mailBody = $mailBody + "`n###################################  Windows Defender  #####################################`n"
    $mailBody = $mailBody + "New signature from  " + $newDate + " installed.`n"
    $mailBody = $mailBody + "New verion: "  + $newVersion + "`n"
    $mailBody = $mailBody + "Old verion: "  + $oldVersion + "`n"
    $mailBody = $mailBody + "Last quick scan: " + $qScan + "`n"
    $mailBody = $mailBody + "Last full scan: "+ $fScan + "`n`n"

} else {

    $mailBody = $mailBody + "`n###################################  Windows Defender  #####################################`n"
    $mailBody = $mailBody + "No updates available.`n"
    $mailBody = $mailBody + "Current signature( "+ $newVersion + ") - " + $newDate + " `n"
    $mailBody = $mailBody + "Last quick scan: " + $qScan + "`n"
    $mailBody = $mailBody + "Last full scan: "+ $fScan + "`n`n"

}

#
# Send Mail

if ( ( $oldVersion -ne $newVersion ) -or  ( $count -gt 0 ) ) {

    # to force TLS11 or higher
    [System.Net.ServicePointManager]::SecurityProtocol = 'TLS11,TLS12'

    $PSEmailServer = $h.Get_Item("mailSmtpURL")
    $pw = Get-Content $h.Get_Item("mailPw") | ConvertTo-SecureString

    $to = $h.Get_Item("mailTo")
    $from = $h.Get_Item("mailFrom")
    $cc = $h.Get_Item("mailCC")
    $cred = New-Object System.Management.Automation.PSCredential $from, $pw

    if ( $oldVersion -ne $newVersion ) {
        $mailSubject = "New security signature installed on " + $Computername + ""
    } else {
        $mailSubject = "Updates on " + $Computername + " available!"
    }

    Send-MailMessage -Credential $cred -from $from -to $to -CC $cc -Subject $mailSubject -body $mailBody  -encoding ([System.Text.Encoding]::UTF8) -UseSSL

 } else {

    if ( $h.Get_Item("SendMailAlway") -eq "True" ) {
         # to force TLS11 or higher
        [System.Net.ServicePointManager]::SecurityProtocol = 'TLS11,TLS12'

        $PSEmailServer = $h.Get_Item("mailSmtpURL")
        $pw = Get-Content $h.Get_Item("mailPw") | ConvertTo-SecureString

        $to = $h.Get_Item("mailTo")
        $from = $h.Get_Item("mailFrom")
        $cc = $h.Get_Item("mailCC")
        $cred = New-Object System.Management.Automation.PSCredential $from, $pw
        $mailSubject = "No Updates on " + $Computername + " available"

        Send-MailMessage -Credential $cred -from $from -to $to -CC $cc -Subject $mailSubject -body $mailBody  -encoding ([System.Text.Encoding]::UTF8) -UseSSL
    }
 }
