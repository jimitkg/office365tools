
<#
Disable junk mail configuration on new mailboxes. Daily Schedule
Can be scheduled via windows task scheduler
#>


Connect-ExchangeOnline -CertificateFilePath "C:\Certs\democert.pfx" -AppID "App ID"  -Organization "demo.onmicrosoft.com"

$lookupday = (get-date).AddDays(-7)

$mailboxesinscope = Get-User -Filter {(RecipientType -eq 'UserMailbox')}|Where-Object {$_.WhenCreated -gt $lookupday} 

$output = "Mailboxes in scope are :  $mailboxesinscope"
$output += "Setting JunkEmailConfig to false"
$mailboxesinscope|Get-EXOMailbox | Set-MailboxJunkEmailConfiguration -Enabled:$false

$results  = $mailboxesinscope|Get-EXOMailbox | Get-MailboxJunkEmailConfiguration


#New-EventLog -Source "ExchangeOnline" -LogName "Application"

foreach ($result in $results) {
    $mailid = $result.Identity.ToString()
    $mailstatus = $result.Enabled
    $outputmsg = "Get-MailboxJunkEmailConfiguration for $mailid is $mailstatus"

    if ($mailstatus) {
        Write-EventLog -LogName Application -Source "ExchangeOnline” -EntryType Error -EventId 2 -Message $outputmsg

        $From = "Notification <notification@demo.com.au>"
        $To = "Admin@demo.com.au"
        $Subject = "Alert - Disable junkmail config failed"
        $Body = "Alert - Disable junkmail config failed for one or more Office 365 mailboxes. Check event viewer or job histroy for more details.</br></br>$outputmsg"
        $SMTPServer = "smtpserver.demo.com.au"
        $SMTPPort = "25"
        $anonUsername = "anonymous"
        $anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
        $anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)
        Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -Credential $anonCredentials  –DeliveryNotificationOption OnSuccess -BodyAsHtml
    }
    ELSE {
            Write-EventLog -LogName Application -Source "ExchangeOnline” -EntryType Information -EventId 1 -Message $outputmsg
    }
}

   Disconnect-ExchangeOnline

