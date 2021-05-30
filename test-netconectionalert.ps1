param (
[string]$URLAddress='app.powerbi.com'
,[int]$URLport=443
,[string]$From = "Notification <notification@demo.com.au>"
,[string[]]$To = "Admin@demo.com.au"
,[string]$Subject = "Alert - Cannot reach : app.powerbi.com"
,[string]$Body = "Internet connection test to app.powerbi.com failed."
,[string]$SMTPServer = "smtp.demo.com.au"
,[int]$SMTPPort = "25"
)


$ConTestResult = Test-NetConnection -ComputerName $URLAddress -Port $URLport
$TcPTestResult = $ConTestResult.TcpTestSucceeded

if ($TcPTestResult)
    {
        Write-EventLog -LogName Application -Source "Internet Connection Test Script” -EntryType Information -EventId 1 -Message "Test connection to app.powerbi.com port 443: Success"
    }
ELSE
    {
        $anonUsername = "anonymous"
        $anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
        $anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

        Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -Credential $anonCredentials  –DeliveryNotificationOption OnSuccess
        Write-EventLog -LogName Application -Source "Internet Connection Test Script” -EntryType Warning -EventId 2 -Message "Test connection to app.powerbi.com port 443: Failed"
    }
