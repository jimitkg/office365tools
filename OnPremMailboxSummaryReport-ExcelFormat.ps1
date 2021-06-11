$reportoutputpath = "C:\Temp"
$reportgeneratedtag = (Get-Date -Format "yyyyMMdd-HHmmss")
$reportoutputfilename = "ExchangeMailboxSummaryDocument_$reportgeneratedtag.xlsx" 
$reportdatafullpath = Join-Path $reportoutputpath -ChildPath $reportoutputfilename   
$AdminCredential = Get-Credential 

 $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ExchangeServer/PowerShell/ -Authentication Kerberos -Credential $AdminCredential  
 
 Import-PSSession $Session -DisableNameChecking

$mailboxlist = get-mailbox 

$reportdata = @()

    foreach ($mailbox in $mailboxlist) {
    $mailboxstats = $mailbox | Get-MailboxStatistics 
    $aduser = Get-ADUser -Identity ($mailbox.SamAccountName) -Properties Title,Description,Department,Division,LastLogon,LastLogonDate,PasswordExpired,PasswordLastSet,AccountExpirationDate,PasswordNeverExpires,Info
    $LastLogonConverted = [DateTime]::FromFileTime($aduser.LastLogon)
            $data = New-Object -TypeName PSCustomObject -Property @{
                mb_SamAccountName = $mailbox.SamAccountName 
                mb_DisplayName = $mailbox.DisplayName 
                ad_Title = $aduser.Title
                ad_Description = $aduser.Description
                ad_Department = $aduser.Department
                ad_Division = $aduser.Division
                ad_Notes = $aduser.Info
                ad_LastLogon = $LastLogonConverted
                #ad_LastLogon = $aduser.LastLogonDate
                mb_UserPrincipalName = $mailbox.UserPrincipalName 
                ad_AccountEnabled = $aduser.Enabled
                ad_AccountExpirationDate = $aduser.AccountExpirationDate
                ad_PasswordNeverExpires = $aduser.PasswordNeverExpires
                ad_PasswordExpired = $aduser.PasswordExpired
                ad_PasswordLastSet = $aduser.PasswordLastSet
                mb_IsMailboxEnabled = $mailbox.IsMailboxEnabled 
                mb_RecipientTypeDetails = $mailbox.RecipientTypeDetails 
                mb_WhenCreated = $mailbox.WhenCreated 
                mb_WhenMailboxCreated = $mailbox.WhenMailboxCreated 
                mb_WhenChanged = $mailbox.WhenChanged 
                mb_ForwardingAddress = $mailbox.ForwardingAddress 
                mb_ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress 
                mb_RoomMailboxAccountEnabled = $mailbox.RoomMailboxAccountEnabled 
                mb_ServerName = $mailbox.ServerName 
                mb_Office = $mailbox.Office 
                mb_UMEnabled = $mailbox.UMEnabled 
                mb_PrimarySmtpAddress = $mailbox.PrimarySmtpAddress 
                mb_WhenSoftDeleted = $mailbox.WhenSoftDeleted 
                mb_EmailAddressPolicyEnabled = $mailbox.EmailAddressPolicyEnabled 
                mb_UnifiedMailbox = $mailbox.UnifiedMailbox 
                mstats_DisplayName = $mailboxstats.DisplayName 
                mstats_MailboxGuid = $mailboxstats.MailboxGuid 
                mstats_ItemCount = $mailboxstats.ItemCount 
                mstats_TotalItemSize = $mailboxstats.TotalItemSize 
                mstats_LastLogonTime = $mailboxstats.LastLogonTime 
                mstats_DatabaseName = $mailboxstats.DatabaseName 
                mstats_DisconnectReason = $mailboxstats.DisconnectReason 
                mstats_DisconnectDate = $mailboxstats.DisconnectDate 
                mb_Alias = $mailbox.Alias 
                mb_Name = $mailbox.Name 
                mb_Identity = $mailbox.Identity 
                mb_OrganizationalUnit = $mailbox.OrganizationalUnit 

            } 
        $reportdata +=$data
    } 

    [datetime] $olddt = (get-date 1900-01-01 )

    $reportdata |
    Select-object mb_SamAccountName ,mb_DisplayName ,ad_Title ,ad_Description , @{Name = 'ad_LastLogon'; Expression = {if($_.ad_LastLogon -lt $olddt) {$null} else {$_.ad_LastLogon}}} ,mstats_LastLogonTime ,mstats_ItemCount,mstats_TotalItemSize,ad_AccountEnabled ,mb_RecipientTypeDetails,ad_AccountExpirationDate ,ad_PasswordNeverExpires,ad_PasswordExpired ,ad_PasswordLastSet ,mb_IsMailboxEnabled, mb_UserPrincipalName ,mb_WhenCreated ,mb_WhenMailboxCreated ,mb_WhenChanged ,mb_ForwardingAddress ,mb_ForwardingSmtpAddress ,mb_RoomMailboxAccountEnabled ,mb_ServerName ,mb_Office ,mb_UMEnabled ,mb_PrimarySmtpAddress ,mb_WhenSoftDeleted ,mb_UnifiedMailbox ,mstats_DisplayName ,mstats_MailboxGuid  ,mstats_DatabaseName ,mstats_DisconnectReason ,mstats_DisconnectDate,mb_Alias,mb_Name ,mb_Identity ,mb_OrganizationalUnit,ad_Department ,ad_Division, ad_Notes |

    Export-Excel -Path $reportdatafullpath -TableName tbl_exchangedata
