
# This maintenance script deletes log files older than 60 days
# Version 1.0

$removedate = ((get-date).AddDays(-60))

#e5 log directory
$logdir1 = 'C:\inetpub\logs\LogFiles\'



Get-ChildItem $logdir1 -File | Where-Object {$_.LastWriteTime -lt $removedate} | Remove-Item
