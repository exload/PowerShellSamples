Start-Transcript 'C:\Scripts\dc-backup-debug.txt'
$attach = 'C:\Scripts\dc-backup-debug.txt'
#predefine variables
[string]$path = ''
[string]$week_day = ''

#get enviroment variables
[string]$domain_name = $env:USERDNSDOMAIN
[string]$computer_name = $env:COMPUTERNAME

if($domain_name -eq 'имя домена'){
	$path = '\\путь к папке\' + $computer_name
}
elseif ($domain_name -eq 'имя домена'){
	$path = '\\путь к папке\' + $computer_name
}
elseif ($domain_name -eq 'имя домена'){
	$path = '\\путь к папке\' + $computer_name
}

#finding index of current day
$arr_dest_folder = @('blank','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday')

[string]$week_day = (get-date).dayofweek
$day_int = [array]::indexof($arr_dest_folder,$week_day)

#constructing network path
$path += '\' + $day_int

#create backup policy
$DC_WB_Policy = New-WBPolicy

#set BareMetal recovery 
Add-WBBareMetalRecovery -Policy $DC_WB_Policy

#set destination folder
$user = 'имя спец УЗ'
$secure_string = "от пароля спец УЗ"

$secpasswd = $secure_string | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PSCredential($user, $secpasswd)

$Target_folder = New-WBBackupTarget -Credential $credential -NetworkPath $path -NonInheritAcl
Add-WBBackupTarget -Policy $DC_WB_Policy -Target $Target_folder -Force

#set VSS setting
Set-WBVssBackupOption -Policy $DC_WB_Policy -VssFullBackup

#set delete old backups
#Set-WBPolicy -Policy $DC_WB_Policy -AllowDeleteOldBackups -Force

#delete previous backup
write-host ''
Get-WBBackupSet -BackupTarget $Target_folder | Remove-WBBackupSet -Force -Verbose
write-host ''
Remove-Item -Path $path -Recurse -Force -Verbose
write-host ''

#recreate deleted backup target folder
write-host ''
New-Item $path -ItemType Directory
write-host ''

#starting backup process
Start-WBBackup -Policy $DC_WB_Policy

Stop-Transcript

#send a backup log by e-mail
$sender = 'адрес почты отправителя'
$to = 'адрес получателя'

$smtpServer = 'ip или имя сервера'
$port = 587

$date = get-date -Format dd.MM.y
$subject = 'Отчёт о резервном копировании ' + $domain_name + ' ' + $date

wbadmin get versions > C:\Scripts\tmp.txt
$backup_summ = get-content C:\Scripts\tmp.txt | out-string
remove-item -path C:\Scripts\tmp.txt -force

$body_text = '' + $computer_name + "`n `n" + $backup_summ

#password hash
$secure_string = "отправителя"

$secpasswd = $secure_string | ConvertTo-SecureString

$credential = New-Object System.Management.Automation.PSCredential($sender, $secpasswd)

#turn off certificate validation
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }

#send e-mail
Send-MailMessage -From $sender -To $to -SmtpServer $smtpServer -Port $port -UseSsl -Encoding utf8 -Subject $subject -Body $body_text -Credential $credential -Attachments $attach