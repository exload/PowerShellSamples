# Common name
$common_name = Get-Date -Format 'dd.MM.yyy HH.mm'

# Folder from backup. Change to actual values
$source = 'E:\'
$source_folder_name = 'Test'
$source_path = $source + $source_folder_name + '\'

# Folder to backup. Change to actual values
$target = 'E:\dest\'
$target_archive = $target + $common_name + '.7z'

# Temp folder. Change as you wish
$temp = $env:TEMP
$temp_folder = $temp + '\' + $source_folder_name + '\'

# Transcript file name
$transcript_file = $target + $common_name + '.txt'

# 7zip exe path. Change to actual values. Tested on v 15.14
$zip = 'D:\Program Files\7-Zip\7z.exe'

Start-Transcript -Path $transcript_file

Copy-Item -Path $source_path -Destination $temp_folder -Verbose -Recurse

&$zip a -mx9 -r -sdel $target_archive $temp_folder | Out-String

Stop-Transcript