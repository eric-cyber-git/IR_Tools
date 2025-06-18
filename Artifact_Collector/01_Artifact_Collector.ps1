Write-Host "Services Collection In Progress"
###########  Begin - Capture the Servicess ########### 
get-wmiobject win32_service | select name,state,displayname,processid,startmode,pathname,startname | export-csv C:\temp\_temp\services.csv -NoTypeInformation
###########  End - Capture the Servicess ########### 


Start-Sleep -s 2

Write-Host "Scheduled Tasks Collection In Progress"
###########  Begin - Capture the Scheduled Tasks ########### 
schtasks /query /V /FO csv | convertfrom-csv | where taskname -ne "TaskName" | select hostname,taskname, 'next run time',status,'logon mode','last run time', author, 'task to run','start in',comment,'scheduled task state','run as user' | Export-Csv "C:\temp\_temp\scheduled_tasks.csv" -NoTypeInformation
###########  End - Capture the Scheduled Tasks ########### 


Start-Sleep -s 2

Write-Host "DNS Cache Collection In Progress"
###########  Begin - Capture DNS Cache ########### 
Get-DnsClientCache | select-object Data,Name,Type| Export-Csv "C:\temp\_temp\dns_cache.csv" -NoTypeInformation
###########  End - Capture DNS Cache ########### 


Start-Sleep -s 2

Write-Host "Process list Collection In Progress"
###########  Begin - Capture the Process List ########### 
get-wmiobject win32_process | select name,executablepath,processid,parentprocessid,commandline | export-csv -Path "C:\temp\_temp\plist.csv" -NoTypeInformation
###########  End - Capture the Process Lists ########### 


Start-Sleep -s 2

Write-Host "Autoruns Collection In Progress"
###########  Begin - Capture the AutoRuns Details ########### 
$ARoutput = C:\temp\_temp\autorunsc.exe -c -h -s -nobanner -accepteula
$ARoutput | ConvertFrom-Csv | Export-Csv -Path "C:\temp\_temp\AutoRuns.csv" -NoTypeInformation
###########  End - Capture the AutoRuns Details ########### 


Start-Sleep -s 2
Write-Host "Netstat Collection In Progress"

###########  BEGIN - Capture the Netstat Detailst ########### 
$timespan = New-TimeSpan -Minutes 1
$timer = [diagnostics.stopwatch]::startnew()
while ($timer.elapsed -lt $timespan){
$gettcpconnections = Get-NetTCPConnection | Where-Object state -ne "Bound" | Select-Object localaddress,localport,remoteaddress,remoteport,state,owningprocess, @{Name="process";Expression={(Get-Process -id $_.OwningProcess).ProcessName}} | Export-Csv C:\Temp\_temp\net_connect.csv -NoTypeInformation -Append
$gettcpconnections
Write-Host "Working....."
start-sleep -Seconds 3

}
Write-Host "Capture Complete"
###########  END - Capture the Netstat Detailst ########### 


Start-Sleep -Seconds 2

###########  BEGIN - Create Target Directory ########### 
mkdir Target
###########  END - Create Target Directory ########### 



###########  BEGIN - Write the host + Date\Time info to .txt ########### 
$name = hostname
$date = get-date -Format s
# Remove : from date information
$date_formatted = $date -replace ":", "-"
# Add the Hostname and Date\Time Details to variable
$file_name = $date_formatted + "_" + $name + ".txt"
#Create a new file with Timestamp
New-Item -Path "C:\temp\_temp\Target" -Name "$file_name" -ItemType "File"
###########  END - Write the host + Date\Time info to .txt ########### 


###########  Begin - Folder Manipulation ########### 

# Move all of the generated Target files to the Target folder. 
move-Item C:\temp\_temp\*.csv C:\Temp\_temp\Target
move-Item C:\temp\_temp\*.txt C:\Temp\_temp\Target
#Move-Item -path "C:\temp\_temp\Target" "\\epasseno\C$\Case_Data\"

###########  END - Folder Manipulation  ########### 

