#####################################################################################################################################################
#Script Name : 
#Script Version: Version 0.1
#
#
#Script Description : Script to read Date column from csv file & fetch data depending on threshold
#
#Pre-Requisites:1) Csv file should contain one Column with name "Date" which will have all dates[in same date format]
#              
#               
#
#
#####################################################################################################################################################
#########################################################   INPUT SECTION   #########################################################################
#####################################################################################################################################################
#fetching script path

$scriptParentPath = (Resolve-Path "..\..\").Path

$ConfPath =  Join-Path -path $scriptParentPath -ChildPath "/conf"
$libPath =  Join-Path -path $scriptParentPath -ChildPath "/lib"
$logPath =  Join-Path -path $scriptParentPath -ChildPath "/log"
$masterBotPath =  Join-Path -path $scriptParentPath -ChildPath "/scripts/MasterBot"
$microBotsPath =  Join-Path -path $scriptParentPath -ChildPath "/scripts/MicroBots"
$jsonConfig = Get-content -Path "$confPath\config.json" |  ConvertFrom-Json

#$jsonfile = Join-Path -Path $ConfPath -ChildPath "/config.json" 
#$logentrybot = $jsonfile.logentrybot

$logentrybot = $jsonConfig.logentrybot
$logfilepath = $jsonConfig.logfilepath
$smtpServer = $jsonConfig.smtpServer
$fromEmail = $jsonConfig.fromEmail
$sendEmail = $jsonConfig.sendEmail
$threshold = $jsonConfig.threshold
$cc = $jsonConfig.cc
$file = $jsonConfig.file | Import-Csv

$html_output = $jsonConfig.html_output


#$file = Import-Csv $file





$before = "<style>"
$before = $before + "TABLE{border-width: 1px;border-style: solid;border-color:black;}"
$before = $before + "Table{background-color:#DFFFFF;border-collapse: collapse;}"
$before = $before + "TH{border-width:1px;padding:0px;border-style:solid;border-color:black;}"
$before = $before + "TD{border-width:1px;padding-left:5px;border-style:solid;border-color:black;}"
$before = $before + "</style>"

$between = "<style>"
$between = $between + "TABLE{border-width: 1px;border-style: solid;border-color:black;}"
$between = $between + "Table{background-color:#DFFFFF;border-collapse: collapse;}"
$between = $between + "TH{border-width:1px;padding:0px;border-style:solid;border-color:black;}"
$between = $between + "TD{border-width:1px;padding-left:5px;border-style:solid;border-color:black;}"
$between = $between + "</style>"

$after = "<style>"
$after = $after + "TABLE{border-width: 1px;border-style: solid;border-color:black;}"
$after = $after + "Table{background-color:#DFFFFF;border-collapse: collapse;}"
$after = $after + "TH{border-width:1px;padding:5px;border-style:solid;border-color:black;background-color:grey;color:white}"
$after = $after + "TD{border-width:1px;padding-left:5px;border-style:solid;border-color:black;}"
$after = $after + "</style>"

#####################################################################################################################################################
#########################################################   SCRIPT SECTION   ########################################################################
#####################################################################################################################################################

$reportDate = (Get-Date -Format "dd-MM-yyyy @ hh:mm:ss" )
#"----------------------------------------------------[$reportDate]-----------------------------------------------------------------------"| LogMe -display -progress


#intializing arrays to store data
$betCurrAndThresholdData =@()
$greaterThanThresholdData =@()
$lessThanCurrData =@()
$allData =@()

#"Removing Old HTML report" | LogMe -display -progress
& $logentrybot "Removing Old HTML report" -LogFile $logfilepath

try {

$ErrorActionPreference="continue"
#removing old HTML Report
if(test-path $html_output)
{
Remove-Item -path $html_output
& $logentrybot "Removed Old HTML report" -LogFile $logfilepath
}}

catch{
Write-Error " Old HTML report not found" $_.
& $logentrybot "Error while removing html file" -LogFile $logfilepath
} 

#Number of days to look for expiring date

#$threshold = 60 #Read-Host "enter threshold value[in days]"  
#"Threshold is set to $threshold days" | LogMe -display -progress
& $logentrybot "Threshold is set to $threshold days" -LogFile $logfilepath
#getting current date
$currDate = get-date
#based on threshold getting deadline date
$deadline = (Get-Date).AddDays($threshold)   #Set deadline date


#"Fetching Data having dates older than current date" | LogMe -display -progress
& $logentrybot "Fetching Data having dates older than current date" -LogFile $logfilepath
#fetching all data which comes before CurrentDate i.e expire App/Certificate/*
$lessThanCurr = $file | Where-Object {($_.Date -as [datetime] -lt $currDate) } | Select-Object 

try
{

$ErrorActionPreference="stop"
# fetching each entry & storing it in report
if($lessThanCurr -eq $null)
{
#"No data existing before current date[$currDate]- No Expired Data." | LogMe -warning

& $logentrybot "No data existing before current date[$currDate]- No Expired Data." -LogFile $logfilepath

$lessThanCurrData = "  "
$lessThanCurrData = $lessThanCurrData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, DaysCount -title "Expiry Information" -head $between -As Table -pre "<b>Before Current Date [Expired].</b>" -PostContent "<br>No data existing before current date[$currDate]- No Expired Data.<br><br>"
}
else
{
#"Creating HTML report for data older than current date" | LogMe -display -progress
& $logentrybot "Creating HTML report for data older than current date" -LogFile $logfilepath

Foreach($entry in $lessThanCurr)
{
$otherEntries = $entry | Select-Object * -ExcludeProperty Date
$name = $entry.Name
[datetime]$d2 = $entry.Date
[datetime]$d1 = $currDate
$ts = New-TimeSpan -Start $d1 -End $d2

$lessThanCurrData += ""| Select-Object @{Label="OtherColumns";Expression={"$otherEntries"}},@{Label="ExpiryDate";Expression={$entry.Date}}, @{Label="DaysCount";Expression={$ts.Days}}
}
$lessThanCurrData = $lessThanCurrData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, DaysCount -title "Expiry Information" -head $before -As Table -pre "<b>Before Current Date [Expired].</b>" -PostContent "<br>"
}
}

catch{

Write-Error "Error" $_
}

#fetching all data which comes between CurrentDate & Deadline Date  i.e to be expired before deadline App/Certificate/*
$betCurrAndThreshold = $file | Where-Object {($_.Date -as [datetime] -le $deadline) -and ($_.Date -as [datetime] -ge $currDate) } | Select-Object 

try
{

$ErrorActionPreference="stop"
# fetching each entry & storing it in report
if($betCurrAndThreshold -eq $null)
{
#"No data existing between current date[$currDate] and threshold date[$deadline]- No 'About to expire' Data." | LogMe -warning
& $logentrybot "No data existing between current date[$currDate] and threshold date[$deadline]- No 'About to expire' Data." -LogFile $logfilepath
$betCurrAndThresholdData = "  "
$betCurrAndThresholdData = $betCurrAndThresholdData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, DaysCount -title "Expiry Information" -head $between -As Table -pre "<b>Between Current Date & Threshold [About to Expire].</b>" -PostContent "<br>No data existing between current date[$currDate] and threshold date[$deadline]- No 'About to expire' Data.<br><br>"
}
else
{
#"Creating HTML report for data between current date & threshold date" | LogMe -display -progress
& $logentrybot "Creating HTML report for data between current date & threshold date" -LogFile $logfilepath
Foreach($entry1 in $betCurrAndThreshold)
{
$name1 = $entry1.Name
$otherEntries = $entry1 | Select-Object * -ExcludeProperty Date
$d1 = $entry1.Date
$d2 = $deadline
$ts = New-TimeSpan -Start $d1 -End $d2

$betCurrAndThresholdData
$betCurrAndThresholdData += ""| Select-Object @{Label="OtherColumns";Expression={"$otherEntries"}},@{Label="ExpiryDate";Expression={$entry1.Date}}, @{Label="DaysCount";Expression={$ts.Days}}
}
$betCurrAndThresholdData = $betCurrAndThresholdData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, DaysCount -title "Expiry Information" -head $between -As Table -pre "<b>Between Current Date & Threshold [About to Expire].</b>" -PostContent "<br>"
}
}

catch{
Write-Error "Error" $_

}

#fetching all data which comes after Deadline Date i.e data which is still yet to fall below deadline threshold
$greaterThanThreshold = $file | Where-Object {($_.Date -as [datetime] -gt $deadline) } | Select-Object 

try
{

$ErrorActionPreference="stop"
if($greaterThanThreshold -eq $null)
{
#"No data existing after threshold date[$deadline]." | LogMe -warning
& $logentrybot "No data existing after threshold date[$deadline]." -LogFile $logfilepath
$betCurrAndThresholdData = "  "
$greaterThanThresholdData = $greaterThanThresholdData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, DaysCount -title "Expiry Information" -head $between -As Table -pre "<b>After Threshold [Expiry after threshold].</b>" -PostContent "<br>No data existing after threshold date[$deadline].<br><br>"
}
else
{
#"Creating HTML report for data after threshold date" | LogMe -display -progress
& $logentrybot "Creating HTML report for data after threshold date" -LogFile $logfilepath
#fetching each entry & storing it in report
Foreach($entry2 in $greaterThanThreshold)
{
$name2 = $entry2.Name
$otherEntries = $entry2 | Select-Object * -ExcludeProperty Date
$d1 = $deadline
$d2 = $entry2.Date
$ts = New-TimeSpan -Start $d1 -End $d2

$greaterThanThresholdData += ""| Select-Object @{Label="OtherColumns";Expression={"$otherEntries"}},@{Label="ExpiryDate";Expression={$entry2.Date}}, @{Label="DaysCount";Expression={'+'+$ts.Days}}
}
$greaterThanThresholdData = $greaterThanThresholdData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, DaysCount -title "Expiry Information" -head $after -As Table -pre "<b>After Threshold [Expiry after threshold].</b>" -PostContent "<br>Done! $(Get-Date)" 
}

}

catch
{

Write-Error "Error" $_
}


$allData += $lessThanCurrData

$allData += $betCurrAndThresholdData

$allData += $greaterThanThresholdData

#"Creating HTML File of all records" | LogMe -progress

$allData |` Out-File $html_output -Force

#$betCurrAndThresholdData | out-file $html_output -Force

#"" | LogMe

& $logentrybot "" -LogFile $logfilepath 
& $logentrybot "-----------------------------------------------END--------------------------------------------------------------" -LogFile $logfilepath

#"-----------------------------------------------END--------------------------------------------------------------" |LogMe -display -progress

<#try{

$ErrorActionPreference ="stop"

Send-MailMessage -From $fromEmail -to "$sendemail" -Cc "$cc" -Subject "Certificate expiry details" -SmtpServer $smtpServer -attachments $html_output
}

catch {

Write-Error "Error in sending email" $_
} #>

