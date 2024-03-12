#####################################################################################################################################################
#Script Name : 
#Script Version: Version 0.1
#
#
#Script Description : Script to read Expiry date from Excel and determine near to expiry certificate
#
#Pre-Requisites:1) Csv file should contain one Column with name "Date" which will have all dates[in same date format]
#                  Date Format is "dd-MM-yyyy @ hh:mm:ss"
#              
#               
#
#
#####################################################################################################################################################
#########################################################   INPUT SECTION   #########################################################################
#####################################################################################################################################################
#Assigning relative path

$scriptParentPath = (Resolve-Path "..\..\").Path

$ConfPath =  Join-Path -path $scriptParentPath -ChildPath "/conf"
$libPath =  Join-Path -path $scriptParentPath -ChildPath "/lib"
$logfilepath =  Join-Path -path $scriptParentPath -ChildPath "/log/log.txt"
$masterBotPath =  Join-Path -path $scriptParentPath -ChildPath "/scripts/MasterBot"
$logentrybot =  Join-Path -path $scriptParentPath -ChildPath "/scripts/MicroBots/logentrybot.ps1"
$jsonConfig = Get-content -Path "$confPath\Config\config.json" |  ConvertFrom-Json
$html_output = "$scriptParentPath/output/ReadExcelExpiryDate-Report.html"
$file = "$scriptParentPath/inputs/test.xlsx" | Import-Excel

#Taking input from json file

$smtpServer = $jsonConfig.smtpServer
$fromEmail = $jsonConfig.fromEmail
$sendEmail = $jsonConfig.sendEmail
$threshold = $jsonConfig.threshold
$cc = $jsonConfig.cc

#Setting up CSS formatting
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

#Initializing date format as dd-MM-yyyy @ hh:mm:ss"
$reportDate = (Get-Date -Format "dd-MM-yyyy @ hh:mm:ss" )
#"----------------------------------------------------[$reportDate]-----------------------------------------------------------------------"| LogMe -display -progress


#Intializing arrays to store data
$betCurrAndThresholdData =@()
$greaterThanThresholdData =@()
$lessThanCurrData =@()
$allData =@()

#"Removing Old HTML report" | LogMe -display -progress


try 
{
$ErrorActionPreference="continue"
#removing old HTML Report
& $logentrybot "Removing Old HTML report" -LogFile $logfilepath
    if(test-path $html_output)
    {
      Remove-Item -path $html_output
      & $logentrybot "Removed Old HTML report" -LogFile $logfilepath
    }  
}

catch
    {
       Write-Error "Old HTML report not found $_."
       & $logentrybot "Error while removing html file" -LogFile $logfilepath
    } 

#Number of days to look for expiring date

& $logentrybot "Threshold is set to $threshold days" -LogFile $logfilepath
#Getting current date
$currDate = Get-Date 
#Based on threshold getting deadline date
$deadline = (Get-Date).AddDays($threshold) 


#Fetching expired data
& $logentrybot "Fetching expired data" -LogFile $logfilepath
$lessThanCurr = $file | Where-Object {($_.Date  -as [datetime]  -lt $currDate) } | Select-Object 

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
               $d2 = $entry.Date
               $d1 = $currDate
               $ts = New-TimeSpan -Start $d1 -End $d2


               $lessThanCurrData += ""| Select-Object @{Label="OtherColumns";Expression={"$otherEntries"}},@{Label="ExpiryDate";Expression={$entry.Date}}, @{Label="ElapsedDate";Expression={$ts.Days}}
              }
               $lessThanCurrData = $lessThanCurrData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, ElapsedDate -title "Expiry Information" -head $before -As Table -pre "<b>Before Current Date [Expired].</b>" -PostContent "<br>"
        }
   }

catch
    {

        Write-Error "Error in working with less than current date $_"
    }

#fetching all data which comes between CurrentDate & Deadline Date  i.e to be expired before deadline App/Certificate/*

$betCurrAndThreshold = $file | Where-Object {($_.Date -as [datetime] -le $deadline) -and ($_.Date -ge $currDate) } | Select-Object # removed -as [datetime]

try
{

      $ErrorActionPreference="stop"
      #Fetching each entry & storing it in report
        if($betCurrAndThreshold -eq $null)
            {
               & $logentrybot "No data existing between current date[$currDate] and threshold date[$deadline]- No 'About to expire' Data." -LogFile $logfilepath
               $betCurrAndThresholdData = "  "
               $betCurrAndThresholdData = $betCurrAndThresholdData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, DaysCount -title "Expiry Information" -head $between -As Table -pre "<b>Between Current Date & Threshold [About to Expire].</b>" -PostContent "<br>No data existing between current date[$currDate] and threshold date[$deadline]- No 'About to expire' Data.<br><br>"
            }
else
    {
      & $logentrybot "Creating HTML report for data between current date & threshold date" -LogFile $logfilepath
      Foreach($entry1 in $betCurrAndThreshold)
        {
            $name1 = $entry1.Name
            $otherEntries = $entry1 | Select-Object * -ExcludeProperty Date
            $d1 = $entry1.Date
            $d2 = $deadline
            $ts = New-TimeSpan -Start $d1 -End $d2
            #$ts = (60 - $ts.Days)
            #$betCurrAndThresholdData += ""| Select-Object @{Label="OtherColumns";Expression={"$otherEntries"}},@{Label="ExpiryDate";Expression={$entry1.Date}}, @{Label="About to Expire within";Expression={$ts.Days}}
            $betCurrAndThresholdData += ""| Select-Object @{Label="OtherColumns";Expression={"$otherEntries"}},@{Label="ExpiryDate";Expression={$entry1.Date}}, @{Label="About to Expire within";Expression={(60 - $ts.Days)}}
        }
            $betCurrAndThresholdData = $betCurrAndThresholdData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, "About to Expire within" -title "Expiry Information" -head $between -As Table -pre "<b>Between Current Date & Threshold [About to Expire].</b>" -PostContent "<br>"
    }
}

catch
    {
     Write-Error "Error in working with between current date & threshold date $_"
    }

#Fetching all data which comes after Deadline Date i.e data which is still yet to fall below deadline threshold
$greaterThanThreshold = $file | Where-Object {($_.Date -as [datetime] -gt $deadline) } | Select-Object #Removed -as [datetime]

try
{
    $ErrorActionPreference="stop"
    if($greaterThanThreshold -eq $null)
      {
       & $logentrybot "No data existing after threshold date[$deadline]." -LogFile $logfilepath
       $betCurrAndThresholdData = "  "
       $greaterThanThresholdData = $greaterThanThresholdData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, DaysCount -title "Expiry Information" -head $between -As Table -pre "<b>After Threshold [Expiry after threshold].</b>" -PostContent "<br>No data existing after threshold date[$deadline].<br><br>"
      }
    else
    {
         & $logentrybot "Creating HTML report for data after threshold date" -LogFile $logfilepath
         #fetching each entry & storing it in report
         Foreach($entry2 in $greaterThanThreshold)
           {
             $name2 = $entry2.Name
             $otherEntries = $entry2 | Select-Object * -ExcludeProperty Date
             $d1 = $deadline
             $d2 = $entry2.Date
             $ts = New-TimeSpan -Start $d1 -End $d2

             $greaterThanThresholdData += ""| Select-Object @{Label="OtherColumns";Expression={"$otherEntries"}},@{Label="ExpiryDate";Expression={$entry2.Date}}, @{Label="Above Deadline";Expression={'+'+$ts.Days}}
           }
             $greaterThanThresholdData = $greaterThanThresholdData | ConvertTo-Html -Property OtherColumns, Name, ExpiryDate, 'Above Deadline' -title "Expiry Information" -head $after -As Table -pre "<b>After Threshold [Expiry after threshold].</b>" -PostContent "<br>Done! $(Get-Date)" 
    }

}

catch
   {
     Write-Error "Error $_"
   }

#$allData += $betCurrAndThresholdData
#$allData += $greaterThanThresholdData
#$allData += $lessThanCurrData
$allData += $lessThanCurrData

$allData += $betCurrAndThresholdData

$allData += $greaterThanThresholdData

#"Creating HTML File of all records" | LogMe -progress

$allData | ` Out-File $html_output -Force

#$betCurrAndThresholdData | out-file $html_output -Force

#"" | LogMe

& $logentrybot "" -LogFile $logfilepath 
& $logentrybot "-----------------------------------------------END--------------------------------------------------------------" -LogFile $logfilepath

#"-----------------------------------------------END--------------------------------------------------------------" |LogMe -display -progress

try{

$ErrorActionPreference ="stop"

#$password = 'hvby oflp uboo nhso'
$password = Get-Content -Path "$ConfPath\Creds\securestring.txt" | ConvertTo-SecureString
 
#[SecureString]$securepassword = $password | ConvertTo-SecureString -AsPlainText -Force
#$credential = New-Object System.Management.Automation.PSCredential -ArgumentList $fromEmail, $securepassword
$credential = New-Object System.Management.Automation.PSCredential -ArgumentList $fromEmail, $password

Send-MailMessage -SmtpServer $smtpServer -Port 587 -UseSsl -cc $cc -From $fromEmail -To $sendemail -Subject "Certificate expiry details" -Body 'Test message' -Credential $credential -Attachments $html_output

#Send-MailMessage -From $fromEmail -to "$sendemail" -Cc "$cc" -Subject "Certificate expiry details" -SmtpServer $smtpServer -attachments $html_output
}

catch {

Write-Error "Error in sending email $_"
} 

