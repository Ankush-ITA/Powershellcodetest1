################ Creating folder path constant ################
# Folder path destination
$scriptParentPath = (Resolve-Path "..\..\").Path
$confPath =  Join-Path -path $scriptParentPath -ChildPath "/conf"
$libPath =  Join-Path -path $scriptParentPath -ChildPath "/lib"
$logFilePath =  Join-Path -path $scriptParentPath -ChildPath "/log/log.txt"
$masterBotPath =  Join-Path -path $scriptParentPath -ChildPath "/scripts/MasterBot"
$logentryBotPath =  Join-Path -path $scriptParentPath -ChildPath "/scripts/MicroBots/logentrybot.ps1"
$jsonConfig = Get-content -Path "$confPath\Config\config.json" |  ConvertFrom-Json
$csvFilePath = "$scriptParentPath/output/E3_User_Count.csv"
$credFolderPath = Join-Path $confPath -ChildPath "\Creds"
$outputFolder = Join-Path $scriptParentPath -ChildPath "\output"
#$password = Get-Content -Path "$credFolderPath\emailpassword.txt" | ConvertTo-SecureString
$netScalerBot = Join-Path -path $scriptParentPath -ChildPath "/scripts/MicroBots/Netscaler.ps1"
$rdpBot = Join-Path -path $scriptParentPath -ChildPath "/scripts/MicroBots/RDP.ps1"
$inputfilepath = Join-Path -Path $scriptParentPath -ChildPath "/input/URLlist1.txt"
$outputfilepath = Join-Path -Path $scriptParentPath -ChildPath "/output/netscaler.html"
$outputfilepath1 = Join-Path -Path $scriptParentPath -ChildPath "/output/netscalerall.html"
$outputfilepath1rdp = Join-Path -Path $scriptParentPath -ChildPath "/output/rdp.csv"
$serverListFilePath = Join-Path -Path $scriptParentPath -ChildPath "/input/serverlist1.txt"
$outputfilepathhtml = Join-Path -Path $scriptParentPath -ChildPath "/output/rdp.html"
$outputfilepathhtml1 = Join-Path -Path $scriptParentPath -ChildPath "/output/rdp1.html"

$ErrorActionPreference = "SilentlyContinue"

$microbotOutput = & "..\MicroBots\Netscaler.ps1" -inputfilepath $inputfilepath -outputfilepath $outputfilepath -outputfilepath1 $outputfilepath1
$microbotOutput1 = & "..\MicroBots\RDP.ps1" -serverListFilePath $serverListFilePath -outputfilepath1rdp $outputfilepath1rdp -outputfilepathhtml $outputfilepathhtml -outputfilepathhtml1 $outputfilepathhtml1

