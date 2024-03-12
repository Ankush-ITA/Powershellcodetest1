

$script = split-path -parent $MyInvocation.MyCommand.Definition
#$script

$script1= Join-Path -path $script -ChildPath "C:\Users\ankushku\OneDrive - Capgemini\Redeployement\BID283_Certificate-Expiry_Check\Scripts\MicroBots"

#$script1


$scrip3 = "$script1\Test1.txt"

$checking1 =Get-Content "$scrip3"