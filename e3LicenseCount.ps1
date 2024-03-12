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
 

# define the secretkey file path
$secretkeyFilePath = Join-Path -Path $credFolderPath -ChildPath "secretkey.txt"
# define the vectorBytes file path
$vectorkeyFilePath = Join-Path -Path $credFolderPath -ChildPath "vectorkey.txt"
# Specify the path to the output text file
$licenseFilePath = Join-Path -Path $confPath -ChildPath "/License/license.txt"



################################ Decrypt Credential File ################################
function decryptkey {
    param (
        [string]$filePath
    )
    [Byte[]] $key = (1..16)
    $secureString = Get-Content $filePath | ConvertTo-SecureString -Key $key
    $decryptedText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString))
    return $decryptedText
}
######################## Checking the License File if it's valid or not ########################
$secretKey = decryptkey -filePath $secretkeyFilePath
$iv = decryptkey -filePath $vectorkeyFilePath
$password = decryptkey -filePath "$credFolderPath\emailpassword.txt"

$global:currentDate = Get-Date
$encryptionKey = [System.Text.Encoding]::UTF8.GetBytes($secretKey)
$ivBytes = [System.Text.Encoding]::UTF8.GetBytes($iv)
# Decrypt and Validate License from the file
$encryptedLicenseInfo = Get-Content -Path $licenseFilePath -Encoding Byte

$decryptedLicenseInfo = $null
try {
    $decryptedLicenseInfo = [System.Security.Cryptography.Aes]::Create() | ForEach-Object {
        $_.Key = $encryptionKey
        $_.IV = $ivBytes
        $_.Mode = [System.Security.Cryptography.CipherMode]::CBC  # Use CBC mode
        $_.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $_.CreateDecryptor().TransformFinalBlock($encryptedLicenseInfo, 0, $encryptedLicenseInfo.Length)
    }
    if ($decryptedLicenseInfo) {
        $decryptedLicenseInfo = [System.Text.Encoding]::UTF8.GetString($decryptedLicenseInfo)
        $licenseData = $decryptedLicenseInfo | ConvertFrom-Json
        if ([DateTime]::Parse($licenseData.StartDate) -le $global:currentDate -and [DateTime]::Parse($licenseData.ExpiryDate) -ge $currentDate) { 
            $global:daysRemaining = [math]::Ceiling(([DateTime]::Parse($licenseData.ExpiryDate) - $global:currentDate ).Totaldays)
            Write-Output "License is valid. Proceeding with the script execution."
            Write-Output "Current Date: $global:currentDate"
            Write-Output "Start Date: $($licenseData.StartDate)"
            Write-Output "Expiry Date: $($licenseData.ExpiryDate)"
            Write-Output "Days Remaining: $daysRemaining"
            & $logentryBotPath "SUCCESS - License Validation Complete!!" -LogFile $logFilePath
        }
        elseif ([DateTime]::Parse($licenseData.StartDate) -lt $global:currentDate) {
        } 
    }
    else {
        Write-Output "License has expired."
        & $logentryBotPath "ERROR - License Expired!!" -LogFile $logFilePath
        exit
    }
} 
catch {
    Write-Output "An error occurred while decrypting the license: $_append"
    & $logentryBotPath "ERROR - While Loading License File!!" -LogFile $logFilePath
    exit
}
########################################################################################
$jsonConfig = Get-content -Path "$confPath\Config\config.json" |  ConvertFrom-Json

#Taking input from json file

$smtpServer = $jsonConfig.smtpServer
$fromEmail = $jsonConfig.fromEmail
$sendEmail = $jsonConfig.sendEmail
$threshold = $jsonConfig.threshold
$cc = $jsonConfig.cc
$AccountSkuId = $jsonConfig.AccountSkuId
$emailSubject = $jsonConfig.emailSubject

#Creating connection with MsolService
try {
    $ErrorActionPreference = "Stop"
    $credential = Get-credential
    Connect-MsolService -Credential $credential
    & $logentryBotPath "SUCCESS -MsolService connected" -LogFile $logFilePath
}
catch {
    Write-Output "An error occurred while connecting to Msol service: $($_.Exception.Message)"
    & $logentryBotPath "An error occurred while connecting to Msol service" -LogFile $logFilePath
    exit
}

#Taking usesr list having E3 license
$e3LicenseUsers = Get-MsolUser -All | Where-Object {($_.Licenses | ForEach-Object { $_.AccountSkuId }) -contains "$AccountSkuId"}

# Count the number of E3 license users
$e3UserCount = $e3LicenseUsers.Count

# Export the user count to a CSV file
$attachment= $e3LicenseUsers | Select-Object UserPrincipalName | Export-Csv -Path $csvFilePath -NoTypeInformation

# Calculating total number of E3 license left
$e3LicenseCount = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -eq "your_e3_sku_id"} | Select-Object -ExpandProperty ActiveUnits,ConsumedUnits
$activeUnits = $e3LicenseCount.ActiveUnits
$consumedUnits = $e3LicenseCount.ConsumedUnits
$e3LicenseAvailable = $activeUnits-$consumedUnits

# Taking License available count in licenseLeft.txt file
$timestamp = Get-Date -Format "dd-MM-yy hh:mm:ss tt"
"Timestamp: $timestamp : Number of E3 license left:$e3LicenseAvailable" | Out-File -Append "$outputFolder\licenseLeft.txt"




# HTML email content
$emailBody = @"
<html>
<head>
<style>
  body { font-family: Arial, sans-serif; }
  h2 { color: #333; }
  p { margin: 1em 0; }
</style>
</head>
<body>
<h2>E3 License User Count Report</h2>
<p>Total E3 License User Count: $e3UserCount</p>
</body>
</html>
"@

if ($jsonConfig.isEmailRequired ="yes")
{
# Send an email
#Send-MailMessage -SmtpServer "your_smtp_server" -From $adminEmail -To $adminEmail -Subject $emailSubject -Body $emailBody -BodyAsHtml
$credential = New-Object System.Management.Automation.PSCredential -ArgumentList $fromEmail, $password
#Send-MailMessage -SmtpServer $smtpServer -Port 587 -UseSsl -cc $cc -From $fromEmail -To $sendemail -Subject "Certificate expiry details" -Body 'Test message' -Credential $credential -Attachments $html_output
Send-MailMessage -SmtpServer $smtpServer -Port 587 -UseSsl -From $fromEmail -To $sendemail -cc $cc -Subject "$emailSubject" -Body "$emailBody" -BodyAsHtml -Credential $credential -Attachments $attachment

}

else {
exit
}