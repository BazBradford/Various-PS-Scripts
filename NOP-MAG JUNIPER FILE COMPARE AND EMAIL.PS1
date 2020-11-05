     #Variables
$Hashstore = "d:\baseline.txt"
$HashCompare = "d:\hashcompare.txt"
$HashTemp = "d:\hashtemp.txt"
$FileDir = "d:\New2"
$DateTime = Get-Date -format M.d.yyyy.hh.mm.ss

     #Email Variables
$smtp_server = 'NOPPEREXCH01.NOPSA.GOV.AU'
$to_email = 'information@nopsema.gov.au'
$from_email = 'noreply@nopsema.gov.au'
$dns_server = "<yourExternalDNSServer>"
$domain = "nopsema.gov.au"

#Check if Baseline.txt Exists
If (Test-Path $Hashstore)
  # // File exists
{}
Else {
  # // File does not exist - Should never happen!
$RefreshHash = dir $FileDir | Get-FileHash -Algorithm MD5
$RefreshHash | Out-File $Hashstore
  }

# Generate new Compare Hash.txt
$HashNew = dir $FileDir -Recurse | Get-FileHash -Algorithm MD5
$HashNew | Out-File $HashCompare

# Get Hash of baseline.txt
$HashBaseline = Get-FileHash -Path d:\baseline.txt -Algorithm MD5

#Get Hash of hashcompare.txt
$HashDiff = Get-FileHash -Path d:\hashcompare.txt -Algorithm MD5

#If changed, output hash to storage, and flag changes
If ($HashBaseline.hash -eq $HashDiff.hash)
{
    Add-Content -Path d:\success.$DateTime.txt -Value " Source Files ARE EQUAL </p>"
    }
else 
{
    Add-Content -Path d:\failure.$DateTime.html -Value "Source Files NOT EQUAL </p>"
    $HashNew | Out-File $HashTemp
}

# Compare two logs, send email if there is a change

If ($diff_results)
{
    #$evt_message = Get-Content .\domain.new.txt | Out-String
    #Write-EventLog -LogName Application -EventId 9000 -EntryType Error -Source "Maximo Validation Script" -Message $evt_message
    #Send-MailMessage -To $to_email -From $from_email -SmtpServer $smtp_server -Attachments .\domain.new.txt -Subject "ALERT! Change in Records" -Body "A change has been detected in the JUNIPER CONFIGURATION.`n`n`tACTION REQUIRED!`n`nVerify that this change was authorized."

    }

If ($HashNew.HashString -eq $Hashstore.HashString)
{
    }
else 
{
    $HashTemp | Out-File $HashStore
}