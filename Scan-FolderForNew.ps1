<#
.SYNOPSIS
	Script to check if files have been updated or added in the past 10 minutes.
.DESCRIPTION
	Script to check if files have been updated or added in the past 10 minutes, with the expectation
	that the script will be run as a scheduled task every 10 minutes.  If there is a hit the script
	will send an email with the changed/added files in the body of the email.
	
	Make sure to edit the script and update the default Param values to match your environment.  Then
	store the script in a central location and run it from multiple servers (if you want) using 
	the parameters to alter the script behavior.
.PARAMETER Path
	Full path name of the folder you wish to monitor.
.PARAMETER SMTPServer
	Name or IP address of your SMTP relay server.  This server must be setup to allow email relay.
.PARAMETER From
	Designate what email address you wish the update notice to come from.
.PARAMETER To
	Designate what email address you wish the update notice to be sent to.
.PARAMETER Subject
	Simple text for the subject of the update notice email.  This text will be updated by the
	string to add the $Path results as well.
.EXAMPLE
	.\Scan-FolderForNew.ps1
	Will run the script with full defaults.  
.EXAMPLE
	.\Scan-FolderForNew.ps1 -path c:\newpath -From newpath@domain.com -Subject "Newpath scan new files"
	Will run the path and scan the c:\newpath path, and if there is an add/udpate send an email to the
	default user from newpath@domain.com with a subject of "Newpath scan new files".
.LINK
	Inspiration:		http://community.spiceworks.com/topic/257320-need-some-help-with-a-script
	Source:		http://community.spiceworks.com/scripts/show/1595-scan-folder-for-changed-added-files
	Blog:		http://thesurlyadmin.com/2012/09/11/scan-a-folder-for-changes/
#>
Param (
	[string]$Path = "C:\Temp\",
	[string]$SMTPServer = "pdcexc04",
	[string]$From = "noreply@empired.com",
	[string]$To = "barry.bradford@empired.com",
	[string]$Subject = "New document(s) detected"
	)

$SMTPMessage = @{
    To = $To
    From = $From
	Subject = "$Subject on the Z:\Temp network drive"
    Smtpserver = $SMTPServer
}

$File = Get-ChildItem -recurse $Path | Where { $_.CreateTime -ge [datetime]::Now.AddMinutes(-2) }
If ($File)
{	$SMTPBody = "`nThe following files have recently been added/changed:`n`n"
	$File | ForEach { $SMTPBody += "$($_.FullName)`n" }
	Send-MailMessage @SMTPMessage -Body $SMTPBody
	
}