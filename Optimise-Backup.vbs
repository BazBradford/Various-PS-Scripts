Option Explicit
' Objective backup script: Stops relevant objective and related services,  cleans up logs, backs up database and files, then restarts services
' Note that the diles in Objective document store volume 1 are only backed up on the first of the month or if the parameter ForceBackupVolume1 is passed to the script.

' LIVE CONFIG
const LOG_FILE = "e:\scripts\logs\Objective-Optimise-Backup.log"
const WAIT_TIME = 600
const SMTP_SERVER = "nopperexc01"
const DB_SERVER = "NOPPERDBS04"
const LOG_HEADER = "-- Starting Objective Optimisation and Backup Script --"
const LOG_FOOTER = "---- End Objective Optimisation and backup Script -----"
const BACKUP_SOURCE_OBJVOL1 = "\\nopperapp28\e$\Objective\volume1"
const BACKUP_SOURCE_OBJVOL2 = "\\nopperapp28\g$\Objective\volume2"
const BACKUP_SOURCE_OBJVOL3 = "\\nopperapp28\f$\Objective\volume3"
const BACKUP_DEST_OBJVOL1 = "E:\Backup_Staging\volume1"
const BACKUP_DEST_OBJVOL2 = "E:\Backup_Staging\volume2"
const BACKUP_DEST_OBJVOL3 = "E:\Backup_Staging\volume3"
const BACKUP_SOURCE_CLOUDVIEW = "\\NOPPERAPP29\E$\Objective\ObjSearch\"
const BACKUP_DEST_CLOUDVIEW = "E:\Backup_Staging\ObjSearch"

const BACKUP_LOG = "E:\Backup_Staging\logs\robocopy.log"
const MAINTENANCE_PLAN = """D:\Program Files\Microsoft SQL Server\100\DTS\Binn\DTEXEC.exe"" /SQL ""Maintenance Plans\Full Database Backups"" /SERVER NOPPERDBS04  /MAXCONCURRENT "" -1 "" /CHECKPOINTING OFF /SET ""\Package\Subplan_1.Disable"";false /REPORTING E"
const SENDER = "svc-objbackup@nopsa.gov.au"
const RECIPIENT = "ObjectiveOptimisationNotification@nopsema.gov.au"

Rem ' ZETTA DEV CONFIG
Rem const LOG_FILE = "C:\Dev\play\objective\logs\Objective-Optimise-Backup.log"
Rem const WAIT_TIME = 600
Rem const SMTP_SERVER = "exchange.zettaserve.com"
Rem const DB_SERVER = "NOPPERDBS04"
Rem const LOG_HEADER = "-- Starting Objective Optimisation and Backup Script --"
Rem const LOG_FOOTER = "---- End Objective Optimisation and backup Script -----"
Rem const BACKUP_SOURCE_OBJVOL1 = "\\zettaserve.com\shares\Temp\keith\objective\volume1"
Rem const BACKUP_SOURCE_OBJVOL2 = "\\zettaserve.com\shares\Temp\keith\objective\volume2"
Rem const BACKUP_SOURCE_OBJVOL3 = "\\zettaserve.com\shares\Temp\keith\objective\volume3"
Rem const BACKUP_DEST_OBJVOL1 = "C:\Dev\play\objective\Backup_Staging\volume1"
Rem const BACKUP_DEST_OBJVOL2 = "C:\Dev\play\objective\Backup_Staging\volume2"
Rem const BACKUP_DEST_OBJVOL3 = "C:\Dev\play\objective\Backup_Staging\volume3"
Rem const BACKUP_SOURCE_CLOUDVIEW = "\\zettaserve.com\shares\Temp\keith\objective\ObjSearch\"
Rem const BACKUP_DEST_CLOUDVIEW = "C:\Dev\play\objective\Backup_Staging\ObjSearch"

Rem const BACKUP_LOG = "C:\Dev\play\objective\logs\robocopy.log"
Rem const MAINTENANCE_PLAN = """C:\Program Files\Microsoft SQL Server\100\DTS\Binn\DTEXEC.exe"" /SQL ""Maintenance Plans\Objective Prod Daily DB"" /SERVER NOPPERDBS01  /MAXCONCURRENT "" -1 "" /CHECKPOINTING OFF /SET ""\Package\Subplan_1.Disable"";false /REPORTING E"
Rem const SENDER = "keith.williams@zetta.com.au"
Rem const RECIPIENT = "keith.williams@zetta.com.au"


' Objective Servers
const OBJSEARCHSVR = "NOPPERAPP29"
const OBJECMSVR = "NOPPERAPP28"
const ADLIBSVR = "NOPPERAPP30"
const CONNECTLINKSVR = "NOPPERAPP20"

const UTILITIES = "c:\Objective\Utilities"

Dim objArgs, intArgCount
Dim bForceBackupVolume1: bForceBackupVolume1 = False
' Gather Arguments
Set objArgs = WScript.Arguments
For intArgCount = 0 to objArgs.Count - 1
	Select Case LCase(objArgs(intArgCount)) 
		Case "forcebackupvolume1"
			bForceBackupVolume1 = True
	End Select
Next


Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
Dim sEmail : sEmail = "Objective Maintenance started at: " & Now()
Dim iMainError : iMainError = -1


Sub WriteMsg(sMessage)
	dim oLog
	Wscript.echo Now() & ": " & sMessage
	If(oFso.FileExists(LOG_FILE) = false ) then
		Set oLog = oFso.CreateTextFile(LOG_FILE)
		oLog.WriteLine Now() & ": " & sMessage
	Else
		Set oLog = oFso.OpenTextFile(LOG_FILE, 8)
		oLog.writeLine Now() & ": "& sMessage
	End If 
	oLog.close
	sEmail = sEmail & vblf & chr(10) & sMessage
End sub


Function GetScriptPath()
' Gets the current script path
	Dim ScriptPathArr
	ScriptPathArr=Split(WScript.ScriptFullName, "\")
	ScriptPathArr(UBound(ScriptPathArr))=""
	GetScriptPath = Join(ScriptPathArr, "\")
End Function


Function StartStopService(sServer, sService, bStart)
	Err.Clear
	On Error Resume Next
	' Return codes
	'  0 Success
	' -1 General Error
	' -2 Null query returned from WMI - Service does not exist
	' -3 No matching services - service does not exist
	' -4 Service found and running, but instruction to Service failed
	' -5 Timeout while waiting for service operation
	Dim sDesiredState
	Dim sUndesiredState
	Dim sAction
	If bStart Then
		sDesiredState = "Running"
		sUndesiredState = "Stopped"
		sAction = "Starting"
	Else
		sDesiredState = "Stopped"
		sUndesiredState = "Running"
		sAction = "Stopping"
	End If
	Dim iError : iError = -1
	Dim oWMI : Set oWMI = GetObject("WINMGMTS:{impersonationLevel=impersonate}//" & sServer & "/root/cimv2")
	Dim oQuery : Set oQuery = oWMI.ExecQuery("SELECT * FROM Win32_Service WHERE Name='"& sService &"'")
	Dim oService
	Dim iReturn
	If ( IsNull(oQuery) = False ) Then
		iError = -3
		For Each oService in oQuery
			iError = -1
			If oService.State = sUndesiredState Then
				If bStart Then
					iReturn = oService.StartService
				Else
					iReturn = oService.StopService
				End If
				If iReturn <> 0 Then
					iError = -4
				Else
					Dim iCount : iCount = 0
					Dim oServiceCheck : Set oServiceCheck = oService
					Dim oServiceEnum
					Dim oQueryCheck
					Do
						If oServiceCheck.State = sDesiredState Then
							iError = 0
							Exit Do
						End If
						Set oQueryCheck = oWMI.ExecQuery("SELECT * FROM Win32_Service WHERE Name='"& sService &"'")
						For Each oServiceEnum in oQueryCheck
							set oServiceCheck = oServiceEnum
						Next
						WScript.Sleep(1000)
						iCount = iCount + 1
						If iCount > WAIT_TIME Then
							iError = -5
							Exit Do
						End If
					Loop
				End If
			End If
			if oService.State = sDesiredState Then
				iError = 0
			End If
		Next
	Else
		iError = -2
	End If
	Select Case iError
	Case -1:
		WriteMsg "Error " & sAction & " Service " & sService & " on " & sServer & ": -1: General Error"
	Case -2:
		WriteMsg "Error " & sAction & " Service " & sService & " on " & sServer & ": -2: Null query returned from WMI - Service does not exist"
	Case -3:
		WriteMsg "Error " & sAction & " Service " & sService & " on " & sServer & ": -3: No matching services - service does not exist"
	Case -4:
		WriteMsg "Error " & sAction & " Service " & sService & " on " & sServer & ": -4: Service found and " & sUndesiredState & ", but operation failed"
	Case -5:
		WriteMsg "Error " & sAction & " Service " & sService & " on " & sServer & ": -5: Timeout while waiting for service operation"
	End Select
	StartStopService = iError
	If Err.Number <> 0 Then
		WriteMsg "Error " & sAction & " Service " & sService & " on " & sServer & ": " & Err.Number & " :" & Err.Description
		Err.Clear
		StartStopService = -10
	End If
	On Error Goto 0
End Function


Function SendMail(sSubject, sBody)
	Err.Clear
	On Error Resume Next
	Dim oEmail : Set oEmail = CreateObject("CDO.Message")
	oEmail.Subject = sSubject
	oEmail.TextBody = sBody
	oEmail.Sender = SENDER
	oEmail.To = RECIPIENT
	oEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_SERVER
	oEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	oEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	oEmail.Configuration.Fields.Update
	oEmail.Send
	If Err.Number <> 0 then
		WriteMsg "Error sending email: " & Err.Number & " " & Err.Description
		Err.Clear
	End If
	On Error Goto 0
End Function


Function Quote(sInput)
	If Instr(sInput, " ") Then
		Quote = """" & sInput & """"
	Else
		Quote = sInput
	End If
End Function


Function GetRandomString(len)   
	dim i, s   
	const startChr ="a", range = 26   
	Randomize   
	s = ""  
	for i = 0 to len-1     
		s = s + Chr(asc(startChr) + Rnd() * range )    
	next   
	GetRandomString = s   
End Function


Function GetNewFileName(Directory, Seed, Extension)
'Gets a new filename (will not overwrite existing files)
	Dim fsoLocal
	Dim NewFileName
	Dim Counter
	Counter = UCase(GetRandomString(8))
	Set fsoLocal = CreateObject("Scripting.FileSystemObject")
	GetNewFileName = Directory & Seed & "-" & Counter & Extension
	If fsoLocal.FileExists(GetNewFileName) Then
		Dim Success
		Success = False
		Do Until Success
			GetNewFileName = Directory & Seed & "-" & Counter & Extension
			If fsoLocal.FileExists(GetNewFileName) Then
			Else
				Success = True
			End If
			Counter = UCase(GetRandomString(8))
		Loop
	End If
End Function

' See end of script for return codes.
Function RoboBackup(source, dest)
	Err.Clear
	On Error Resume Next
	Dim oShell : Set oShell = CreateObject("WScript.Shell")
	'Dim sTemp : sTemp = GetNewFileName(StandardDir(oFso.GetSpecialFolder(1)), "Temp", ".tmp")
	'Dim sTempCmd : sTempCmd = Quote(sTemp)
	Dim sTempCmd : sTempCmd = Quote(BACKUP_LOG)
	Dim sRobo : sRobo = Quote(StandardDir(oFso.GetSpecialFolder(1)) & "ROBOCOPY.EXE")
	Dim sSource : sSource = Quote(source)
	Dim sDest : sDest = Quote(dest)
	'Dim iRet : iRet = oShell.Run(sRobo & " " & sSource & " " & sDest & " /MIR /LOG:" & sTempCmd, 1, True)
	Dim iRet : iRet = oShell.Run(sRobo & " " & sSource & " " & sDest & " /MIR /LOG+:" & sTempCmd, 1, True)
	'Dim oFile : Set oFile = oFso.OpenTextFile(sTemp, 1)
	'Do Until oFile.AtEndOfStream
	'	WriteMsg oFile.ReadLine
	'Loop
	'oFile.Close
	'oFso.DeleteFile sTemp, True
	If iRet = 0 OR iRet = 1 OR iRet = 3 Then
		WriteMsg "RoboCopy Success. Code: " & iRet
	Else
		WriteMsg "RoboCopy Backup Failed. Code: " & iRet
	End If
	RoboBackup = iRet
	If Err.Number <> 0 then
		WriteMsg "Error Backing up Objective Document Store: " & Err.Number & " " & Err.Description
		Err.Clear
		RoboBackup = -10
	End If
	On Error Goto 0
End Function

' See end of script for return codes.
Function KillExe(exename,server)
	Err.Clear
	On Error Resume Next
	Dim oShell : Set oShell = CreateObject("WScript.Shell")
	Dim iRet : iRet = oShell.Run("taskkill /S " & server & " /F /IM " & exename, 1, True)
	If iRet = 0 Then
		WriteMsg "Kill " & server & " " & exename & " Success. Code: " & iRet
	Else
		WriteMsg "Kill Failed. Code: " & iRet
	End If
	KillExe = iRet
	If Err.Number <> 0 then
		WriteMsg "Error with Kill (if already successfully stopped by service stopped, then valid to get an error when trying to kill process: " & Err.Number & " " & Err.Description
		Err.Clear
		KillExe = -10
	End If
	On Error Goto 0
End Function

' See end of script for return codes.
Function RunLogCleanup()
	Err.Clear
	On Error Resume Next	
	Dim sDTExec : sDTExec = Quote(StandardDir(GetScriptPath()) & "psexec.exe")
	Dim oShell : Set oShell = CreateObject("WScript.Shell")

	Dim iRet : iRet = oShell.Run(sDTExec & " \\" & OBJECMSVR & " -i -d cmd /c " & UTILITIES & "\logCleanup.bat", 1, True)
	
	If iRet = 0 Then
		WriteMsg "RunLogCleanup Success. Code: " & iRet
	Else
		WriteMsg "RunLogCleanup Failed, likely no files to delete which outputs Errors to screen for each file to delete - ok. Code: " & iRet
	End If
	RunLogCleanup = iRet
	If Err.Number <> 0 then
		WriteMsg "Error with RunLogCleanup: " & Err.Number & " " & Err.Description
		Err.Clear
		RunLogCleanup = -10
	End If
	On Error Goto 0
End Function


Function StandardDir(sDir)
    If StrComp(Left(StrReverse(sDir), 1), "\", 1) <> 0 Then
        StandardDir = sDir & "\"
    Else
        StandardDir = sDir
    End If
End Function


Function DBBackup
	Err.Clear
	On Error Resume Next
	Dim sDTExec : sDTExec = Quote(StandardDir(GetScriptPath()) & "psexec.exe")
	Dim oShell : Set oShell = CreateObject("WScript.Shell")

	'Accept EULA
	oShell.Run sDTExec & " /accepteula", 1, True
        

	Dim iRet : iRet = oShell.Run(sDTExec & " \\" & DB_SERVER & " " & MAINTENANCE_PLAN, 1, True)
	If iRet <> 0 Then
		WriteMsg "Database backup failed. Code: " & iRet
	End If
	DBBackup = iRet
	If Err.Number <> 0 then
		WriteMsg "Error backing up database: " & Err.Number & " " & Err.Description
		Err.Clear
		DBBackup = -10
	End If
	On Error Goto 0
End Function




Function WaitNSeconds(nseconds)
	Dim iCount : iCount = 0
	Do
		WScript.Sleep(1000)
		iCount = iCount + 1
		If iCount > nseconds Then
			Exit Do
		End If
	Loop
End Function


Function Main
	Dim iError, bBackupVolume1
		Dim bSuccess : bSuccess = True
	WriteMsg "Stopping Objective Services"

Dim Counter00
Counter00 = 5 

Do	
	WriteMsg "Kill Objective Connect Link process"
	iError = KillExe("java.exe", CONNECTLINKSVR )
	If iError = True Then
	WaitNSeconds(60)
	Counter00 = Counter00 - 1
	End If
	Loop While iError = True And counter > 0


Dim Counter
Counter = 5

Do
'Objective Application Server
	iError = StartStopService(OBJECMSVR, "JBAS61SVC", false)
	If iError = True Then
	WaitNSeconds(60)
	Counter = Counter - 1
	End If
	Loop While iError = True And counter > 0
	
	
'Objective Automated Email Capture
Dim Counter0
Counter0 = 5

Do
	iError = StartStopService(OBJECMSVR, "ObjectivePROD_AecS", false)
	If iError = True Then
	WaitNSeconds(60)
	Counter0 = Counter0 - 1
	End If
	Loop While iError = True And Counter0 > 0

Dim Counter1
Counter1 = 5

Do
	iError = StartStopService(OBJECMSVR, "ObjectivePROD", false)
	If iError = True Then
	WaitNSeconds(60)
	Counter1 = Counter - 1
	End If
	Loop While iError = True And Counter1 > 0

Dim Counter2
Counter2 = 5

Do	
	iError = StartStopService(ADLIBSVR, "Adlib Process Manager", false)
		If iError = True Then
	WaitNSeconds(60)
	Counter2 = Counter2 - 1
	End If
	Loop While iError = True And Counter2 > 0
	
	
Dim Counter3
Counter3 = 5

Do	
	iError = StartStopService(ADLIBSVR, "Adlib System Manager Service", false)
		If iError = True Then
	WaitNSeconds(60)
	Counter3 = Counter3 - 1
	End If
	Loop While iError = True And Counter3 > 0
		
Dim Counter4
Counter4 = 5

Do
	iError = StartStopService(ADLIBSVR, "Adlib FMR", false)
		If iError = True Then
	WaitNSeconds(60)
	Counter4 = Counter4 - 1
	End If
	Loop While iError = True And Counter4 > 0
	
	
Dim Counter5
Counter5 = 5

Do
	iError = StartStopService(ADLIBSVR, "Adlib Job Management Service", false)
		If iError = True Then
	WaitNSeconds(60)
	Counter5 = Counter5 - 1
	End If
	Loop While iError = True And Counter5 > 0
	
Dim Counter6
Counter6 = 5

Do
	iError = StartStopService(ADLIBSVR, "Adlib Management Console Service", false)
		If iError = True Then
	WaitNSeconds(60)
	Counter6 = Counter6 - 1
	End If
	Loop While iError = True And Counter6 > 0

	WriteMsg "Ensure Objective Application Server is stopped - Kill any remaining Objective Application server e.g. java.exe"

Dim Counter7
Counter7 = 5

Do
	iError = KillExe("java.exe", OBJECMSVR )
		If iError = True Then
	WaitNSeconds(60)	
		Counter7 = Counter7 - 1
		End If
		Loop While iError = True And Counter7 > 0
		
Dim Counter8
Counter8 = 10

Do	
	iError = StartStopService(OBJSEARCHSVR, "Exalead Cloudview - PROD", false)
	If iError = True Then  
	WaitNSeconds(60)
	Counter8 = Counter8 - 1
	End If
	Loop While iError = True And Counter8 > 0
	
	' Log Cleanup
	WriteMsg "Run Logcleanup.bat"
	iError = RunLogCleanup()
	If iError <> 0 Then
		bSuccess = False
	End If
	
	
	
	WriteMsg "Backing up Objective Production Database"
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	WriteMsg "Backing up Objective Production Database"
	iError = DBBackup()
	If iError <> 0 Then
		bSuccess = False
	End If


	'TODO: ZettaServe - Backup CloudIndexes e.g. BACKUP_SOURCE_TWO, BACKUP_DEST_TWO
	
	WriteMsg "Backing up Cloudview Indexes"
	iError = RoboBackup(BACKUP_SOURCE_CLOUDVIEW, BACKUP_DEST_CLOUDVIEW)
	If iError <> 0 AND iError <> 1 AND iError <> 2 AND iError <> 3 Then
		bSuccess = False
	End If
	
	'TODO: ZettaServe - add Other Objective Volumes to backup e.g. Volume 1 once per month, Volume 2-3 daily.
' 		e.g. 	
'	 	BACKUP_SOURCE = "\\nopperapp28\g$\Objective\volume2"
'		BACKUP_SOURCE_THREE = "\\nopperapp28\f$\Objective\volume3"
'		BACKUP_DEST = "E:\Backup_Staging\volume2"
' 		BACKUP_DEST_THREE = "E:\Backup_Staging\volume3"


	' Backup Volume 1 once per month
	If (Day(Date()) = 1) Or bForceBackupVolume1 Then
		bBackupVolume1 = True
	Else
		bBackupVolume1  = False
	End If

	If bBackupVolume1 Then
		WriteMsg "Backing up Objective Volume 1 "
		iError = RoboBackup(BACKUP_SOURCE_OBJVOL1, BACKUP_DEST_OBJVOL1)
		If iError <> 0 AND iError <> 1 AND iError <> 2 AND iError <> 3 Then
			bSuccess = False
		End If
	Else
		WriteMsg "Skipping backup of Objective Volume 1 "
	End If

	WriteMsg "Backing up Objective Volume 2 "
	iError = RoboBackup(BACKUP_SOURCE_OBJVOL2, BACKUP_DEST_OBJVOL2)
	If iError <> 0 AND iError <> 1 AND iError <> 2 AND iError <> 3 Then
		bSuccess = False
	End If
	WriteMsg "Backing up Objective Volume 3 "
	iError = RoboBackup(BACKUP_SOURCE_OBJVOL3, BACKUP_DEST_OBJVOL3)
	If iError <> 0 AND iError <> 1 AND iError <> 2 AND iError <> 3 Then
		bSuccess = False
	End If	

	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	WriteMsg "Starting Cloudview Service"
	iError = StartStopService(OBJSEARCHSVR, "Exalead Cloudview - PROD", true)
	If iError <> 0 Then
		Exit Function
	End If

	WriteMsg "Wait 1min for Cloudview server to start up."
	WaitNSeconds(60)	' Wait 1 min to start the other services as the above takes a little time to start up.

	WriteMsg "Resume Starting Objective Adlib Services"
	iError = StartStopService(ADLIBSVR, "Adlib Management Console Service", true)
	If iError <> 0 Then
		Exit Function
	End If
	WaitNSeconds(10)	' Wait to start the other services as the above takes a little time to start up.
	iError = StartStopService(ADLIBSVR, "Adlib Job Management Service", true)
	If iError <> 0 Then
		Exit Function
	End If
	WaitNSeconds(10)	' Wait to start the other services as the above takes a little time to start up.
	iError = StartStopService(ADLIBSVR, "Adlib FMR", true)
	If iError <> 0 Then
		Exit Function
	End If
	WaitNSeconds(10)	' Wait to start the other services as the above takes a little time to start up.
	iError = StartStopService(ADLIBSVR, "Adlib System Manager Service", true)
	If iError <> 0 Then
		Exit Function
	End If
	WaitNSeconds(10)	' Wait to start the other services as the above takes a little time to start up.
	iError = StartStopService(ADLIBSVR, "Adlib Process Manager", true)
	If iError <> 0 Then
		Exit Function
	End If
	WaitNSeconds(10)	' Wait to start the other services as the above takes a little time to start up.
	
	
	WriteMsg "Resume Starting Objective Services"
	iError = StartStopService(OBJECMSVR, "ObjectivePROD", true)
	If iError <> 0 Then
		Exit Function
	End If
	WaitNSeconds(90)	' Wait to start the other services as the above takes a little time to start up.
	
	'Start Objective AEC
	iError = StartStopService(OBJECMSVR, "ObjectivePROD_AecS", true)
	If iError <> 0 Then
		Exit Function
	End If
	WaitNSeconds(10)	' Wait to start the other services as the above takes a little time to start up.
	
	'Start Objective Application Service
	iError = StartStopService(OBJECMSVR, "JBAS61SVC", true)
	If iError <> 0 Then
		Exit Function
	End If
	WaitNSeconds(30)	' Wait to start the other services as the above takes a little time to start up.
	
	'Start Objective Connect Link Service
	iError = StartStopService(CONNECTLINKSVR, "CONNAGNTJBOSS", true)
	If iError <> 0 Then
		Exit Function
	End If
	
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	
	If bSuccess Then
		WriteMsg "Objective tasks = Success"
		iMainError = 0
	Else
		WriteMsg "Objective tasks = Error"
		iMainError = -1
	End If
End Function




'On Error Resume Next
WriteMsg LOG_HEADER
Main
If iMainError = 0 Then
	SendMail "Objective-Backup-Optimise script: Success", sEmail
Else
	SendMail "Objective-Backup-Optimise script: Failure", sEmail
End If
WriteMsg LOG_FOOTER
WriteMsg ""



'    0×10  16       Serious error. Robocopy did not copy any files.
'                   Either a usage error or an error due to insufficient access privileges
'                   on the source or destination directories.
'    0×08   8       Some files or directories could not be copied
'                   (copy errors occurred and the retry limit was exceeded).
'                   Check these errors further.
'    0×04   4       Some Mismatched files or directories were detected.
'                   Examine the output log. Some housekeeping may be needed.
'    0×02   2       Some Extra files or directories were detected.
'                   Examine the output log for details. 
'    0×01   1       One or more files were copied successfully (that is, new files have arrived).
'    0×00   0       No errors occurred, and no copying was done.
'                   The source and destination directory trees are completely synchronized. 