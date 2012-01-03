' ===========================================================================
' Wincron.vbs
' 2003-ish
' by Dale Anderson (http://www.daleanderson.ca/)
' See README file for full details.
' ===========================================================================


''' config ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

WScript.Quit Main()


'============================
'============================
Class cConfig

	Public EmailNotify_From
	Public EmailNotify_SmtpHost
	Public EmailNotify_Rcpt
	Public TCP_Machine_Name

	Private Sub Class_Initialize
		EmailNotify_From = "watchdog@your-domain-name.com"
		EmailNotify_SmtpHost = "localhost"
		EmailNotify_Rcpt = "errors@your-domain-name.com"
		TCP_Machine_Name = RegRead("HKLM\SYSTEM\CurrentControlSet\Services\TCPIP\Parameters\Hostname")
	End Sub
	
	Private Sub Class_Terminate
	End Sub
	
End Class




'-----------------------------------------------
'-----------------------------------------------
'-----------------------------------------------
'-----------------------------------------------
Function Main()
	Main = 0 
	Dim errmsg
	If Not IsCscript(errmsg) Then
		fwrite errmsg
		WScript.Echo ErrMsg
		Main = 97
		Exit Function
	End If	
	
	Dim oConfig
	Set oConfig = New cConfig

	Const namR1P = "\run-once-pending"
	Const namR1F = "\run-once-finished"
	Const namRE = "\recurring"
	Const namLogFile = "\WinCron.log"
	
	Dim oFso
	Set oFso = CreateObject("Scripting.FileSystemObject")

	Dim currentDir
	currentDir = oFso.GetParentFolderName(WScript.ScriptFullName)

	Call prepFolder(currentDir & namR1P)
	Call prepFolder(currentDir & namR1F)
	Call prepFolder(currentDir & namRE)

	Dim workingDir

	Dim iResult
	
	'''' Do Recurring tasks first.
	workingdir = currentDir & namRE
	Call runTasks(workingDir, False, oFso, oConfig) 'leave the files where they are when we are done.

	'''' now Do the one-time tasks.
	workingdir = currentDir & namR1P
	Call runTasks(workingDir, True, oFso, oConfig) 'move the files when we are done with them.
	

	Set oFso = nothing

	Set oConfig = Nothing
	
End Function









'--------------------------------
'2008-02-11
'--------------------------------
Sub DebugL(title, value)
	Dim s
	s = Title 
	s = s & ": ["
	on error resume next
	s = s & Value
	on error goto 0 
	s = s & "] ("
	s = s & TypeName(Value)
	s = s & ")"
 	Debug s
End Sub
'--------------------------------
'2008-02-11
'--------------------------------
Sub Debug(s)
 	Echo s
End Sub

'----------------------------------------
'----------------------------------------
Function Echo(s)
	WScript.Echo s
	fwrite s
End Function






'''''''' subs & functions '''''''''''''''''''''''''''''''''''

Sub runTasks(workingDir, RemoveWhenDone, oFso, oConfig)
	
	Dim today
	today = now()

	Dim yyyy
	yyyy = datepart("yyyy",today)

	Dim mm
	mm = zerofill(datepart("m",today),2)

	Dim dd
	dd = zerofill(datepart("d", today),2)

	Dim hh
	hh = zerofill(datepart("h", today),2)

	Dim nn
	nn = zerofill(datepart("n", today),2)

	Dim ss
	ss = zerofill(datepart("s", today),2)


	Dim oFolder
	Set oFolder = oFso.GetFolder(workingDir)

	Dim oFiles
	Set oFiles = oFolder.Files

	Dim iResult
	

	Dim oFile
	Dim filePath
	Dim fileName
	Dim extensionName
	Dim ErrMsg
	For Each oFile In oFiles
		filePath = oFile.path
		filename = oFile.name
		extensionName = lcase(oFso.GetExtensionName(filePath))
		ErrMsg = "" 'reset this for each time.

		Select Case extensionName

			Case "bat", "cmd", "exe" 
				iResult = Run(filePath, ErrMsg)

			Case "vbs", "js"
				iResult = Run("cscript.exe " & filePath, ErrMsg)

			Case Else
				iResult = 0 ' there is nothing to do for undefined file types.

		End Select
		

		If iResult <> 0 Then
			Call NotifyAdmin(oConfig, "iResult: " & iResult & vbNewLine & "Path: " & filePath & vbNewLine & "ErrMsg: " & ErrMsg)
		End If
		
		If RemoveWhenDone = true Then 
			' move Each of the files to the "finished" folder.
			oFso.moveFile filePath, currentDir & namR1F & "\" & yyyy & mm & dd & hh & nn & ss & "-" & filename
		End If

	Next

	Set oFiles = nothing

	Set oFolder = nothing

End Sub




'---------------------------------------------------------------
'---------------------------------------------------------------
Sub NotifyAdmin(oConfig, ByVal Message)
  Dim FileSystemObj, LogString
  Dim jmail, SentOK
  Set jmail = createobject("jmail.smtpmail")
  jmail.serveraddress = oConfig.EmailNotify_SmtpHost
  jmail.silent = True
  jmail.logging = True
  jmail.sender = oConfig.EmailNotify_From
  JMail.Subject = "\\" & oconfig.TCP_Machine_Name & "\" & WScript.ScriptFullName
  jmail.addrecipient oConfig.EmailNotify_Rcpt
  jmail.Body = Message
  JMail.ISOEncodeHeaders = False
  SentOK = jmail.execute
  LogString = "MailSent=" & sentOk & "; ReplyTo=" & JMail.Sender & "; RCPT=" & oconfig.EmailNotify_Rcpt
  if SentOK then 
    fwrite LogString
  else
    fwrite jmail.log
  end if
  set jmail = nothing
End Sub 





'----------------------------------------
' 2008-07-02, Dale C. Anderson
' Adapted from Christian d'Heureuse's Exec() function.
'----------------------------------------
Function Run(ByVal cmd, ErrMsg)
	Echo "Executing " & Cmd
	Dim sh
	Set sh = CreateObject("WScript.Shell")
	Dim wsx
	On Error Resume Next
	Set wsx = Sh.Exec(cmd)
	If Err.Number <> 0 Then
		Run = Err.Number
		ErrMsg = Err.Description
		Echo ErrMsg
		On Error GoTo 0
		Set sh = Nothing
		Exit Function
	End If
	On Error GoTo 0
	
	If wsx.ProcessID = 0 And wsx.Status = 1 Then
		' (The Win98 version of VBScript does not detect WshShell.Exec errors)
		ErrMsg = "WshShell.Exec failed. Are you trying to run this on a non-NT system?"
		Run = 2750
		Exit Function 
	End If


	' Status and Error are two different things. 
	' Error 0 means A-OK. Anything else means something went wrong,
	
	Dim iStatus
	iStatus = 0

	Dim iError
	iError = 0 


	Dim sErr 
	
	Do

		' Status 0 means the porgram is still running.
		' Status 1 means the program is done.
		iStatus = wsx.Status     	

		fwrite wsx.StdOut.ReadAll()

		'Since runtime errors wont return any clear information, we have to check the stderror stream.
		sErr = wsx.StdErr.ReadAll()
		If Len(sErr) > 0 Then
			iError = 9200
		End If

		If iStatus <> 0 Then  'Checking for program completion
			Exit Do
		End If

		WScript.Sleep 100 'Dont want to kill the processor.

	Loop

	If iError = 0 Then
		iError = wsx.ExitCode  ' Only set the error number to the program's exit code if there was nothing in the STDERR stream.
	End If 

	Set wsx = Nothing
	Set sh = Nothing

	Run = iError

End Function



'

'********************************************************************
'* Function IsCscript() 
'* Purpose: Determines which program is used to run this script.
'* Input:   ErrMsg (referenced string)
'* Output:  boolean, ErrMsg
'  
'********************************************************************
Function IsCscript(errmsg)

	IsCscript = False
	
	On Error Resume Next
	Dim strFullName

	strFullName = WScript.FullName

	If Err.Number Then
		ErrMsg =  "IsCscript(), line 317, Error 0x" & CStr(Hex(Err.Number)) & ": " & Err.Description
		On Error GoTo 0
		Exit Function
	End If
	On Error GoTo 0

	Dim i
	i = InStr(1, strFullName, ".exe", 1)
	If i = 0 Then
		ErrMsg =  "IsCscript(), line 326, '.exe' was not found in WScript.FullName"
		Exit Function
	End If
	
	Dim j
	j = InStrRev(strFullName, "\", i, 1)
	If j = 0 Then
		ErrMsg =  "IsCscript(), line 326, '\' was not found in WScript.FullName"
		Exit Function
	End If

	Dim strCommand
	strCommand = Mid(strFullName, j+1, i-j-1)
	Select Case LCase(strCommand)
		Case "cscript"
			IsCscript = True 'YAY!
		Case "wscript"
			ErrMsg = "WScript.exe cannot be used to run this script. Must use CScript.exe instead."
		Case Else
			ErrMsg = "IsCscript(), line 326, An unexpected program was used to run this script. Only CScript.Exe or WScript.Exe can be used to run this script." 
	End Select

End Function


'----------------------------
'----------------------------
Function ZeroFill(MyString,LenShouldBe)
	Dim Length
	Length = Len(CStr(MyString))
	If Length < LenShouldBe Then
		Do While Length < LenShouldBe 
			MyString = "0" & MyString
			Length = Length + 1
		Loop
	End If
	ZeroFill = MyString
End Function

'----------------------------
'----------------------------
Function RegRead(ByVal What)
  Dim oSh
  Set oSh = CreateObject("WScript.Shell")
  RegRead = oSh.regRead(what)
  Set oSh = Nothing
End Function
  



















'#######################################################################
'
' =====[ ClsLogIt ]====================
' 
' Date:     2004-12-20: initial build.
'           2005 01 09: Newline characters are now optional
'           2005 01 09: Option to overwrite previous contents
'           2005-01-12: ScriptName(), CurrentDir() and PrepFolder() are now
'                       more easily accessible.
'           2005-01-24: SetLogFileName() externalized.
'
' Author:   Dale C Anderson http://www.daleanderson.ca/
'
' Benefits: (1) Enables INSTANT ability to start logging timestamped entries to a
'           log file, whether you're using ASP on the server, or a VBS file
'           on your desktop.
'
'           (2) ALL parts of this class will serve
'           you, whether you whether you are using ASP or VBS!!!!!!!!!!!!!
'           Which makes things very nice if you like to copy & paste
'           code between both worlds.
'
' Installation:
'
'           You need 3 things in your script to start using this: 
'           (1) global declaration of "gLogFile" 
'           (2) the logit class
'           (3) well this one isnt necessary, but its handy;  an 
'           "fwrite" function, or something similar makes this thing
'               VERY convenient to use.
'           
'           For asp use: make sure iusr_<machinename> has 
'           write / change permissions on the nt directories
'           that it needs to write to.
'
'
' Methods
' ------------------------------------
'     .Write(SomeTextToWrite)   
'
'     Accepts......:  String
'     Returns......:  Nothing
'     Description..:  Give it something to write, and it does just what it says. 
'                     Kind of the whole point of this excersize, no?
'
'
'     .PrepFolder(SomeDirectoryPath)   
'
'     Accepts......:  String
'     Returns......:  Nothing
'     Description..:  This sub will create a directory structure any number of levels deep.
'                     For instance, if you only have one directory on your computer (i.e. "C:\DOS"),
'                     you could pass the path "c:\windows\system32\inetsrv\iisadmin\etc\etc\etc" and 
'                     the sub would create that directory structure.
'
'
' Functions 
' ------------------------------------
'     .CurrentDir()  

'     Accepts......:  Nothing
'     Returns......:  String Value
'     Description..:  Returns the full path of the directory that the script 
'                     is running in.
'
'
'     .ScriptName()  
' 
'     Accepts......:  Nothing
'     Returns......:  String Value
'     Description..:  Returns only the name of the script you are running, without the path.
'
'
'
' Properties
' ------------------------------------
'     .LogfileName(YourFavoriteFilename)
'
'     Accepts......:  String value
'     Returns......:  String value
'     Mandatory....:  No
'     Description..:  If you dont specify this, logging will take place 
'                     in the same directory that the script is running in, 
'                     in a file that has the same name as the script, with 
'                     but with ".log" appended to the filename.
'
'
'  
'     .UseTimestamps(TrueOrFalse)
'
'     Accepts......:  Boolean
'     Returns......:  Boolean
'     Mandatory....:  No
'     Description..:  True by default. Causes timestamps plus a VbTab 
'                     to be inserted in front of each line that you log.
' 
	
	'_____________________________________________
	Dim gLogFile 'The only thing that must be declared outside of the class for everything to work.


	'_____________________________________________
	Sub fwrite(s) 
		Call InitClsLogit()
		' gLogFile.NewlineCharacter = "" ' default is VbCrLF.
		' gLogFile.LogfileName = "C:\path\to\gLogFile.txt" 	'Defaults to "%CurrentDir%\%ScriptName%.log"
		' gLogFile.UseTimestamps = False 				'default is true.
		gLogFile.OverwriteIfFileExists = True     ' default is False.
		gLogFile.Write s
	End Sub
	
	'__________________________________
	Function CurrentDir()
		Call InitClsLogit()
		CurrentDir = gLogFile.CurrentDir()
	End Function

	'__________________________________
	Function ScriptName()
		Call InitClsLogit()
		ScriptName = gLogFile.ScriptName()
	End Function

	'__________________________________
	Sub PrepFolder(ByVal strMyFolder)
		Call InitClsLogit()
		gLogFile.PrepFolder(strMyFolder)
	End Sub

	'__________________________________
	Sub SetLogFileName(ByVal strLogFileName)
		Call InitClsLogit()
		gLogFile.LogfileName = CurrentDir() & "\" & strLogFileName
	End Sub
	'__________________________________
	Sub InitClsLogit()
		If Not IsObject(gLogFile) Then
			Set gLogFile = New clsLogit
		End If
	End Sub

	'__________________________________
	Class clsLogIt 

		Private m_fso
		Private m_LogfileName
		Private m_bUseTimestamps
		Private m_TextStreamObject
		Private m_OverwriteIfFileExists
		Private m_NewlineCharacter
		'__________________________________
		Private Sub Class_Initialize
			Set m_fso = CreateObject("scripting.filesystemobject")
			m_LogfileName = CurrentDir & "\" & ScriptName & ".log" 'just a default
			m_bUseTimestamps = True
			m_OverwriteIfFileExists = False
			m_NewlineCharacter = VbCrLf
		End Sub
		'__________________________________
		Private Sub Class_Terminate()
			If IsObject(m_TextStreamObject) Then 
				m_TextStreamObject.close
			End If
			set m_TextStreamObject = Nothing
			Set m_fso = nothing
		End Sub
		'__________________________________
		Public Function CurrentDir
			Dim FullPath
			Dim Result
			If isWscript() Then 
				FullPath =  wscript.ScriptFullName
				Result = m_fso.GetParentFolderName(FullPath) 
			ElseIf IsAsp() Then 
				FullPath = Server.MapPath(Request.ServerVariables("PATH_INFO"))
				Result = m_fso.GetParentFolderName(FullPath) 
			Else
				Err.Raise 1, "clsLogIt.CurrentDir: Couldn't determine script engine."
			End If
			CurrentDir = Result
		End Function
		'__________________________________
		Public Function ScriptName
			Dim FullPath
			Dim Result
			If IsWscript() Then 
				FullPath =  wscript.ScriptFullName
				Result = m_fso.GetBaseName(FullPath) & Ext(FullPath)
			ElseIf IsAsp() Then 
				FullPath =  Request.ServerVariables("PATH_INFO")
				Result = m_fso.GetBaseName(FullPath) & Ext(FullPath)
			Else
				Err.Raise 1, "clsLogIt.ScriptName: Couldn't determine script engine."
			End If
			ScriptName = Result
		End Function
		'__________________________________
		Public Property Let NewlineCharacter(p_data)
			m_NewlineCharacter = p_data
		End Property
		'__________________________________
		Public Property Get NewlineCharacter()
			NewlineCharacter = m_NewlineCharacter
		End Property
		'__________________________________
		Public Property Let LogfileName(p_data)
			m_LogfileName = p_data
		End Property
		'__________________________________
		Public Property Get LogfileName()
			LogfileName = m_LogfileName
		End Property
		'__________________________________
		Public Property Let UseTimestamps(p_data)
			m_bUseTimestamps = CBool(p_data)
		End Property
		'__________________________________
		Public Property Get UseTimestamps()
			UseTimestamps = m_bUseTimestamps
		End Property
		'__________________________________
		Public Property Let OverwriteIfFileExists(p_data)
			m_OverwriteIfFileExists = CBool(p_data)
		End Property
		'__________________________________
		Public Property Get OverwriteIfFileExists()
			OverwriteIfFileExists = m_OverwriteIfFileExists
		End Property
		'__________________________________
		Public Sub Write(p_data)
			Const ForAppending = 8
			Const FileFormatAscii = 0
			Const CreateIfNotExists = true 
			If Not IsObject(m_TextStreamObject) Then
				Call prepfolder(m_fso.GetParentFolderName(m_LogfileName))
				If m_OverwriteIfFileExists Then 
					Set m_TextStreamObject = m_fso.CreateTextFile(m_LogFileName, m_OverwriteIfFileExists)
				Else
					Set m_TextStreamObject = m_fso.OpenTextFile(m_LogfileName, ForAppending, CreateIfNotExists, FileFormatAscii)
				End If  				
			End If
			If m_bUseTimestamps Then p_data = Now() & vbTab & p_data
			m_TextStreamObject.write m_NewlineCharacter & p_data
		End sub
		'__________________________________
		Public Sub PrepFolder(ByVal strMyFolder)
			Dim strParentFolder
			If Not m_fso.FolderExists(strMyFolder) Then
				 strParentFolder = m_fso.GetParentFolderName(strMyFolder)
				 Call PrepFolder(strParentFolder)
				 m_fso.CreateFolder (strMyFolder)
			End If
		End Sub
		'__________________________________
		Public function ext(byval fname)
			ext = lcase(mid(fname,inStrRev(fname,".")))
		end function 
		'__________________________________
		Public Function IsAsp()
			On Error Resume Next
			Dim result
			Result = False
			If IsObject(Session) _
			And IsObject(Application) _
			And IsObject(Request) _
			And IsObject(ObjectContext) _
			Then 
				Result = True
			End If
			If err.number <> 0 Then result = False
			On Error GoTo 0
			IsAsp = Result
		End Function
		'__________________________________
		Public Function IsWscript()
			On Error Resume Next
			Dim result
			Result = False
			If IsObject(Wscript) _
			And Not IsAsp() _
			Then 
				Result = True
			End If
			If err.number <> 0 Then result = False
			On Error GoTo 0
			IsWscript = Result
		End Function
		
	End Class 



'#######################################################################