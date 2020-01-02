
Dim o42: Set o42 = CreateObject("Scripting.FileSystemObject")
Dim shell42: Set shell42 = CreateObject("Wscript.Shell")
Dim errors42: errors42 = ""

'Load ScriptWise configuration
Dim swConfig42: swConfig42 = ScriptWisePath & "ScriptWise.config"
If Not (o42.FileExists(swConfig42)) Then
	LocalScriptDebug.MessageBox("Could not find ScriptWise config file: " & swConfig42)
	WScript.Quit
End If
IncludeNoDebug swConfig42

'Set the base path
if Not (o42.FolderExists(BasePath)) Then
	LocalScriptDebug.MessageBox("Could not find base path: " & BasePath)
End If
shell42.CurrentDirectory = BasePath
'LocalScriptDebug.MessageBox "Current Directory: " & shell42.CurrentDirectory

'Load Script configuration
If Not (o42.FileExists(ScriptConfig)) Then
	LocalScriptDebug.MessageBox("Could not find script config file: " & ScriptConfig)
	WScript.Quit
End If
IncludeNoDebug ScriptConfig

'Open log file
Call CreateFilesPath(LogPath & LogFile)
Dim l42
If LogFileAppend Then
	Set l42 = o42.OpenTextFile(LogPath & LogFile, 8, True)
Else
	Set l42 = o42.OpenTextFile(LogPath & LogFile, 2, True)
End If

'Log base path
Log "----------------------------------------", 1
Log "ScriptWise  Base Path: " & shell42.CurrentDirectory, 5

Function ProcessFile(xml)
	If Not (DumpRequestFile = "") Then
		Log "ScriptWise  Dumping Request XML: " & LogPath & DumpRequestFile, 8
		Call SaveXML(xml, LogPath & DumpRequestFile)
	Else
		Log "ScriptWise  Skipped Dumping Request XML", 8
	End If
	
	Log "ScriptWise  Loading Script: " & ScriptFile, 5
	Include ScriptFile
	Log "ScriptWise  Calling Script", 9
	ProcessFile = ProcessFile(xml)
	
	If Not (DumpReplyFile = "") Then
		Log "ScriptWise  Dumping Reply XML: " & LogPath & DumpReplyFile, 8
		Call SaveXML(ProcessFile, LogPath & DumpReplyFile)
	Else
		Log "ScriptWise  Skipped Dumping Reply XML", 8
	End If
	
	If Not (errors42 = "") Then
		LocalScriptDebug.MessageBox FailurePrompt & vbCrLf & errors42
	ElseIf Not (SuccessPrompt = "") Then
		LocalScriptDebug.MessageBox SuccessPrompt,64,"ScriptWise"
	ElseIf Not (SuccessPrompt = "") Then
		LocalScriptDebug.MessageBox SuccessPrompt,64,"ScriptWise"
	End If
	
	Log "ScriptWise  Complete", 5
	l42.close()
End Function

Function Include(ByVal n42)
	If Not (o42.FileExists(n42)) Then
		Log "Could not find include file: " & n42, 1
		LocalScriptDebug.MessageBox("Could not find include file: " & n42)
		WScript.Quit
	End If
	If LogLevel > 8 Then
		IncludeDebug(n42)
	Else
		IncludeNoDebug(n42)
	End If
End Function

Function IncludeNoDebug(ByVal n42)
	Dim f42: Set f42 = o42.OpenTextFile(n42, 1)
	Dim s42: s42 = f42.ReadAll()
	ExecuteGlobal s42
	f42.close()
End Function

Function IncludeDebug(ByVal n42)
	Dim f42: Set f42 = o42.OpenTextFile(n42, 1)
	Dim s42: s42 = ""
	Dim ln42: ln42 = 0
	Dim p42, c42
	Dim insideClass: insideClass = False
	Do Until f42.AtEndOfStream
	    ln42 = ln42 + 1
		c42 = f42.ReadLine
		c42 = Trim42(c42)
		cc42 = Replace(c42 ,"""","""""")
		
		If InStr(1, c42, "Class", VBTEXTCOMPARE) = 1 Then
			insideClass = True
		End If
		
		splitLine2 = False
		If splitLine1 Then
			splitLine2 = True
		End If
		splitLine1 = False
		If Right(c42, 1) = "_" Then
			splitLine1 = True
		End If
		
		If c42 = "" Or insideClass Or splitLine2 Or InStr(1, c42, "'", VBTEXTCOMPARE) = 1 Then
			' Dont add debug line
			s42 = s42 & c42 & vbCrLf
		Elseif  Not splitLine1 And (InStr(1, c42, "Function", VBTEXTCOMPARE) = 1  Or InStr(1, c42, "Sub", VBTEXTCOMPARE) = 1 Or InStr(1, c42, "Else", VBTEXTCOMPARE) = 1 Or InStr(1, c42, "Case", VBTEXTCOMPARE) = 1 Or InStr(1, c42, "End Select", VBTEXTCOMPARE) = 1 Or InStr(1, c42, "End If", VBTEXTCOMPARE) = 1) Then
			' Add debug line after
			s42 = s42 & c42 & vbCrLf
			s42 = s42 & "Log """ & n42 & ":" & ln42 & "  " & cc42 & """, 9 " & vbCrLf
		Else
			' By default add debug line before
			s42 = s42 & "Log """ & n42 & ":" & ln42 & "  " & cc42 & """, 9 " & vbCrLf
			s42 = s42 & c42 & vbCrLf
		End If
		
		If InStr(1, c42, "End Class", VBTEXTCOMPARE) = 1 Then
			insideClass = False
		End If
	Loop

	Log "Executing:" & vbCrLf & s42 & vbCrLf, 10	
	ExecuteGlobal s42
	f42.close()
End Function

Function Log(ByVal s42, ByVal n42)
	If n42 <= LogLevel Then
		l42.WriteLine(Now() & "  " & s42)
	End If
End Function

Function Error(ByVal sErrorMsg)
	Log "WARNING: " & sErrorMsg, 1
	errors42 = errors42 & "WARNING: " & sErrorMsg & vbCrLf
End Function

Function CriticalError(ByVal sErrorMsg)
	Log "ERROR: " & sErrorMsg, 1
	Call LocalScriptDebug.MessageBox("ERROR: " & sErrorMsg)
	l42.close()
	Call WScript.Quit()
End Function

Function Trim42(str)
	Dim re: Set re = New RegExp
	re.Pattern = "^\s*"
	re.Multiline = False
	Trim42 = re.Replace(str, "")
End Function

Function SaveXML(ByVal x42, ByVal n42)
	Log "ScriptWise  Saving XML: " & n42, 8
	Call CreateFilesPath(n42)
	Dim o42: set o42 = CreateObject("Scripting.FileSystemObject")
	Dim f42: set f42 = o42.OpenTextFile(n42, 2, True)
	f42.Write(x42)
	f42.close()
	Log "ScriptWise  Saved XML: " & n42, 8
End Function

Function CreateFilesPath(ByVal sFullPath)
	sFullPath = o42.GetParentFolderName(sFullPath)
	If Not o42.FolderExists(sFullPath) Then
		Call CreateFilesPath(o42.GetParentFolderName(sFullPath))
		o42.CreateFolder sFullPath
		'LocalScriptDebug.MessageBox "Created Folder: " & sFullPath
	End If
End Function
