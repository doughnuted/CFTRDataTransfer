'ScriptWise Bootstrap

ScriptConfig = "MALDITOF/MALDITOF_CFTR_Analysis_Import.config"

'*** DO NOT EDIT below this line ***
Dim ScriptWisePath
Function ProcessFile(xml)
	Dim ws42: Set ws42 = CreateObject( "WScript.Shell" )
	ScriptWisePath = ws42.ExpandEnvironmentStrings( "%HELIX_SCRIPT_PATH%" ) & "\"
	Dim fn42: fn42 = ScriptWisePath & "ScriptWise.vbs"
	Dim o42: Set o42 = CreateObject("Scripting.FileSystemObject")
	If Not (o42.FileExists(fn42)) Then
		LocalScriptDebug.MessageBox("Could not load ScriptWise: " & fn42)
		ProcessFile = "<ReplyData xmlns=""http://www.cerner.com/ptl/localscript/reply""><FileName></FileName><StatusData><Status>Failure</Status><ErrorDescription>Could not load ScriptWise: " & fn42 & "</ErrorDescription></StatusData></ReplyData>"
		Exit Function
	End If
	Dim f42: Set f42 = o42.OpenTextFile(fn42, 1)
	ExecuteGlobal f42.ReadAll()
	f42.close()
	ProcessFile = ProcessFile(xml)
End Function