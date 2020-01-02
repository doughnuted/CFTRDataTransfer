'Common functions that are specific to Helix
Class ContainerValue
	Public OrderID
	Public ContainerID
	Public ResultText
End Class

Function GetRequest(sRequestXml)
	Set oRequest = CreateObject("MSXML2.DOMDocument.6.0")
	oRequest.setProperty "SelectionNamespaces", "xmlns:req='http://www.cerner.com/ptl/localscript/request'"
	oRequest.LoadXML(sRequestXml)
	Set GetRequest = oRequest
End Function

Function GetReply(sRequestXml)
	sReplyXml = LocalScriptAccess.CreateLocalScriptReply("", True, "", True)
	Set oReply = CreateObject("MSXML2.DOMDocument.6.0")
	oReply.setProperty "SelectionNamespaces", "xmlns:rep='http://www.cerner.com/ptl/localscript/reply'"
	oReply.LoadXml(sReplyXml)
	Set GetReply = oReply
End Function

Function GetFilename(oRequest)
	GetFilename = oRequest.selectSingleNode("req:RequestData/req:FullFileName").Text
End Function

'Converts any accession numbers format to unformatted accessions i.e. 15-245-00001A or 1524500001A to 000002015245000001
Function ConvertAccession(sAccession)
	If Len(sAccession) = 12 Or Len(sAccession) = 13 Then
		ConvertAccession = "0000020" & Mid(sAccession,1,2) & Mid(sAccession,4,3) & "0" & Mid(sAccession,8,5)
	ElseIf Len(sAccession) = 10 Or Len(sAccession) = 11 Then
		ConvertAccession = "0000020" & Mid(sAccession,1,2) & Mid(sAccession,3,3) & "0" & Mid(sAccession,6,5)
	ElseIf Len(sAccession) = 18 Then
		ConvertAccession = sAccession
	Else
		Error "Accession doesn't match known formats: " & sAccession
		ConvertAccession = ""
	End If
End Function

'Test accession numbers if valid format i.e. 15-245-00001A or 1524500001A to 000002015245000001
Function IsValidAccession(sAccession)
	IsValidAccession = False
	If Len(sAccession) >= 10 And Len(sAccession) <= 13 Or Len(sAccession) = 18 Then
		If IsNumeric(Left(sAccession, 2)) Then
			IsValidAccession = True
		End If
	End If
End Function

' Get Output Container Values for the specified Order Number and Worksheet Column Name
Function GetOutPutContainerValues(sOrderID, sWSColumn)
		Dim oResults
		
		Set oValues = CreateObject("Scripting.Dictionary")
		Set oUCMCCLAccess = CreateObject("Cerner.Helix.UCMCCLAccess.CCLAccess")
		sCCLParam = sOrderID & "," & sWSColumn
		
'		Log "CCLParam: " & sCCLParam, 1
		
		sScriptReply = oUCMCCLAccess.ExecuteScript("GET_OUTPUT_CNTR_VAL_FOR_ORDER", sCCLParam, "")
'		Log "CCLResult: " & sScriptReply, 1
		
		set oXMLDoc = CreateObject("MSXML2.DOMDocument.3.0")
		Call oXMLDoc.loadxml(sScriptReply)
		
		Set oResults = oXMLDoc.selectNodes("/CCLREC/LIST/ITEM")	
		
		For Each oResult In oResults
		
			Set oContainerValue = New ContainerValue
			
			dPtlBatchItemId = oResult.getElementsByTagName("BATCHITEMID").item(0).attributes.getNamedItem("value").value
			oContainerValue.OrderID = oResult.getElementsByTagName("ORDERID").item(0).attributes.getNamedItem("value").value
			oContainerValue.ContainerID = oResult.getElementsByTagName("CONTAINERID").item(0).attributes.getNamedItem("value").value
			oContainerValue.ResultText = oResult.selectSingleNode("RESULTTEXT").Text
			
'			Log sResultText, 1
			
			oValues.Add dPtlBatchItemId, oContainerValue
		
		Next	
		
		Set GetOutPutContainerValues = oValues
		
End Function

Function GetContainerSuffix(rContainerID)

		Dim containerAlias
		Dim sScriptReply
		containerAlias = Array("", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")

		Set oUCMCCLAccess = CreateObject("Cerner.Helix.UCMCCLAccess.CCLAccess")
		
		sCCLParam = rContainerID
			
		sScriptReply = oUCMCCLAccess.ExecuteScript("GET_CONTAINER_PREFIX", sCCLParam, "")
		Log sScriptReply, 1
		
		If sScriptReply <> "" Then
			i = cInt(sScriptReply)
			GetContainerSuffix = containerAlias(i)
		Else
			GetContainerSuffix = ""
		End If

End Function