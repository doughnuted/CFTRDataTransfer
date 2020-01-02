
Function ImportPolyTResult(oReqDoc, oRepDoc, sLine, sAssayIdMatch, sAllele1, sAllele2, sAllele3, sWellField, sCallField, sDescField)

	Log "Reading line: " & sLine, 9
	
	aFields = Split(sLine, ",")
	'Define fields required for import to Helix
	sSampleId = aFields(aCsvFieldMap(0))
	aSampleId = Split(sSampleId, "_")
	sSampleId = aSampleId(3)
	sSampleId = Mid(sSampleId,1,2) & "-" & Mid(sSampleId,3,3) & "-" & Mid(sSampleId,6,5)
	sSampleDescription = aFields(aCsvFieldMap(4))
	sCall = aFields(aCsvFieldMap(2))
	sAssayId = aFields(aCsvFieldMap(3))
	sDescription = aFields(aCsvFieldMap(5))

	'Check if accession number present in Request XML
	Set oReqBatchItem = LocalScriptAccess.GetRequestBatchItemNodeByAccession (oReqDoc, sSampleId)
	If Not (oReqBatchItem Is Nothing) Then
	
		sOrderId = LocalScriptAccess.XMLGetChildNodeText(oReqBatchItem, "req:OrderId")
		sContainerId = LocalScriptAccess.XMLGetChildNodeText(oReqBatchItem, "req:ContainerId")
		
		If Not (sSampleId = sCurrentSampleId) Then
			sCurrentSampleId = sSampleId
			'Add batch item to XML if not present
			Set oBatchItems = oRepDoc.selectSingleNode("/rep:ReplyData/rep:Protocol/rep:BatchItems")
			Set oBatchItem = LocalScriptAccess.AddBatchItemInfo(oBatchItems, sOrderId, sContainerId, 0)
		End If
	
		If sAssayId = sAssayIdMatch Then
			If InStr(1, "|" & sAllele1 & "|", "|" & sCall & "|") > 0 Then
				sCallFinal = sAllele1
			ElseIf InStr(1, "|" & sAllele2 & "|", "|" & sCall & "|") > 0 Then
				sCallFinal = sAllele2
			ElseIf InStr(1, "|" & sAllele3 & "|", "|" & sCall & "|") > 0 Then
				sCallFinal = sAllele3
			Else
				sCallFinal = ""
				Log "WARNING Call Not Valid: " & sAssayId & " " & sCall, 1
				'LocalScriptDebug.MessageBox("WARNING Call Not Valid: " & sAssayId)
			End If

			'Add results to XML
			'Set oProtocolResults = oBatchItem.selectSingleNode("rep:ProtocolResults") not required as panel is multi well
			'Call LocalScriptAccess.AddResultInfo(oProtocolResults, "ProtocolResult", sWellField, sSampleDescription, False) not required as panel is multi well
			Set oProtocolResults = oBatchItem.selectSingleNode("rep:ProtocolResults")
			Call LocalScriptAccess.AddResultInfo(oProtocolResults, "ProtocolResult", sDescField, sDescription, False)
			
			If Not (sCallFinal = "") Then
				Set oProtocolResults = oBatchItem.selectSingleNode("rep:ProtocolResults")
				Call LocalScriptAccess.AddResultInfo(oProtocolResults, "ProtocolResult", sCallField, sCallFinal, False)
			End If

		End If
	End If
End Function