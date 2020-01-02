
Include "HelixLib.vbs"
Include "MALDITOF/MALDITOF_CFTR_Common_Import.vbs"
Include "MALDITOF/MALDITOF_PolyT_Import.vbs"

Dim sCurrentSampleId, oBatchItem

Dim aCsvFieldMap(6): aCsvFieldMap(0) = 0 : aCsvFieldMap(1) = 1 : aCsvFieldMap(2) = 2 : aCsvFieldMap(3) = 3 : aCsvFieldMap(4) = 4 : aCsvFieldMap(5) = 5

Function ProcessFile(sXml)

  Set oLSA = LocalScriptAccess
  Set oReqDoc = GetRequest(sXml)
  Set oRepDoc = GetReply(sXml)
  sFilename = GetFilename(oReqDoc)
	
  Log "Writing request XML to file: " & sFilename, 1
  
	'Load CSV
  Set oFS = CreateObject("Scripting.FileSystemObject")
  Set oCSV = oFS.OpenTextFile(sFileName)
  
  'Drop header rows
  For i = 0 To 2
  oCSV.ReadLine
  Next
  
  'Loop through CSV
  Do Until oCSV.AtEndOfStream
	sLine = oCSV.ReadLine

	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S549N", "WT", "WT.S549N", "S549N", "Text1", "Results1", "Results2")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "3199del6", "WT", "WT.DEL6", "DEL6", "Text1", "Results3", "Results4")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S549R_CGT", "WT", "S549R.WT", "S549R", "Text1", "Results5", "Results6")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "G542X", "WT", "WT.G542X", "G542X", "Text1", "Results7", "Results8")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "3791delC:c.3659delC", "WT", "WT.3791delC", "3791delC", "Text1", "Results9", "Results10")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "2143delT", "WT", "WT.2143delT", "2143delT", "Text1", "Results11", "Results12")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R117H", "WT", "R117H.WT", "R117H", "Text1", "Results13", "Results14")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "E60X", "WT", "WT.E60X", "E60X", "Text1", "Results15", "Results16")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R75X", "WT", "WT.R75X", "R75X", "Text1", "Results17", "Results18")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "406-1G->A c.274-1G>A", "WT", "WT.406-1G->A", "406-1G->A", "Text1", "Results19", "Results20")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "Q890X", "WT", "WT.Q890X", "Q890X", "Text1", "Results21", "Results22")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "2307insA c.2175_2176insA", "WT", "WT.2307insA", "2307insA", "Text1", "Results23", "Results24")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "W1089X", "WT", "WT.W1089X", "W1089X", "Text1", "Results25", "Results26")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "D1152H", "WT", "WT.D1152H", "D1152H", "Text1", "Results27", "Results28")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "K710X", "WT", "WT.K710X", "K710X", "Text1", "Results29", "Results30")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "G330X", "WT", "WT.G330X", "G330X", "Text1", "Results31", "Results32")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R1066C", "WT", "WT.R1066C", "R1066C", "Text1", "Results33", "Results34")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "3905insT", "WT", "WT.3905insT", "3905insT", "Text1", "Results35", "Results36")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S1196X", "WT", "S1196X.WT", "S1196X", "Text1", "Results37", "Results38")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "1677delTA c.1545_1546delTA", "WT", "WT.1677delTA", "1677delTA", "Text1", "Results39", "Results40")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "2183AA->G c.2051_2052delAAinsG", "WT", "2183AA->G.WT", "2183AA->G", "Text1", "Results41", "Results42")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "3876delA c.3744delA", "WT", "3876delA.WT", "3876delA", "Text1", "Results43", "Results44")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "deltaI507 (I507 del)", "WT|WT.I507del|I507del", "WT.I507V|I507V.I507del", "I507V", "Text1", "Results45", "Results46") '2 targets I507V
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "deltaI507 (I507 del)", "WT|WT.I507V|I507V", "WT.I507del|I507V.I507del", "I507del", "Text1", "Results47", "Results48") '2 targets I507del
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "L206W", "WT", "L206W.WT", "L206W", "Text1", "Results49", "Results50")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "Y1092X_TAA", "WT", "WT.Y1092X_TAA", "Y1092X_TAA", "Text1", "Results51", "Results52")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "W1282X", "WT", "WT.W1282X", "W1282X", "Text1", "Results53", "Results54")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "M1101K", "WT", "M1101K.WT", "M1101K", "Text1", "Results55", "Results56")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R1162X", "WT", "WT.R1162X", "R1162X", "Text1", "Results57", "Results58")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "A455E", "WT", "WT.A455E", "A455E", "Text1", "Results59", "Results60")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "G1349D", "WT", "WT.G1349D", "G1349D", "Text1", "Results61", "Results62")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R117C", "WT", "WT.R117C", "R117C", "Text1", "Results63", "Results64")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "N1303K", "WT", "N1303K.WT", "N1303K", "Text1", "Results65", "Results66")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "G551S", "WT", "G551S.WT", "G551S", "Text1", "Results67", "Results68")
	Call ImportPolyTResult(oReqDoc, oRepDoc, sLine, "PolyT_T5/T7", "T5", "T5.T7", "T7", "Text1", "Results69", "Results70") 'T5 and T7 are targets
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "1078delT", "WT", "1078delT.WT", "1078delT", "Text1", "Results71", "Results72")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "dele2-3_3prime", "WT", "dele2-3_3prime.WT", "dele2-3_3prime", "Text1", "Results73", "Results74")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "G178R", "WT", "WT.G178R", "G178R", "Text1", "Results75", "Results76")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "935delA:c.803delA", "WT", "WT.935delA", "935delA", "Text1", "Results77", "Results78")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R553X", "WT", "R553X.WT", "R553X", "Text1", "Results79", "Results80")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "2789+5G->A c.2657+5G>A", "WT", "WT.2789+5G->A", "2789+5G->A", "Text1", "Results81", "Results82")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "Y122X", "WT", "WT.Y122X", "Y122X", "Text1", "Results83", "Results84")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "3849+10kbC->T", "WT", "WT.3849+10kbC->T", "3849+10kbC->T", "Text1", "Results85", "Results86")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R560T", "WT", "WT.R560T", "R560T", "Text1", "Results87", "Results88")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "G551D", "WT", "WT.G551D", "G551D", "Text1", "Results89", "Results90")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R1158X", "WT", "WT.R1158X", "R1158X", "Text1", "Results91", "Results92")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "1717-1G->A c.1585-1G>A", "WT", "1717-1G>A.WT", "1717-1G>A", "Text1", "Results93", "Results94")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "I506V", "WT", "WT.MUT", "MUT", "Text1", "Results95", "Results96")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "1898+5G->T", "WT", "WT.1898+5G->T", "1898+5G->T", "Text1", "Results97", "Results98")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S1255P", "WT", "S1255P.WT", "S1255P", "Text1", "Results99", "Results100")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S1255X", "WT|S1255L|WT.S1255L", "S1255X.WT|S1255L.S1255X", "S1255X", "Text1", "Results101", "Results102") '2 targets S1255X
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S1255X", "WT|S1255X|WT.S1255X", "S1255L.WT|S1255L.S1255X", "S1255L", "Text1", "Results103", "Results104") '2 targets S1255L
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S549R-AGA-AGG", "WT|S549R|WT.S549R", "S549R-AGG.WT|S549R-AGG.S549R", "S549R-AGG", "Text1", "Results105", "Results106") '2 targets S549R-AGG
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S549R-AGA-AGG", "WT|S549R-AGG|S549R-AGG.WT", "WT.S549R|S549R-AGG.S549R", "S549R", "Text1", "Results107", "Results108") '2 targets S549R
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "V520F", "WT", "WT.V520F", "V520F", "Text1", "Results109", "Results110")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "S1251N", "WT", "S1251N.WT", "S1251N", "Text1", "Results111", "Results112")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "G1244E", "WT", "WT.G1244E", "G1244E", "Text1", "Results113", "Results114")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R347PH", "R347H|WT|R347H.WT", "R347P.R347H|R347P.WT", "R347P", "Text1", "Results115", "Results116") '2 targets R347P
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R347PH", "R347P|WT|R347P.WT", "R347P.R347H|R347H.WT", "R347H", "Text1", "Results117", "Results118") '2 targets R347H
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "dele2-3_5prime", "WT", "WT.dele2-3_5prime", "dele2-3_5prime", "Text1", "Results119", "Results120")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "deltaF508(F508del)", "WT", "WT.F508del", "F508del", "Text1", "Results121", "Results122")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "2055del9>A: c.1923_1931del9insA", "WT", "WT.2055del9>A", "2055del9>A", "Text1", "Results123", "Results124")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "2184delA c.2052delA", "WT", "2184delA.WT", "2184delA", "Text1", "Results125", "Results126")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "Q493X", "WT", "Q493X.WT", "Q493X", "Text1", "Results127", "Results128")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R334W", "WT", "WT.R334W", "R334W", "Text1", "Results129", "Results130")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "711+1G->T", "WT", "WT.711+1G->T", "711+1G->T", "Text1", "Results131", "Results132")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "A559T", "WT", "A559T.WT", "A559T", "Text1", "Results133", "Results134")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "621+1G->T", "WT", "WT.621+1G->T", "621+1G->T", "Text1", "Results135", "Results136")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "394delTT c.262_263delTT", "WT", "WT.394delTT", "394delTT", "Text1", "Results137", "Results138")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "G85E", "WT", "G85E.WT", "G85E", "Text1", "Results139", "Results140")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R560KT", "WT|R560K|WT.R560K", "WT.R560T|R560T.R560K", "R560T", "Text1", "Results141", "Results142") '2 targets R560T
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R560KT", "WT|R560T|WT.R560T", "R560T.R560K|WT.R560K", "R560K", "Text1", "Results143", "Results144") '2 targets R560K
	Call ImportPolyTResult(oReqDoc, oRepDoc, sLine, "PolyT_T7/T9", "T7", "T7.T9", "T9", "Text1", "Results145", "Results146") 'T7 and T9 are targets
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R1162X-ALT", "WT|R1162X|WT.R1162X", "WT.R1162Q|R1162Q.R1162X", "R1162Q", "Text1", "Results147", "Results148") '2 targets R1162Q
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "R1162X-ALT", "WT|R1162Q|WT.R1162Q", "R1162Q.R1162X|WT.R1162X", "R1162X", "Text1", "Results149", "Results150") '2 targets R1162X
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "3659delC c.3528delC", "WT", "WT.3659delC", "3659delC", "Text1", "Results151", "Results152")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "3120+1G->A c.2988G>A", "WT", "WT.3120+1G->A", "3120+1G->A", "Text1", "Results153", "Results154")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "1898+1G->A", "WT", "18981G>A.WT", "18981G>A", "Text1", "Results155", "Results156")
	Call ImportSpecificResult(oReqDoc, oRepDoc, sLine, "F508C", "WT", "F508C.WT", "F508C", "Text1", "Results157", "Results158")
  Loop
  
  ProcessFile = oRepDoc.xml
End Function