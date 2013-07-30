Sub Main
	strJaar = "2013"
	strMaand = "01_Jan"
	
	strSource = "F:\" & strJaar & "\P-" & strJaar & "\Process_Onverdicht.mdb"
	strDestination = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Process_Onverdicht.mdb"
	
	FileCopy strSource, strDestination 
	
	strSource = "F:\" & strJaar & "\P-" & strJaar & "\Looncomponent omschr.imd"
	strDestination = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Looncomponent omschr.imd"
	
	FileCopy strSource, strDestination 
	
	strSource = "F:\" & strJaar & "\P-" & strJaar & "\LC_Steek.imd"
	strDestination = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\LC_Steek.imd"

	FileCopy strSource, strDestination 
	
	strSource = "F:\" & strJaar & "\P-" & strJaar & "\Objecten.imd"
	strDestination = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Objecten.imd"
	
	FileCopy strSource, strDestination 

	strTitle = "     P-Steekproef " & UCase(Right(strmaand, 3)) & "     "
	strPrompt = "Zijn de Access bewerkingen uitgevoerd?"

	Client.RunIDEAScriptEx  "00_1_Startup.iss", strJaar, strMaand, "", ""
	
	intDoorgaan = MsgBox (strPrompt, MB_YESNO, strTitle)
	
	If intDoorgaan = IDYES Then
		Client.RunIDEAScriptEx  "00_2_Startup.iss", strJaar, strMaand, "", ""
	End If

	'Kill "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Onverdicht.mdb"
	'Kill "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\KP.mdb"
	'Kill "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Onverdicht.csv"
End Sub
