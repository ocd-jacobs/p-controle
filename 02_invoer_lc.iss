Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call ReportReaderImport(strJaar, strMaand)	'F:\2009\P-2009\12_Dec\LC.txt	
	Client.RefreshFileExplorer 
End Sub


' Bestand - Import Assistent: Report Reader
Function ReportReaderImport(strJaar, strMaand)
	dbName = "LC.IMD"
	Client.ImportPrintReport "F:\" & strJaar & "\P-" & strJaar & "\LC - T9AI09.jpm", "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\LC.txt", dbname, FALSE
	Client.OpenDatabase (dbName)
End Function