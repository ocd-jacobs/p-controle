Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call ReportReaderImport(strJaar, strMaand)	'F:\2009\P-2009\11_Nov\KP.txt
	Client.RefreshFileExplorer 
End Sub


' Bestand - Import Assistent: Report Reader
Function ReportReaderImport(strJaar, strMaand)
	dbName = "KP.IMD"
	Client.ImportPrintReport "F:\" & strJaar & "\P-" & strJaar & "\KP - T9AI08.jpm", "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\KP.txt", dbname, FALSE
	Client.OpenDatabase (dbName)
End Function