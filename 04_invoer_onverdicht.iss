Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call TextImport(strJaar, strMaand)	'F:\2009\P-2009\12_Dec\onverdicht_0210_20091001_20091021-022456-000.csv
	Client.RefreshFileExplorer 
End Sub


' Bestand - Import Assistent: Vaste Lengte Tekst
Function TextImport(strJaar, strMaand)
	dbName = "Onverdicht.IMD"
	Client.ImportDatabase "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand & "\Onverdicht.csv", dbName, FALSE, FALSE, "",  "F:\" & strJaar & "\P-" & strJaar & "\onverdicht_0905.RDF"
	Client.OpenDatabase (dbName)
End Function