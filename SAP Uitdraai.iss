Begin Dialog frmSAPUitdraai 50,40,259,120,"SAP Uitdraai", .NieuwDialoogvenster
  TextBox 50,10,185,9, .edtPost
  Text 15,10,30,8, "Post nr.:", .Text1
  TextBox 50,26,186,9, .edtNaam
  Text 15,27,30,8, "Naam:", .Text1
  TextBox 50,44,185,9, .edtMaand
  Text 15,45,30,8, "Maand:", .Text1
  TextBox 50,61,185,10, .edtPersNr
  Text 15,61,30,8, "Pers. Nnr.:", .Text1
  OKButton 83,82,40,14, "OK", .OKButton1
  CancelButton 138,82,40,14, "Annuleren", .CancelButton1
End Dialog
Dim strPost As String
Dim strNaam As String
Dim strMaand As String
Dim strPersNr As String

Sub Main
	Dim a As frmSAPUitdraai
	Button = Dialog(frmSAPUitdraai)
	
	If Button <> -1 Then
		Exit Sub
	End If
	
	strJaar = "2013"
	
	strPost = frmSAPUitdraai.edtPost
	strNaam = frmSAPUitdraai.edtNaam
	strMaand = frmSAPUitdraai.edtMaand
	strpersNr = frmSAPUitdraai.edtPersNr
	
	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	'Client.workingDirectory = "F:\2011\P-2011\01_Jan" 
		
	Call JoinDatabase()	'Onverdicht.IMD
	Call ExportDatabaseXLS8()	'20002082 - jun.IMD
End Sub


' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("Onverdicht.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Looncomponent omschr.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "OMSCHRIJVING"
	task.AddMatchKey "LGART", "LC", "A"
	task.Criteria = "PERNR = " & Chr(34) & strPersNr & Chr(34)
	dbName = "Post " & strPost & " - " & strNaam & " - " & strMaand & ".IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bestand - Database Exporteren: XLS8
Function ExportDatabaseXLS8
	Set db = Client.OpenDatabase("Post " & strPost & " - " & strNaam & " - " & strMaand & ".IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "H:\"& "Post " & strPost & " - " & strNaam & " - " & strMaand & ".xls", "Database", "XLS8", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function