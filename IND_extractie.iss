Sub Main
	Call DirectExtraction()	'Onverdicht_KP_LC_DK.IMD
	Call DirectExtraction1()	'IND Onverdicht feb 2011.IMD
	Call JoinDatabase()	'IND Onverdicht kostenrekeningen.IMD
	Call JoinDatabase1()	'IND Onverdicht + LC_Oms.IMD
	Call ExportDatabaseXLS8()	'IND Onverdicht Aangevuld.IMD
End Sub


' Gegevens: Directe Selectie
Function DirectExtraction
	Set db = Client.OpenDatabase("Onverdicht_KP_LC_DK.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "PERNR"
	task.AddFieldToInc "PERNR_OUD"
	task.AddFieldToInc "FPBEG"
	task.AddFieldToInc "FPEND"
	task.AddFieldToInc "IPBEG"
	task.AddFieldToInc "IPEND"
	task.AddFieldToInc "BUKRS"
	task.AddFieldToInc "KOSTL"
	task.AddFieldToInc "LGART"
	task.AddFieldToInc "BTZNR"
	task.AddFieldToInc "BETRG"
	task.AddFieldToInc "KP_MANDT"
	task.AddFieldToInc "KP_BUKRS"
	task.AddFieldToInc "KP_DKNTR"
	task.AddFieldToInc "KP_KOSTL"
	task.AddFieldToInc "KP_ZZHKONT"
	task.AddFieldToInc "KP_KASBH"
	task.AddFieldToInc "KP_ZZKASBH2"
	task.AddFieldToInc "LC_MANDT"
	task.AddFieldToInc "LC_LGART"
	task.AddFieldToInc "LC_KSTAR"
	task.AddFieldToInc "LC_KSTR2"
	task.AddFieldToInc "LC_KSTR3"
	task.AddFieldToInc "VERDICHTNGS_KS"
	task.AddFieldToInc "KOSTENPL"
	task.AddFieldToInc "KOSTENSOORT_NUM"
	task.AddFieldToInc "DK_OMSCHRIJVING"
	task.AddFieldToInc "DK_DIRECTIE"
	task.AddFieldToInc "DK_DIENST"
	task.AddFieldToInc "DK_ARTIKEL"
	dbName = "IND Onverdicht feb 2011.IMD"
	task.AddExtraction dbName, "", "DK_DIENST = ""IND"""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Gegevens: Directe Selectie
Function DirectExtraction1
	Set db = Client.OpenDatabase("IND Onverdicht feb 2011.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "IND Onverdicht kostenrekeningen.IMD"
	task.AddExtraction dbName, "", "KOSTENSOORT_NUM > 4000000"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("IND Onverdicht kostenrekeningen.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Looncomponent omschr.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "OMSCHRIJVING"
	task.AddMatchKey "LGART", "LC", "A"
	dbName = "IND Onverdicht + LC_Oms.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bestand: Databases Combineren
Function JoinDatabase1
	Set db = Client.OpenDatabase("IND Onverdicht + LC_Oms.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Tot-tblQryStamgegevens.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ADR_LN1"
	task.AddMatchKey "PERNR", "PERSNR", "A"
	dbName = "IND Onverdicht Aangevuld.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Bestand - Database Exporteren: XLS8
Function ExportDatabaseXLS8
	Set db = Client.OpenDatabase("IND Onverdicht Aangevuld.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "F:\2011\P-2011\02_Feb\IND Onverdicht Aangevuld.XLS", "Database", "XLS8", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function