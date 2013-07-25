Sub Main
	Call JoinDatabase()	'Steekproef.IMD
	Call FieldManipulationAppendFields()	'Steekproef.IMD
End Sub


' Bestand: Databases Combineren
Function JoinDatabase
	Set db = Client.OpenDatabase("MUS aangevuld.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Tot-tblQryStamgegevens.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "PERNR", "PERSNR", "A"
	dbName = "Steekproef.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Gegevens: Velden Bewerken - Velden Samenvoegen
Function FieldManipulationAppendFields
	Set db = Client.OpenDatabase("Steekproef.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	eqn = """"""
	field.Name = "POSTNUMMER"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = eqn
	field.Length = 10
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
End Function
