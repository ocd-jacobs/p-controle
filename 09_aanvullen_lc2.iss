Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call FieldManipulationAppendFields()	'Onverdicht_KP_LC.IMD
	Client.RefreshFileExplorer 
End Sub


' Gegevens: Velden Bewerken - Velden Samenvoegen
Function FieldManipulationAppendFields
	Set db = Client.OpenDatabase("Onverdicht_KP_LC.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	eqn = "@left(KOSTL;7)"
	field.Name = "LONG_FAI"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = eqn
	field.Length = 7
	task.AppendField field
	Set field = table.NewField
	eqn = "@if(@len(@trim(LC_KSTR2)) > 0 ; LC_KSTR2 ; LC_KSTAR)"
	field.Name = "VERDICHTNGS_KS"
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