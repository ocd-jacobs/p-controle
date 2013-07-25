Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call FieldManipulationAppendFields()	'KP.IMD
	Client.RefreshFileExplorer 
End Sub


' Gegevens: Velden Bewerken - Velden Samenvoegen
Function FieldManipulationAppendFields
	Set db = Client.OpenDatabase("KP.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	eqn = "@RegExpr( KP_DKNTR  ;  ""[^*]+"")"
	field.Name = "KP_ZOEK_WAARDE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = eqn
	field.Length = 10
	task.AppendField field
	Set field = table.NewField
	eqn = "@Len(@AllTrim(KP_ZOEK_WAARDE))"
	field.Name = "KP_ZOEK_LENGTE"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = eqn
	field.Decimals = 0
	task.AppendField field
	Set field = table.NewField
	eqn = "@Left(KP_DKNTR; 4)"
	field.Name = "KP_ZOEK_INDEX"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = eqn
	field.Length = 4
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
End Function