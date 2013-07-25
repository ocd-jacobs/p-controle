Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.workingDirectory = "F:\" & strJaar & "\P-" & strJaar & "\" & strMaand 
	Call FieldManipulationAppendFields()	'Onverdicht_KP_LC_Aanv.IMD
	Client.RefreshFileExplorer 
End Sub


' Gegevens: Velden Bewerken - Velden Samenvoegen
Function FieldManipulationAppendFields
	Set db = Client.OpenDatabase("Onverdicht_KP_LC_Aanv.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	eqn = "@if(@AllTrim(KOSTL) = """" ; KOSTENPL ; KOSTL)"
	field.Name = "TOT_KP"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = eqn
	field.Length = 10
	task.AppendField field
	Set field = table.NewField
	eqn = "@left(TOT_KP; 4)"
	field.Name = "TOT_FEITALG"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = eqn
	field.Length = 4
	task.AppendField field
	eqn = "@left(TOT_KP;7)"
	field.Name = "DIENSTKANTOOR_LANG"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = eqn
	field.Length = 7
	task.AppendField field
	Set field = table.NewField
	eqn = "@val(VERDICHTNGS_KS)"
	field.Name = "KOSTENSOORT_NUM"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = eqn
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
End Function