Sub Main
	Call FieldManipulationAppendFields()	'Tot-DienstKantorenBoek.IMD
End Sub

' Gegevens: Velden Bewerken - Velden Samenvoegen
Function FieldManipulationAppendFields
	Set db = Client.OpenDatabase("Onverdicht_KP_LC_Aanv.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	eqn = "@IF(@List(FEITALG; ""7802""; ""7822""; ""7842""; ""7844""; ""7862""; ""7864""; ""7863""); DIENSTKANTOOR_LANG; @If(@len(@alltrim(FEITALG)) = 0;@Left(TOT_KP;4);FEITALG) + ""000"")"
	field.Name = "ZOEKW_DKNBOEK"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = eqn
	field.Length = 7
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
End Function
