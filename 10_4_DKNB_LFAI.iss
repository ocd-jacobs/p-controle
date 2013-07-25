Sub Main
	Call FieldManipulationAppendFields()	'Tot-DienstKantorenBoek.IMD
End Sub


' Gegevens: Velden Bewerken - Velden Samenvoegen
Function FieldManipulationAppendFields
	Set db = Client.OpenDatabase("Tot-DienstKantorenBoek.IMD")
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
	eqn = "DK_A+DK_B+DK_C+@LEFT(DK_D;1)"
	field.Name = "DK_LONG_FAI"
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