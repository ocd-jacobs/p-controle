Sub Main
	Call IndexDatabase()	'Onverdicht_Object_Aanv.IMD
End Sub


' Gegevens: Indexeer Database
Function IndexDatabase
	Set db = Client.OpenDatabase("Onverdicht_Object_Aanv.IMD")
	Set task = db.Index
	task.AddKey "MASSA", "A"
	task.AddKey "DAD_OBJECT", "A"
	task.Index FALSE
	Set task = Nothing
	Set db = Nothing
End Function