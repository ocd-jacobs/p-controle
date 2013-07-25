Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.RunIDEAScriptEx  "01_Invoer_auditbest.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "02_Invoer_lc.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "03_Invoer_kp.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "04_Invoer_onverdicht.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "05_Aanvullen_kp.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "06_Export_Access.iss", strJaar, strMaand, "", ""
End Sub
