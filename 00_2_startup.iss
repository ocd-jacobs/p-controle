Sub Main
	strJaar = arg1
	strMaand = arg2

	Client.RunIDEAScriptEx  "07_Import_Access.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "08_Aanvullen_lc.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "09_Aanvullen_lc2.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "10_0_Hernoemen_DK.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "10_1_Aanvulling.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "10_2_Aanvulling2.iss", strJaar, strMaand, "", ""

	Client.RunIDEAScriptEx  "10_3_ZOEKW_DKNB.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "10_4_DKNB_LFAI.iss", strJaar, strMaand, "", ""	
	
	Client.RunIDEAScriptEx  "11_Toevoegen_DK.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "12_Uitsluiten_RVDR.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "13_Aansluit_Totalen.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "14_Aanv_Object.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "15_Aansluiting.iss", strJaar, strMaand, "", ""
	
	Client.RunIDEAScriptEx  "16_Aanv_massa.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "17_Index_massa.iss", strJaar, strMaand, "", ""

	Client.RunIDEAScriptEx  "18a_Deelmassa.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "18b_Deelmassa GVKA.iss", strJaar, strMaand, "", ""
	
	Client.RunIDEAScriptEx  "19a_MUS_1_1.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "19b_MUS_1_2.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "19c_MUS_1_3.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "19d_MUS_1_4.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "19e_MUS_1_5.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "19f_MUS_1_6.iss", strJaar, strMaand, "", ""
	
	Client.RunIDEAScriptEx  "20a_MUS_2_1.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "20b_MUS_2_2.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "20c_MUS_2_3.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "20d_MUS_2_4.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "20e_MUS_2_5.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "20f_MUS_2_6.iss", strJaar, strMaand, "", ""
	
	Client.RunIDEAScriptEx  "22_Steekproef_TOTAAL.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "23_LC_OMS.iss", strJaar, strMaand, "", ""
	Client.RunIDEAScriptEx  "24_Aanvullen_Steekproef.iss", strJaar, strMaand, "", ""

	Client.RunIDEAScriptEx  "25_LC_telling.iss", strJaar, strMaand, "", ""	
	Client.RunIDEAScriptEx  "26_export_steekproef.iss", strJaar, strMaand, "", ""
End Sub
