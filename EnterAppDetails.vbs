Dim CurrID, CurrApp, CurrVer, CurrFree, CurrOS, CurrReason, CurrPC, CurrPlans, CurrUpdate, CurrURL, CurrQTH, CurrVar
Dim adoconn
Dim rs
Dim str
dim WPData 'Web page text
dim xmlhttp : set xmlhttp = createobject("msxml2.xmlhttp.3.0")

'''''''''''''''''''
'Required Variables

'Database
DBPass = "P@ssword1" 'Password to access database on localhost

'''''''''''''''''''

'Ask for App Reference ID
CurrID = inputbox("Enter the reference ID of the application that you would like to update:", "Software Matrix", "")

if len(CurrID) > 0 and isnumeric(CurrID) then
	Set adoconn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	adoconn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=localhost;" & _
                   "Database=software_matrix; User=root; Password=" & DBPass & ";"
	
	str = "Select * from discoveredapplications where ID = '" & CurrID & "';"
	rs.Open str, adoconn, 3, 3 'OpenType, LockType
	
	if not rs.eof then
		rs.movefirst
		CurrApp = rs("Name") & ""
		CurrVer = rs("Version_Newest") & ""
		CurrFree = rs("Free") & ""
		CurrOS = rs("OpenSource") & ""
		CurrReason = rs("ReasonForSoftware") & ""
		CurrPC = rs("NeededOnMachines") & ""
		CurrPlans = rs("PlansForRemoval") & ""
		CurrUpdate = rs("Update Method") & ""
		CurrURL = rs("UpdateURL") & ""
		CurrQTH = rs("UpdatePageQTH") & ""
		CurrVar = rs("UpdatePageQTHVarience") & ""
		
		'Prompt for data
		msgbox "You are updating: " & CurrApp
		CurrFree = inputbox("Is " & CurrApp & " free? (Y/N)", "Software Matrix", CurrFree)
		CurrOS = inputbox("Is " & CurrApp & " Open Source? (Y/N)", "Software Matrix", CurrOS)
		if CurrFree = "Y" or CurrOS = "Y" then msgbox "You are adding FOSS!!!" & vbCrlf & "(Good for you)"
		CurrReason = inputbox("What is the reason for adding " & CurrApp & " to the network?", "Software Matrix", CurrReason)
		CurrPC = inputbox("Generally speaking, which machines will " & CurrApp & " be used on?", "Software Matrix", CurrPC)
		CurrPlans = inputbox("Generally speaking, what are the plans to remove " & CurrApp & "?", "Software Matrix", CurrPlans)
		CurrUpdate = inputbox("How is " & CurrApp & " updated? (Manual/Automatic/None)", "Software Matrix", CurrUpdate)
		if CurrUpdate = "None" then
			CurrURL = ""
			CurrQTH = 0
			CurrVar = 10
		else
			CurrURL = inputbox("Enter the URL where the version number (" & CurrVer & ") can be found:", "Software Matrix", CurrURL)
			if len(CurrURL) = 0 then
				CurrQTH = 0
				CurrVar = 10
			else
				'Pull website
				xmlhttp.open "get", CurrURL, false
				xmlhttp.send
				WPData = xmlhttp.responseText
				set xmlhttp = nothing
				if instr(1,WPData,CurrVer,0)>0 then
					CurrQTH = instr(1,WPData,CurrVer,0)
					CurrQTH = inputbox("This is the location where the version (" & CurrVer & ") is found on the URL that was entered. You should leave this AS IS unless alerts are triggered and version is current. In that case, clear this:", "Software Matrix", CurrQTH)
				else
					CurrQTH = inputbox("The version (" & CurrVer & ") was not found on the URL that was entered. You should leave this blank and verify that the URL is correct:", "Software Matrix", CurrQTH)
				end if
				if not CurrQTH = "" and not CurrQTH = "0" then CurrVar = inputbox("Leave this at 10 unless you know what you are doing:", "Software Matrix", CurrVar) else CurrVar = 10
			end if
		end if
		msgbox CurrApp & " will now be updated"
		
		rs("Name") = CurrApp
		rs("Version_Newest") = CurrVer
		rs("Free") = CurrFree
		rs("OpenSource") = CurrOS
		if CurrFree = "Y" or CurrOS = "Y" then rs("FOSS") = "Y" else rs("FOSS") = "N"
		rs("ReasonForSoftware") = CurrReason
		rs("NeededOnMachines") = CurrPC
		rs("PlansForRemoval") = CurrPlans
		rs("Update Method") = CurrUpdate
		rs("UpdateURL") = CurrURL
		if len(CurrQTH) = 0 then rs("UpdatePageQTH") = 0 else rs("UpdatePageQTH") = CurrQTH
		rs("UpdatePageQTHVarience") = CurrVar
		rs.update
		rs.close
	else
		msgbox "ID entered was not found in the DB!"
	end if
end if
