dim OVersion 'Oldest Installed version
dim CVersion 'Installed version
dim outputl 'Email body
Dim AllApps 'Data from CSV
dim WPData 'Web page text
Dim yfound 'For new apps, series of tests to find similar apps
Dim UpdatePageQTH, UpdatePageQTHVarience 'Used to fix any integer values in the two fields that are actually NULL
Dim adoconn
Dim rs
Dim str
set filesys=CreateObject("Scripting.FileSystemObject")
set xmlhttp = createobject("msxml2.serverxmlhttp.3.0")
Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir = filesys.GetParentFolderName(Wscript.ScriptFullName)
Dim AppRenames()
Dim RenameTo, RenameEx 'AI App Renaming
Dim Response 'For answers to prompts
Dim PSSchema, PSTbl 'Define schema and table names
PSSchema = "software_matrix"
PSTbl = "discoveredapplications"

'Gather variables from smapp.ini or prompt for them and save them for next time
If filesys.FileExists(strCurDir & "\smapp.ini") then
	'Database
	CSVPath = ReadIni(strCurDir & "\smapp.ini", "Database", "CSVPath" )
	DBLocation = ReadIni(strCurDir & "\smapp.ini", "Database", "DBLocation" )
	DBUser = ReadIni(strCurDir & "\smapp.ini", "Database", "DBUser" )
	DBPass = ReadIni(strCurDir & "\smapp.ini", "Database", "DBPass" )
	
	'Email - Defaults to anonymous login
	RptToEmail = ReadIni(strCurDir & "\smapp.ini", "Email", "RptToEmail" )
	RptFromEmail = ReadIni(strCurDir & "\smapp.ini", "Email", "RptFromEmail" )
	EmailSvr = ReadIni(strCurDir & "\smapp.ini", "Email", "EmailSvr" )
	EmailPort = ReadIni(strCurDir & "\smapp.ini", "Email", "EmailPort" )
	EmailAuthType = ReadIni(strCurDir & "\smapp.ini", "Email", "EmailAuthType" )
	EmailUserName = ReadIni(strCurDir & "\smapp.ini", "Email", "EmailUserName" )
	EmailPassword = ReadIni(strCurDir & "\smapp.ini", "Email", "EmailPassword" )
	'Additional email settings found in smapp.ini
	
	'WebGUI
	BaseURL = ReadIni(strCurDir & "\smapp.ini", "WebGUI", "BaseURL" )
else
	msgbox "INI file not found at: " & strCurDir & "\smapp.ini" & vbCrlf & "You will now be prompted with questions to create it."
	
	'Database
	CSVPath = inputbox("Enter the location where the CSV file with the software dump can be found (UNC path recommended):", "Software Matrix", strCurDir & "\Applications.csv")
	DBLocation = inputbox("Enter the IP address or hostname for the location of the database:", "Software Matrix", "localhost")
	DBUser = inputbox("Enter the user name to access database on " & DBLocation & ":", "Software Matrix", "user")
	DBPass = inputbox("Enter the password to access database on " & DBLocation & ":", "Software Matrix", "P@ssword1")
	
	'Check to see if DB exists
	CheckForTables
	
	'Email - Defaults to anonymous login
	RptToEmail = inputbox("Enter the report email's To address:", "Software Matrix", "admin@company.com")
	RptFromEmail = inputbox("Enter the report email's From address:", "Software Matrix", "admin@company.com")
	EmailSvr = inputbox("Enter the FQDN or IP address of email server:", "Software Matrix", "mail.server.com")
	EmailPort = "25"
	EmailAuthType = "0" '0 = Do not authenticate, 1 = basic (clear-text) authentication, 2 = NTLM
	EmailUserName = RptFromEmail
	EmailPassword = ""
	msgbox "Additional email settings found in smapp.ini"
	
	'WebGUI
	BaseURL = inputbox("Enter the base URL for the Software Matrix GUI (Web GUI available at https://github.com/compuvin/SoftwareMatrix-GUI):", "Software Matrix", "http://www.intranet.com")
		
	'Write the data to INI file
	WriteIni strCurDir & "\smapp.ini", "Database", "CSVPath", CSVPath
	WriteIni strCurDir & "\smapp.ini", "Database", "DBLocation", DBLocation
	WriteIni strCurDir & "\smapp.ini", "Database", "DBUser", DBUser
	WriteIni strCurDir & "\smapp.ini", "Database", "DBPass", DBPass
	WriteIni strCurDir & "\smapp.ini", "Email", "RptToEmail", RptToEmail
	WriteIni strCurDir & "\smapp.ini", "Email", "RptFromEmail", RptFromEmail
	WriteIni strCurDir & "\smapp.ini", "Email", "EmailSvr", EmailSvr
	WriteIni strCurDir & "\smapp.ini", "Email", "EmailPort", EmailPort
	WriteIni strCurDir & "\smapp.ini", "Email", "EmailAuthType", EmailAuthType
	WriteIni strCurDir & "\smapp.ini", "Email", "EmailUserName", EmailUserName
	WriteIni strCurDir & "\smapp.ini", "Email", "EmailPassword", EmailPassword
	WriteIni strCurDir & "\smapp.ini", "WebGUI", "BaseURL", BaseURL
end if
			   
outputl = ""

If filesys.FileExists(CSVPath) then
	AllApps = GetFile(CSVPath)

	if len(AllApps) > 100 then
		'Fix for if a single quote exists in file
		AllApps = replace(AllApps,"'","''")
		
		Set adoconn = CreateObject("ADODB.Connection")
		Set rs = CreateObject("ADODB.Recordset")
		adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & DBLocation & ";" & _
					   "Database=" & PSSchema & "; User=" & DBUser & "; Password=" & DBPass & ";"

		Get_App_Renames 'Uses apprename table to check regex for apps that match patterns and should be listed as the same app
		
		Get_PC_New_Updated 'List software Added/Updated from each PC
		Get_PC_Removed 'List software removed from each PC

		if outputl <> "" then
			outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body>" & vbcrlf & outputl
			SendMail RptToEmail, "Software Matrix: Change Report"
			outputl = ""
		end if

		Get_Organization_New 'List software that has never been seen in the organization before
		Get_Organization_Removed 'List software that no longer exists in the organization

		if outputl <> "" then
			outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body>" & vbcrlf & outputl
			SendMail RptToEmail, "Software Matrix: Security Report"
			outputl = ""
		end if
		
		'msgbox "CSV Exists"
		
		filesys.DeleteFile CSVPath, force
	End if
End if


Function Get_PC_New_Updated()
	Dim AllApps_Org, CurrApp, CurrVer 'For Orgs
	Dim CurrPC, CurrPub 'For PCs only
	Dim CurrAppNoVer, TestFOSS, TestFOSSFree, TestFOSSOS 'App without version 
	
	'Update the Organization stats
	AllApps_Org = right(AllApps,len(AllApps)-32)
	do while len(AllApps_Org) > 10
		'Ignore PC name
		AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,",",1))
		'Get application
		if left(AllApps_Org,1)="""" then
			CurrApp = mid(AllApps_org,2,instr(1,AllApps_org,""",",1)-2)
			AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,""",",1)-1)
		else
			CurrApp = mid(AllApps_org,1,instr(1,AllApps_org,",",1)-1)
			AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,",",1))
		end if
		'Check for apps that need to be renamed based on apprename table
		Set re = New RegExp
		for i = 0 to ubound(AppRenames,1)
			re.Pattern = AppRenames(i,0)
			If re.Test(CurrApp) then
				CurrApp = AppRenames(i,1)
				i = ubound(AppRenames,1)
			End if
		next
		'msgbox CurrApp
		'Ignore publisher
		if left(AllApps_Org,1)="""" then
			AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,""",",1)-1)
		else
			AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,",",1))
		end if
		'Get version
		if left(AllApps_Org,1)="""" then
			CurrVer = mid(AllApps_org,2,instr(1,AllApps_org,vbCrlf,1)-3)
			AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,vbCrlf,1)-1)
			if len(CurrVer) = 0 then CurrVer = "0" 'No version!
		elseif instr(1,AllApps_org,vbCrlf,1) - 1 <= 0 then
			CurrVer = "0"
			AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,vbCrlf,1)-1)
			'msgbox CurrApp & " No version!"
		else
			CurrVer = mid(AllApps_org,1,instr(1,AllApps_org,vbCrlf,1)-1)
			AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,vbCrlf,1)-1)
		end if
		'msgbox CurrVer
		
		str = "Select * from discoveredapplications where Name='" & CurrApp & "';"
		rs.Open str, adoconn, 3, 3 'OpenType, LockType
		if not rs.eof then
			rs.MoveFirst
			if len(rs("LastDiscovered") & "") = 0 then rs("LastDiscovered") = "2001-01-01" 'Fix DB issues
			if len(rs("FirstDiscovered") & "") = 0 then rs("FirstDiscovered") = format(date()-1, "YYYY-MM-DD") 'Fix DB issues
			if format(rs("LastDiscovered"), "YYYY-MM-DD") <> format(date(), "YYYY-MM-DD") then
				if len(rs("Version_Newest") & "") = 0 then rs("Version_Newest") = 0 'Fix DB issues
				rs("Version_Oldest") = rs("Version_Newest")
				rs("LastDiscovered") = format(date(), "YYYY-MM-DD")
				'msgbox "date"
			end if
			
			if isnumeric(replace(CurrVer,".","")) and isnumeric(replace(rs("Version_Oldest"),".","")) and isnumeric(replace(rs("Version_Newest"),".","")) then
				if CompareVersions(PadVersion(CurrVer), PadVersion(rs("Version_Oldest"))) = 2 then
					rs("Version_Oldest") = CurrVer
					'msgbox CurrApp & " Updated -"
				end if
				if CompareVersions(PadVersion(CurrVer), PadVersion(rs("Version_Newest"))) = 1 then
					rs("Version_Newest") = CurrVer
					'msgbox CurrApp & " Updated +"
				end if
			end if
			'msgbox "Got it"
			
			rs.update
			rs.close
		else
			rs.close
			yfound = False
			
			'Check existing software for similar apps
			CurrAppNoVer = replace(CurrApp, ".", "_")
			for i=0 to 9
				CurrAppNoVer = replace(CurrAppNoVer, i, "_")
			next
			i = 0
			do while right(CurrAppNoVer,1) = "_"
				CurrAppNoVer = left(CurrAppNoVer, len(CurrAppNoVer) - 1)
				i = 1
			loop
			if i = 1 then CurrAppNoVer = CurrAppNoVer & "%"
			str = "Select * from discoveredapplications where Name like '" & CurrAppNoVer & "';"
			rs.Open str, adoconn, 2, 1 'OpenType, LockType
			
			if not rs.eof then
				yfound = True
				'msgbox "New app - minor version change (1)" & vbCrlf & CurrApp
				
				rs.MoveFirst
				if len(rs("UpdateURL")) > 1 then
					if int(rs("UpdatePageQTH")) & 0 = 0 then UpdatePageQTH = 0 else UpdatePageQTH = int(rs("UpdatePageQTH")) 'Fix NULL entries in integer field
					if int(rs("UpdatePageQTHVarience")) & 0 = 0 then UpdatePageQTHVarience = 0 else UpdatePageQTHVarience = int(rs("UpdatePageQTHVarience")) 'Fix NULL entries in integer field
					if UpdatePageQTHVarience = 10 + len(rs("Version_Newest")) then UpdatePageQTHVarience = 10 + len(CurrVer) 'Update the varience if it has never been updated and the version length is different
					str = "INSERT INTO discoveredapplications(Name,Version_Oldest,Version_Newest,LastDiscovered,FirstDiscovered,Free,OpenSource,FOSS,ReasonForSoftware,NeededOnMachines,PlansForRemoval,`Update Method`,UpdateURL,UpdatePageQTH,UpdatePageQTHVarience) values('" & CurrApp & "','" & CurrVer & "','" & CurrVer & "','" & format(date(), "YYYY-MM-DD")  & "','" & format(date(), "YYYY-MM-DD") & "','" & rs("Free") & "','" & rs("OpenSource") & "','" & rs("FOSS") & "','" & rs("ReasonForSoftware") & "','" & rs("NeededOnMachines") & "','" & rs("PlansForRemoval") & "','" & rs("Update Method") & "','" & rs("UpdateURL") & "','" & UpdatePageQTH & "','" & UpdatePageQTHVarience & "');"
				else
					str = "INSERT INTO discoveredapplications(Name,Version_Oldest,Version_Newest,LastDiscovered,FirstDiscovered,Free,OpenSource,FOSS,ReasonForSoftware,NeededOnMachines,PlansForRemoval,`Update Method`) values('" & CurrApp & "','" & CurrVer & "','" & CurrVer & "','" & format(date(), "YYYY-MM-DD")  & "','" & format(date(), "YYYY-MM-DD") & "','" & rs("Free") & "','" & rs("OpenSource") & "','" & rs("FOSS") & "','" & rs("ReasonForSoftware") & "','" & rs("NeededOnMachines") & "','" & rs("PlansForRemoval") & "','" & rs("Update Method") & "');"
				end if
				adoconn.Execute(str)
			end if
			rs.close
			
			'Machine Learning App Renames
			if yfound = True then
				yfound = false 'temporary so we can use the variable
				Set re = New RegExp
				RenameEx = Replace(CurrAppNoVer,"_",".") 'To meet RegEx format
				RenameEx = Replace(RenameEx," %","*") 'To meet RegEx format
				RenameEx = Replace(RenameEx,"%","*") 'To meet RegEx format
				RenameEx = Replace(Replace(RenameEx,"(","\("),")","\)") 'Excape ( and )
				RenameTo = replace(replace(Replace(CurrAppNoVer," %",""),"%",""),"_","")
				
				str = "Select * from apprename where RenameTo = '" & RenameTo & "';"
				rs.Open str, adoconn, 3, 3 'OpenType, LockType
				
				if not rs.eof then
					rs.MoveFirst
				end if

				do while not rs.eof
					re.Pattern = rs("RegEx")
					If re.Test(CurrApp) then
						rs("Hits") = int(rs("Hits")) + 1
						yfound = True
						rs.update
						rs.movenext
					end if
				loop
				
				if yfound = false then
					str = "INSERT INTO apprename(RegEx,RenameTo,Hits,Confirmed) values('" & RenameEx & "','" & RenameTo & "','1','0');"
					adoconn.Execute(str)
				end if
				
				rs.close
				yfound = True 'Set is back to true
			end if
			
			if yfound = false then
				str = "Select * from discoveredapplications where Name like '" & left(CurrApp,len(CurrApp)/2) & "%' and not UpdateURL = '' and UpdateURL IS NOT NULL;"
				rs.Open str, adoconn, 2, 1 'OpenType, LockType
				
				if not rs.eof then
					rs.MoveFirst
					
					'Pull website
					On error resume next
					xmlhttp.open "get", rs("UpdateURL"), false
					xmlhttp.send
					WPData = xmlhttp.responseText
					
					'Check to see if exists
					if instr(1,WPData,CurrVer,0)>0 then
						yfound = True
						'msgbox "New app - major version change (2)" & vbCrlf & CurrApp
						
						UpdatePageQTHVarience = int(rs("UpdatePageQTHVarience"))
						if UpdatePageQTHVarience = 10 + len(rs("Version_Newest")) then UpdatePageQTHVarience = 10 + len(CurrVer) 'Update the varience if it has never been updated and the version length is different
						str = "INSERT INTO discoveredapplications(Name,Version_Oldest,Version_Newest,LastDiscovered,FirstDiscovered,Free,OpenSource,FOSS,ReasonForSoftware,NeededOnMachines,PlansForRemoval,`Update Method`,UpdateURL,UpdatePageQTH,UpdatePageQTHVarience) values('" & CurrApp & "','" & CurrVer & "','" & CurrVer & "','" & format(date(), "YYYY-MM-DD")  & "','" & format(date(), "YYYY-MM-DD") & "','" & rs("Free") & "','" & rs("OpenSource") & "','" & rs("FOSS") & "','" & rs("ReasonForSoftware") & "','" & rs("NeededOnMachines") & "','" & rs("PlansForRemoval") & "','" & rs("Update Method") & "','" & rs("UpdateURL") & "','" & instr(1,WPData,CurrVer,0) & "','" & UpdatePageQTHVarience & "');"
						adoconn.Execute(str)
						RenameTo = rs("Name")
					end if
				end if
				rs.close
				
				'Machine Learning App Renames
				if yfound = True then
					yfound = false 'temporary so we can use the variable
					Set re = New RegExp
					for i = 1 to len(CurrApp)
						if not left(CurrApp,i) = Left(RenameTo,i) then
							If right(RenameTo,1) = " " then i = i - 1 'Remove space at the end if there is one
							RenameTo = Left(RenameTo,i)
							RenameEx = Left(RenameTo,i) & "*"
							i = len(CurrApp)
						end if
					next
					str = "Select * from apprename where RenameTo = '" & RenameTo & "';"
					rs.Open str, adoconn, 3, 3 'OpenType, LockType
					
					if not rs.eof then
						rs.MoveFirst
					end if

					do while not rs.eof
						re.Pattern = rs("RegEx")
						If re.Test(CurrApp) then
							rs("Hits") = int(rs("Hits")) + 1
							yfound = True
							rs.update
							rs.movenext
						end if
					loop
					
					if yfound = false then
						str = "INSERT INTO apprename(RegEx,RenameTo,Hits,Confirmed) values('" & RenameEx & "','" & RenameTo & "','1','0');"
						adoconn.Execute(str)
					end if
					
					rs.close
					yfound = True 'Set is back to true
				end if
			end if
			
			if yfound = false then
				'msgbox "New app - brand new (3)" & vbCrlf & CurrApp
				CurrAppNoVer = replace(CurrApp, ".", "")
				CurrAppNoVer = replace(CurrAppNoVer, "x86", "")
				CurrAppNoVer = replace(CurrAppNoVer, "x64", "")
				CurrAppNoVer = replace(CurrAppNoVer, "(", "")
				CurrAppNoVer = replace(CurrAppNoVer, ")", "")
				for i=0 to 9
					CurrAppNoVer = replace(CurrAppNoVer, i, "")
				next
				CurrAppNoVer = trim(CurrAppNoVer)
				TestFOSSfree = ""
				TestFOSSOS = ""
				TestFOSS = ""
				
				'Test FOSS at OpenHub.net
				xmlhttp.open "get", "https://www.openhub.net/p?query=" & CurrAppNoVer, false
				xmlhttp.send
				WPData = xmlhttp.responseText
				if instr(1,WPData," - did not match anything.",1) = 0 then
					TestFOSSFree = "Y"
					TestFOSSOS = "Y"
				end if
				
				'Test FOSS at chocolatey.org
				xmlhttp.open "get", "https://chocolatey.org/packages?q=" & CurrAppNoVer, false
				xmlhttp.send
				WPData = xmlhttp.responseText
				if instr(1,WPData,"Returned 0 <text>Package</text>s",1) = 0 then
					TestFOSSFree = "Y"
				end if
				
				if TestFOSSFree = "Y" or TestFOSSOS = "Y" then TestFOSS = "Y"
		  
				str = "INSERT INTO discoveredapplications(Name,Version_Oldest,Version_Newest,LastDiscovered,FirstDiscovered,Free,OpenSource,FOSS,UpdatePageQTHVarience) values('" & CurrApp & "','" & CurrVer & "','" & CurrVer & "','" & format(date(), "YYYY-MM-DD")  & "','" & format(date(), "YYYY-MM-DD") & "','" & TestFOSSFree & "','" & TestFOSSOS & "','" & TestFOSS & "','" & (10 + len(CurrVer)) & "');"
				adoconn.Execute(str)
				
				'msgbox "Added: " & CurrApp & " - " & CurrVer
			end if
		end if
		
	loop
	
	'PCs - Whats new/old/changed
	'AllApps = right(AllApps,len(AllApps)-33)
	AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,vbCrlf,1)-1)
	do while len(AllApps) > 10
		'Get PC name
		if left(AllApps,1)="""" then
			CurrPC = mid(AllApps,2,instr(1,AllApps,""",",1)-2)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,""",",1)-1)
		else
			CurrPC = mid(AllApps,1,instr(1,AllApps,",",1)-1)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,",",1))
		end if
		'msgbox CurrPC
		'Get application
		if left(AllApps,1)="""" then
			CurrApp = mid(AllApps,2,instr(1,AllApps,""",",1)-2)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,""",",1)-1)
		else
			CurrApp = mid(AllApps,1,instr(1,AllApps,",",1)-1)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,",",1))
		end if
		'Check for apps that need to be renamed based on apprename table
		Set re = New RegExp
		for i = 0 to ubound(AppRenames,1)
			re.Pattern = AppRenames(i,0)
			If re.Test(CurrApp) then
				CurrApp = AppRenames(i,1)
				i = ubound(AppRenames,1)
			End if
		next
		'msgbox CurrApp
		'Get publisher
		if left(AllApps,1)="""" then
			CurrPub = mid(AllApps,2,instr(1,AllApps,""",",1)-2)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,""",",1)-1)
		else
			CurrPub = mid(AllApps,1,instr(1,AllApps,",",1)-1)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,",",1))
		end if
		'msgbox CurrPub
		'Get version
		if left(AllApps,1)="""" then
			CurrVer = mid(AllApps,2,instr(1,AllApps,vbCrlf,1)-3)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,vbCrlf,1)-1)
			if len(CurrVer) = 0 then CurrVer = "0" 'No version!
		elseif instr(1,AllApps,vbCrlf,1) - 1 <= 0 then
			CurrVer = "0"
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,vbCrlf,1)-1)
			'msgbox CurrApp & " No version!"
		else
			CurrVer = mid(AllApps,1,instr(1,AllApps,vbCrlf,1)-1)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,vbCrlf,1)-1)
		end if
		'msgbox CurrVer
		
		'msgbox CurrPC & vbCrlf & CurrApp & vbCrlf & CurrPub & vbCrlf & CurrVer
		
		str = "Select * from applicationsdump where Computer='" & CurrPC & "' and Name='" & CurrApp & "';"
		rs.Open str, adoconn, 3, 3 'OpenType, LockType
		if not rs.eof then
			rs.MoveFirst
			if len(rs("LastDiscovered") & "") = 0 then rs("LastDiscovered") = "2001-01-01" 'Fix DB issues
			if len(rs("FirstDiscovered") & "") = 0 then rs("FirstDiscovered") = format(date()-1, "YYYY-MM-DD") 'Fix DB issues
			if format(rs("LastDiscovered"), "YYYY-MM-DD") <> format(date(), "YYYY-MM-DD") then
				rs("LastDiscovered") = format(date(), "YYYY-MM-DD")
				'msgbox "date"
			end if
			
			if not rs("Version") = CurrVer then
				if instr(1,outputl,"<p><b>Software Added or Changed:</b></p>",1) = 0 then
					'Header Info
					outputl = outputl & "<p><b>Software Added or Changed:</b></p>" & vbcrlf
					outputl = outputl & "<table>" & vbcrlf
					outputl = outputl & "<tr>" & vbcrlf
					outputl = outputl & "  <th>Computer</th>" & vbcrlf
					outputl = outputl & "  <th>Application</th>" & vbcrlf
					outputl = outputl & "  <th>Publisher</th>" & vbcrlf
					outputl = outputl & "  <th>Previous Version</th>" & vbcrlf
					outputl = outputl & "  <th>New Version</th>" & vbcrlf
					outputl = outputl & "</tr>" & vbcrlf
				end if
				
				outputl = outputl & "<tr>" & vbcrlf
				outputl = outputl & "  <td>" & CurrPC & "</td>" & vbcrlf
				outputl = outputl & "  <td>" & CurrApp & "</td>" & vbcrlf
				outputl = outputl & "  <td>" & CurrPub & "</td>" & vbcrlf
				outputl = outputl & "  <td>" & rs("Version") & "</td>" & vbcrlf
				outputl = outputl & "  <td>" & CurrVer & "</td>" & vbcrlf
				outputl = outputl & "</tr>" & vbcrlf
				
				'msgbox CurrApp & ": Updated on " & CurrPC & " from " & rs("Version") & " to " & CurrVer
				rs("Version") = CurrVer
				rs("Publisher") = CurrPub
			end if
			'msgbox CurrPC & "|" & CurrApp & ": finished updating"
			
			rs.update
		else
			if instr(1,outputl,"<p><b>Software Added or Changed:</b></p>",1) = 0 then
				'Header Info
				outputl = outputl & "<p><b>Software Added or Changed:</b></p>" & vbcrlf
				outputl = outputl & "<table>" & vbcrlf
				outputl = outputl & "<tr>" & vbcrlf
				outputl = outputl & "  <th>Computer</th>" & vbcrlf
				outputl = outputl & "  <th>Application</th>" & vbcrlf
				outputl = outputl & "  <th>Publisher</th>" & vbcrlf
				outputl = outputl & "  <th>Previous Version</th>" & vbcrlf
				outputl = outputl & "  <th>New Version</th>" & vbcrlf
				outputl = outputl & "</tr>" & vbcrlf
			end if
			
			outputl = outputl & "<tr>" & vbcrlf
			outputl = outputl & "  <td>" & CurrPC & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & CurrApp & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & CurrPub & "</td>" & vbcrlf
			outputl = outputl & "  <td></td>" & vbcrlf
			outputl = outputl & "  <td>" & CurrVer & "</td>" & vbcrlf
			outputl = outputl & "</tr>" & vbcrlf
			
			str = "INSERT INTO applicationsdump(Computer,Name,Publisher,Version,LastDiscovered,FirstDiscovered) values('" & CurrPC & "','" & CurrApp & "','" & CurrPub & "','" & CurrVer & "','" & format(date(), "YYYY-MM-DD")  & "','" & format(date(), "YYYY-MM-DD") & "');"
			adoconn.Execute(str)
			
			'msgbox "Added: " & CurrPC & "|" & CurrApp & " - " & CurrVer
		end if
		rs.close
		
	loop
	
	if instr(1,outputl,"<p><b>Software Added or Changed:</b></p>",1) > 0 then outputl = outputl & "</table>" & vbcrlf
End function

Function Get_PC_Removed()
	str = "Select * from applicationsdump where LastDiscovered IS NOT NULL and not LastDiscovered = '" & format(date(), "YYYY-MM-DD") & "' order by Computer;"
	rs.Open str, adoconn, 3, 3 'OpenType, LockType
	if not rs.eof then
		'Header Info
		outputl = outputl & "<p><b>Software Removed:</b></p>" & vbcrlf
		outputl = outputl & "<table>" & vbcrlf
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <th>Computer</th>" & vbcrlf
		outputl = outputl & "  <th>Application</th>" & vbcrlf
		outputl = outputl & "  <th>Publisher</th>" & vbcrlf
		outputl = outputl & "  <th>Version</th>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
		
		rs.MoveFirst
	end if

	do while not rs.eof
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Computer") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Publisher") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Version") & "</td>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
		
		rs.delete
		rs.movenext
		if rs.eof then outputl = outputl & "</table>" & vbcrlf
	loop
	
	rs.close
End function

Function Get_Organization_New()
	str = "Select * from discoveredapplications where FirstDiscovered = '" & format(date(), "YYYY-MM-DD") & "' order by Name;"
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
	if not rs.eof then
		'Header Info
		outputl = outputl & "<p><b>New Software that has been added to organization:</b></p>" & vbcrlf
		outputl = outputl & "<table>" & vbcrlf
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <th>Application</th>" & vbcrlf
		outputl = outputl & "  <th>FOSS</th>" & vbcrlf
		outputl = outputl & "  <th>Purpose</th>" & vbcrlf
		outputl = outputl & "  <th>Reference ID</th>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
		
		rs.MoveFirst
	end if

	do while not rs.eof
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
		if rs("FOSS") & "" = "Y" then
			outputl = outputl & "  <td bgcolor=#00CCFF>Y</td>" & vbcrlf
		else
			outputl = outputl & "  <td>" & rs("FOSS") & "</td>" & vbcrlf
		end if
		outputl = outputl & "  <td>" & rs("ReasonForSoftware") & "</td>" & vbcrlf
		if BaseURL = "" then
			outputl = outputl & "  <td>" & rs("ID") & "</td>" & vbcrlf
		else
			outputl = outputl & "  <td><a href=""" & BaseURL & "/edit-appinfo.php?id=" & rs("ID") & """>" & rs("ID") & "</a></td>" & vbcrlf
		end if
		outputl = outputl & "</tr>" & vbcrlf
	
		rs.movenext
		if rs.eof then outputl = outputl & "</table>" & vbcrlf
	loop
	
	rs.close
End function

Function Get_Organization_Removed()
	Dim RenameConf()
	str = "Select * from discoveredapplications where LastDiscovered IS NOT NULL and not LastDiscovered = '" & format(date(), "YYYY-MM-DD") & "' order by LastDiscovered DESC, Name;"
	redim RenameConf(cint((adoconn.Execute(replace(str,"*","count(*)")))(0)) - 1)
	i = 0
	'msgbox UBound(RenameConf)
	
	rs.Open str, adoconn, 3, 3 'OpenType, LockType
	if not rs.eof then
		'Header Info
		outputl = outputl & "<p><b>Software that no longer exists in the organization:</b></p>" & vbcrlf
		outputl = outputl & "<table>" & vbcrlf
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <th>Application</th>" & vbcrlf
		outputl = outputl & "  <th>FOSS</th>" & vbcrlf
		outputl = outputl & "  <th>Purpose</th>" & vbcrlf
		outputl = outputl & "  <th>Last Seen</th>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
		
		rs.MoveFirst
	end if

	do while not rs.eof
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("FOSS") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("ReasonForSoftware") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("LastDiscovered") & "</td>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
	
		if cdate(rs("LastDiscovered")) < (Date() - 7) then
			RenameConf(i) = rs("Name")
			rs.delete
		Else
			RenameConf(i) = ""
		end if
		i = i + 1
		rs.movenext
		if rs.eof then outputl = outputl & "</table>" & vbcrlf
	loop
	
	rs.close
	
	'Machine Learning App Renames - Confirm
	if UBound(RenameConf) > 0 then
		Set re = New RegExp
		
		str = "Select * from apprename where Confirmed < 5;"
		rs.Open str, adoconn, 3, 3 'OpenType, LockType
		
		if not rs.eof then
			rs.MoveFirst
		end if

		do while not rs.eof
			re.Pattern = rs("RegEx")
			for i = 0 to UBound(RenameConf)
				If re.Test(RenameConf(i)) and not RenameConf(i) = "" then
					'msgbox rs("RenameTo") & vbCrlf & RenameConf(i)
					rs("Confirmed") = int(rs("Confirmed")) + 1
					rs.update
				end if
			next
			rs.movenext
	loop
	end if
End function

Function Get_App_Renames()	
	str = "select count(*) from apprename where Hits = 5 and Confirmed = 5;"
	redim AppRenames(cint((adoconn.Execute(str))(0)) - 1,1)
	
	str = "Select * from apprename where Hits = 5 and Confirmed = 5;"
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
		
	i = 0

	do while not rs.eof
		AppRenames(i,0) = rs("RegEx")
		AppRenames(i,1) = rs("RenameTo")
		i = i + 1
		
		rs.movenext
	loop
	
	rs.close
End function

Function PadVersion(InputVersion)
	Dim PaddedVersion
	
	j = 0
	for i = 1 to len(InputVersion)
		if mid(InputVersion,i,1) = "." Then
			'msgbox mid(InputVersion,i - j,j)
			if j < 6 then
				PaddedVersion = PaddedVersion & left("000000",6 - j) & mid(InputVersion,i - j,j)
			Else
				PaddedVersion = PaddedVersion & mid(InputVersion,i - j,j)
			end if
			j = - 1
		end if
		j = j + 1
	Next
	if j > 0 and j < 6 then
		PaddedVersion = PaddedVersion & left("000000",6 - j) & mid(InputVersion,len(InputVersion) - j + 1,j)
	Else
		if PaddedVersion = "" then
			PaddedVersion = InputVersion
		Else
			PaddedVersion = PaddedVersion & mid(InputVersion,len(InputVersion) - j + 1,j)
		end if
	end if
	
	'msgbox PaddedVersion

	PadVersion = PaddedVersion
End Function

Function CompareVersions(NumberOne, NumberTwo)
	Dim WinningNum
	
	WinningNum = 0

	If Len(NumberOne) > len(NumberTwo) Then
		WinningNum = 1
		for i = 1 to Len(NumberTwo)
			If mid(NumberOne,i,1) > mid(NumberTwo,i,1) Then
				WinningNum = 1
				i = len(NumberTwo)
			elseif mid(NumberOne,i,1) < mid(NumberTwo,i,1) then
				WinningNum = 2
				i = len(NumberTwo)
			end if
		next
	elseif Len(NumberOne) < len(NumberTwo) then
		WinningNum = 2
		for i = 1 to Len(NumberOne)
			If mid(NumberOne,i,1) > mid(NumberTwo,i,1) Then
				WinningNum = 1
				i = len(NumberOne)
			elseif mid(NumberOne,i,1) < mid(NumberTwo,i,1) then
				WinningNum = 2
				i = len(NumberOne)
			end if
		next
	Else
		for i = 1 to Len(NumberOne)
			If mid(NumberOne,i,1) > mid(NumberTwo,i,1) Then
				WinningNum = 1
				i = len(NumberOne)
			elseif mid(NumberOne,i,1) < mid(NumberTwo,i,1) then
				WinningNum = 2
				i = len(NumberOne)
			end if
		next
	end if
	
	'msgbox WinningNum
	
	CompareVersions = WinningNum
End Function 

Function SendMail(TextRcv,TextSubject)
  Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
  Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 

  Const cdoAnonymous = 0 'Do not authenticate
  Const cdoBasic = 1 'basic (clear-text) authentication
  Const cdoNTLM = 2 'NTLM

  Set objMessage = CreateObject("CDO.Message") 
  objMessage.Subject = TextSubject 
  objMessage.From = RptFromEmail 
  objMessage.To = TextRcv
  objMessage.HTMLBody = outputl

  '==This section provides the configuration information for the remote SMTP server.

  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

  'Name or IP of Remote SMTP Server
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailSvr
  
  'Type of authentication, NONE, Basic (Base64 encoded), NTLM
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = EmailAuthType
  
  if EmailAuthType > 0 then

	  'Your UserID on the SMTP server
	  objMessage.Configuration.Fields.Item _
	  ("http://schemas.microsoft.com/cdo/configuration/sendusername") = EmailUserName

	  'Your password on the SMTP server
	  objMessage.Configuration.Fields.Item _
	  ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = EmailPassword
  
  end if

  'Server port (typically 25)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = EmailPort

  'Use SSL for the connection (False or True)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

  'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

  objMessage.Configuration.Fields.Update

  '==End remote SMTP server configuration section==

  objMessage.Send
End Function

Function Format(vExpression, sFormat)
  Dim nExpression
  nExpression = sFormat
  
  if isnull(vExpression) = False then
    if instr(1,sFormat,"Y") > 0 or instr(1,sFormat,"M") > 0 or instr(1,sFormat,"D") > 0 or instr(1,sFormat,"H") > 0 or instr(1,sFormat,"S") > 0 then 'Time/Date Format
      vExpression = cdate(vExpression)
	  if instr(1,sFormat,"AM/PM") > 0 and int(hour(vExpression)) > 12 then
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression)-12,2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)-12) '1 character hour
		nExpression = replace(nExpression,"AM/PM","PM") 'If if its afternoon, its PM
	  else
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression),2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)) '1 character hour
		if int(hour(vExpression)) = 12 then nExpression = replace(nExpression,"AM/PM","PM") '12 noon is PM while anything else in this section is AM (fixed 04/19/2019 thanks to our HR Dept.)
		nExpression = replace(nExpression,"AM/PM","AM") 'If its not PM, its AM
	  end if
	  nExpression = replace(nExpression,":MM",":" & right("00" & minute(vExpression),2)) '2 character minute
	  nExpression = replace(nExpression,"SS",right("00" & second(vExpression),2)) '2 character second
	  nExpression = replace(nExpression,"YYYY",year(vExpression)) '4 character year
	  nExpression = replace(nExpression,"YY",right(year(vExpression),2)) '2 character year
	  nExpression = replace(nExpression,"DD",right("00" & day(vExpression),2)) '2 character day
	  nExpression = replace(nExpression,"D",day(vExpression)) '(N)N format day
	  nExpression = replace(nExpression,"MMM",left(MonthName(month(vExpression)),3)) '3 character month name
	  if instr(1,sFormat,"MM") > 0 then
	    nExpression = replace(nExpression,"MM",right("00" & month(vExpression),2)) '2 character month
	  else
	    nExpression = replace(nExpression,"M",month(vExpression)) '(N)N format month
	  end if
    elseif instr(1,sFormat,"N") > 0 then 'Number format
	  nExpression = vExpression
	  if instr(1,sFormat,".") > 0 then 'Decimal format
	    if instr(1,nExpression,".") > 0 then 'Both have decimals
		  do while instr(1,sFormat,".") > instr(1,nExpression,".")
		    nExpression = "0" & nExpression
		  loop
		  if len(nExpression)-instr(1,nExpression,".") >= len(sFormat)-instr(1,sFormat,".") then
		    nExpression = left(nExpression,instr(1,nExpression,".")+len(sFormat)-instr(1,sFormat,"."))
	      else
		    do while len(nExpression)-instr(1,nExpression,".") < len(sFormat)-instr(1,sFormat,".")
			  nExpression = nExpression & "0"
			loop
	      end if
		else
		  nExpression = nExpression & "."
		  do while len(nExpression) < len(sFormat)
			nExpression = nExpression & "0"
		  loop
	    end if
	  else
		do while len(nExpression) < sFormat
		  nExpression = "0" and nExpression
		loop
	  end if
	else
      msgbox "Formating issue on page. Unrecognized format: " & sFormat
	end if
	
	Format = nExpression
  end if
End Function

'Read text file
function GetFile(FileName)
  If FileName<>"" Then
    Dim FS, FileStream
    Set FS = CreateObject("Scripting.FileSystemObject")
      on error resume Next
      Set FileStream = FS.OpenTextFile(FileName)
      GetFile = FileStream.ReadAll
  End If
End Function

Function ReadIni( myFilePath, mySection, myKey ) 'Thanks to http://www.robvanderwoude.com
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function

Sub WriteIni( myFilePath, mySection, myKey, myValue ) 'Thanks to http://www.robvanderwoude.com
    ' This subroutine writes a value to an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be written
    ' myValue     [string]  the value to be written (myKey will be
    '                       deleted if myValue is <DELETE_THIS_VALUE>)
    '
    ' Returns:
    ' N/A
    '
    ' CAVEAT:     WriteIni function needs ReadIni function to run
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre, Johan Pol and Rob van der Woude

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
    Dim intEqualPos
    Dim objFSO, objNewIni, objOrgIni
    Dim strFilePath, strFolderPath, strKey, strLeftString
    Dim strLine, strSection, strTempDir, strTempFile, strValue

    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
    strValue    = Trim( myValue )

    Set objFSO   = CreateObject( "Scripting.FileSystemObject" )

    strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

    Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
    Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )

    blnInSection     = False
    blnSectionExists = False
    ' Check if the specified key already exists
    blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
    blnWritten       = False

    ' Check if path to INI file exists, quit if not
    strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
    If Not objFSO.FolderExists ( strFolderPath ) Then
        WScript.Echo "Error: WriteIni failed, folder path (" _
                   & strFolderPath & ") to ini file " _
                   & strFilePath & " not found!"
        Set objOrgIni = Nothing
        Set objNewIni = Nothing
        Set objFSO    = Nothing
        WScript.Quit 1
    End If

    While objOrgIni.AtEndOfStream = False
        strLine = Trim( objOrgIni.ReadLine )
        If blnWritten = False Then
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                blnSectionExists = True
                blnInSection = True
            ElseIf InStr( strLine, "[" ) = 1 Then
                blnInSection = False
            End If
        End If

        If blnInSection Then
            If blnKeyExists Then
                intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
                If intEqualPos > 0 Then
                    strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                    If LCase( strLeftString ) = LCase( strKey ) Then
                        ' Only write the key if the value isn't empty
                        ' Modification by Johan Pol
                        If strValue <> "<DELETE_THIS_VALUE>" Then
                            objNewIni.WriteLine strKey & "=" & strValue
                        End If
                        blnWritten   = True
                        blnInSection = False
                    End If
                End If
                If Not blnWritten Then
                    objNewIni.WriteLine strLine
                End If
            Else
                objNewIni.WriteLine strLine
                    ' Only write the key if the value isn't empty
                    ' Modification by Johan Pol
                    If strValue <> "<DELETE_THIS_VALUE>" Then
                        objNewIni.WriteLine strKey & "=" & strValue
                    End If
                blnWritten   = True
                blnInSection = False
            End If
        Else
            objNewIni.WriteLine strLine
        End If
    Wend

    If blnSectionExists = False Then ' section doesn't exist
        objNewIni.WriteLine
        objNewIni.WriteLine "[" & strSection & "]"
            ' Only write the key if the value isn't empty
            ' Modification by Johan Pol
            If strValue <> "<DELETE_THIS_VALUE>" Then
                objNewIni.WriteLine strKey & "=" & strValue
            End If
    End If

    objOrgIni.Close
    objNewIni.Close

    ' Delete old INI file
    objFSO.DeleteFile strFilePath, True
    ' Rename new INI file
    objFSO.MoveFile strTempFile, strFilePath

    Set objOrgIni = Nothing
    Set objNewIni = Nothing
    Set objFSO    = Nothing
End Sub

'Check to see if database and tables exist
Function CheckForTables()
	Dim CreatePS2DB 'Boolean for DB creation
	CreatePS2DB = False
	
	Set adoconn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & DBLocation & ";" & _
			"User=" & DBUser & "; Password=" & DBPass & ";"
			
	str = "SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '" & PSSchema & "'"
	rs.CursorLocation = 3 'adUseClient
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
	
	if rs.eof then
		Response = msgbox("The database does not exist. Would you like to create it now? (Make sure the user """ & DBUser & """ has permission to do so)", vbYesNo)
		if Response = vbYes then
			CreatePS2DB = True
		else
			WScript.Quit
		end if
		rs.close
	else
		'msgbox "DB exists"
		rs.close
		
		'Double check to make sure table is also there
		str = "SELECT * FROM information_schema.tables WHERE table_schema = '" & PSSchema & "' AND table_name = '" & PSTbl & "' LIMIT 1;"
		rs.Open str, adoconn, 2, 1 'OpenType, LockType
	
		if rs.eof then
			Response = msgbox("The database exists but the table does not exist. Would you like to create it now?", vbYesNo)
			if Response = vbYes then
				CreatePS2DB = True
			else
				WScript.Quit
			end if
			rs.close
		else
			'msgbox "Table exists"
			rs.close
		end if
	end if
	
	'Create schema and/or table if needed
	if CreatePS2DB = True then
		'Create schema if not there
		str = "CREATE DATABASE IF NOT EXISTS " & PSSchema & ";"
		adoconn.Execute(str)
		
		'Create tables
		PSTbl = "discoveredapplications"
		str = "CREATE TABLE " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, Name text, `Version_Oldest` text, `Version_Newest` text, Computers int(11) DEFAULT NULL, Free text, OpenSource text, FOSS text, ReasonForSoftware text, NeededOnMachines text, PlansForRemoval text, `Update Method` text, FirstDiscovered date DEFAULT NULL, LastDiscovered date DEFAULT NULL, UpdateURL text, UpdatePageQTH int(11) DEFAULT NULL, UpdatePageQTHVarience int(11) DEFAULT '10');"
		adoconn.Execute(str)
		
		PSTbl = "applicationsdump"
		str = "CREATE TABLE " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, Computer text, Name text, Publisher text, Version text, FirstDiscovered date DEFAULT NULL, LastDiscovered date DEFAULT NULL);"
		adoconn.Execute(str)
		
		PSTbl = "highriskapps"
		str = "CREATE TABLE " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, Name text, DateAdded date DEFAULT NULL, Source text);"
		adoconn.Execute(str)
		
		PSTbl = "licensedapps"
		str = "CREATE TABLE " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, Name VARCHAR(255) UNIQUE, Publisher text, Amount int(11) DEFAULT NULL, Comments text);"
		adoconn.Execute(str)
		
		PSTbl = "apprename"
		str = "CREATE TABLE " & PSSchema & "." & PSTbl & " (ID INT PRIMARY KEY AUTO_INCREMENT, RegEx text, RenameTo text, Hits INT NOT NULL DEFAULT '0', Confirmed INT NOT NULL DEFAULT '0');"
		adoconn.Execute(str)
	end if
	
	Set adoconn = Nothing
	Set rs = Nothing
End Function