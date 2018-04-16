dim OVersion 'Oldest Installed version
dim CVersion 'Installed version
dim outputl 'Email body
Dim AllApps 'Data from CSV
dim WPData 'Web page text
Dim adoconn
Dim rs
Dim str
set filesys=CreateObject("Scripting.FileSystemObject")
set xmlhttp = createobject("msxml2.xmlhttp.3.0")

'''''''''''''''''''
'Required Variables

'Database
CSVPath = "C:\SoftwareMatrix\Applications.csv" 'Full path to the CSV file
DBPass = "P@ssword1" 'Password to access database on localhost

'Email - Defaults to anonymous login
RptToEmail = "admin@company.com" 'Report email's To address
RptFromEmail = "admin@company.com" 'Report email's From address
EmailSvr = "mail.server.com" 'FQDN or IP address of email server
'Additional email settings found in Function SendMail()

'''''''''''''''''''

outputl = ""

If filesys.FileExists(CSVPath) then
	AllApps = GetFile(CSVPath)

	if len(AllApps) > 100 then
		Set adoconn = CreateObject("ADODB.Connection")
		Set rs = CreateObject("ADODB.Recordset")
		adoconn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=localhost;" & _
					   "Database=software_matrix; User=root; Password=" & DBPass & ";"

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
			AllApps_Org = right(AllApps_Org,len(AllApps_Org)-instr(1,AllApps_Org,vbCrlf,1)-3)
		elseif instr(1,AllApps_org,vbCrlf,1) - 1 =< 0 then
			CurrVer = "0"
			'msgbox CurrApp & " No version!"
		else
			CurrVer = mid(AllApps_org,1,instr(1,AllApps_org,vbCrlf,1)-1)
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
				if int(replace(CurrVer,".","")) < int(replace(rs("Version_Oldest"),".","")) then
					rs("Version_Oldest") = CurrVer
					'msgbox CurrApp & " Updated -"
				end if
				if int(replace(CurrVer,".","")) > int(replace(rs("Version_Newest"),".","")) then
					rs("Version_Newest") = CurrVer
					'msgbox CurrApp & " Updated +"
				end if
			end if
			'msgbox "Got it"
			
			rs.update
		else
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
			
			'Test FOSS at FOSShub.com
			xmlhttp.open "get", "https://www.fosshub.com/search/" & CurrAppNoVer, false
			xmlhttp.send
			WPData = xmlhttp.responseText
			if instr(1,WPData,"There is <span>0</span> app",1) = 0 then
				TestFOSSFree = "Y"
				TestFOSSOS = "Y"
			end if
			
			'Test FOSS at chocolatey.org
			xmlhttp.open "get", "https://chocolatey.org/packages?q=" & CurrAppNoVer, false
			xmlhttp.send
			WPData = xmlhttp.responseText
			if instr(1,WPData,"returned 0 packages",1) = 0 then
				TestFOSSFree = "Y"
			end if
			
			if TestFOSSFree = "Y" or TestFOSSOS = "Y" then TestFOSS = "Y"
	  
			str = "INSERT INTO discoveredapplications(Name,Version_Oldest,Version_Newest,LastDiscovered,FirstDiscovered,Free,OpenSource,FOSS) values('" & CurrApp & "','" & CurrVer & "','" & CurrVer & "','" & format(date(), "YYYY-MM-DD")  & "','" & format(date(), "YYYY-MM-DD") & "','" & TestFOSSFree & "','" & TestFOSSOS & "','" & TestFOSS & "');"
			adoconn.Execute(str)
			
			'msgbox "Added: " & CurrApp & " - " & CurrVer
		end if
		rs.close
		
	loop
	
	'PCs - Whats new/old/changed
	AllApps = right(AllApps,len(AllApps)-33)
	do while len(AllApps) > 10
		'Get PC name
		CurrPC = mid(AllApps,1,instr(1,AllApps,",",1)-1)
		AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,",",1))
		'msgbox CurrPC
		'Get application
		if left(AllApps,1)="""" then
			CurrApp = mid(AllApps,2,instr(1,AllApps,""",",1)-2)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,""",",1)-1)
		else
			CurrApp = mid(AllApps,1,instr(1,AllApps,",",1)-1)
			AllApps = right(AllApps,len(AllApps)-instr(1,AllApps,",",1))
		end if
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
		elseif instr(1,AllApps,vbCrlf,1) - 1 =< 0 then
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
	str = "Select * from applicationsdump where not LastDiscovered = '' and LastDiscovered IS NOT NULL and not LastDiscovered = '" & format(date(), "YYYY-MM-DD") & "' order by Computer;"
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
		outputl = outputl & "  <td>" & rs("ID") & "</td>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
	
		rs.movenext
		if rs.eof then outputl = outputl & "</table>" & vbcrlf
	loop
	
	rs.close
End function

Function Get_Organization_Removed()
	str = "Select * from discoveredapplications where not LastDiscovered = '' and LastDiscovered IS NOT NULL and not LastDiscovered = '" & format(date(), "YYYY-MM-DD") & "' order by Name;"
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
	
		if cdate(rs("LastDiscovered")) < (Date() - 7) then rs.delete
		rs.movenext
		if rs.eof then outputl = outputl & "</table>" & vbcrlf
	loop
	
	rs.close
End function

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
  ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoAnonymous

  'Server port (typically 25)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

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