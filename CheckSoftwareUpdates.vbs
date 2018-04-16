dim WPURL 'Web page URL
dim WPData 'Web page text
dim WPQTH 'Update location
dim WPVar 'Location Varience
dim OVersion 'Oldest Installed version
dim CVersion 'Installed version
dim outputl 'Email body
dim URLERR 'URLS with errors
set xmlhttp = createobject("msxml2.xmlhttp.3.0")
Dim adoconn
Dim rs
Dim str

'''''''''''''''''''
'Required Variables

'Database
DBPass = "P@ssword1" 'Password to access database on localhost

'Email - Defaults to anonymous login
RptToEmail = "admin@company.com" 'Report email's To address
RptFromEmail = "admin@company.com" 'Report email's From address
EmailSvr = "mail.server.com" 'FQDN or IP address of email server
'Additional email settings found in Function SendMail()

'''''''''''''''''''

outputl = "There is currently no out-of-date software."
URLERR = ""

Set adoconn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
adoconn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=localhost;" & _
                   "Database=software_matrix; User=root; Password=" & DBPass & ";"
				   
str = "Select * from discoveredapplications where (not UpdateURL = '' and UpdateURL IS NOT NULL) or Version_Oldest <> Version_Newest order by Name;"
rs.Open str, adoconn, 3, 3 'OpenType, LockType
if not rs.eof then rs.MoveFirst

do while not rs.eof
	WPURL = rs("UpdateURL")
	WPQTH = rs("UpdatePageQTH")
	WPVar = rs("UpdatePageQTHVarience")
	OVersion = rs("Version_Oldest")
	CVersion = rs("Version_Newest")

	if len(WPURL) > 0 then
	  'Pull website
	  On error resume next
	  xmlhttp.open "get", WPURL, false
	  xmlhttp.send
	  WPData = xmlhttp.responseText
	  if WPData = "" then URLERR = URLERR & WPURL & "<br>" & vbCrlf

	  'Check to see if exists
	  if instr(1,WPData,CVersion,0)>0 then
		if instr(1,WPData,CVersion,0) => WPQTH + WPVar or instr(1,WPData,CVersion,0) =< WPQTH - WPVar then
			if outputl = "There is currently no out-of-date software." then
				'Header Info
				outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body> <table>" & vbcrlf
				outputl = outputl & "<tr>" & vbcrlf
				outputl = outputl & "  <th>Application</th>" & vbcrlf
				outputl = outputl & "  <th>Oldest Version</th>" & vbcrlf
				outputl = outputl & "  <th>Newest Version</th>" & vbcrlf
				outputl = outputl & "  <th>Varience</th>" & vbcrlf
				outputl = outputl & "  <th>Update URL</th>" & vbcrlf
				outputl = outputl & "</tr>" & vbcrlf
			end if
		
			outputl = outputl & "<tr>" & vbcrlf
			outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & OVersion & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & CVersion & "</td>" & vbcrlf
			if instr(1,WPData,CVersion,0) => WPQTH + WPVar then
				outputl = outputl & "  <td bgcolor=yellow>" & instr(1,WPData,CVersion,0) - WPQTH & "</td>" & vbcrlf
			elseif instr(1,WPData,CVersion,0) =< WPQTH - WPVar then
				outputl = outputl & "  <td bgcolor=yellow>" & WPQTH - instr(1,WPData,CVersion,0) & "</td>" & vbcrlf
			end if
			outputl = outputl & "  <td><a href=""" & WPURL & """>Download</a></td>" & vbcrlf  
			outputl = outputl & "</tr>" & vbcrlf
		elseif OVersion <> CVersion then
			if outputl = "There is currently no out-of-date software." then
				'Header Info
				outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body> <table>" & vbcrlf
				outputl = outputl & "<tr>" & vbcrlf
				outputl = outputl & "  <th>Application</th>" & vbcrlf
				outputl = outputl & "  <th>Oldest Version</th>" & vbcrlf
				outputl = outputl & "  <th>Newest Version</th>" & vbcrlf
				outputl = outputl & "  <th>Varience</th>" & vbcrlf
				outputl = outputl & "  <th>Update URL</th>" & vbcrlf
				outputl = outputl & "</tr>" & vbcrlf
			end if
			
			outputl = outputl & "<tr>" & vbcrlf
			outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
			outputl = outputl & "  <td bgcolor=yellow>" & OVersion & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & CVersion & "</td>" & vbcrlf
			outputl = outputl & "  <td bgcolor=green>0</td>" & vbcrlf
			outputl = outputl & "  <td><a href=""" & WPURL & """>Download</a></td>" & vbcrlf
			outputl = outputl & "</tr>" & vbcrlf
		end if
		
		'msgbox len(WPQTH & "")
		'msgbox instr(1,WPData,CVersion,0)
		if len(WPQTH & "") = 0 or WPQTH = 0 then
		  rs("UpdatePageQTH") = instr(1,WPData,CVersion,0)
		  rs.update
		end if
	  
		'msgbox rs("Name") & ": The installed version, " & CVersion & ", is the latest version."
	  else
		if outputl = "There is currently no out-of-date software." then
			'Header Info
			outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body> <table>" & vbcrlf
			outputl = outputl & "<tr>" & vbcrlf
			outputl = outputl & "  <th>Application</th>" & vbcrlf
			outputl = outputl & "  <th>Oldest Version</th>" & vbcrlf
			outputl = outputl & "  <th>Newest Version</th>" & vbcrlf
			outputl = outputl & "  <th>Varience</th>" & vbcrlf
			outputl = outputl & "  <th>Update URL</th>" & vbcrlf
			outputl = outputl & "</tr>" & vbcrlf
		end if
		
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & OVersion & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & CVersion & "</td>" & vbcrlf
		outputl = outputl & "  <td bgcolor=red>N/A</td>" & vbcrlf
		outputl = outputl & "  <td><a href=""" & WPURL & """>Download</a></td>" & vbcrlf  
		outputl = outputl & "</tr>" & vbcrlf
		
		'msgbox rs("Name") & ": The installed version, " & CVersion & ", is not the latest version. Please download new version at: " & WPURL
	  End if
	else
		if outputl = "There is currently no out-of-date software." then
			'Header Info
			outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body> <table>" & vbcrlf
			outputl = outputl & "<tr>" & vbcrlf
			outputl = outputl & "  <th>Application</th>" & vbcrlf
			outputl = outputl & "  <th>Oldest Version</th>" & vbcrlf
			outputl = outputl & "  <th>Newest Version</th>" & vbcrlf
			outputl = outputl & "  <th>Varience</th>" & vbcrlf
			outputl = outputl & "  <th>Update URL</th>" & vbcrlf
			outputl = outputl & "</tr>" & vbcrlf
		end if
		
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
		outputl = outputl & "  <td bgcolor=yellow>" & OVersion & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & CVersion & "</td>" & vbcrlf
		outputl = outputl & "  <td>N/A</td>" & vbcrlf
		outputl = outputl & "  <td></td>" & vbcrlf  
		outputl = outputl & "</tr>" & vbcrlf
	end if
	rs.MoveNext
loop

rs.close

'Clean up and output errors to bottom of email
outputl = outputl & "</table>"
if not URLERR = "" then outputl = outputl & "<br><br>URLs with Errors:<br>" & URLERR 

SendMail RptToEmail, "Software Matrix: Update Report"

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