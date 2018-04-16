Dim CurrID, CurrApp, CurrVer, CurrFree, CurrOS, CurrReason, CurrPC, CurrPlans, CurrUpdate, CurrURL, CurrQTH, CurrVar
Dim adoconn
Dim rs
Dim str
set xmlhttp = createobject("msxml2.xmlhttp.3.0")
dim WPURL 'Web page URL
dim WPData 'Web page text

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

WPURL = "https://nvd.nist.gov/download/nvd-rss-analyzed.xml"

xmlhttp.open "get", WPURL, false
xmlhttp.send
WPData = xmlhttp.responseText

outputl = "No installed applications matched vulnerabilities added within the last week."

if len(WPData) > 100 then
	Set adoconn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	adoconn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=localhost;" & _
                   "Database=software_matrix; User=root; Password=" & DBPass & ";"
	
	str = "Select * from discoveredapplications;"
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
	rs.movefirst
	
	Do while not rs.eof
		if instr(1,WPData,rs("Name"),1) > 0 then
		
			if 	outputl = "No installed applications matched vulnerabilities added within the last week." then
				'Header Info
				outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body> <table>" & vbcrlf
				outputl = outputl & "<tr>" & vbcrlf
				outputl = outputl & "  <th>Application</th>" & vbcrlf
				outputl = outputl & "  <th>Oldest Version</th>" & vbcrlf
				outputl = outputl & "  <th>Newest Version</th>" & vbcrlf
				outputl = outputl & "  <th>App Found</th>" & vbcrlf
				outputl = outputl & "  <th>Version Found</th>" & vbcrlf
				outputl = outputl & "</tr>" & vbcrlf
			end if
			
			outputl = outputl & "<tr>" & vbcrlf
			outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & rs("Version_Oldest") & "</td>" & vbcrlf
			outputl = outputl & "  <td>" & rs("Version_Newest") & "</td>" & vbcrlf
			outputl = outputl & "  <td>Y</td>" & vbcrlf
			if instr(1,WPData,rs("Version_Newest"),1) > 0 and instr(1,WPData,rs("Version_Newest"),1) - instr(1,WPData,rs("Name"),1) < 50 then
				outputl = outputl & "  <td>Y</td>" & vbcrlf
			else
				outputl = outputl & "  <td>N</td>" & vbcrlf
			end if
		End if
		rs.movenext
	loop
	rs.close
end if

if not outputl = "No installed applications matched vulnerabilities added within the last week." then
	outputl = outputl & "</table>"
	outputl = outputl & "<br><br>" & vbCrlf & vbCrlf & "Vulnerability database is accessible at: <a href=https://nvd.nist.gov/>https://nvd.nist.gov/</a>"
end if

SendMail RptToEmail, "Software Matrix: Vulnerability Report"

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