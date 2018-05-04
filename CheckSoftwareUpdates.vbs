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
set filesys=CreateObject("Scripting.FileSystemObject")
Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir = WshShell.CurrentDirectory

'Gather variables from smapp.ini
If filesys.FileExists(strCurDir & "\smapp.ini") then
	'Database
	DBPass = ReadIni(strCurDir & "\smapp.ini", "Database", "DBPass" )
	
	'Email - Defaults to anonymous login
	RptToEmail = ReadIni(strCurDir & "\smapp.ini", "Email", "RptToEmail" )
	RptFromEmail = ReadIni(strCurDir & "\smapp.ini", "Email", "RptFromEmail" )
	EmailSvr = ReadIni(strCurDir & "\smapp.ini", "Email", "EmailSvr" )
	'Additional email settings found in Function SendMail()
else
	msgbox "INI file not found at: " & strCurDir & "\smapp.ini" & vbCrlf & "Please run IngestCSV.vbs first before running this file."
end if


outputl = "There is currently no out-of-date software."
URLERR = ""

Set adoconn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;" & _
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