Dim adoconn
Dim rs
Dim str
Dim i 'Counter
dim outputl 'Email body
dim CountName 'Count for each app
set filesys=CreateObject("Scripting.FileSystemObject")
Dim strCurDir
strCurDir = filesys.GetParentFolderName(Wscript.ScriptFullName)

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


outputl = ""


Set adoconn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")
adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;" & _
		   "Database=software_matrix; User=root; Password=" & DBPass & ";"

CountApps 'Count apps and update Computers column on discoveredapplications table
CheckLicenses 'Check licensed apps vs. actual installed count

if outputl <> "" then
	outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body>" & vbcrlf & outputl
	SendMail RptToEmail, "Software Matrix: Licensing Report"
	outputl = ""
end if

CountHighRiskApps 'Report on top 10 high risk apps

if outputl <> "" then
	outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body>" & vbcrlf & outputl
	SendMail RptToEmail, "Software Matrix: High Risk Software Report"
	outputl = ""
end if

Set adoconn = Nothing
Set rs = Nothing

Function CountApps()
	str = "Select * from discoveredapplications where not LastDiscovered = '' and LastDiscovered IS NOT NULL and LastDiscovered > '" & format(date() - 7, "YYYY-MM-DD") & "' order by Name;"
	rs.Open str, adoconn, 3, 3 'OpenType, LockType

	do while not rs.eof
		str = "select count(*) from applicationsdump where Name = '" & rs("Name") & "';"
		CountName = adoconn.Execute(str) 'Kind of a hack way of doing this, results in an array with 0 being the count
		rs("Computers") = CountName(0)
		'msgbox rs("Name") & vbCrlf & CountName(0)
		
		rs.update
		rs.movenext
	loop
	
	rs.close
End Function

Function CheckLicenses()
	str = "SELECT L.Name, L.Publisher, L.Amount, D.Computers FROM software_matrix.licensedapps as L inner join software_matrix.discoveredapplications as D on L.Name = D.Name and D.Computers > L.Amount order by L.Name;"
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
	if not rs.eof then
		'Header Info
		outputl = outputl & "<p><b>Software Licensing Report:</b></p>" & vbcrlf
		outputl = outputl & "<table>" & vbcrlf
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <th>Name</th>" & vbcrlf
		outputl = outputl & "  <th>Publisher</th>" & vbcrlf
		outputl = outputl & "  <th>Licensed Amount</th>" & vbcrlf
		outputl = outputl & "  <th>Installed Amount</th>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
		
		rs.MoveFirst
	end if

	do while not rs.eof
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Publisher") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Amount") & "</td>" & vbcrlf
		outputl = outputl & "  <td bgcolor=yellow>" & rs("Computers") & "</td>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
	
		rs.movenext
		if rs.eof then outputl = outputl & "</table>" & vbcrlf
	loop
	
	rs.close
End Function

Function CountHighRiskApps()
	str = "select Name, count(Name), max(DateAdded), (select Computers from software_matrix.discoveredapplications where discoveredapplications.Name = highriskapps.Name) Computers from highriskapps where DateAdded > '" & format(date() - 365, "YYYY-MM-DD") & "' group by Name order by count(Name) DESC, max(DateAdded) DESC, Computers DESC;"
	rs.Open str, adoconn, 2, 1 'OpenType, LockType
	
	if not rs.eof then
		'Header Info
		outputl = outputl & "<p><b>High Risk Application Report (top 10):</b></p>" & vbcrlf
		outputl = outputl & "<table>" & vbcrlf
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <th>Name</th>" & vbcrlf
		outputl = outputl & "  <th>Vulnerabilities prior year</th>" & vbcrlf
		outputl = outputl & "  <th>Date of Last Vulnerability</th>" & vbcrlf
		outputl = outputl & "  <th>Installed Amount</th>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
		
		rs.MoveFirst
	end if

	i = 1
	do while not rs.eof and not i > 10
		outputl = outputl & "<tr>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Name") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("count(Name)") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("max(DateAdded)") & "</td>" & vbcrlf
		outputl = outputl & "  <td>" & rs("Computers") & "</td>" & vbcrlf
		outputl = outputl & "</tr>" & vbcrlf
	
		i = i + 1
		rs.movenext
		if rs.eof or i > 10 then outputl = outputl & "</table>" & vbcrlf
	loop
	
	rs.close
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