Dim CurrID, CurrApp, CurrVer, CurrFree, CurrOS, CurrReason, CurrPC, CurrPlans, CurrUpdate, CurrURL, CurrQTH, CurrVar
Dim adoconn
Dim rs
Dim str
dim WPData 'Web page text
dim xmlhttp : set xmlhttp = createobject("msxml2.xmlhttp.3.0")
set filesys=CreateObject("Scripting.FileSystemObject")
Dim strCurDir
strCurDir = filesys.GetParentFolderName(Wscript.ScriptFullName)

'Gather variables from smapp.ini
If filesys.FileExists(strCurDir & "\smapp.ini") then
	'Database
	DBPass = ReadIni(strCurDir & "\smapp.ini", "Database", "DBPass" )
else
	msgbox "INI file not found at: " & strCurDir & "\smapp.ini" & vbCrlf & "Please run IngestCSV.vbs first before running this file."
end if

'Ask for App Reference ID
CurrID = inputbox("Enter the reference ID of the application that you would like to update:", "Software Matrix", "")

if len(CurrID) > 0 and isnumeric(CurrID) then
	Set adoconn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	adoconn.Open "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;" & _
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
