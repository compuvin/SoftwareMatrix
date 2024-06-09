[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")


function Get-IniContent ($filePath)
{
    $ini = @{}
    switch -regex -file $FilePath
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        "^(;.*)$" # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = “Comment” + $CommentCount
            $ini[$section][$name] = $value
        }
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}

$ScriptFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Source
$IniContent = Get-IniContent ($ScriptFolder + "\smapp.ini")

[String] $Server = $IniContent["Database"]["DBLocation"]
$PSUser = $IniContent["Database"]["DBUser"]
$PSPassword = $IniContent["Database"]["DBPass"]
$IgnoreVulnerabilities = "," + $IniContent["AppSpecific"]["IgnoreVulnerabilities"] + ","
$PSSchema = "software_matrix"

#If nothing matches, this text will be sent in the email
$NoMatchesEmailBody = "No installed applications matched vulnerabilities added within the last period."

#Open MySQL Connection
$myconnection = New-Object MySql.Data.MySqlClient.MySqlConnection
$myconnection.ConnectionString = "Database=" + $PSSchema + ";server=" + $Server + ";Persist Security Info=false;user id=" + $PSUser + ";pwd=" + $PSPassword + ";"
$myconnection.Open()
$command = $myconnection.CreateCommand()

#Correct DateAdded column to be datetime rather than just date
$command.CommandText = "SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE table_name = 'highriskapps' AND COLUMN_NAME = 'DateAdded';"
if ($command.ExecuteScalar() -eq "date") {
	$command.CommandText = "ALTER TABLE `software_matrix`.`highriskapps` CHANGE COLUMN `DateAdded` `DateAdded` DATETIME NULL DEFAULT NULL;"
	$command.ExecuteNonQuery()
}

#Get the date of the last vulnerability in the highriskapps table (our database)
$command.CommandText = "SELECT max(DateAdded) FROM highriskapps;"
$LastVulnDate = $command.ExecuteScalar()

#Check to see if this script has been run (and has found anything that has matched) within the last 60 days
#60 days worth of vulnerabilities are quite a lot so this assures the script will finish within a reasonable time
if ($LastVulnDate -lt (Get-Date).AddDays(-60)) {$LastVulnDate = (Get-Date).AddDays(-60)}

#Get latest vulnerability data
$Vuln = Invoke-RestMethod "https://services.nvd.nist.gov/rest/json/cves/2.0?noRejected&lastModStartDate=$($LastVulnDate.ToString(""yyyy-MM-ddTHH:mm:ss.fffK""))&lastModEndDate=$((get-date).ToUniversalTime().ToString(""yyyy-MM-ddTHH:mm:ss.fffK""))" -Method GET
$VulnCVE = $Vuln.vulnerabilities.cve

#If there are more than the maximum of 2,000 records
$startIndex = 2000
while ($($Vuln.totalResults -1) -gt ($VulnCVE | measure).count) {
	$Vuln = Invoke-RestMethod "https://services.nvd.nist.gov/rest/json/cves/2.0?noRejected&lastModStartDate=$($LastVulnDate.ToString(""yyyy-MM-ddTHH:mm:ss.fffK""))&lastModEndDate=$((get-date).ToUniversalTime().ToString(""yyyy-MM-ddTHH:mm:ss.fffK""))&startIndex=$startIndex" -Method GET
	$VulnCVE += $Vuln.vulnerabilities.cve
	$startIndex = $startIndex + 2000
	Write-Output "Total: $($Vuln.totalResults), Gathered: $(($VulnCVE | measure).count)"
	sleep 10
}


#Query Database
$command.CommandText = "select * from discoveredapplications;"
$reader = $command.ExecuteReader()

$MatchedVulns = @()

while ($reader.Read()) {
	$AppFriendlyName = $Reader['Name'].ToString().replace("+","\+").Replace("'","\'") #Excape out some symbols that will mess up the search
	if (($IgnoreVulnerabilities | Select-String ",$AppFriendlyName," | measure).count -eq 0 -and ($VulnCVE | where {$_.descriptions -like "*$AppFriendlyName*"} | measure).count -gt 0) #instr(1,WPData,rs("Name"),1) > 0 and instr(1,IgnoreVulnerabilities,"," & rs("Name") & ",",1) = 0 then
	{
		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $Reader["Name"].ToString()
		$obj | Add-Member -MemberType NoteProperty -Name "Version_Oldest" -Value $Reader["Version_Oldest"].ToString()
		$obj | Add-Member -MemberType NoteProperty -Name "Version_Newest" -Value $Reader["Version_Newest"].ToString()
		$obj | Add-Member -MemberType NoteProperty -Name "DateAdded" -Value ($VulnCVE | where {$_.descriptions -like "*$($Reader['Name'].ToString())*"} | Sort-Object lastModified | select -last 1).lastModified
		$obj | Add-Member -MemberType NoteProperty -Name "VersionFound" -Value $(if (($VulnCVE | where {$_.descriptions -like "*$($Reader['Name'].ToString())*" -and $_.descriptions -like "*$($Reader['Version_Newest'].ToString())*"} | Measure-Object).count -gt 0) {"Y"} else {"N"})
		$obj | Add-Member -MemberType NoteProperty -Name "Source" -Value "NIST"

		$MatchedVulns += $obj
	}
}
$reader.close()

#Prepare the email body
if (($MatchedVulns | Measure-Object).count -eq 0) {
	$outputl = $NoMatchesEmailBody
} else {
	$outputl = "<html><head> <style>BODY{font-family: Arial; font-size: 10pt;}TABLE{border: 1px solid black; border-collapse: collapse;}TH{border: 1px solid black; background: #dddddd; padding: 5px; }TD{border: 1px solid black; padding: 5px; }</style> </head><body> <table>`n"
	$outputl = $outputl + "<tr>`n"
	$outputl = $outputl + "  <th>Application</th>`n"
	$outputl = $outputl + "  <th>Oldest Version</th>`n"
	$outputl = $outputl + "  <th>Newest Version</th>`n"
	$outputl = $outputl + "  <th>App Found</th>`n"
	$outputl = $outputl + "  <th>Version Found</th>`n"
	$outputl = $outputl + "</tr>`n"
}

#Sort the matched vulnerabilities by date
$MatchedVulns = $MatchedVulns | Sort-Object DateAdded

#Report on vulnerabilities and enter into DB
foreach ($Vulnerability in $MatchedVulns)
{
	$outputl = $outputl + "<tr>`n"
	$outputl = $outputl + "  <td>" + $Vulnerability.Name + "</td>`n"
	$outputl = $outputl + "  <td>" + $Vulnerability.Version_Oldest + "</td>`n"
	$outputl = $outputl + "  <td>" + $Vulnerability.Version_Newest + "</td>`n"
	$outputl = $outputl + "  <td>Y</td>`n"
	$outputl = $outputl + "  <td>" + $Vulnerability.VersionFound + "</td>`n"
	
	#Write new found vulnerability to High Risk database
	$command.CommandText = "INSERT INTO highriskapps(Name,DateAdded,Source) values('$($Vulnerability.Name)','$($Vulnerability.DateAdded)','$($Vulnerability.Source)');"
	$reader = $command.ExecuteNonQuery()
}

#Add footer to email
if (($MatchedVulns | Measure-Object).count -gt 0) {
	$outputl = $outputl + "</table>`n"
	$outputl = $outputl + "<br><br>`n`nVulnerability database is accessible at: <a href=https://nvd.nist.gov/>https://nvd.nist.gov/</a>"
}

#Send email
$Cred = New-Object System.Management.Automation.PSCredential ($IniContent["Email"]["EmailUserName"], $($IniContent["Email"]["EmailPassword"] | ConvertTo-SecureString -AsPlainText -Force))
Send-MailMessage -From $IniContent["Email"]["RptFromEmail"] -To ([string[]]($IniContent["Email"]["RptToEmail"]).Split(',')) -SmtpServer $IniContent["Email"]["EmailSvr"] -Port $IniContent["Email"]["EmailPort"] -Credential $Cred -Subject "Software Matrix: Vulnerability Report" -Body $outputl -BodyAsHtml

#Close MySQL Connection
$myconnection.Close()
