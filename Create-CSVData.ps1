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
$PSSchema = "software_matrix"
$PSTbl = "smcrawler"

$CSVData = @()

#Open MySQL Connection
$myconnection = New-Object MySql.Data.MySqlClient.MySqlConnection
$myconnection.ConnectionString = "Database=" + $PSSchema + ";server=" + $Server + ";Persist Security Info=false;user id=" + $PSUser + ";pwd=" + $PSPassword + ";"
$myconnection.Open()
$command = $myconnection.CreateCommand()

#Query Database before creating CSV
$command.CommandText = "select * from " + $PSTbl + ";"
$reader = $command.ExecuteReader()

while ($reader.Read()) {
	$obj = New-Object PSObject
	$obj | Add-Member -MemberType NoteProperty -Name "Computer" -Value $Reader["Computer"].ToString()
	$obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $Reader["Name"].ToString()
	$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $Reader["Publisher"].ToString()
	$obj | Add-Member -MemberType NoteProperty -Name "Version" -Value $Reader["Version"].ToString()

	$CSVData += $obj
}
$reader.close()

#Save data to CSV
$CSVData | export-csv $IniContent["Database"]["CSVPath"] -notypeinformation

#Close MySQL Connection
$myconnection.Close()

#Run VBS
start-process ($ScriptFolder + "\IngestCSV.vbs")