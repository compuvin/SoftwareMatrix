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
$ParentFolder = Split-Path -Path $ScriptFolder -Parent -erroraction SilentlyContinue
if ((split-path -Leaf -path $ScriptFolder) -eq "Optional") {
	if (test-Path "$ParentFolder\smapp.ini") {
		$IniContent = Get-IniContent "$ParentFolder\smapp.ini"
	} else {
		Write-Error "Main Software Matrix scripts detected! Run Ingest-CSV first"
		exit
	}
} else {
	if (test-Path "$ScriptFolder\smapp.ini") {
		$IniContent = Get-IniContent "$ScriptFolder\smapp.ini"
	} else {
		"[Database]`r`nDBPass=P@ssword1`r`nDBUser=user`r`nDBLocation=localhost" | Out-File "$ScriptFolder\smapp.ini"
		Write-Error "First run detected! Edit smapp.ini"
		exit
	}
}

[String] $Server = $IniContent["Database"]["DBLocation"]
$PSUser = $IniContent["Database"]["DBUser"]
$PSPassword = $IniContent["Database"]["DBPass"]
$PSSchema = "software_matrix"
$PSTbl = "smcrawler"

$UpdateDate = Get-Date -Format "yyyy-MM-dd"

$array = @()

$reg=[microsoft.win32.registrykey]::OpenBaseKey('LocalMachine',0) 

#32bit programs
$UninstallKey="SOFTWARE\\WOW6432Node\Microsoft\\Windows\\CurrentVersion\\Uninstall"
$regkey=$reg.OpenSubKey($UninstallKey) 
$subkeys=$regkey.GetSubKeyNames() 

foreach($key in $subkeys){

	$thisKey=$UninstallKey+"\\"+$key 

	$thisSubKey=$reg.OpenSubKey($thisKey) 
	
	if (!$thisSubKey.GetValue("SystemComponent") -eq 1) {

		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $env:computername
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
		$obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
		$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))

		$array += $obj
	}
}

#64bit programs
$UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
$regkey=$reg.OpenSubKey($UninstallKey) 
$subkeys=$regkey.GetSubKeyNames() 

foreach($key in $subkeys){

	$thisKey=$UninstallKey+"\\"+$key 

	$thisSubKey=$reg.OpenSubKey($thisKey) 

	if (!$thisSubKey.GetValue("SystemComponent") -eq 1) {

		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $env:computername
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
		$obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
		$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))

		$array += $obj
	}

}

#Programs in user space
$PatternSID = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'
 
# Get Username, SID, and location of ntuser.dat for all users
$ProfileList = gp 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object {$_.PSChildName -match $PatternSID} | 
    Select  @{name="SID";expression={$_.PSChildName}}, 
            @{name="UserHive";expression={"$($_.ProfileImagePath)\ntuser.dat"}}, 
            @{name="Username";expression={$_.ProfileImagePath -replace '^(.*[\\\/])', ''}}
 
# Get all user SIDs found in HKEY_USERS (ntuder.dat files that are loaded)
$LoadedHives = gci Registry::HKEY_USERS | ? {$_.PSChildname -match $PatternSID} | Select @{name="SID";expression={$_.PSChildName}}
 
# Get all users that are not currently logged
$UnloadedHives = Compare-Object $ProfileList.SID $LoadedHives.SID | Select @{name="SID";expression={$_.InputObject}}, UserHive, Username
 
# Loop through each profile on the machine
Foreach ($item in $ProfileList) {
    # Load User ntuser.dat if it's not already loaded
    IF ($item.SID -in $UnloadedHives.SID) {
        reg load HKU\$($Item.SID) $($Item.UserHive) | Out-Null
    }
 
    #####################################################################
    # This is where you can read/modify a users portion of the registry 
 
    # This example lists the Uninstall keys for each user registry hive
    "{0}" -f $($item.Username) | Write-Output
    Get-ItemProperty registry::HKEY_USERS\$($Item.SID)\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Foreach {
			
			$obj = New-Object PSObject
			$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $env:computername
			$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($_.DisplayName)
			$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($_.DisplayVersion)
			$obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($_.InstallLocation)
			$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($_.Publisher)

			$array += $obj
		}
    Get-ItemProperty registry::HKEY_USERS\$($Item.SID)\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Foreach {
			
			$obj = New-Object PSObject
			$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $env:computername
			$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($_.DisplayName)
			$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($_.DisplayVersion)
			$obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($_.InstallLocation)
			$obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($_.Publisher)

			$array += $obj
		}
    
    #####################################################################
 
    # Unload ntuser.dat 
    IF ($item.SID -in $UnloadedHives.SID) {
        ### Garbage collection and closing of ntuser.dat ###
        [gc]::Collect()
        reg unload HKU\$($Item.SID) | Out-Null
    }
}

$array = $array | Where-Object { $_.DisplayName }
$array | select ComputerName, DisplayName, DisplayVersion, Publisher | ft -auto

#Open MySQL Connection
$myconnection = New-Object MySql.Data.MySqlClient.MySqlConnection
$myconnection.ConnectionString = "Database=" + $PSSchema + ";server=" + $Server + ";Persist Security Info=false;user id=" + $PSUser + ";pwd=" + $PSPassword + ";"
$myconnection.Open()
$command = $myconnection.CreateCommand()

#Create table for this module if it doesn't exist
$command.CommandText = "CREATE TABLE IF NOT EXISTS " + $PSSchema + "." + $PSTbl + " (ID INT PRIMARY KEY AUTO_INCREMENT, Computer Text, Name text, Publisher text, Version text, LastDiscovered date DEFAULT NULL)";
$reader = $command.ExecuteNonQuery()

#Clear old data for this device
$command.CommandText = "DELETE FROM " + $PSSchema + "." + $PSTbl + " WHERE (`Computer` = '" + $env:computername + "')";
$reader = $command.ExecuteNonQuery()

#Add/Update table entries
foreach ($row in $array)
{
	#$row
	if ($row.DisplayVersion + "" -ne "") {$row.DisplayVersion = $row.DisplayVersion.replace("`0","")}
	$command.CommandText = "INSERT INTO " + $PSSchema + "." + $PSTbl + "(Computer, Name, Publisher, Version, LastDiscovered) values('" + $row.ComputerName.ToUpper() + "','" + $row.DisplayName.Replace("'","\'") + "','" + $row.Publisher + "','" + $row.DisplayVersion + "','" + $UpdateDate + "')";
	$reader = $command.ExecuteNonQuery()
}


#Close MySQL Connection
$myconnection.Close()