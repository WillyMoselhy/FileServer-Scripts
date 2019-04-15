[CmdletBinding()]
param
(
[Parameter(Mandatory=$True,
ValueFromPipeline=$True,
ValueFromPipelineByPropertyName=$True,
    HelpMessage='Input File Path')]
[Alias('host')]
[ValidateLength(3,30)]
[string]$Path,
		
[string]$ComputerName
)

process {

write-verbose "Looking for open files."
    
    $ErrorActionPreference = "stop"
    #Set default parameter if computer is defined
    if($ComputerName){
        $CIMSession = New-CimSession -ComputerName $ComputerName
        $PSDefaultParameterValues = @{
            "*SMB*:CIMSession" = $CIMSession 
        }
    }
    else{
        $PSDefaultParameterValues = @{}
    }
    

    #Query for Open Files in this path
	$OpenFiles = Get-SmbOpenFile | Where-Object {$_.Path -like "*$Path*"}
    
    #Logic if no open files found
    if(!($OpenFiles)){
        return "No open files found at this path"
    }
    
    #Query for sessions related to this path
    $SMBSessions = foreach ($SessionID in ($OpenFiles.SessionId | select -Unique)){
        $Session = Get-SmbSession -SessionId  $SessionID
        $ClientName = $null
        $ClientName = Resolve-DnsName -DnsOnly -Name $Session.ClientComputerName -ErrorAction SilentlyContinue
        if($ClientName){
            $ClientName = $ClientName.NameHost
        }

        [PSCustomObject]@{
            SessionID       = $Session.SessionId
            CLientIPAddress = $Session.ClientUserName
            ClientName      = $ClientName
            IdleTime        = $Session.SecondsIdle
            IdleTimeminutes = "$(New-TimeSpan -Seconds $Session.SecondsIdle)" -f {hh:mm:ss}

        }
        
        
    }
    "Found the following open files"
    $FileSession = foreach ($File in $OpenFiles){
        $Session = $SMBSessions | Where-Object {$_.SessionID -eq $File.SessionID}
        [PSCustomObject]@{
            Path            = $File.Path
            Username        = $File.ClientUserName
            ClientIPAddress = $File.ClientComputerName
            ClientName      = $Session.ClientName
            SessionIdleTime = $Session.IdleTimeminutes
        }
    }
    
    
    $FileSession |ft

    $Title = "Action"
    $Info = "What would you like to do?"
 
    $options = [System.Management.Automation.Host.ChoiceDescription[]] @(
            "Close &All open files", 
            "Close files for specific &User", 
            "Close files for specific &IP Addres",
            "&Export list to CSV"
            "&Quit"
        )
    [int]$defaultchoice = 4
    $opt = $host.UI.PromptForChoice($Title , $Info , $Options,$defaultchoice)
    switch($opt)
    {
    0 { 
        Write-Host "Close all open files" -ForegroundColor Green
        $OpenFiles | Close-SmbOpenFile -Force -ErrorAction Continue
      }
    1 { Write-Host "Close files for specific user" -ForegroundColor Green
        $Username = Read-Host -Prompt "Please enter username"
        $UserCheck = $OpenFiles | Where-Object {$_.ClientUserName -like "*$Username"} | select -First 1
        if ($UserCheck){
            $Username = $UserCheck.ClientUserName
        }
        else{
            Write-Error "User '$Username' does not have any open files"
        }
        
        $UserFiles = $OpenFiles | Where-Object {$_.ClientUserName -eq $Username}
        if ($UserFiles){
            $UserFiles | Close-SmbOpenFile -Force -ErrorAction Continue
            Write-Host "Files closed successfully" -ForegroundColor Green
        }
        else{
            Write-Error "User '$Username' does not have any open files"
        }
      }
    2 { 
        Write-Host "Close files for specific IP Address" -ForegroundColor Green
        $IPAddress = Read-Host -Prompt "Please enter IP Address"
        $IPFiles = $OpenFiles | Where-Object {$_.ClientComputerName -eq $IPAddress}
        if ($IPFiles){
            $IPFiles | Close-SmbOpenFile -Force -ErrorAction Continue
            Write-Host "Files closed successfully" -ForegroundColor Green
        }
        else{
            Write-Error "IP Address '$IPAddress' does not have any open files"
        }

      }
    3 { 
        Write-Host "Export list to CSV" -ForegroundColor Green
        $ExportPath = Read-Host "Please input file path"
        $FileSession| Export-Csv -Path $ExportPath -NoTypeInformation
      }

    4 { Write-Host "Good Bye!" -ForegroundColor Green}
    }
  
}
