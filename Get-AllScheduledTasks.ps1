#Requires -Modules ActiveDirectory
Import-Module ActiveDirectory

# AD OU to search for computers
$OUDistinguishedName = "OU=Virtual,OU=Desktops,OU=MyDomain,DC=mydomain,DC=com"

Function Connect-TaskScheduler
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [String]$ComputerName
    )

    Begin
    {
        $objScheduledTask = New-Object -ComObject("Schedule.Service")
    }
    Process
    {
        Try
        {
            Write-Verbose "Connecting to $ComputerName"
            $objScheduledTask.Connect("$ComputerName")
            Write-Verbose "Connected: $($objScheduledTask.Connected)"
        }
        Catch
        {
            Throw "Failed to connect to $ComputerName"
        }
    }
    End
    {
        Return $objScheduledTask
    }
}

Function Get-AllScheduledTasks
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [ValidateScript({If ($_.GetType().BaseType.Name -eq "MarshalByRefObject" -and $_.Connected -eq $True){$True}Else{Throw "Not a valid Task Scheduler connection object"}})]
        $Session,

        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Path = "\",

        [Parameter()]
        [Switch]$Recurse = $False
    )

    Begin
    {
        $Tasks = @()
        $Paths = @()
    }

    Process
    {
        $Paths += Get-TaskSchedulerPaths -Session $Session -Path $Path -Recurse:$Recurse
        $Tasks += $Paths | ForEach-Object {$_.Path | Get-TaskSchedulerTasks -Session $Session}
    }
    End
    {
        Return $Tasks
    }
}

Function Get-TaskSchedulerPaths
{
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$True)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Path = "\",

        [Parameter(Mandatory=$True)]
        [ValidateScript({If ($_.GetType().BaseType.Name -eq "MarshalByRefObject" -and $_.Connected -eq $True){$True}Else{Throw "Not a valid Task Scheduler connection object"}})]
        $Session,

        [Parameter()]
        [Switch]$Recurse = $False
    )

    Begin
    {
        $Paths = @()
    }
    Process
    {
        $BasePath = $Session.GetFolder("$Path")
        $Paths += $BasePath
        $Paths += $BasePath.GetFolders(1)
        If ($Recurse)
        {
            ForEach ($P in $Paths)
            {
                If ($P -eq $BasePath)
                {
                    Continue
                }
                Else
                {
                    $Paths += Get-TaskSchedulerPaths -Session $Session -Path $P.Path -Recurse
                }
            }
        }
    }
    End
    {
        Return $Paths
    }
}

Function Get-TaskSchedulerTasks
{
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline=$True)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Path = "\",

        [Parameter(Mandatory=$True)]
        [ValidateScript({If ($_.GetType().BaseType.Name -eq "MarshalByRefObject" -and $_.Connected -eq $True){$True}Else{Throw "Not a valid Task Scheduler connection object"}})]
        $Session
    )

    Begin
    {
        $Domain = (Get-WmiObject Win32_ComputerSystem -ComputerName $Session.TargetServer).Domain
        $AllTasks = @()
    }

    Process
    {
        $Folder = $Session.GetFolder("$Path")
        $FolderTasks = $Folder.GetTasks(0)

        ForEach ($Task in $FolderTasks)
        {
            $RunAsID = $Task | Select-Object @{Name="RunAs";Expression={[xml]$xml = $_.xml ; $xml.Task.Principals.Principal.UserId}} | Select-Object -ExpandProperty RunAs -ErrorAction Stop
            If ($RunAsID -match "S-\d-\d-\d\d.-\d*")
            {
                Try
                {
                    $RunAs = Get-ADUser -Identity $RunAsID | Select-Object -ExpandProperty SamAccountName -ErrorAction Stop
                }
                Catch
                {
                    Try
                    {
                        $RunAs = Get-ADGroup -Identity $RunAsID | Select-Object -ExpandProperty SamAccountName -ErrorAction Stop
                    }
                    Catch
                    {
                        $RunAs = $RunAsID
                    }
                }
            }
            Else
            {
                $RunAs = $RunAsID
            }
            $HashTable = [Ordered]@{
                ComputerName = $Session.TargetServer
                Name = $Task.Name
                Path = $Task.Path
                RunAs = $RunAs
                LastRunTime = $Task.LastRunTime
                NextRunTime = $Task.NextRunTime
                TaskEnabled = $Task.Enabled
            }
            $AllTasks += New-Object PSObject -Property $HashTable
        }
    }

    End
    {
        Return $AllTasks
    }
}

$Tasks = @()

ForEach ($Computer in (Get-ADComputer -Filter * -Properties OperatingSystem -SearchBase $OUDistinguishedName -SearchScope Subtree | Where-Object {$_.Enabled -eq $True -and $_.OperatingSystem -like "*Windows*"} | Select-Object -ExpandProperty Name | Sort-Object))
{
    $Connection = Connect-TaskScheduler -ComputerName $Computer -Verbose
    $Tasks += Get-AllScheduledTasks -Session $Connection -Recurse
}

$Tasks | Export-Csv -LiteralPath "$PSScriptRoot\ScheduledTasks.csv" -NoTypeInformation
