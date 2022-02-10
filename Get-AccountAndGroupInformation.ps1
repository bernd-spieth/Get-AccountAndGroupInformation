<#
    This script collects information about the local computer and saves the collected data into a file share
#>

#region variables
$computerName = $env:COMPUTERNAME
$currentDate = Get-Date -Format "yyyy-MM-dd--HH-mm-ss"

$exportRootFolder = "C:\Temp"

$accountInformation = @()
$accountInformationRootFolder = Join-Path $exportRootFolder "Accounts"
$accountInformationExportFileName = ("{0}-Accounts.csv" -f $computerName)
$accountInformationExportPath = Join-Path $accountInformationRootFolder $accountInformationExportFileName

$serviceInformation = @()
$serviceInformationRootFolder = Join-Path $exportRootFolder "Services"
$serviceInformationExportFileName = ("{0}-Services.csv" -f $computerName)
$serviceInformationExportPath = Join-Path $serviceInformationRootFolder $serviceInformationExportFileName

$localGroupSIDs = @("S-1-5-32-555", "S-1-5-32-544")

# Check which PowerShell version we are running on because we collect service information differntly in versions
$global:powerShellMajorVersion = ($PSVersionTable).PSVersion.Major

#endregion

#region functions

function Get-GroupMemberEntries
{
    <#
        Create new custom objects from group members that can easily be eyported to a csv file
    #>
    param
    (
        [parameter(Mandatory=$true)]$GroupMembers,
        [parameter(Mandatory=$true)]$GroupName,
        [parameter(Mandatory=$true)]$ComputerName,
        [parameter(Mandatory=$true)]$Date
    )

    $csvEntries = @()

    foreach($groupMember in $GroupMembers)
    {
        $csvEntry = [pscustomobject]@{

            Computername        = $ComputerName
            Type                = $groupMember.ObjectClass
            PrincipalSource     = $groupMember.PrincipalSource
            Name                = $groupMember.Name
            GroupName           = $GroupName
            CollectionDate      = $Date
        }

        $csvEntries += $csvEntry
    }

    $csvEntries
}


function Get-AccountType
{
    <#
        Get the type of a group member (user or group)
    #>
    param
    (
        [parameter(Mandatory=$true)]$GroupMember
    )

    $type = "Unknown"

    try
    {
        switch($GroupMember.PartComponent.CimClass.CimClassName)
        {
            "Win32_UserAccount"
            {
                $type = "User"
                break
            }
            "Win32_SystemAccount"
            {
                $type = "User"
                break
            }
            "Win32_Group"
            {
                $type = "Group"
                break
            }
        }
    }
    catch
    {

    }

    $type
}

function Get-PrincipalSource
{
    <#
        Get the source (Local or ActiveDirectory) of a group member
    #>
    param
    (
        [parameter(Mandatory=$true)]$GroupMember
    )

    $source = "Unknown"

    try
    {
        if($GroupMember.PartComponent.Domain -eq $env:COMPUTERNAME)
        {
            $source = "Local"
        }
        else
        {
            $source = "ActiveDirectory"
        }
    }
    catch
    {

    }

    $source
}

function Get-GroupMemberName
{
    <#
        Get the name of a group member
    #>
    param
    (
        [parameter(Mandatory=$true)]$GroupMember
    )

    $name = "Unknown"

    try
    {
        if($GroupMember.PartComponent.Domain -eq $env:COMPUTERNAME)
        {
            $name = $GroupMember.PartComponent.Name
        }
        else
        {
            $name = $GroupMember.PartComponent.Domain + "\" + $GroupMember.PartComponent.Name
        }
    }
    catch
    {

    }

    $name
}

function Get-LocalGroupMemberPre2016
{
    <#
        Get the members of a group on a pre Server 2016 system
    #>

    param
    (
        [parameter(Mandatory=$true)]$GroupName,
        [parameter(Mandatory=$true)]$ComputerName,
        [parameter(Mandatory=$true)]$Date
    )

    # If the group name has a prefix remove it first
    $groupName = $GroupName.Substring($GroupName.IndexOf("\") + 1)

    $result = @()

    try
    {
        $members = Get-CimInstance -ClassName Win32_GroupUser

        foreach($member in $members)
        {
            if($member.GroupComponent.Name -eq $groupName)
            {
                $type = Get-AccountType $member
                $principalSource = Get-PrincipalSource $member
                $name = Get-GroupMemberName $member 

                $csvEntry = [pscustomobject]@{

                    Computername        = $ComputerName
                    Type                = $type
                    PrincipalSource     = $principalSource
                    Name                = $name
                    GroupName           = $groupName
                    CollectionDate      = $Date
                }

                $result +=  $csvEntry
            }
        }
    }
    catch
    {

    }

    $result
}

function Get-ServiceEntries
{
    <#
        Create new custom objects from the services
    #>
    param
    (
        [parameter(Mandatory=$true)]$Services,
        [parameter(Mandatory=$true)]$ComputerName,
        [parameter(Mandatory=$true)]$Date
    )

    $csvEntries = @()

    foreach($service in $services)
    {
        $csvEntry = [pscustomobject]@{

            Computername        = $ComputerName
            CollectionDate      = $Date
            Name                = $service.Name
            AccountName         = $Service.UserName
            StartType           = $Service.StartType
            Status              = $Service.Status
            BinaryPathName      = $service.BinaryPathName
            DisplayName         = $service.DisplayName
            Description         = $service.Description
        }

        $csvEntries += $csvEntry
    }

    $csvEntries
}

function Get-GroupNameForSID
{
    param
    (
        [parameter(Mandatory=$true)]$SID
    )

    $result = $SID

    try{
        $sidAsObject = New-Object System.Security.Principal.SecurityIdentifier($SID)
        $group = $sidAsObject.Translate([System.Security.Principal.NTAccount])
        $result = $group.Value
    }
    catch {

        $e = $_.Exception
        $message = $e.Message
    
        while ($e.InnerException){
            $e = $e.InnerException
            $message += "`n" + $e.Message
        }

        $message
    }

    $result
}

function Get-ServicesForExport
{
    # The cmdlets used to get the services are different between PowerShell and PowerShell Core
    $services = @()
    $serviceInformation = @()

    if($global:powerShellMajorVersion -le 5)
    {
        $services = Get-WmiObject Win32_Service | Select-Object Name, @{N='UserName'; E={$_.StartName}}, @{N='StartType'; E={$_.StartMode}}, @{N='Status'; E={$_.State}}, @{N='BinaryPathName'; E={$_.PathName}}, DisplayName, Description
    }
    elseif($global:powerShellMajorVersion -gt 5)
    {
        $services = Get-Service | Select-Object Name, UserName, StartType, Status, BinaryPathName, DisplayName, Description
    }

    foreach ($service in $services)
    {
        $serviceInformation += Get-ServiceEntries -Services $service -ComputerName $computerName -Date $currentDate
    }

    $serviceInformation
}

function Get-GroupMembersForExport
{
    param
    (
        [parameter(Mandatory=$true)]$LocalGroupSIDs
    )

    $accountInformation = @()

    if($global:powerShellMajorVersion -le 5)
    {
        # Get the members of local group
        foreach ($localGroupSID in $LocalGroupSIDs)
        {
            $groupName = Get-GroupNameForSID $localGroupSID
            $accountInformation += Get-LocalGroupMemberPre2016 -GroupName $groupName -ComputerName $computerName -Date $currentDate
        }
    }
    elseif($global:powerShellMajorVersion -gt 5)
    {
        # Get the members of local group
        foreach ($localGroupSID in $LocalGroupSIDs)
        {
            $groupName = Get-GroupNameForSID $localGroupSID
            $users = Get-LocalGroupMember -SID $localGroupSID

            if($null -ne $users)
            {
                $accountInformation += Get-GroupMemberEntries -GroupMembers $users -GroupName $groupName -ComputerName $computerName -Date $currentDate
            }
        }
    }

    $accountInformation
}
#endregions

# Create the paths if they do not exist
if(!(Test-Path $accountInformationRootFolder)){New-Item -Path $accountInformationRootFolder -ItemType "Directory" -ErrorAction SilentlyContinue}
if(!(Test-Path $serviceInformationRootFolder)){New-Item -Path $serviceInformationRootFolder -ItemType "Directory" -ErrorAction SilentlyContinue}


# Export the members of the given groups
$accountInformation = Get-GroupMembersForExport -LocalGroupSIDs $localGroupSIDs

# Export to csv file. If there is already a file with this name delete it first
if(Test-Path $accountInformationExportPath){Remove-Item -Path $accountInformationExportPath}
$accountInformation | Export-Csv -Path $accountInformationExportPath -Force -NoClobber -NoTypeInformation -Encoding UTF8 -Delimiter `t


# Export all servies
$serviceInformation = Get-ServicesForExport

# Export to csv file. If the file already exists delete first
if(Test-Path $serviceInformationExportPath){Remove-Item -Path $serviceInformationExportPath}
$serviceInformation | Export-Csv -Path $serviceInformationExportPath -Force -NoClobber -NoTypeInformation -Encoding UTF8 -Delimiter `t
