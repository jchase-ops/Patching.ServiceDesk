# .ExternalHelp $PSScriptRoot\Move-ChangeOrder-help.xml
function Move-ChangeOrder {

    [CmdletBinding()]

    Param (

        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Object]
        $Session = $script:Config.Session,

        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]
        $SessionID = $script:Config.SessionID,

        [Parameter(Mandatory, Position = 2)]
        [ValidatePattern("^\d{4,10}$")]
        [System.String]
        $ChangeOrder,

        [Parameter(Position = 3)]
        [ValidateNotNull()]
        [AllowEmptyString()]
        [System.String]
        $Description = '',

        [Parameter(Position = 4)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Assignee,

        [Parameter(Position = 5)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Group,

        [Parameter(Position = 6)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Organization,
        
        [Parameter()]
        [Switch]
        $Quiet
    )

    $windowVisible = if ($(Get-Process -Id $([System.Diagnostics.Process]::GetCurrentProcess().Id)).MainWindowHandle -eq 0) { $false } else { $true }

    if ($null -eq $script:Config.Session -or ($null -eq $script:Config.SessionID -or $script:Config.SessionID -eq 0)) {
        if (!($windowVisible) -or $Quiet) {
            Connect-ServiceDesk -Quiet
        }
        else {
            Connect-ServiceDesk
        }
        $Session = $script:Config.Session
        $SessionID = $script:Config.SessionID
    }
    else {
        if ($Session.serverStatus($SessionID) -ne 0) {
            if (!($windowVisible) -or $Quiet) {
                Connect-ServiceDesk -Quiet
            }
            else {
                Connect-ServiceDesk
            }
            $Session = $script:Config.Session
            $SessionID = $script:Config.SessionID
        }
    }

    $creator_handle = $Session.getHandleForUserId($SessionID, "$($Session.Credentials.Domain)\$($Session.Credentials.UserName)")
    $co_handle = Get-ServiceDeskHandle -ObjectType chg -WhereClause "chg_ref_num = '$ChangeOrder'"

    $set = [PSCustomObject]@{
        Assignee     = $false
        Group        = $false
        Organization = $false
    }

    $setHandles = [PSCustomObject]@{
        Assignee     = $null
        Group        = $null
        Organization = $null
    }

    'Assignee', 'Group', 'Organization' | ForEach-Object {
        if ($PSBoundParameters[$($_)]) {
            $set.$($_) = $true
            Switch ($_) {
                'Assignee' { $setHandles.$($_) = $Session.getHandleForUserId($SessionID, $Assignee) }
                'Group' { $setHandles.$($_) = Get-ServiceDeskHandle -ObjectType grp -WhereClause "last_name = '$Group'" }
                'Organization' { $setHandles.$($_) = Get-ServiceDeskHandle -ObjectType org -WhereClause "name = '$Organization'" }
            }
        }
    }

    $handle = ([xml]$Session.transfer($SessionID, $creator_handle, $co_handle, $Description, $set.Assignee, $setHandles.Assignee, $set.Group, $setHandles.Group, $set.Organization, $setHandles.Organization)).UDSObject.Handle

    if ($MyInvocation.MyCommand.Name -eq $(Split-Path -Path $PSCommandPath -Leaf).TrimEnd('.ps1')) {
        if (!($windowVisible) -or $Quiet) {
            Disconnect-ServiceDesk -Quiet
        }
        else {
            Disconnect-ServiceDesk
        }
    }

    $handle
}
