# .ExternalHelp $PSScriptRoot\Set-ChangeOrderStatus-help.xml
function Set-ChangeOrderStatus {

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

        [Parameter(Mandatory, Position = 3)]
        [ValidateNotNull()]
        [AllowEmptyString()]
        [System.String]
        $Description = '',

        [Parameter(Position = 4)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Status,

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

    $coHandle = Get-ServiceDeskHandle -ObjectType chg -WhereClause "chg_ref_num = '$ChangeOrder'"
    $creatorHandle = $Session.getHandleForUserId($SessionID, "$($Session.Credentials.Domain)\$($Session.Credentials.UserName)")
    $statusHandle = Get-ServiceDeskHandle -ObjectType chgstat -WhereClause "sym = '$Status'"
    $handle = ([xml]$Session.changeStatus($SessionID, $creatorHandle, $coHandle, $Description, $statusHandle)).UDSObjectList.UDSObject.Handle

    if ($($MyInvocation.MyCommand.Name) -eq $((Split-Path -Path $PSCommandPath -Leaf).TrimEnd('.ps1'))) {
        if (!($windowVisible) -or $Quiet) {
            Disconnect-ServiceDesk -Quiet
        }
        else {
            Disconnect-ServiceDesk
        }
    }

    $handle
}
