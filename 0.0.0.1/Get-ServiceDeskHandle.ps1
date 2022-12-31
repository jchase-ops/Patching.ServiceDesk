# .ExternalHelp $PSScriptRoot\Get-ServiceDeskHandle-help.xml
function Get-ServiceDeskHandle {

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
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ObjectType,

        [Parameter(Mandatory, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $WhereClause,

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

    ([xml]$Session.doSelect($SessionID, $ObjectType, $WhereClause, -1, @('id'))).UDSObjectList.UDSObject.Handle
}
