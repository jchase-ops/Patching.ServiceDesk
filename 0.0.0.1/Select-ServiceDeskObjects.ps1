# .ExternalHelp $PSScriptRoot\Select-ServiceDeskObjects-help.xml
function Select-ServiceDeskObjects {

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

        [Parameter(Position = 4)]
        [ValidateNotNullOrEmpty()]
        [AllowEmptyCollection()]
        [System.String[]]
        $Attributes = @(),

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

    $results = [System.Collections.Generic.List[System.Object]]::New()
    $queryLimit = 250

    $queryList = $Session.doQuery($SessionID, $ObjectType, $WhereClause)
    if ($queryList.listLength -le $queryLimit) {
        $xmlList = @(([xml]$Session.getListValues($SessionID, $queryList.listHandle, 0, $queryList.listLength - 1, $Attributes)).UDSObjectList.UDSObject)
    }
    else {
        $xmlList = [System.Collections.Generic.List[System.Object]]::New()
        $startIndex = 0
        $endIndex = $($queryLimit - 1)
        do {
                ([xml]$Session.getListValues($SessionID, $queryList.listHandle, $startIndex, $endIndex, $Attributes)).UDSObjectList.UDSObject | ForEach-Object {
                $xmlList.Add($_)
            }
            if (($queryList.listLength - $endIndex) -ge $queryLimit) {
                $startIndex = $startIndex + $queryLimit
                $endIndex = $endIndex + $queryLimit
            }
            else {
                $startIndex = $startIndex + $queryLimit
                $endIndex = $queryList.listLength - 1
            }
        } until ($startIndex -ge $($queryList.listLength - 1))
    }
    $Session.freeListHandles($SessionID, $queryList.listHandle)
        
    ForEach ($xml in $xmlList) {
        $params = [System.Collections.Specialized.OrderedDictionary]::New()
        ForEach ($xmlAttribute in $($xml.Attributes.Attribute | Sort-Object -Property AttrName)) {
            $params.Add($xmlAttribute.AttrName, $xmlAttribute.AttrValue)
        }
        $results.Add($(New-Object -TypeName PSObject -Property $params))
    }

    if ($($MyInvocation.MyCommand.Name) -eq $((Split-Path -Path $PSCommandPath -Leaf).TrimEnd('.ps1'))) {
        if (!($windowVisible) -or $Quiet) {
            Disconnect-ServiceDesk -Quiet
        }
        else {
            Disconnect-ServiceDesk
        }
    }
    
    $results | Sort-Object -Property chg_ref_num
}
