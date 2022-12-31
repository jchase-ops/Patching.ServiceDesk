# .ExternalHelp $PSScriptRoot\New-ChangeOrder-help.xml
function New-ChangeOrder {

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

        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Requester = "$($Session.Credentials.Domain)\$($Session.Credentials.UserName)",

        [Parameter(Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $AffectedEndUser = "$($Session.Credentials.Domain)\$($Session.Credentials.UserName)",

        [Parameter(Mandatory, Position = 4)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Category,

        [Parameter(Position = 5)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Status = 'RFC',

        [Parameter(Position = 6)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Priority = '3',

        [Parameter(Position = 7)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Type = 'Normal',

        [Parameter(Position = 8)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Assignee,

        [Parameter(Position = 9)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Group = 'NOC',

        [Parameter(Position = 10)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $CAB = 'CAB',

        [Parameter(Mandatory, Position = 11)]
        [ValidateNotNullOrEmpty()]
        [System.DateTime]
        $ScheduleStartDate,

        [Parameter(Position = 12)]
        [ValidateNotNullOrEmpty()]
        [System.TimeSpan]
        $Duration = $(New-TimeSpan -Hours 8),

        [Parameter(Mandatory, Position = 13)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Summary,

        [Parameter(Mandatory, Position = 14)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Description,

        [Parameter(Position = 15)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [System.String[]]
        $Properties = $null,

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

    $attrValueHash = [System.Collections.Specialized.OrderedDictionary]::New()
    $attrValueHash.Add('requestor', $Session.getHandleForUserId($SessionID, $Requester))
    $attrValueHash.Add('affected_contact', $Session.getHandleForUserId($SessionID, $AffectedEndUser))
    $attrValueHash.Add('category', $(Get-ServiceDeskHandle -ObjectType chgcat -WhereClause "sym = '$Category'"))
    $attrValueHash.Add('status', $(Get-ServiceDeskHandle -ObjectType chgstat -WhereClause "sym = '$Status'"))
    $attrValueHash.Add('priority', $(Get-ServiceDeskHandle -ObjectType pri -WhereClause "sym = '$Priority'"))
    $attrValueHash.Add('chgtype', $(Get-ServiceDeskHandle -ObjectType chgtype -WhereClause "sym = '$Type'"))
    if ($Assignee) { $attrValueHash.Add('assignee', $Session.getHandleForUserId($SessionID, $Assignee)) }
    $attrValueHash.Add('group', $(Get-ServiceDeskHandle -ObjectType grp -WhereClause "last_name = '$Group'"))
    $attrValueHash.Add('cab', $(Get-ServiceDeskHandle -ObjectType grp -WhereClause "last_name = '$CAB'"))
    $attrValueHash.Add('active', 1)
    $attrValueHash.Add('sched_start_date', $([DateTimeOffset]::New($ScheduleStartDate.ToUniversalTime()).ToUnixTimeSeconds()))
    $attrValueHash.Add('sched_duration', $Duration.TotalSeconds)
    $attrValueHash.Add('summary', $Summary)
    $attrValueHash.Add('description', $Description)

    $attrValues = [System.Collections.Generic.List[System.String]]::New()
    ForEach ($key in $attrValueHash.Keys) {
        $attrValues.Add($key)
        $attrValues.Add($attrValueHash.$key)
    }

    $new_change_handle = $null
    $new_change_number = $null

    $return = [xml]$Session.createChangeOrder($SessionID, $creator_handle, $attrValues, $Properties, $null, @(), [ref]$new_change_handle, [ref]$new_change_number)

    if ($($MyInvocation.MyCommand.Name) -eq $((Split-Path -Path $PSCommandPath -Leaf).TrimEnd('.ps1'))) {
        if (!($windowVisible) -or $Quiet) {
            Disconnect-ServiceDesk -Quiet
        }
        else {
            Disconnect-ServiceDesk
        }
    }

    $return, $new_change_handle, $new_change_number
}
