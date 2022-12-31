# .ExternalHelp $PSScriptRoot\Get-ChangeOrder-help.xml
function Get-ChangeOrder {

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
        [ValidateScript({ $_ -in $script:Config.ChangeOrderAttributes })]
        [AllowEmptyCollection()]
        [System.String[]]
        $ChangeOrderAttributes = $script:Config.ChangeOrderAttributes,

        [Parameter(Position = 4)]
        [ValidateScript({ $_ -in $script:Config.WorkflowTaskAttributes })]
        [AllowEmptyCollection()]
        [System.String[]]
        $WorkflowTaskAttributes = $script:Config.WorkflowTaskAttributes,

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

    $unixEpoch = [System.DateTime]::New(1970, 1, 1, 0, 0, 0, 0)
    $coHandle = Get-ServiceDeskHandle -ObjectType chg -WhereClause "chg_ref_num = '$ChangeOrder'"

    ([xml]$Session.getObjectValues($SessionID, $coHandle, $ChangeOrderAttributes)).UDSObject.Attributes | ForEach-Object {
        $co = [PSCustomObject]@{
            ChangeOrder       = $ChangeOrder
            Handle            = $coHandle
            Requester         = $_.Attribute[0].AttrValue
            AffectedEndUser   = $_.Attribute[1].AttrValue
            Category          = $_.Attribute[2].AttrValue
            Status            = $_.Attribute[3].AttrValue
            Priority          = $_.Attribute[4].AttrValue
            Type              = $_.Attribute[5].AttrValue
            CreatedBy         = $_.Attribute[6].AttrValue
            Assignee          = $_.Attribute[7].AttrValue
            Group             = $_.Attribute[8].AttrValue
            CAB               = $_.Attribute[9].AttrValue
            Active            = $_.Attribute[10].AttrValue
            ScheduleStartDate = $unixEpoch.AddSeconds($_.Attribute[11].AttrValue).ToLocalTime()
            Organization      = $_.Attribute[12].AttrValue
            Summary           = $_.Attribute[13].AttrValue
            ScheduleDuration  = New-TimeSpan -Seconds $_.Attribute[14].AttrValue
            ScheduleEndDate   = $unixEpoch.AddSeconds($_.Attribute[11].AttrValue).ToLocalTime().AddSeconds($_.Attribute[14].AttrValue)
            Approval          = $_.Attribute[15].AttrValue
            OpenDate          = $unixEpoch.AddSeconds($_.Attribute[16].AttrValue).ToLocalTime()
            CreatedVia        = $_.Attribute[17].AttrValue
            WebUrl            = $_.Attribute[18].AttrValue
            Description       = $_.Attribute[19].AttrValue
            Properties        = [System.Collections.Specialized.OrderedDictionary]::New()
            WorkflowTasks     = [System.Collections.Generic.List[System.Object]]::New()
        }
    }

    ([xml]$Session.doSelect($SessionID, 'prp', "object_id = $($coHandle.TrimStart('chg:').ToInt32($null))", -1, @('label', 'sequence', 'value'))).UDSObjectList.UDSObject | ForEach-Object {
        $co.Properties.Add($_.Attributes.Attribute[1].AttrValue, $_.Attributes.Attribute[2].AttrValue)
    }

    ([xml]$Session.doSelect($SessionID, 'wf', "chg = $($co.Handle.TrimStart('chg:').ToInt32($null))", -1, $WorkflowTaskAttributes)).UDSObjectList.UDSObject | ForEach-Object {
        if ($null -ne $_) {
            $wf = [PSCustomObject]@{
                SequenceID        = $_.Attributes.Attribute[0].AttrValue
                Task              = $_.Attributes.Attribute[1].AttrValue
                Assignee          = $_.Attributes.Attribute[2].AttrValue
                Group             = $_.Attributes.Attribute[3].AttrValue
                Status            = $_.Attributes.Attribute[4].AttrValue
                Comments          = $_.Attributes.Attribute[5].AttrValue
                ScheduleStartDate = $unixEpoch.AddSeconds($_.Attributes.Attribute[6].AttrValue).ToLocalTime()
                StartDate         = $unixEpoch.AddSeconds($_.Attributes.Attribute[7].AttrValue).ToLocalTime()
                CompletionDate    = $unixEpoch.AddSeconds($_.Attributes.Attribute[8].AttrValue).ToLocalTime()
                Description       = $_.Attributes.Attribute[9].AttrValue
            }
            $co.WorkflowTasks.Add($wf)
        }
    }

    if ($($MyInvocation.MyCommand.Name) -eq $((Split-Path -Path $PSCommandPath -Leaf).TrimEnd('.ps1'))) {
        if (!($windowVisible) -or $Quiet) {
            Disconnect-ServiceDesk -Quiet
        }
        else {
            Disconnect-ServiceDesk
        }
    }

    $co
}
