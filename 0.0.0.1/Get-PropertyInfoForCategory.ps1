# .ExternalHelp $PSScriptRoot\Get-PropertyInfoForCategory-help.xml
function Get-PropertyInfoForCategory {

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
        [System.String[]]
        $Category,

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

    $categoryAttributes = 'label', 'sequence', 'required', 'description'
    $results = [System.Collections.Generic.List[System.Object]]::New()

    ForEach ($cat in $Category) {
        $catInfo = [PSCustomObject]@{
            Name       = $cat
            Properties = [System.Collections.Generic.List[System.Object]]::New()
        }
        $catHandle = Get-ServiceDeskHandle -ObjectType chgcat -WhereClause "sym = '$cat'"
            ([xml]$Session.getPropertyInfoForCategory($SessionID, $catHandle, $categoryAttributes)).UDSObjectList.UDSObject | ForEach-Object {
            $xml = $_
            if ($null -ne $xml) {
                $prop = [PSCustomObject]@{
                    Label       = $xml.Attributes.Attribute[0].AttrValue
                    Sequence    = $xml.Attributes.Attribute[1].AttrValue
                    Required    = if ($xml.Attributes.Attribute[2].AttrValue -eq 1) { $true } else { $false }
                    Description = $xml.Attributes.Attribute[3].AttrValue
                }
                $catInfo.Properties.Add($prop)
            }
        }
        $catInfo.Properties = $catInfo.Properties | Sort-Object -Property Label
        $results.Add($catInfo)
    }

    if ($($MyInvocation.MyCommand.Name) -eq $((Split-Path -Path $PSCommandPath -Leaf).TrimEnd('.ps1'))) {
        if (!($windowVisible) -or $Quiet) {
            Disconnect-ServiceDesk -Quiet
        }
        else {
            Disconnect-ServiceDesk
        }
    }
    
    $results | Sort-Object -Property Name
}
