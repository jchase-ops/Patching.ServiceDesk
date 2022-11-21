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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZcJ+UKS/HEYI1LWBeuOdjvam
# vWGgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
# AQUFADAWMRQwEgYDVQQDDAtDZXJ0LTAzNDU2MDAeFw0yMTEyMDIwNDU5MTJaFw0y
# MjEyMDIwNTE5MTJaMBYxFDASBgNVBAMMC0NlcnQtMDM0NTYwMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEA8daSAcUBI0Xx8sMMlSpsCV+24lY46RsxX8iC
# bB7ZM19b/GBjwMo0TCb28ssbZ/P8liNJICrSbyIkQDrIrjqtAdyAPdPAYHONTHad
# 0fuOQQT5MkO5HAxUYLz/6H/xq92lKQFxz5Wgzw+3KOyignY8V8ZZ379z/WqQbNCV
# +29zb9YWOK7eXQ9x8s4+SOizqUE3zkOuijf86I9vZmzMYhsxE7if0R0UlQsLlvTA
# kH/m4IjHem8rl/kC+O71lU7l9475XrUUR3Fxebqh9YoCEZh2eE81TLQcnvK8zgqP
# F+X4INdNPD6zO4T1Nbz0Ccev7mj37+pk/eL5R5aV+NJgqAzhvQIDAQABo0YwRDAO
# BgNVHQ8BAf8EBAMCBaAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFFNN
# e4x6JSqbcnTR354fVSEgQ0VYMA0GCSqGSIb3DQEBBQUAA4IBAQBXfA8VgaMD2c/v
# Sv8gnS/LWri51BBqcUFE9JYMxEIzlEt2ZfJsG+INaQqzBoyCDx/oMQH7wdFRvDjQ
# QsXpNTo7wH7WytFe9KJrOz2uGG0EnIYHK0dTFIMVOcM9VsWWPG40EAzD//55xX/d
# pBL+L4SSTujbR3ptni8Agu5GiRhTpxwl1L/HLC2QYYMoUKiAxL1p61+cHRj6wMzl
# jxnrMIcBhKioaXnwWdKPCN66Jk8IYdqr8afcRYiwtDi+8Hk2/9nB9HwPox3Dtf8H
# jH0O2/8NiJTeOBFSfrWPM9r4j4NWR8IuLwsqHUfXJEQa9SOxhHvxaNMR/Fhq1GVj
# qUClZiXiMYIByzCCAccCAQEwKjAWMRQwEgYDVQQDDAtDZXJ0LTAzNDU2MAIQFnL4
# oVNG56NIRjNfzwNXejAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUgNPUvq27LQRWbSV7AbUDJ03r
# bOgwDQYJKoZIhvcNAQEBBQAEggEAttMWwRz+TzH1KozzRRE0BYu/PKuIG8EzRwUM
# UZfjeqoY3KrOqi9WW0oU0toG+oVjk/YXmnWpaS5zrJPz2qNyrRPy6Seur0vlK73+
# f12CXRjebysMItHI3UR+Pi0FLbgvtqmhvtsEp/ricnl5evpDaeVHrvAHOpzvvcz3
# bi/xdJVgsR3w/8hvUPJbW4CdJyIipdz7jv4UJQuySFeaXKSyA8IXmqLuj98AXK4s
# YmmVfzK6G0FKOu6VcKIJwWf7u5C9KupWeGXknOMzi+EtkNKUZ6qsvrlGJ4b/eLD6
# qWsvTvUNE3O06jXqt/fJDPnYpV62ytIefk7X73Hl1aBhH/h28w==
# SIG # End signature block
