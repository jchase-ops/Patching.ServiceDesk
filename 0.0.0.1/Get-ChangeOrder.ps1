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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqD0XKqMPdojlcEkZWIehhl02
# 2l2gggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUJiLAgpsl0uOdfPqC5Gl4ZRGi
# uTwwDQYJKoZIhvcNAQEBBQAEggEAH8SsXvdgWVd+weQB3pm1dhsuYEcosNRbzDGH
# 9jnftkpzBSWMe7DTC+8jDeock8rRY1iB9frDmPoUOnLN1c+Tm1QVBhEsTsM7DJND
# HX7uasy+KE2avlxLLdZ/oGbnzuX+ogHUSEONfJ/rP4YRnTrY9knMky81J356lX5e
# c8cPis9FyNfqWZld8Mbhzobc2N6piIzPxVjxLI2uqV2/NXIgtavCmjBZAneDZ/vN
# Fg8QC4LMjZpAaUKEKeP4/kjPpRZshlX5ZECg3jnaAAnbcQwM0qGYiyNPQIJ/stdu
# 6ekoKx2smVkZ5Cks71FkEKVjN3HPuxXQ4ScMX+ArxLmCSQbAMA==
# SIG # End signature block
