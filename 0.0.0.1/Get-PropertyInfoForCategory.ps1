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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUEx1dZQ7Y9MduN8T8G29vF5CD
# bB2gggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUs4p4Ot1uZUmG+88UFGOLcW60
# Z6cwDQYJKoZIhvcNAQEBBQAEggEAyOXgHcRBWSTpJstnSGcaNiC1BTzyzNmkgBV5
# GdLC0f7LGSQUrMpg9dqm+uVmWvCGqY6EIX3gHU3wPHtpgduLQzxWIPwFIwNVkA9v
# fBanlUCfEc0JLUeDNEEyn6Nsn/BGbQylDTJNLKWdV4/wk4yvYwRTSDyvwWM4gG5e
# 2oRrenUhJ6OpQFk+58STwQfyUe3yELNRzv6q8cMTw5hEhgYGPXgCpPBnZlDkNGnX
# XUdP48XkJyMw1JzEgEc3MeSTYBRnqX26QJGkVa62rabT237UBLvMbIgmZpb5gMq9
# CDJM8hgfdwAHHhap+4aAY/7gte0HLbHs0thRkCU76V4kc1ZvhQ==
# SIG # End signature block
