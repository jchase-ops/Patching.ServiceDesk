# .ExternalHelp $PSScriptRoot\Get-ObjectTypeInformation-help.xml
function Get-ObjectTypeInformation {

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
        $ObjectType,

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

    $attributeProperties = 'Name', 'DataType', 'Factory', 'Required', 'Size'
    $results = [System.Collections.Generic.List[System.Object]]::New()

    ForEach ($ot in $ObjectType) {
        $objectInfo = [PSCustomObject]@{
            Name       = $ot
            ObjectType = $ot
            Attributes = [System.Collections.Generic.List[System.Object]]::New()
        }
        $xml = [xml]$Session.getObjectTypeInformation($SessionID, $ot)
        ForEach ($attributeName in $($xml.UDSObject.Attributes | Get-Member -MemberType Property | Select-Object -ExpandProperty Name)) {
            $attr = $xml.UDSObject.Attributes.$attributeName | Select-Object -Property $attributeProperties
            $attr.Name = $attributeName
            $attr.DataType = $script:Config.DataTypeHash.$($attr.DataType)
            $attr.Required = if ($attr.Required -eq 1) { $true } else { $false }
            $objectInfo.Attributes.Add($attr)
        }
        $results.Add($objectInfo)
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
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUnHfxnMUYBTCh+ZgQ194Nhdea
# RxSgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUATz1YZoqs/ImwNtEKmx5sYj+
# 7TcwDQYJKoZIhvcNAQEBBQAEggEAkRlY1e9/gNRR0aMTZanuws4JzQuQytHFgk1X
# 4U6akR2Mh5D0Clp03JOdeC2AJ289xXlQdIH9P67HE+T0/N/N2UPvXdQhrQCLBvTM
# 1SnSAq5qcS5jocz8/XoThUH89Do/kPB9r1+qRxs4AaB065ozlci7W4LJYlDvc2+6
# T1VUv13hxgD6NfJzdkfdaKAj4y9BbR7pR6K1dfzij83sbUHnufrtCtAg2Xw5YudT
# rpo3SyvgsNEdgdIdPb4IP5VP1e6+xC+x5vMAF3eJtbFfZVM0o1MQ7hNKbHhcnV+8
# KrLCPMXCrpn/ZWuEqA+mrrX7DPV5NQll5FoKcQB8ygrpsTlNXA==
# SIG # End signature block
