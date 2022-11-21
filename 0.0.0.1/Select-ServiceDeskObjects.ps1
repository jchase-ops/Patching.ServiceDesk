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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/JoIm9OSeMwg5Gh/J4DQT7XA
# R5igggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUQnaAFck+E9D0rit4grCBxO/I
# TnMwDQYJKoZIhvcNAQEBBQAEggEA56pIyA5Fnj/EhvmUkj4OybPjw5dhG6KK/0AH
# x6C/f6xVtgdVgyZfK05m9uYaTuC2Bh1TLvILqEb1n4i+oZG8Q9U/XZ74pPCVU6G0
# OaDRny1p9unU56DNh1U+s3qhbTQOqkdr8lHhh5CkU/lZXti70Dy+EijKDwosUiPR
# 11LwCMg1TYmU1mJZyVtZ2x6G5YjGPadMGdCgNxVS4afnUcmNjWUH5bVscTsqcq5T
# qgkTQD06Lp7notYUR4tP4PlJ5UNy0sKwNLbBQz7ijCB9HfwN5FE5tytQVIvq/ZQ8
# fqXEbS5138JbhFez4JT7hsrK1SPeXxwPZynTzbk/fbnUBpPBxw==
# SIG # End signature block
