# .ExternalHelp $PSScriptRoot\Move-ChangeOrder-help.xml
function Move-ChangeOrder {

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
        [ValidateNotNull()]
        [AllowEmptyString()]
        [System.String]
        $Description = '',

        [Parameter(Position = 4)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Assignee,

        [Parameter(Position = 5)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Group,

        [Parameter(Position = 6)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Organization,
        
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
    $co_handle = Get-ServiceDeskHandle -ObjectType chg -WhereClause "chg_ref_num = '$ChangeOrder'"

    $set = [PSCustomObject]@{
        Assignee     = $false
        Group        = $false
        Organization = $false
    }

    $setHandles = [PSCustomObject]@{
        Assignee     = $null
        Group        = $null
        Organization = $null
    }

    'Assignee', 'Group', 'Organization' | ForEach-Object {
        if ($PSBoundParameters[$($_)]) {
            $set.$($_) = $true
            Switch ($_) {
                'Assignee' { $setHandles.$($_) = $Session.getHandleForUserId($SessionID, $Assignee) }
                'Group' { $setHandles.$($_) = Get-ServiceDeskHandle -ObjectType grp -WhereClause "last_name = '$Group'" }
                'Organization' { $setHandles.$($_) = Get-ServiceDeskHandle -ObjectType org -WhereClause "name = '$Organization'" }
            }
        }
    }

    $handle = ([xml]$Session.transfer($SessionID, $creator_handle, $co_handle, $Description, $set.Assignee, $setHandles.Assignee, $set.Group, $setHandles.Group, $set.Organization, $setHandles.Organization)).UDSObject.Handle

    if ($MyInvocation.MyCommand.Name -eq $(Split-Path -Path $PSCommandPath -Leaf).TrimEnd('.ps1')) {
        if (!($windowVisible) -or $Quiet) {
            Disconnect-ServiceDesk -Quiet
        }
        else {
            Disconnect-ServiceDesk
        }
    }

    $handle
}
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHHMkh+2/d1HLjhu9VBmLAbeB
# 9FugggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUpmW9/+qX1xKFgR7K7UloILH9
# RPUwDQYJKoZIhvcNAQEBBQAEggEAnDUs+h6wpgAKlpslgR/aYqhEkhQIag2QgivD
# X0Z4JhuzCV4GMddDG6j0QYXQ4aRW9NfES7Iyvjr6b9Su+3WNHGE7ALOk2Grbz+Xw
# QBM6J0Vr4GKzYKlV1obJlfaw8I1MgVYJvfMne+ZCsn0tG8/7JCUIqfEHPBbg2zBg
# FzRIVtWJv+jkmV10RWm8etHPDQbvzdTR7KR0t+yRi6vcZKpwICSj29nqJf2tTEqF
# IqtDfWeGrMUMTVe+3TDo+ynOZpKB4Lrcse7z83WLEv4XTuFYlJ6EOImmlQDIdv/H
# 6CYKMaKu4Fn0NyKjFmfvn8wPy6WWbZct2nhzCAFUGFab9il+cw==
# SIG # End signature block
