# .ExternalHelp $PSScriptRoot\Connect-ServiceDesk-help.xml
function Connect-ServiceDesk {

    [CmdletBinding()]

    Param (

        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        $Credential,

        [Parameter()]
        [Switch]
        $SaveCredential,

        [Parameter()]
        [Switch]
        $PassThru,

        [Parameter()]
        [Switch]
        $Quiet
    )

    $windowVisible = if ($(Get-Process -Id $([System.Diagnostics.Process]::GetCurrentProcess().Id)).MainWindowHandle -eq 0) { $false } else { $true }

    if ($null -eq $script:Config.Uri) {
        if ($windowVisible -and !($Quiet)) {
            $script:Config.Uri = Read-Host -Prompt 'Enter ServiceDesk Uri'
            $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
        }
        else {
            exit
        }
    }

    if (!($Credential)) {
        if ($null -eq $script:Config.Credential) {
            if ($windowVisible -and !($Quiet)) {
                $script:Config.Credential = $Host.UI.PromptForCredential("Service Desk Credentials", "Enter password for ${env:USERDOMAIN}\${env:USERNAME}", "${env:USERDOMAIN}\${env:USERNAME}", '')
                if ($SaveCredential) {
                    $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
                }
            }
        }
    }
    else {
        $script:Config.Credential = $Credential
        if ($SaveCredential) {
            $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
        }
    }

    if ($null -ne $script:Config.Session) {
        if ($script:Config.SessionID -eq 0 -or $null -eq $script:Config.SessionID) {
            $script:Config.Session = $null
            $script:Config.SessionID = $null
        }
    }
    else {
        if ($null -ne $script:Config.SessionID){
            $script:Config.SessionID = $null
        }
    }
    
    $script:Config.Session = New-WebServiceProxy -Uri $script:Config.Uri
    $script:Config.Session.Credentials = $script:Config.Credential
    $script:Config.Session.AllowAutoRedirect = $true

    try {
        $script:Config.SessionID = $script:Config.Session.login($script:Config.Credential.UserName, $script:Config.Credential.GetNetworkCredential().Password)
    }
    catch {
        $script:Config.SessionID = 0
        
    }
    finally {
        if ($script:Config.SessionID -ne 0) {
            if ($windowVisible -and !($Quiet)) {
                Write-Host "Established Service Desk session." -ForegroundColor Green
            }
        }
        else {
            if ($windowVisible -and !($Quiet)) {
                Write-Host "Failed to establish Service Desk session." -ForegroundColor Red
            }
        }

        if ($PassThru) {
            $script:Config.Session, $script:Config.SessionID
        }
    }
}
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUDJWLzZai3rIQjWifvQvKc+Uq
# bRegggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU2jWi/rmBCqsgVZKsCLIOWPd4
# Aw0wDQYJKoZIhvcNAQEBBQAEggEAkEYpSHHNO86KyXVfg0ubmZdqAyf7mjHKnAJx
# maqzIjlwmaR49LAa2btd5y9GLgTj8wIAOA9iFgy5e/coZJZhHuR/mzsPuJRI3xKp
# UawJJZ04aD+iY29M4RfV5FXyGU8+jEVdOrufOpb8bara8JSbhhZ4I69x6Q+oNHKR
# OXSXteVaNqerAatDpIeKqAOapcdY91AdI21I+pjopQHZKTmU62hkGjaAFajNBxY6
# VwL6Xm/L69Yi3qQF4kaO+5QYF8Rq8hqtlMiBrRJQMfzVme35WfGF30Qv0i+Dz/o3
# c1FmxgtnmOefjZi+q9cnF74xVVY2W8+8ki61BGKTJqmMkhxUIA==
# SIG # End signature block
