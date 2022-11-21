# Patching.ServiceDesk PS Module

#region Classes
################################################################################
#                                                                              #
#                                 CLASSES                                      #
#                                                                              #
################################################################################
# . "$PSScriptRoot\$(Split-Path -Path $(Split-Path -Path $PSScriptRoot -Parent) -Leaf).Classes.ps1"
#endregion

#region Variables
################################################################################
#                                                                              #
#                               VARIABLES                                      #
#                                                                              #
################################################################################
try {
    $script:Config = Import-Clixml -Path "$PSScriptRoot\config.xml"
}
catch {
    $script:Config = [ordered]@{
        Credential     = $null
        Session        = $null
        SessionID      = $null
        Uri            = $null
        ChangeOrderAttributes = @(
            'requestor.combo_name'
            'affected_contact.combo_name'
            'category.sym'
            'status.sym'
            'priority.sym'
            'chgtype.sym'
            'log_agent.combo_name'
            'assignee.combo_name'
            'group.combo_name'
            'cab.combo_name'
            'active.sym'
            'sched_start_date'
            'organization.name'
            'summary'
            'sched_duration'
            'cab_approval'
            'open_date'
            'created_via.sym'
            'web_url'
            'description'
        )
        WorkflowTaskAttributes = @(
            'sequence'
            'task.sym'
            'assignee.combo_name'
            'group.combo_name'
            'status.sym'
            'comments'
            'date_created'
            'start_date'
            'completion_date'
            'description'
        )
        DataTypeHash = [ordered]@{
            '2001' = 'Integer'
            '2002' = 'String'
            '2003' = 'Duration'
            '2004' = 'Date'
            '2005' = 'SREL'
            '2006' = 'UNKNOWN'
            '2007' = 'List (QREL/BREL)'
            '2008' = 'Lrel (many-to-many)'
            '2009' = 'UUID' 
        }
    }
    $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
}
#endregion

#region DotSourcedScripts
################################################################################
#                                                                              #
#                           DOT-SOURCED SCRIPTS                                #
#                                                                              #
################################################################################
. "$PSScriptRoot\Connect-ServiceDesk.ps1"
. "$PSScriptRoot\Disconnect-ServiceDesk.ps1"
. "$PSScriptRoot\Get-ObjectTypeInformation.ps1"
. "$PSScriptRoot\Get-PropertyInfoForCategory.ps1"
. "$PSScriptRoot\Get-ChangeOrder.ps1"
. "$PSScriptRoot\Get-ServiceDeskHandle.ps1"
. "$PSScriptRoot\Move-ChangeOrder.ps1"
. "$PSScriptRoot\New-ChangeOrder.ps1"
. "$PSScriptRoot\Select-ServiceDeskObjects.ps1"
. "$PSScriptRoot\Set-ChangeOrderStatus.ps1"
#endregion

#region ModuleMembers
################################################################################
#                                                                              #
#                              MODULE MEMBERS                                  #
#                                                                              #
################################################################################
Export-ModuleMember -Function Connect-ServiceDesk
Export-ModuleMember -Function Disconnect-ServiceDesk
Export-ModuleMember -Function Get-ObjectTypeInformation
Export-ModuleMember -Function Get-PropertyInfoForCategory
Export-ModuleMember -Function Get-ChangeOrder
Export-ModuleMember -Function Get-ServiceDeskHandle
Export-ModuleMember -Function Move-ChangeOrder
Export-ModuleMember -Function New-ChangeOrder
Export-ModuleMember -Function Select-ServiceDeskObjects
Export-ModuleMember -Function Set-ChangeOrderStatus
#endregion
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUscE5lK4eRrLwFUeEOBwh57Jm
# T/SgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU9kVcNYFLdXWxmtPyl5NFDRBD
# UdcwDQYJKoZIhvcNAQEBBQAEggEA1HM+DuYGdDTQ2UyWaaCXMCi/Bt8gFH3kzDFF
# sUI5diRT1fs8ISGZA0FBY7kzynmr8LtpaJbvuii2Vn9XSZH+e+6umYelGPUTS1cB
# lDCjW5Z8biCMI4rDscPhNWJd9l99A18PpgdVCgick4iDdMGxlBnbXV0t6tp92xz+
# TKCsuRqoYkwA5xGi2T6no/LoD5QT1X21hlnBKXZpsYkDTYA+PFh0OUgA6et4zwO3
# qPtgy33e9uowyq1lH72T8B2seRKezQJOSEttrUK7yiStS3RI1rWi4TNC0D00zcZH
# /jBXDlMtNQ4aZqG1bkpaZYeKjCJkCe+tVz9RVeHK5KH4AHOQ/w==
# SIG # End signature block
