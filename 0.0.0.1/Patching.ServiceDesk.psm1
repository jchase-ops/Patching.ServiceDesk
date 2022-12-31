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
