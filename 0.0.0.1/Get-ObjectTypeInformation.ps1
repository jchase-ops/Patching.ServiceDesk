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
