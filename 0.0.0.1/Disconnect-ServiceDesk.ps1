# .ExternalHelp $PSScriptRoot\Disconnect-ServiceDesk-help.xml
function Disconnect-ServiceDesk {

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

        [Parameter()]
        [Switch]
        $Quiet
    )

    $windowVisible = if ($(Get-Process -Id $([System.Diagnostics.Process]::GetCurrentProcess().Id)).MainWindowHandle -eq 0) { $false } else { $true }

    if ($Session.serverStatus($SessionID) -eq 0) {
        $Session.logout($SessionID)
        if ($?) {
            $script:Config.Session = $null
            $script:Config.SessionID = $null
            if ($windowVisible -and !($Quiet)) {
                Write-Host "Disconnected from Service Desk" -ForegroundColor Green
            }
        }
        else {
            if ($windowVisible -and !($Quiet)) {
                Write-Host "Failed to disconnect from Service Desk" -ForegroundColor Red
            }
        }
    }
}
