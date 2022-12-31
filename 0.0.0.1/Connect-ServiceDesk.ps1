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
