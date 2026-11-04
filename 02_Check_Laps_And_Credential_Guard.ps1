<#
.SYNOPSIS
  Check Windows LAPS policy signals and Credential Guard / VBS status.
#>
. "$PSScriptRoot\00_Config.ps1"

$laps = [ordered]@{
    LegacyAdmPwdPolicyPresent = $false
    WindowsLapsPolicyPresent = $false
    Details = @{}
}
$legacy = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft Services\AdmPwd" -ErrorAction SilentlyContinue
if ($legacy) {
    $laps.LegacyAdmPwdPolicyPresent = $true
    $laps.Details.Legacy = @{
        AdmPwdEnabled = $legacy.AdmPwdEnabled
        PasswordAgeDays = $legacy.PasswordAgeDays
        PasswordLength = $legacy.PasswordLength
    }
}
$wlaps = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS" -ErrorAction SilentlyContinue
if ($wlaps) {
    $laps.WindowsLapsPolicyPresent = $true
    $laps.Details.WindowsLaps = ($wlaps.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object { @{ $_.Name = $_.Value } })
}

$cg = $null
try {
    $obj = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction Stop
    # 2 = Credential Guard in SecurityServices* arrays (Microsoft docs)
    $cg = [ordered]@{
        VirtualizationBasedSecurityStatus = $obj.VirtualizationBasedSecurityStatus
        SecurityServicesConfigured = @($obj.SecurityServicesConfigured)
        SecurityServicesRunning = @($obj.SecurityServicesRunning)
        CredentialGuardConfigured = (@($obj.SecurityServicesConfigured) -contains 2)
        CredentialGuardRunning = (@($obj.SecurityServicesRunning) -contains 2)
    }
} catch {
    $cg = [ordered]@{ Error = $_.Exception.Message }
}

$out = [ordered]@{
    Organisation = $Org
    Host = $env:COMPUTERNAME
    CollectedUtc = (Get-Date).ToUniversalTime().ToString("o")
    Mode = "LiveDefensiveInventory"
    LAPS = $laps
    CredentialGuard = $cg
    Guidance = @(
        "Deploy Windows LAPS so each device has a unique, rotating local admin password.",
        "Remove daily users from local Administrators; use just-in-time elevation for IT.",
        "Enable Credential Guard where hardware/OS supports it; pair with ASR rules via Defender."
    )
}
Write-EvidenceJson -Prefix "LAPS_CredentialGuard" -Object $out | Out-Null
Write-Host ("LAPS legacy={0} windows={1}; CredGuard running={2}" -f $laps.LegacyAdmPwdPolicyPresent, $laps.WindowsLapsPolicyPresent, $cg.CredentialGuardRunning)
