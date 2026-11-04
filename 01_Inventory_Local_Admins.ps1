<#
.SYNOPSIS
  Defensive inventory of local Administrators group and current user elevation.
#>
. "$PSScriptRoot\00_Config.ps1"

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

$members = @()
try {
    $members = @(Get-LocalGroupMember -Group "Administrators" | ForEach-Object {
        [ordered]@{ Name = $_.Name; ObjectClass = $_.ObjectClass; PrincipalSource = [string]$_.PrincipalSource }
    })
} catch {
    $members = @([ordered]@{ Name = "(enumeration failed)"; Error = $_.Exception.Message })
}

$out = [ordered]@{
    Organisation = $Org
    Host = $env:COMPUTERNAME
    User = "$env:USERDOMAIN\$env:USERNAME"
    CollectedUtc = (Get-Date).ToUniversalTime().ToString("o")
    Mode = "LiveDefensiveInventory"
    Ethics = "Read-only local group enumeration. No credential access."
    CurrentUserIsElevatedAdmin = $isAdmin
    LocalAdministrators = $members
    FindingHints = @(
        "Any interactive daily account in Administrators is standing local privilege.",
        "Shared or well-known local Administrator passwords enable lateral movement after one host is compromised."
    )
}
Write-EvidenceJson -Prefix "Local_Admin_Inventory" -Object $out | Out-Null
Write-Host ("Elevated={0}; AdminCount={1}" -f $isAdmin, $members.Count)
