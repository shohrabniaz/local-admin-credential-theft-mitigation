<#
.SYNOPSIS
  Score privilege/credential exposure and post-mitigation residual using inventory + design coverage.
#>
. "$PSScriptRoot\00_Config.ps1"

& "$PSScriptRoot\01_Inventory_Local_Admins.ps1"
& "$PSScriptRoot\02_Check_Laps_And_Credential_Guard.ps1"

function Latest([string]$suffix) {
    Get-ChildItem $EvidenceDir -Filter "*_$suffix.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

$invFile = Latest "Local_Admin_Inventory"
$cgFile = Latest "LAPS_CredentialGuard"
$inv = if ($invFile) { Get-Content $invFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }
$cg = if ($cgFile) { Get-Content $cgFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json } else { $null }

$adminCount = if ($inv) { @($inv.LocalAdministrators).Count } else { 0 }
$dailyAdminLikely = $adminCount -gt 1  # built-in Administrator + at least one other
$lapsPresent = $false
if ($cg -and $cg.LAPS) {
    $lapsPresent = [bool]$cg.LAPS.LegacyAdmPwdPolicyPresent -or [bool]$cg.LAPS.WindowsLapsPolicyPresent
}
$cgRunning = $false
if ($cg -and $cg.CredentialGuard) {
    $cgRunning = [bool]$cg.CredentialGuard.CredentialGuardRunning
}

# Exposure factors (pre-mitigation intent for SME fleet modelled from this host)
$factors = @(
    [ordered]@{ Id = "EXP-01"; Title = "Standing local admin for daily users"; Pre = $(if ($dailyAdminLikely) { 90 } else { 40 }); PostDesign = 15; Max = 100; Note = "Remove interactive users from Administrators" }
    [ordered]@{ Id = "EXP-02"; Title = "Unmanaged / shared local admin password"; Pre = $(if ($lapsPresent) { 25 } else { 85 }); PostDesign = 20; Max = 100; Note = "Windows LAPS unique rotating passwords" }
    [ordered]@{ Id = "EXP-03"; Title = "Credential theft from LSASS / memory"; Pre = $(if ($cgRunning) { 35 } else { 80 }); PostDesign = $(if ($cgRunning) { 30 } else { 40 }); Max = 100; Note = "Credential Guard + ASR + least privilege" }
    [ordered]@{ Id = "EXP-04"; Title = "Lateral movement via reused local admin"; Pre = $(if ($lapsPresent) { 30 } else { 90 }); PostDesign = 25; Max = 100; Note = "LAPS + deny network logon for local admin where feasible" }
    [ordered]@{ Id = "EXP-05"; Title = "IT support friction / shadow elevation"; Pre = 50; PostDesign = 35; Max = 100; Note = "JIT elevation helpdesk process; measure ticket volume" }
)

function Agg($field) {
    $s = ($factors | ForEach-Object { $_.$field } | Measure-Object -Sum).Sum
    $m = ($factors | ForEach-Object { $_.Max } | Measure-Object -Sum).Sum
    [math]::Round(100.0 * $s / $m, 1)
}

$pre = Agg "Pre"
$post = Agg "PostDesign"

$out = [ordered]@{
    Organisation = $Org
    Host = $env:COMPUTERNAME
    CollectedUtc = (Get-Date).ToUniversalTime().ToString("o")
    Mode = "LiveInventoryPlusDesignCoverage"
    LabSignals = [ordered]@{
        AdminMemberCount = $adminCount
        DailyAdminLikely = $dailyAdminLikely
        LapsPolicyPresent = $lapsPresent
        CredentialGuardRunning = $cgRunning
    }
    PreExposurePercent = $pre
    PostMitigationDesignPercent = $post
    ReductionPoints = [math]::Round($pre - $post, 1)
    Factors = $factors
    ValidationNote = "Lab validation = configuration inventory + Essential Eight–aligned mitigation coverage. No simulated attack tooling."
}
Write-EvidenceJson -Prefix "Privilege_Exposure_Score" -Object $out | Out-Null
Write-Host ("Exposure {0}% -> design residual {1}% (reduction {2})" -f $pre, $post, $out.ReductionPoints)
