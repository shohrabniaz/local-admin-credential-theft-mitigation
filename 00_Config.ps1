$script:ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:EvidenceDir = Join-Path $ProjectRoot "evidence"
$script:Org = "Coastal Managed IT (fictional) — Windows endpoint privilege lab on host"
if (-not (Test-Path $EvidenceDir)) { New-Item -ItemType Directory -Path $EvidenceDir | Out-Null }
function Get-Stamp { Get-Date -Format "yyyyMMdd_HHmmss" }
function Write-EvidenceJson([string]$Prefix, $Object) {
    $path = Join-Path $EvidenceDir ("{0}_{1}.json" -f (Get-Stamp), $Prefix)
    $Object | ConvertTo-Json -Depth 10 | Set-Content -Path $path -Encoding UTF8
    Write-Host "Wrote $path"
    return $path
}
