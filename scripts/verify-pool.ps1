param()

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\src\lib.ps1')

$pool = Read-CsvData -Path (Join-ProjectPath 'config/stock_pool.csv')
$byCode = @{}
foreach ($d in (Get-AStockUniverse)) { $byCode[[string]$d.f12] = $d }

$bad = 0
foreach ($p in $pool) {
    $d = $byCode[[string]$p.code]
    if ($null -eq $d) {
        Write-Host ("MISSING  {0} {1}" -f $p.code, $p.name)
        $bad++
    }
    elseif ([string]$d.f14 -ne $p.name) {
        Write-Host ("MISMATCH {0}: config={1} live={2}" -f $p.code, $p.name, $d.f14)
        $bad++
    }
}
Write-Host ("checked {0} stocks, issues={1}" -f $pool.Count, $bad)
if ($bad -gt 0) { exit 1 }
