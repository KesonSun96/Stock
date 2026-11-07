param(
    [string]$Date = '',
    [switch]$RefreshFinancials,
    [switch]$SkipQuotes
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\src\lib.ps1')
. (Join-Path $PSScriptRoot '..\src\score.ps1')
. (Join-Path $PSScriptRoot '..\src\report.ps1')

if ($Date -eq '') {
    $Date = (Get-EpochTradeDate).ToString('yyyy-MM-dd')
}
else {
    $Date = ([datetime]$Date).ToString('yyyy-MM-dd')
}

$quoteFile = Join-ProjectPath "data/quotes/quotes-$Date.csv"
if (-not $SkipQuotes -or -not (Test-Path -LiteralPath $quoteFile)) {
    Write-Host "[1/3] Fetching quotes for $Date ..."
    $quoteResult = Get-PoolQuotes -TradeDate $Date
    Write-CsvUtf8 -Path $quoteFile -Data $quoteResult.rows
    Write-Host ("      saved {0} rows -> {1}" -f $quoteResult.rows.Count, $quoteFile)
    if ($quoteResult.missing.Count -gt 0) {
        Write-Host ("      missing: " + ($quoteResult.missing -join ', '))
    }
}
else {
    Write-Host "[1/3] Reusing quotes: $quoteFile"
}

Write-Host "[2/3] Financial reports ..."
$finResult = Get-PoolFinancials -Force:$RefreshFinancials
Write-Host ("      {0} rows (cached={1}) -> {2}" -f $finResult.rows.Count, $finResult.cached, $finResult.file)

Write-Host "[3/3] Scoring and ranking ..."
$model = Read-JsonFile -Path (Join-ProjectPath 'config/model.json')
$ranked = Invoke-Scoring -TradeDate $Date
$out = Save-DailyOutputs -Ranked $ranked -TradeDate $Date -Model $model

Write-Host ''
Write-Host ("Ranked {0} stocks for {1}" -f $ranked.Count, $Date)
Write-Host ("Report : {0}" -f $out.report_file)
Write-Host ("CSV    : {0}" -f $out.rank_file)
Write-Host ''
$ranked | Select-Object -First 15 rank, rank_change, code, name, group, close_price, market_cap_yi, base_return_pct, final_score, confidence |
    Format-Table -AutoSize | Out-String | Write-Host
