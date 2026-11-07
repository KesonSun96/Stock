# Shared helpers for the AI high-odds stock pool project.
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:ProjectRoot = Split-Path -Parent $PSScriptRoot
$script:Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Get-ProjectRoot { return $script:ProjectRoot }

function Join-ProjectPath {
    param([Parameter(Mandatory = $true)][string]$Child)
    return (Join-Path $script:ProjectRoot $Child)
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)
    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

function Read-CsvData {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    return @(Import-Csv -Path $Path -Encoding UTF8)
}

function Write-TextUtf8 {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [switch]$Bom
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $enc = if ($Bom) { $script:Utf8Bom } else { $script:Utf8NoBom }
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
}

function Write-CsvUtf8 {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Data
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $rows = @($Data)
    if ($rows.Count -eq 0) { Write-TextUtf8 -Path $Path -Text '' -Bom; return }
    $csv = $rows | ConvertTo-Csv -NoTypeInformation
    $text = ($csv -join "`r`n") + "`r`n"
    Write-TextUtf8 -Path $Path -Text $text -Bom
}

function Invoke-ApiJson {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [int]$Retry = 3,
        [int]$TimeoutSec = 40
    )
    $headers = @{
        'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
        'Referer'    = 'https://quote.eastmoney.com/'
        'Accept'     = 'application/json,text/plain,*/*'
    }
    for ($i = 1; $i -le $Retry; $i++) {
        try {
            return Invoke-RestMethod -Uri $Uri -Headers $headers -TimeoutSec $TimeoutSec
        }
        catch {
            if ($i -eq $Retry) { throw }
            Start-Sleep -Seconds (2 * $i)
        }
    }
}

function Invoke-ApiJsonUtf8 {
    # Some Eastmoney endpoints omit charset; decode raw bytes as UTF-8 to avoid mojibake.
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [int]$Retry = 3,
        [int]$TimeoutSec = 40
    )
    $headers = @{
        'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
        'Referer'    = 'https://data.eastmoney.com/'
        'Accept'     = 'application/json,text/plain,*/*'
    }
    for ($i = 1; $i -le $Retry; $i++) {
        try {
            $resp = Invoke-WebRequest -Uri $Uri -Headers $headers -TimeoutSec $TimeoutSec -UseBasicParsing
            $text = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
            return ($text | ConvertFrom-Json)
        }
        catch {
            if ($i -eq $Retry) { throw }
            Start-Sleep -Seconds (2 * $i)
        }
    }
}

function ConvertTo-Number {
    param($Value, [double]$Default = 0)
    if ($null -eq $Value) { return $Default }
    $s = ([string]$Value).Trim()
    if ($s -eq '' -or $s -eq '-' -or $s -eq '--') { return $Default }
    $out = 0.0
    if ([double]::TryParse($s, [ref]$out)) { return $out }
    return $Default
}

function Format-Pct {
    param($Value, [int]$Digits = 1)
    if ($null -eq $Value -or "$Value" -eq '') { return '-' }
    return (('{0:N' + $Digits + '}%') -f [double]$Value)
}

function Format-Number {
    param($Value, [int]$Digits = 2)
    if ($null -eq $Value -or "$Value" -eq '') { return '-' }
    return (('{0:N' + $Digits + '}') -f [double]$Value)
}

function Get-EpochTradeDate {
    # Resolve the latest A-share trading date from the SSE Composite quote timestamp.
    try {
        $j = Invoke-ApiJson -Uri 'https://push2.eastmoney.com/api/qt/stock/get?secid=1.000001&fields=f43,f86,f58' -Retry 2
        $epoch = ConvertTo-Number $j.data.f86
        if ($epoch -gt 1000000000) {
            return ([DateTimeOffset]::FromUnixTimeSeconds([int64]$epoch)).ToOffset([TimeSpan]::FromHours(8)).Date
        }
    }
    catch { }
    $fallback = (Get-Date).Date
    while ($fallback.DayOfWeek -eq [DayOfWeek]::Saturday -or $fallback.DayOfWeek -eq [DayOfWeek]::Sunday) {
        $fallback = $fallback.AddDays(-1)
    }
    return $fallback
}

function Get-AStockUniverse {
    # Full Shanghai/Shenzhen/ChiNext/STAR list with quote fields.
    $fields = 'f12,f13,f14,f2,f20,f21,f9,f23,f3,f6,f8,f100,f18,f15,f16,f5,f6'
    $fs = 'm:0+t:6,m:0+t:80,m:1+t:2,m:1+t:23'
    $all = New-Object System.Collections.ArrayList
    $page = 1
    $total = 0
    do {
        $uri = 'https://push2.eastmoney.com/api/qt/clist/get?pn={0}&pz=100&po=1&np=1&fltt=2&invt=2&fid=f3&fs={1}&fields={2}' -f $page, $fs, $fields
        $resp = Invoke-ApiJson -Uri $uri
        $diff = @($resp.data.diff)
        if ($diff.Count -eq 0) { break }
        [void]$all.AddRange($diff)
        $total = [int](ConvertTo-Number $resp.data.total)
        $page++
    } while ($all.Count -lt $total -and $page -le 80)
    return $all
}

function Get-PoolQuotes {
    param([string]$TradeDate = '')
    $pool = Read-CsvData -Path (Join-ProjectPath 'config/stock_pool.csv')
    $byCode = @{}
    foreach ($d in (Get-AStockUniverse)) { $byCode[[string]$d.f12] = $d }

    $rows = New-Object System.Collections.ArrayList
    $missing = New-Object System.Collections.ArrayList
    foreach ($p in $pool) {
        $d = $byCode[$p.code]
        if ($null -eq $d) {
            [void]$missing.Add($p.name)
            [void]$rows.Add([pscustomobject]@{
                code = $p.code; name = $p.name; group = $p.group; exchange = $p.exchange
                price = ''; market_cap_yi = ''; float_cap_yi = ''; pe_ttm = ''; pb = ''
                change_pct = ''; amount_yi = ''; prev_close = ''; high = ''; low = ''
                industry = ''; quote_date = $TradeDate; data_status = '未获取到行情'
            })
            continue
        }
        $price = ConvertTo-Number $d.f2 -1
        $cap = ConvertTo-Number $d.f20
        [void]$rows.Add([pscustomobject]@{
            code        = $p.code
            name        = $p.name
            group       = $p.group
            exchange    = $p.exchange
            price       = if ($price -gt 0) { [math]::Round($price, 2) } else { '' }
            market_cap_yi = if ($cap -gt 0) { [math]::Round($cap / 100000000, 2) } else { '' }
            float_cap_yi  = if ((ConvertTo-Number $d.f21) -gt 0) { [math]::Round((ConvertTo-Number $d.f21) / 100000000, 2) } else { '' }
            pe_ttm      = if ((ConvertTo-Number $d.f9) -ne 0) { [math]::Round((ConvertTo-Number $d.f9), 1) } else { '' }
            pb          = if ((ConvertTo-Number $d.f23) -gt 0) { [math]::Round((ConvertTo-Number $d.f23), 2) } else { '' }
            change_pct  = if ($null -ne $d.f3 -and "$($d.f3)" -ne '-') { [math]::Round((ConvertTo-Number $d.f3), 2) } else { '' }
            amount_yi   = if ((ConvertTo-Number $d.f6) -gt 0) { [math]::Round((ConvertTo-Number $d.f6) / 100000000, 2) } else { '' }
            prev_close  = if ((ConvertTo-Number $d.f18) -gt 0) { [math]::Round((ConvertTo-Number $d.f18), 2) } else { '' }
            high        = if ((ConvertTo-Number $d.f15) -gt 0) { [math]::Round((ConvertTo-Number $d.f15), 2) } else { '' }
            low         = if ((ConvertTo-Number $d.f16) -gt 0) { [math]::Round((ConvertTo-Number $d.f16), 2) } else { '' }
            industry    = [string]$d.f100
            quote_date  = $TradeDate
            data_status = '已核实行情'
        })
    }
    return [pscustomobject]@{ rows = @($rows); missing = @($missing); total = $pool.Count }
}

function Get-PoolFinancials {
    param([switch]$Force)
    $dir = Join-ProjectPath 'data/financials'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $latest = Get-ChildItem -Path $dir -Filter '*.csv' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $Force -and $latest) {
        return [pscustomobject]@{ rows = @(Read-CsvData -Path $latest.FullName); file = $latest.FullName; cached = $true }
    }

    $pool = Read-CsvData -Path (Join-ProjectPath 'config/stock_pool.csv')
    $rows = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($p in $pool) {
        $i++
        Write-Progress -Activity 'Fetching financial reports' -Status $p.name -PercentComplete ([int]($i * 100 / $pool.Count))
        $uri = 'https://datacenter-web.eastmoney.com/api/data/v1/get?reportName=RPT_LICO_FN_CPD&columns=ALL&filter=(SECURITY_CODE%3D%22{0}%22)&pageNumber=1&pageSize=2&sortColumns=REPORTDATE&sortTypes=-1' -f $p.code
        $row = [pscustomobject]@{
            code = $p.code; name = $p.name; group = $p.group
            report_date = ''; report_name = ''; revenue_yi = ''; net_profit_yi = ''
            revenue_yoy = ''; profit_yoy = ''; gross_margin = ''; roe = ''; eps = ''; notice_date = ''
            data_status = '未获取到财报'
        }
        try {
            $j = Invoke-ApiJson -Uri $uri -Retry 2
            $d = @($j.result.data)
            if ($d.Count -gt 0) {
                $x = $d[0]
                $row.report_date = ([string]$x.REPORTDATE).Substring(0, 10)
                $row.report_name = [string]$x.DATATYPE
                $row.revenue_yi = [math]::Round((ConvertTo-Number $x.TOTAL_OPERATE_INCOME) / 100000000, 2)
                $row.net_profit_yi = [math]::Round((ConvertTo-Number $x.PARENT_NETPROFIT) / 100000000, 2)
                $row.revenue_yoy = [math]::Round((ConvertTo-Number $x.YSTZ), 2)
                $row.profit_yoy = [math]::Round((ConvertTo-Number $x.SJLTZ), 2)
                $row.gross_margin = [math]::Round((ConvertTo-Number $x.XSMLL), 2)
                $row.roe = [math]::Round((ConvertTo-Number $x.WEIGHTAVG_ROE), 2)
                $row.eps = [math]::Round((ConvertTo-Number $x.BASIC_EPS), 3)
                $row.notice_date = if ($x.NOTICE_DATE) { ([string]$x.NOTICE_DATE).Substring(0, 10) } else { '' }
                $row.data_status = '已核实财报'
            }
        }
        catch { }
        [void]$rows.Add($row)
        Start-Sleep -Milliseconds 120
    }
    Write-Progress -Activity 'Fetching financial reports' -Completed

    $stamp = (Get-Date).ToString('yyyy-MM-dd')
    $file = Join-Path $dir ("financials-$stamp.csv")
    Write-CsvUtf8 -Path $file -Data $rows
    return [pscustomobject]@{ rows = @($rows); file = $file; cached = $false }
}
