param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\src\lib.ps1')
. (Join-Path $PSScriptRoot '..\src\report.ps1')
. (Join-Path $PSScriptRoot '..\src\profiles.ps1')

$sectors = @(Read-JsonFile -Path (Join-ProjectPath 'config/sectors.json'))
$pool = Read-CsvData -Path (Join-ProjectPath 'config/stock_pool.csv')
$assumptions = Read-CsvData -Path (Join-ProjectPath 'config/assumptions.csv')
$model = Read-JsonFile -Path (Join-ProjectPath 'config/model.json')
$aMap = @{}; foreach ($a in $assumptions) { $aMap[[string]$a.code] = $a }

# Latest snapshots: ranking, quotes, financials, profiles.
$rankFile = Join-ProjectPath 'output/latest/rank.csv'
if (-not (Test-Path -LiteralPath $rankFile)) {
    $rankFile = Get-ChildItem -Path (Join-ProjectPath 'output') -Filter 'rank.csv' -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
}
$rankMap = @{}
if ($rankFile) { foreach ($r in (Read-CsvData -Path $rankFile)) { $rankMap[[string]$r.code] = $r } }

$quoteFile = Get-ChildItem -Path (Join-ProjectPath 'data/quotes') -Filter '*.csv' -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
$quoteMap = @{}
if ($quoteFile) { foreach ($q in (Read-CsvData -Path $quoteFile.FullName)) { $quoteMap[[string]$q.code] = $q } }

$finFile = Get-ChildItem -Path (Join-ProjectPath 'data/financials') -Filter '*.csv' -File -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
$finMap = @{}
if ($finFile) { foreach ($f in (Read-CsvData -Path $finFile.FullName)) { $finMap[[string]$f.code] = $f } }

$profileMap = @{}
foreach ($p in (Read-CsvData -Path (Join-ProjectPath 'data/profiles/company_profiles.csv'))) { $profileMap[[string]$p.code] = $p }
$mixMap = @{}
foreach ($m in (Read-CsvData -Path (Join-ProjectPath 'data/profiles/revenue_mix.csv'))) {
    if (-not $mixMap.ContainsKey([string]$m.code)) { $mixMap[[string]$m.code] = New-Object System.Collections.ArrayList }
    [void]$mixMap[[string]$m.code].Add($m)
}
$annMap = @{}
foreach ($a in (Read-CsvData -Path (Join-ProjectPath 'data/profiles/announcements.csv'))) {
    if (-not $annMap.ContainsKey([string]$a.code)) { $annMap[[string]$a.code] = New-Object System.Collections.ArrayList }
    [void]$annMap[[string]$a.code].Add($a)
}

$outDir = Join-ProjectPath 'docs/sectors'
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

$now = Get-Date -Format 'yyyy-MM-dd'
$created = 0
$skipped = 0
$quoteFileName = if ($quoteFile) { $quoteFile.Name } else { '无' }
$finFileName = if ($finFile) { $finFile.Name } else { '无' }

function Get-PositionText {
    param([double]$Value, [array]$Values)
    if ($Value -le 0 -or $Values.Count -eq 0) { return '-' }
    $better = @($Values | Where-Object { $_ -gt $Value }).Count
    return ("{0}/{1}" -f ($better + 1), $Values.Count)
}

foreach ($sec in ($sectors | Sort-Object -Property order)) {
    $folderName = [System.IO.Path]::GetFileNameWithoutExtension($sec.file)
    $secDir = Join-Path $outDir $folderName
    if (-not (Test-Path -LiteralPath $secDir)) { New-Item -ItemType Directory -Force -Path $secDir | Out-Null }

    $oldFile = Join-Path $outDir $sec.file
    if (Test-Path -LiteralPath $oldFile) { Remove-Item -LiteralPath $oldFile -Force }

    $members = @($pool | Where-Object { $_.group -eq $sec.key })
    $members = @($members | Sort-Object -Property { [int]$aMap[[string]$_.code].old_rank })

    # Peer comparison arrays (pool-internal only).
    $peerCap = @($members | ForEach-Object { $r = $rankMap[[string]$_.code]; if ($r) { ConvertTo-Number $r.market_cap_yi } } | Where-Object { $_ -gt 0 })
    $peerRev = @($members | ForEach-Object { $f = $finMap[[string]$_.code]; if ($f -and "$($f.revenue_yi)" -ne '') { ConvertTo-Number $f.revenue_yi } } | Where-Object { $_ -gt 0 })
    $peerProfit = @($members | ForEach-Object { $f = $finMap[[string]$_.code]; if ($f -and "$($f.net_profit_yi)" -ne '') { ConvertTo-Number $f.net_profit_yi } })
    $peerGm = @($members | ForEach-Object { $f = $finMap[[string]$_.code]; if ($f -and "$($f.gross_margin)" -ne '') { ConvertTo-Number $f.gross_margin } } | Where-Object { $_ -ne 0 })
    $peerScore = @($members | ForEach-Object { $r = $rankMap[[string]$_.code]; if ($r) { ConvertTo-Number $r.final_score } } | Where-Object { $_ -ne 0 })

    # ---------- Sector overview ----------
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# $($sec.name)（$($sec.key)）板块总览")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("> 静态板块底稿，成分股 $($members.Count) 只。个股研究见本目录下的独立文件；每日排名请运行 ``scripts/run-daily.ps1``。")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## 板块定位')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine($sec.positioning)
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## AI关联强度')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine($sec.ai_link)
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("## 成分股（$($members.Count) 只）")
    [void]$sb.AppendLine('')
    $rows = New-Object System.Collections.ArrayList
    foreach ($m in $members) {
        $a = $aMap[[string]$m.code]
        if ($null -eq $a) { continue }
        $rk = $rankMap[[string]$m.code]
        $curRank = if ($rk) { $rk.rank } else { '-' }
        $curScore = if ($rk) { $rk.final_score } else { '-' }
        $link = "[$($m.name)]($($m.code)-$($m.name).md)"
        [void]$rows.Add(@(
                $a.old_rank,
                $curRank,
                $m.code,
                $link,
                $a.old_cap_yi,
                $a.old_forward_pe,
                (Format-Pct $a.bear_pct),
                (Format-Pct $a.base_pct),
                (Format-Pct $a.bull_pct),
                $curScore,
                $a.notes
            ))
    }
    [void]$sb.AppendLine((New-MarkdownTable -Headers @('旧排名', '现排名', '代码', '股票', '旧市值(亿)', '旧ForwardPE', 'Bear%', 'Base%', 'Bull%', '现总分', '备注') -Rows $rows))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## 板块级尽调问题')
    [void]$sb.AppendLine('')
    foreach ($q in $sec.questions) { [void]$sb.AppendLine("- $q") }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## 数据与口径待办')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('- 核对最新收盘价、总市值、股本（含增发/A+H口径），标注数据日期与来源。')
    [void]$sb.AppendLine('- 获取最近一期财报、公告、订单与客户认证，区分公司指引、机构一致预期与自建预测。')
    [void]$sb.AppendLine('- 在 `config/assumptions.csv` 填写 2027E/2028E 利润与合理 PE，并覆盖板块占位评分。')
    [void]$sb.AppendLine('- Base>40% 的标的逐只完成订单→收入→毛利→费用→利润的利润桥。')
    [void]$sb.AppendLine('')
    Write-TextUtf8 -Path (Join-Path $secDir '00-板块总览.md') -Text $sb.ToString()

    # ---------- Per-stock research notes ----------
    foreach ($m in $members) {
        $a = $aMap[[string]$m.code]
        if ($null -eq $a) { continue }
        $rk = $rankMap[[string]$m.code]
        $q = $quoteMap[[string]$m.code]
        $f = $finMap[[string]$m.code]
        $pf = $profileMap[[string]$m.code]
        $mix = if ($mixMap.ContainsKey([string]$m.code)) { @($mixMap[[string]$m.code]) } else { @() }
        $anns = if ($annMap.ContainsKey([string]$m.code)) { @($annMap[[string]$m.code]) } else { @() }

        $stockFile = Join-Path $secDir ("{0}-{1}.md" -f $m.code, $m.name)
        if ((Test-Path -LiteralPath $stockFile) -and -not $Force) {
            $skipped++
            continue
        }

        $priceDate = if ($rk) { $rk.price_date } elseif ($q) { $q.quote_date } else { '未知' }
        $hasQuote = ($null -ne $q -and "$($q.price)" -ne '')
        $hasFin = ($null -ne $f -and "$($f.report_date)" -ne '')
        $hasForecast = ($null -ne $rk -and "$($rk.profit_basis)" -eq '人工2027E/2028E加权')
        $hasProfile = ($null -ne $pf -and "$($pf.org_name)" -ne '')
        $conf = if ($rk) { $rk.confidence } else { '低' }

        $oldCap = ConvertTo-Number $a.old_cap_yi
        $bearCap = $oldCap * (1 + (ConvertTo-Number $a.bear_pct) / 100)
        $baseCap = $oldCap * (1 + (ConvertTo-Number $a.base_pct) / 100)
        $bullCap = $oldCap * (1 + (ConvertTo-Number $a.bull_pct) / 100)
        $cap = if ($rk) { ConvertTo-Number $rk.market_cap_yi } elseif ($q) { ConvertTo-Number $q.market_cap_yi } else { $oldCap }
        $fwdProfit = if ($rk) { ConvertTo-Number $rk.forward_profit_yi } else { 0 }
        $fwdPe = if ($rk) { $rk.forward_pe } else { '' }
        $reasonablePe = ConvertTo-Number $a.reasonable_pe ([double]$model.reasonable_pe_default)

        $ssb = New-Object System.Text.StringBuilder
        [void]$ssb.AppendLine("# $($m.name)（$($m.code)）研究底稿")
        [void]$ssb.AppendLine('')
        [void]$ssb.AppendLine("> 板块：$($sec.name)（$($m.group)） ｜ 数据截至：$priceDate ｜ 生成：$now ｜ 置信度：$conf")
        [void]$ssb.AppendLine('>')
        [void]$ssb.AppendLine("> 状态：行情$(if ($hasQuote) { '已核实' } else { '未获取' })；财报$(if ($hasFin) { '已核实' } else { '未获取' })；公司画像$(if ($hasProfile) { '已获取' } else { '未获取' })；2027E/2028E 预测$(if ($hasForecast) { '已填写' } else { '未填写（暂估）' })。")
        [void]$ssb.AppendLine('>')
        [void]$ssb.AppendLine('> 本文件为研究底稿，不构成投资建议。估值锚来自 2026-11-03 旧基准，须以最新利润预测逐股重估。')
        [void]$ssb.AppendLine('')

        # 0. Dashboard
        [void]$ssb.AppendLine('## 0. 速览')
        [void]$ssb.AppendLine('')
        $dashRows = New-Object System.Collections.ArrayList
        [void]$dashRows.Add(@('现排名 / 池内', $(if ($rk) { "$($rk.rank)/74" } else { '-' }), '系统总分', $(if ($rk) { $rk.final_score } else { '-' })))
        [void]$dashRows.Add(@('Base 收益率', $(if ($rk) { Format-Pct $rk.base_return_pct } else { '-' }), 'Bull 收益率', $(if ($rk) { Format-Pct $rk.bull_return_pct } else { '-' })))
        [void]$dashRows.Add(@('总市值(亿)', $(if ($rk) { Format-Number $rk.market_cap_yi 2 } else { '-' }), 'PE(TTM)', $(if ($rk -and "$($rk.pe_ttm)" -ne '') { $rk.pe_ttm } else { '-' })))
        [void]$dashRows.Add(@('营收(亿)', $(if ($hasFin) { $f.revenue_yi } else { '-' }), '归母净利(亿)', $(if ($hasFin) { $f.net_profit_yi } else { '-' })))
        [void]$dashRows.Add(@('营收同比', $(if ($hasFin) { Format-Pct $f.revenue_yoy } else { '-' }), '净利同比', $(if ($hasFin) { Format-Pct $f.profit_yoy } else { '-' })))
        [void]$dashRows.Add(@('毛利率', $(if ($hasFin) { Format-Pct $f.gross_margin } else { '-' }), '最新报告期', $(if ($hasFin) { $f.report_date } else { '-' })))
        [void]$ssb.AppendLine((New-MarkdownTable -Headers @('指标', '数值', '指标', '数值') -Rows $dashRows))
        [void]$ssb.AppendLine('')
        if ($rk) {
            $scoreWidth = [int][math]::Round(([double]$rk.final_score / 102) * 30)
            if ($scoreWidth -lt 0) { $scoreWidth = 0 }
            if ($scoreWidth -gt 30) { $scoreWidth = 30 }
            $scoreBar = ('█' * $scoreWidth) + ('░' * (30 - $scoreWidth))
            $baseVal = [double]$rk.base_return_pct
            $baseWidth = [int][math]::Round([math]::Max(0, [math]::Min(120, $baseVal)) / 120 * 30)
            $baseBar = ('█' * $baseWidth) + ('░' * (30 - $baseWidth))
            [void]$ssb.AppendLine("- 总分：``$scoreBar`` $($rk.final_score) / 102")
            [void]$ssb.AppendLine("- Base：``$baseBar`` $(Format-Pct $baseVal)（条长按 0—120% 映射）")
            [void]$ssb.AppendLine('')
        }

        # 1. Company profile
        [void]$ssb.AppendLine('## 1. 公司画像')
        [void]$ssb.AppendLine('')
        if ($hasProfile) {
            $info = New-Object System.Collections.ArrayList
            [void]$info.Add(@('公司全称', $pf.org_name))
            [void]$info.Add(@('东财行业', $(if ("$($pf.industry_em)" -ne '') { $pf.industry_em } else { '-' })))
            [void]$info.Add(@('证监会行业', $(if ("$($pf.industry_csrc)" -ne '') { $pf.industry_csrc } else { '-' })))
            [void]$info.Add(@('地区 / 员工', "$(if ("$($pf.province)" -ne '') { $pf.province } else { '-' }) / $(if ("$($pf.employees)" -ne '') { "$($pf.employees) 人" } else { '-' })"))
            [void]$info.Add(@('董事长 / 注册资本', "$(if ("$($pf.chairman)" -ne '') { $pf.chairman } else { '-' }) / $(if ("$($pf.reg_capital_yi)" -ne '') { "$($pf.reg_capital_yi) 亿" } else { '-' })"))
            [void]$info.Add(@('官网', $(if ("$($pf.website)" -ne '') { $pf.website } else { '-' })))
            [void]$ssb.AppendLine((New-MarkdownTable -Headers @('项目', '内容') -Rows $info))
            [void]$ssb.AppendLine('')
            if ("$($pf.business_scope)" -ne '') {
                [void]$ssb.AppendLine('**主营范围**：' + $pf.business_scope)
                [void]$ssb.AppendLine('')
            }
            if ("$($pf.profile)" -ne '') {
                [void]$ssb.AppendLine('**公司简介**：' + $pf.profile)
                [void]$ssb.AppendLine('')
            }
        }
        else {
            [void]$ssb.AppendLine('未获取公司画像。请运行 `scripts/fetch-profiles.ps1` 后重新生成本文件。')
            [void]$ssb.AppendLine('')
        }

        # 2. Products and revenue mix
        [void]$ssb.AppendLine('## 2. 产品结构与收入占比')
        [void]$ssb.AppendLine('')
        if ($mix.Count -gt 0) {
            $reportDate = $mix[0].report_date
            [void]$ssb.AppendLine("> 数据报告期：$reportDate（最新披露的主营构成）。占比为营业收入口径；毛利率 = 主营利润 ÷ 主营收入。")
            [void]$ssb.AppendLine('')
            $productRows = @($mix | Where-Object { $_.dimension -eq '按产品' })
            if ($productRows.Count -eq 0) { $productRows = @($mix | Where-Object { $_.dimension -eq '按行业' }) }
            if ($productRows.Count -gt 0) {
                [void]$ssb.AppendLine('### 2.1 按产品/业务')
                [void]$ssb.AppendLine('')
                $t = New-Object System.Collections.ArrayList
                foreach ($x in ($productRows | Sort-Object -Property { -[double]$_.share_pct })) {
                    $bar = New-TextBar -Pct ([double]$x.share_pct) -Width 20
                    $gmText = if ([double]$x.gross_margin_pct -gt 0) { Format-Pct $x.gross_margin_pct } else { '-' }
                    [void]$t.Add(@($x.item, (Format-Number $x.revenue_yi 2), (Format-Pct $x.share_pct), $gmText, "``$bar``"))
                }
                [void]$ssb.AppendLine((New-MarkdownTable -Headers @('产品/业务', '收入(亿)', '收入占比', '毛利率', '占比可视化') -Rows $t))
                [void]$ssb.AppendLine('')
            }
            $industryRows = @($mix | Where-Object { $_.dimension -eq '按行业' })
            if ($industryRows.Count -gt 0) {
                [void]$ssb.AppendLine('### 2.2 按行业')
                [void]$ssb.AppendLine('')
                $t2 = New-Object System.Collections.ArrayList
                foreach ($x in ($industryRows | Sort-Object -Property { -[double]$_.share_pct } | Select-Object -First 10)) {
                    $gmText = if ([double]$x.gross_margin_pct -gt 0) { Format-Pct $x.gross_margin_pct } else { '-' }
                    [void]$t2.Add(@($x.item, (Format-Number $x.revenue_yi 2), (Format-Pct $x.share_pct), $gmText))
                }
                [void]$ssb.AppendLine((New-MarkdownTable -Headers @('行业', '收入(亿)', '收入占比', '毛利率') -Rows $t2))
                [void]$ssb.AppendLine('')
            }
            $regionRows = @($mix | Where-Object { $_.dimension -eq '按地区' })
            if ($regionRows.Count -gt 0) {
                [void]$ssb.AppendLine('### 2.3 按地区')
                [void]$ssb.AppendLine('')
                $t3 = New-Object System.Collections.ArrayList
                foreach ($x in ($regionRows | Sort-Object -Property { -[double]$_.share_pct } | Select-Object -First 10)) {
                    $gmText = if ([double]$x.gross_margin_pct -gt 0) { Format-Pct $x.gross_margin_pct } else { '-' }
                    [void]$t3.Add(@($x.item, (Format-Number $x.revenue_yi 2), (Format-Pct $x.share_pct), $gmText))
                }
                [void]$ssb.AppendLine((New-MarkdownTable -Headers @('地区', '收入(亿)', '收入占比', '毛利率') -Rows $t3))
                [void]$ssb.AppendLine('')
            }
            [void]$ssb.AppendLine('> 注意：部分公司最新报告期只有年报口径；占比合计可能因“其他(补充)”不等于100%。')
            [void]$ssb.AppendLine('')
        }
        else {
            [void]$ssb.AppendLine('未获取主营构成数据。请运行 `scripts/fetch-profiles.ps1` 后重新生成本文件。')
            [void]$ssb.AppendLine('')
        }

        # 3. Orders and business dynamics
        [void]$ssb.AppendLine('## 3. 订单与经营动态')
        [void]$ssb.AppendLine('')
        $reviewText = if ($pf) { [string]$pf.review_text } else { '' }
        $orderHits = Get-SentenceHighlights -Text $reviewText -Pattern '订单|在手|排产|交付|中标|合同|供不应求' -Max 4
        $orderHits += Get-SentenceHighlights -Text $reviewText -Pattern '扩产|产能|投产|量产|客户|认证|涨价|提价' -Max 4
        $orderHits = @($orderHits | Select-Object -Unique | Select-Object -First 6)
        if ($orderHits.Count -gt 0) {
            [void]$ssb.AppendLine("### 3.1 经营评述中的订单/产能线索（$(if ($pf) { $pf.review_date } else { '未知' })）")
            [void]$ssb.AppendLine('')
            foreach ($h in $orderHits) { [void]$ssb.AppendLine("- $h") }
            [void]$ssb.AppendLine('')
        }
        else {
            [void]$ssb.AppendLine('经营评述中未提取到明确的订单/产能表述。')
            [void]$ssb.AppendLine('')
        }
        [void]$ssb.AppendLine('### 3.2 公告线索（订单/合同/扩产/客户等关键词）')
        [void]$ssb.AppendLine('')
        $annPattern = '订单|中标|合同|框架协议|扩产|产能|投产|量产|客户|认证|定点|供应|合作|涨价|提价|业绩预告|业绩快报|调研|项目'
        $relevant = @($anns | Where-Object { "$($_.title)" -match $annPattern } | Select-Object -First 12)
        if ($relevant.Count -gt 0) {
            $ta = New-Object System.Collections.ArrayList
            foreach ($x in $relevant) { [void]$ta.Add(@($x.date, $x.type, "[$($x.title -replace '\|', '｜')]($($x.url))")) }
            [void]$ssb.AppendLine((New-MarkdownTable -Headers @('日期', '类型', '公告') -Rows $ta))
        }
        else {
            [void]$ssb.AppendLine('近 40 条公告中未出现订单/合同/扩产等关键词。')
        }
        [void]$ssb.AppendLine('')
        [void]$ssb.AppendLine('### 3.3 最近公告')
        [void]$ssb.AppendLine('')
        if ($anns.Count -gt 0) {
            $tl = New-Object System.Collections.ArrayList
            foreach ($x in ($anns | Select-Object -First 5)) { [void]$tl.Add(@($x.date, $x.type, "[$($x.title -replace '\|', '｜')]($($x.url))")) }
            [void]$ssb.AppendLine((New-MarkdownTable -Headers @('日期', '类型', '公告') -Rows $tl))
        }
        else {
            [void]$ssb.AppendLine('未获取公告数据。')
        }
        [void]$ssb.AppendLine('')
        [void]$ssb.AppendLine('> **在手订单口径说明**：公开接口没有结构化的“在手订单金额”。上表是公告标题与经营评述的线索，不等于完整订单簿；订单金额、交付周期与收入确认需回查定期报告、业绩说明会或调研纪要。')
        [void]$ssb.AppendLine('')

        # 4. Competitive position
        [void]$ssb.AppendLine('## 4. 竞争地位')
        [void]$ssb.AppendLine('')
        [void]$ssb.AppendLine("### 4.1 池内同业对比（$($sec.name)）")
        [void]$ssb.AppendLine('')
        $peerRows = New-Object System.Collections.ArrayList
        $peerSorted = @($members | Sort-Object -Property { $r = $rankMap[[string]$_.code]; if ($r) { -[double](ConvertTo-Number $r.market_cap_yi) } else { 0 } })
        foreach ($pm in $peerSorted) {
            $pr = $rankMap[[string]$pm.code]
            $pfin = $finMap[[string]$pm.code]
            $nameText = if ($pm.code -eq $m.code) { "**$($pm.name)**" } else { $pm.name }
            [void]$peerRows.Add(@(
                    $nameText,
                    $(if ($pr) { Format-Number $pr.market_cap_yi 0 } else { '-' }),
                    $(if ($pfin -and "$($pfin.revenue_yi)" -ne '') { Format-Number $pfin.revenue_yi 1 } else { '-' }),
                    $(if ($pfin -and "$($pfin.net_profit_yi)" -ne '') { Format-Number $pfin.net_profit_yi 1 } else { '-' }),
                    $(if ($pfin -and "$($pfin.revenue_yoy)" -ne '') { Format-Pct $pfin.revenue_yoy } else { '-' }),
                    $(if ($pfin -and "$($pfin.gross_margin)" -ne '') { Format-Pct $pfin.gross_margin } else { '-' }),
                    $(if ($pr -and "$($pr.pe_ttm)" -ne '') { $pr.pe_ttm } else { '-' }),
                    $(if ($pr) { $pr.final_score } else { '-' }),
                    $(if ($pr) { $pr.rank } else { '-' })
                ))
        }
        [void]$ssb.AppendLine((New-MarkdownTable -Headers @('股票', '市值(亿)', '营收(亿)', '归母净利(亿)', '营收同比', '毛利率', 'PE(TTM)', '总分', '池内排名') -Rows $peerRows))
        [void]$ssb.AppendLine('')
        $revVal = if ($hasFin) { ConvertTo-Number $f.revenue_yi } else { 0 }
        $profitVal = if ($hasFin) { ConvertTo-Number $f.net_profit_yi } else { 0 }
        $gmVal = if ($hasFin) { ConvertTo-Number $f.gross_margin } else { 0 }
        $scoreVal = if ($rk) { ConvertTo-Number $rk.final_score } else { 0 }
        [void]$ssb.AppendLine("**池内位置**：总市值 $(Get-PositionText $cap $peerCap)；营收 $(Get-PositionText $revVal $peerRev)；归母净利 $(Get-PositionText $profitVal $peerProfit)；毛利率 $(Get-PositionText $gmVal $peerGm)；系统总分 $(Get-PositionText $scoreVal $peerScore)。")
        [void]$ssb.AppendLine('')
        $positionText = if ($pf) { "$($pf.profile) $reviewText" } else { $reviewText }
        $positionHits = Get-SentenceHighlights -Text $positionText -Pattern '公司.*(市场占有率|市场份额|排名|第[一二三四五六七八九十0-9]+|领先|龙头|首家|唯一|主导)|(市场占有率|市场份额|排名|领先|龙头|首家|唯一|主导).*公司' -Max 5
        if ($positionHits.Count -eq 0) {
            $positionHits = Get-SentenceHighlights -Text $positionText -Pattern '市场占有率|市场份额|排名|第[一二三四五六七八九十0-9]+|领先|龙头|首家|唯一|主导' -Max 3
        }
        if ($positionHits.Count -gt 0) {
            [void]$ssb.AppendLine('### 4.2 行业地位线索（自动摘录，需核对原文）')
            [void]$ssb.AppendLine('')
            foreach ($h in $positionHits) { [void]$ssb.AppendLine("- $h") }
            [void]$ssb.AppendLine('')
        }
        [void]$ssb.AppendLine('> 池内排名仅覆盖本项目 74 只股票，不代表全行业排名；全行业份额与竞争对手需结合 Prismark/行业协会/公司公告核对。')
        [void]$ssb.AppendLine('')

        # 5. Market data and financials
        [void]$ssb.AppendLine('## 5. 行情、财务与估值')
        [void]$ssb.AppendLine('')
        $infoRows = New-Object System.Collections.ArrayList
        [void]$infoRows.Add(@('代码 / 交易所', "$($m.code) / $($m.exchange)"))
        [void]$infoRows.Add(@('所属板块', $sec.name))
        [void]$infoRows.Add(@('旧排名 / 旧总分', "$($a.old_rank) / $($a.old_score)"))
        [void]$infoRows.Add(@('当日排名 / 变化', $(if ($rk) { "$($rk.rank) / $(Format-Signed $rk.rank_change)" } else { '-' })))
        [void]$infoRows.Add(@('收盘价', $(if ($hasQuote) { $q.price } else { '-' })))
        [void]$infoRows.Add(@('总市值（亿元）', $(if ($rk) { Format-Number $rk.market_cap_yi 2 } else { '-' })))
        [void]$infoRows.Add(@('PE(TTM) / PB', "$(if ($rk -and "$($rk.pe_ttm)" -ne '') { $rk.pe_ttm } else { '-' }) / $(if ($q -and "$($q.pb)" -ne '') { $q.pb } else { '-' })"))
        [void]$infoRows.Add(@('旧市值 / 旧Forward PE', "$($a.old_cap_yi) 亿 / $($a.old_forward_pe)x"))
        [void]$infoRows.Add(@('当前Forward PE', $(if ($fwdPe) { "$fwdPe x" } else { '-' })))
        [void]$infoRows.Add(@('最新报告期', $(if ($hasFin) { "$($f.report_date) $($f.report_name)" } else { '-' })))
        [void]$infoRows.Add(@('营收 / 归母净利', $(if ($hasFin) { "$($f.revenue_yi) 亿 / $($f.net_profit_yi) 亿" } else { '-' })))
        [void]$infoRows.Add(@('营收同比 / 净利同比', $(if ($hasFin) { "$(Format-Pct $f.revenue_yoy) / $(Format-Pct $f.profit_yoy)" } else { '-' })))
        [void]$infoRows.Add(@('毛利率 / ROE', $(if ($hasFin) { "$(Format-Pct $f.gross_margin) / $(Format-Pct $f.roe)" } else { '-' })))
        [void]$ssb.AppendLine((New-MarkdownTable -Headers @('项目', '数值') -Rows $infoRows))
        [void]$ssb.AppendLine('')

        # 6. Scenarios
        [void]$ssb.AppendLine('## 6. 三情景（旧锚敏感性）')
        [void]$ssb.AppendLine('')
        $scenarioRows = New-Object System.Collections.ArrayList
        $scenarios = @(
            @{ name = 'Bear'; cap = $bearCap; pct = $a.bear_pct },
            @{ name = 'Base'; cap = $baseCap; pct = $a.base_pct },
            @{ name = 'Bull'; cap = $bullCap; pct = $a.bull_pct }
        )
        foreach ($s in $scenarios) {
            $rel = if ($cap -gt 0) { (($s.cap / $cap) - 1) * 100 } else { 0 }
            $impliedPe = if ($fwdProfit -gt 0) { [math]::Round($s.cap / $fwdProfit, 1) } else { '-' }
            $needProfit = if ($reasonablePe -gt 0) { [math]::Round($s.cap / $reasonablePe, 1) } else { '-' }
            [void]$scenarioRows.Add(@($s.name, (Format-Number $s.cap 1), (Format-Pct $rel), $impliedPe, $needProfit))
        }
        [void]$ssb.AppendLine((New-MarkdownTable -Headers @('情景', '目标市值(亿)', '相对当日市值', '隐含PE', "所需利润(${reasonablePe}x)") -Rows $scenarioRows))
        [void]$ssb.AppendLine('')
        [void]$ssb.AppendLine('- 目标市值 = 旧市值 ×（1 + 旧情景收益率）；相对当日市值用当日总市值重算。')
        [void]$ssb.AppendLine('- 隐含PE = 目标市值 ÷ 当前远期利润；当前远期利润在未填预测时由旧市值 ÷ 旧 Forward PE 反推。')
        [void]$ssb.AppendLine('- 该表是敏感性分析，不是新估值；必须用最新订单和利润预测替换旧锚。')
        [void]$ssb.AppendLine('')

        # 7. Scores
        [void]$ssb.AppendLine("## 7. 评分明细（$priceDate 暂定）")
        [void]$ssb.AppendLine('')
        $scoreRows = New-Object System.Collections.ArrayList
        if ($rk) {
            [void]$scoreRows.Add(@('未来1—2年收益率', $rk.score_return, '25', 'Base收益率线性映射'))
            [void]$scoreRows.Add(@('上下行赔率', $rk.score_odds, '25', '期望收益 ÷ 下行幅度，Bear<−60%打8折'))
            [void]$scoreRows.Add(@('未计价期权/第二曲线', $rk.score_option, '17', '板块级占位，待逐股覆盖'))
            [void]$scoreRows.Add(@('AI/产业需求弹性', $rk.score_ai, '6', '板块级占位，待逐股覆盖'))
            [void]$scoreRows.Add(@('盈利加速/预期上修', $rk.score_earnings, '11', $rk.earnings_method))
            [void]$scoreRows.Add(@('未来6—12个月催化', $rk.score_catalyst, '7', '板块级占位，待逐股覆盖'))
            [void]$scoreRows.Add(@('产业地位+订单+确定性', $rk.score_certainty, '9', $rk.certainty_method))
            $riskText = if ((ConvertTo-Number $rk.risk_deduction) -eq 0) { '0' } else { '−' + $rk.risk_deduction }
            [void]$scoreRows.Add(@('风险扣分', $riskText, '0—20', 'PE、Bear尾部、负Base、亏损规则'))
            [void]$scoreRows.Add(@('高凸性微调', ('+' + $rk.convexity_bonus), '0—2', 'Bull/Base组合规则'))
            [void]$scoreRows.Add(@('**总分**', "**$($rk.final_score)**", '−20—102', "置信度：$conf"))
        }
        else {
            [void]$scoreRows.Add(@('暂无评分', '-', '-', '请先运行 scripts/run-daily.ps1'))
        }
        [void]$ssb.AppendLine((New-MarkdownTable -Headers @('维度', '得分', '满分', '方法/证据') -Rows $scoreRows))
        [void]$ssb.AppendLine('')

        # 8. Profit bridge
        [void]$ssb.AppendLine('## 8. 利润桥（Base>40% 强制）')
        [void]$ssb.AppendLine('')
        if ($rk -and "$($rk.bridge_flag)" -ne '') {
            $bridgeRows = New-Object System.Collections.ArrayList
            [void]$bridgeRows.Add(@('当前总市值', (Format-Number $rk.market_cap_yi 2)))
            [void]$bridgeRows.Add(@('Base目标市值', (Format-Number $rk.base_target_cap_yi 2)))
            [void]$bridgeRows.Add(@('当前远期利润', (Format-Number $rk.forward_profit_yi 2)))
            [void]$bridgeRows.Add(@('所需利润', (Format-Number $rk.required_profit_yi 2)))
            [void]$bridgeRows.Add(@('利润缺口', (Format-Number $rk.profit_gap_yi 2)))
            [void]$bridgeRows.Add(@('初检结论', $rk.bridge_flag))
            [void]$bridgeRows.Add(@('利润口径', $rk.profit_basis))
            [void]$ssb.AppendLine((New-MarkdownTable -Headers @('项目', '数值(亿元)') -Rows $bridgeRows))
            [void]$ssb.AppendLine('')
            [void]$ssb.AppendLine('> 还须直接核对：订单 → 收入 → 毛利率 → 费用 → 税率 → 归母净利润，并解释新增市值来自利润增长还是估值扩张。')
        }
        else {
            [void]$ssb.AppendLine('Base收益率未超过40%，未触发强制利润桥；仍需验证2027E/2028E利润假设。')
        }
        [void]$ssb.AppendLine('')

        # 9. Original assumptions
        [void]$ssb.AppendLine('## 9. 原始假设与备注')
        [void]$ssb.AppendLine('')
        [void]$ssb.AppendLine($(if ("$($a.notes)" -ne '') { $a.notes } else { '暂无。' }))
        [void]$ssb.AppendLine('')

        # 10. Due diligence questions
        [void]$ssb.AppendLine('## 10. 尽调问题')
        [void]$ssb.AppendLine('')
        foreach ($question in $sec.questions) { [void]$ssb.AppendLine("- $question") }
        [void]$ssb.AppendLine('- 该公司AI相关收入占比、订单能见度、客户集中度分别是多少？')
        [void]$ssb.AppendLine('- 2027E/2028E 利润预测与合理 PE 的依据是什么（公司指引/一致预期/自建）？')
        [void]$ssb.AppendLine('- 最近一期财报、公告或订单是否改变了原假设？')
        [void]$ssb.AppendLine('')

        # 11. Catalysts and risks
        [void]$ssb.AppendLine('## 11. 催化与风险')
        [void]$ssb.AppendLine('')
        $catalystRows = New-Object System.Collections.ArrayList
        [void]$catalystRows.Add(@('催化', '（待填：认证/量产/订单/产能/财报时点）', '（时间）', '未验证'))
        [void]$catalystRows.Add(@('风险', '（待填：主业利空/估值/客户集中/兑现困难）', '（时间）', '未验证'))
        [void]$ssb.AppendLine((New-MarkdownTable -Headers @('类型', '内容', '时点', '状态') -Rows $catalystRows))
        [void]$ssb.AppendLine('')

        # 12. Evidence
        [void]$ssb.AppendLine('## 12. 证据与待核实')
        [void]$ssb.AppendLine('')
        [void]$ssb.AppendLine("- 行情：东方财富公开接口，$priceDate 收盘快照（``data/quotes/$quoteFileName``）。")
        [void]$ssb.AppendLine("- 财报：东方财富业绩报表接口，报告期 $(if ($hasFin) { $f.report_date } else { '未获取' })（``data/financials/$finFileName``）。")
        [void]$ssb.AppendLine("- 公司画像/主营构成/经营评述：东方财富 F10，报告期 $(if ($pf) { $pf.review_date } else { '未获取' })（``data/profiles/``）。")
        [void]$ssb.AppendLine('- 未核实：2027E/2028E 利润、合理 PE、订单金额、客户份额、产能数据、最新公告影响。')
        [void]$ssb.AppendLine('- 待补：公司公告、投资者关系记录、机构研报、产业链交叉验证。')
        [void]$ssb.AppendLine('')

        # 13. Research log
        [void]$ssb.AppendLine('## 13. 研究日志')
        [void]$ssb.AppendLine('')
        $logRows = New-Object System.Collections.ArrayList
        [void]$logRows.Add(@($now, '系统初版', '-', '-', '待研究'))
        [void]$ssb.AppendLine((New-MarkdownTable -Headers @('日期', '原假设', '新事实', '情景/评分变化', '结论') -Rows $logRows))
        [void]$ssb.AppendLine('')

        # 14. Appendix: full business review
        if ($reviewText -ne '') {
            [void]$ssb.AppendLine('## 14. 附录：经营评述全文')
            [void]$ssb.AppendLine('')
            [void]$ssb.AppendLine("> 来源：东方财富 F10 经营评述，报告期 $(if ($pf) { $pf.review_date } else { '未知' })。原文照录，未做事实核验。")
            [void]$ssb.AppendLine('')
            $paragraphs = $reviewText -split "`r?`n"
            foreach ($para in $paragraphs) {
                $t = $para.Trim()
                if ($t -ne '') {
                    [void]$ssb.AppendLine($t)
                    [void]$ssb.AppendLine('')
                }
            }
        }

        Write-TextUtf8 -Path $stockFile -Text $ssb.ToString()
        $created++
    }

    Write-Host ("{0}: {1} stocks" -f $folderName, $members.Count)
}

Write-Host ("created/updated={0}, skipped(existing)={1}" -f $created, $skipped)
if ($skipped -gt 0) { Write-Host 'Use -Force to regenerate existing per-stock notes.' }
