# Markdown / CSV report generation.
$ErrorActionPreference = 'Stop'

function Format-Signed {
    param($Value)
    $n = [int]$Value
    if ($n -gt 0) { return "+$n" }
    return "$n"
}

function Format-Cell {
    param($Value, [string]$Suffix = '')
    if ($null -eq $Value -or "$Value" -eq '') { return '-' }
    return ("$Value$Suffix")
}

function Format-NameList {
    param($Items)
    $arr = @($Items | Where-Object { $null -ne $_ -and "$_" -ne '' })
    if ($arr.Count -eq 0) { return '无' }
    return ($arr -join '、')
}

function New-MarkdownTable {
    param([string[]]$Headers, [array]$Rows)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('| ' + ($Headers -join ' | ') + ' |')
    [void]$sb.AppendLine('|' + (($Headers | ForEach-Object { '---' }) -join '|') + '|')
    foreach ($r in $Rows) {
        [void]$sb.AppendLine('| ' + (($r | ForEach-Object { if ($null -eq $_) { '' } else { [string]$_ } }) -join ' | ') + ' |')
    }
    return $sb.ToString()
}

function New-DailyReport {
    param([array]$Ranked, [string]$TradeDate, $Model)

    $sb = New-Object System.Text.StringBuilder
    $now = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $count = $Ranked.Count
    $withQuote = @($Ranked | Where-Object { $_.market_cap_status -eq '已核实行情市值' }).Count
    $withFin = @($Ranked | Where-Object { "$($_.report_date)" -ne '' }).Count
    $withForecast = @($Ranked | Where-Object { $_.profit_basis -eq '人工2027E/2028E加权' }).Count
    $highConf = @($Ranked | Where-Object { $_.confidence -eq '高' }).Count
    $baseOver40 = @($Ranked | Where-Object { [double]$_.base_return_pct -gt 40 })
    $top10 = @($Ranked | Select-Object -First 10)
    $avgBase = if ($count -gt 0) { ($Ranked | Measure-Object -Property base_return_pct -Average).Average } else { 0 }

    [void]$sb.AppendLine("# AI高赔率股票池 $($Model.model_version) 暂定排名")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("> 交易日：$TradeDate ｜ 生成时间：$now ｜ 覆盖：$count 只")
    [void]$sb.AppendLine('>')
    [void]$sb.AppendLine('> **本报告为系统暂定排名，不是已核实研究结论。** 行情来自东方财富公开接口；利润预测未填写时以旧Forward PE反推，仅作敏感性起点；定性评分为板块级占位值。所有未人工核实项须按《项目交接》要求标注【暂估】。')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## 1. 当日摘要')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("- 已核实行情：$withQuote / $count；已核实财报：$withFin / $count；已填2027E/2028E预测：$withForecast / $count；高置信度：$highConf / $count")
    [void]$sb.AppendLine("- 全池Base收益率均值：$([math]::Round($avgBase, 1))%；Base>40%共 $($baseOver40.Count) 只（逐只利润桥见第4节）")
    [void]$sb.AppendLine("- 暂定前10：" + (($top10 | ForEach-Object { "$($_.rank).$($_.name)" }) -join '、'))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## 2. 完整排名（1—' + $count + '）')
    [void]$sb.AppendLine('')

    $headers = @('排名', '变化', '代码', '股票', '板块', '收盘价', '市值(亿)', 'PE(TTM)', 'ForwardPE', 'Bear%', 'Base%', 'Bull%', '期望%', '赔率', '总分', '数据')
    $tableRows = New-Object System.Collections.ArrayList
    foreach ($r in $Ranked) {
        [void]$tableRows.Add(@(
                $r.rank,
                (Format-Signed $r.rank_change),
                $r.code,
                $r.name,
                $r.group,
                (Format-Cell $r.close_price),
                (Format-Number $r.market_cap_yi 0),
                (Format-Cell $r.pe_ttm),
                (Format-Cell $r.forward_pe),
                (Format-Pct $r.bear_return_pct),
                (Format-Pct $r.base_return_pct),
                (Format-Pct $r.bull_return_pct),
                (Format-Pct $r.expected_return_pct),
                (Format-Number $r.upside_downside_ratio 2),
                (Format-Number $r.final_score 1),
                $r.confidence
            ))
    }
    [void]$sb.AppendLine((New-MarkdownTable -Headers $headers -Rows $tableRows))
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('## 3. 前20与重大名次变化')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('### 3.1 暂定前20')
    [void]$sb.AppendLine('')
    $top20 = @($Ranked | Select-Object -First 20)
    $t20 = New-Object System.Collections.ArrayList
    foreach ($r in $top20) {
        [void]$t20.Add(@($r.rank, (Format-Signed $r.rank_change), $r.name, $r.group, (Format-Pct $r.base_return_pct), (Format-Pct $r.bear_return_pct), (Format-Number $r.final_score 1), $r.confidence))
    }
    [void]$sb.AppendLine((New-MarkdownTable -Headers @('排名', '变化', '股票', '板块', 'Base%', 'Bear%', '总分', '数据') -Rows $t20))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('### 3.2 重大升降（名次变化绝对值≥10）')
    [void]$sb.AppendLine('')
    $movers = @($Ranked | Where-Object { [math]::Abs([int]$_.rank_change) -ge 10 } | Sort-Object -Property @{ Expression = { [math]::Abs([int]$_.rank_change) }; Descending = $true })
    if ($movers.Count -eq 0) {
        [void]$sb.AppendLine('无。')
    }
    else {
        $mr = New-Object System.Collections.ArrayList
        foreach ($r in $movers) {
            [void]$mr.Add(@($r.name, $r.group, $r.old_rank, $r.rank, (Format-Signed $r.rank_change), (Format-Pct $r.base_return_pct), (Format-Number $r.final_score 1), $r.confidence))
        }
        [void]$sb.AppendLine((New-MarkdownTable -Headers @('股票', '板块', '旧排名', '新排名', '变化', 'Base%', '总分', '数据') -Rows $mr))
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('## 4. Base>40% 市值—利润兑现初检')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('> 所需利润 = Base目标市值 ÷ 合理PE（默认25×）；当前远期利润在未填预测时由旧市值÷旧Forward PE反推。该桥仅作辅助检验，不替代逐股订单→收入→利润拆解。')
    [void]$sb.AppendLine('')
    if ($baseOver40.Count -eq 0) {
        [void]$sb.AppendLine('无。')
    }
    else {
        $br = New-Object System.Collections.ArrayList
        foreach ($r in $baseOver40) {
            [void]$br.Add(@($r.rank, $r.name, (Format-Number $r.market_cap_yi 0), (Format-Number $r.base_target_cap_yi 0), (Format-Number $r.forward_profit_yi 1), (Format-Number $r.required_profit_yi 1), (Format-Number $r.profit_gap_yi 1), $r.bridge_flag, $r.profit_basis))
        }
        [void]$sb.AppendLine((New-MarkdownTable -Headers @('排名', '股票', '当前市值', 'Base目标市值', '当前远期利润', '所需利润', '利润缺口', '初检', '利润口径') -Rows $br))
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('## 5. 板块汇总')
    [void]$sb.AppendLine('')
    $sr = New-Object System.Collections.ArrayList
    foreach ($g in ($Ranked | Group-Object -Property group | Sort-Object -Property Name)) {
        $avgScore = ($g.Group | Measure-Object -Property final_score -Average).Average
        $avgBaseG = ($g.Group | Measure-Object -Property base_return_pct -Average).Average
        $best = $g.Group | Sort-Object -Property final_score -Descending | Select-Object -First 1
        [void]$sr.Add(@($g.Name, $g.Count, (Format-Number $avgScore 1), (Format-Pct $avgBaseG), $best.name, (Format-Number $best.final_score 1)))
    }
    [void]$sb.AppendLine((New-MarkdownTable -Headers @('板块', '数量', '平均总分', '平均Base%', '板块第一', '其总分') -Rows $sr))
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('## 6. 数据缺口与待办')
    [void]$sb.AppendLine('')
    $noQuote = @($Ranked | Where-Object { $_.market_cap_status -ne '已核实行情市值' })
    $noFin = @($Ranked | Where-Object { "$($_.report_date)" -eq '' })
    $noForecast = @($Ranked | Where-Object { $_.profit_basis -ne '人工2027E/2028E加权' })
    $strongCheck = @($Ranked | Where-Object { $_.bridge_flag -eq '需强验证' })
    [void]$sb.AppendLine("- 未获取行情（$($noQuote.Count)）：" + (Format-NameList ($noQuote | ForEach-Object { $_.name })))
    [void]$sb.AppendLine("- 未获取财报（$($noFin.Count)）：" + (Format-NameList ($noFin | ForEach-Object { $_.name })))
    [void]$sb.AppendLine("- 未填2027E/2028E预测（$($noForecast.Count)）：" + (Format-NameList ($noForecast | ForEach-Object { $_.name })))
    [void]$sb.AppendLine("- Base>40%且利润桥需强验证（$($strongCheck.Count)）：" + (Format-NameList ($strongCheck | ForEach-Object { $_.name })))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('## 7. 方法与免责声明')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('- 评分：七维合计（收益25/赔率25/期权17/AI 6/加速11/催化7/确定性9）− 风险扣分（0—20）+ 凸性（0—2）。')
    [void]$sb.AppendLine('- 估值：沿用2026-11-03旧基准Bear/Base/Bull目标市值锚，用当日市值重算情景收益率；这是敏感性口径，不是重新估值。')
    [void]$sb.AppendLine('- 情景概率为暂定25%/50%/25%，未经用户冻结；期望收益与赔率仅用于排序辅助。')
    [void]$sb.AppendLine('- 定性维度（期权/AI/催化）当前为板块级占位值，需逐股研究后在 `config/assumptions.csv` 中覆盖。')
    [void]$sb.AppendLine('- 本报告不构成投资建议；旧名次仅用于变化检查，不作为评分输入。')
    return $sb.ToString()
}

function Save-DailyOutputs {
    param([array]$Ranked, [string]$TradeDate, $Model)

    $dir = Join-ProjectPath ("output/$TradeDate")
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

    $rankFile = Join-Path $dir 'rank.csv'
    Write-CsvUtf8 -Path $rankFile -Data $Ranked

    $report = New-DailyReport -Ranked $Ranked -TradeDate $TradeDate -Model $Model
    $reportFile = Join-Path $dir 'report.md'
    Write-TextUtf8 -Path $reportFile -Text $report

    $latest = Join-ProjectPath 'output/latest'
    if (Test-Path -LiteralPath $latest) { Remove-Item -LiteralPath $latest -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $latest | Out-Null
    Copy-Item -LiteralPath $rankFile -Destination (Join-Path $latest 'rank.csv') -Force
    Copy-Item -LiteralPath $reportFile -Destination (Join-Path $latest 'report.md') -Force

    return [pscustomobject]@{ rank_file = $rankFile; report_file = $reportFile; latest_file = (Join-Path $latest 'report.md') }
}
