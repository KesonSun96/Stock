# Scoring and valuation engine for model V2.4.3.
$ErrorActionPreference = 'Stop'

function Clamp-Value {
    param([double]$Value, [double]$Min = 0, [double]$Max = 1)
    if ($Value -lt $Min) { return $Min }
    if ($Value -gt $Max) { return $Max }
    return $Value
}

function Invoke-Scoring {
    param([string]$TradeDate)

    $model = Read-JsonFile -Path (Join-ProjectPath 'config/model.json')
    $pool = Read-CsvData -Path (Join-ProjectPath 'config/stock_pool.csv')
    $assumptions = Read-CsvData -Path (Join-ProjectPath 'config/assumptions.csv')
    $quotes = Read-CsvData -Path (Join-ProjectPath ("data/quotes/quotes-$TradeDate.csv"))
    $finFile = Get-ChildItem -Path (Join-ProjectPath 'data/financials') -Filter '*.csv' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $financials = if ($finFile) { Read-CsvData -Path $finFile.FullName } else { @() }

    $aMap = @{}; foreach ($a in $assumptions) { $aMap[[string]$a.code] = $a }
    $qMap = @{}; foreach ($q in $quotes) { $qMap[[string]$q.code] = $q }
    $fMap = @{}; foreach ($f in $financials) { $fMap[[string]$f.code] = $f }

    $w = $model.weights
    $p = $model.probabilities
    $rows = New-Object System.Collections.ArrayList

    foreach ($s in $pool) {
        $a = $aMap[[string]$s.code]
        $q = $qMap[[string]$s.code]
        $f = $fMap[[string]$s.code]
        if ($null -eq $a) { continue }

        $oldCap = ConvertTo-Number $a.old_cap_yi
        $oldFpe = ConvertTo-Number $a.old_forward_pe
        $bear0 = (ConvertTo-Number $a.bear_pct) / 100
        $base0 = (ConvertTo-Number $a.base_pct) / 100
        $bull0 = (ConvertTo-Number $a.bull_pct) / 100

        $cap = if ($q) { ConvertTo-Number $q.market_cap_yi } else { 0 }
        $capStatus = '已核实行情市值'
        if ($cap -le 0) { $cap = $oldCap; $capStatus = '旧基准市值（未获取行情）' }

        $bearCap = $oldCap * (1 + $bear0)
        $baseCap = $oldCap * (1 + $base0)
        $bullCap = $oldCap * (1 + $bull0)

        $bear = $bearCap / $cap - 1
        $base = $baseCap / $cap - 1
        $bull = $bullCap / $cap - 1

        $exp = [double]$p.bear * $bear + [double]$p.base * $base + [double]$p.bull * $bull
        $downside = [math]::Max([math]::Abs($bear), 0.05)
        $ratio = [math]::Max($exp, 0) / $downside

        $returnScore = [double]$w.return * (Clamp-Value (($base * 100 - [double]$model.return_curve.zero_at_pct) / ([double]$model.return_curve.full_at_pct - [double]$model.return_curve.zero_at_pct)))
        $oddsScore = [double]$w.odds * (Clamp-Value ($ratio / [double]$model.odds_curve.full_ratio))
        if ($bear * 100 -lt [double]$model.odds_curve.tail_penalty_bear_pct) {
            $oddsScore = $oddsScore * [double]$model.odds_curve.tail_penalty_factor
        }

        $hasFin = $f -and $f.data_status -eq '已核实财报' -and "$($f.net_profit_yi)" -ne ''
        $profitYoy = if ($hasFin) { ConvertTo-Number $f.profit_yoy -999 } else { -999 }
        $revenueYoy = if ($hasFin) { ConvertTo-Number $f.revenue_yoy -999 } else { -999 }
        $netProfit = if ($hasFin) { ConvertTo-Number $f.net_profit_yi } else { 0 }
        $grossMargin = if ($hasFin) { ConvertTo-Number $f.gross_margin } else { 0 }

        $earningsScoreText = '板块占位（缺财报）'
        if (-not ($a.score_earnings_override -eq '' -or $null -eq $a.score_earnings_override)) {
            $earningsScore = [double]$a.score_earnings_override
            $earningsScoreText = '人工覆盖'
        }
        elseif ($hasFin) {
            $pn = if ($profitYoy -gt 0) { Clamp-Value ($profitYoy / [double]$model.earnings_curve.profit_yoy_full_pct) } else { 0 }
            $rn = if ($revenueYoy -gt 0) { Clamp-Value ($revenueYoy / [double]$model.earnings_curve.revenue_yoy_full_pct) } else { 0 }
            $earningsScore = [double]$w.earnings * (0.6 * $pn + 0.4 * $rn)
            $earningsScoreText = '财报同比规则计算'
        }
        else {
            $earningsScore = [double]$w.earnings * 0.5
        }

        $certaintyScoreText = '板块占位（缺财报）'
        if (-not ($a.score_certainty_override -eq '' -or $null -eq $a.score_certainty_override)) {
            $certaintyScore = [double]$a.score_certainty_override
            $certaintyScoreText = '人工覆盖'
        }
        elseif ($hasFin) {
            $profitComp = Clamp-Value ($netProfit / [double]$model.certainty.profit_full_yi)
            $positive = if ($netProfit -gt 0) { 1 } else { 0 }
            $gmComp = Clamp-Value ($grossMargin / [double]$model.certainty.gross_margin_full_pct)
            $certaintyScore = [double]$w.certainty * (0.5 * $profitComp + 0.3 * $positive + 0.2 * $gmComp)
            $certaintyScoreText = '财报规模/盈利规则计算'
        }
        else {
            $certaintyScore = [double]$w.certainty * 0.5
        }

        $optionScore = ConvertTo-Number $a.score_option ([double]$w.option * 0.5)
        $aiScore = ConvertTo-Number $a.score_ai ([double]$w.ai * 0.5)
        $catalystScore = ConvertTo-Number $a.score_catalyst ([double]$w.catalyst * 0.5)

        $pe = if ($q -and (ConvertTo-Number $q.pe_ttm) -ne 0) { ConvertTo-Number $q.pe_ttm } else { 0 }
        $risk = 0.0
        if ($pe -gt [double]$model.risk.pe_high) { $risk += [double]$model.risk.pe_high_deduction }
        elseif ($pe -gt [double]$model.risk.pe_mid) { $risk += [double]$model.risk.pe_mid_deduction }
        if ($bear * 100 -lt [double]$model.odds_curve.tail_penalty_bear_pct) { $risk += [double]$model.risk.bear_tail_deduction }
        if ($base -lt 0) { $risk += [double]$model.risk.negative_base_deduction }
        if ($hasFin -and $netProfit -lt 0) { $risk += [double]$model.risk.loss_deduction }
        $risk = Clamp-Value $risk 0 ([double]$w.risk_max)
        if (-not ($a.risk_override -eq '' -or $null -eq $a.risk_override)) { $risk = [double]$a.risk_override }

        $convexity = 0
        if ($bull -ge 1.5 -and $base -ge 0.3) { $convexity = 2 }
        elseif ($bull -ge 1.0) { $convexity = 1 }

        $total = $returnScore + $oddsScore + $optionScore + $aiScore + $earningsScore + $catalystScore + $certaintyScore - $risk + $convexity
        if ($total -lt -20) { $total = -20 }
        if ($total -gt 102) { $total = 102 }

        $fwdProfit = 0.0
        $profitStatus = '旧Forward PE反推（暂估）'
        $p27 = ConvertTo-Number $a.profit_2027e -1
        $p28 = ConvertTo-Number $a.profit_2028e -1
        if ($p27 -gt 0 -and $p28 -gt 0) {
            $fwdProfit = 0.7 * $p27 + 0.3 * $p28
            $profitStatus = '人工2027E/2028E加权'
        }
        elseif ($oldFpe -gt 0) {
            $fwdProfit = $oldCap / $oldFpe
        }
        $forwardPe = if ($fwdProfit -gt 0) { $cap / $fwdProfit } else { 0 }

        $reasonablePe = ConvertTo-Number $a.reasonable_pe ([double]$model.reasonable_pe_default)
        $needProfit = if ($reasonablePe -gt 0) { $baseCap / $reasonablePe } else { 0 }
        $profitGap = $needProfit - $fwdProfit
        $gapRatio = if ($fwdProfit -gt 0) { $profitGap / $fwdProfit } else { 0 }
        $bridgeFlag = ''
        if ($base -gt ([double]$model.bridge_threshold_pct / 100)) {
            if ($gapRatio -le 0.5) { $bridgeFlag = '通过初检' }
            elseif ($gapRatio -le 1.0) { $bridgeFlag = '需验证' }
            else { $bridgeFlag = '需强验证' }
        }

        $conf = '低'
        $hasQuote = ($null -ne $q -and "$($q.price)" -ne '')
        if ($hasQuote -and $hasFin -and $p27 -gt 0 -and $p28 -gt 0) { $conf = '高' }
        elseif ($hasQuote -and $hasFin) { $conf = '中' }

        [void]$rows.Add([pscustomobject]@{
            code = $s.code; name = $s.name; group = $s.group; price_date = $TradeDate
            close_price = if ($q) { $q.price } else { '' }
            market_cap_yi = [math]::Round($cap, 2)
            market_cap_status = $capStatus
            pe_ttm = if ($pe -gt 0) { [math]::Round($pe, 1) } else { '' }
            forward_profit_yi = [math]::Round($fwdProfit, 2)
            profit_basis = $profitStatus
            forward_pe = if ($forwardPe -gt 0) { [math]::Round($forwardPe, 1) } else { '' }
            bear_return_pct = [math]::Round($bear * 100, 1)
            base_return_pct = [math]::Round($base * 100, 1)
            bull_return_pct = [math]::Round($bull * 100, 1)
            expected_return_pct = [math]::Round($exp * 100, 1)
            upside_downside_ratio = [math]::Round($ratio, 2)
            score_return = [math]::Round($returnScore, 2)
            score_odds = [math]::Round($oddsScore, 2)
            score_option = [math]::Round($optionScore, 2)
            score_ai = [math]::Round($aiScore, 2)
            score_earnings = [math]::Round($earningsScore, 2)
            score_catalyst = [math]::Round($catalystScore, 2)
            score_certainty = [math]::Round($certaintyScore, 2)
            risk_deduction = [math]::Round($risk, 2)
            convexity_bonus = $convexity
            final_score = [math]::Round($total, 1)
            confidence = $conf
            base_target_cap_yi = [math]::Round($baseCap, 2)
            required_profit_yi = [math]::Round($needProfit, 2)
            profit_gap_yi = [math]::Round($profitGap, 2)
            bridge_flag = $bridgeFlag
            revenue_yi = if ($hasFin) { $f.revenue_yi } else { '' }
            net_profit_yi = if ($hasFin) { $f.net_profit_yi } else { '' }
            revenue_yoy_pct = if ($hasFin) { $f.revenue_yoy } else { '' }
            profit_yoy_pct = if ($hasFin) { $f.profit_yoy } else { '' }
            gross_margin_pct = if ($hasFin) { $f.gross_margin } else { '' }
            report_date = if ($hasFin) { $f.report_date } else { '' }
            old_rank = [int]$a.old_rank
            old_score = [double]$a.old_score
            earnings_method = $earningsScoreText
            certainty_method = $certaintyScoreText
            notes = [string]$a.notes
        })
    }

    $sorted = @($rows | Sort-Object -Property @{ Expression = 'final_score'; Descending = $true }, @{ Expression = 'base_return_pct'; Descending = $true })
    for ($i = 0; $i -lt $sorted.Count; $i++) {
        $sorted[$i] | Add-Member -NotePropertyName rank -NotePropertyValue ($i + 1) -Force
        $sorted[$i] | Add-Member -NotePropertyName rank_change -NotePropertyValue ($sorted[$i].old_rank - ($i + 1)) -Force
    }
    return $sorted
}
