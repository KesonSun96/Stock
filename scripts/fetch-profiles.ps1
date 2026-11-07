param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\src\lib.ps1')
. (Join-Path $PSScriptRoot '..\src\profiles.ps1')

$outDir = Join-ProjectPath 'data/profiles'
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
$profileFile = Join-Path $outDir 'company_profiles.csv'
$mixFile = Join-Path $outDir 'revenue_mix.csv'
$annFile = Join-Path $outDir 'announcements.csv'

if ((Test-Path -LiteralPath $profileFile) -and -not $Force) {
    Write-Host 'Profile data already exists. Use -Force to refresh.'
    exit 0
}

$pool = Read-CsvData -Path (Join-ProjectPath 'config/stock_pool.csv')
$profileRows = New-Object System.Collections.ArrayList
$mixRows = New-Object System.Collections.ArrayList
$annRows = New-Object System.Collections.ArrayList
$headers = @{
    'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'
    'Referer'    = 'https://emweb.securities.eastmoney.com/'
    'Accept'     = 'application/json,text/plain,*/*'
}

$i = 0
foreach ($p in $pool) {
    $i++
    Write-Progress -Activity 'Fetching company profiles' -Status $p.name -PercentComplete ([int]($i * 100 / $pool.Count))
    $emCode = if ($p.exchange -eq 'SH') { "SH$($p.code)" } else { "SZ$($p.code)" }

    $row = [pscustomobject]@{
        code = $p.code; name = $p.name; group = $p.group
        org_name = ''; industry_em = ''; industry_csrc = ''; province = ''
        employees = ''; website = ''; reg_capital_yi = ''; chairman = ''
        profile = ''; business_scope = ''; review_date = ''; review_text = ''
        data_status = '未获取'
    }

    try {
        $survey = Invoke-ApiJson -Uri ("https://emweb.securities.eastmoney.com/PC_HSF10/CompanySurvey/PageAjax?code=$emCode") -Retry 2
        $jb = @($survey.jbzl)[0]
        if ($jb) {
            $row.org_name = [string]$jb.ORG_NAME
            $row.industry_em = [string]$jb.EM2016
            $row.industry_csrc = [string]$jb.INDUSTRYCSRC1
            $row.province = [string]$jb.PROVINCE
            $row.employees = if ($jb.EMP_NUM) { [string]$jb.EMP_NUM } else { '' }
            $row.website = [string]$jb.ORG_WEB
            $row.reg_capital_yi = if ($jb.REG_CAPITAL) { [math]::Round((ConvertTo-Number $jb.REG_CAPITAL) / 10000, 2) } else { '' }
            $row.chairman = [string]$jb.CHAIRMAN
            $row.profile = ([string]$jb.ORG_PROFILE).Trim()
            $row.business_scope = ([string]$jb.BUSINESS_SCOPE).Trim()
            $row.data_status = '已获取公司概况'
        }
    }
    catch { }

    try {
        $ba = Invoke-ApiJson -Uri ("https://emweb.securities.eastmoney.com/PC_HSF10/BusinessAnalysis/PageAjax?code=$emCode") -Retry 2
        $review = Get-LatestByDate -Rows @($ba.jyps)
        if ($review) {
            $row.review_date = ([string]$review.REPORT_DATE).Substring(0, 10)
            $row.review_text = ([string]$review.BUSINESS_REVIEW).Trim()
        }

        $zy = @($ba.zygcfx)
        if ($zy.Count -gt 0) {
            $productDates = @($zy | Where-Object { [string]$_.MAINOP_TYPE -eq '1' -and $_.REPORT_DATE } | ForEach-Object { [string]$_.REPORT_DATE } | Sort-Object -Descending -Unique)
            $latestDate = if ($productDates.Count -gt 0) { $productDates[0] } else { (@($zy | ForEach-Object { [string]$_.REPORT_DATE } | Sort-Object -Descending -Unique))[0] }
            if ($latestDate) {
            foreach ($z in @($zy | Where-Object { [string]$_.REPORT_DATE -eq $latestDate })) {
                $type = [string]$z.MAINOP_TYPE
                $dimLabel = switch ($type) { '1' { '按产品' } '2' { '按行业' } '3' { '按地区' } default { "类型$type" } }
                [void]$mixRows.Add([pscustomobject]@{
                    code = $p.code; name = $p.name
                    report_date = $latestDate.Substring(0, 10)
                    dimension = $dimLabel
                    item = [string]$z.ITEM_NAME
                    revenue_yi = [math]::Round((ConvertTo-Number $z.MAIN_BUSINESS_INCOME) / 100000000, 2)
                    cost_yi = [math]::Round((ConvertTo-Number $z.MAIN_BUSINESS_COST) / 100000000, 2)
                    gross_margin_pct = [math]::Round((ConvertTo-Number $z.GROSS_RPOFIT_RATIO) * 100, 1)
                    share_pct = [math]::Round((ConvertTo-Number $z.MBI_RATIO) * 100, 1)
                })
            }
            }
        }
    }
    catch { }

    try {
        $annUri = "https://np-anotice-stock.eastmoney.com/api/security/ann?sr=-1&page_size=40&page_index=1&ann_type=A&client_source=web&stock_list=$($p.code)&f_node=0&s_node=0"
        $ann = Invoke-ApiJsonUtf8 -Uri $annUri -Retry 2
        foreach ($a in @($ann.data.list)) {
            $date = if ($a.notice_date) { ([string]$a.notice_date).Substring(0, 10) } else { '' }
            $type = ''
            if (@($a.columns).Count -gt 0) { $type = [string]@($a.columns)[0].column_name }
            [void]$annRows.Add([pscustomobject]@{
                code = $p.code; name = $p.name; date = $date
                title = ([string]$a.title).Trim(); type = $type
                url = "https://data.eastmoney.com/notices/detail/$($p.code)/$($a.art_code).html"
            })
        }
    }
    catch { }

    [void]$profileRows.Add($row)
    Start-Sleep -Milliseconds 120
}
Write-Progress -Activity 'Fetching company profiles' -Completed

Write-CsvUtf8 -Path $profileFile -Data $profileRows
Write-CsvUtf8 -Path $mixFile -Data $mixRows
Write-CsvUtf8 -Path $annFile -Data $annRows
Write-Host ("profiles={0}, mix_rows={1}, announcements={2}" -f $profileRows.Count, $mixRows.Count, $annRows.Count)
Write-Host ("saved to {0}" -f $outDir)
