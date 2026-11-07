# Helpers for company profile / revenue mix / announcement data.
$ErrorActionPreference = 'Stop'

function New-TextBar {
    param([double]$Pct, [int]$Width = 20)
    $p = [math]::Max(0, [math]::Min(100, $Pct))
    $fill = [int][math]::Round($p / 100 * $Width)
    return ('█' * $fill) + ('░' * ($Width - $fill))
}

function Get-SentenceHighlights {
    param(
        [string]$Text,
        [string]$Pattern,
        [int]$Max = 6,
        [int]$MaxLen = 220
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $clean = $Text -replace '\s+', ' '
    $parts = $clean -split '[。；;]'
    $hits = New-Object System.Collections.ArrayList
    foreach ($part in $parts) {
        $t = $part.Trim()
        if ($t.Length -lt 8) { continue }
        if ($t -match $Pattern) {
            if ($t.Length -gt $MaxLen) { $t = $t.Substring(0, $MaxLen) + '…' }
            if (-not $hits.Contains($t)) { [void]$hits.Add($t) }
            if ($hits.Count -ge $Max) { break }
        }
    }
    return @($hits)
}

function Get-LatestByDate {
    param([array]$Rows, [string]$Field = 'REPORT_DATE')
    $rowsWithField = @($Rows | Where-Object { $_.$Field })
    if ($rowsWithField.Count -eq 0) { return $null }
    return ($rowsWithField | Sort-Object -Property @{ Expression = { [string]$_.$Field }; Descending = $true } | Select-Object -First 1)
}
