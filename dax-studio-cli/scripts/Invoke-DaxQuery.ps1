<#
.SYNOPSIS
  Runs a DAX query file against a live semantic model with DAX Studio's dscmd.exe (export or benchmark).

.EXAMPLE
  .\Invoke-DaxQuery.ps1 -QueryFile q.dax -Server "powerbi://api.powerbi.com/v1.0/myorg/My Workspace" -Database "My Model" `
      -UserId me@tenant.com -Mode csv -Out .\results\q_S1.csv -Params @{ CalendarYear = "2026"; Currency = "USD" }

.EXAMPLE
  .\Invoke-DaxQuery.ps1 -QueryFile q.dax -Server ... -Database ... -UserId ... -Mode bench -Out .\results\bench_q.csv -Cold 3 -Warm 3

.NOTES
  - @Name tokens in the query are replaced with DAX string literals ("value", quotes doubled), the way paginated
    reports bind parameters. A token with no value in -Params is an error (so a typo cannot silently run).
  - The resolved query is written next to the output (<out>.resolved.dax) so the exact run can be repeated.
  - Always non-interactive: if the account is not cached, it fails instead of prompting (see SKILL.md).
#>
param(
    [Parameter(Mandatory)] [string]$QueryFile,
    [Parameter(Mandatory)] [string]$Server,
    [string]$Database,
    [Parameter(Mandatory)] [string]$UserId,
    [ValidateSet('csv', 'bench')] [string]$Mode = 'csv',
    [Parameter(Mandatory)] [string]$Out,
    [hashtable]$Params = @{},
    [int]$Cold = 3,
    [int]$Warm = 3,
    [string]$Role,
    [string]$EffectiveUser,
    [string]$DsCmd
)

if (-not $DsCmd) {
    $DsCmd = @("C:\Program Files\DAX Studio\dscmd.exe", "$env:LOCALAPPDATA\Programs\DAX Studio\dscmd.exe") |
             Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $DsCmd) { throw "dscmd.exe not found; install DAX Studio 3.6.2+ or pass -DsCmd" }
}

$text = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $QueryFile).Path)
# Check every token first: a throw inside a regex MatchEvaluator script block does not stop the script.
$tokens = @([regex]::Matches($text, '@(\w+)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
$missing = @($tokens | Where-Object { -not $Params.ContainsKey($_) })
if ($missing.Count) { throw "No value supplied for parameter(s): $(($missing | ForEach-Object { '@' + $_ }) -join ', ')" }
$resolved = [regex]::Replace($text, '@(\w+)', {
    param($m)
    '"' + ([string]$Params[$m.Groups[1].Value]).Replace('"', '""') + '"'
})
$outFull = [IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory -Force (Split-Path $outFull) | Out-Null
$qOut = [IO.Path]::ChangeExtension($outFull, '.resolved.dax')   # never the source file
if ($qOut -eq (Resolve-Path -LiteralPath $QueryFile).Path) { throw "Resolved query path would overwrite the source query" }
[IO.File]::WriteAllText($qOut, $resolved, (New-Object Text.UTF8Encoding $false))

$conn = @('-s', $Server, '-u', $UserId, '--non-interactive', '-f', $qOut)
if ($Database) { $conn += @('-d', $Database) }

$sw = [Diagnostics.Stopwatch]::StartNew()
if ($Mode -eq 'csv') {
    $log = & $DsCmd csv $outFull @conn 2>&1
} else {
    $extra = @('--cold', $Cold, '--warm', $Warm, '--silent')
    if ($Role) { $extra += @('--role', $Role) }
    if ($EffectiveUser) { $extra += @('--effective-user', $EffectiveUser) }
    $log = & $DsCmd benchmark $outFull @conn @extra 2>&1
}
$code = $LASTEXITCODE
$sw.Stop()
$log | Where-Object { "$_" -match 'ERR|Error|rows returned|denied|sign-in|token' } | ForEach-Object { "  dscmd: $_" }
"exit=$code elapsed_ms=$($sw.ElapsedMilliseconds) out=$outFull"

if ($code -eq 0 -and $Mode -eq 'bench' -and (Test-Path $outFull)) {
    # Per-iteration rows come before the "# Summary" line. Exclude failed iterations (Error column) and trace-fallback
    # iterations (trace could not start: FE = SE = SE queries = 0, total is client wall-clock), and say how many.
    $lines = Get-Content $outFull
    $cut = [Array]::FindIndex([string[]]$lines, [Predicate[string]] { param($l) $l -like '#*' })
    if ($cut -lt 0) { $cut = $lines.Count }
    $iters = $lines[0..($cut - 1)] | ConvertFrom-Csv
    $errs = @($iters | Where-Object { $_.PSObject.Properties['Error'] -and $_.Error })
    $fallback = @($iters | Where-Object { -not ($_.PSObject.Properties['Error'] -and $_.Error) -and [int]$_.TotalDuration_ms -gt 0 -and [int]$_.StorageEngineQueryCount -eq 0 -and [int]$_.FormulaEngineDuration_ms -eq 0 -and [int]$_.StorageEngineDuration_ms -eq 0 })   # traced total ~ FE + SE, so total > 0 with FE = SE = 0 means no trace
    if ($errs.Count) { "ERRORS in $($errs.Count) iteration(s): $($errs[0].Error)" }
    if ($fallback.Count) { "$($fallback.Count) iteration(s) used wall-clock fallback (no trace) and are excluded below" }
    foreach ($c in 'Cold', 'Warm') {
        $ok = @($iters | Where-Object { $_.Cache -eq $c -and $errs -notcontains $_ -and $fallback -notcontains $_ })
        if ($ok.Count) {
            $t = $ok | ForEach-Object { [double]$_.TotalDuration_ms }
            $fe = ($ok | ForEach-Object { [double]$_.FormulaEngineDuration_ms } | Measure-Object -Average).Average
            $se = ($ok | ForEach-Object { [double]$_.StorageEngineDuration_ms } | Measure-Object -Average).Average
            "{0}: n={1} avg={2:N0} ms min={3:N0} ms FE={4:N0} SE={5:N0} SEqueries={6} rows={7}" -f $c, $ok.Count,
                ($t | Measure-Object -Average).Average, ($t | Measure-Object -Minimum).Minimum, $fe, $se, $ok[0].StorageEngineQueryCount, $ok[0].RowCount
        }
    }
}
