<#
.SYNOPSIS
  Compares two dscmd CSV exports as multisets of rows: rows matched on key columns, numeric columns within a tolerance.

.DESCRIPTION
  Value columns (compared numerically) are those matching -ValuePattern; if omitted, every column whose non-empty
  values in the expected file all parse as numbers. All other columns (minus -IgnoreColumns) form the row key.
  Text formatting differences (quoting, 0 vs 0.00) do not matter. BLANK vs 0 is counted separately and not a failure.
  Prints MATCH / NOT A MATCH, counts, the largest difference, and samples of mismatches.

.EXAMPLE
  .\Compare-DaxCsv.ps1 -Expected orig.csv -Actual new.csv -ValuePattern 'Amt' -Tol 0.0001
#>
param(
    [Parameter(Mandatory)] [string]$Expected,
    [Parameter(Mandatory)] [string]$Actual,
    [double]$Tol = 0.0001,
    [string]$ValuePattern,
    [string[]]$IgnoreColumns = @(),
    [int]$Show = 10
)
$inv = [Globalization.CultureInfo]::InvariantCulture
foreach ($p in @($Expected, $Actual)) { if (-not (Test-Path -LiteralPath $p)) { "MISSING FILE: $p"; "NOT A MATCH"; return } }
$e = @(Import-Csv -LiteralPath $Expected); $a = @(Import-Csv -LiteralPath $Actual)
if ($e.Count -eq 0 -or $a.Count -eq 0) { "rows expected=$($e.Count) actual=$($a.Count)"; if ($e.Count -eq $a.Count) { "MATCH (both empty)" } else { "NOT A MATCH" }; return }
$cols = @($e[0].PSObject.Properties.Name); $colsA = @($a[0].PSObject.Properties.Name)
if (($cols -join '|') -ne ($colsA -join '|')) { "HEADER MISMATCH"; "  expected: $($cols -join ',')"; "  actual:   $($colsA -join ',')"; "NOT A MATCH"; return }

function IsNum([string]$s) { $d = 0.0; [double]::TryParse($s, [Globalization.NumberStyles]::Float, $inv, [ref]$d) }

# dscmd quotes text columns and leaves numbers/booleans/dates bare. Use the RAW quoting of the expected file to tell
# text from numbers (a code like "6730" is text even though it parses as a number).
function QuotedFlags([string]$line) {
    $flags = New-Object Collections.Generic.List[bool]; $i = 0; $n = $line.Length
    while ($true) {
        if ($i -lt $n -and $line[$i] -eq '"') {
            $i++; while ($i -lt $n) { if ($line[$i] -eq '"') { if ($i + 1 -lt $n -and $line[$i + 1] -eq '"') { $i += 2; continue }; break }; $i++ }
            $i++; $flags.Add($true)
        } else {
            while ($i -lt $n -and $line[$i] -ne ',') { $i++ }; $flags.Add($false)
        }
        if ($i -ge $n) { break }; $i++   # skip comma
        if ($i -eq $n) { $flags.Add($false); break }
    }
    , $flags
}
if ($ValuePattern) { $val = @($cols | Where-Object { $_ -match $ValuePattern }) }
else {
    $quoted = @{}; $raw = [IO.File]::ReadLines((Resolve-Path -LiteralPath $Expected).Path) | Select-Object -Skip 1 -First 200
    foreach ($l in $raw) { $f = QuotedFlags $l; if ($f.Count -ne $cols.Count) { continue }; for ($j = 0; $j -lt $cols.Count; $j++) { if ($f[$j]) { $quoted[$cols[$j]] = $true } } }
    # value column = never quoted and every non-empty value numeric (excludes True/False and dates)
    $val = @($cols | Where-Object { $c = $_; -not $quoted.ContainsKey($c) -and
        -not ($e | Where-Object { $v = [string]$_.$c; $v -ne '' -and -not (IsNum $v) } | Select-Object -First 1) })
}
$val = @($val | Where-Object { $IgnoreColumns -notcontains $_ })
$key = @($cols | Where-Object { $val -notcontains $_ -and $IgnoreColumns -notcontains $_ })
"key columns: $($key.Count)  value columns: $($val.Count)  ignored: $($IgnoreColumns.Count)"

$sep = [string][char]31
function K($r) { ($key | ForEach-Object { [string]$r.$_ }) -join $sep }
$idx = @{}
foreach ($r in $a) { $k = K $r; if (-not $idx.ContainsKey($k)) { $idx[$k] = New-Object Collections.Generic.Queue[object] }; $idx[$k].Enqueue($r) }

$onlyE = New-Object Collections.Generic.List[string]; $matched = 0; $bad = 0; $blank0 = 0; $maxAbs = 0.0
$samples = New-Object Collections.Generic.List[string]
foreach ($r in $e) {
    $k = K $r
    if (-not $idx.ContainsKey($k) -or $idx[$k].Count -eq 0) { $onlyE.Add($k); continue }
    $o = $idx[$k].Dequeue(); $matched++
    foreach ($c in $val) {
        $x = [string]$r.$c; $y = [string]$o.$c
        if ($x -eq $y) { continue }
        if ($x -eq '' -or $y -eq '') {
            $other = if ($x -eq '') { $y } else { $x }
            if ((IsNum $other) -and [double]::Parse($other, $inv) -eq 0) { $blank0++; continue }
            $bad++; if ($samples.Count -lt $Show) { $samples.Add("  [$c] expected '$x' actual '$y' | " + $k.Replace($sep, ' | ')) }; continue
        }
        if (-not ((IsNum $x) -and (IsNum $y))) { $bad++; if ($samples.Count -lt $Show) { $samples.Add("  [$c] expected '$x' actual '$y' | " + $k.Replace($sep, ' | ')) }; continue }
        $d = [Math]::Abs([double]::Parse($x, $inv) - [double]::Parse($y, $inv))
        if ($d -gt $maxAbs) { $maxAbs = $d }
        if ($d -gt $Tol) { $bad++; if ($samples.Count -lt $Show) { $samples.Add("  [$c] expected $x actual $y (diff $d) | " + $k.Replace($sep, ' | ')) } }
    }
}
$onlyA = 0; foreach ($q in $idx.Values) { $onlyA += $q.Count }
"rows expected=$($e.Count) actual=$($a.Count) matched=$matched only_expected=$($onlyE.Count) only_actual=$onlyA"
"value cells beyond tolerance=$bad  blank_vs_zero=$blank0  max_abs_diff=$maxAbs  (tol=$Tol)"
if ($onlyE.Count) { "sample rows only in expected:"; $onlyE | Select-Object -First $Show | ForEach-Object { "  " + $_.Replace($sep, ' | ') } }
if ($samples.Count) { "sample value differences:"; $samples }
if ($onlyE.Count -eq 0 -and $onlyA -eq 0 -and $bad -eq 0) { "MATCH" } else { "NOT A MATCH" }
