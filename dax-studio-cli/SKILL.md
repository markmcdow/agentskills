---
name: dax-studio-cli
description: Run DAX queries against a live Power BI / Fabric / Analysis Services semantic model from the command line with DAX Studio's dscmd.exe - export results to CSV/Parquet/XLSX, benchmark cold vs warm cache with formula-engine/storage-engine timings, test RLS roles, and capture VertiPaq (vpax) statistics. Use this whenever the work would be better informed by real data or real timings from a deployed model - validating a DAX or model change against the original results, performance tuning a query or paginated-report dataset, profiling data, checking row counts or relationships, reconciling numbers with a source system - and the Tabular Editor MCP (which has no timings and may only see an offline .bim) is not enough. Windows only.
---

# DAX Studio CLI (dscmd.exe) against a live semantic model

`dscmd.exe` ships with DAX Studio. It runs a DAX query (or a model-level operation) against any XMLA-reachable
tabular model and writes the result to a file. Because it is non-interactive and scriptable, it lets you:

- **check results**: export a query result and compare it with another query or with a source system;
- **measure performance**: `benchmark` clears the cache, runs cold and warm iterations, and reports total, formula engine (FE) and storage engine (SE) ms, plus SE query count;
- **test security**: run with an RLS role or an impersonated user (`benchmark --role / --effective-user`);
- **inspect the model**: `vpax` gives row counts, column cardinality and size; `export` dumps tables.

Prefer it over the Tabular Editor MCP's `execute_dax_query` when you need timings, large result sets saved to
disk, repeatable runs, or the **deployed** model (the MCP may have an offline `Model.bim` open, which has no data).

## Prerequisites (check these first)

| Requirement | Detail | How to check |
|---|---|---|
| OS | Windows (dscmd is a .NET Windows app) | — |
| DAX Studio | **3.6.0 or later** (`benchmark` was introduced in 3.6.0). This skill was written and tested against `dscmd v3.6.2`. Older builds can still export (`csv`/`file`/`xlsx`/`vpax`) but lack `benchmark`, and may lack `auth`, `--non-interactive` and `--role`/`--effective-user`. Portable builds ship `dscmd.exe` too. Free download: https://daxstudio.org | `& "C:\Program Files\DAX Studio\dscmd.exe" --help` (first line shows the version) |
| Location | Default `C:\Program Files\DAX Studio\dscmd.exe`; per-user installs go under `%LOCALAPPDATA%\Programs\DAX Studio` | `Get-ChildItem "C:\Program Files\DAX Studio*", "$env:LOCALAPPDATA\Programs\DAX Studio*"` |
| Power BI / Fabric model | Workspace on **Premium, PPU or Fabric capacity**, with the **XMLA endpoint** at least *Read* (capacity setting) and tenant setting *Allow XMLA endpoints and Analyze in Excel with on-premises datasets* enabled. Pro/shared-capacity workspaces cannot be queried this way. | Connecting fails with an XMLA/capacity error if not met |
| Permissions | Read queries: **Build** permission on the semantic model, or workspace Viewer with build, or Contributor and above. Cold-cache benchmarks clear the cache, which typically needs **write** access (Contributor, Member or Admin). Without it, "cold" numbers may not be cold. | A permission error on connect or on cache clear |
| Authentication | Microsoft Entra ID account **cached by DAX Studio** (see below). For Azure Analysis Services the same applies; local SSAS/Power BI Desktop use Windows auth. | `dscmd auth --list` |
| Network | HTTPS to `api.powerbi.com` / `*.analysis.windows.net` (or the AAS region host) | — |

### Connection strings (`-s` server, `-d` database)

| Target | `-s` | `-d` |
|---|---|---|
| Power BI / Fabric workspace | `powerbi://api.powerbi.com/v1.0/myorg/<Workspace Name>` (spaces allowed; quote it) | semantic model name |
| Azure Analysis Services | `asazure://<region>.asazure.windows.net/<server>` | database name |
| SSAS on-premises | `<host>\<instance>` | database name |
| Power BI Desktop (open file) | `<filename>.pbix` / `.pbip` (name only, no path) or `localhost:<port>`. **Desktop must already be running with the file open and the model loaded**; dscmd does not start it. | optional |
| Power Pivot (Excel) | **not supported**; use the DAX Studio Excel add-in instead | — |

The workspace connection string is under Workspace settings → License info → Connection link.
`-d` is required for every target except Power BI Desktop.
`-c "<full connection string>"` can replace `-s` and `-d`, but cannot be combined with them.
Put the `User ID` in the connection string, or keep using `-u` / `DSCMD_USER`.
Official docs: https://daxstudio.org/docs/features/command-line/ (Commands, Authentication, Connecting, Syntax pages).

### Authentication: unattended runs

Claude cannot complete an interactive sign-in. To get a silent token, the account must already be in the DAX Studio token cache.

How dscmd picks an Entra account:
- No cached account: it prompts.
- Exactly one cached account: it tries that account silently.
- Several cached accounts: it prompts unless `-u` names one.
- `-u` doesn't match a cached account: it prompts for that account.
- MFA or Conditional Access can force a prompt even for a cached account.
- With `--non-interactive`, or when the process has no window (for example a Windows service), every one of these prompts becomes a failure with a non-zero exit code.

Local SSAS and Power BI Desktop use the Windows identity of the calling process, so no Entra sign-in is involved.

1. **List cached accounts:** `dscmd auth --list` (no token is acquired). Many consultants have one account per client tenant.
2. **If the account isn't cached**, ask the user to sign in once interactively. They can type it in the Claude Code prompt:
   `! "C:\Program Files\DAX Studio\dscmd.exe" auth -u <user@tenant> -s "powerbi://api.powerbi.com/v1.0/myorg/<Workspace>"`
   or open DAX Studio and connect to the model once.
3. **Always pass `-u <account>` together with `--non-interactive`.** On a machine with several cached accounts, `-u` picks the right tenant. `--non-interactive` makes the command fail with a clear error instead of hanging on a hidden prompt.
4. **Pre-flight check** for a long batch: `dscmd auth -u <account> -s <server> --non-interactive`. It authenticates without exposing a token; a non-zero exit code means you need to re-authenticate. `dscmd accesstoken ... --non-interactive` also works, but it prints the token. **Never print or log a token:** redirect the output to `$null` and check only `$LASTEXITCODE`.
5. **Environment variables** can replace the flags. A command-line option overrides its variable.
   - `DSCMD_USER`: the account to use.
   - `DSCMD_NON_INTERACTIVE`: set to `1`, `true` or `yes`.
   - `DSCMD_PASSWORD`: the password or secret.

**Other authentication modes** (only when the user sets them up; never invent or ask for secrets):
- **Service principal:** the user ID is `-u "app:<clientId>@<tenantId>"`. Pass the client secret with `-p`, or `-p "cert:<thumbprint>"` for a certificate. Keep the secret in `DSCMD_PASSWORD`, set by the user, rather than in scripts. The service principal must be allowed to use Power BI APIs in tenant settings, and must be a member of the workspace or have Build permission on the model.
- **Access token reuse:** `dscmd accesstoken ...` prints an Entra token, usually valid for 60–90 minutes. Pass it later as `-p <token>` **without** `-u`. It is useful for many runs after one interactive sign-in, but treat it as a password: never echo or store it, and avoid it on shared machines. Cached accounts with `-u --non-interactive` are usually simpler.

**Do not put user passwords in `-p`.** Don't ask the user for passwords, and never write secrets into scripts or files.

## Core commands

```powershell
$dscmd = "C:\Program Files\DAX Studio\dscmd.exe"
$conn  = @('-s', 'powerbi://api.powerbi.com/v1.0/myorg/<Workspace>', '-d', '<Model>', '-u', '<user@tenant>', '--non-interactive')

& $dscmd csv  out.csv      @conn -f query.dax          # query result -> CSV
& $dscmd file out.json     @conn -f query.dax          # type from extension: .csv UTF-8 CSV, .txt tab-delimited, .json JSON
& $dscmd file out.parquet  @conn -f query.dax -t PARQUET   # -t UTF8CSV|UNICODECSV|JSON|TAB|PARQUET overrides the extension
& $dscmd csv  out.csv      @conn -f query.dax -m Year=2026 -m Currency=USD   # real query parameters, referenced as @Year in the query
& $dscmd xlsx out.xlsx     @conn -f query.dax
& $dscmd benchmark bench.csv @conn -f query.dax --cold 3 --warm 3 --silent
& $dscmd benchmark bench.csv @conn -f query.dax --cold 1 --warm 1 --role "GL - READ"            # RLS role
& $dscmd benchmark bench.csv @conn -f query.dax --cold 1 --warm 1 --effective-user someone@x.com # impersonation
& $dscmd vpax model.vpax   @conn                       # model stats (VertiPaq Analyzer)
& $dscmd export csv C:\out\folder @conn -t "Table A" "Table B"   # dump tables
```

Helper scripts in this skill's `scripts/` folder:

- `Invoke-DaxQuery.ps1`: runs a query file in `csv` or `bench` mode. It substitutes `@Param` tokens (as paginated reports do), writes the resolved query next to the output so it can be rerun, and prints the benchmark summary.
- `Compare-DaxCsv.ps1`: compares two exported results as multisets. Rows are matched on key columns, and numeric columns are compared within a tolerance. Use it to prove a rewritten query or a model change returns the same data.

## Practices that matter (learned the hard way)

**Queries**
- **Put queries in files and use `-f`** (`.dax`, `.daxx` or plain text), not `-q`. Inline queries break on PowerShell/cmd quoting (for example the error "ROW expects a column name"). The official PowerShell example has to escape quotes inside a here-string, which a file avoids.
- **`@Param` queries (paginated reports, parameterised tests)** have two options:
  - **`-m Name=Value`** (repeat it per parameter). It sends real query parameters, as the report does. It is available on `csv`, `file`, `xlsx` and `accesstoken`, **but not on `benchmark`**.
  - **Textual substitution** of DAX string literals, as `Invoke-DaxQuery.ps1` does. It works for every command, and it gives `csv` and `benchmark` exactly the same query text, which matters when results and timings must correspond.
  - **Commented-out tokens count.** The report still binds a parameter whose only use is in a `//` line, and textual substitution replaces it there too. Supply every `@Name` that appears anywhere in the text; `Invoke-DaxQuery.ps1` stops and lists any that are missing.
  - **Either way, use the exact values the report sends.** For multi-select parameters that is often a pipe-joined list rather than "All". Keep one JSON file of named scenarios (for example "all selected, USD, latest period") so runs are repeatable.
- **Run one query at a time.** These are production capacities, and heavy queries consume capacity units (CU) and can hit the memory limit (for example the 10 GB error on wide `SUMMARIZECOLUMNS` with `IF`-wrapped measures).
- **Long queries:** run them in the background (`run_in_background`) with a generous timeout, and expect cold runs of minutes on large models.

**Reading CSV output**
- **Header names:** columns lose their table prefix (`'Account'[Object Account]` becomes `Object Account`), and query columns keep their alias.
- **Case-insensitive duplicates:** if two columns differ only by case (for example `SEQ` and `Seq`), PowerShell's `Import-Csv` fails. Rename one of them in the query.
- **Formatting:**
  - BLANK is an empty field.
  - Strings are quoted.
  - Dates come out as `yyyy-MM-dd HH:mm:ss.fff`.
  - Fixed-decimal (Currency) columns print as `0.00`, but Double columns print `0`. Compare numbers, not text.
- **Equality across a model change:** use a numeric tolerance (for example 0.0001) and report BLANK-vs-0 separately. Byte-identical text comparison only works when both queries run on the same model with the same data types.

**Benchmarking**
- **How it runs:** each **cold** iteration sends the XMLA `ClearCache` command before running the query; **warm** iterations don't. FE and SE times come from a server-timings trace, the same as in the DAX Studio UI. `--cold` and `--warm` both default to 5; set at least one above 0.
- **Output layout:**
  - A per-iteration table: `Sequence, Cache, TotalDuration_ms, FormulaEngineDuration_ms, StorageEngineDuration_ms, StorageEngineQueryCount, StorageEngineCpu_ms, TotalCpu_ms, VertipaqCacheMatches, RowCount, Error`.
  - Then a `# Summary` block with Average, StdDev, Min and Max per cache state. It leaves out the CPU and `Error` columns.
- **Check the `Error` column.** A failed iteration still writes a row.
- **Trace fallback:** if the trace can't start (permissions or endpoint limits), dscmd falls back to wall-clock timing. FE, SE and SE-query counts are then 0, and the total includes client and network time. Recognise it by **total > 0 while FE = SE = 0**. A traced trivial query can legitimately be all zeros, total included. These rows are not comparable with traced ones, so exclude them from averages (the helper script does). If every row is a fallback, report the numbers as wall-clock only. Report the minimum alongside the average.
- **Reading the split:**
  - High SE time when cold: a scan problem (data volume, relationships, snowflakes).
  - High FE time when warm: a calculation-shape problem (row-by-row iteration, `CALCULATE` per cell, callbacks).
  - A high SE query count: per-cell evaluation that should be set-based.
- **Fair comparisons:** benchmark the original and the candidate under the same scenario, close together in time (capacity load varies). Re-export the original before comparing if the model has been refreshed since.

**Validating a change**
1. Export the original query's result for each scenario.
2. Export the candidate's result.
3. Run `Compare-DaxCsv.ps1`, which must report MATCH.
4. Only then benchmark.
5. Log the commands, row counts, results and timings in the project's testing notes.

**Privacy and safety**
- **Results contain client data.** Keep them in a git-ignored folder (`results/*.csv`), and don't paste large extracts into chat or send them to external services.
- **Read only:** dscmd here is for queries. Don't use `export sql` to push data anywhere unless the user explicitly asks.
- **Name the target:** say which workspace and model you are querying, and which account, before running against production.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| Error mentioning interactive sign-in / no cached account under `--non-interactive` | The account isn't cached or the token expired. Ask the user to run `dscmd auth -u <account>` interactively or connect once in DAX Studio. |
| Signs in to the wrong tenant / "database not found" | Wrong cached account picked; pass `-u` explicitly. Check the `-d` spelling (it is the semantic model name). |
| XMLA / capacity error on connect | Workspace not on Premium/PPU/Fabric, or XMLA endpoint disabled. This is an admin setting. |
| `Import-Csv` "member already present" | Two column names differ only by case. Rename one in the query. |
| Every row differs in a text comparison | Formatting differences (Currency vs Double, quoting). Use `Compare-DaxCsv.ps1`, and if rows still don't match, compare one row side by side to find the real difference. |
| Memory error (for example 10 GB) | The query materializes too much. Simplify the measures in `SUMMARIZECOLUMNS` (bare measure references), pre-aggregate, or filter. |
| Cold timings look warm | `ClearCache` not permitted for the account. Use an account with write access, or treat the numbers as warm. |
| All benchmark FE/SE columns are 0 | The trace couldn't start, so the times are wall-clock (see Benchmarking). Check permissions, or accept wall-clock numbers. |
| Prompts or fails although the account is cached | MFA or Conditional Access requires interaction, or several accounts are cached and `-u` is missing. |
| `-c` rejected together with `-s` or `-d` | Use either a connection string or server plus database, not both. |
| "You must specify a <database> when using the <server> parameter" | `-d` is required with `-s` (except for a .pbix/.pbip target). |
