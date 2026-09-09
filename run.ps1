param(
    [ValidateSet('build', 'refresh', 'value')][string]$Command = 'build',
    [string]$FromWorkbook,
    [string]$Config,
    [switch]$NoPreviews
)
$ErrorActionPreference = 'Stop'
$dcfPython = Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
if (-not (Test-Path -LiteralPath $dcfPython)) { $dcfPython = 'python' }
$dcfArguments = @('-m', 'dcf_model.cli', $Command)
if ($FromWorkbook) { $dcfArguments += @('--from-workbook', (Resolve-Path -LiteralPath $FromWorkbook).Path) }
if ($Config) { $dcfArguments += @('--config', (Resolve-Path -LiteralPath $Config).Path) }
if ($NoPreviews) { $dcfArguments += '--no-previews' }
Push-Location (Split-Path -Parent $PSScriptRoot)
try { & $dcfPython @dcfArguments; if ($LASTEXITCODE -ne 0) { throw "DCF command failed with exit code $LASTEXITCODE" } }
finally { Pop-Location }
