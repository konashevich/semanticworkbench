<#
Sync Office dev -> prod
Usage:
  .\sync-office-dev-to-prod.ps1 [-WhatIf] [-CreateBackup] [-ExcludeList <string[]>]

Options:
  -WhatIf (switch): perform a dry run using Robocopy's /L to list actions.
  -CreateBackup (switch): before mirroring, copy the current prod folder to prod_backup_<timestamp>.
  -ExcludeList: array of relative paths to exclude (default excludes .venv, .venv_shared, .git, __pycache__).

Notes:
- Script uses Robocopy and preserves timestamps and attributes. It's Windows-only.
- It excludes common venv folders and build artifacts by default.
#>
param(
    [switch]$WhatIf,
    [switch]$CreateBackup,
    [string[]]$ExcludeList = @('.venv','.venv_shared','.venv*','venv','.git','dist','build','__pycache__','.pytest_cache')
)

$root = Join-Path $PSScriptRoot '..' | Resolve-Path
$dev = Join-Path $root 'mcp-servers\mcp-server-office'
$prod = Join-Path $root 'mcp-servers\mcp-server-office-prod'

if (-not (Test-Path $dev)) { Write-Error "Dev folder not found: $dev"; exit 1 }
if (-not (Test-Path $prod)) { Write-Error "Prod folder not found: $prod"; exit 1 }

if ($CreateBackup) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backup = "$prod_backup_$timestamp"
    Write-Output "Creating backup of prod -> $backup"
    robocopy $prod $backup /MIR /FFT /Z /XA:SH /W:1 /R:1 | Out-Null
}

# Build robocopy exclude switches
$xdArgs = @()
foreach ($ex in $ExcludeList) { $xdArgs += '/XD'; $xdArgs += $ex }

# Use /MIR to mirror. If WhatIf, use /L to list but not copy
$robocopyArgs = @($dev, $prod, '/MIR', '/FFT', '/Z', '/XA:SH', '/W:1', '/R:1') + $xdArgs
if ($WhatIf) { $robocopyArgs += '/L' }

Write-Output "Running Robocopy from:`n  $dev`nTo:`n  $prod
Excluding: $($ExcludeList -join ', ')"

$rc = robocopy @robocopyArgs
$exit = $LASTEXITCODE

if ($WhatIf) {
    Write-Output "Dry-run finished. Robocopy exit code: $exit"
    exit 0
}

# Robocopy exit codes: 0 no files, 1 files copied, >=8 failures
if ($exit -ge 8) {
    Write-Error "Robocopy failed with exit code $exit"
    exit $exit
}

Write-Output "Sync complete. Robocopy exit code: $exit"
