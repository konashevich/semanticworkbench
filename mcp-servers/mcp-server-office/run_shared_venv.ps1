# Launch mcp-server-office using the shared virtual environment
$root = Split-Path -Parent $PSCommandPath
$venv = Resolve-Path "$root\..\..\.venv_shared" -ErrorAction SilentlyContinue
if (-not $venv) {
    Write-Error "Shared venv not found at ../../.venv_shared. Create it first."
    exit 1
}
$python = Join-Path $venv "Scripts\python.exe"
& $python -m mcp_server.start @Args
