# ============================================================
#  One-time setup for the Excel Refresh Orchestrator.
#    - Installs uv (a fast Python tool) if it is missing.
#    - Creates the project's private environment (.venv) and installs
#      its dependency into it.
#
#  About Python: uv uses a Python you ALREADY have (3.14 or newer) if you
#  have one, and only downloads an isolated copy if you don't.
#
#  Run from the project folder, in PowerShell:
#      powershell -ExecutionPolicy Bypass -File setup.ps1
# ============================================================

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

Write-Host "[1/2] Checking for uv..." -ForegroundColor Cyan
$uv = Get-Command uv -ErrorAction SilentlyContinue
if (-not $uv) {
    Write-Host "      Installing uv..." -ForegroundColor Yellow
    Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    # The installer updates PATH for new shells; make uv usable in THIS one too.
    $env:Path = "$env:USERPROFILE\.local\bin;$env:Path"
}
else {
    Write-Host "      Found: $($uv.Source)" -ForegroundColor Green
}

Write-Host "[2/2] Creating the environment and installing dependencies..." -ForegroundColor Cyan
uv sync

Write-Host ""
Write-Host "Setup complete." -ForegroundColor Green
Write-Host "Next: edit config.toml, then test with:  uv run python refresh_reports.py"
