param(
    [switch]$NoBrowser
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ScriptDir
Set-Location $RepoRoot

function Add-UvToPath {
    $candidates = @(
        Join-Path $env:USERPROFILE ".local\bin",
        Join-Path $env:USERPROFILE ".cargo\bin",
        Join-Path $env:LOCALAPPDATA "Programs\Python\Python313\Scripts",
        Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\Scripts",
        Join-Path $env:LOCALAPPDATA "Programs\Python\Python311\Scripts",
        Join-Path $env:LOCALAPPDATA "Programs\Python\Python310\Scripts"
    )

    foreach ($candidate in $candidates) {
        if ((Test-Path (Join-Path $candidate "uv.exe")) -and
            ($env:Path -notlike "*$candidate*")) {
            $env:Path = "$candidate;$env:Path"
        }
    }
}

Write-Host "OriginLab MCP setup"
Write-Host "Repository: $RepoRoot"

Add-UvToPath
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Host "uv not found. Installing uv..."
    Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    Add-UvToPath
}

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    throw "uv was installed but is not available on PATH. Restart PowerShell and run this script again."
}

Write-Host "Installing project dependencies..."
uv sync

if ($NoBrowser) {
    $env:ORIGINLAB_MCP_UI_NO_BROWSER = "1"
} elseif (Test-Path Env:ORIGINLAB_MCP_UI_NO_BROWSER) {
    Remove-Item Env:ORIGINLAB_MCP_UI_NO_BROWSER
}

Write-Host "Opening OriginLab MCP UI at http://127.0.0.1:8765/"
uv run python -m originlab_mcp.ui
