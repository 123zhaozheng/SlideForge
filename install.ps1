$ErrorActionPreference = "Stop"
$SkillDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "[SlideForge] Checking environment..." -ForegroundColor Cyan

if (-not (Test-Path (Join-Path $SkillDir ".env"))) {
    Copy-Item (Join-Path $SkillDir ".env.example") (Join-Path $SkillDir ".env")
    Write-Host "[SlideForge] Created .env from .env.example — please fill in your OPENAI_API_KEY." -ForegroundColor Yellow
}

if (Get-Command uv -ErrorAction SilentlyContinue) {
    Write-Host "[SlideForge] uv detected — validating dependencies..." -ForegroundColor Green
    uv run --with requests --with pillow --with numpy --with python-pptx --with python-dotenv python -c "import requests, PIL, numpy, pptx, dotenv; print('[ok] All Python dependencies ready!')"
} else {
    Write-Host "[SlideForge] Installing requirements via pip..." -ForegroundColor Cyan
    python -m pip install -r (Join-Path $SkillDir "requirements.txt")
}

Write-Host "[SlideForge] Setup complete! Read SKILL.md or README.md to get started." -ForegroundColor Green
