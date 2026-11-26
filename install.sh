#!/usr/bin/env bash
set -euo pipefail

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
echo "[SlideForge] Checking environment..."

if [ ! -f "$SKILL_DIR/.env" ]; then
  cp "$SKILL_DIR/.env.example" "$SKILL_DIR/.env"
  echo "[SlideForge] Created .env from .env.example — please fill in your OPENAI_API_KEY."
fi

if command -v uv >/dev/null 2>&1; then
  echo "[SlideForge] uv detected — validating dependencies..."
  uv run --with requests --with pillow --with numpy --with python-pptx --with python-dotenv \
    python -c "import requests, PIL, numpy, pptx, dotenv; print('[ok] All Python dependencies ready!')"
else
  echo "[SlideForge] Installing requirements via pip..."
  python3 -m pip install -r "$SKILL_DIR/requirements.txt"
fi

echo "[SlideForge] Setup complete! Read SKILL.md or README.md to get started."
