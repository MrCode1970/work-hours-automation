#!/usr/bin/env bash
set -e

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$PROJECT_DIR/.venv"

if [ ! -d "$VENV_DIR" ]; then
  echo "❌ Virtual environment not found (.venv)"
  echo "Run: python3.10 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt"
  exit 1
fi

source "$VENV_DIR/bin/activate"

echo "▶ Using python: $(which python)"
python --version

python run.py
