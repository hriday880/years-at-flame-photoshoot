#!/bin/zsh
set -e
cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "Creating local Python environment..."
  python3 -m venv .venv
fi

source .venv/bin/activate

if ! python3 -c "import flask, openpyxl, waitress" 2>/dev/null; then
  echo "Installing dependencies (flask, openpyxl, waitress)..."
  pip install --quiet flask openpyxl waitress
fi

python3 AppUpdate.py
