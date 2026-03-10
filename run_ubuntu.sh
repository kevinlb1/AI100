#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${VENV_DIR:-$SCRIPT_DIR/.venv}"
PYTHON_BIN="${PYTHON_BIN:-python3}"
URL="${APP_URL:-http://127.0.0.1:8000/launch}"

cd "$SCRIPT_DIR"

if [ ! -f "$VENV_DIR/bin/activate" ]; then
  echo "Virtual environment not found at $VENV_DIR."
  echo "Run ./install_ubuntu.sh first."
  exit 1
fi

# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"

if command -v xdg-open >/dev/null 2>&1 && [ -n "${DISPLAY:-}${WAYLAND_DISPLAY:-}" ]; then
  (
    sleep 1
    xdg-open "$URL" >/dev/null 2>&1 || true
  ) &
fi

echo "Starting app on http://127.0.0.1:8000"
exec "$PYTHON_BIN" app.py
