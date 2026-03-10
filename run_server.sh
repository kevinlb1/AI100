#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${VENV_DIR:-$SCRIPT_DIR/.venv}"
APP_BASE_PATH="${APP_BASE_PATH:-/AI100}"
export APP_BASE_PATH

cd "$SCRIPT_DIR"

if [ ! -f "$VENV_DIR/bin/activate" ]; then
  echo "Virtual environment not found at $VENV_DIR."
  echo "Run ./install_ubuntu.sh first."
  exit 1
fi

# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"

if ! command -v gunicorn >/dev/null 2>&1; then
  echo "gunicorn is not installed in the virtual environment."
  echo "Run ./install_server_ubuntu.sh or: source .venv/bin/activate && python -m pip install gunicorn"
  exit 1
fi

exec gunicorn --config "$SCRIPT_DIR/gunicorn.conf.py" app:application
