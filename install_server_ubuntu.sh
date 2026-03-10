#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [ "$(id -u)" -eq 0 ]; then
  APT_CMD=""
else
  if ! command -v sudo >/dev/null 2>&1; then
    echo "This script needs sudo for apt package installation."
    exit 1
  fi
  APT_CMD="sudo"
fi

cd "$SCRIPT_DIR"

$APT_CMD apt-get update
$APT_CMD apt-get install -y python3 python3-venv python3-pip nginx

"$SCRIPT_DIR/install_ubuntu.sh"

# shellcheck disable=SC1091
source "$SCRIPT_DIR/.venv/bin/activate"
python -m pip install gunicorn

echo
echo "Server prerequisites installed."
echo "Next steps:"
echo "  1. Review ai100.service.example and nginx-ai100.conf.example"
echo "  2. Copy them into /etc/systemd/system/ and /etc/nginx/sites-available/"
echo "  3. Enable the systemd service and nginx site"
