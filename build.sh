#!/usr/bin/env bash
# Local macOS / Linux build helper
set -euo pipefail
cd "$(dirname "$0")"

if [[ ! -d .venv ]]; then
  python3 -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate
pip install -q -r requirements-build.txt
rm -rf build dist
pyinstaller --noconfirm control_cable_scheduler.spec

echo
echo "Build complete."
if [[ "$(uname -s)" == "Darwin" ]]; then
  echo "  App: dist/Control Cable Scheduler.app"
else
  echo "  Dir: dist/Control Cable Scheduler/"
fi
