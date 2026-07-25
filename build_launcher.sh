#!/bin/sh
# Builds "Inventory Manager.app" from launcher.applescript (macOS only).
# Double-clicking the result starts the app with no Terminal window.
set -e

cd "$(dirname "$0")"

if ! command -v osacompile >/dev/null 2>&1; then
  echo "osacompile not found — this build script only works on macOS." >&2
  exit 1
fi

rm -rf "Inventory Manager.app"
osacompile -o "Inventory Manager.app" launcher.applescript

echo "Built: $(pwd)/Inventory Manager.app"
echo "Double-click it to start the Inventory Manager."
