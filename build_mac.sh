#!/bin/zsh
set -e

cd "$(dirname "$0")"

python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements-mac.txt
python -m PyInstaller --noconfirm --clean CarnetVolTelepilotePRO_mac.spec

echo
echo "Application creee : dist/CarnetVolTelepilotePRO.app"
