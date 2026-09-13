#!/bin/sh
set -eu
cd "$(dirname "$0")"
python3 -m PyInstaller --clean --noconfirm main.spec
