#!/bin/bash
# macOS/Linux launcher. Keep this file in the same folder as
# report_builder.py. First time only: right-click > Open (macOS may warn
# about an unidentified developer - that's expected for your own script),
# or run: chmod +x Run_AlloyBench.command

cd "$(dirname "$0")"

if command -v python3 >/dev/null 2>&1; then
    python3 report_builder.py
elif command -v python >/dev/null 2>&1; then
    python report_builder.py
else
    echo "Python was not found. Install it from https://python.org"
    read -p "Press Enter to close..."
    exit 1
fi
