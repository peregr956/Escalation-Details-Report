#!/bin/bash
# run.sh - Run the Executive Business Review generator
# Handles venv activation and argument passthrough

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${SCRIPT_DIR}/venv"

# Check if venv exists
if [ ! -d "$VENV_DIR" ]; then
    echo "Error: Virtual environment not found at $VENV_DIR"
    echo ""
    echo "Please run setup first:"
    echo "  ./setup.sh"
    exit 1
fi

# Check if requirements are satisfied
if [ ! -f "$VENV_DIR/bin/python" ]; then
    echo "Error: Virtual environment appears corrupted."
    echo ""
    echo "Please re-run setup:"
    echo "  ./setup.sh"
    exit 1
fi

# Activate venv and run the generator with all passed arguments
source "$VENV_DIR/bin/activate"
python "${SCRIPT_DIR}/generate_presentation.py" "$@"
