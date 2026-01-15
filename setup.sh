#!/bin/bash
# setup.sh - One-command development environment setup
# Executive Business Review PowerPoint Generator

set -e  # Exit on error

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${SCRIPT_DIR}/venv"
PYTHON_MIN_VERSION="3.10"

echo "======================================"
echo "Executive Business Review - Setup"
echo "======================================"
echo ""

# Check Python version
check_python_version() {
    local python_cmd="${1:-python3}"

    if ! command -v "$python_cmd" &> /dev/null; then
        echo "Error: Python not found. Please install Python 3.10+"
        exit 1
    fi

    local version=$($python_cmd -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')
    local major=$(echo "$version" | cut -d. -f1)
    local minor=$(echo "$version" | cut -d. -f2)

    if [ "$major" -lt 3 ] || ([ "$major" -eq 3 ] && [ "$minor" -lt 10 ]); then
        echo "Error: Python 3.10+ required, found $version"
        exit 1
    fi
    echo "Found Python $version"
}

# Create virtual environment
create_venv() {
    if [ -d "$VENV_DIR" ]; then
        echo "Virtual environment already exists at $VENV_DIR"
        read -p "Recreate it? [y/N]: " confirm
        if [[ "$confirm" =~ ^[Yy]$ ]]; then
            echo "Removing existing virtual environment..."
            rm -rf "$VENV_DIR"
        else
            echo "Using existing virtual environment."
            return 0
        fi
    fi

    echo "Creating virtual environment..."
    python3 -m venv "$VENV_DIR"
    echo "Virtual environment created."
}

# Install dependencies
install_deps() {
    echo ""
    echo "Installing Python dependencies..."
    "$VENV_DIR/bin/pip" install --upgrade pip --quiet
    "$VENV_DIR/bin/pip" install -r "${SCRIPT_DIR}/requirements.txt" --quiet
    echo "Python dependencies installed."

    echo ""
    echo "Installing Playwright browsers (this may take a minute)..."
    "$VENV_DIR/bin/playwright" install chromium --with-deps 2>/dev/null || \
        "$VENV_DIR/bin/playwright" install chromium
    echo "Playwright browsers installed."
}

# Verify installation
verify_install() {
    echo ""
    echo "Verifying installation..."
    "$VENV_DIR/bin/python" -c "
import pptx
import playwright
import openpyxl
import yaml
from PIL import Image
print('All core dependencies verified.')
"
}

# Create output directory
setup_directories() {
    mkdir -p "${SCRIPT_DIR}/output"
    mkdir -p "${SCRIPT_DIR}/temp_charts"
}

# Main execution
main() {
    check_python_version
    create_venv

    # Activate venv for subsequent commands
    source "$VENV_DIR/bin/activate"

    install_deps
    verify_install
    setup_directories

    echo ""
    echo "======================================"
    echo "Setup complete!"
    echo "======================================"
    echo ""
    echo "To activate the environment:"
    echo "  source venv/bin/activate"
    echo ""
    echo "To generate a presentation:"
    echo "  ./run.sh --config clients/sample.yaml"
    echo ""
    echo "For help:"
    echo "  ./run.sh --help"
    echo ""
}

main "$@"
