#!/bin/bash
# batch_generate.sh - Generate reports for multiple clients
# Executive Business Review PowerPoint Generator

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

# Usage help
show_help() {
    echo "Usage: $0 [OPTIONS] [CLIENT_IDS...]"
    echo ""
    echo "Generate Executive Business Review presentations for multiple clients."
    echo ""
    echo "Options:"
    echo "  -h, --help          Show this help message"
    echo "  -a, --all           Generate for all clients in registry"
    echo "  -t, --tier TIER     Generate for all clients of specified tier"
    echo "  -o, --output DIR    Output directory (default: output/)"
    echo ""
    echo "Examples:"
    echo "  $0 acme burlington             # Generate for specific clients"
    echo "  $0 --all                       # Generate for all clients"
    echo "  $0 --tier 'Signature Tier'     # Generate for all Signature Tier clients"
    echo ""
}

# Parse arguments
OUTPUT_DIR="${PROJECT_DIR}/output"
CLIENTS=()
ALL_CLIENTS=false
TIER_FILTER=""

while [[ $# -gt 0 ]]; do
    case $1 in
        -h|--help)
            show_help
            exit 0
            ;;
        -a|--all)
            ALL_CLIENTS=true
            shift
            ;;
        -t|--tier)
            TIER_FILTER="$2"
            shift 2
            ;;
        -o|--output)
            OUTPUT_DIR="$2"
            shift 2
            ;;
        -*)
            echo "Unknown option: $1"
            show_help
            exit 1
            ;;
        *)
            CLIENTS+=("$1")
            shift
            ;;
    esac
done

# Determine client list
if [ "$ALL_CLIENTS" = true ]; then
    echo "Fetching all clients from registry..."
    CLIENTS=($("$PROJECT_DIR/venv/bin/python" -c "
from client_registry import list_clients
print(' '.join(list_clients()))
" 2>/dev/null || echo ""))
elif [ -n "$TIER_FILTER" ]; then
    echo "Fetching clients for tier: $TIER_FILTER"
    CLIENTS=($("$PROJECT_DIR/venv/bin/python" -c "
from client_registry import list_clients
print(' '.join(list_clients(filter_tier='$TIER_FILTER')))
" 2>/dev/null || echo ""))
fi

if [ ${#CLIENTS[@]} -eq 0 ]; then
    echo "Error: No clients specified."
    echo ""
    show_help
    exit 1
fi

echo "======================================"
echo "Batch Report Generation"
echo "======================================"
echo "Clients: ${CLIENTS[*]}"
echo "Output:  $OUTPUT_DIR"
echo ""

# Generate reports
SUCCESS_COUNT=0
FAIL_COUNT=0
FAILED_CLIENTS=()

for client in "${CLIENTS[@]}"; do
    echo "--------------------------------------"
    echo "Generating report for: $client"
    echo "--------------------------------------"

    if "$PROJECT_DIR/run.sh" --client "$client" --output-dir "$OUTPUT_DIR" 2>&1; then
        echo "SUCCESS: $client"
        ((SUCCESS_COUNT++))
    else
        echo "FAILED: $client"
        ((FAIL_COUNT++))
        FAILED_CLIENTS+=("$client")
    fi
    echo ""
done

# Summary
echo "======================================"
echo "Batch Generation Complete"
echo "======================================"
echo "Success: $SUCCESS_COUNT"
echo "Failed:  $FAIL_COUNT"

if [ ${#FAILED_CLIENTS[@]} -gt 0 ]; then
    echo ""
    echo "Failed clients: ${FAILED_CLIENTS[*]}"
    exit 1
fi
