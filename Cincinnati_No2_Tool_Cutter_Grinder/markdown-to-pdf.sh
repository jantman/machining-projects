#!/bin/bash
#
# Convert a markdown file to PDF with GitHub-style rendering
# Usage: markdown-to-pdf.sh <markdown-file>
#

set -e

# Check if file argument is provided
if [ $# -eq 0 ]; then
    echo "Usage: $0 <markdown-file>"
    exit 1
fi

MARKDOWN_FILE="$1"

# Check if file exists
if [ ! -f "$MARKDOWN_FILE" ]; then
    echo "Error: File '$MARKDOWN_FILE' not found"
    exit 1
fi

# Check if grip is available
GRIP_PATH="$HOME/venvs/current/bin/grip"
if [ ! -x "$GRIP_PATH" ]; then
    echo "Error: grip not found at $GRIP_PATH"
    exit 1
fi

# Check if chromium is available
if ! command -v chromium &> /dev/null; then
    echo "Error: chromium not found"
    exit 1
fi

# Get the base filename without extension
BASE_NAME="${MARKDOWN_FILE%.md}"

# Generate output filenames
HTML_FILE="${BASE_NAME}_github.html"
PDF_FILE="${BASE_NAME}.pdf"

echo "Converting $MARKDOWN_FILE to PDF..."
echo "Step 1: Generating GitHub-styled HTML..."

# Export markdown to HTML with GitHub styling
"$GRIP_PATH" "$MARKDOWN_FILE" --export "$HTML_FILE"

echo "Step 2: Converting HTML to PDF..."

# Convert HTML to PDF without headers/footers
chromium --headless --disable-gpu \
    --print-to-pdf="$PDF_FILE" \
    --no-pdf-header-footer \
    "$HTML_FILE"

# Clean up the intermediate HTML file
rm -f "$HTML_FILE"

echo "Done! PDF saved to: $PDF_FILE"
