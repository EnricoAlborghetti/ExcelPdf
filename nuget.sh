#!/usr/bin/env bash

# Exit immediately on errors, treat unset variables as errors
set -euo pipefail

SOURCE="https://nuget.point2point.it/nuget"
API_KEY="${NUGET_API_KEY:-Polenta9999}"

DIRS=("")

echo "Starting NuGet package push to $SOURCE..."

BIN_DIR="./bin"

if [[ ! -d "$BIN_DIR" ]]; then
    echo "⚠️  Directory not found: $BIN_DIR (Skipping)"
    continue
fi

# Find the newest package safely across all operating systems
PACKAGE=""

# Process find output safely, even with spaces in filenames
while IFS= read -r -d '' file; do
    # Ignore symbol packages
    if [[ "$file" != *.symbols.nupkg && "$file" != *.snupkg ]]; then
        # If PACKAGE is empty, OR if this file is newer than the current PACKAGE
        if [[ -z "$PACKAGE" || "$file" -nt "$PACKAGE" ]]; then
            PACKAGE="$file"
        fi
    fi
done < <(find "$BIN_DIR" -type f -name "*.nupkg" -print0 2>/dev/null || true)

if [[ -n "$PACKAGE" && -f "$PACKAGE" ]]; then
    echo "📦 Pushing $PACKAGE..."
    dotnet nuget push "$PACKAGE" -k "$API_KEY" -s "$SOURCE"
    echo "✅ Successfully pushed $PACKAGE"
    echo "----------------------------------------"
else
    echo "ℹ️  No valid .nupkg found in $BIN_DIR"
fi

echo "🎉 All done!"