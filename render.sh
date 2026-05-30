#!/usr/bin/env bash
# Render the deck to PDF (and optionally PNGs) using headless LibreOffice.
# Requires: libreoffice-impress, poppler-utils (pdftoppm).
#
# The two things that make headless conversion reliable in a container:
#   1. the libreoffice-impress component must be installed (not just -core)
#   2. an explicit, writable UserInstallation profile dir via -env:
set -euo pipefail

DECK="${1:-Athanase_Engaged_Ownership_Allocator_Deck.pptx}"
OUTDIR="${2:-render}"
mkdir -p "$OUTDIR"

soffice --headless \
  -env:UserInstallation=file:///tmp/lo_profile \
  --convert-to pdf --outdir "$OUTDIR" "$DECK"

PDF="$OUTDIR/$(basename "${DECK%.pptx}").pdf"
echo "PDF: $PDF"

# Optional: pass "png" as a 3rd arg to also rasterise every slide.
if [[ "${3:-}" == "png" ]]; then
  pdftoppm -png -r 110 "$PDF" "$OUTDIR/slide"
  echo "PNGs: $OUTDIR/slide-*.png"
fi
