#!/usr/bin/env bash
# Nauta Implementation Tracker — local build + push to GitHub Pages
set -euo pipefail

cd "$(dirname "$0")"

echo "→ Generando index.html ..."
python3 generate.py

if git diff --quiet index.html 2>/dev/null; then
  echo "✓ Sin cambios en index.html. Nada que desplegar."
  exit 0
fi

STAMP=$(date '+%Y-%m-%d %H:%M')
echo "→ Commit + push (gh-pages workflow vía main branch) ..."
git add index.html data/Nauta_Implementation_Tracker.xlsx
git commit -m "dashboard: update ${STAMP}"
git push origin main

REPO_URL=$(gh repo view --json url -q .url 2>/dev/null || echo "")
echo ""
echo "✅ Deploy enviado. GitHub Pages tarda ~30-60s en publicar."
echo "   Repo:  ${REPO_URL}"
echo "   URL:   https://santiago-suarez-nauta.github.io/nauta-implementation-tracker/"
echo "   Fecha: ${STAMP}"
