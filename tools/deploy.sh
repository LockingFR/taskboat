#!/bin/sh
# Publie uniquement les fichiers du site (ni tests, ni docs). Usage : sh tools/deploy.sh [branche]
set -e
cd "$(dirname "$0")/.."
rm -rf dist && mkdir dist
cp -r _headers index.html 404.html robots.txt sitemap.xml assets en exporter-microsoft-to-do microsoft-to-do-vers-todoist mentions-legales dist/
npx --yes wrangler@latest pages deploy dist --project-name taskboat --branch "${1:-main}" --commit-dirty=true
