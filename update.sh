#!/bin/sh
set -e
echo "=== MCP Outlook Update ==="
git pull
docker compose up --build -d
echo "=== Update abgeschlossen ==="
