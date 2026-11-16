#!/bin/bash
# Start the Galaxium back office: web view and MCP server on http://localhost:8090
# Usage: ./start.sh [--pack win-loss] [--port 8090]
set -e
cd "$(dirname "${BASH_SOURCE[0]}")"
if [ ! -x .venv/bin/python ]; then
  echo "First start: setting up (about a minute, needs internet once)..."
  python3 -m venv .venv
  .venv/bin/pip install --quiet --disable-pip-version-check -r requirements.txt
fi
exec .venv/bin/python -m backoffice "$@"
