#!/bin/bash
# Start the Galaxium back office: web view and MCP server on http://localhost:8090
# Usage: ./start.sh [--pack win-loss] [--port 8090]
# A back office already serving on the port is stopped first, so this always restarts cleanly.
set -e
cd "$(dirname "${BASH_SOURCE[0]}")"
if [ ! -x .venv/bin/python ]; then
  echo "First start: setting up (about a minute, needs internet once)..."
  python3 -m venv .venv
  .venv/bin/pip install --quiet --disable-pip-version-check -r requirements.txt
fi

# Find the port and whether this run serves over HTTP (not 'reset', 'packs' or --stdio).
port=8090
serve=1
prev=
for arg in "$@"; do
  case "$arg" in
    --port=*) port="${arg#--port=}" ;;
    --stdio|reset|packs) serve=0 ;;
  esac
  [ "$prev" = "--port" ] && port="$arg"
  prev="$arg"
done

if [ "$serve" = 1 ]; then
  pids=$(lsof -nP -t -iTCP:"$port" -sTCP:LISTEN 2>/dev/null || true)
  for pid in $pids; do
    cmd=$(ps -o command= -p "$pid" 2>/dev/null || true)
    if [[ "$cmd" != *backoffice* ]]; then
      echo "Port $port is in use by another program (pid $pid: $cmd). Stop it or pass --port." >&2
      exit 1
    fi
    echo "Stopping the back office already running on port $port (pid $pid)..."
    kill "$pid" 2>/dev/null || true
  done
  # Wait up to 5 seconds for the port to free, then force it.
  for _ in $(seq 50); do
    lsof -nP -t -iTCP:"$port" -sTCP:LISTEN >/dev/null 2>&1 || break
    sleep 0.1
  done
  pids=$(lsof -nP -t -iTCP:"$port" -sTCP:LISTEN 2>/dev/null || true)
  [ -n "$pids" ] && kill -9 $pids 2>/dev/null || true
fi

exec .venv/bin/python -m backoffice "$@"
