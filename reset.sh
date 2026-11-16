#!/bin/bash
# Throw away everything the agent wrote and return to the starting state.
set -e
cd "$(dirname "${BASH_SOURCE[0]}")"
exec .venv/bin/python -m backoffice reset
