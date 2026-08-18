#!/usr/bin/env bash

set -euo pipefail

PROJECT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
cd "$PROJECT_DIR"

if [[ -x "$PROJECT_DIR/.venv/bin/python" ]]; then
  LOOKUP_STUDIO_PYTHON="$PROJECT_DIR/.venv/bin/python"
else
  LOOKUP_STUDIO_PYTHON="${LOOKUP_STUDIO_PYTHON:-python3}"
fi

if ! command -v "$LOOKUP_STUDIO_PYTHON" >/dev/null 2>&1; then
  echo "Python was not found. Install Python 3 or create a .venv in $PROJECT_DIR." >&2
  exit 1
fi

if ! "$LOOKUP_STUDIO_PYTHON" -c "import streamlit" >/dev/null 2>&1; then
  echo "Streamlit is not installed for $LOOKUP_STUDIO_PYTHON." >&2
  echo "Install the project dependencies with:" >&2
  echo "  $LOOKUP_STUDIO_PYTHON -m pip install -r $PROJECT_DIR/requirements.txt" >&2
  exit 1
fi

exec "$LOOKUP_STUDIO_PYTHON" -m streamlit run "$PROJECT_DIR/app.py" "$@"
