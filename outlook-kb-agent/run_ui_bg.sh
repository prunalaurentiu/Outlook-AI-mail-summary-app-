#!/usr/bin/env bash
set -euo pipefail

PORT="${PORT:-8000}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

PYTHON_BIN="$(command -v python3 || command -v python || true)"
if [[ -z "$PYTHON_BIN" ]]; then
  echo "[run_ui_bg] Python 3 is required but was not found in PATH." >&2
  exit 1
fi

if [[ ! -x ".venv/bin/python" ]]; then
  "$PYTHON_BIN" -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate

python - <<'PY'
import importlib, subprocess, sys
required = {
    "fastapi": "fastapi",
    "uvicorn": "uvicorn[standard]",
    "multipart": "python-multipart",
}
missing = []
for module, package in required.items():
    try:
        importlib.import_module(module)
    except Exception:
        missing.append(package)
if missing:
    print("[run_ui_bg] Installing:", ", ".join(missing))
    subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
PY

# Stop previous instance if the port is busy
if command -v lsof >/dev/null 2>&1; then
  if pids="$(lsof -ti tcp:"$PORT" || true)" && [[ -n "$pids" ]]; then
    echo "[run_ui_bg] Stopping existing process on port $PORT ($pids)"
    kill -9 $pids 2>/dev/null || true
  fi
elif command -v fuser >/dev/null 2>&1; then
  if fuser -k "$PORT"/tcp >/dev/null 2>&1; then
    echo "[run_ui_bg] Freed port $PORT via fuser"
  fi
fi

LOG_DIR="$SCRIPT_DIR/.logs"
mkdir -p "$LOG_DIR"
LOG_FILE="$LOG_DIR/ui.log"

nohup uvicorn app:app --port "$PORT" --log-level info > "$LOG_FILE" 2>&1 &
UVICORN_PID=$!

echo "[run_ui_bg] Uvicorn started with PID $UVICORN_PID"
URL="http://127.0.0.1:$PORT/"

if command -v open >/dev/null 2>&1; then
  open "$URL" >/dev/null 2>&1 || true
elif command -v xdg-open >/dev/null 2>&1; then
  xdg-open "$URL" >/dev/null 2>&1 || true
elif command -v start >/dev/null 2>&1; then
  start "$URL" >/dev/null 2>&1 || true
else
  echo "[run_ui_bg] Navigate to $URL"
fi

echo "[run_ui_bg] Logs: $LOG_FILE"
