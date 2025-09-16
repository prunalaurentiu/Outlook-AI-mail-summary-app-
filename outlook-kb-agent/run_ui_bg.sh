#!/bin/zsh
set -euo pipefail

proj="/Users/prunalaurentiu/Documents/outlook-kb-agent"
cd "$proj"

# Activează venv
if [[ ! -x ".venv/bin/python" ]]; then
  /usr/bin/python3 -m venv .venv
fi
source .venv/bin/activate

# (opțional) verifică dependențele minime pentru UI
python - <<'PY'
import importlib, sys, subprocess
missing=[]
# FastAPI se importă drept "fastapi", python-multipart drept "multipart"
for mod in ("fastapi","uvicorn","multipart"):
    try: importlib.import_module(mod)
    except Exception: missing.append(mod if mod!="multipart" else "python-multipart")
if missing:
    print("[setup] Installing:", ", ".join(missing))
    subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
PY

# Oprește instanța anterioară (dacă există)
lsof -ti tcp:8000 | xargs -r kill -9 || true

# Pornește serverul în background și loghează în Librărie
nohup uvicorn app:app --port 8000 --log-level info > "$HOME/Library/Logs/outlook-kb-ui.log" 2>&1 &

# Deschide browserul pe UI
sleep 1
open "http://127.0.0.1:8000/"
