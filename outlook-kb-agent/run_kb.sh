#!/bin/zsh
cd "$(dirname "$0")"
source .venv/bin/activate
python kb_mail.py "$@"
