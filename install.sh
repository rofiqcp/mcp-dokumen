#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

info() { printf '[INFO] %s\n' "$*"; }
warn() { printf '[WARN] %s\n' "$*"; }

PY_CMD=""
for candidate in python3.10 python3 python; do
  if command -v "$candidate" >/dev/null 2>&1; then
    ver="$($candidate - <<'PY'
import sys
print('.'.join(map(str, sys.version_info[:3])))
PY
)"
    major="${ver%%.*}"
    minor="${ver#*.}"; minor="${minor%%.*}"
    if [[ "$major" -gt 3 || ( "$major" -eq 3 && "$minor" -ge 10 ) ]]; then
      PY_CMD="$candidate"
      break
    fi
  fi
done

if [[ -z "$PY_CMD" ]]; then
  warn "Python 3.10+ tidak ditemukan. Instal Python 3.10.x terlebih dahulu (lihat panduan di README)."
  exit 1
fi

info "Menggunakan Python: $PY_CMD"

if [[ ! -d "$REPO_ROOT/.venv" ]]; then
  info "Membuat virtual environment di $REPO_ROOT/.venv"
  "$PY_CMD" -m venv "$REPO_ROOT/.venv"
fi

source "$REPO_ROOT/.venv/bin/activate"

info "Memperbarui pip"
python -m pip install --upgrade pip

info "Meng-install paket dalam mode editable"
python -m pip install -e "$REPO_ROOT"

CONFIG_PATH="$HOME/.config/Code/User/mcp.json"

mkdir -p "$(dirname "$CONFIG_PATH")"

export REPO_ROOT
export CONFIG_PATH

python - <<'PY'
import json
import os
import pathlib
import sys

config_path = pathlib.Path(os.environ["CONFIG_PATH"])
repo_root = pathlib.Path(os.environ["REPO_ROOT"]).resolve()

if config_path.exists():
    backup = config_path.with_suffix(config_path.suffix + ".bak")
    backup.write_bytes(config_path.read_bytes())

data = {}
if config_path.exists() and config_path.stat().st_size > 0:
    try:
        data = json.loads(config_path.read_text())
    except Exception as exc:  # noqa: BLE001
        print(f"[ERROR] Gagal membaca {config_path}: {exc}")
        sys.exit(1)

if not isinstance(data, dict):
    data = {}

servers = data.get("servers")
if not isinstance(servers, dict):
    servers = {}
    data["servers"] = servers

servers["mcpdocx"] = {
    "type": "stdio",
    "command": str(repo_root / ".venv/bin/python"),
    "args": ["-m", "mcpdocx"],
    "cwd": str(repo_root),
    "env": {
        "PYTHONPATH": str(repo_root)
    }
}

if "inputs" not in data:
    data["inputs"] = []

# Hapus mcpServers jika ada (format lama tidak menggunakannya)
if "mcpServers" in data:
    del data["mcpServers"]

config_path.write_text(json.dumps(data, indent=4))
print(f"[OK] Konfigurasi MCP disimpan ke {config_path}")
PY

info "Instalasi selesai. Aktifkan venv dengan: source .venv/bin/activate"
