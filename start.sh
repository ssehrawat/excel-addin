#!/usr/bin/env bash
# =============================================================================
# start.sh — Boot both Workbook Copilot backend and frontend
# =============================================================================
# Usage:  bash start.sh [--install]
#
#   --install   Force re-install of Python and Node dependencies
#
# Requires: Python 3.11+, Node.js 18+, npm
# Platform: Windows (Git Bash / WSL) or macOS / Linux
# =============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BACKEND_DIR="$SCRIPT_DIR/backend"
FRONTEND_DIR="$SCRIPT_DIR/frontend"

# ---------------------------------------------------------------------------
# Colors (graceful fallback)
# ---------------------------------------------------------------------------
if [[ -t 1 ]]; then
  GREEN='\033[0;32m'; YELLOW='\033[1;33m'; RED='\033[0;31m'; NC='\033[0m'
else
  GREEN=''; YELLOW=''; RED=''; NC=''
fi

info()  { echo -e "${GREEN}[OK]${NC} $*"; }
warn()  { echo -e "${YELLOW}[WARN]${NC} $*"; }
error() { echo -e "${RED}[ERROR]${NC} $*"; }

# ---------------------------------------------------------------------------
# Parse flags
# ---------------------------------------------------------------------------
FORCE_INSTALL=false
for arg in "$@"; do
  case "$arg" in
    --install) FORCE_INSTALL=true ;;
    -h|--help)
      echo "Usage: ./start.sh [--install]"
      echo "  --install  Force re-install of Python and Node dependencies"
      exit 0
      ;;
    *) error "Unknown flag: $arg"; exit 1 ;;
  esac
done

# ---------------------------------------------------------------------------
# Prerequisite checks
# ---------------------------------------------------------------------------
PYTHON=""
for cmd in python python3; do
  if command -v "$cmd" &>/dev/null; then
    ver=$("$cmd" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>/dev/null || echo "0.0")
    major="${ver%%.*}"
    minor="${ver#*.}"
    if [[ "$major" -ge 3 && "$minor" -ge 11 ]]; then
      PYTHON="$cmd"
      break
    fi
  fi
done

if [[ -z "$PYTHON" ]]; then
  error "Python 3.11+ is required but not found. Install it from https://www.python.org/downloads/"
  exit 1
fi
info "Python: $($PYTHON --version)"

if ! command -v node &>/dev/null; then
  error "Node.js is required but not found. Install it from https://nodejs.org/"
  exit 1
fi

NODE_MAJOR=$(node -e "console.log(process.version.split('.')[0].slice(1))")
if [[ "$NODE_MAJOR" -lt 18 ]]; then
  error "Node.js 18+ is required (found v$NODE_MAJOR). Update from https://nodejs.org/"
  exit 1
fi
info "Node.js: $(node --version)"

if ! command -v npm &>/dev/null; then
  error "npm is required but not found."
  exit 1
fi
info "npm: $(npm --version)"

# ---------------------------------------------------------------------------
# SSL certificate check
# ---------------------------------------------------------------------------
CERT_DIR="${USERPROFILE:-$HOME}/.office-addin-dev-certs"
CERT_DIR="${CERT_DIR//\\//}"

CERT_FILE="$CERT_DIR/localhost.crt"
KEY_FILE="$CERT_DIR/localhost.key"

if [[ ! -f "$CERT_FILE" || ! -f "$KEY_FILE" ]]; then
  warn "SSL certificates not found at $CERT_DIR"
  echo "  Excel sideloading requires HTTPS. Generate certs by running:"
  echo ""
  echo "    cd frontend && npm install && npm run dev:cert"
  echo ""
  echo "  Then re-run this script."
  exit 1
fi
info "SSL certs found"

# ---------------------------------------------------------------------------
# Backend: venv + dependencies
# ---------------------------------------------------------------------------
VENV_DIR="$BACKEND_DIR/.venv"

if [[ ! -d "$VENV_DIR" ]]; then
  info "Creating Python virtual environment..."
  "$PYTHON" -m venv "$VENV_DIR"
  FORCE_INSTALL=true
fi

# Activate venv
if [[ -f "$VENV_DIR/Scripts/activate" ]]; then
  source "$VENV_DIR/Scripts/activate"    # Windows (Git Bash)
elif [[ -f "$VENV_DIR/bin/activate" ]]; then
  source "$VENV_DIR/bin/activate"        # macOS / Linux / WSL
else
  error "Cannot find venv activate script in $VENV_DIR"
  exit 1
fi
info "Virtual environment activated"

if [[ "$FORCE_INSTALL" == true ]]; then
  info "Installing Python dependencies..."
  pip install -q -r "$BACKEND_DIR/requirements.txt"
  info "Python dependencies installed"
fi

# ---------------------------------------------------------------------------
# Backend: .env bootstrap
# ---------------------------------------------------------------------------
if [[ ! -f "$BACKEND_DIR/.env" ]]; then
  if [[ -f "$BACKEND_DIR/.env.example" ]]; then
    cp "$BACKEND_DIR/.env.example" "$BACKEND_DIR/.env"
    info "Created backend/.env from .env.example — edit it to add your API keys"
  else
    warn "No backend/.env found. Backend will use defaults (mock provider only)."
  fi
fi

# ---------------------------------------------------------------------------
# Frontend: node_modules
# ---------------------------------------------------------------------------
if [[ ! -d "$FRONTEND_DIR/node_modules" ]] || [[ "$FORCE_INSTALL" == true ]]; then
  info "Installing Node dependencies..."
  (cd "$FRONTEND_DIR" && npm install --no-audit --no-fund)
  info "Node dependencies installed"
fi

# ---------------------------------------------------------------------------
# Start both services, kill both on exit
# ---------------------------------------------------------------------------
BACKEND_PID=""
FRONTEND_PID=""

cleanup() {
  echo ""
  info "Shutting down..."
  [[ -n "$BACKEND_PID" ]]  && kill "$BACKEND_PID"  2>/dev/null
  [[ -n "$FRONTEND_PID" ]] && kill "$FRONTEND_PID" 2>/dev/null
  wait 2>/dev/null
  info "Stopped."
}
trap cleanup SIGINT SIGTERM EXIT

info "Starting backend on https://localhost:8000 ..."
(cd "$BACKEND_DIR" && python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000 \
  --ssl-certfile "$CERT_FILE" \
  --ssl-keyfile "$KEY_FILE") &
BACKEND_PID=$!

info "Starting frontend on https://localhost:3000 ..."
(cd "$FRONTEND_DIR" && npm start) &
FRONTEND_PID=$!

echo ""
info "Both services running. Press Ctrl+C to stop."
echo ""

# Wait for either process to exit
wait -n 2>/dev/null || wait
