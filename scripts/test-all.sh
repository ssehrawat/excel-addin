#!/usr/bin/env bash
# =============================================================================
# test-all.sh — Run all E2E test layers in one shot
# =============================================================================
# Usage:
#   bash scripts/test-all.sh            # Layers 1-4 replay (no API keys)
#   bash scripts/test-all.sh --live     # Layers 1-4 with live LLM calls
#
# Exit code: 0 if all pass, 1 if any fail.
# =============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"
BACKEND_DIR="$ROOT_DIR/backend"
FRONTEND_DIR="$ROOT_DIR/frontend"

# WHY: Save PATH before venv activation so Node/npx remain findable
PRE_VENV_PATH="$PATH"

LIVE_FLAG=""
for arg in "$@"; do
  case "$arg" in
    --live) LIVE_FLAG="--live" ;;
  esac
done

# Colors
if [[ -t 1 ]]; then
  GREEN='\033[0;32m'; RED='\033[0;31m'; CYAN='\033[0;36m'; NC='\033[0m'
else
  GREEN=''; RED=''; CYAN=''; NC=''
fi

FAILED=0

step_pass() { echo -e "${GREEN}✓ $1 passed${NC}"; }
step_fail() { echo -e "${RED}✗ $1 FAILED${NC}"; FAILED=1; }

# ---------------------------------------------------------------------------
# Layer 1: Stateful Office.js Workbook Simulator (frontend, Vitest)
# ---------------------------------------------------------------------------
echo -e "\n${CYAN}━━━ Layer 1 — Workbook Simulator ━━━${NC}"
if (cd "$FRONTEND_DIR" && PATH="$PRE_VENV_PATH" npx vitest run src/__tests__/e2e.simulator.test.ts); then
  step_pass "Layer 1"
else
  step_fail "Layer 1"
fi

# ---------------------------------------------------------------------------
# Activate backend venv
# ---------------------------------------------------------------------------
if [[ -f "$BACKEND_DIR/.venv/Scripts/activate" ]]; then
  source "$BACKEND_DIR/.venv/Scripts/activate"
elif [[ -f "$BACKEND_DIR/.venv/bin/activate" ]]; then
  source "$BACKEND_DIR/.venv/bin/activate"
fi

# ---------------------------------------------------------------------------
# Layer 2: Full-Pipeline Integration Tests (backend, pytest)
# ---------------------------------------------------------------------------
echo -e "\n${CYAN}━━━ Layer 2 — Pipeline Integration ━━━${NC}"
if (cd "$BACKEND_DIR" && python -m pytest tests/integration/test_e2e_pipeline.py -v); then
  step_pass "Layer 2"
else
  step_fail "Layer 2"
fi

# ---------------------------------------------------------------------------
# Layer 3: Record-Replay LLM Parsing (backend, pytest)
# ---------------------------------------------------------------------------
echo -e "\n${CYAN}━━━ Layer 3 — Record-Replay Parsing ━━━${NC}"
if (cd "$BACKEND_DIR" && python -m pytest tests/eval/test_replay.py -v); then
  step_pass "Layer 3"
else
  step_fail "Layer 3"
fi

# ---------------------------------------------------------------------------
# Layer 4: Live LLM E2E / Cassette Replay (backend, pytest)
# ---------------------------------------------------------------------------
if [[ -n "$LIVE_FLAG" ]]; then
  echo -e "\n${CYAN}━━━ Layer 4 — Live LLM E2E (recording) ━━━${NC}"
  if (cd "$BACKEND_DIR" && python -m pytest tests/e2e/test_live_llm.py -v --live); then
    step_pass "Layer 4 (live)"
  else
    step_fail "Layer 4 (live)"
  fi
else
  echo -e "\n${CYAN}━━━ Layer 4 — LLM Cassette Replay ━━━${NC}"
  if (cd "$BACKEND_DIR" && python -m pytest tests/e2e/test_live_llm.py -v); then
    step_pass "Layer 4 (replay)"
  else
    step_fail "Layer 4 (replay)"
  fi
fi

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
echo ""
if [[ "$FAILED" -eq 0 ]]; then
  echo -e "${GREEN}━━━ All layers passed ━━━${NC}"
else
  echo -e "${RED}━━━ Some layers FAILED ━━━${NC}"
  exit 1
fi
