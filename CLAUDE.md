# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Is

**Workbook Copilot** — an Excel Office Add-in (taskpane) that lets users chat with their workbook data. The frontend reads cell selections via Office.js, POSTs them to a FastAPI backend, which routes through an LLM provider (OpenAI / Anthropic) and optionally calls external MCP tool servers. Responses stream back as NDJSON and are applied to Excel (cell values, formatting, charts).

---

## Commands

### Backend

```bash
cd backend
python -m venv .venv
.venv/Scripts/activate                     # Windows
pip install -r requirements.txt

# Run with SSL (required for Excel sideloading)
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000 \
  --ssl-certfile "$USERPROFILE/.office-addin-dev-certs/localhost.crt" \
  --ssl-keyfile  "$USERPROFILE/.office-addin-dev-certs/localhost.key"

# Run without SSL (browser testing only)
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### Frontend

```bash
cd frontend
npm install
npm run dev:cert      # Install trusted HTTPS cert (run once, or on cert expiry)
npm start             # Serve taskpane at https://localhost:3000

npm run build         # Production bundle
npm run lint          # ESLint (ts, tsx)
npm run typecheck     # tsc --noEmit
```

### Override API URL or default provider

```bash
API_BASE_URL="https://localhost:8000" DEFAULT_PROVIDER="anthropic" npm start
```

---

## Configuration (`backend/.env`)

All settings use the `COPILOT_` prefix (handled by `pydantic-settings`). Key variables:

| Variable | Purpose |
|---|---|
| `COPILOT_ANTHROPIC_API_KEY` | Anthropic API key |
| `COPILOT_ANTHROPIC_MODEL` | Default: `claude-3-5-sonnet-20240620` |
| `COPILOT_OPENAI_API_KEY` | OpenAI API key |
| `COPILOT_OPENAI_MODEL` | Default: `gpt-4o-mini` |
| `COPILOT_REQUEST_TIMEOUT_SECONDS` | Overall request timeout |

Settings are loaded once and cached via `@lru_cache` in `config.py`. To reload settings during development, restart uvicorn.

---

## Architecture

### Request Flow

```
User types prompt → App.tsx handleSend()
  ├─ getWorkbookMetadata()          (once at init, re-used)       [excel.ts]
  ├─ initPreviewCache()             (once at init, registers event listeners) [excel.ts]
  ├─ getUserContext()                (fresh each send)             [excel.ts]
  ├─ getLightweightSheetPreview(50)  (event-cached CSV preview)    [excel.ts]
  ├─ getSelectionsFromPrompt() / getCurrentSelection()             [excel.ts]
  └─ POST /chat { prompt, provider, messages, selection,
                  workbookMetadata, userContext, activeSheetPreview }
       → LangGraphOrchestrator.stream()                           [orchestrator.py]
           ReAct loop (up to MAX_REACT_ITERATIONS=8):
           → _get_live_mcp_tools()  (per-request MCP health check)
           → provider.stream_result()                             [providers.py]
           → If LLM returns MCP tool call (mcp__<server>__<tool>):
               invoke MCP server-side, inject result, loop again
           → If LLM returns Excel tool call (needs_data):
               emit {type:"tool_call_required", payload:[...]}
               emit {type:"done"}
               → Frontend executes Office.js tool → re-POST with toolResults
               → (max 3 tool-call rounds on frontend side)
           → If LLM returns answer:
               stream message_start / message_delta / message_done
               emit cell_updates, format_updates, chart_inserts, pivot_table_inserts
               emit telemetry, done
       → handleStreamEvent()                                      [App.tsx]
           → applyCellUpdates()      [excel.ts]
           → applyFormatUpdates()    [excel.ts]
           → insertCharts()          [excel.ts]
           → insertPivotTables()     [excel.ts]

=MYEXCELCOMPANION.ASKAI(query, ranges...) → functions.ts askAI()
  ├─ Check fingerprint cache → if input changed, return cached result
  ├─ If same inputs (manual recalc), evict cache → re-fetch
  ├─ getSharedProvider() → read provider from window global
  ├─ rangesToContext(ranges) → serialize cell data
  └─ POST /chat (NDJSON, single-shot, no tool calls)
      → cell shows #BUSY! → final answer (spill 2D)
```

### Backend Modules

- **`main.py`** — FastAPI app factory. Registers CORS, constructs `MCPServerService` and `LangGraphOrchestrator`. Owns all route handlers.
- **`orchestrator.py`** — Central logic. ReAct agent loop (up to `MAX_REACT_ITERATIONS = 8`). Per-request MCP health check via `_get_live_mcp_tools()`. MCP tool calls (`mcp__<server_id>__<tool_name>`) are routed server-side; Excel tool calls emit `tool_call_required` for browser round-trip. LangGraph graph handles the non-streaming fallback (`run()`).
- **`providers.py`** — Two providers: `OpenAIProvider`, `AnthropicProvider`. `MCPToolEntry` dataclass + `build_system_prompt(mcp_tools)` dynamically construct the system prompt with available MCP tools. Provider `__init__` accepts `mcp_tools`. `stream_result()` method on all providers. Both OpenAI and Anthropic attempt JSON mode and fall back to lax parsing.
- **`mcp.py`** — MCP server persistence (`MCPServerStore` → JSON file at `data/mcp_servers.json`). Single transport client: `MCPJsonRpcClient` (JSON-RPC 2.0 with session init/terminate). `MCPServerService` uses JSON-RPC exclusively.
- **`schemas.py`** — All Pydantic models shared across the backend. Includes `WorkbookMetadata`, `SheetMetadata`, `UserContext`, `WorkbookToolCall`, `WorkbookToolResult`. Source of truth for message kinds (`thought`, `step`, `final`, `suggestion`, `context`), cell/format/chart update shapes.
- **`config.py`** — `Settings` (pydantic-settings, `extra="ignore"`). Always use `get_settings()` — never instantiate `Settings` directly.

### Frontend Modules

- **`App.tsx`** — Root component. State: messages, workbookMetadata (init once), thinkingSteps, provider selection, MCP server list, busy flags. Init calls `getWorkbookMetadata()` and `initPreviewCache()`. `handleSend()` pre-reads userContext + activeSheetPreview on every send. `streamRound()` handles one HTTP POST+stream round. Tool-call loop retries up to 3 times when `tool_call_required` is received. `handleNewChat()` resets messages and thinking steps for a fresh conversation.
- **`excel.ts`** — All Office.js interactions. `getWorkbookMetadata()` collects filename + per-sheet dims. `getUserContext()` returns active sheet name + selected range. `initPreviewCache()` registers `onChanged`/`onActivated` listeners for cache invalidation. `getLightweightSheetPreview(maxRows)` returns event-cached CSV of first N rows (only re-reads when data changes or sheet switches). Five on-demand read tools: `getXlCellRanges` (batched single sync), `getXlRangeAsCsv`, `xlSearchData`, `getAllXlObjects` (two-phase sync), `executeXlOfficeJs`. `executeWorkbookTool(call)` dispatches to the right function. `getSelectionsFromPrompt()` extracts explicit range refs before falling back to `getCurrentSelection()`. `applyCellUpdates()` supports `replace` and `append` modes. `insertPivotTables()` supports explicit `destinationWorksheet` and `destinationAddress` (including `"Sheet2!E1"` format).
- **`functions/functions.ts`** — `=MYEXCELCOMPANION.ASKAI()` custom function. Non-streaming async (`Promise<string[][]>`), cached, spill-aware. Participates in Excel's standard calculation engine so F2+Enter and Ctrl+Shift+F9 trigger fresh API calls. Uses `fetch` with NDJSON parsing, `parseAnswerTo2D()` for 2D output, fingerprint-based caching (`computeCacheKey()`), `rangesToContext()` for serializing cell data.
- **`functions/metadata.json`** — Custom function schema registered with Office runtime. Declares parameters, result shape, and `cancelable`/`requiresAddress` options (`stream: false`).
- **`sharedState.ts`** — Window-global-backed provider and cache state shared between taskpane and custom function bundles. Provides `getSharedProvider()`, `setSharedProvider()`, `getAskAICache()`, `clearAskAICache()`.
- **`config.ts`** — `API_BASE_URL` and `DEFAULT_PROVIDER` are injectable via webpack `process.env`.
- **`types.ts`** — TypeScript mirror of backend schemas. Includes `WorkbookMetadata`, `UserContext`, `WorkbookToolCall/Result`. `ChatStreamEvent` discriminated union includes `tool_call_required`, `status`, and `suggestion` event types.

---

## Key Patterns & Conventions

### LLM Response Format
All providers share `WORKBOOK_COPILOT_SYSTEM_PROMPT` (in `providers.py`). The LLM returns strict JSON in one of two formats:
- **Option A** (direct answer): `{"answer": "...", "cell_updates": [...], "format_updates": [...], "chart_inserts": [...], "pivot_table_inserts": [...], "suggestion": "..."}`
- **Option B** (needs data): `{"needs_data": true, "tool_call": {"tool": "get_xl_range_as_csv", "args": {...}}}`

The system prompt is the contract — changes there affect all parsing in `assemble_messages_from_payload()` and `build_*` helpers. Both OpenAI and Anthropic providers attempt JSON mode and fall back to lax parsing on `response_format` errors.

### Workbook Context Pre-Read Pattern
On every message send, the frontend collects fresh context in parallel before POSTing:
1. `workbookMetadata` — collected once at add-in init, re-used (filename + all sheets with dims)
2. `userContext` — active sheet name + selected range address (fresh each send)
3. `activeSheetPreview` — first 50 rows of active sheet as CSV (event-cached; only re-reads when sheet data changes or user switches sheets, via listeners registered by `initPreviewCache()`)
4. `selection` — explicit range refs from prompt text, or current selection

### Excel Read Tool Round-Trip
When the LLM needs more data than the pre-read provides, it returns `needs_data: true` with a `tool_call`. The orchestrator emits `{type:"tool_call_required"}`, the frontend executes the Office.js tool, and re-POSTs with `toolResults`. Max 3 rounds per conversation turn. The five tools are: `get_xl_cell_ranges`, `get_xl_range_as_csv`, `xl_search_data`, `get_all_xl_objects`, `execute_xl_office_js`.

### MCP Tool Result Handling
`orchestrator._summarize_tool_result()` tries to render results as a markdown table (up to 10 rows). It handles MongoDB-style extended JSON (`$date`, `$oid`, `$numberDouble`). Results are injected as `MessageKind.CONTEXT` / `MessageRole.SYSTEM` messages before the LLM call, not as separate API tool calls.

### Chart Type Normalization
Chart type aliases are defined in **both** `providers.py` (`CHART_TYPE_ALIASES` dict) and `excel.ts` (`CHART_TYPE_ALIASES` record). If a new chart type is added, update both. The frontend does a second normalization pass against `Excel.ChartType` enum values.

### Streaming vs Non-Streaming
`POST /chat` checks the `Accept` header. `application/x-ndjson` → streaming path. Anything else → `orchestrator.run()` returns a single `ChatResponse`. The frontend always requests NDJSON.

### MCP Server State
MCP server configs persist to `backend/app/data/mcp_servers.json` (path configurable via `COPILOT_MCP_CONFIG_PATH`). The store uses atomic write (`.tmp` file then rename) with a threading `Lock`.

### Custom Function Cache Pattern
`=ASKAI` results are cached in a `Map<string, AskAICacheEntry>` on `window.__MYEXCELCOMPANION_CACHE`. Each entry stores the result and a range fingerprint. The cache key is `callerAddress + "||" + query`. Fingerprint-based logic distinguishes auto-recalc (input data changed → return cached result) from manual recalc (F2+Enter or Ctrl+Shift+F9 → same inputs → evict & re-fetch). Because the function is non-streaming (`stream: false`), it participates in Excel's standard calculation engine: F2+Enter and Ctrl+Shift+F9 both re-invoke the function. `clearAskAICache()` is still available to force a full cache clear.

### Shared Runtime State
The taskpane and custom function bundles are separate webpack chunks that share state via typed window globals (`sharedState.ts`). `setSharedProvider()` is called from `App.tsx` whenever the provider changes. `getSharedProvider()` is called from `functions.ts` on each `=ASKAI` invocation. The `=ASKAI` function collects `cell_updates`, `format_updates`, `chart_inserts`, and `pivot_table_inserts` mutations from the stream and applies them via the taskpane mutation handler bridge. For pivot/chart mutations, the formula cell shows a brief confirmation (e.g. "Pivot table created.") instead of the verbose LLM answer. Destination addresses for pivots/charts are injected from the caller address or an active-cell fallback.

---

## Git & Tracking

- Get changes reviewed by user before committing and pushing.
- After each meaningful task, record the GITSHA in `enhancements.md` as a Phase Checkpoint.
- When adding dependencies: update `requirements.txt` (backend) or `package-lock.json` (frontend).
- `.env`, `manifest.md`, and SSL certs are gitignored.

## Documentation

- Python: Google-style docstrings (Args, Returns, Raises) for all classes and public methods.
- TypeScript: JSDoc (`@param`, `@returns`, `@throws`) for all exported functions.
- Inline comments should explain *why*, not *what*.
