## Workbook Copilot

An Excel Office Add-in (taskpane) that lets users chat with their workbook data. The frontend reads cell selections via Office.js, POSTs them to a FastAPI backend, which routes through an LLM provider (OpenAI / Anthropic) and optionally calls external MCP tool servers. Responses stream back as NDJSON and are applied to Excel — cell values, formatting, charts, and pivot tables.

### Features

- **Streaming chat** — thinking steps show progress with live spinners, then the final response appears
- **Cell updates** — LLM can write values into cells (replace or append mode)
- **Format updates** — fill color, font color, bold, italic, number format, and borders applied to specified ranges
- **Chart inserts** — bar, column, line, scatter, bubble with axis titles
- **Pivot table inserts** — rows, columns, values, filters inferred from sheet context
- **MCP tool servers** — connect external data sources (databases, APIs) that the LLM can query server-side
- **Excel read tools** — LLM can request additional workbook data mid-conversation (up to 3 tool-call rounds per message)
- **Multi-provider** — switch between OpenAI and Anthropic at runtime

### Quick Start

The fastest way to get everything running:

```bash
# 1. Generate SSL certs (first time only)
cd frontend && npm install && npm run dev:cert && cd ..

# 2. Start both backend and frontend
bash start.sh
```

The startup script checks prerequisites, creates a Python venv, installs dependencies,
copies `.env.example` to `.env` if needed, and boots both services with HTTPS.

Use `bash start.sh --install` to force-reinstall all dependencies.

### Prerequisites

- Excel desktop with sideloading enabled (Microsoft 365 recommended)
- Node.js 18+
- npm 9+
- Python 3.11+

### Backend setup

```bash
cd backend
python -m venv .venv
.venv/Scripts/activate                     # Windows
pip install -r requirements.txt
cp .env.example .env          # then edit .env to add your API keys

# Run with SSL (required for Excel sideloading)
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000 \
  --ssl-certfile "$USERPROFILE/.office-addin-dev-certs/localhost.crt" \
  --ssl-keyfile  "$USERPROFILE/.office-addin-dev-certs/localhost.key"
```

The server exposes:

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Readiness probe |
| GET | `/providers` | List available LLM providers |
| POST | `/chat` | Chat endpoint. Send `Accept: application/x-ndjson` for streaming |
| GET | `/mcp/servers` | List configured MCP servers |
| POST | `/mcp/servers` | Add a new MCP server |
| PATCH | `/mcp/servers/{id}` | Update an MCP server (enable/disable) |
| DELETE | `/mcp/servers/{id}` | Remove an MCP server |
| POST | `/mcp/servers/{id}/refresh` | Reload tools from an MCP server |

### Frontend setup

```bash
cd frontend
npm install
npm run dev:cert      # install trusted HTTPS cert (run once)
npm start             # serve taskpane at https://localhost:3000
```

To override the backend URL or default provider:

```bash
API_BASE_URL="https://localhost:8000" DEFAULT_PROVIDER="anthropic" npm start
```

Other useful scripts:

| Script | Purpose |
|--------|---------|
| `npm run build` | Production webpack bundle |
| `npm run lint` | ESLint (ts, tsx) |
| `npm run typecheck` | `tsc --noEmit` |
| `npm test` | Run Vitest test suite |
| `npm run test:watch` | Vitest in watch mode |
| `npm run test:coverage` | Vitest with coverage |

### Sideload in Excel

1. Ensure the backend (`https://localhost:8000`) and frontend (`https://localhost:3000`) are both running over HTTPS.
2. In Excel, go to **Insert → My Add-ins → Upload My Add-in** and select the `manifest.xml` at the repo root.
3. The taskpane button will appear under the **MyExcelCompanion** tab in the ribbon.

### Using the chatbot

- Select any cells you want to include as context before sending a prompt.
- If your prompt contains explicit range references (e.g. `Sheet1!B2:C5` or `A1:B3`), those ranges are read instead of the current selection.
- The first 50 rows of the active sheet are sent as a CSV preview on every message, so the LLM has immediate context.
- Responses stream into the panel — thinking steps update with spinners as the LLM works, then the final answer appears.
- Cell updates, formatting, charts, and pivot tables are applied to the workbook automatically.
- Use the settings gear to switch providers or manage MCP servers.
- Click **New Chat** to reset the conversation.

### Using the `=ASKAI` formula

The add-in also provides a custom Excel function that queries the LLM directly from a cell formula:

```
=MYEXCELCOMPANION.ASKAI(query, [range1], [range2], ...)
```

- **query** — The question or instruction to send to the AI (required)
- **range1, range2, ...** — Optional cell ranges to include as context (variadic)

The function uses the same LLM backend and whichever provider is currently selected in the taskpane.

**Examples:**
- `=MYEXCELCOMPANION.ASKAI("What is the capital of France?")` — simple question
- `=MYEXCELCOMPANION.ASKAI("Summarize this data", A1:B20)` — with cell data context
- `=MYEXCELCOMPANION.ASKAI("Compare revenue vs cost", A1:A10, C1:C10)` — multiple ranges

**Status:** While the AI processes your query, the cell shows `#BUSY!` until the final answer is ready.

**Spill behavior:** Multi-line or tabular answers spill into adjacent cells as a 2D array. Tab-separated responses produce NxM grids; plain multi-line responses produce Nx1 vertical spills.

**Recalculation:** Results are cached per unique query+data combination. If input data changes, the cached result is returned instantly (no API call). To force a fresh API call, press **F2+Enter** on the cell or **Ctrl+Shift+F9** to recalculate all ASKAI cells. You can also click "Clear AI Cache" in the taskpane to wipe the cache entirely.

### Configuration

All backend settings use the `COPILOT_` prefix (via pydantic-settings). Key variables:

| Variable | Purpose |
|---|---|
| `COPILOT_ANTHROPIC_API_KEY` | Anthropic API key |
| `COPILOT_ANTHROPIC_MODEL` | Default: `claude-3-5-sonnet-20240620` |
| `COPILOT_OPENAI_API_KEY` | OpenAI API key |
| `COPILOT_OPENAI_MODEL` | Default: `gpt-4o-mini` |
| `COPILOT_REQUEST_TIMEOUT_SECONDS` | Overall request timeout |

### Docker (optional)

Run both services in containers using Docker Compose:

```bash
cp backend/.env.example backend/.env   # then edit with your API keys
docker-compose up --build
```

Services will be available at:
- Backend: https://localhost:8000
- Frontend: https://localhost:3000

Both services use self-signed certificates. Your browser will show a security warning on first access — accept it to proceed.

To stop: `docker-compose down`. To also remove persisted MCP server data: `docker-compose down -v`.

### Packaging (optional)

For production, build the frontend and host the `dist/` output:

```bash
cd frontend
npm run build
```

Update `manifest.xml` to point `SourceLocation` and icon URLs to the deployed host.

### Troubleshooting

- If Excel reports certificate errors, rerun `npm run dev:cert` and restart Excel.
- Backend logs surface provider misconfiguration (missing API keys, timeouts).
- Use browser dev tools (F12) on the taskpane window for frontend diagnostics.
