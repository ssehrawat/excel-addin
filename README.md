## Workbook Copilot

An Excel Office Add-in (taskpane) that lets users chat with their workbook data. The frontend reads cell selections via Office.js, POSTs them to a FastAPI backend, which routes through an LLM provider (mock / OpenAI / Anthropic) and optionally calls external MCP tool servers. Responses stream back as NDJSON and are applied to Excel — cell values, formatting, charts, and pivot tables.

### Features

- **Streaming chat** — thinking steps show progress with live spinners, then the final response appears
- **Cell updates** — LLM can write values into cells (replace or append mode)
- **Format updates** — fill color, font color, bold, italic, number format, and borders applied to specified ranges
- **Chart inserts** — bar, column, line, scatter, bubble with axis titles
- **Pivot table inserts** — rows, columns, values, filters inferred from sheet context
- **MCP tool servers** — connect external data sources (databases, APIs) that the LLM can query server-side
- **Excel read tools** — LLM can request additional workbook data mid-conversation (up to 3 tool-call rounds per message)
- **Multi-provider** — switch between mock, OpenAI, and Anthropic at runtime

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
# create backend/.env with your API keys (see Configuration section below)

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

### Configuration

All backend settings use the `COPILOT_` prefix (via pydantic-settings). Key variables:

| Variable | Purpose |
|---|---|
| `COPILOT_ANTHROPIC_API_KEY` | Anthropic API key |
| `COPILOT_ANTHROPIC_MODEL` | Default: `claude-3-5-sonnet-20240620` |
| `COPILOT_OPENAI_API_KEY` | OpenAI API key |
| `COPILOT_OPENAI_MODEL` | Default: `gpt-4o-mini` |
| `COPILOT_MOCK_PROVIDER_ENABLED` | `true`/`false` — enable mock provider |
| `COPILOT_REQUEST_TIMEOUT_SECONDS` | Overall request timeout |

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
