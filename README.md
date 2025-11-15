## Workbook Copilot

Chat with your Excel workbooks through an Office taskpane add-in backed by a LangGraph orchestrator.

### Prerequisites

- Excel desktop with sideloading enabled (Office 365 recommended)
- Node.js 18+
- npm 9+
- Python 3.11+
- OpenSSL (for HTTPS dev certificate) – install via `npm install -g office-addin-dev-certs` if missing

### Backend setup

```bash
cd excel_addin_simple/backend
python -m venv .venv
.venv\Scripts\activate            # Windows PowerShell
pip install -r requirements.txt
copy .env.example .env            # fill in API keys as needed
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000 ^
  --ssl-certfile "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" ^
  --ssl-keyfile  "%USERPROFILE%\.office-addin-dev-certs\localhost.key"
```

The server exposes:

- `GET /health` – readiness probe
- `GET /providers` – configured providers
- `POST /chat` – chat endpoint used by the add-in. Send `Accept: application/x-ndjson`
  to receive streaming message/cell-update events as they are produced.

### Frontend setup

```bash
cd excel_addin_simple/frontend
npm install
npm run dev:cert                  # creates trusted HTTPS cert (run once)
npm start                         # serves taskpane at https://localhost:3000
```

To change the backend URL or default provider, set environment variables before invoking `npm start`:

```bash
$env:API_BASE_URL="https://localhost:8000"; npm start
```

### Sideload in Excel

1. Ensure the backend (`https://localhost:8000`) and frontend (`https://localhost:3000`) are both running over HTTPS.
2. In Excel, open **File → Options → Trust Center → Trust Center Settings… → Trusted Add-in Catalogs** and add the folder containing `manifest.xml`, or sideload manually via **Insert → My Add-ins → Upload My Add-in**.
3. Select `excel_addin_simple/manifest.xml`.
4. The taskpane button **Workbook Copilot** will appear under the **Copilot** tab.

### Using the chatbot

- Select any cells you want to include as context before sending a prompt.
- If your prompt contains explicit references (e.g. `Sheet1!B2:C5` or `A1:B3`), those ranges are fetched and sent instead of the current selection.
- Responses stream into the panel when the backend emits NDJSON events, so
  thought/step/final bubbles appear as they are produced.
- Formatting instructions (`format_updates`) such as fill color, font color, bold/italic are applied to the specified ranges.
- Use the settings gear to switch between configured providers (`mock`, `openai`, `anthropic`).

### Packaging (optional)

For production, build the frontend and host the `dist/` output:

```bash
npm run build
```

Update `manifest.xml` to point `SourceLocation` and icon URLs to the deployed host.

### Troubleshooting

- If Excel reports certificate errors, rerun `npm run dev:cert` and restart Excel.
- Backend logs surface provider misconfiguration (missing API keys, timeouts).
- Use browser dev tools (F12) on the taskpane window for frontend diagnostics.

