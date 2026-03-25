from __future__ import annotations

import asyncio
import json
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from uuid import uuid4
from typing import (
    Any,
    AsyncIterator,
    Dict,
    Iterable,
    List,
    Optional,
    Protocol,
    TypedDict,
)

from fastapi import HTTPException
from pydantic import ValidationError

from .config import get_settings
from .schemas import (
    AggregationFunction,
    ChartInsert,
    ChartSeriesBy,
    ChatMessage,
    ChatRequest,
    ChatResponse,
    CellUpdate,
    FormatUpdate,
    MessageKind,
    MessageRole,
    PivotTableDataField,
    PivotTableInsert,
    Telemetry,
    WorkbookToolCall,
)

logger = logging.getLogger(__name__)


STREAM_MESSAGE_DELAY_SECONDS = 0.35
STREAM_MESSAGE_CHUNK_DELAY_SECONDS = 0.08
STREAM_MESSAGE_CHUNK_SIZE = 48

CHART_TYPE_ALIASES: Dict[str, str] = {
    "scatter": "XYScatter",
    "scatterplot": "XYScatter",
    "scatterchart": "XYScatter",
    "scattermarkers": "XYScatter",
    "xyscatter": "XYScatter",
    "xyscattermarkers": "XYScatter",
    "scatterlines": "XYScatterLines",
    "scatterline": "XYScatterLines",
    "xyscatterlines": "XYScatterLines",
    "scatterlinenomarkers": "XYScatterLinesNoMarkers",
    "scatterlinesnomarkers": "XYScatterLinesNoMarkers",
    "xyscatterlinesnomarkers": "XYScatterLinesNoMarkers",
    "bubble": "Bubble",
    "line": "LineMarkers",
    "column": "ColumnClustered",
    "bar": "BarClustered",
}


@dataclass
class MCPToolEntry:
    """A live MCP tool available for the current request.

    Instances are created by the orchestrator's per-request health check and
    injected into the system prompt so the LLM can select and call them.

    Attributes:
        namespaced_name: Unique tool identifier used in LLM tool_call responses.
            Format: ``mcp__{server_id}__{tool_name}``.
        server_id: ID of the owning MCPServerRecord.
        server_name: Human-readable server name shown in status events.
        description: Tool description injected verbatim into the system prompt.
        input_schema: JSON Schema dict describing the tool's accepted arguments.
    """

    namespaced_name: str
    server_id: str
    server_name: str
    description: str
    input_schema: Dict[str, Any]


def build_system_prompt(mcp_tools: List[MCPToolEntry]) -> str:
    """Build the LLM system prompt, dynamically including live MCP tools.

    Generates a complete system prompt that lists all available tools
    (always Excel tools, plus any live MCP tools for this request). The
    DECISION RULES section explicitly instructs the LLM to use tools only
    when they directly help answer the query.

    Args:
        mcp_tools: Live MCP tools from the orchestrator's health check.
            When empty, the ``EXTERNAL MCP TOOLS`` section is omitted
            entirely so the LLM does not attempt MCP calls.

    Returns:
        Complete system prompt string ready to send as the ``system``
        role message to the LLM.
    """
    parts: List[str] = [
        "You are Workbook Copilot, an AI assistant for Microsoft Excel.\n\n"
        "You receive structured context:\n"
        "- workbook_metadata: Filename, all sheets with row/column dimensions and "
        "column headers (first-row values for every sheet via columnHeaders)\n"
        "- user_context: Active sheet name, currently selected range address\n"
        "- active_sheet_preview: First rows of active sheet as CSV. The first row "
        "contains column-letter markers ([A],[B],[C],...) mapping positions to "
        "Excel column letters.\n"
        "- selection: Values of cells the user has explicitly selected\n\n",
        "EXCEL TOOLS (request when you need more data):\n"
        "- get_xl_cell_ranges: Read specific ranges with formulas and formatting. "
        "Args: {ranges: [\"A1:C10\"], sheetName: \"Sheet1\"}\n"
        "- get_xl_range_as_csv: Read large data ranges as CSV. "
        "Args: {sheetName: \"Sheet1\", range: \"A1:D200\", maxRows: 200, offset: 0}\n"
        "- xl_search_data: Find values across sheets. "
        "Args: {query: \"...\", sheetName?: \"Sheet1\", caseSensitive?: false, "
        "regex?: false, matchEntireCell?: false}\n"
        "- get_all_xl_objects: List charts/tables/pivot tables. "
        "Args: {sheetName?: \"Sheet1\", objectType?: \"chart\"|\"table\"|\"pivot\"}\n"
        "- execute_xl_office_js: Run Office.js code snippet (reads, sorts, filters, "
        "formatting, and other mutations). The code runs inside Excel.run() — "
        "`context` (RequestContext) and `Excel` are already available. "
        "Do NOT wrap your code in Excel.run(); use `context` directly. "
        "Args: {code: \"...\"}\n"
        "  Sort example (ALWAYS read headers to find column indexes — never hardcode them):\n"
        "  const sheet = context.workbook.worksheets.getItem('SheetName');\n"
        "  const used = sheet.getUsedRange();\n"
        "  const headerRow = used.getRow(0); headerRow.load('values');\n"
        "  used.load(['rowCount','columnCount']); await context.sync();\n"
        "  const headers = headerRow.values[0].map(h => String(h).trim());\n"
        "  const col = (name) => { const i = headers.indexOf(name); "
        "if (i === -1) throw new Error('Column not found: ' + name); return i; };\n"
        "  const data = sheet.getRangeByIndexes(1, 0, used.rowCount - 1, used.columnCount);\n"
        "  data.sort.apply([{ key: col('ColA'), ascending: true }, "
        "{ key: col('ColB'), ascending: false }]);\n"
        "  await context.sync();\n\n",
    ]

    if mcp_tools:
        mcp_lines = [
            "EXTERNAL MCP TOOLS (call when they provide relevant external data):\n"
        ]
        for tool in mcp_tools:
            schema_str = json.dumps(tool.input_schema)
            mcp_lines.append(
                f"- {tool.namespaced_name}: [{tool.server_name}] {tool.description}\n"
                f"  Args schema: {schema_str}"
            )
        parts.append("\n".join(mcp_lines) + "\n\n")

    parts.append(
        "CLARIFICATION:\n"
        "- STRONGLY prefer acting over asking. Make reasonable assumptions using the "
        "workbook metadata, sheet preview, selection context, and column headers. "
        "For example: if only one sheet has data, use it; if the user says "
        "\"create a chart\" without specifying type, pick the most appropriate one; "
        "if headers are visible, infer the relevant columns.\n"
        "- Only ask a clarifying question when it is genuinely IMPOSSIBLE to proceed "
        "— e.g. the user says \"analyze\" with zero context about what to analyze "
        "and the workbook has multiple unrelated sheets with no obvious target. "
        "Even then, prefer making a best-effort attempt with a note about your "
        "assumptions over asking.\n"
        "- When you must ask, return it in the \"answer\" field as natural language. "
        "Do NOT use needs_data for clarification.\n\n"
    )

    parts.append(
        "DECISION RULES:\n"
        "- If the question can be answered from provided context → answer directly\n"
        "- If you need more data → return a needs_data response "
        "(one tool call per response, max 3 rounds total)\n"
        "- Prefer get_xl_range_as_csv for data analysis; "
        "get_xl_cell_ranges for formula/format inspection\n"
        "- For large sheets (> 200 rows), use offset pagination or xl_search_data first\n"
        "- Use MCP tools only when they provide relevant external data for the query. "
        "Do NOT use tools unless they directly help answer the user's question.\n"
        "- After a tool returns a success result (including {\"success\": true}), "
        "produce a final answer confirming the operation. Do NOT re-call the "
        "same tool — the operation already completed.\n\n"
    )

    option_b = (
        "{\"needs_data\": true, \"tool_call\": "
        "{\"tool\": \"get_xl_range_as_csv\", \"args\": {\"sheetName\": \"Sheet1\", "
        "\"range\": \"A1:D200\", \"maxRows\": 200}}}"
    )
    if mcp_tools:
        option_b += (
            "\n  — or MCP tool:\n"
            "{\"needs_data\": true, \"tool_call\": "
            "{\"tool\": \"mcp__<server_id>__<tool_name>\", \"args\": {...}}}"
        )

    parts.append(
        "RESPONSE FORMAT — Option A (direct answer):\n"
        "{\"answer\": \"...\", \"cell_updates\": [...], \"format_updates\": [...], "
        "\"chart_inserts\": [...], \"pivot_table_inserts\": [...]}\n\n"
        f"RESPONSE FORMAT — Option B (needs more data):\n{option_b}\n\n"
    )

    parts.append(
        "RULES FOR CELL UPDATES:\n"
        "Every cell update MUST include: `address` (A1 notation, include worksheet like "
        "'Sheet1!E7:E9' when known), `mode` ('replace' or 'append'), and `values` "
        "(two-dimensional array; wrap single rows like [[\"Header\"], [123]]).\n"
        "**FORMULA RULE (CRITICAL)**: ALWAYS use Excel formulas instead of hardcoded computed values. "
        "Examples: =SUM(A1:A10) not 5432, =AVERAGE(B2:B50) not 78.5, ='Sector Allocation'!B2 not \"42%\". "
        "Only use static values for text labels/headers or data not derived from existing cells. "
        "When writing formulas, ALWAYS cross-check column letters against the column markers in "
        "active_sheet_preview or the columnHeaders in workbook_metadata. Never guess column "
        "positions — verify from provided context.\n"
        "PLACEMENT RULE: When placing new data (cell_updates, chart_inserts), choose an address "
        "in an empty area — below or to the right of existing data — unless the user explicitly "
        "specifies where. Do NOT overwrite existing cell values. Check active_sheet_preview to "
        "identify the used range and place output beyond it. For charts, omit topLeftCell and "
        "the system will auto-position in empty space.\n"
        "HEADER ROW RULE: The first row of data is almost always a header row — verify "
        "by checking columnHeaders in workbook_metadata and the first data row in "
        "active_sheet_preview. When sorting, reordering, or rearranging rows, NEVER "
        "include the header row in the operation. Keep it fixed in row 1 and only "
        "sort/reorder rows 2 onwards. Always include the original header row unchanged "
        "at the top of your cell_updates. Only modify the header row if the user "
        "explicitly asks to change headers.\n"
        "RULES FOR FORMAT UPDATES: Only include when user EXPLICITLY requests formatting. "
        "Each format update has: `address` (A1 notation with sheet), and any combination of: "
        "`fillColor` (hex like '#4472C4'), `fontColor` (hex like '#FFFFFF'), `bold` (true/false), "
        "`italic` (true/false), `numberFormat` (Excel format string), `borderColor` (hex), "
        "`borderStyle` ('Continuous'|'Dash'|'None'|...), `borderWeight` ('Thin'|'Medium'|'Thick'). "
        "Only include non-null fields. Example: {\"address\":\"Sheet1!A1:D1\",\"fillColor\":\"#4472C4\",\"fontColor\":\"#FFFFFF\",\"bold\":true}\n"
        "CHART SOURCE RULE: When the user has selected specific ranges, use those exact ranges "
        "as sourceAddress. For non-contiguous selections, use comma-separated format "
        "(e.g. 'A1:A13,G1:G13'). Do NOT expand into a single contiguous block that includes "
        "unwanted columns.\n"
        "RULES FOR CHART INSERTS: Only include when user EXPLICITLY requests a chart. "
        "Every chart insert MUST include `chartType` (Excel.ChartType), `sourceAddress`, "
        "and a descriptive `title`, plus `xAxisTitle` and `yAxisTitle` for the axes. "
        "Use official identifiers like 'XYScatter', 'ColumnClustered', 'LineMarkers'.\n"
        "MODIFYING EXISTING CHARTS: To change properties on an existing chart (title, "
        "axis titles, legend, formatting), return a needs_data response using "
        "execute_xl_office_js with Office.js code. Example for setting chart and axis "
        "titles:\n"
        '{"needs_data": true, "tool_call": {"tool": "execute_xl_office_js", "args": '
        '{"code": "const chart = context.workbook.worksheets.getItem(\'Sheet1\')'
        ".charts.getItem('Chart 1'); chart.title.text = 'My Title'; "
        "chart.title.visible = true; chart.axes.categoryAxis.title.text = 'X Axis'; "
        "chart.axes.categoryAxis.title.visible = true; chart.axes.valueAxis.title.text = 'Y Axis'; "
        "chart.axes.valueAxis.title.visible = true; await context.sync();"
        '"}}}\n'
        "RULES FOR PIVOT TABLE INSERTS: Only include when user EXPLICITLY requests a pivot table. "
        "Every pivot table insert MUST include `name` (unique identifier) and `sourceAddress` "
        "(data range in A1 notation). By default the pivot table is placed on the active sheet "
        "in an empty area next to the data — do NOT include `destinationAddress` or "
        "`destinationWorksheet` unless the user explicitly specifies where to put it. "
        "If the user asks to place it on a new sheet, set `destinationWorksheet` to a descriptive name. "
        "Infer the pivot structure (rows, columns, values, filters) from the user's request "
        "and the sheet preview data. Do NOT ask the user to specify fields that can be "
        "reasonably inferred from context. Use the column headers visible in the sheet preview "
        "to determine sourceAddress and map fields to rows/columns/values as the user's "
        "description implies. "
        "Include `rows`, `columns`, `values`, and `filters` arrays "
        "to define the pivot table structure. Each entry in `values` must have `name` (column "
        "header) and optionally `summarizeBy` (sum, count, average, max, min, product, "
        "countNumbers, standardDeviation, standardDeviationP, variance, varianceP). "
        "Optional: `destinationWorksheet` (sheet name for the pivot output) and "
        "`destinationAddress` (cell where the pivot starts, e.g. \"E1\" or \"Sheet2!E1\"). "
        "If omitted, the pivot is placed automatically to the right of used data on the source sheet. "
        "Example:\n"
        '{"pivot_table_inserts": [{"name": "SalesPivot", "sourceAddress": "A1:D100", '
        '"destinationWorksheet": "PivotSheet", "destinationAddress": "A1", '
        '"rows": ["Region"], "columns": ["Quarter"], '
        '"values": [{"name": "Revenue", "summarizeBy": "sum"}], '
        '"filters": ["Category"]}]}\n'
        "Return strictly valid JSON with fully expanded arrays — no code or list comprehensions."
    )

    return "".join(parts)


def _timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


def _message_id() -> str:
    return uuid4().hex


def _ensure_iterable(value: Any) -> Iterable[Any]:
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        return value
    return [value]


def _is_response_format_error(error: Exception) -> bool:
    text = str(error).lower()
    if "response_format" not in text:
        return False
    return any(
        keyword in text
        for keyword in ["not supported", "unsupported", "only available"]
    )


@dataclass
class ProviderResult:
    """Result returned by a provider's generate() call.

    Attributes:
        messages: Chat messages to include in the response.
        cell_updates: Excel cell writes.
        format_updates: Excel formatting changes.
        chart_inserts: Excel chart creations.
        telemetry: Performance data.
        tool_call_required: Set when the LLM needs more Excel data before answering.
    """

    messages: List[ChatMessage]
    cell_updates: List[CellUpdate] = field(default_factory=list)
    format_updates: List[FormatUpdate] = field(default_factory=list)
    chart_inserts: List[ChartInsert] = field(default_factory=list)
    pivot_table_inserts: List[PivotTableInsert] = field(default_factory=list)
    telemetry: Optional[Telemetry] = None
    tool_call_required: Optional[WorkbookToolCall] = None

    def to_response(self) -> ChatResponse:
        """Convert to a ChatResponse for the non-streaming path."""
        return ChatResponse(
            messages=self.messages,
            cell_updates=self.cell_updates,
            format_updates=self.format_updates,
            chart_inserts=self.chart_inserts,
            pivot_table_inserts=self.pivot_table_inserts,
            telemetry=self.telemetry,
        )


class ProviderStreamEvent(TypedDict, total=False):
    type: str
    payload: Any


def _chunk_text(text: str, size: int = STREAM_MESSAGE_CHUNK_SIZE) -> Iterable[str]:
    if not text:
        return []
    return [text[i : i + size] for i in range(0, len(text), size)]


async def _stream_result(result: ProviderResult) -> AsyncIterator[ProviderStreamEvent]:
    """Stream a ProviderResult as NDJSON events.

    Only FINAL and MESSAGE kind messages are streamed as visible bubbles.
    THOUGHT/STEP/CONTEXT messages are silently dropped (they stay internal).
    tool_call_required causes a ``tool_call_required`` event and early return.
    """
    if result.tool_call_required:
        call = result.tool_call_required
        yield {
            "type": "tool_call_required",
            "payload": [call.model_dump()],
        }
        return

    visible_messages = [
        m
        for m in result.messages
        if m.kind in (MessageKind.FINAL, MessageKind.MESSAGE)
    ]

    for index, message in enumerate(visible_messages):
        if index > 0:
            await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)

        payload = message.model_dump(by_alias=True)
        payload["content"] = ""
        yield {"type": "message_start", "payload": payload}

        for chunk in _chunk_text(message.content):
            await asyncio.sleep(STREAM_MESSAGE_CHUNK_DELAY_SECONDS)
            yield {
                "type": "message_delta",
                "payload": {"id": message.id, "delta": chunk},
            }

        await asyncio.sleep(STREAM_MESSAGE_CHUNK_DELAY_SECONDS)
        yield {"type": "message_done", "payload": message.model_dump(by_alias=True)}

    if result.cell_updates:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "cell_updates",
            "payload": [item.model_dump(by_alias=True) for item in result.cell_updates],
        }
    if result.format_updates:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "format_updates",
            "payload": [
                item.model_dump(by_alias=True) for item in result.format_updates
            ],
        }
    if result.chart_inserts:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "chart_inserts",
            "payload": [
                item.model_dump(by_alias=True) for item in result.chart_inserts
            ],
        }
    if result.pivot_table_inserts:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "pivot_table_inserts",
            "payload": [
                item.model_dump(by_alias=True)
                for item in result.pivot_table_inserts
            ],
        }
    if result.telemetry:
        await asyncio.sleep(STREAM_MESSAGE_DELAY_SECONDS)
        yield {
            "type": "telemetry",
            "payload": result.telemetry.model_dump(by_alias=True),
        }


class LLMProvider(Protocol):
    id: str
    label: str
    description: str
    requires_key: bool

    async def generate(self, request: ChatRequest) -> ProviderResult: ...

    async def stream(
        self, request: ChatRequest
    ) -> AsyncIterator[ProviderStreamEvent]: ...

    async def stream_result(
        self, result: ProviderResult
    ) -> AsyncIterator[ProviderStreamEvent]: ...


class MockProvider:
    id = "mock"
    label = "Mock provider"
    description = (
        "Deterministic, local responses for development without external API calls."
    )
    requires_key = False

    async def generate(self, request: ChatRequest) -> ProviderResult:
        selection_preview = [
            f"{sel.address} = {sel.values}" for sel in request.selection[:3]
        ]
        messages = []

        if selection_preview:
            messages.append(
                ChatMessage(
                    id=_message_id(),
                    role=MessageRole.ASSISTANT,
                    kind=MessageKind.CONTEXT,
                    content="Selected data:\n" + "\n".join(selection_preview),
                    created_at=_timestamp(),
                )
            )

        messages.append(
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.FINAL,
                content=(
                    f"(mock) Answer to: '{request.prompt}'. "
                    "Configure a real provider in the backend environment to get real answers."
                ),
                created_at=_timestamp(),
            )
        )
        cell_updates: List[CellUpdate] = []
        format_updates: List[FormatUpdate] = []
        if request.selection:
            first_address = request.selection[0].address
            cell_updates.append(
                CellUpdate(
                    address=first_address,
                    values=[["(mock) Result written back to Excel"]],
                )
            )
            if "color" in request.prompt.lower():
                format_updates.append(
                    FormatUpdate(
                        address=first_address,
                        worksheet=request.selection[0].worksheet,
                        fill_color="#FFF2CC",
                    )
                )

        chart_inserts: List[ChartInsert] = []
        if request.selection and any(
            keyword in request.prompt.lower() for keyword in ("chart", "graph")
        ):
            chart_inserts.append(
                ChartInsert(
                    chart_type="columnClustered",
                    source_address=request.selection[0].address,
                    source_worksheet=request.selection[0].worksheet,
                    title="(mock) Selection overview",
                    series_by=ChartSeriesBy.AUTO,
                )
            )

        pivot_table_inserts: List[PivotTableInsert] = []
        if request.selection and "pivot" in request.prompt.lower():
            pivot_table_inserts.append(
                PivotTableInsert(
                    name="MockPivotTable",
                    source_address=request.selection[0].address,
                    source_worksheet=request.selection[0].worksheet,
                    rows=["Category"],
                    values=[
                        PivotTableDataField(
                            name="Amount",
                            summarize_by=AggregationFunction.SUM,
                        )
                    ],
                )
            )

        telemetry = Telemetry(provider=self.id, model="mock-local", latency_ms=5)

        return ProviderResult(
            messages=messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
            chart_inserts=chart_inserts,
            pivot_table_inserts=pivot_table_inserts,
            telemetry=telemetry,
        )

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        result = await self.generate(request)
        async for event in _stream_result(result):
            yield event

    async def stream_result(
        self, result: ProviderResult
    ) -> AsyncIterator[ProviderStreamEvent]:
        """Stream a completed ProviderResult as NDJSON events.

        Called by the orchestrator's ReAct loop after ``generate()`` returns
        a final answer (i.e. ``result.tool_call_required is None``).

        Args:
            result: A completed ProviderResult with no pending tool call.

        Yields:
            ProviderStreamEvent dicts (message_start, message_delta,
            message_done, cell_updates, format_updates,
            chart_inserts, telemetry).
        """
        async for event in _stream_result(result):
            yield event


class OpenAIProvider:
    id = "openai"
    label = "OpenAI"
    description = "Leverage OpenAI models such as GPT-4o."
    requires_key = True

    def __init__(self, mcp_tools: Optional[List[MCPToolEntry]] = None) -> None:
        settings = get_settings()
        self.api_key = settings.openai_api_key
        self.model = settings.openai_model
        self.temperature = settings.openai_temperature
        if not self.api_key:
            raise HTTPException(
                status_code=503,
                detail="OpenAI API key is not configured on the server.",
            )
        try:
            from langchain_openai import ChatOpenAI  # type: ignore
        except ImportError as exc:  # pragma: no cover
            raise HTTPException(
                status_code=500,
                detail="langchain-openai is not installed on the server.",
            ) from exc

        self._settings = settings
        self._system_prompt = build_system_prompt(mcp_tools or [])
        self._json_mode_enabled = True
        self._client = self._instantiate_openai_client(json_mode=True)

    def _instantiate_openai_client(self, json_mode: bool):  # type: ignore[return-type]
        from langchain_openai import ChatOpenAI  # type: ignore

        client_kwargs = {
            "model": self.model,
            "api_key": self.api_key,
            "timeout": self._settings.request_timeout_seconds,
        }
        if self.temperature is not None:
            client_kwargs["temperature"] = self.temperature
        if json_mode:
            client_kwargs["model_kwargs"] = {"response_format": {"type": "json_object"}}
        client = ChatOpenAI(**client_kwargs)
        self._json_mode_enabled = json_mode
        return client

    async def generate(self, request: ChatRequest) -> ProviderResult:
        messages = [
            {"role": "system", "content": self._system_prompt},
            {"role": "user", "content": build_prompt_payload(request)},
        ]

        try:
            result = await asyncio.to_thread(self._client.invoke, messages)
        except Exception as exc:
            if self._json_mode_enabled and _is_response_format_error(exc):
                logger.warning(
                    "OpenAI model '%s' does not support JSON mode. Falling back to lax parsing.",
                    self.model,
                )
                self._client = self._instantiate_openai_client(json_mode=False)
                result = await asyncio.to_thread(self._client.invoke, messages)
            else:
                raise

        parsed = parse_structured_response(result.content)

        telemetry = Telemetry(
            provider=self.id,
            model=self.model,
            raw={"openai_response": getattr(result, "response_metadata", {})},
        )

        # Handle needs_data response (LLM requests more data via tool call)
        if parsed.get("needs_data") and parsed.get("tool_call"):
            tool_call = _build_tool_call(parsed["tool_call"])
            if tool_call:
                return ProviderResult(
                    messages=[],
                    tool_call_required=tool_call,
                    telemetry=telemetry,
                )

        provider_messages = assemble_messages_from_payload(parsed, request.prompt)
        cell_updates = build_cell_updates(parsed.get("cell_updates", []))
        format_updates = build_format_updates(parsed.get("format_updates", []))
        chart_inserts = build_chart_inserts(parsed.get("chart_inserts", []))
        pivot_table_inserts = build_pivot_table_inserts(
            parsed.get("pivot_table_inserts", [])
        )

        return ProviderResult(
            messages=provider_messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
            chart_inserts=chart_inserts,
            pivot_table_inserts=pivot_table_inserts,
            telemetry=telemetry,
        )

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        result = await self.generate(request)
        async for event in _stream_result(result):
            yield event

    async def stream_result(
        self, result: ProviderResult
    ) -> AsyncIterator[ProviderStreamEvent]:
        """Stream a completed ProviderResult as NDJSON events.

        Called by the orchestrator's ReAct loop after ``generate()`` returns
        a final answer (i.e. ``result.tool_call_required is None``).

        Args:
            result: A completed ProviderResult with no pending tool call.

        Yields:
            ProviderStreamEvent dicts (message_start, message_delta,
            message_done, cell_updates, format_updates,
            chart_inserts, pivot_table_inserts, telemetry).
        """
        async for event in _stream_result(result):
            yield event


class AnthropicProvider:
    id = "anthropic"
    label = "Anthropic"
    description = "Access Claude models via Anthropic API."
    requires_key = True

    def __init__(self, mcp_tools: Optional[List[MCPToolEntry]] = None) -> None:
        settings = get_settings()
        self.api_key = settings.anthropic_api_key
        self.model = settings.anthropic_model
        self.temperature = settings.anthropic_temperature
        if not self.api_key:
            raise HTTPException(
                status_code=503,
                detail="Anthropic API key is not configured on the server.",
            )
        try:
            from langchain_anthropic import ChatAnthropic  # type: ignore
        except ImportError as exc:  # pragma: no cover
            raise HTTPException(
                status_code=500,
                detail="langchain-anthropic is not installed on the server.",
            ) from exc

        self._settings = settings
        self._system_prompt = build_system_prompt(mcp_tools or [])
        self._json_mode_enabled = True
        self._client = self._instantiate_anthropic_client(json_mode=True)

    def _instantiate_anthropic_client(self, json_mode: bool):  # type: ignore[return-type]
        from langchain_anthropic import ChatAnthropic  # type: ignore

        client_kwargs = {
            "model": self.model,
            "anthropic_api_key": self.api_key,
            "timeout": self._settings.request_timeout_seconds,
        }
        if self.temperature is not None:
            client_kwargs["temperature"] = self.temperature
        if json_mode:
            client_kwargs.setdefault("model_kwargs", {})["response_format"] = {
                "type": "json_object"
            }
        client = ChatAnthropic(**client_kwargs)
        self._json_mode_enabled = json_mode
        return client

    async def generate(self, request: ChatRequest) -> ProviderResult:
        messages = [
            {"role": "system", "content": self._system_prompt},
            {"role": "user", "content": build_prompt_payload(request)},
        ]

        try:
            result = await asyncio.to_thread(self._client.invoke, messages)
        except Exception as exc:
            if self._json_mode_enabled and _is_response_format_error(exc):
                logger.warning(
                    "Anthropic model '%s' does not support JSON mode. Falling back to lax parsing.",
                    self.model,
                )
                self._client = self._instantiate_anthropic_client(json_mode=False)
                result = await asyncio.to_thread(self._client.invoke, messages)
            else:
                raise

        parsed = parse_structured_response(result.content)

        telemetry = Telemetry(
            provider=self.id,
            model=self.model,
            raw={"anthropic_response": getattr(result, "response_metadata", {})},
        )

        # Handle needs_data response (LLM requests more data via tool call)
        if parsed.get("needs_data") and parsed.get("tool_call"):
            tool_call = _build_tool_call(parsed["tool_call"])
            if tool_call:
                return ProviderResult(
                    messages=[],
                    tool_call_required=tool_call,
                    telemetry=telemetry,
                )

        provider_messages = assemble_messages_from_payload(parsed, request.prompt)
        cell_updates = build_cell_updates(parsed.get("cell_updates", []))
        format_updates = build_format_updates(parsed.get("format_updates", []))
        chart_inserts = build_chart_inserts(parsed.get("chart_inserts", []))
        pivot_table_inserts = build_pivot_table_inserts(
            parsed.get("pivot_table_inserts", [])
        )

        return ProviderResult(
            messages=provider_messages,
            cell_updates=cell_updates,
            format_updates=format_updates,
            chart_inserts=chart_inserts,
            pivot_table_inserts=pivot_table_inserts,
            telemetry=telemetry,
        )

    async def stream(self, request: ChatRequest) -> AsyncIterator[ProviderStreamEvent]:
        result = await self.generate(request)
        async for event in _stream_result(result):
            yield event

    async def stream_result(
        self, result: ProviderResult
    ) -> AsyncIterator[ProviderStreamEvent]:
        """Stream a completed ProviderResult as NDJSON events.

        Called by the orchestrator's ReAct loop after ``generate()`` returns
        a final answer (i.e. ``result.tool_call_required is None``).

        Args:
            result: A completed ProviderResult with no pending tool call.

        Yields:
            ProviderStreamEvent dicts (message_start, message_delta,
            message_done, cell_updates, format_updates,
            chart_inserts, pivot_table_inserts, telemetry).
        """
        async for event in _stream_result(result):
            yield event


def build_prompt_payload(request: ChatRequest) -> str:
    """Build the user-turn JSON payload sent to the LLM.

    Includes workbook metadata, user context, active sheet preview, selection
    data, and conversation history.

    Args:
        request: The incoming ChatRequest.

    Returns:
        JSON string to use as the user message content.
    """
    selection_text = "\n".join(
        f"{sel.address}: {sel.values}" for sel in request.selection
    )
    history_messages: List[str] = []
    tool_context_segments: List[str] = []
    for msg in request.messages:
        entry = f"[{msg.role}] {msg.content}"
        if msg.kind == MessageKind.CONTEXT:
            tool_context_segments.append(entry)
        else:
            history_messages.append(entry)

    trimmed_history = "\n".join(history_messages[-6:])
    context_block = "\n".join(tool_context_segments)
    if trimmed_history and context_block:
        history = f"{trimmed_history}\n\n[tool_context]\n{context_block}"
    elif context_block:
        history = context_block
    else:
        history = trimmed_history

    payload: Dict[str, Any] = {
        "conversation_history": history,
        "user_prompt": request.prompt,
        "selection": selection_text,
    }

    # Include workbook context when available
    if request.workbook_metadata:
        payload["workbook_metadata"] = request.workbook_metadata.model_dump(
            by_alias=True
        )
    if request.user_context:
        payload["user_context"] = request.user_context.model_dump(by_alias=True)
    if request.active_sheet_preview:
        payload["active_sheet_preview"] = request.active_sheet_preview

    return json.dumps(payload)


def _build_tool_call(raw: Any) -> Optional[WorkbookToolCall]:
    """Build a WorkbookToolCall from the LLM's needs_data tool_call dict.

    The ``tool`` field may contain an MCP-namespaced tool name in the format
    ``mcp__<server_id>__<tool_name>``. The orchestrator routes based on this
    prefix — no additional schema fields are needed.

    Args:
        raw: The raw ``tool_call`` value from the LLM response.

    Returns:
        A WorkbookToolCall, or None if the data is invalid.
    """
    if not isinstance(raw, dict):
        return None
    tool_name = raw.get("tool")
    if not isinstance(tool_name, str) or not tool_name.strip():
        return None
    args = raw.get("args") or {}
    if not isinstance(args, dict):
        args = {}
    return WorkbookToolCall(id=_message_id(), tool=tool_name.strip(), args=args)


def _normalize_identifier(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.strip().lower().lstrip("xl"))


def _normalize_chart_type(value: Any) -> Optional[str]:
    if not isinstance(value, str):
        return None
    normalized = _normalize_identifier(value)
    if not normalized:
        return None
    if normalized in CHART_TYPE_ALIASES:
        return CHART_TYPE_ALIASES[normalized]
    return value


def parse_structured_response(content: Any) -> Dict[str, Any]:
    """Parse a potentially-JSON LLM response into a dict.

    Args:
        content: Raw LLM output (string or dict).

    Returns:
        Parsed dict.  Falls back to a minimal default structure on failure.
    """
    if isinstance(content, dict):
        return content
    if isinstance(content, str):
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            pass
        # LLM may have returned multiple JSON objects on separate lines
        # (e.g. a needs_data tool call followed by a pre-emptive answer).
        # Parse line-by-line; prefer a needs_data tool call if present.
        first_parsed: Optional[Dict[str, Any]] = None
        for line in content.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(obj, dict):
                if obj.get("needs_data") and obj.get("tool_call"):
                    return obj
                if first_parsed is None:
                    first_parsed = obj
        if first_parsed is not None:
            return first_parsed
        logger.warning("LLM returned non-JSON output: %s", content)
    return {
        "answer": content if isinstance(content, str) else "No answer produced.",
        "cell_updates": [],
        "format_updates": [],
        "chart_inserts": [],
        "pivot_table_inserts": [],
    }


def assemble_messages_from_payload(
    payload: Dict[str, Any], prompt: str
) -> List[ChatMessage]:
    """Build ChatMessage objects from a parsed LLM response dict.

    Only the ``answer`` field produces visible (FINAL) messages.
    Thoughts and steps are intentionally omitted from the message list.

    Args:
        payload: Parsed LLM response dict.
        prompt: Original user prompt (used for fallback answer).

    Returns:
        List of ChatMessage instances.
    """
    messages: List[ChatMessage] = []

    # Only create a FINAL message for the answer
    answer = payload.get("answer")
    if answer and str(answer).strip():
        messages.append(
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.FINAL,
                content=str(answer).strip(),
                created_at=_timestamp(),
            )
        )

    # Ensure we always have a final answer
    if not any(msg.kind == MessageKind.FINAL for msg in messages):
        # If the payload contains mutations the LLM performed an action;
        # use a concise confirmation instead of echoing the prompt.
        has_mutations = any(
            payload.get(k)
            for k in (
                "cell_updates",
                "format_updates",
                "chart_inserts",
                "pivot_table_inserts",
            )
        )
        fallback = "Done." if has_mutations else f"Here is my best effort answer to: {prompt}"
        messages.append(
            ChatMessage(
                id=_message_id(),
                role=MessageRole.ASSISTANT,
                kind=MessageKind.FINAL,
                content=fallback,
                created_at=_timestamp(),
            )
        )

    return messages


def build_cell_updates(raw_updates: Any) -> List[CellUpdate]:
    """Parse raw LLM cell_updates into CellUpdate models.

    Args:
        raw_updates: Raw value from the LLM response (should be a list of dicts).

    Returns:
        List of valid CellUpdate instances.
    """
    updates: List[CellUpdate] = []
    for candidate in _ensure_iterable(raw_updates):
        if not isinstance(candidate, dict):
            continue
        address = candidate.get("address")
        values = candidate.get("values")
        if not address or not isinstance(values, list):
            continue
        mode = candidate.get("mode", "replace")
        if isinstance(mode, str):
            normalized_mode = mode.lower()
            if normalized_mode not in {"replace", "append"}:
                normalized_mode = "replace"
            mode = normalized_mode
        worksheet = candidate.get("worksheet")

        try:
            normalized_values = values
            if isinstance(values, list) and values and not isinstance(values[0], list):
                normalized_values = [values]

            updates.append(
                CellUpdate(
                    address=address,
                    values=normalized_values,
                    mode=mode,
                    worksheet=worksheet,
                )
            )
        except ValidationError as error:
            logger.warning(
                "Skipping invalid cell update %s due to validation error: %s",
                candidate,
                error,
            )
    return updates


def build_format_updates(raw_updates: Any) -> List[FormatUpdate]:
    """Parse raw LLM format_updates into FormatUpdate models.

    Args:
        raw_updates: Raw value from the LLM response.

    Returns:
        List of valid FormatUpdate instances.
    """
    updates: List[FormatUpdate] = []
    for candidate in _ensure_iterable(raw_updates):
        if not isinstance(candidate, dict):
            continue
        address = candidate.get("address")
        if not address:
            continue
        worksheet = candidate.get("worksheet")
        fill_color = candidate.get("fill_color", candidate.get("fillColor"))
        font_color = candidate.get("font_color", candidate.get("fontColor"))
        bold = candidate.get("bold")
        italic = candidate.get("italic")
        number_format = candidate.get("number_format", candidate.get("numberFormat"))
        border_color = candidate.get("border_color", candidate.get("borderColor"))
        border_style = candidate.get("border_style", candidate.get("borderStyle"))
        border_weight = candidate.get("border_weight", candidate.get("borderWeight"))
        try:
            updates.append(
                FormatUpdate(
                    address=address,
                    worksheet=worksheet,
                    fill_color=fill_color,
                    font_color=font_color,
                    bold=bold if isinstance(bold, bool) else None,
                    italic=italic if isinstance(italic, bool) else None,
                    number_format=number_format,
                    border_color=border_color,
                    border_style=border_style,
                    border_weight=border_weight,
                )
            )
        except ValidationError as error:
            logger.warning(
                "Skipping invalid format update %s due to validation error: %s",
                candidate,
                error,
            )
    return updates


def build_chart_inserts(raw_inserts: Any) -> List[ChartInsert]:
    """Parse raw LLM chart_inserts into ChartInsert models.

    Args:
        raw_inserts: Raw value from the LLM response.

    Returns:
        List of valid ChartInsert instances.
    """
    inserts: List[ChartInsert] = []
    for candidate in _ensure_iterable(raw_inserts):
        if not isinstance(candidate, dict):
            continue
        chart_type_raw = candidate.get("chart_type") or candidate.get("chartType")
        source_address = candidate.get("source_address") or candidate.get(
            "sourceAddress"
        )
        chart_type = _normalize_chart_type(chart_type_raw)
        if not chart_type or not source_address:
            if not chart_type:
                logger.warning(
                    "Skipping chart insert due to unsupported chart type: %s",
                    chart_type_raw,
                )
            continue
        source_worksheet = candidate.get("source_worksheet") or candidate.get(
            "sourceWorksheet"
        )
        destination_worksheet = candidate.get("destination_worksheet") or candidate.get(
            "destinationWorksheet"
        )
        top_left_cell = candidate.get("top_left_cell") or candidate.get("topLeftCell")
        bottom_right_cell = candidate.get("bottom_right_cell") or candidate.get(
            "bottomRightCell"
        )
        name = candidate.get("name")
        title = candidate.get("title")
        x_axis_title = candidate.get("x_axis_title") or candidate.get("xAxisTitle")
        y_axis_title = candidate.get("y_axis_title") or candidate.get("yAxisTitle")
        series_by_raw = candidate.get("series_by") or candidate.get("seriesBy")
        series_by = ChartSeriesBy.AUTO
        if isinstance(series_by_raw, str):
            normalized_series_by = series_by_raw.lower()
            if normalized_series_by in {item.value for item in ChartSeriesBy}:
                series_by = ChartSeriesBy(normalized_series_by)
        try:
            inserts.append(
                ChartInsert(
                    chart_type=chart_type,
                    source_address=source_address,
                    source_worksheet=source_worksheet,
                    destination_worksheet=destination_worksheet,
                    top_left_cell=top_left_cell,
                    bottom_right_cell=bottom_right_cell,
                    name=name,
                    title=title,
                    x_axis_title=x_axis_title,
                    y_axis_title=y_axis_title,
                    series_by=series_by,
                )
            )
        except ValidationError as error:
            logger.warning(
                "Skipping invalid chart insert %s due to validation error: %s",
                candidate,
                error,
            )
    return inserts


AGGREGATION_ALIASES: Dict[str, str] = {
    "avg": "average",
    "mean": "average",
    "cnt": "count",
    "total": "sum",
    "std": "standardDeviation",
    "stdev": "standardDeviation",
    "stdevp": "standardDeviationP",
    "var": "variance",
    "varp": "varianceP",
    "countnums": "countNumbers",
}

_VALID_AGGREGATIONS = {item.value for item in AggregationFunction}


def _normalize_aggregation(raw: Any) -> str:
    """Normalize an aggregation function string to a valid AggregationFunction value."""
    if not isinstance(raw, str):
        return "sum"
    lowered = raw.strip().lower()
    if lowered in AGGREGATION_ALIASES:
        lowered = AGGREGATION_ALIASES[lowered]
    if lowered in _VALID_AGGREGATIONS:
        return lowered
    return "sum"


def build_pivot_table_inserts(raw_inserts: Any) -> List[PivotTableInsert]:
    """Parse raw LLM pivot_table_inserts into PivotTableInsert models.

    Args:
        raw_inserts: Raw value from the LLM response (should be a list of dicts).

    Returns:
        List of valid PivotTableInsert instances.
    """
    inserts: List[PivotTableInsert] = []
    for candidate in _ensure_iterable(raw_inserts):
        if not isinstance(candidate, dict):
            continue
        name = candidate.get("name")
        source_address = candidate.get("source_address") or candidate.get(
            "sourceAddress"
        )
        if not name or not source_address:
            logger.warning(
                "Skipping pivot table insert missing name or sourceAddress: %s",
                candidate,
            )
            continue

        source_worksheet = candidate.get("source_worksheet") or candidate.get(
            "sourceWorksheet"
        )
        destination_address = (
            candidate.get("destination_address")
            or candidate.get("destinationAddress")
            or None
        )
        destination_worksheet = candidate.get(
            "destination_worksheet"
        ) or candidate.get("destinationWorksheet")
        rows = candidate.get("rows", [])
        if not isinstance(rows, list):
            rows = []
        columns = candidate.get("columns", [])
        if not isinstance(columns, list):
            columns = []
        filters = candidate.get("filters", [])
        if not isinstance(filters, list):
            filters = []

        raw_values = candidate.get("values", [])
        if not isinstance(raw_values, list):
            raw_values = []
        data_fields: List[PivotTableDataField] = []
        for val in raw_values:
            if isinstance(val, str):
                data_fields.append(
                    PivotTableDataField(name=val, summarize_by=AggregationFunction.SUM)
                )
            elif isinstance(val, dict):
                field_name = val.get("name")
                if not field_name:
                    continue
                agg = _normalize_aggregation(val.get("summarizeBy") or val.get("summarize_by"))
                data_fields.append(
                    PivotTableDataField(
                        name=field_name,
                        summarize_by=AggregationFunction(agg),
                    )
                )

        try:
            inserts.append(
                PivotTableInsert(
                    name=name,
                    source_address=source_address,
                    source_worksheet=source_worksheet,
                    destination_address=destination_address,
                    destination_worksheet=destination_worksheet,
                    rows=rows,
                    columns=columns,
                    values=data_fields,
                    filters=filters,
                )
            )
        except ValidationError as error:
            logger.warning(
                "Skipping invalid pivot table insert %s due to validation error: %s",
                candidate,
                error,
            )
    return inserts


def available_providers() -> List[Dict[str, Any]]:
    """Return the list of available providers based on current settings.

    Returns:
        List of provider info dicts for the /providers endpoint.
    """
    settings = get_settings()
    providers: List[Dict[str, Any]] = []
    if settings.mock_provider_enabled:
        providers.append(
            {
                "id": MockProvider.id,
                "label": MockProvider.label,
                "description": MockProvider.description,
                "requiresKey": MockProvider.requires_key,
            }
        )
    providers.append(
        {
            "id": OpenAIProvider.id,
            "label": OpenAIProvider.label,
            "description": OpenAIProvider.description,
            "requiresKey": OpenAIProvider.requires_key,
        }
    )
    providers.append(
        {
            "id": AnthropicProvider.id,
            "label": AnthropicProvider.label,
            "description": AnthropicProvider.description,
            "requiresKey": AnthropicProvider.requires_key,
        }
    )
    return providers


def create_provider(
    provider_id: str,
    mcp_tools: Optional[List[MCPToolEntry]] = None,
) -> LLMProvider:
    """Instantiate an LLM provider by ID, injecting live MCP tools.

    Args:
        provider_id: Provider identifier string (e.g. ``"openai"``).
        mcp_tools: Live MCP tools from the orchestrator's health check.
            Passed to the provider constructor to build the dynamic system
            prompt. Defaults to ``None`` (treated as empty list).

    Returns:
        An ``LLMProvider`` instance configured with the given tools.

    Raises:
        HTTPException: 400 if the provider ID is unknown or disabled.
    """
    provider_id = provider_id.lower()
    if provider_id == MockProvider.id:
        if not get_settings().mock_provider_enabled:
            raise HTTPException(
                status_code=400, detail="Mock provider is disabled on the server."
            )
        return MockProvider()
    if provider_id == OpenAIProvider.id:
        return OpenAIProvider(mcp_tools=mcp_tools)
    if provider_id == AnthropicProvider.id:
        return AnthropicProvider(mcp_tools=mcp_tools)
    raise HTTPException(status_code=400, detail=f"Unknown provider '{provider_id}'.")
