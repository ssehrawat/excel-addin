from enum import Enum
from typing import Any, Dict, List, Literal, Optional

from pydantic import AnyHttpUrl, BaseModel, Field


class MessageRole(str, Enum):
    USER = "user"
    ASSISTANT = "assistant"
    SYSTEM = "system"


class MessageKind(str, Enum):
    MESSAGE = "message"
    THOUGHT = "thought"
    STEP = "step"
    SUGGESTION = "suggestion"
    CONTEXT = "context"
    FINAL = "final"


class ChatMessage(BaseModel):
    id: str
    role: MessageRole
    kind: MessageKind
    content: str
    created_at: str = Field(..., alias="createdAt")

    class Config:
        populate_by_name = True


class CellSelection(BaseModel):
    address: str
    values: List[List[Any]]
    worksheet: Optional[str] = None


class CellUpdateMode(str, Enum):
    REPLACE = "replace"
    APPEND = "append"


class CellUpdate(BaseModel):
    address: str
    values: List[List[Any]]
    mode: CellUpdateMode = CellUpdateMode.REPLACE
    worksheet: Optional[str] = None


class FormatUpdate(BaseModel):
    address: str
    worksheet: Optional[str] = None
    fill_color: Optional[str] = Field(default=None, alias="fillColor")
    font_color: Optional[str] = Field(default=None, alias="fontColor")
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    number_format: Optional[str] = Field(default=None, alias="numberFormat")
    border_color: Optional[str] = Field(default=None, alias="borderColor")
    border_style: Optional[str] = Field(default=None, alias="borderStyle")
    border_weight: Optional[str] = Field(default=None, alias="borderWeight")

    class Config:
        populate_by_name = True


class ChartSeriesBy(str, Enum):
    AUTO = "auto"
    ROWS = "rows"
    COLUMNS = "columns"


class AggregationFunction(str, Enum):
    SUM = "sum"
    COUNT = "count"
    AVERAGE = "average"
    MAX = "max"
    MIN = "min"
    PRODUCT = "product"
    COUNT_NUMBERS = "countNumbers"
    STANDARD_DEVIATION = "standardDeviation"
    STANDARD_DEVIATION_P = "standardDeviationP"
    VARIANCE = "variance"
    VARIANCE_P = "varianceP"


class PivotTableDataField(BaseModel):
    """A data (values) field with aggregation for a pivot table."""

    name: str
    summarize_by: AggregationFunction = Field(
        default=AggregationFunction.SUM, alias="summarizeBy"
    )

    class Config:
        populate_by_name = True


class PivotTableInsert(BaseModel):
    """Schema for creating a pivot table via Office.js."""

    name: str
    source_address: str = Field(..., alias="sourceAddress")
    source_worksheet: Optional[str] = Field(default=None, alias="sourceWorksheet")
    destination_address: Optional[str] = Field(default=None, alias="destinationAddress")
    destination_worksheet: Optional[str] = Field(
        default=None, alias="destinationWorksheet"
    )
    rows: List[str] = Field(default_factory=list)
    columns: List[str] = Field(default_factory=list)
    values: List[PivotTableDataField] = Field(default_factory=list)
    filters: List[str] = Field(default_factory=list)

    class Config:
        populate_by_name = True


class ChartInsert(BaseModel):
    chart_type: str = Field(..., alias="chartType")
    source_address: str = Field(..., alias="sourceAddress")
    source_worksheet: Optional[str] = Field(default=None, alias="sourceWorksheet")
    destination_worksheet: Optional[str] = Field(
        default=None, alias="destinationWorksheet"
    )
    name: Optional[str] = None
    title: Optional[str] = None
    x_axis_title: Optional[str] = Field(default=None, alias="xAxisTitle")
    y_axis_title: Optional[str] = Field(default=None, alias="yAxisTitle")
    top_left_cell: Optional[str] = Field(default=None, alias="topLeftCell")
    bottom_right_cell: Optional[str] = Field(default=None, alias="bottomRightCell")
    series_by: ChartSeriesBy = Field(default=ChartSeriesBy.AUTO, alias="seriesBy")

    class Config:
        populate_by_name = True


# ---------------------------------------------------------------------------
# Workbook context models (new for refactor)
# ---------------------------------------------------------------------------

class SheetMetadata(BaseModel):
    """Metadata for a single worksheet."""

    id: str
    name: str
    index: int = 0
    max_rows: int = Field(default=0, alias="maxRows")
    max_columns: int = Field(default=0, alias="maxColumns")

    class Config:
        populate_by_name = True


class WorkbookMetadata(BaseModel):
    """Workbook-level metadata collected at add-in init."""

    success: bool = True
    file_name: str = Field(default="", alias="fileName")
    sheets_metadata: List[SheetMetadata] = Field(
        default_factory=list, alias="sheetsMetadata"
    )
    total_sheets: int = Field(default=0, alias="totalSheets")

    class Config:
        populate_by_name = True


class UserContext(BaseModel):
    """Per-request context: active sheet and current selection."""

    current_active_sheet_name: str = Field(default="", alias="currentActiveSheetName")
    selected_ranges: str = Field(default="", alias="selectedRanges")

    class Config:
        populate_by_name = True


class WorkbookToolCall(BaseModel):
    """An Excel read tool call requested by the LLM."""

    id: str
    tool: str  # e.g. "get_xl_range_as_csv"
    args: Dict[str, Any] = Field(default_factory=dict)


class WorkbookToolResult(BaseModel):
    """Result of an Excel read tool executed by the frontend."""

    id: str
    tool: str
    result: Any = None
    error: Optional[str] = None


# ---------------------------------------------------------------------------
# Chat request / response
# ---------------------------------------------------------------------------

class ChatRequest(BaseModel):
    prompt: str
    provider: str
    messages: List[ChatMessage]
    selection: List[CellSelection] = Field(default_factory=list)
    metadata: Dict[str, Any] = Field(default_factory=dict)
    # Workbook context (new)
    workbook_metadata: Optional[WorkbookMetadata] = Field(
        default=None, alias="workbookMetadata"
    )
    user_context: Optional[UserContext] = Field(default=None, alias="userContext")
    tool_results: List[WorkbookToolResult] = Field(
        default_factory=list, alias="toolResults"
    )
    active_sheet_preview: Optional[str] = Field(
        default=None, alias="activeSheetPreview"
    )

    class Config:
        populate_by_name = True


class Telemetry(BaseModel):
    latency_ms: Optional[float] = None
    provider: Optional[str] = None
    model: Optional[str] = None
    tokens_prompt: Optional[int] = None
    tokens_completion: Optional[int] = None
    raw: Optional[Dict[str, Any]] = None


class ChatResponse(BaseModel):
    messages: List[ChatMessage]
    cell_updates: List[CellUpdate] = Field(default_factory=list, alias="cell_updates")
    format_updates: List[FormatUpdate] = Field(
        default_factory=list, alias="format_updates"
    )
    chart_inserts: List[ChartInsert] = Field(
        default_factory=list, alias="chart_inserts"
    )
    pivot_table_inserts: List[PivotTableInsert] = Field(
        default_factory=list, alias="pivot_table_inserts"
    )
    telemetry: Optional[Telemetry] = None

    class Config:
        populate_by_name = True


class MCPTool(BaseModel):
    name: str
    description: Optional[str] = None
    input_schema: Dict[str, Any] = Field(default_factory=dict, alias="inputSchema")

    class Config:
        populate_by_name = True


class MCPServerPublic(BaseModel):
    id: str
    name: str
    base_url: AnyHttpUrl = Field(..., alias="baseUrl")
    description: Optional[str] = None
    enabled: bool = True
    status: Literal["online", "offline", "error", "unknown"] = "unknown"
    last_refreshed_at: Optional[str] = Field(default=None, alias="lastRefreshedAt")
    tools: List[MCPTool] = Field(default_factory=list)
    protocol: Literal["auto", "rest", "mcp"] = "auto"
    created_at: Optional[str] = Field(default=None, alias="createdAt")
    updated_at: Optional[str] = Field(default=None, alias="updatedAt")

    class Config:
        populate_by_name = True


class MCPServerCreateRequest(BaseModel):
    name: str
    base_url: AnyHttpUrl = Field(..., alias="baseUrl")
    description: Optional[str] = None
    api_key: Optional[str] = Field(default=None, alias="apiKey")
    enabled: bool = True
    auto_refresh: bool = Field(default=True, alias="autoRefresh")
    protocol: Literal["auto", "rest", "mcp"] = "auto"

    class Config:
        populate_by_name = True


class MCPServerUpdateRequest(BaseModel):
    name: Optional[str] = None
    base_url: Optional[AnyHttpUrl] = Field(default=None, alias="baseUrl")
    description: Optional[str] = None
    api_key: Optional[str] = Field(default=None, alias="apiKey")
    enabled: Optional[bool] = None
    protocol: Optional[Literal["auto", "rest", "mcp"]] = None

    class Config:
        populate_by_name = True


class MCPServerResponse(BaseModel):
    server: MCPServerPublic


class MCPServerListResponse(BaseModel):
    servers: List[MCPServerPublic]


class ProvidersResponse(BaseModel):
    providers: List[Dict[str, Any]]


class HealthResponse(BaseModel):
    status: Literal["ok"] = "ok"
