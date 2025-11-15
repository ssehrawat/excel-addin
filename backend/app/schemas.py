from enum import Enum
from typing import Any, Dict, List, Literal, Optional

from pydantic import BaseModel, Field


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


class ChartInsert(BaseModel):
    chart_type: str = Field(..., alias="chartType")
    source_address: str = Field(..., alias="sourceAddress")
    source_worksheet: Optional[str] = Field(default=None, alias="sourceWorksheet")
    destination_worksheet: Optional[str] = Field(
        default=None, alias="destinationWorksheet"
    )
    name: Optional[str] = None
    title: Optional[str] = None
    top_left_cell: Optional[str] = Field(default=None, alias="topLeftCell")
    bottom_right_cell: Optional[str] = Field(default=None, alias="bottomRightCell")
    series_by: ChartSeriesBy = Field(default=ChartSeriesBy.AUTO, alias="seriesBy")

    class Config:
        populate_by_name = True


class ChatRequest(BaseModel):
    prompt: str
    provider: str
    messages: List[ChatMessage]
    selection: List[CellSelection] = Field(default_factory=list)
    metadata: Dict[str, Any] = Field(default_factory=dict)


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
    telemetry: Optional[Telemetry] = None

    class Config:
        populate_by_name = True


class ProvidersResponse(BaseModel):
    providers: List[Dict[str, Any]]


class HealthResponse(BaseModel):
    status: Literal["ok"] = "ok"

