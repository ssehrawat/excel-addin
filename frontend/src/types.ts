export type MessageRole = "user" | "assistant" | "system";

export type MessageKind =
  | "message"
  | "thought"
  | "step"
  | "suggestion"
  | "context"
  | "final";

export interface ChatMessage {
  id: string;
  role: MessageRole;
  kind: MessageKind;
  content: string;
  createdAt: string;
}

export interface CellSelection {
  address: string;
  values: (string | number | boolean | null)[][];
  worksheet?: string | null;
}

export interface CellUpdate {
  address: string;
  values: (string | number | boolean | null)[][];
  mode: "replace" | "append";
  worksheet?: string | null;
}

export interface FormatUpdate {
  address: string;
  worksheet?: string | null;
  fillColor?: string | null;
  fontColor?: string | null;
  bold?: boolean;
  italic?: boolean;
  numberFormat?: string | null;
  borderColor?: string | null;
  borderStyle?: string | null;
  borderWeight?: string | null;
}

export type ChartSeriesBy = "auto" | "rows" | "columns";

export interface ChartInsert {
  chartType: string;
  sourceAddress: string;
  sourceWorksheet?: string | null;
  destinationWorksheet?: string | null;
  name?: string | null;
  title?: string | null;
  topLeftCell?: string | null;
  bottomRightCell?: string | null;
  seriesBy?: ChartSeriesBy;
}

export interface ChatRequest {
  prompt: string;
  provider: string;
  messages: ChatMessage[];
  selection: CellSelection[];
  metadata?: Record<string, unknown>;
}

export interface ChatResponse {
  messages: ChatMessage[];
  cell_updates?: CellUpdate[];
  format_updates?: FormatUpdate[];
  chart_inserts?: ChartInsert[];
  telemetry?: Telemetry | null;
}

export interface ProviderOption {
  id: string;
  label: string;
  description: string;
  requiresKey: boolean;
}

export interface MCPTool {
  name: string;
  description?: string | null;
  inputSchema?: Record<string, unknown>;
}

export type MCPServerStatus = "online" | "offline" | "error" | "unknown";

export interface MCPServer {
  id: string;
  name: string;
  baseUrl: string;
  description?: string | null;
  enabled: boolean;
  status: MCPServerStatus;
  lastRefreshedAt?: string | null;
  tools: MCPTool[];
  createdAt?: string;
  updatedAt?: string;
}

export interface CreateMCPServerPayload {
  name: string;
  baseUrl: string;
  description?: string;
  apiKey?: string;
  enabled?: boolean;
  autoRefresh?: boolean;
}

export interface UpdateMCPServerPayload {
  name?: string;
  baseUrl?: string;
  description?: string;
  apiKey?: string;
  enabled?: boolean;
}

export interface Telemetry {
  latency_ms?: number;
  provider?: string;
  model?: string;
  tokens_prompt?: number;
  tokens_completion?: number;
  raw?: Record<string, unknown> | null;
}

export type ChatStreamEvent =
  | { type: "message_start"; payload: ChatMessage }
  | { type: "message_delta"; payload: { id: string; delta: string } }
  | { type: "message_done"; payload: ChatMessage }
  | { type: "message"; payload: ChatMessage }
  | { type: "cell_updates"; payload: CellUpdate[] }
  | { type: "format_updates"; payload: FormatUpdate[] }
  | { type: "chart_inserts"; payload: ChartInsert[] }
  | { type: "telemetry"; payload: Telemetry | null }
  | { type: "done"; payload?: null }
  | { type: "error"; payload: { message: string } };

