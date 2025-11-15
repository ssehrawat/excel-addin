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
  telemetry?: Telemetry | null;
}

export interface ProviderOption {
  id: string;
  label: string;
  description: string;
  requiresKey: boolean;
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
  | { type: "message"; payload: ChatMessage }
  | { type: "cell_updates"; payload: CellUpdate[] }
  | { type: "format_updates"; payload: FormatUpdate[] }
  | { type: "telemetry"; payload: Telemetry | null }
  | { type: "done"; payload?: null }
  | { type: "error"; payload: { message: string } };

