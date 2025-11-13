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
  telemetry?: Record<string, unknown>;
}

export interface ProviderOption {
  id: string;
  label: string;
  description: string;
  requiresKey: boolean;
}

