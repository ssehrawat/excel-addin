/* global Office */

import { useCallback, useEffect, useRef, useState } from "react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { ChatPanel } from "./components/ChatPanel";
import { SettingsDialog } from "./components/SettingsDialog";
import { setSharedProvider, setMutationHandler } from "./sharedState";
import {
  API_BASE_URL,
  DEFAULT_PROVIDER,
  FALLBACK_PROVIDERS
} from "./config";
import {
  ChatMessage,
  ChatRequest,
  ChatResponse,
  ChartInsert,
  PivotTableInsert,
  ProviderOption,
  CellUpdate,
  FormatUpdate,
  Telemetry,
  ChatStreamEvent,
  MCPServer,
  CreateMCPServerPayload,
  WorkbookToolCall,
  WorkbookToolResult
} from "./types";
import {
  applyCellUpdates,
  applyFormatUpdates,
  insertCharts,
  insertPivotTables,
  getCurrentSelection,
  getSelectionsFromPrompt,
  getWorkbookMetadata,
  getUserContext,
  getLightweightSheetPreview,
  initPreviewCache,
  executeWorkbookTool
} from "./excel";

const INITIAL_MESSAGES: ChatMessage[] = [
  {
    id: crypto.randomUUID(),
    role: "assistant",
    kind: "message",
    content:
      "Hi! I can analyze your workbook, reference selected ranges, run calculations, and write results back into Excel. Select some cells and ask away.",
    createdAt: new Date().toISOString()
  }
];

export async function readApiError(res: Response): Promise<string> {
  try {
    const contentType = res.headers.get("content-type") ?? "";
    if (contentType.includes("application/json")) {
      const data = await res.json();
      if (data && typeof data === "object") {
        const maybeDetail = (data as any).detail;
        if (typeof maybeDetail === "string" && maybeDetail.trim()) {
          return maybeDetail;
        }
        const maybeMessage = (data as any).message;
        if (typeof maybeMessage === "string" && maybeMessage.trim()) {
          return maybeMessage;
        }
      }
    }
  } catch {
    // ignore parse errors; fall back to text
  }
  try {
    const text = await res.text();
    return text || `Status ${res.status}`;
  } catch {
    return `Status ${res.status}`;
  }
}

export function formatNetworkError(error: unknown, baseUrl: string): string {
  if (error instanceof TypeError) {
    const msg = error.message || "";
    if (msg.toLowerCase().includes("fetch")) {
      return `Could not reach the backend at ${baseUrl}. Ensure the backend is running and your localhost HTTPS certificate is trusted.`;
    }
  }
  return error instanceof Error ? error.message : "Network error";
}

// Typed callback aliases for streamRound.
// eslint-disable-next-line no-unused-vars
type MessageFn = (msg: ChatMessage) => void;
// eslint-disable-next-line no-unused-vars
type DeltaFn = (id: string, delta: string) => void;
// eslint-disable-next-line no-unused-vars
type StringFn = (text: string) => void;
/** Thinking step progress event from the backend. */
// eslint-disable-next-line no-unused-vars
type StepFn = (step: { id: string; text: string; status: "active" | "done" }) => void;

/** Result of a single streaming round with the backend. */
interface StreamRoundResult {
  cellUpdates: CellUpdate[];
  formatUpdates: FormatUpdate[];
  chartInserts: ChartInsert[];
  pivotTableInserts: PivotTableInsert[];
  telemetry: Telemetry | null;
  /** Non-null when the LLM requested Excel tool data. */
  toolCallRequired: WorkbookToolCall[] | null;
}

export function App() {
  const [messages, setMessages] = useState<ChatMessage[]>(INITIAL_MESSAGES);
  const [provider, setProviderState] = useState<string>(DEFAULT_PROVIDER);
  const [providers, setProviders] = useState(FALLBACK_PROVIDERS);
  const [mcpServers, setMcpServers] = useState<MCPServer[]>([]);
  const [mcpServersLoading, setMcpServersLoading] = useState(false);
  const [mcpBusyIds, setMcpBusyIds] = useState<string[]>([]);
  const [mcpError, setMcpError] = useState<string | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [isBusy, setIsBusy] = useState(false);
  const [thinkingSteps, setThinkingSteps] = useState<
    { id: string; text: string; status: "active" | "done" }[]
  >([]);

  /** Update both React state and shared window global for custom functions. */
  const setProvider = useCallback((next: string) => {
    setProviderState(next);
    setSharedProvider(next);
  }, []);

  const messagesRef = useRef(messages);
  useEffect(() => {
    messagesRef.current = messages;
  }, [messages]);

  // Collect workbook metadata and register preview cache listeners on mount
  useEffect(() => {
    setSharedProvider(DEFAULT_PROVIDER);

    // Bridge mutation-applying functions so =ASKAI() custom functions can
    // dispatch pivot tables, charts, cell/format updates via the taskpane.
    setMutationHandler(async (mutations) => {
      if (mutations.cellUpdates?.length)       await applyCellUpdates(mutations.cellUpdates);
      if (mutations.formatUpdates?.length)     await applyFormatUpdates(mutations.formatUpdates);
      if (mutations.chartInserts?.length)      await insertCharts(mutations.chartInserts);
      if (mutations.pivotTableInserts?.length) await insertPivotTables(mutations.pivotTableInserts);
    });

    initPreviewCache().catch((err) =>
      console.warn("Failed to init preview cache:", err)
    );
  }, []);

  // Load available providers from backend
  useEffect(() => {
    const loadProviders = async () => {
      try {
        const res = await fetch(`${API_BASE_URL}/providers`);
        if (!res.ok) {
          throw new Error(await readApiError(res));
        }
        const data = await res.json();
        if (Array.isArray(data.providers) && data.providers.length > 0) {
          const normalized = data.providers.map((item: any) => ({
            id: String(item.id),
            label: item.label ?? item.id,
            description: item.description ?? "",
            requiresKey: Boolean(item.requiresKey)
          }));
          setProviders(normalized);
          if (!normalized.some((item: ProviderOption) => item.id === provider)) {
            setProvider(normalized[0].id);
          }
        }
      } catch (error) {
        console.warn("Falling back to bundled provider list", error);
      }
    };
    void loadProviders();
  }, []);

  const loadMcpServers = useCallback(async () => {
    setMcpServersLoading(true);
    try {
      const res = await fetch(`${API_BASE_URL}/mcp/servers`);
      if (!res.ok) {
        throw new Error(await readApiError(res));
      }
      const data = await res.json();
      if (Array.isArray(data.servers)) {
        setMcpServers(data.servers);
      }
      setMcpError(null);
    } catch (error) {
      setMcpError(formatNetworkError(error, API_BASE_URL));
    } finally {
      setMcpServersLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadMcpServers();
  }, [loadMcpServers]);

  const markMcpBusy = (id: string, busy: boolean) => {
    setMcpBusyIds((prev) => {
      if (busy) {
        if (prev.includes(id)) return prev;
        return [...prev, id];
      }
      return prev.filter((item) => item !== id);
    });
  };

  const handleCreateMcpServer = async (payload: CreateMCPServerPayload) => {
    try {
      setMcpError(null);
      const res = await fetch(`${API_BASE_URL}/mcp/servers`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (!res.ok) throw new Error(await readApiError(res));
      await loadMcpServers();
    } catch (error) {
      setMcpError(formatNetworkError(error, API_BASE_URL));
    }
  };

  const handleToggleMcpServer = async (id: string, enabled: boolean) => {
    markMcpBusy(id, true);
    try {
      setMcpError(null);
      const res = await fetch(`${API_BASE_URL}/mcp/servers/${id}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ enabled })
      });
      if (!res.ok) throw new Error(await readApiError(res));
      await loadMcpServers();
    } catch (error) {
      setMcpError(formatNetworkError(error, API_BASE_URL));
    } finally {
      markMcpBusy(id, false);
    }
  };

  const handleRefreshMcpServer = async (id: string) => {
    markMcpBusy(id, true);
    try {
      setMcpError(null);
      const res = await fetch(`${API_BASE_URL}/mcp/servers/${id}/refresh`, {
        method: "POST"
      });
      if (!res.ok) throw new Error(await readApiError(res));
      await loadMcpServers();
    } catch (error) {
      setMcpError(formatNetworkError(error, API_BASE_URL));
    } finally {
      markMcpBusy(id, false);
    }
  };

  /**
   * Opens the settings dialog and immediately triggers a concurrent refresh
   * of all enabled MCP servers so the status badges reflect current reachability
   * rather than the last persisted value.
   *
   * The dialog opens synchronously; refreshes run in the background and
   * the server list is reloaded once all settle. Per-server failures are
   * silently swallowed so a single offline server cannot block the UI.
   */
  const handleOpenSettings = useCallback(async () => {
    setSettingsOpen(true);
    // Refresh all enabled servers concurrently so status badges reflect reality.
    const enabledServers = mcpServers.filter((s) => s.enabled);
    if (enabledServers.length === 0) return;
    await Promise.allSettled(
      enabledServers.map((s) =>
        fetch(`${API_BASE_URL}/mcp/servers/${s.id}/refresh`, { method: "POST" })
          .catch(() => {})
      )
    );
    await loadMcpServers();
  }, [mcpServers, loadMcpServers]);

  const handleDeleteMcpServer = async (id: string) => {
    markMcpBusy(id, true);
    try {
      setMcpError(null);
      const res = await fetch(`${API_BASE_URL}/mcp/servers/${id}`, {
        method: "DELETE"
      });
      if (!res.ok && res.status !== 204) {
        throw new Error(await readApiError(res));
      }
      await loadMcpServers();
    } catch (error) {
      setMcpError(formatNetworkError(error, API_BASE_URL));
    } finally {
      markMcpBusy(id, false);
    }
  };

  /**
   * Execute a single streaming round with the backend.
   * Handles message_start/delta/done streaming and accumulates Excel mutations.
   * Returns tool call info if the LLM needs more Excel data.
   */
  const streamRound = async (
    payload: ChatRequest,
    onAppendMessage: MessageFn,
    onAppendDelta: DeltaFn,
    onFinalizeMessage: MessageFn,
    onStatus: StringFn,
    onStep: StepFn
  ): Promise<StreamRoundResult> => {
    const result: StreamRoundResult = {
      cellUpdates: [],
      formatUpdates: [],
      chartInserts: [],
      pivotTableInserts: [],
      telemetry: null,
      toolCallRequired: null
    };

    const response = await fetch(`${API_BASE_URL}/chat`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/x-ndjson"
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Backend error (${response.status}): ${errorText || "Unknown"}`);
    }

    const contentType = response.headers.get("content-type") ?? "";

    if (contentType.includes("application/x-ndjson")) {
      let buffer = "";

      const handleEvent = (event: ChatStreamEvent) => {
        switch (event.type) {
          case "message_start":
            setThinkingSteps([]);
            onAppendMessage(event.payload);
            break;
          case "message_delta":
            onAppendDelta(event.payload.id, event.payload.delta);
            break;
          case "message_done":
            onFinalizeMessage(event.payload);
            break;
          case "message":
            onFinalizeMessage(event.payload);
            break;
          case "cell_updates":
            if (event.payload?.length) {
              result.cellUpdates = result.cellUpdates.concat(event.payload);
            }
            break;
          case "format_updates":
            if (event.payload?.length) {
              result.formatUpdates = result.formatUpdates.concat(event.payload);
            }
            break;
          case "chart_inserts":
            if (event.payload?.length) {
              result.chartInserts = result.chartInserts.concat(event.payload);
            }
            break;
          case "pivot_table_inserts":
            if (event.payload?.length) {
              result.pivotTableInserts = result.pivotTableInserts.concat(event.payload);
            }
            break;
          case "telemetry":
            result.telemetry = event.payload ?? null;
            break;
          case "tool_call_required":
            result.toolCallRequired = event.payload ?? null;
            break;
          case "status":
            onStatus(event.payload);
            break;
          case "step":
            onStep(event.payload);
            break;
          case "error":
            throw new Error(event.payload?.message ?? "Streaming error");
          case "done":
          default:
            break;
        }
      };

      const drainBuffer = (flush: boolean) => {
        let newlineIndex = buffer.indexOf("\n");
        while (newlineIndex !== -1) {
          const line = buffer.slice(0, newlineIndex).trim();
          buffer = buffer.slice(newlineIndex + 1);
          if (line) {
            handleEvent(JSON.parse(line) as ChatStreamEvent);
          }
          newlineIndex = buffer.indexOf("\n");
        }
        if (flush) {
          const remaining = buffer.trim();
          buffer = "";
          if (remaining) {
            handleEvent(JSON.parse(remaining) as ChatStreamEvent);
          }
        }
      };

      if (response.body) {
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let reading = true;
        while (reading) {
          const { value, done } = await reader.read();
          if (value) {
            buffer += decoder.decode(value, { stream: true });
            drainBuffer(false);
          }
          if (done) reading = false;
        }
        buffer += decoder.decode();
        drainBuffer(true);
      } else {
        buffer = await response.text();
        drainBuffer(true);
      }
    } else {
      // Non-streaming fallback
      const json = (await response.json()) as ChatResponse;
      for (const msg of json.messages) {
        if (msg.kind === "final" || msg.kind === "message") {
          onFinalizeMessage(msg);
        }
      }
      if (json.cell_updates?.length) {
        result.cellUpdates = json.cell_updates;
      }
      if (json.format_updates?.length) {
        result.formatUpdates = json.format_updates;
      }
      if (json.chart_inserts?.length) {
        result.chartInserts = json.chart_inserts;
      }
      if (json.pivot_table_inserts?.length) {
        result.pivotTableInserts = json.pivot_table_inserts;
      }
    }

    return result;
  };

  const handleSend = async (prompt: string) => {
    const userMessage: ChatMessage = {
      id: crypto.randomUUID(),
      role: "user",
      kind: "message",
      content: prompt,
      createdAt: new Date().toISOString()
    };

    setMessages((prev) => {
      const next = [...prev, userMessage];
      messagesRef.current = next;
      return next;
    });
    setIsBusy(true);
    setThinkingSteps([]);

    const appendMessage = (message: ChatMessage) => {
      // Only add visible message kinds to state
      if (
        message.role !== "user" &&
        message.kind !== "final" &&
        message.kind !== "message"
      ) {
        return;
      }
      setMessages((prev) => {
        const next = [...prev, message];
        messagesRef.current = next;
        return next;
      });
    };

    const appendMessageDelta = (id: string, delta: string) => {
      setMessages((prev) => {
        const next = prev.map((msg) =>
          msg.id === id ? { ...msg, content: `${msg.content}${delta}` } : msg
        );
        messagesRef.current = next;
        return next;
      });
    };

    const finalizeMessage = (message: ChatMessage) => {
      if (
        message.role !== "user" &&
        message.kind !== "final" &&
        message.kind !== "message"
      ) {
        return;
      }
      setMessages((prev) => {
        const exists = prev.some((msg) => msg.id === message.id);
        const next = exists
          ? prev.map((msg) => (msg.id === message.id ? message : msg))
          : [...prev, message];
        messagesRef.current = next;
        return next;
      });
    };

    // eslint-disable-next-line @typescript-eslint/no-empty-function
    const onStatus: StringFn = () => {};
    const onStep = (step: { id: string; text: string; status: "active" | "done" }) => {
      setThinkingSteps((prev) => {
        const idx = prev.findIndex((s) => s.id === step.id);
        if (idx >= 0) {
          const next = [...prev];
          next[idx] = step;
          return next;
        }
        return [...prev, step];
      });
    };

    try {
      if (typeof Office === "undefined") {
        throw new Error("Office runtime is not available. Please run inside Excel.");
      }

      // Collect per-request context in parallel
      const [freshMetadata, userContext, preview] = await Promise.all([
        getWorkbookMetadata(),
        getUserContext(),
        getLightweightSheetPreview(50)
      ]);

      let selection = await getSelectionsFromPrompt(prompt);
      if (selection.length === 0) {
        selection = await getCurrentSelection();
      }

      let payload: ChatRequest = {
        prompt,
        provider,
        messages: messagesRef.current,
        selection,
        workbookMetadata: freshMetadata.success ? freshMetadata : undefined,
        userContext,
        activeSheetPreview: preview ?? undefined,
        metadata: {
          platform: "excel",
          version: Office.context?.diagnostics?.version
        }
      };

      // Accumulated Excel mutations across all rounds
      let allCellUpdates: CellUpdate[] = [];
      let allFormatUpdates: FormatUpdate[] = [];
      let allChartInserts: ChartInsert[] = [];
      let allPivotTableInserts: PivotTableInsert[] = [];
      let pendingTelemetry: Telemetry | null = null;

      // Up to MAX_TOOL_ROUNDS of tool-call round-trips
      const MAX_TOOL_ROUNDS = 3;
      for (let round = 0; round < MAX_TOOL_ROUNDS; round++) {
        const roundResult = await streamRound(
          payload,
          appendMessage,
          appendMessageDelta,
          finalizeMessage,
          onStatus,
          onStep
        );

        allCellUpdates = allCellUpdates.concat(roundResult.cellUpdates);
        allFormatUpdates = allFormatUpdates.concat(roundResult.formatUpdates);
        allChartInserts = allChartInserts.concat(roundResult.chartInserts);
        allPivotTableInserts = allPivotTableInserts.concat(roundResult.pivotTableInserts);
        if (roundResult.telemetry) {
          pendingTelemetry = roundResult.telemetry;
        }

        if (!roundResult.toolCallRequired) {
          // LLM answered directly — we're done
          break;
        }

        // Execute the requested Excel tools and re-POST
        const toolResults: WorkbookToolResult[] = await Promise.all(
          roundResult.toolCallRequired.map((call) => executeWorkbookTool(call))
        );

        payload = {
          ...payload,
          toolResults,
          messages: messagesRef.current
        };
      }

      // Apply all accumulated Excel mutations after stream completes
      if (allCellUpdates.length > 0) {
        await applyCellUpdates(allCellUpdates);
      }
      if (allFormatUpdates.length > 0) {
        await applyFormatUpdates(allFormatUpdates);
      }
      if (allChartInserts.length > 0) {
        await insertCharts(allChartInserts);
      }
      if (allPivotTableInserts.length > 0) {
        await insertPivotTables(allPivotTableInserts);
      }
      if (pendingTelemetry) {
        console.debug("Chat telemetry", pendingTelemetry);
      }
    } catch (error) {
      console.error(error);
      const errorMessage: ChatMessage = {
        id: crypto.randomUUID(),
        role: "assistant",
        kind: "message",
        content:
          error instanceof Error
            ? `Something went wrong: ${error.message}`
            : "Something went wrong. Please try again.",
        createdAt: new Date().toISOString()
      };
      setMessages((prev) => {
        const next = [...prev, errorMessage];
        messagesRef.current = next;
        return next;
      });
    } finally {
      setIsBusy(false);
      setThinkingSteps([]);
    }
  };

  return (
    <FluentProvider theme={webLightTheme} style={{ height: "100%" }}>
      <ChatPanel
        messages={messages}
        isBusy={isBusy}
        thinkingSteps={thinkingSteps}
        onSend={handleSend}
        onOpenSettings={handleOpenSettings}
        onNewChat={() => {
          setMessages(INITIAL_MESSAGES);
          setThinkingSteps([]);
        }}
      />
      <SettingsDialog
        open={settingsOpen}
        providers={providers}
        selectedProvider={provider}
        onSelect={(next) => setProvider(next)}
        onClose={() => setSettingsOpen(false)}
        mcpServers={mcpServers}
        mcpServersLoading={mcpServersLoading}
        mcpBusyIds={mcpBusyIds}
        mcpError={mcpError}
        onCreateMcpServer={handleCreateMcpServer}
        onToggleMcpServer={handleToggleMcpServer}
        onRefreshMcpServer={handleRefreshMcpServer}
        onDeleteMcpServer={handleDeleteMcpServer}
      />
    </FluentProvider>
  );
}
