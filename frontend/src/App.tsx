/* global Office */

import { useEffect, useRef, useState } from "react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { ChatPanel } from "./components/ChatPanel";
import { SettingsDialog } from "./components/SettingsDialog";
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
  ProviderOption,
  CellUpdate,
  FormatUpdate,
  Telemetry,
  ChatStreamEvent
} from "./types";
import {
  applyCellUpdates,
  applyFormatUpdates,
  insertCharts,
  getCurrentSelection,
  getSelectionsFromPrompt
} from "./excel";

const INITIAL_MESSAGES: ChatMessage[] = [
  {
    id: crypto.randomUUID(),
    role: "assistant",
    kind: "context",
    content:
      "Hi! I can analyze your workbook, reference selected ranges, run calculations, and write results back into Excel. Select some cells and ask away.",
    createdAt: new Date().toISOString()
  }
];

export function App() {
  const [messages, setMessages] = useState<ChatMessage[]>(INITIAL_MESSAGES);
  const [provider, setProvider] = useState<string>(DEFAULT_PROVIDER);
  const [providers, setProviders] =
    useState(FALLBACK_PROVIDERS);
  useEffect(() => {
    const loadProviders = async () => {
      try {
        const res = await fetch(`${API_BASE_URL}/providers`);
        if (!res.ok) {
          throw new Error(`Status ${res.status}`);
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
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [isBusy, setIsBusy] = useState(false);

  const messagesRef = useRef(messages);
  useEffect(() => {
    messagesRef.current = messages;
  }, [messages]);

  const handleSend = async (prompt: string) => {
    const userMessage: ChatMessage = {
      id: crypto.randomUUID(),
      role: "user",
      kind: "message",
      content: prompt,
      createdAt: new Date().toISOString()
    };

    const history = [...messagesRef.current, userMessage];
    setMessages(history);
    messagesRef.current = history;
    setIsBusy(true);

    const appendMessage = (message: ChatMessage) => {
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
      setMessages((prev) => {
        const exists = prev.some((msg) => msg.id === message.id);
        const next = exists
          ? prev.map((msg) => (msg.id === message.id ? message : msg))
          : [...prev, message];
        messagesRef.current = next;
        return next;
      });
    };

    try {
      if (typeof Office === "undefined") {
        throw new Error("Office runtime is not available. Please run inside Excel.");
      }

      let selection = await getSelectionsFromPrompt(prompt);
      if (selection.length === 0) {
        selection = await getCurrentSelection();
      }

      const payload: ChatRequest = {
        prompt,
        provider,
        messages: history,
        selection,
        metadata: {
          platform: "excel",
          version: Office.context?.diagnostics?.version
        }
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
        throw new Error(
          `Backend error (${response.status}): ${errorText || "Unknown"}`
        );
      }

      const contentType = response.headers.get("content-type") ?? "";
      if (contentType.includes("application/x-ndjson")) {
        let pendingCellUpdates: CellUpdate[] = [];
        let pendingFormatUpdates: FormatUpdate[] = [];
        let pendingChartInserts: ChartInsert[] = [];
        let pendingTelemetry: Telemetry | null = null;

        let buffer = "";

        const handleStreamEvent = (event: ChatStreamEvent) => {
          switch (event.type) {
            case "message_start":
              appendMessage(event.payload);
              break;
            case "message_delta":
              appendMessageDelta(event.payload.id, event.payload.delta);
              break;
            case "message_done":
              finalizeMessage(event.payload);
              break;
            case "message":
              finalizeMessage(event.payload);
              break;
            case "cell_updates":
              if (event.payload && event.payload.length > 0) {
                pendingCellUpdates = pendingCellUpdates.concat(event.payload);
              }
              break;
            case "format_updates":
              if (event.payload && event.payload.length > 0) {
                pendingFormatUpdates = pendingFormatUpdates.concat(event.payload);
              }
              break;
            case "chart_inserts":
              if (event.payload && event.payload.length > 0) {
                pendingChartInserts = pendingChartInserts.concat(event.payload);
              }
              break;
            case "telemetry":
              pendingTelemetry = event.payload ?? null;
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
              const event = JSON.parse(line) as ChatStreamEvent;
              handleStreamEvent(event);
            }
            newlineIndex = buffer.indexOf("\n");
          }
          if (flush) {
            const remaining = buffer.trim();
            buffer = "";
            if (remaining) {
              const event = JSON.parse(remaining) as ChatStreamEvent;
              handleStreamEvent(event);
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
            if (done) {
              reading = false;
            }
          }

          buffer += decoder.decode();
          drainBuffer(true);
        } else {
          const textPayload = await response.text();
          buffer = textPayload;
          drainBuffer(true);
        }

        if (pendingCellUpdates.length > 0) {
          await applyCellUpdates(pendingCellUpdates);
        }
        if (pendingFormatUpdates.length > 0) {
          await applyFormatUpdates(pendingFormatUpdates);
        }
        if (pendingChartInserts.length > 0) {
          await insertCharts(pendingChartInserts);
        }
        if (pendingTelemetry) {
          console.debug("Chat telemetry", pendingTelemetry);
        }
      } else {
        const json = (await response.json()) as ChatResponse;
        const combinedMessages = [...history, ...json.messages];
        setMessages(combinedMessages);
        messagesRef.current = combinedMessages;

        if (json.cell_updates && json.cell_updates.length > 0) {
          await applyCellUpdates(json.cell_updates);
        }
        if (json.format_updates && json.format_updates.length > 0) {
          await applyFormatUpdates(json.format_updates);
        }
        if (json.chart_inserts && json.chart_inserts.length > 0) {
          await insertCharts(json.chart_inserts);
        }
      }
    } catch (error) {
      console.error(error);
      const errorMessage: ChatMessage = {
        id: crypto.randomUUID(),
        role: "assistant",
        kind: "step",
        content:
          error instanceof Error
            ? `Something went wrong: ${error.message}`
            : "Something went wrong. Please try again.",
        createdAt: new Date().toISOString()
      };
      const combined = [...messagesRef.current, errorMessage];
      setMessages(combined);
      messagesRef.current = combined;
    } finally {
      setIsBusy(false);
    }
  };

  return (
    <FluentProvider theme={webLightTheme} style={{ height: "100%" }}>
      <ChatPanel
        messages={messages}
        isBusy={isBusy}
        onSend={handleSend}
        onOpenSettings={() => setSettingsOpen(true)}
      />
      <SettingsDialog
        open={settingsOpen}
        providers={providers}
        selectedProvider={provider}
        onSelect={(next) => setProvider(next)}
        onClose={() => setSettingsOpen(false)}
      />
    </FluentProvider>
  );
}

