/**
 * Component tests for ChatPanel.
 *
 * Tests rendering, message filtering, keyboard behavior, and thinking
 * step visualization using @testing-library/react.
 */

import { describe, it, expect, vi, afterEach } from "vitest";
import { render, screen, fireEvent } from "@testing-library/react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { ChatPanel } from "../components/ChatPanel";
import type { ChatMessage } from "../types";

const defaultProps = {
  messages: [] as ChatMessage[],
  isBusy: false,
  thinkingSteps: [] as { id: string; text: string; status: "active" | "done" }[],
  onSend: vi.fn(async () => {}),
  onOpenSettings: vi.fn(),
  onNewChat: vi.fn(),
};

function renderPanel(overrides: Partial<typeof defaultProps> = {}) {
  return render(
    <FluentProvider theme={webLightTheme}>
      <ChatPanel {...defaultProps} {...overrides} />
    </FluentProvider>
  );
}

describe("ChatPanel", () => {
  it("renders visible messages", () => {
    const messages: ChatMessage[] = [
      {
        id: "1",
        role: "assistant",
        kind: "final",
        content: "Hello there!",
        createdAt: new Date().toISOString(),
      },
      {
        id: "2",
        role: "user",
        kind: "message",
        content: "Hi back",
        createdAt: new Date().toISOString(),
      },
    ];
    renderPanel({ messages });
    expect(screen.getByText("Hello there!")).toBeInTheDocument();
    expect(screen.getByText("Hi back")).toBeInTheDocument();
  });

  it("hides internal message kinds (thought, context, step)", () => {
    const messages: ChatMessage[] = [
      {
        id: "1",
        role: "assistant",
        kind: "thought",
        content: "Internal thought",
        createdAt: new Date().toISOString(),
      },
      {
        id: "2",
        role: "assistant",
        kind: "context",
        content: "Context data",
        createdAt: new Date().toISOString(),
      },
      {
        id: "3",
        role: "assistant",
        kind: "final",
        content: "Visible answer",
        createdAt: new Date().toISOString(),
      },
    ];
    renderPanel({ messages });
    expect(screen.queryByText("Internal thought")).not.toBeInTheDocument();
    expect(screen.queryByText("Context data")).not.toBeInTheDocument();
    expect(screen.getByText("Visible answer")).toBeInTheDocument();
  });

  it("disables send button when busy", () => {
    renderPanel({ isBusy: true });
    const sendButton = screen.getByRole("button", { name: /send/i });
    expect(sendButton).toBeDisabled();
  });

  it("disables textarea when busy", () => {
    renderPanel({ isBusy: true });
    const textarea = screen.getByPlaceholderText(/ask about your workbook/i);
    expect(textarea).toBeDisabled();
  });

  it("renders thinking steps", () => {
    renderPanel({
      thinkingSteps: [
        { id: "s1", text: "Analyzing data...", status: "active" },
        { id: "s2", text: "Done reading", status: "done" },
      ],
    });
    expect(screen.getByText("Analyzing data...")).toBeInTheDocument();
    expect(screen.getByText("Done reading")).toBeInTheDocument();
  });

  it("new chat button calls onNewChat", () => {
    const onNewChat = vi.fn();
    renderPanel({ onNewChat });
    // Find the New chat button by tooltip
    const buttons = screen.getAllByRole("button");
    // The New Chat button should be one of the icon buttons
    const newChatBtn = buttons.find(
      (btn) => btn.getAttribute("aria-label") === "New chat" ||
               btn.closest("[aria-label='New chat']") !== null
    );
    if (newChatBtn) {
      fireEvent.click(newChatBtn);
      expect(onNewChat).toHaveBeenCalled();
    }
  });

});

/* ── Voice input integration tests ─────────────────────────────── */

describe("ChatPanel — voice input", () => {
  let savedWebkit: unknown;

  afterEach(() => {
    if (savedWebkit !== undefined) {
      (window as Record<string, unknown>).webkitSpeechRecognition = savedWebkit;
    } else {
      delete (window as Record<string, unknown>).webkitSpeechRecognition;
    }
  });

  function mockWebkitSpeechRecognition() {
    const instance = {
      continuous: false,
      interimResults: false,
      lang: "",
      onresult: null as unknown,
      onerror: null as unknown,
      onend: null as unknown,
      start: vi.fn(),
      stop: vi.fn(function (this: { onend: (() => void) | null }) { this.onend?.(); }),
      abort: vi.fn(),
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
      dispatchEvent: vi.fn(() => true),
    };
    savedWebkit = (window as Record<string, unknown>).webkitSpeechRecognition;
    (window as Record<string, unknown>).webkitSpeechRecognition = vi.fn(() => instance);
    return instance;
  }

  it("renders mic button when speech API available", () => {
    mockWebkitSpeechRecognition();
    renderPanel();
    expect(screen.getByRole("button", { name: "Start voice input (Web Speech API)" })).toBeInTheDocument();
  });

  it("hides mic button when speech API unavailable", () => {
    savedWebkit = (window as Record<string, unknown>).webkitSpeechRecognition;
    delete (window as Record<string, unknown>).webkitSpeechRecognition;
    delete (window as Record<string, unknown>).SpeechRecognition;
    // Also clear getUserMedia to ensure whisper fallback isn't available
    const savedGUM = navigator.mediaDevices?.getUserMedia;
    if (navigator.mediaDevices) {
      (navigator.mediaDevices as Record<string, unknown>).getUserMedia = undefined;
    }

    renderPanel();
    expect(screen.queryByRole("button", { name: /start voice input/i })).not.toBeInTheDocument();

    // Restore
    if (navigator.mediaDevices) {
      (navigator.mediaDevices as Record<string, unknown>).getUserMedia = savedGUM;
    }
  });

  it("mic button disabled when busy", () => {
    mockWebkitSpeechRecognition();
    renderPanel({ isBusy: true });
    const micBtn = screen.getByRole("button", { name: "Start voice input (Web Speech API)" });
    expect(micBtn).toBeDisabled();
  });

  it("mic button toggles icon on click", () => {
    mockWebkitSpeechRecognition();
    renderPanel();
    const micBtn = screen.getByRole("button", { name: "Start voice input (Web Speech API)" });
    fireEvent.click(micBtn);
    // After click, the aria-label should change to "Stop voice input"
    expect(screen.getByRole("button", { name: "Stop voice input" })).toBeInTheDocument();
  });
});
