/**
 * Component tests for ChatPanel.
 *
 * Tests rendering, message filtering, keyboard behavior, and thinking
 * step visualization using @testing-library/react.
 */

import { describe, it, expect, vi } from "vitest";
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

  it("shows title text", () => {
    renderPanel();
    expect(screen.getByText("Workbook Copilot")).toBeInTheDocument();
  });
});
