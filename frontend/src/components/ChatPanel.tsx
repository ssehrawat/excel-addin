import {
  Button,
  Spinner,
  Tooltip,
  makeStyles,
  shorthands
} from "@fluentui/react-components";
import { Add24Regular, CheckmarkCircle16Filled, Send24Filled, Settings24Regular } from "@fluentui/react-icons";
import { FormEvent, KeyboardEvent, useEffect, useRef, useState } from "react";
import { ChatMessage } from "../types";

const useStyles = makeStyles({
  root: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    backgroundColor: "#f7f7f7"
  },
  header: {
    padding: "12px 16px",
    borderBottom: "1px solid #e4e4e4",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    backgroundColor: "#ffffff"
  },
  headerBrand: {
    display: "flex",
    alignItems: "center",
    gap: "10px"
  },
  headerTitle: {
    fontSize: "16px",
    fontWeight: 600,
    color: "#111827"
  },
  messages: {
    flex: 1,
    overflowY: "auto",
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "10px"
  },
  messageBubble: {
    ...shorthands.padding("10px", "14px"),
    borderRadius: "14px",
    maxWidth: "88%",
    lineHeight: "1.5",
    fontSize: "14px",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word"
  },
  bubbleUser: {
    alignSelf: "flex-end",
    backgroundColor: "#2563eb",
    color: "#ffffff",
    borderRadius: "14px 14px 4px 14px"
  },
  bubbleAssistant: {
    alignSelf: "flex-start",
    backgroundColor: "#ffffff",
    color: "#111827",
    border: "1px solid #e5e7eb",
    borderRadius: "14px 14px 14px 4px"
  },
  thinkingStep: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "4px 0",
    fontSize: "13px",
    alignSelf: "flex-start"
  },
  thinkingStepDone: {
    color: "#10b981"
  },
  thinkingStepActive: {
    color: "#6b7280"
  },
  composer: {
    borderTop: "1px solid #e4e4e4",
    padding: "12px",
    backgroundColor: "#ffffff"
  },
  composerRow: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-end"
  },
  textarea: {
    flex: 1,
    resize: "none",
    border: "1px solid #d1d5db",
    borderRadius: "8px",
    padding: "10px 12px",
    fontSize: "14px",
    fontFamily: "inherit",
    lineHeight: "1.5",
    outline: "none",
    minHeight: "44px",
    maxHeight: "160px",
    overflowY: "auto"
  }
});

interface ChatPanelProps {
  messages: ChatMessage[];
  isBusy: boolean;
  thinkingSteps: { id: string; text: string; status: "active" | "done" }[];
  // eslint-disable-next-line no-unused-vars
  onSend: (text: string) => Promise<void>;
  onOpenSettings: () => void;
  onNewChat: () => void;
}

export function ChatPanel({
  messages,
  isBusy,
  thinkingSteps,
  onSend,
  onOpenSettings,
  onNewChat
}: ChatPanelProps) {
  const styles = useStyles();
  const [input, setInput] = useState("");
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Auto-scroll to bottom whenever messages or status change
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, thinkingSteps]);

  // Auto-resize textarea
  useEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = `${Math.min(el.scrollHeight, 160)}px`;
  }, [input]);

  const handleSubmit = async (event: FormEvent) => {
    event.preventDefault();
    const trimmed = input.trim();
    if (!trimmed || isBusy) return;
    setInput("");
    await onSend(trimmed);
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>) => {
    // Enter sends; Shift+Enter inserts a newline
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void handleSubmit(event as unknown as FormEvent);
    }
  };

  // Only render user messages and final/message assistant answers
  const visibleMessages = messages.filter(
    (m) => m.role === "user" || m.kind === "final" || m.kind === "message"
  );

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <div className={styles.headerBrand}>
          <div className={styles.headerTitle}>Workbook Copilot</div>
        </div>
        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
          <Tooltip content="New chat" relationship="label">
            <Button
              icon={<Add24Regular />}
              appearance="subtle"
              onClick={onNewChat}
              disabled={isBusy}
            />
          </Tooltip>
          <Tooltip content="Settings" relationship="label">
            <Button
              icon={<Settings24Regular />}
              appearance="subtle"
              onClick={onOpenSettings}
              disabled={isBusy}
            />
          </Tooltip>
        </div>
      </div>

      <div className={styles.messages}>
        {visibleMessages.map((message) => (
          <div
            key={message.id}
            className={`${styles.messageBubble} ${
              message.role === "user" ? styles.bubbleUser : styles.bubbleAssistant
            }`}
          >
            {message.content}
          </div>
        ))}

        {/* Stacking thinking steps — shown while busy */}
        {thinkingSteps.length > 0 &&
          thinkingSteps.map((step) => (
            <div key={step.id} className={styles.thinkingStep}>
              {step.status === "done" ? (
                <CheckmarkCircle16Filled className={styles.thinkingStepDone} />
              ) : (
                <Spinner size="extra-tiny" />
              )}
              <span
                className={
                  step.status === "done"
                    ? styles.thinkingStepDone
                    : styles.thinkingStepActive
                }
              >
                {step.text}
              </span>
            </div>
          ))}

        <div ref={messagesEndRef} />
      </div>

      <form className={styles.composer} onSubmit={handleSubmit}>
        <div className={styles.composerRow}>
          <textarea
            ref={textareaRef}
            className={styles.textarea}
            placeholder="Ask about your workbook… (Enter to send, Shift+Enter for newline)"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={handleKeyDown}
            rows={1}
            disabled={isBusy}
          />
          <Button
            appearance="primary"
            icon={<Send24Filled />}
            type="submit"
            disabled={isBusy || !input.trim()}
          >
            Send
          </Button>
        </div>
      </form>
    </div>
  );
}
