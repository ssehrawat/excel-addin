import {
  Button,
  Field,
  Input,
  Spinner,
  Tooltip,
  makeStyles,
  shorthands
} from "@fluentui/react-components";
import { Send24Filled, Settings24Regular } from "@fluentui/react-icons";
import { FormEvent, useState } from "react";
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
  headerTitle: {
    fontSize: "16px",
    fontWeight: 600
  },
  messages: {
    flex: 1,
    overflowY: "auto",
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px"
  },
  messageBubble: {
    ...shorthands.padding("12px", "14px"),
    borderRadius: "14px",
    boxShadow: "0 1px 2px rgba(0,0,0,0.08)",
    maxWidth: "90%"
  },
  bubbleUser: {
    alignSelf: "flex-end",
    backgroundColor: "#2563eb",
    color: "#ffffff"
  },
  bubbleAssistant: {
    alignSelf: "flex-start",
    backgroundColor: "#ffffff",
    color: "#111827",
    border: "1px solid #e5e7eb"
  },
  bubbleMeta: {
    fontSize: "11px",
    opacity: 0.7,
    marginBottom: "4px",
    textTransform: "uppercase",
    letterSpacing: "0.04em"
  },
  composer: {
    borderTop: "1px solid #e4e4e4",
    padding: "12px",
    backgroundColor: "#ffffff"
  },
  composerRow: {
    display: "flex",
    gap: "8px",
    alignItems: "center"
  }
});

interface ChatPanelProps {
  messages: ChatMessage[];
  isBusy: boolean;
  // eslint-disable-next-line no-unused-vars
  onSend: (text: string) => Promise<void>;
  onOpenSettings: () => void;
}

export function ChatPanel({
  messages,
  isBusy,
  onSend,
  onOpenSettings
}: ChatPanelProps) {
  const styles = useStyles();
  const [input, setInput] = useState("");

  const handleSubmit = async (event: FormEvent) => {
    event.preventDefault();
    const trimmed = input.trim();
    if (!trimmed || isBusy) {
      return;
    }
    setInput("");
    await onSend(trimmed);
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <div className={styles.headerTitle}>Workbook Copilot</div>
        <div style={{ display: "flex", gap: "8px" }}>
          {isBusy ? (
            <Spinner size="tiny" label="Thinking..." />
          ) : (
            <Tooltip content="Change model provider" relationship="label">
              <Button icon={<Settings24Regular />} onClick={onOpenSettings} />
            </Tooltip>
          )}
        </div>
      </div>
      <div className={styles.messages}>
        {messages.map((message) => (
          <div
            key={message.id}
            className={`${styles.messageBubble} ${
              message.role === "user"
                ? styles.bubbleUser
                : styles.bubbleAssistant
            }`}
          >
            <div className={styles.bubbleMeta}>
              {message.role === "user" ? "You" : formatKind(message.kind)}
            </div>
            <div>{message.content}</div>
          </div>
        ))}
        {isBusy && (
          <div className={`${styles.messageBubble} ${styles.bubbleAssistant}`}>
            <div className={styles.bubbleMeta}>Assistant</div>
            <div>Processing your request…</div>
          </div>
        )}
      </div>
      <form className={styles.composer} onSubmit={handleSubmit}>
        <Field>
          <div className={styles.composerRow}>
            <Input
              appearance="outline"
              placeholder="Ask about your workbook…"
              value={input}
              onChange={(_, data) => setInput(data.value)}
              size="large"
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
        </Field>
      </form>
    </div>
  );
}

function formatKind(kind: ChatMessage["kind"]): string {
  switch (kind) {
    case "thought":
      return "Thought";
    case "step":
      return "Step";
    case "suggestion":
      return "Suggestion";
    case "context":
      return "Context";
    case "final":
      return "Answer";
    default:
      return "Assistant";
  }
}

