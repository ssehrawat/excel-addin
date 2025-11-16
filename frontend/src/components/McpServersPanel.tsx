import {
  Badge,
  Button,
  Field,
  Input,
  Switch,
  Text,
  Textarea,
  makeStyles,
  shorthands,
  Spinner
} from "@fluentui/react-components";
import {
  ArrowClockwise24Regular,
  Delete16Regular
} from "@fluentui/react-icons";
import { FormEvent, useMemo, useState } from "react";
import {
  CreateMCPServerPayload,
  MCPServer,
  MCPServerStatus
} from "../types";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "16px"
  },
  form: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
    gap: "12px"
  },
  serverCard: {
    ...shorthands.border("1px", "solid", "#e5e7eb"),
    borderRadius: "12px",
    padding: "12px 14px",
    backgroundColor: "#ffffff",
    display: "flex",
    flexDirection: "column",
    gap: "8px"
  },
  serverHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "8px"
  },
  toolsList: {
    margin: 0,
    paddingLeft: "18px",
    color: "#4b5563",
    fontSize: "13px"
  },
  buttonRow: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap"
  },
  emptyState: {
    padding: "12px",
    borderRadius: "8px",
    backgroundColor: "#f9fafb",
    color: "#4b5563",
    fontSize: "13px"
  },
  error: {
    color: "#b01030",
    fontSize: "13px"
  }
});

interface McpServersPanelProps {
  servers: MCPServer[];
  isLoading: boolean;
  busyIds: string[];
  error?: string | null;
  // eslint-disable-next-line no-unused-vars
  onCreate: (payload: CreateMCPServerPayload) => Promise<void>;
  // eslint-disable-next-line no-unused-vars
  onToggle: (serverId: string, enabled: boolean) => Promise<void>;
  // eslint-disable-next-line no-unused-vars
  onRefresh: (serverId: string) => Promise<void>;
  // eslint-disable-next-line no-unused-vars
  onDelete: (serverId: string) => Promise<void>;
}

const initialFormState = {
  name: "",
  baseUrl: "",
  apiKey: "",
  description: ""
};

const STATUS_COLORS: Record<MCPServerStatus, "success" | "danger" | "warning"> = {
    online: "success",
    offline: "danger",
    error: "danger",
    unknown: "warning"
  };

export function McpServersPanel({
  servers,
  isLoading,
  busyIds,
  error,
  onCreate,
  onToggle,
  onRefresh,
  onDelete
}: McpServersPanelProps) {
  const styles = useStyles();
  const [form, setForm] = useState(initialFormState);
  const [submitting, setSubmitting] = useState(false);

  const busySet = useMemo(() => new Set(busyIds), [busyIds]);

  const handleSubmit = async (event: FormEvent) => {
    event.preventDefault();
    const name = form.name.trim();
    const baseUrl = form.baseUrl.trim();
    if (!name || !baseUrl) {
      return;
    }
    setSubmitting(true);
    try {
      await onCreate({
        name,
        baseUrl,
        description: form.description.trim() || undefined,
        apiKey: form.apiKey.trim() || undefined,
        enabled: true,
        autoRefresh: true
      });
      setForm(initialFormState);
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className={styles.root}>
      <form className={styles.form} onSubmit={handleSubmit}>
        <Field label="Server name">
          <Input
            value={form.name}
            onChange={(_, data) => setForm((prev) => ({ ...prev, name: data.value }))}
            required
          />
        </Field>
        <Field label="Base URL">
          <Input
            value={form.baseUrl}
            onChange={(_, data) =>
              setForm((prev) => ({ ...prev, baseUrl: data.value }))
            }
            required
            placeholder="https://localhost:9000"
          />
        </Field>
        <Field label="API key (optional)">
          <Input
            value={form.apiKey}
            onChange={(_, data) =>
              setForm((prev) => ({ ...prev, apiKey: data.value }))
            }
            type="password"
            placeholder="Token used when calling the server"
          />
        </Field>
        <Field
          label="Description (optional)"
          style={{ gridColumn: "1 / -1" }}
        >
          <Textarea
            value={form.description}
            onChange={(_, data) =>
              setForm((prev) => ({ ...prev, description: data.value }))
            }
            resize="vertical"
          />
        </Field>
        <div>
          <Button
            appearance="primary"
            type="submit"
            disabled={submitting}
          >
            {submitting ? "Adding…" : "Add MCP server"}
          </Button>
        </div>
      </form>
      {error && (
        <Text role="alert" className={styles.error}>
          {error}
        </Text>
      )}
      {isLoading ? (
        <div className={styles.emptyState}>
          <Spinner size="tiny" />
          &nbsp;Loading MCP servers…
        </div>
      ) : servers.length === 0 ? (
        <div className={styles.emptyState}>
          No MCP servers configured yet. Add one to enrich Workbook Copilot with
          your own tools.
        </div>
      ) : (
        servers.map((server) => {
          const busy = busySet.has(server.id);
          return (
            <div key={server.id} className={styles.serverCard}>
              <div className={styles.serverHeader}>
                <div>
                  <Text weight="semibold">{server.name}</Text>
                  <Text block size={200}>
                    {server.baseUrl}
                  </Text>
                </div>
                <Switch
                  checked={server.enabled}
                  disabled={busy}
                  onChange={(_, data) => onToggle(server.id, data.checked)}
                />
              </div>
              <div className={styles.buttonRow}>
                <Badge
                  appearance="filled"
                  color={STATUS_COLORS[server.status]}
                  size="small"
                >
                  {server.status}
                </Badge>
                {server.lastRefreshedAt && (
                  <Text size={200}>
                    Refreshed {new Date(server.lastRefreshedAt).toLocaleString()}
                  </Text>
                )}
              </div>
              {server.description && (
                <Text size={200} wrap>
                  {server.description}
                </Text>
              )}
              <div className={styles.buttonRow}>
                <Button
                  icon={<ArrowClockwise24Regular />}
                  appearance="subtle"
                  size="small"
                  onClick={() => onRefresh(server.id)}
                  disabled={busy}
                >
                  Refresh tools
                </Button>
                <Button
                  icon={<Delete16Regular />}
                  appearance="outline"
                  size="small"
                  onClick={() => onDelete(server.id)}
                  disabled={busy}
                >
                  Remove
                </Button>
              </div>
              {server.enabled && server.tools.length > 0 && (
                <div>
                  <Text weight="semibold" size={200}>
                    Tools
                  </Text>
                  <ul className={styles.toolsList}>
                    {server.tools.map((tool) => (
                      <li key={`${server.id}-${tool.name}`}>
                        <Text weight="semibold">{tool.name}</Text>
                        {tool.description && (
                          <Text>&nbsp;— {tool.description}</Text>
                        )}
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          );
        })
      )}
    </div>
  );
}

