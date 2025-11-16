import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Field,
  Radio,
  RadioGroup,
  Tab,
  TabList,
  TabPanel,
  TabPanels,
  Text
} from "@fluentui/react-components";
import { useState } from "react";
import { McpServersPanel } from "./McpServersPanel";
import {
  CreateMCPServerPayload,
  MCPServer,
  ProviderOption
} from "../types";

interface SettingsDialogProps {
  open: boolean;
  providers: ProviderOption[];
  selectedProvider: string;
  // eslint-disable-next-line no-unused-vars
  onSelect: (providerId: string) => void;
  onClose: () => void;
  mcpServers: MCPServer[];
  mcpServersLoading: boolean;
  mcpBusyIds: string[];
  mcpError?: string | null;
  // eslint-disable-next-line no-unused-vars
  onCreateMcpServer: (payload: CreateMCPServerPayload) => Promise<void>;
  // eslint-disable-next-line no-unused-vars
  onToggleMcpServer: (serverId: string, enabled: boolean) => Promise<void>;
  // eslint-disable-next-line no-unused-vars
  onRefreshMcpServer: (serverId: string) => Promise<void>;
  // eslint-disable-next-line no-unused-vars
  onDeleteMcpServer: (serverId: string) => Promise<void>;
}

export function SettingsDialog({
  open,
  providers,
  selectedProvider,
  onSelect,
  onClose,
  mcpServers,
  mcpServersLoading,
  mcpBusyIds,
  mcpError,
  onCreateMcpServer,
  onToggleMcpServer,
  onRefreshMcpServer,
  onDeleteMcpServer
}: SettingsDialogProps) {
  const [activeTab, setActiveTab] = useState<"providers" | "mcp">("providers");

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Workbook Copilot Settings</DialogTitle>
          <DialogContent>
            <TabList
              selectedValue={activeTab}
              onTabSelect={(_, data) =>
                setActiveTab(data.value as "providers" | "mcp")
              }
              appearance="subtle"
            >
              <Tab value="providers">Providers</Tab>
              <Tab value="mcp">MCP servers</Tab>
            </TabList>
            <TabPanels>
              <TabPanel value="providers">
                <Field label="Choose which backend provider should handle chat requests.">
                  <RadioGroup
                    value={selectedProvider}
                    onChange={(_, data) => onSelect(data.value)}
                  >
                    {providers.map((provider) => (
                      <Radio
                        key={provider.id}
                        value={provider.id}
                        label={
                          <div>
                            <Text weight="semibold">{provider.label}</Text>
                            <Text block size={200}>
                              {provider.description}
                            </Text>
                            {provider.requiresKey && (
                              <Text block size={200} italic>
                                Requires API key configured on backend.
                              </Text>
                            )}
                          </div>
                        }
                      />
                    ))}
                  </RadioGroup>
                </Field>
              </TabPanel>
              <TabPanel value="mcp">
                <McpServersPanel
                  servers={mcpServers}
                  isLoading={mcpServersLoading}
                  busyIds={mcpBusyIds}
                  error={mcpError}
                  onCreate={onCreateMcpServer}
                  onToggle={onToggleMcpServer}
                  onRefresh={onRefreshMcpServer}
                  onDelete={onDeleteMcpServer}
                />
              </TabPanel>
            </TabPanels>
          </DialogContent>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement action="close">
              <Button appearance="secondary" onClick={onClose}>
                Close
              </Button>
            </DialogTrigger>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

