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
  Text
} from "@fluentui/react-components";
import { ProviderOption } from "../types";

interface SettingsDialogProps {
  open: boolean;
  providers: ProviderOption[];
  selectedProvider: string;
  onSelect: (providerId: string) => void;
  onClose: () => void;
}

export function SettingsDialog({
  open,
  providers,
  selectedProvider,
  onSelect,
  onClose
}: SettingsDialogProps) {
  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Model Provider</DialogTitle>
          <DialogContent>
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

