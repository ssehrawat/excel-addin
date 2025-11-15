/* global process */

import { ProviderOption } from "./types";

export const API_BASE_URL =
  process.env.API_BASE_URL ?? "https://localhost:8000";

export const DEFAULT_PROVIDER =
  process.env.DEFAULT_PROVIDER ?? "mock";

export const FALLBACK_PROVIDERS: ProviderOption[] = [
  {
    id: "mock",
    label: "Mock (no external calls)",
    description: "Deterministic responses useful for development and demos.",
    requiresKey: false
  },
  {
    id: "openai",
    label: "OpenAI",
    description:
      "Use OpenAI compatible models such as GPT-4o. Requires an API key on the backend.",
    requiresKey: true
  },
  {
    id: "anthropic",
    label: "Anthropic",
    description: "Use Claude models. Requires API key on the backend.",
    requiresKey: true
  }
];

