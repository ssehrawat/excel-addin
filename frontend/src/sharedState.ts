/**
 * Cross-bundle shared state using window globals.
 *
 * The taskpane and custom-function bundles run in the same shared runtime
 * but are separate webpack chunks — they cannot share module instances.
 * Window globals bridge the two bundles.
 */

/* eslint-disable @typescript-eslint/no-explicit-any */

import { CellUpdate, FormatUpdate, ChartInsert, PivotTableInsert } from "./types";

/** Mutations collected by =ASKAI() to be applied via the taskpane bridge. */
export interface PendingMutations {
  cellUpdates?: CellUpdate[];
  formatUpdates?: FormatUpdate[];
  chartInserts?: ChartInsert[];
  pivotTableInserts?: PivotTableInsert[];
}

declare global {
  interface Window {
    __MYEXCELCOMPANION_PROVIDER?: string;
    __MYEXCELCOMPANION_CACHE?: Map<string, string[][]>;
    __MYEXCELCOMPANION_APPLY_MUTATIONS?: (mutations: PendingMutations) => Promise<void>;
  }
}

/**
 * Returns the currently-selected LLM provider.
 * Falls back to "mock" if none has been set by the taskpane yet.
 */
export function getSharedProvider(): string {
  return window.__MYEXCELCOMPANION_PROVIDER ?? "mock";
}

/**
 * Sets the LLM provider. Called by the taskpane whenever the user
 * switches providers or on initial load.
 * @param provider - Provider id (e.g. "mock", "openai", "anthropic")
 */
export function setSharedProvider(provider: string): void {
  window.__MYEXCELCOMPANION_PROVIDER = provider;
}

/**
 * Returns the shared result cache map, creating it on first access.
 */
export function getAskAICache(): Map<string, string[][]> {
  if (!window.__MYEXCELCOMPANION_CACHE) {
    window.__MYEXCELCOMPANION_CACHE = new Map();
  }
  return window.__MYEXCELCOMPANION_CACHE;
}

/**
 * Clears all cached =ASKAI results.
 * After clearing, the user should press Ctrl+Shift+F9 to force recalculation.
 */
export function clearAskAICache(): void {
  if (window.__MYEXCELCOMPANION_CACHE) {
    window.__MYEXCELCOMPANION_CACHE.clear();
  }
}

/**
 * Registers a mutation handler callback on the window global.
 * Called by the taskpane on mount so that custom functions can dispatch
 * Excel mutations (cell updates, charts, pivot tables) via the bridge.
 * @param handler - Async function that applies mutations using Excel.run()
 */
export function setMutationHandler(handler: (m: PendingMutations) => Promise<void>): void {
  window.__MYEXCELCOMPANION_APPLY_MUTATIONS = handler;
}

/**
 * Returns the mutation handler registered by the taskpane, or undefined
 * if the taskpane has not yet mounted.
 */
export function getMutationHandler(): ((m: PendingMutations) => Promise<void>) | undefined {
  return window.__MYEXCELCOMPANION_APPLY_MUTATIONS;
}
