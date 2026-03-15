/**
 * =MYEXCELCOMPANION.ASKAI() custom function implementation.
 *
 * Non-streaming async custom function that queries the same FastAPI /chat
 * backend as the taskpane chatbot. Supports variadic range arguments, result
 * caching, and 2D spill output. Participates in Excel's standard calculation
 * engine so F2+Enter and Ctrl+Shift+F9 trigger fresh API calls.
 */

import { API_BASE_URL } from "../config";
import { getSharedProvider, getAskAICache, getMutationHandler, PendingMutations } from "../sharedState";
import { ChatStreamEvent } from "../types";

// ---------------------------------------------------------------------------
// Pure helpers (exported for testing)
// ---------------------------------------------------------------------------

/**
 * Parse a delimited line respecting quoted fields.
 * Handles fields wrapped in double-quotes that may contain the delimiter
 * or escaped double-quotes ("").
 * @param line - A single line of text
 * @param delimiter - The field delimiter character
 * @returns Array of field values
 */
export function parseDelimitedLine(line: string, delimiter: string): string[] {
  const fields: string[] = [];
  let current = "";
  let inQuotes = false;
  let i = 0;

  while (i < line.length) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') {
          current += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      current += ch;
      i++;
    } else {
      if (ch === '"') {
        inQuotes = true;
        i++;
      } else if (ch === delimiter) {
        fields.push(current);
        current = "";
        i++;
      } else {
        current += ch;
        i++;
      }
    }
  }
  fields.push(current);
  return fields;
}

/**
 * Parse an AI answer string into a 2D string array for Excel spill output.
 *
 * Rules:
 * 1. Single line, no delimiters → `[[text]]` (1×1)
 * 2. Multiple lines with tabs → split by tab → N×M matrix
 * 3. Multiple lines with commas (no tabs) → CSV parse → N×M matrix
 * 4. Multiple lines, no delimiters → `[[line1],[line2],...]` (N×1 vertical spill)
 * 5. Short rows are padded with empty strings to the max column count.
 *
 * @param text - The raw answer text from the LLM
 * @returns 2D string array suitable for Excel spill
 */
export function parseAnswerTo2D(text: string): string[][] {
  const trimmed = text.trim();
  if (!trimmed) return [[""]];

  const lines = trimmed.split("\n");
  if (lines.length === 1) {
    // Check for tab or comma delimited single line
    if (lines[0].includes("\t")) {
      return [parseDelimitedLine(lines[0], "\t")];
    }
    return [[lines[0]]];
  }

  // Multiple lines — detect delimiter
  const hasTabs = lines.some((l) => l.includes("\t"));
  if (hasTabs) {
    const rows = lines.map((l) => parseDelimitedLine(l, "\t"));
    return padRows(rows);
  }

  const hasCommas = lines.some((l) => l.includes(","));
  if (hasCommas) {
    const rows = lines.map((l) => parseDelimitedLine(l, ","));
    return padRows(rows);
  }

  // Plain multi-line text → vertical spill (N×1)
  return lines.map((l) => [l]);
}

/**
 * Pad all rows to the same column count with empty strings.
 * @param rows - 2D array of strings
 * @returns Padded 2D array
 */
function padRows(rows: string[][]): string[][] {
  const maxCols = Math.max(...rows.map((r) => r.length));
  return rows.map((r) => {
    while (r.length < maxCols) r.push("");
    return r;
  });
}

/**
 * Compute a deterministic cache key from the query and range data.
 * @param query - The user's question
 * @param ranges - Array of range matrices
 * @returns Cache key string
 */
export function computeCacheKey(query: string, ranges: unknown[][][]): string {
  const parts: string[] = [query];
  for (let ri = 0; ri < ranges.length; ri++) {
    parts.push(`__RANGE${ri}__`);
    const matrix = ranges[ri];
    for (const row of matrix) {
      parts.push(row.map((v) => (v == null ? "" : String(v))).join("\t"));
    }
  }
  return parts.join("\n");
}

/**
 * Serialize range matrices into a CSV-like string for the POST body.
 * Multiple ranges are separated by a blank line.
 * @param ranges - Array of range matrices
 * @returns Serialized context string
 */
export function rangesToContext(ranges: unknown[][][]): string {
  if (ranges.length === 0) return "";

  const sections: string[] = [];
  for (const matrix of ranges) {
    const lines: string[] = [];
    for (const row of matrix) {
      const cells = row.map((v) => {
        const s = v == null ? "" : String(v);
        // Quote fields that contain commas, newlines, or double quotes
        if (s.includes(",") || s.includes("\n") || s.includes('"')) {
          return '"' + s.replace(/"/g, '""') + '"';
        }
        return s;
      });
      lines.push(cells.join(","));
    }
    sections.push(lines.join("\n"));
  }
  return sections.join("\n\n");
}

// ---------------------------------------------------------------------------
// Core async custom function
// ---------------------------------------------------------------------------

/**
 * =MYEXCELCOMPANION.ASKAI(query, [range1], [range2], ...)
 *
 * Non-streaming async custom function that sends a query to the LLM backend,
 * optionally with cell data as context. Returns a 2D result that spills into
 * adjacent cells. Participates in Excel's standard recalculation so F2+Enter
 * and Ctrl+Shift+F9 trigger fresh API calls.
 *
 * @param query - The question or instruction
 * @param args - Variadic range matrices followed by the CancelableInvocation
 */
function askAI(query: string, ...args: unknown[]): Promise<string[][]> {
  // Office runtime appends the CancelableInvocation as the last argument
  const invocation = args.pop() as CustomFunctions.CancelableInvocation;
  const callerAddress = (invocation as any).address as string | undefined;
  console.debug("[ASKAI] callerAddress from invocation:", callerAddress);

  // Remaining args are range matrices (each is unknown[][])
  const ranges = args as unknown[][][];

  // Validate query
  if (!query || (typeof query === "string" && !query.trim())) {
    return Promise.resolve([["#ERROR: Query is required"]]);
  }

  // Fingerprint-based cache: distinguish auto-recalc (input change) from
  // manual recalc (F2+Enter / Ctrl+Shift+F9).
  const cache = getAskAICache();
  const cacheKey = (callerAddress ?? "") + "||" + query;
  const currentFingerprint = computeCacheKey("", ranges);

  const cached = cache.get(cacheKey);
  if (cached) {
    if (cached.rangeFingerprint !== currentFingerprint) {
      // Input data changed → auto-recalc → return cached result, update fingerprint
      cached.rangeFingerprint = currentFingerprint;
      return Promise.resolve(cached.result);
    }
    // Same inputs → manual recalc (F2+Enter / Ctrl+Shift+F9) → evict & re-fetch
    cache.delete(cacheKey);
  }

  // Set up cancellation
  const controller = new AbortController();

  invocation.onCanceled = () => {
    controller.abort();
  };

  // Build request payload
  const contextStr = rangesToContext(ranges);
  const selection = contextStr
    ? [{ address: "ASKAI_INPUT", values: ranges.length > 0 ? ranges[0] as (string | number | boolean | null)[][] : [], worksheet: null }]
    : [];

  const body = JSON.stringify({
    prompt: contextStr ? `${query}\n\nContext data:\n${contextStr}` : query,
    provider: getSharedProvider(),
    messages: [],
    selection
  });

  // Execute fetch and return result
  return (async () => {
    try {
      const response = await fetch(`${API_BASE_URL}/chat`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/x-ndjson"
        },
        body,
        signal: controller.signal
      });

      if (!response.ok) {
        return [[`#ERROR: Backend error (${response.status})`]];
      }

      const contentType = response.headers.get("content-type") ?? "";
      let finalAnswer = "";
      let errorOccurred = false;
      let errorResult: string[][] | null = null;
      const pendingMutations: PendingMutations = {};

      if (contentType.includes("application/x-ndjson") && response.body) {
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";
        let reading = true;

        const processEvent = (event: ChatStreamEvent) => {
          switch (event.type) {
            case "message_delta": {
              finalAnswer += event.payload.delta;
              break;
            }
            case "message_done":
            case "message":
              if (event.payload?.content) {
                finalAnswer = event.payload.content;
              }
              break;
            case "error":
              errorOccurred = true;
              errorResult = [[`#ERROR: ${event.payload?.message ?? "Unknown error"}`]];
              break;
            case "cell_updates":
              if (event.payload?.length)
                pendingMutations.cellUpdates = (pendingMutations.cellUpdates ?? []).concat(event.payload);
              break;
            case "format_updates":
              if (event.payload?.length)
                pendingMutations.formatUpdates = (pendingMutations.formatUpdates ?? []).concat(event.payload);
              break;
            case "chart_inserts":
              if (event.payload?.length)
                pendingMutations.chartInserts = (pendingMutations.chartInserts ?? []).concat(event.payload);
              break;
            case "pivot_table_inserts":
              if (event.payload?.length)
                pendingMutations.pivotTableInserts = (pendingMutations.pivotTableInserts ?? []).concat(event.payload);
              break;
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
              try {
                processEvent(JSON.parse(line) as ChatStreamEvent);
              } catch {
                // Skip malformed lines
              }
            }
            newlineIndex = buffer.indexOf("\n");
          }
          if (flush) {
            const remaining = buffer.trim();
            buffer = "";
            if (remaining) {
              try {
                processEvent(JSON.parse(remaining) as ChatStreamEvent);
              } catch {
                // Skip malformed lines
              }
            }
          }
        };

        while (reading) {
          const { value, done } = await reader.read();
          if (value) {
            buffer += decoder.decode(value, { stream: true });
            drainBuffer(false);
          }
          if (done) reading = false;
        }
        buffer += decoder.decode();
        drainBuffer(true);
      } else {
        // Non-streaming fallback
        const json = await response.json();
        if (json.messages?.length) {
          const last = json.messages[json.messages.length - 1];
          finalAnswer = last.content ?? "";
        }
      }

      // Override verbose LLM answer for pivot/chart mutations
      const hasPivotOrChart =
        (pendingMutations.pivotTableInserts?.length ?? 0) > 0 ||
        (pendingMutations.chartInserts?.length ?? 0) > 0;
      if (hasPivotOrChart && finalAnswer) {
        const parts: string[] = [];
        if (pendingMutations.pivotTableInserts?.length) {
          const n = pendingMutations.pivotTableInserts.length;
          parts.push(n === 1 ? "Pivot table created" : `${n} pivot tables created`);
        }
        if (pendingMutations.chartInserts?.length) {
          const n = pendingMutations.chartInserts.length;
          parts.push(n === 1 ? "Chart created" : `${n} charts created`);
        }
        finalAnswer = parts.join("; ") + ".";
      }

      // Determine final result (skip if error already captured)
      let result: string[][];
      if (errorOccurred) {
        result = errorResult!;
      } else if (finalAnswer) {
        result = parseAnswerTo2D(finalAnswer);
        cache.set(cacheKey, { result, rangeFingerprint: currentFingerprint });
      } else {
        result = [[""]];
      }

      // Inject formula cell address as default destination for pivot tables and charts.
      // Fall back to active cell when callerAddress is unavailable (e.g. stale Office cache).
      let resolvedAddress = callerAddress;
      if (!resolvedAddress && hasPivotOrChart && typeof Excel !== "undefined") {
        try {
          resolvedAddress = await Excel.run(async (ctx) => {
            const cell = ctx.workbook.getActiveCell();
            const ws = ctx.workbook.worksheets.getActiveWorksheet();
            cell.load("address"); ws.load("name");
            await ctx.sync();
            return `${ws.name}!${cell.address.split("!").pop()}`;
          });
        } catch { /* ignore — best-effort fallback */ }
      }
      if (resolvedAddress) {
        for (const pt of pendingMutations.pivotTableInserts ?? []) {
          if (!pt.destinationAddress) pt.destinationAddress = resolvedAddress;
        }
        for (const chart of pendingMutations.chartInserts ?? []) {
          if (!chart.topLeftCell) chart.topLeftCell = resolvedAddress;
        }
      }

      // Apply any collected mutations via the taskpane bridge
      const handler = getMutationHandler();
      const hasMutations = pendingMutations.cellUpdates?.length ||
        pendingMutations.formatUpdates?.length ||
        pendingMutations.chartInserts?.length ||
        pendingMutations.pivotTableInserts?.length;
      if (handler && hasMutations) {
        try { await handler(pendingMutations); } catch (e) { console.warn("ASKAI mutation apply failed:", e); }
      }

      return result;
    } catch (err: unknown) {
      if (err instanceof DOMException && err.name === "AbortError") {
        // Request was cancelled — return empty to avoid caching
        return [[""]];
      }
      const message = err instanceof Error ? err.message : "Cannot reach backend";
      return [[`#ERROR: ${message}`]];
    }
  })();
}

// Export for testing
export { askAI };

// Registration is handled by taskpane.tsx inside Office.onReady() (shared runtime).
// Do NOT call CustomFunctions.associate() here — the import into taskpane.tsx
// causes this module to execute in the same runtime, and a duplicate associate()
// call triggers [DuplicatedName] warnings that can break registration.
