/**
 * Streaming tests for the =ASKAI custom function.
 * Uses mocked fetch to simulate NDJSON backend responses.
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { clearAskAICache, getAskAICache, setSharedProvider, setMutationHandler } from "../sharedState";
import { askAI } from "../functions/functions";

// ---------------------------------------------------------------------------
// Helpers to build NDJSON response streams
// ---------------------------------------------------------------------------

function ndjsonBody(events: object[]): ReadableStream<Uint8Array> {
  const encoder = new TextEncoder();
  const lines = events.map((e) => JSON.stringify(e) + "\n").join("");
  return new ReadableStream({
    start(controller) {
      controller.enqueue(encoder.encode(lines));
      controller.close();
    }
  });
}

function mockFetchResponse(events: object[], status = 200): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: new Headers({ "content-type": "application/x-ndjson" }),
    body: ndjsonBody(events),
    text: async () => "",
    json: async () => ({})
  } as unknown as Response;
}

function mockFetchErrorResponse(status: number): Response {
  return {
    ok: false,
    status,
    headers: new Headers({ "content-type": "text/plain" }),
    body: null,
    text: async () => "Server error",
    json: async () => ({})
  } as unknown as Response;
}

// ---------------------------------------------------------------------------
// Fake StreamingInvocation
// ---------------------------------------------------------------------------

function createInvocation(address?: string) {
  const results: unknown[] = [];
  let cancelHandler: (() => void) | null = null;
  return {
    invocation: {
      setResult(value: unknown) {
        results.push(value);
      },
      set onCanceled(fn: () => void) {
        cancelHandler = fn;
      },
      get onCanceled() {
        return cancelHandler;
      },
      ...(address !== undefined ? { address } : {})
    } as unknown as CustomFunctions.StreamingInvocation<string[][]>,
    results,
    cancel: () => cancelHandler?.()
  };
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------

beforeEach(() => {
  clearAskAICache();
  delete (window as any).__MYEXCELCOMPANION_PROVIDER;
  vi.restoreAllMocks();
});

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("askAI streaming function", () => {
  it("shows Thinking... then final answer on successful query", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "status", payload: "Processing..." },
        { type: "message_start", payload: { id: "1", role: "assistant", kind: "message", content: "", createdAt: "" } },
        { type: "message_delta", payload: { id: "1", delta: "42" } },
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "42", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation, results } = createInvocation();
    askAI("What is 6*7?", invocation);

    // Wait for async work
    await new Promise((r) => setTimeout(r, 50));

    expect(fetchSpy).toHaveBeenCalledOnce();
    // Should have: "Thinking...", "Processing...", "Generating response...", delta preview, final
    expect(results[0]).toEqual([["Thinking..."]]);
    expect(results[results.length - 1]).toEqual([["42"]]);
  });

  it("returns cached result immediately without fetching", () => {
    const cache = getAskAICache();
    cache.set("cached query", [["cached answer"]]);

    const fetchSpy = vi.spyOn(globalThis, "fetch");

    const { invocation, results } = createInvocation();
    askAI("cached query", invocation);

    // Cache hit is synchronous
    expect(results).toEqual([[["cached answer"]]]);
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("caches result after successful completion", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "Paris", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation } = createInvocation();
    askAI("capital of France?", invocation);

    await new Promise((r) => setTimeout(r, 50));

    const cache = getAskAICache();
    const key = Array.from(cache.keys()).find((k) => k.includes("capital"));
    expect(key).toBeDefined();
    expect(cache.get(key!)).toEqual([["Paris"]]);
  });

  it("clearAskAICache forces next call to fetch", async () => {
    const cache = getAskAICache();
    cache.set("q1", [["old"]]);

    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "new", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    clearAskAICache();

    const { invocation, results } = createInvocation();
    askAI("q1", invocation);

    await new Promise((r) => setTimeout(r, 50));

    // Should have fetched (cache was cleared)
    expect(results[0]).toEqual([["Thinking..."]]);
    expect(results[results.length - 1]).toEqual([["new"]]);
  });

  it("shows error on stream error event", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "error", payload: { message: "Rate limited" } }
      ])
    );

    const { invocation, results } = createInvocation();
    askAI("test error", invocation);

    await new Promise((r) => setTimeout(r, 50));

    const lastResult = results[results.length - 1];
    expect(lastResult).toEqual([["#ERROR: Rate limited"]]);
  });

  it("shows error on network failure", async () => {
    vi.spyOn(globalThis, "fetch").mockRejectedValue(new TypeError("Failed to fetch"));

    const { invocation, results } = createInvocation();
    askAI("test network", invocation);

    await new Promise((r) => setTimeout(r, 50));

    const lastResult = results[results.length - 1];
    expect(lastResult).toEqual([["#ERROR: Failed to fetch"]]);
  });

  it("shows error on HTTP 500", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(mockFetchErrorResponse(500));

    const { invocation, results } = createInvocation();
    askAI("test 500", invocation);

    await new Promise((r) => setTimeout(r, 50));

    const lastResult = results[results.length - 1];
    expect(lastResult).toEqual([["#ERROR: Backend error (500)"]]);
  });

  it("shows error on empty query", () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");

    const { invocation, results } = createInvocation();
    askAI("", invocation);

    expect(results).toEqual([[["#ERROR: Query is required"]]]);
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("reads provider from shared state", async () => {
    setSharedProvider("anthropic");

    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "ok", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation } = createInvocation();
    askAI("test provider", invocation);

    await new Promise((r) => setTimeout(r, 50));

    const body = JSON.parse(fetchSpy.mock.calls[0][1]?.body as string);
    expect(body.provider).toBe("anthropic");
  });

  // --- Spill integration ---

  it("returns 2D array for tabular response", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "A\tB\n1\t2", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation, results } = createInvocation();
    askAI("tabular data", invocation);

    await new Promise((r) => setTimeout(r, 50));

    expect(results[results.length - 1]).toEqual([["A", "B"], ["1", "2"]]);
  });

  it("returns vertical spill for multi-line response", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "first\nsecond\nthird", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation, results } = createInvocation();
    askAI("multi-line", invocation);

    await new Promise((r) => setTimeout(r, 50));

    expect(results[results.length - 1]).toEqual([["first"], ["second"], ["third"]]);
  });

  it("returns 1×1 array for single-line response", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "Just one line", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation, results } = createInvocation();
    askAI("single line", invocation);

    await new Promise((r) => setTimeout(r, 50));

    expect(results[results.length - 1]).toEqual([["Just one line"]]);
  });

  // --- Pivot / Chart answer override & address injection ---

  it("overrides verbose answer with brief confirmation for pivot mutations", async () => {
    const captured: unknown[] = [];
    setMutationHandler(async (m) => { captured.push(m); });

    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "Pivot table requested: Rows = Gender, Values = Toys (sum). It will be placed on the current sheet next to the data.", createdAt: "" } },
        { type: "pivot_table_inserts", payload: [{ sourceRange: "Sheet1!A1:C6", rows: ["Gender"], values: [{ field: "Toys", function: "sum" }] }] },
        { type: "done", payload: null }
      ])
    );

    const { invocation, results } = createInvocation("Sheet1!E8");
    askAI("create a pivot table", invocation);

    await new Promise((r) => setTimeout(r, 50));

    expect(results[results.length - 1]).toEqual([["Pivot table created."]]);
  });

  it("injects callerAddress into pivot destinationAddress and chart topLeftCell", async () => {
    const captured: unknown[] = [];
    setMutationHandler(async (m) => { captured.push(m); });

    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "Done", createdAt: "" } },
        { type: "pivot_table_inserts", payload: [{ sourceRange: "Sheet1!A1:C6", rows: ["Gender"], values: [{ field: "Toys", function: "sum" }] }] },
        { type: "chart_inserts", payload: [{ type: "bar", dataRange: "Sheet1!A1:C6", title: "Chart" }] },
        { type: "done", payload: null }
      ])
    );

    const { invocation } = createInvocation("Sheet1!G1");
    askAI("create stuff", invocation);

    await new Promise((r) => setTimeout(r, 50));

    expect(captured.length).toBe(1);
    const mutations = captured[0] as any;
    expect(mutations.pivotTableInserts[0].destinationAddress).toBe("Sheet1!G1");
    expect(mutations.chartInserts[0].topLeftCell).toBe("Sheet1!G1");
  });

  it("does NOT override answer when only cell updates are present", async () => {
    setMutationHandler(async () => {});

    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "I updated cells A1:A3 with new values.", createdAt: "" } },
        { type: "cell_updates", payload: [{ address: "A1", value: 10, worksheet: null }] },
        { type: "done", payload: null }
      ])
    );

    const { invocation, results } = createInvocation("Sheet1!E1");
    askAI("update cells", invocation);

    await new Promise((r) => setTimeout(r, 50));

    expect(results[results.length - 1]).toEqual([["I updated cells A1:A3 with new values."]]);
  });
});
