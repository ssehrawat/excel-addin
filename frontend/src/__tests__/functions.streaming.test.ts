/**
 * Tests for the =ASKAI async custom function.
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
// Fake CancelableInvocation
// ---------------------------------------------------------------------------

function createInvocation(address?: string) {
  let cancelHandler: (() => void) | null = null;
  return {
    invocation: {
      set onCanceled(fn: () => void) {
        cancelHandler = fn;
      },
      get onCanceled() {
        return cancelHandler;
      },
      ...(address !== undefined ? { address } : {})
    } as unknown as CustomFunctions.CancelableInvocation,
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

describe("askAI async function", () => {
  it("returns final answer on successful query", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "status", payload: "Processing..." },
        { type: "message_start", payload: { id: "1", role: "assistant", kind: "message", content: "", createdAt: "" } },
        { type: "message_delta", payload: { id: "1", delta: "42" } },
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "42", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation } = createInvocation();
    const result = await askAI("What is 6*7?", invocation);

    expect(fetchSpy).toHaveBeenCalledOnce();
    expect(result).toEqual([["42"]]);
  });

  it("returns cached result immediately without fetching (different fingerprint → cache hit)", async () => {
    const cache = getAskAICache();
    // Pre-seed cache with a different fingerprint so the lookup sees "input data changed" → returns cached
    cache.set("||cached query", { result: [["cached answer"]], rangeFingerprint: "OLD_FP" });

    const fetchSpy = vi.spyOn(globalThis, "fetch");

    const { invocation } = createInvocation();
    // No ranges → currentFingerprint will be "" which differs from "OLD_FP"
    const result = await askAI("cached query", invocation);

    // Cache hit (fingerprint mismatch → auto-recalc path → return cached)
    expect(result).toEqual([["cached answer"]]);
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
    await askAI("capital of France?", invocation);

    const cache = getAskAICache();
    const key = Array.from(cache.keys()).find((k) => k.includes("capital"));
    expect(key).toBeDefined();
    expect(cache.get(key!)?.result).toEqual([["Paris"]]);
  });

  it("clearAskAICache forces next call to fetch", async () => {
    const cache = getAskAICache();
    cache.set("||q1", { result: [["old"]], rangeFingerprint: "" });

    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "new", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    clearAskAICache();

    const { invocation } = createInvocation();
    const result = await askAI("q1", invocation);

    // Should have fetched (cache was cleared)
    expect(result).toEqual([["new"]]);
  });

  it("returns cached result when range data changes, re-fetches on manual recalc (F2+Enter)", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "100", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    // First call — fetches from backend and caches with fingerprint of [["1","2","3"]]
    const { invocation: inv1 } = createInvocation("Sheet1!B1");
    const result1 = await askAI("sum this", [["1", "2", "3"]], inv1);
    expect(fetchSpy).toHaveBeenCalledOnce();
    expect(result1).toEqual([["100"]]);

    fetchSpy.mockClear();

    // Second call — same query & address but DIFFERENT range data
    // Fingerprint mismatch → auto-recalc → returns cached result, updates fingerprint
    const { invocation: inv2 } = createInvocation("Sheet1!B1");
    const result2 = await askAI("sum this", [["4", "5", "6"]], inv2);

    expect(result2).toEqual([["100"]]);
    expect(fetchSpy).not.toHaveBeenCalled();

    fetchSpy.mockClear();
    fetchSpy.mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "200", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    // Third call — same query, address, AND same range data as second call
    // Fingerprint matches (updated in step 2) → manual recalc → evict & re-fetch
    const { invocation: inv3 } = createInvocation("Sheet1!B1");
    const result3 = await askAI("sum this", [["4", "5", "6"]], inv3);

    expect(fetchSpy).toHaveBeenCalledOnce();
    expect(result3).toEqual([["200"]]);
  });

  it("returns error on stream error event", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "error", payload: { message: "Rate limited" } }
      ])
    );

    const { invocation } = createInvocation();
    const result = await askAI("test error", invocation);

    expect(result).toEqual([["#ERROR: Rate limited"]]);
  });

  it("returns error on network failure", async () => {
    vi.spyOn(globalThis, "fetch").mockRejectedValue(new TypeError("Failed to fetch"));

    const { invocation } = createInvocation();
    const result = await askAI("test network", invocation);

    expect(result).toEqual([["#ERROR: Failed to fetch"]]);
  });

  it("returns error on HTTP 500", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(mockFetchErrorResponse(500));

    const { invocation } = createInvocation();
    const result = await askAI("test 500", invocation);

    expect(result).toEqual([["#ERROR: Backend error (500)"]]);
  });

  it("returns error on empty query", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");

    const { invocation } = createInvocation();
    const result = await askAI("", invocation);

    expect(result).toEqual([["#ERROR: Query is required"]]);
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
    await askAI("test provider", invocation);

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

    const { invocation } = createInvocation();
    const result = await askAI("tabular data", invocation);

    expect(result).toEqual([["A", "B"], ["1", "2"]]);
  });

  it("returns vertical spill for multi-line response", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "first\nsecond\nthird", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation } = createInvocation();
    const result = await askAI("multi-line", invocation);

    expect(result).toEqual([["first"], ["second"], ["third"]]);
  });

  it("returns 1×1 array for single-line response", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      mockFetchResponse([
        { type: "message_done", payload: { id: "1", role: "assistant", kind: "final", content: "Just one line", createdAt: "" } },
        { type: "done", payload: null }
      ])
    );

    const { invocation } = createInvocation();
    const result = await askAI("single line", invocation);

    expect(result).toEqual([["Just one line"]]);
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

    const { invocation } = createInvocation("Sheet1!E8");
    const result = await askAI("create a pivot table", invocation);

    expect(result).toEqual([["Pivot table created."]]);
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
    await askAI("create stuff", invocation);

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

    const { invocation } = createInvocation("Sheet1!E1");
    const result = await askAI("update cells", invocation);

    expect(result).toEqual([["I updated cells A1:A3 with new values."]]);
  });
});
