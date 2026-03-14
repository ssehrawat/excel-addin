/**
 * Tests for the shared state module that bridges the taskpane and
 * custom function bundles via window globals.
 */

import { describe, it, expect, beforeEach } from "vitest";
import {
  getSharedProvider,
  setSharedProvider,
  getAskAICache,
  clearAskAICache
} from "../sharedState";

describe("sharedState", () => {
  beforeEach(() => {
    // Reset globals between tests
    delete (window as any).__MYEXCELCOMPANION_PROVIDER;
    delete (window as any).__MYEXCELCOMPANION_CACHE;
  });

  // --- Provider ---

  it("returns default 'mock' when no provider set", () => {
    expect(getSharedProvider()).toBe("mock");
  });

  it("setSharedProvider → getSharedProvider round-trip", () => {
    setSharedProvider("anthropic");
    expect(getSharedProvider()).toBe("anthropic");
  });

  it("overwrites previous provider value", () => {
    setSharedProvider("openai");
    setSharedProvider("anthropic");
    expect(getSharedProvider()).toBe("anthropic");
  });

  it("uses window global as backing store", () => {
    setSharedProvider("openai");
    expect(window.__MYEXCELCOMPANION_PROVIDER).toBe("openai");
  });

  // --- Cache ---

  it("getAskAICache creates a Map on first access", () => {
    const cache = getAskAICache();
    expect(cache).toBeInstanceOf(Map);
    expect(cache.size).toBe(0);
  });

  it("cache persists across multiple get calls", () => {
    const cache1 = getAskAICache();
    cache1.set("key1", [["val"]]);
    const cache2 = getAskAICache();
    expect(cache2.get("key1")).toEqual([["val"]]);
    expect(cache1).toBe(cache2); // same instance
  });

  it("clearAskAICache empties the cache", () => {
    const cache = getAskAICache();
    cache.set("a", [["1"]]);
    cache.set("b", [["2"]]);
    clearAskAICache();
    expect(getAskAICache().size).toBe(0);
  });

  it("clearAskAICache is safe when cache was never created", () => {
    // Should not throw
    clearAskAICache();
    expect(getAskAICache().size).toBe(0);
  });
});
