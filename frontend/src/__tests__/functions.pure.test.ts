/**
 * Pure function unit tests for the =ASKAI custom function helpers.
 * No Office.js or fetch mocking needed — these are pure data transformations.
 */

import { describe, it, expect } from "vitest";
import {
  parseAnswerTo2D,
  computeCacheKey,
  rangesToContext,
  parseDelimitedLine
} from "../functions/functions";

// ---------------------------------------------------------------------------
// parseAnswerTo2D
// ---------------------------------------------------------------------------

describe("parseAnswerTo2D", () => {
  it("returns single word as 1×1", () => {
    expect(parseAnswerTo2D("hello")).toEqual([["hello"]]);
  });

  it("returns single line with no delimiters as 1×1", () => {
    expect(parseAnswerTo2D("The answer is 42")).toEqual([["The answer is 42"]]);
  });

  it("returns multi-line plain text as vertical spill (N×1)", () => {
    expect(parseAnswerTo2D("line1\nline2\nline3")).toEqual([
      ["line1"],
      ["line2"],
      ["line3"]
    ]);
  });

  it("parses tab-separated data into N×M matrix", () => {
    const input = "Name\tAge\nAlice\t30\nBob\t25";
    expect(parseAnswerTo2D(input)).toEqual([
      ["Name", "Age"],
      ["Alice", "30"],
      ["Bob", "25"]
    ]);
  });

  it("parses CSV data into N×M matrix", () => {
    const input = "Name,Age\nAlice,30\nBob,25";
    expect(parseAnswerTo2D(input)).toEqual([
      ["Name", "Age"],
      ["Alice", "30"],
      ["Bob", "25"]
    ]);
  });

  it("handles CSV with quoted fields containing commas", () => {
    const input = '"Last, First",Age\n"Doe, Jane",28';
    expect(parseAnswerTo2D(input)).toEqual([
      ["Last, First", "Age"],
      ["Doe, Jane", "28"]
    ]);
  });

  it("handles CSV with escaped double quotes", () => {
    const input = 'Name,Note\nAlice,"She said ""hello"""\nBob,ok';
    expect(parseAnswerTo2D(input)).toEqual([
      ["Name", "Note"],
      ["Alice", 'She said "hello"'],
      ["Bob", "ok"]
    ]);
  });

  it("prefers tabs over commas when both present", () => {
    // Tabs present → use tab as delimiter, commas are literal
    const input = "a,b\tc\nd,e\tf";
    expect(parseAnswerTo2D(input)).toEqual([
      ["a,b", "c"],
      ["d,e", "f"]
    ]);
  });

  it('returns [[""]] for empty string', () => {
    expect(parseAnswerTo2D("")).toEqual([[""]]);
  });

  it("returns two rows for single newline", () => {
    expect(parseAnswerTo2D("a\nb")).toEqual([["a"], ["b"]]);
  });

  it("pads uneven columns with empty strings", () => {
    const input = "a\tb\tc\nd\te";
    expect(parseAnswerTo2D(input)).toEqual([
      ["a", "b", "c"],
      ["d", "e", ""]
    ]);
  });

  it('returns [[""]] for whitespace-only input', () => {
    expect(parseAnswerTo2D("   \n  \t  ")).toEqual([[""]]);
  });

  it("handles single tab-delimited line", () => {
    expect(parseAnswerTo2D("a\tb\tc")).toEqual([["a", "b", "c"]]);
  });
});

// ---------------------------------------------------------------------------
// computeCacheKey
// ---------------------------------------------------------------------------

describe("computeCacheKey", () => {
  it("same query + same data → same key", () => {
    const ranges = [[["A", "B"], ["C", "D"]]] as unknown[][][];
    expect(computeCacheKey("q1", ranges)).toBe(computeCacheKey("q1", ranges));
  });

  it("different query, same data → different key", () => {
    const ranges = [[["A"]]] as unknown[][][];
    expect(computeCacheKey("q1", ranges)).not.toBe(
      computeCacheKey("q2", ranges)
    );
  });

  it("same query, different data → different key", () => {
    const r1 = [[["A"]]] as unknown[][][];
    const r2 = [[["B"]]] as unknown[][][];
    expect(computeCacheKey("q", r1)).not.toBe(computeCacheKey("q", r2));
  });

  it("empty ranges → deterministic key from query alone", () => {
    const key1 = computeCacheKey("hello", []);
    const key2 = computeCacheKey("hello", []);
    expect(key1).toBe(key2);
    expect(key1).toBe("hello");
  });

  it("multiple ranges produce distinct keys from single range with same data", () => {
    const single = [[["A", "B"]]] as unknown[][][];
    const multi = [[["A"]], [["B"]]] as unknown[][][];
    expect(computeCacheKey("q", single)).not.toBe(
      computeCacheKey("q", multi)
    );
  });

  it("null/undefined cell values coerced to empty string", () => {
    const ranges = [[[null, undefined, "x"]]] as unknown[][][];
    const key = computeCacheKey("q", ranges);
    expect(key).toContain("\t\tx");
  });
});

// ---------------------------------------------------------------------------
// rangesToContext
// ---------------------------------------------------------------------------

describe("rangesToContext", () => {
  it("serializes a single range as CSV", () => {
    const ranges = [[["Name", "Age"], ["Alice", 30]]] as unknown[][][];
    expect(rangesToContext(ranges)).toBe("Name,Age\nAlice,30");
  });

  it("separates multiple ranges with blank line", () => {
    const ranges = [[["A"]], [["B"]]] as unknown[][][];
    expect(rangesToContext(ranges)).toBe("A\n\nB");
  });

  it("returns empty string for empty ranges", () => {
    expect(rangesToContext([])).toBe("");
  });

  it("quotes cells containing commas", () => {
    const ranges = [[["hello, world"]]] as unknown[][][];
    expect(rangesToContext(ranges)).toBe('"hello, world"');
  });

  it("quotes cells containing newlines", () => {
    const ranges = [[["line1\nline2"]]] as unknown[][][];
    expect(rangesToContext(ranges)).toBe('"line1\nline2"');
  });

  it("escapes double quotes in cell values", () => {
    const ranges = [[['say "hi"']]] as unknown[][][];
    expect(rangesToContext(ranges)).toBe('"say ""hi"""');
  });

  it("converts mixed types to strings", () => {
    const ranges = [[["text", 42, true, null]]] as unknown[][][];
    expect(rangesToContext(ranges)).toBe("text,42,true,");
  });
});

// ---------------------------------------------------------------------------
// parseDelimitedLine
// ---------------------------------------------------------------------------

describe("parseDelimitedLine", () => {
  it("splits simple comma-separated values", () => {
    expect(parseDelimitedLine("a,b,c", ",")).toEqual(["a", "b", "c"]);
  });

  it("splits tab-separated values", () => {
    expect(parseDelimitedLine("a\tb\tc", "\t")).toEqual(["a", "b", "c"]);
  });

  it("handles quoted field with comma inside", () => {
    expect(parseDelimitedLine('"a,b",c', ",")).toEqual(["a,b", "c"]);
  });

  it("handles escaped double quote inside quoted field", () => {
    expect(parseDelimitedLine('"say ""hello""",done', ",")).toEqual([
      'say "hello"',
      "done"
    ]);
  });

  it("handles empty fields", () => {
    expect(parseDelimitedLine("a,,c", ",")).toEqual(["a", "", "c"]);
  });

  it("handles trailing delimiter", () => {
    expect(parseDelimitedLine("a,b,", ",")).toEqual(["a", "b", ""]);
  });
});
