/**
 * Tests for utility functions exported from App.tsx.
 *
 * `readApiError` and `formatNetworkError` are pure functions that transform
 * HTTP responses and error objects into user-friendly strings.
 */

import { describe, it, expect } from "vitest";
import { formatNetworkError } from "../App";

// readApiError is async and needs a Response mock — test it with a Response stub
// readApiError is tested below in the "readApiError (via Response mock)" suite.

describe("formatNetworkError", () => {
  it("detects fetch TypeError and includes base URL", () => {
    const error = new TypeError("Failed to fetch");
    const msg = formatNetworkError(error, "https://localhost:8000");
    expect(msg).toContain("localhost:8000");
    expect(msg).toContain("backend");
  });

  it("returns Error.message for non-fetch errors", () => {
    const error = new Error("Timeout exceeded");
    const msg = formatNetworkError(error, "https://localhost:8000");
    expect(msg).toBe("Timeout exceeded");
  });

  it("returns generic message for non-Error", () => {
    const msg = formatNetworkError("some string", "https://localhost:8000");
    expect(msg).toBe("Network error");
  });

  it("handles TypeError without fetch in message", () => {
    const error = new TypeError("Cannot read properties of null");
    const msg = formatNetworkError(error, "https://localhost:8000");
    expect(msg).toBe("Cannot read properties of null");
  });

  it("handles error with empty message", () => {
    const error = new Error("");
    const msg = formatNetworkError(error, "https://localhost:8000");
    expect(msg).toBe("");
  });
});

describe("readApiError (via Response mock)", () => {
  // Dynamic import to get the async function
  let readApiError: (res: Response) => Promise<string>;

  beforeAll(async () => {
    const mod = await import("../App");
    readApiError = mod.readApiError;
  });

  it("extracts detail from JSON response", async () => {
    const res = new Response(JSON.stringify({ detail: "Bad request" }), {
      status: 400,
      headers: { "content-type": "application/json" },
    });
    const msg = await readApiError(res);
    expect(msg).toBe("Bad request");
  });

  it("extracts message from JSON response", async () => {
    const res = new Response(JSON.stringify({ message: "Not found" }), {
      status: 404,
      headers: { "content-type": "application/json" },
    });
    const msg = await readApiError(res);
    expect(msg).toBe("Not found");
  });

  it("falls back to text for non-JSON response", async () => {
    const res = new Response("Plain error text", {
      status: 500,
      headers: { "content-type": "text/plain" },
    });
    const msg = await readApiError(res);
    expect(msg).toBe("Plain error text");
  });

  it("falls back to status code when empty", async () => {
    const res = new Response("", {
      status: 502,
      headers: { "content-type": "text/plain" },
    });
    const msg = await readApiError(res);
    expect(msg).toBe("Status 502");
  });

  it("prefers detail over message", async () => {
    const res = new Response(
      JSON.stringify({ detail: "Detail wins", message: "Message loses" }),
      { status: 400, headers: { "content-type": "application/json" } }
    );
    const msg = await readApiError(res);
    expect(msg).toBe("Detail wins");
  });
});
