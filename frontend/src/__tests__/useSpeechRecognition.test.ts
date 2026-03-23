/**
 * Tests for the useSpeechRecognition hook.
 *
 * Covers strategy detection (Web Speech vs Whisper vs unavailable),
 * Web Speech lifecycle, error mapping, Whisper fallback fetch, and
 * cleanup on unmount.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { renderHook, act } from "@testing-library/react";
import { useSpeechRecognition } from "../hooks/useSpeechRecognition";

/* ── Helpers ────────────────────────────────────────────────────── */

/** Minimal SpeechRecognition stub. */
function createMockRecognition() {
  return {
    continuous: false,
    interimResults: false,
    lang: "",
    onresult: null as ((e: unknown) => void) | null,
    onerror: null as ((e: unknown) => void) | null,
    onend: null as (() => void) | null,
    start: vi.fn(),
    stop: vi.fn(function (this: { onend: (() => void) | null }) {
      this.onend?.();
    }),
    abort: vi.fn(),
    addEventListener: vi.fn(),
    removeEventListener: vi.fn(),
    dispatchEvent: vi.fn(() => true),
  };
}

let savedSpeechRecognition: unknown;
let savedWebkitSpeechRecognition: unknown;
let savedGetUserMedia: unknown;

beforeEach(() => {
  savedSpeechRecognition = (window as Record<string, unknown>).SpeechRecognition;
  savedWebkitSpeechRecognition = (window as Record<string, unknown>).webkitSpeechRecognition;
  savedGetUserMedia = navigator.mediaDevices?.getUserMedia;
});

afterEach(() => {
  (window as Record<string, unknown>).SpeechRecognition = savedSpeechRecognition;
  (window as Record<string, unknown>).webkitSpeechRecognition = savedWebkitSpeechRecognition;
  if (navigator.mediaDevices) {
    (navigator.mediaDevices as Record<string, unknown>).getUserMedia = savedGetUserMedia;
  }
  vi.restoreAllMocks();
});

function clearSpeechAPIs() {
  delete (window as Record<string, unknown>).SpeechRecognition;
  delete (window as Record<string, unknown>).webkitSpeechRecognition;
  // navigator.mediaDevices may be undefined in jsdom
  if (navigator.mediaDevices) {
    (navigator.mediaDevices as Record<string, unknown>).getUserMedia = undefined;
  }
}

function ensureMediaDevices() {
  if (!navigator.mediaDevices) {
    Object.defineProperty(navigator, "mediaDevices", {
      value: {},
      writable: true,
      configurable: true,
    });
  }
}

/* ── Tests ──────────────────────────────────────────────────────── */

describe("useSpeechRecognition", () => {
  it("isAvailable true when webkitSpeechRecognition exists", () => {
    clearSpeechAPIs();
    const mockInstance = createMockRecognition();
    (window as Record<string, unknown>).webkitSpeechRecognition = vi.fn(() => mockInstance);

    const { result } = renderHook(() =>
      useSpeechRecognition({ onTranscript: vi.fn() })
    );
    expect(result.current.isAvailable).toBe(true);
    expect(result.current.strategy).toBe("webSpeech");
  });

  it("isAvailable true when getUserMedia exists but no SpeechRecognition", () => {
    clearSpeechAPIs();
    ensureMediaDevices();
    (navigator.mediaDevices as Record<string, unknown>).getUserMedia = vi.fn();

    const { result } = renderHook(() =>
      useSpeechRecognition({ onTranscript: vi.fn() })
    );
    expect(result.current.isAvailable).toBe(true);
    expect(result.current.strategy).toBe("whisper");
  });

  it("isAvailable false when neither API exists", () => {
    clearSpeechAPIs();

    const { result } = renderHook(() =>
      useSpeechRecognition({ onTranscript: vi.fn() })
    );
    expect(result.current.isAvailable).toBe(false);
    expect(result.current.strategy).toBe("none");
  });

  it("toggleListening starts and stops web speech", () => {
    clearSpeechAPIs();
    const mockInstance = createMockRecognition();
    (window as Record<string, unknown>).webkitSpeechRecognition = vi.fn(() => mockInstance);

    const { result } = renderHook(() =>
      useSpeechRecognition({ onTranscript: vi.fn() })
    );

    // Start
    act(() => result.current.toggleListening());
    expect(result.current.isListening).toBe(true);
    expect(mockInstance.start).toHaveBeenCalled();

    // Stop (mock triggers onend)
    act(() => result.current.toggleListening());
    expect(result.current.isListening).toBe(false);
    expect(mockInstance.stop).toHaveBeenCalled();
  });

  it("onTranscript called with interim and final results", () => {
    clearSpeechAPIs();
    const mockInstance = createMockRecognition();
    (window as Record<string, unknown>).webkitSpeechRecognition = vi.fn(() => mockInstance);

    const onTranscript = vi.fn();
    renderHook(() => useSpeechRecognition({ onTranscript }));

    // Simulate: start recognition so onresult is wired up.
    // The constructor was called but we need to trigger the hook's toggleListening.
    const { result } = renderHook(() => useSpeechRecognition({ onTranscript }));
    act(() => result.current.toggleListening());

    // Simulate interim result
    const interimEvent = {
      resultIndex: 0,
      results: {
        length: 1,
        0: { isFinal: false, length: 1, 0: { transcript: "hel", confidence: 0.8 } },
      },
    };
    act(() => mockInstance.onresult?.(interimEvent));
    expect(onTranscript).toHaveBeenCalledWith("hel", false);

    // Simulate final result
    const finalEvent = {
      resultIndex: 0,
      results: {
        length: 1,
        0: { isFinal: true, length: 1, 0: { transcript: "hello world", confidence: 0.95 } },
      },
    };
    act(() => mockInstance.onresult?.(finalEvent));
    expect(onTranscript).toHaveBeenCalledWith("hello world", true);
  });

  it('error not-allowed maps to "Microphone access denied"', () => {
    clearSpeechAPIs();
    const mockInstance = createMockRecognition();
    (window as Record<string, unknown>).webkitSpeechRecognition = vi.fn(() => mockInstance);

    const onError = vi.fn();
    const { result } = renderHook(() =>
      useSpeechRecognition({ onTranscript: vi.fn(), onError })
    );

    act(() => result.current.toggleListening());
    act(() => mockInstance.onerror?.({ error: "not-allowed", message: "" }));

    expect(onError).toHaveBeenCalledWith("Microphone access denied");
  });

  it("error no-speech silently stops", () => {
    clearSpeechAPIs();
    const mockInstance = createMockRecognition();
    (window as Record<string, unknown>).webkitSpeechRecognition = vi.fn(() => mockInstance);

    const onError = vi.fn();
    const { result } = renderHook(() =>
      useSpeechRecognition({ onTranscript: vi.fn(), onError })
    );

    act(() => result.current.toggleListening());
    act(() => mockInstance.onerror?.({ error: "no-speech", message: "" }));

    expect(onError).not.toHaveBeenCalled();
    expect(result.current.isListening).toBe(false);
  });

  it("whisper fallback POSTs to /transcribe", async () => {
    clearSpeechAPIs();
    ensureMediaDevices();

    // Mock getUserMedia
    const mockStream = {
      getTracks: () => [{ stop: vi.fn() }],
    };
    (navigator.mediaDevices as Record<string, unknown>).getUserMedia = vi.fn(() =>
      Promise.resolve(mockStream)
    );

    // Mock MediaRecorder
    let onstopHandler: (() => void) | null = null;
    let ondataHandler: ((e: { data: Blob }) => void) | null = null;
    const mockRecorder = {
      start: vi.fn(),
      stop: vi.fn(() => onstopHandler?.()),
      get ondataavailable() { return ondataHandler; },
      set ondataavailable(fn) { ondataHandler = fn as ((e: { data: Blob }) => void) | null; },
      get onstop() { return onstopHandler; },
      set onstop(fn) { onstopHandler = fn as (() => void) | null; },
    };
    globalThis.MediaRecorder = vi.fn(() => mockRecorder) as unknown as typeof MediaRecorder;

    // Mock fetch
    const fetchMock = vi.fn(() =>
      Promise.resolve({
        ok: true,
        json: () => Promise.resolve({ text: "transcribed text" }),
      })
    );
    globalThis.fetch = fetchMock as unknown as typeof fetch;

    const onTranscript = vi.fn();
    const { result } = renderHook(() =>
      useSpeechRecognition({ onTranscript })
    );

    // Start recording (async)
    await act(async () => {
      result.current.toggleListening();
      // Allow getUserMedia promise to resolve
      await Promise.resolve();
    });

    // Simulate data chunk
    ondataHandler?.({ data: new Blob(["audio"], { type: "audio/webm" }) });

    // Stop recording
    await act(async () => {
      result.current.toggleListening();
      // Allow onstop handler and fetch to resolve
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();
    });

    expect(fetchMock).toHaveBeenCalledWith(
      expect.stringContaining("/transcribe"),
      expect.objectContaining({ method: "POST" })
    );
  });

  it("cleanup aborts on unmount", () => {
    clearSpeechAPIs();
    const mockInstance = createMockRecognition();
    (window as Record<string, unknown>).webkitSpeechRecognition = vi.fn(() => mockInstance);

    const { result, unmount } = renderHook(() =>
      useSpeechRecognition({ onTranscript: vi.fn() })
    );

    act(() => result.current.toggleListening());
    unmount();

    expect(mockInstance.abort).toHaveBeenCalled();
  });
});
