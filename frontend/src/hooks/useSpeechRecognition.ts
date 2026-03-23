/**
 * Speech-to-text hook with dual strategy: Web Speech API (primary)
 * and MediaRecorder + OpenAI Whisper fallback.
 *
 * @module hooks/useSpeechRecognition
 */

import { useCallback, useEffect, useRef, useState } from "react";
import { API_BASE_URL } from "../config";

/** Options accepted by {@link useSpeechRecognition}. */
export interface UseSpeechRecognitionOptions {
  /** Called with transcribed text. `isFinal` is true once the utterance is complete. */
  onTranscript: (text: string, isFinal: boolean) => void;
  /** Called on recoverable errors (e.g. mic denied). */
  onError?: (error: string) => void;
  /** BCP-47 language tag. @default "en-US" */
  lang?: string;
}

/** Return value of {@link useSpeechRecognition}. */
export interface UseSpeechRecognitionReturn {
  /** False only when no microphone access is possible at all. */
  isAvailable: boolean;
  /** True while actively recording / recognising. */
  isListening: boolean;
  /** The detected speech recognition strategy. */
  strategy: Strategy;
  /** Start if stopped, stop if started. */
  toggleListening: () => void;
  /** Force stop (idempotent). */
  stopListening: () => void;
}

export type Strategy = "webSpeech" | "whisper" | "none";

function detectStrategy(): Strategy {
  if (
    typeof window !== "undefined" &&
    (window.SpeechRecognition || window.webkitSpeechRecognition)
  ) {
    return "webSpeech";
  }
  if (typeof navigator !== "undefined" && typeof navigator.mediaDevices?.getUserMedia === "function") {
    return "whisper";
  }
  return "none";
}

/**
 * React hook providing speech-to-text input via the browser-native
 * Web Speech API with an automatic fallback to MediaRecorder + OpenAI
 * Whisper when the Web Speech API is unavailable.
 *
 * @param options - Configuration callbacks and language.
 * @returns State and controls for speech recognition.
 */
export function useSpeechRecognition(
  options: UseSpeechRecognitionOptions
): UseSpeechRecognitionReturn {
  const { onTranscript, onError, lang = "en-US" } = options;

  const [strategy] = useState<Strategy>(detectStrategy);
  const [isListening, setIsListening] = useState(false);

  // Refs to keep mutable handles across renders.
  const recognitionRef = useRef<SpeechRecognition | null>(null);
  const mediaStreamRef = useRef<MediaStream | null>(null);
  const recorderRef = useRef<MediaRecorder | null>(null);

  // Stable callback refs so effect cleanup doesn't stale-close.
  const onTranscriptRef = useRef(onTranscript);
  onTranscriptRef.current = onTranscript;
  const onErrorRef = useRef(onError);
  onErrorRef.current = onError;

  // ── Web Speech path ────────────────────────────────────────────
  const startWebSpeech = useCallback(() => {
    const SpeechCtor =
      window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechCtor) return;

    const recognition = new SpeechCtor();
    recognition.continuous = false;
    recognition.interimResults = true;
    recognition.lang = lang;

    recognition.onresult = (event: SpeechRecognitionEvent) => {
      let transcript = "";
      let isFinal = false;
      for (let i = event.resultIndex; i < event.results.length; i++) {
        transcript += event.results[i][0].transcript;
        if (event.results[i].isFinal) isFinal = true;
      }
      onTranscriptRef.current(transcript, isFinal);
    };

    recognition.onerror = (event: SpeechRecognitionErrorEvent) => {
      if (event.error === "no-speech") {
        // Silent stop — not an actionable error.
        setIsListening(false);
        return;
      }
      const msg =
        event.error === "not-allowed"
          ? "Microphone access denied"
          : event.error === "network"
            ? "Voice input unavailable offline"
            : `Speech recognition error: ${event.error}`;
      onErrorRef.current?.(msg);
      setIsListening(false);
    };

    recognition.onend = () => {
      setIsListening(false);
    };

    recognitionRef.current = recognition;
    recognition.start();
    setIsListening(true);
  }, [lang]);

  const stopWebSpeech = useCallback(() => {
    recognitionRef.current?.stop();
    recognitionRef.current = null;
  }, []);

  // ── Whisper fallback path ──────────────────────────────────────
  const startWhisper = useCallback(async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      mediaStreamRef.current = stream;

      const recorder = new MediaRecorder(stream, { mimeType: "audio/webm" });
      const chunks: Blob[] = [];

      recorder.ondataavailable = (e) => {
        if (e.data.size > 0) chunks.push(e.data);
      };

      recorder.onstop = async () => {
        // Release mic immediately.
        stream.getTracks().forEach((t) => t.stop());
        mediaStreamRef.current = null;

        const blob = new Blob(chunks, { type: "audio/webm" });
        const form = new FormData();
        form.append("audio", blob, "recording.webm");

        try {
          const res = await fetch(`${API_BASE_URL}/transcribe`, {
            method: "POST",
            body: form,
          });
          if (!res.ok) {
            const detail = await res.text();
            onErrorRef.current?.(`Transcription failed: ${detail}`);
            return;
          }
          const data = (await res.json()) as { text: string };
          onTranscriptRef.current(data.text, true);
        } catch (err) {
          onErrorRef.current?.(
            `Transcription request failed: ${err instanceof Error ? err.message : String(err)}`
          );
        } finally {
          setIsListening(false);
        }
      };

      recorderRef.current = recorder;
      recorder.start();
      setIsListening(true);
    } catch {
      onErrorRef.current?.("Microphone access denied");
      setIsListening(false);
    }
  }, []);

  const stopWhisper = useCallback(() => {
    recorderRef.current?.stop();
    recorderRef.current = null;
  }, []);

  // ── Public API ─────────────────────────────────────────────────
  const toggleListening = useCallback(() => {
    if (isListening) {
      if (strategy === "webSpeech") stopWebSpeech();
      else stopWhisper();
    } else {
      if (strategy === "webSpeech") startWebSpeech();
      else if (strategy === "whisper") void startWhisper();
    }
  }, [isListening, strategy, startWebSpeech, stopWebSpeech, startWhisper, stopWhisper]);

  const stopListening = useCallback(() => {
    if (!isListening) return;
    if (strategy === "webSpeech") stopWebSpeech();
    else stopWhisper();
  }, [isListening, strategy, stopWebSpeech, stopWhisper]);

  // Cleanup on unmount.
  useEffect(() => {
    return () => {
      recognitionRef.current?.abort();
      recorderRef.current?.stop();
      mediaStreamRef.current?.getTracks().forEach((t) => t.stop());
    };
  }, []);

  return {
    isAvailable: strategy !== "none",
    isListening,
    strategy,
    toggleListening,
    stopListening,
  };
}
