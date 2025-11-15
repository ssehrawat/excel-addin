/* global Office */

import { createRoot } from "react-dom/client";
import { App } from "./App";
import "./taskpane.css";

const bootstrap = () => {
  const container = document.getElementById("root");
  if (!container) {
    throw new Error("Root container not found");
  }
  const root = createRoot(container);
  root.render(<App />);
};

if ((window as any).Office) {
  Office.onReady(() => {
    bootstrap();
  });
} else {
  // Running outside of Office (for local development in a browser)
  bootstrap();
}

