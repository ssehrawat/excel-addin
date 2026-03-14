/* global Office, CustomFunctions */

import { createRoot } from "react-dom/client";
import { App } from "./App";
import { askAI } from "./functions/functions";
import "./taskpane.css";

Office.onReady(() => {
  // Register custom function in the shared runtime
  if (typeof CustomFunctions !== "undefined") {
    CustomFunctions.associate("ASKAI", askAI);
  }

  const container = document.getElementById("root");
  if (!container) {
    throw new Error("Root container not found");
  }
  const root = createRoot(container);
  root.render(<App />);
});
