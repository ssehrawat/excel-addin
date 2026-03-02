/**
 * Vitest setup file.
 * Loads jest-dom matchers and initializes the Office.js mock globals.
 */

import "@testing-library/jest-dom/vitest";
import "./officeMock";

// jsdom does not implement scrollIntoView
Element.prototype.scrollIntoView = () => {};
