/**
 * E2E Simulator Test Runner — Layer 1
 *
 * Data-driven tests that load scenario fixtures, set up a stateful
 * SimWorkbook, execute real excel.ts mutation functions against it, and
 * assert on the resulting workbook state. This catches the same class of
 * bugs that manual Excel testing catches (mutation correctness, chart
 * auto-positioning, cross-sheet writes, etc.) without requiring Excel.
 *
 * Key exports: none (test file)
 *
 * @module e2e.simulator.test
 */

import { describe, it, expect, afterEach } from "vitest";
import fs from "fs";
import path from "path";

import {
  SimWorkbook,
  SimChart,
  SimPivotTable,
  buildSimWorkbook,
  installSimulator,
  type WorkbookFixtureDef,
} from "../test/workbookSimulator";

import { applyCellUpdates } from "../excel";
import { applyFormatUpdates } from "../excel";
import { insertCharts } from "../excel";
import { insertPivotTables } from "../excel";

import type {
  CellUpdate,
  FormatUpdate,
  ChartInsert,
  PivotTableInsert,
} from "../types";

// ---------------------------------------------------------------------------
// Scenario types
// ---------------------------------------------------------------------------

interface ScenarioMutation {
  type: "cell_updates" | "format_updates" | "chart_inserts" | "pivot_table_inserts";
  payload: unknown[];
}

interface ScenarioAssertion {
  type: string;
  sheet?: string;
  address?: string;
  expected?: unknown;
  prop?: string;
  chartIndex?: number;
  pivotIndex?: number;
  borderIndex?: string;
  minRow?: number;
  expectedRows?: number;
  expectedCols?: number;
  axis?: "category" | "value";
  description?: string;
}

interface Scenario {
  name: string;
  description?: string;
  initialWorkbook: WorkbookFixtureDef;
  mutations: ScenarioMutation[];
  assertions: ScenarioAssertion[];
}

// ---------------------------------------------------------------------------
// Fixture loading
// ---------------------------------------------------------------------------

const FIXTURE_DIR = path.resolve(
  __dirname,
  "../test/fixtures/workbook-scenarios"
);

function loadScenarios(): Scenario[] {
  const files = fs.readdirSync(FIXTURE_DIR).filter((f) => f.endsWith(".json"));
  return files.map((f) => {
    const raw = fs.readFileSync(path.join(FIXTURE_DIR, f), "utf-8");
    return JSON.parse(raw) as Scenario;
  });
}

// ---------------------------------------------------------------------------
// Mutation executor
// ---------------------------------------------------------------------------

async function executeMutations(mutations: ScenarioMutation[]): Promise<void> {
  for (const mutation of mutations) {
    switch (mutation.type) {
      case "cell_updates":
        await applyCellUpdates(mutation.payload as CellUpdate[]);
        break;
      case "format_updates":
        await applyFormatUpdates(mutation.payload as FormatUpdate[]);
        break;
      case "chart_inserts":
        await insertCharts(mutation.payload as ChartInsert[]);
        break;
      case "pivot_table_inserts":
        await insertPivotTables(mutation.payload as PivotTableInsert[]);
        break;
      default:
        throw new Error(`Unknown mutation type: ${mutation.type}`);
    }
  }
}

// ---------------------------------------------------------------------------
// Assertion evaluator
// ---------------------------------------------------------------------------

function getNestedProp(obj: any, propPath: string): unknown {
  return propPath.split(".").reduce((o, key) => o?.[key], obj);
}

function evaluateAssertion(
  assertion: ScenarioAssertion,
  workbook: SimWorkbook
): void {
  const sheet = assertion.sheet ? workbook.getSheet(assertion.sheet) : undefined;

  switch (assertion.type) {
    case "cellValue": {
      const val = sheet!.getCellValue(assertion.address!);
      expect(val).toEqual(assertion.expected);
      break;
    }

    case "usedRange": {
      expect(sheet!.usedRangeRows).toBe(assertion.expectedRows);
      expect(sheet!.usedRangeCols).toBe(assertion.expectedCols);
      break;
    }

    case "format": {
      const fmt = sheet!.getCellFormat(assertion.address!);
      expect((fmt as any)[assertion.prop!]).toEqual(assertion.expected);
      break;
    }

    case "border": {
      const border = sheet!.getCellBorder(
        assertion.address!,
        assertion.borderIndex!
      );
      expect((border as any)[assertion.prop!]).toEqual(assertion.expected);
      break;
    }

    case "chartCount": {
      expect(sheet!._chartsArray.length).toBe(assertion.expected);
      break;
    }

    case "chartProperty": {
      const chart = sheet!._chartsArray[assertion.chartIndex!];
      expect(chart).toBeDefined();
      expect(getNestedProp(chart, assertion.prop!)).toEqual(assertion.expected);
      break;
    }

    case "chartPositionBelowRow": {
      const chart = sheet!._chartsArray[assertion.chartIndex!];
      expect(chart).toBeDefined();
      // WHY: positionRow is set by SimChart.setPosition() from the topLeft
      // range's _startRow. A chart auto-positioned below usedRange should
      // have positionRow >= minRow.
      expect(chart.positionRow).toBeGreaterThanOrEqual(assertion.minRow!);
      break;
    }

    case "chartAxisTitle": {
      const chart = sheet!._chartsArray[assertion.chartIndex!];
      expect(chart).toBeDefined();
      const axis = (assertion as any).axis as "category" | "value";
      if (axis === "category") {
        expect(chart.axes.categoryAxis.title.text).toEqual(assertion.expected);
      } else {
        expect(chart.axes.valueAxis.title.text).toEqual(assertion.expected);
      }
      break;
    }

    case "chartAxisCorrectness": {
      // WHY: For BarClustered charts, the categoryAxis is the VERTICAL axis
      // and the valueAxis is the HORIZONTAL axis — opposite of ColumnClustered.
      // This assertion catches the bug where axis titles are correct but the
      // data is visually swapped because the code maps xAxisTitle→categoryAxis
      // regardless of chart type.
      const chart = sheet!._chartsArray[assertion.chartIndex!];
      expect(chart).toBeDefined();

      const BAR_TYPES = ["BarClustered", "BarStacked", "BarStacked100"];
      const isBarChart = BAR_TYPES.includes(chart.chartType);

      if (isBarChart) {
        // In bar charts: categoryAxis = vertical (Y), valueAxis = horizontal (X)
        // So categoryAxis title should describe the category labels (e.g. "Ticker")
        // and valueAxis title should describe the numeric values (e.g. "Market Value")
        const catTitle = chart.axes.categoryAxis.title.text;
        const valTitle = chart.axes.valueAxis.title.text;

        // The category title should NOT be a numeric-sounding label
        // and the value title should NOT be a label-sounding name
        // This is a heuristic — if the LLM sent xAxisTitle="Ticker" for a bar
        // chart, that title ends up on categoryAxis (vertical), which is correct
        // data-wise but the USER sees "Ticker" on the Y-axis, not the X-axis
        // they expected. This flags the mismatch.
        expect(catTitle).toBeTruthy();
        expect(valTitle).toBeTruthy();

        // Verify the category axis has the label-like title (not numeric)
        // and value axis has the numeric-like title
        const numericKeywords = ["value", "amount", "price", "total", "sum", "revenue", "sales", "volume", "cost", "market"];
        const catTitleLower = catTitle.toLowerCase();
        const valTitleLower = valTitle.toLowerCase();
        const catSoundsNumeric = numericKeywords.some(kw => catTitleLower.includes(kw));
        const valSoundsNumeric = numericKeywords.some(kw => valTitleLower.includes(kw));

        // FAIL if the category axis (vertical in bar chart) has a numeric title
        // because that means titles are swapped relative to the data
        expect(catSoundsNumeric).toBe(false);
        expect(valSoundsNumeric).toBe(true);
      }
      break;
    }

    case "pivotCount": {
      expect(sheet!._pivotsArray.length).toBe(assertion.expected);
      break;
    }

    case "pivotProperty": {
      const pivot = sheet!._pivotsArray[assertion.pivotIndex!];
      expect(pivot).toBeDefined();
      expect(getNestedProp(pivot, assertion.prop!)).toEqual(assertion.expected);
      break;
    }

    case "sheetExists": {
      const exists = workbook._sheets.has(assertion.sheet!);
      expect(exists).toBe(assertion.expected);
      break;
    }

    default:
      throw new Error(`Unknown assertion type: ${assertion.type}`);
  }
}

// ---------------------------------------------------------------------------
// Test suite
// ---------------------------------------------------------------------------

describe("E2E Simulator — workbook scenario tests", () => {
  let restoreFn: (() => void) | null = null;

  afterEach(() => {
    restoreFn?.();
    restoreFn = null;
  });

  const scenarios = loadScenarios();

  for (const scenario of scenarios) {
    it(`${scenario.name}: ${scenario.description ?? ""}`, async () => {
      // 1. Build workbook from fixture
      const workbook = buildSimWorkbook(scenario.initialWorkbook);

      // 2. Install simulator as global Excel.run
      const sim = installSimulator(workbook);
      restoreFn = sim.restore;

      // 3. Execute mutations (calls real excel.ts functions)
      await executeMutations(scenario.mutations);

      // 4. Evaluate assertions against simulator state
      for (const assertion of scenario.assertions) {
        evaluateAssertion(assertion, workbook);
      }
    });
  }
});
