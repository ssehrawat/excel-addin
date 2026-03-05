/**
 * Minimal Office.js and Excel global stubs for unit testing.
 *
 * Provides enough surface area for the pure-function tests in excel.ts
 * (chart type aliases, range reference parsing) and mocked Office.js tests.
 * Functions that need deeper mocking should set up their own `Excel.run`
 * implementation per test.
 */

import { vi } from "vitest";

// --- Excel.ChartType enum stub ---
const ChartType: Record<string, string> = {
  xyscatter: "XYScatter",
  xyscatterLines: "XYScatterLines",
  xyscatterLinesNoMarkers: "XYScatterLinesNoMarkers",
  bubble: "Bubble",
  lineMarkers: "LineMarkers",
  columnClustered: "ColumnClustered",
  barClustered: "BarClustered",
  line: "Line",
  pie: "Pie",
  area: "Area",
};

// --- Excel.ChartSeriesBy enum stub ---
const ChartSeriesBy: Record<string, string> = {
  auto: "Auto",
  rows: "Rows",
  columns: "Columns",
};

// --- Excel.BorderIndex enum stub ---
const BorderIndex: Record<string, string> = {
  edgeTop: "EdgeTop",
  edgeBottom: "EdgeBottom",
  edgeLeft: "EdgeLeft",
  edgeRight: "EdgeRight",
};

// --- Excel.AggregationFunction enum stub ---
const AggregationFunction: Record<string, string> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min",
  product: "Product",
  countNumbers: "CountNumbers",
  standardDeviation: "StandardDeviation",
  standardDeviationP: "StandardDeviationP",
  variance: "Variance",
  varianceP: "VarianceP",
};

// --- Excel.run stub ---
const excelRun = vi.fn(async (callback: (context: any) => Promise<any>) => {
  // Default: no-op context. Override in individual tests.
  return callback({
    workbook: {
      name: "TestWorkbook.xlsx",
      load: vi.fn(),
      worksheets: {
        items: [],
        load: vi.fn(),
        getActiveWorksheet: vi.fn(),
        getItem: vi.fn(),
      },
      getSelectedRange: vi.fn(),
    },
    sync: vi.fn(),
  });
});

// --- Global Excel object ---
(globalThis as any).Excel = {
  ChartType,
  ChartSeriesBy,
  BorderIndex,
  AggregationFunction,
  run: excelRun,
};

// --- Global Office object ---
(globalThis as any).Office = {
  context: {
    host: "Excel" as any,
    diagnostics: { version: "test" },
  },
  HostType: {
    Excel: "Excel",
    Word: "Word",
  },
  onReady: vi.fn((cb?: () => void) => {
    cb?.();
    return Promise.resolve();
  }),
};

export { excelRun, ChartType, ChartSeriesBy, BorderIndex, AggregationFunction };
