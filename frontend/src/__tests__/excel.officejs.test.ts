/**
 * Mocked Office.js tests for excel.ts functions that call `Excel.run`.
 *
 * Uses the global `Excel.run` stub from our setup file and overrides it
 * per-test to control the mock context.
 */

import { describe, it, expect, vi, beforeEach } from "vitest";
import { excelRun } from "../test/officeMock";
import {
  getWorkbookMetadata,
  getUserContext,
  executeWorkbookTool,
} from "../excel";

beforeEach(() => {
  excelRun.mockReset();
  // Restore Office.context for each test
  (globalThis as any).Office.context = {
    host: "Excel",
    diagnostics: { version: "test" },
  };
});

describe("getWorkbookMetadata", () => {
  it("returns failure when Office is unavailable", async () => {
    (globalThis as any).Office.context = null;
    const result = await getWorkbookMetadata();
    expect(result.success).toBe(false);
    expect(result.sheetsMetadata).toEqual([]);
  });

  it("returns success on normal path", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          name: "Test.xlsx",
          load: vi.fn(),
          worksheets: {
            items: [
              {
                id: "s1",
                name: "Sheet1",
                position: 0,
                load: vi.fn(),
                getUsedRangeOrNullObject: () => ({
                  isNullObject: false,
                  rowCount: 10,
                  columnCount: 5,
                  load: vi.fn(),
                }),
              },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getWorkbookMetadata();
    expect(result.success).toBe(true);
    expect(result.fileName).toBe("Test.xlsx");
    expect(result.sheetsMetadata).toHaveLength(1);
    expect(result.totalSheets).toBe(1);
  });

  it("handles Excel.run error gracefully", async () => {
    excelRun.mockRejectedValue(new Error("Excel not ready"));
    const result = await getWorkbookMetadata();
    expect(result.success).toBe(false);
  });

  it("handles null usedRange", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          name: "Empty.xlsx",
          load: vi.fn(),
          worksheets: {
            items: [
              {
                id: "s1",
                name: "Sheet1",
                position: 0,
                load: vi.fn(),
                getUsedRangeOrNullObject: () => ({
                  isNullObject: true,
                  rowCount: 0,
                  columnCount: 0,
                  load: vi.fn(),
                }),
              },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getWorkbookMetadata();
    expect(result.success).toBe(true);
    expect(result.sheetsMetadata[0].maxRows).toBe(0);
  });

  it("handles multiple sheets", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      const makeSheet = (id: string, name: string, pos: number) => ({
        id,
        name,
        position: pos,
        load: vi.fn(),
        getUsedRangeOrNullObject: () => ({
          isNullObject: false,
          rowCount: 5,
          columnCount: 3,
          load: vi.fn(),
        }),
      });
      return cb({
        workbook: {
          name: "Multi.xlsx",
          load: vi.fn(),
          worksheets: {
            items: [
              makeSheet("s1", "Sheet1", 0),
              makeSheet("s2", "Sheet2", 1),
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getWorkbookMetadata();
    expect(result.sheetsMetadata).toHaveLength(2);
    expect(result.totalSheets).toBe(2);
  });
});

describe("getUserContext", () => {
  it("returns active sheet and selection", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              name: "Sales",
              load: vi.fn(),
            }),
          },
          getSelectedRange: () => ({
            address: "Sales!A1:B5",
            load: vi.fn(),
          }),
        },
        sync: vi.fn(),
      });
    });

    const ctx = await getUserContext();
    expect(ctx.currentActiveSheetName).toBe("Sales");
    expect(ctx.selectedRanges).toBe("Sales!A1:B5");
  });

  it("returns empty when Office unavailable", async () => {
    (globalThis as any).Office.context = null;
    const ctx = await getUserContext();
    expect(ctx.currentActiveSheetName).toBe("");
  });

  it("handles error gracefully", async () => {
    excelRun.mockRejectedValue(new Error("fail"));
    const ctx = await getUserContext();
    expect(ctx.currentActiveSheetName).toBe("");
    expect(ctx.selectedRanges).toBe("");
  });
});

describe("executeWorkbookTool", () => {
  it("dispatches get_xl_cell_ranges", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getRange: () => ({
                address: "Sheet1!A1:B2",
                values: [[1, 2]],
                formulas: [["1", "2"]],
                numberFormat: [["General", "General"]],
                load: vi.fn(),
                format: {
                  fill: { color: "#FFFFFF", load: vi.fn() },
                  font: { color: "#000000", bold: false, italic: false, load: vi.fn() },
                },
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await executeWorkbookTool({
      id: "tc-1",
      tool: "get_xl_cell_ranges",
      args: { ranges: ["A1:B2"] },
    });
    expect(result.id).toBe("tc-1");
    expect(result.error).toBeUndefined();
  });

  it("returns error for unknown tool", async () => {
    const result = await executeWorkbookTool({
      id: "tc-2",
      tool: "unknown_tool",
      args: {},
    });
    expect(result.error).toContain("Unknown tool");
  });

  it("catches Excel.run exceptions", async () => {
    excelRun.mockRejectedValue(new Error("Excel crash"));
    const result = await executeWorkbookTool({
      id: "tc-3",
      tool: "get_xl_cell_ranges",
      args: { ranges: ["A1"] },
    });
    expect(result.error).toBe("Excel crash");
  });

  it("dispatches get_xl_range_as_csv", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getItem: () => ({
              getRange: () => ({
                values: [
                  ["a", "b"],
                  [1, 2],
                ],
                rowCount: 2,
                load: vi.fn(),
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await executeWorkbookTool({
      id: "tc-4",
      tool: "get_xl_range_as_csv",
      args: { sheetName: "Sheet1", range: "A1:B2" },
    });
    expect(result.error).toBeUndefined();
    expect(typeof result.result).toBe("string");
  });

  it("dispatches xl_search_data", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            items: [
              {
                name: "Sheet1",
                load: vi.fn(),
                findAllOrNullObject: () => ({
                  isNullObject: true,
                  areas: { items: [], load: vi.fn() },
                  load: vi.fn(),
                }),
              },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await executeWorkbookTool({
      id: "tc-5",
      tool: "xl_search_data",
      args: { query: "test" },
    });
    expect(result.error).toBeUndefined();
  });

  it("dispatches get_all_xl_objects", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            items: [
              {
                name: "Sheet1",
                load: vi.fn(),
                charts: { items: [], load: vi.fn() },
                tables: { items: [], load: vi.fn() },
                pivotTables: { items: [], load: vi.fn() },
              },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await executeWorkbookTool({
      id: "tc-6",
      tool: "get_all_xl_objects",
      args: {},
    });
    expect(result.error).toBeUndefined();
  });

  it("dispatches execute_xl_office_js", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {},
        sync: vi.fn(),
      });
    });

    const result = await executeWorkbookTool({
      id: "tc-7",
      tool: "execute_xl_office_js",
      args: { code: "return 42;" },
    });
    expect(result.error).toBeUndefined();
  });
});
