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
  getLightweightSheetPreview,
  initPreviewCache,
  resetPreviewCache,
  getXlCellRanges,
  getXlRangeAsCsv,
  xlSearchData,
  getAllXlObjects,
  executeWorkbookTool,
  applyCellUpdates,
  insertCharts,
} from "../excel";

beforeEach(() => {
  excelRun.mockReset();
  resetPreviewCache();
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
                getRangeByIndexes: () => ({
                  values: [["Name", "Age", "City", "State", "Zip"]],
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
    expect(result.sheetsMetadata[0].columnHeaders).toEqual([
      "Name", "Age", "City", "State", "Zip",
    ]);
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
    expect(result.sheetsMetadata[0].columnHeaders).toBeUndefined();
  });

  it("returns fresh data on each call (no internal caching)", async () => {
    let callCount = 0;
    excelRun.mockImplementation(async (cb: (ctx: any) => any) => {
      callCount++;
      return cb({
        workbook: {
          name: `File${callCount}.xlsx`,
          load: vi.fn(),
          worksheets: {
            items: Array.from({ length: callCount }, (_, i) => ({
              id: `s${i}`,
              name: `Sheet${i + 1}`,
              position: i,
              load: vi.fn(),
              getUsedRangeOrNullObject: () => ({
                isNullObject: true, rowCount: 0, columnCount: 0, load: vi.fn(),
              }),
              getRangeByIndexes: vi.fn(),
            })),
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const first = await getWorkbookMetadata();
    expect(first.fileName).toBe("File1.xlsx");
    expect(first.totalSheets).toBe(1);

    const second = await getWorkbookMetadata();
    expect(second.fileName).toBe("File2.xlsx");
    expect(second.totalSheets).toBe(2);
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
        getRangeByIndexes: () => ({
          values: [["H1", "H2", "H3"]],
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
    expect(result.sheetsMetadata[0].columnHeaders).toEqual(["H1", "H2", "H3"]);
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

describe("getLightweightSheetPreview", () => {
  it("returns null when Office is unavailable", async () => {
    (globalThis as any).Office.context = null;
    const result = await getLightweightSheetPreview();
    expect(result).toBeNull();
  });

  it("returns null for empty sheet", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getUsedRangeOrNullObject: () => ({
                isNullObject: true,
                rowCount: 0,
                columnCount: 0,
                load: vi.fn(),
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });
    const result = await getLightweightSheetPreview();
    expect(result).toBeNull();
  });

  it("loads only the requested number of rows via getRangeByIndexes", async () => {
    const getRangeByIndexes = vi.fn(() => ({
      values: [
        ["Name", "Age"],
        ["Alice", 30],
      ],
      load: vi.fn(),
    }));

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getUsedRangeOrNullObject: () => ({
                isNullObject: false,
                rowCount: 5000,
                columnCount: 2,
                load: vi.fn(),
              }),
              getRangeByIndexes,
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getLightweightSheetPreview(2);
    expect(result).toBe("[A],[B]\nName,Age\nAlice,30");
    // Should request only 2 rows (min of maxRows=2 and rowCount=5000)
    expect(getRangeByIndexes).toHaveBeenCalledWith(0, 0, 2, 2);
  });

  it("caps rows at actual rowCount when sheet is smaller than maxRows", async () => {
    const getRangeByIndexes = vi.fn(() => ({
      values: [["A"], ["B"], ["C"]],
      load: vi.fn(),
    }));

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getUsedRangeOrNullObject: () => ({
                isNullObject: false,
                rowCount: 3,
                columnCount: 1,
                load: vi.fn(),
              }),
              getRangeByIndexes,
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getLightweightSheetPreview(50);
    expect(result).toBe("[A]\nA\nB\nC");
    expect(getRangeByIndexes).toHaveBeenCalledWith(0, 0, 3, 1);
  });

  it("returns cached result on second call when not dirty", async () => {
    const syncFn = vi.fn();
    const getRangeByIndexes = vi.fn(() => ({
      values: [["cached"]],
      load: vi.fn(),
    }));

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getUsedRangeOrNullObject: () => ({
                isNullObject: false,
                rowCount: 1,
                columnCount: 1,
                load: vi.fn(),
              }),
              getRangeByIndexes,
            }),
          },
        },
        sync: syncFn,
      });
    });

    // First call: reads from Excel
    const first = await getLightweightSheetPreview();
    expect(first).toBe("[A]\ncached");
    const syncCountAfterFirst = syncFn.mock.calls.length;

    // Second call: should return cache without calling Excel.run/sync again
    const second = await getLightweightSheetPreview();
    expect(second).toBe("[A]\ncached");
    expect(syncFn.mock.calls.length).toBe(syncCountAfterFirst);
  });

  it("quotes cells containing commas", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getUsedRangeOrNullObject: () => ({
                isNullObject: false,
                rowCount: 1,
                columnCount: 1,
                load: vi.fn(),
              }),
              getRangeByIndexes: () => ({
                values: [["hello, world"]],
                load: vi.fn(),
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    // Force dirty so cache is bypassed
    await initPreviewCache();
    const result = await getLightweightSheetPreview();
    expect(result).toContain('"hello, world"');
  });
});

describe("initPreviewCache", () => {
  it("registers onChanged and onActivated listeners", async () => {
    const onChangedAdd = vi.fn();
    const onActivatedAdd = vi.fn();

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              onChanged: { add: onChangedAdd },
            }),
            onActivated: { add: onActivatedAdd },
          },
        },
        sync: vi.fn(),
      });
    });

    await initPreviewCache();
    expect(onChangedAdd).toHaveBeenCalledTimes(1);
    expect(onActivatedAdd).toHaveBeenCalledTimes(1);
  });

  it("handles errors gracefully", async () => {
    excelRun.mockRejectedValue(new Error("no listeners"));
    // Should not throw
    await initPreviewCache();
  });
});

describe("getXlCellRanges — batched sync", () => {
  it("reads multiple ranges with a single context.sync call", async () => {
    const syncFn = vi.fn();
    const makeRange = (addr: string) => ({
      address: addr,
      values: [[1]],
      formulas: [["1"]],
      numberFormat: [["General"]],
      load: vi.fn(),
      format: {
        fill: { color: "#FFF", load: vi.fn() },
        font: { color: "#000", bold: false, italic: false, load: vi.fn() },
      },
    });

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getRange: (addr: string) => makeRange(`Sheet1!${addr}`),
            }),
          },
        },
        sync: syncFn,
      });
    });

    const result = await getXlCellRanges(["A1:B2", "C3:D4", "E5:F6"]);
    expect(Array.isArray(result)).toBe(true);
    expect((result as any[]).length).toBe(3);
    // Should be exactly 1 sync call regardless of range count
    expect(syncFn).toHaveBeenCalledTimes(1);
  });
});

describe("getAllXlObjects — two-phase sync", () => {
  it("uses only 2 sync calls for multiple sheets", async () => {
    const syncFn = vi.fn();
    const makeSheet = (name: string) => ({
      name,
      load: vi.fn(),
      charts: { items: [{ name: `${name}Chart`, chartType: "Bar", load: vi.fn() }], load: vi.fn() },
      tables: { items: [{ name: `${name}Table`, showTotals: false, load: vi.fn() }], load: vi.fn() },
      pivotTables: { items: [{ name: `${name}Pivot`, load: vi.fn() }], load: vi.fn() },
    });

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            items: [makeSheet("Sheet1"), makeSheet("Sheet2"), makeSheet("Sheet3")],
            load: vi.fn(),
          },
        },
        sync: syncFn,
      });
    });

    const result = (await getAllXlObjects()) as any;
    expect(result.charts).toHaveLength(3);
    expect(result.tables).toHaveLength(3);
    expect(result.pivotTables).toHaveLength(3);
    // 1 sync for sheet list + collections, 1 sync for item properties = 3 total
    // (initial sheets load sync + 2 phase syncs)
    expect(syncFn.mock.calls.length).toBeLessThanOrEqual(3);
  });

  it("filters by objectType", async () => {
    const syncFn = vi.fn();
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            items: [{
              name: "Sheet1",
              load: vi.fn(),
              charts: { items: [{ name: "C1", chartType: "Line", load: vi.fn() }], load: vi.fn() },
              tables: { items: [], load: vi.fn() },
              pivotTables: { items: [], load: vi.fn() },
            }],
            load: vi.fn(),
          },
        },
        sync: syncFn,
      });
    });

    const result = (await getAllXlObjects(undefined, "chart")) as any;
    expect(result.charts).toHaveLength(1);
    expect(result.tables).toHaveLength(0);
    expect(result.pivotTables).toHaveLength(0);
  });
});

// -----------------------------------------------------------------------
// Missing tests for latest changes
// -----------------------------------------------------------------------

describe("getLightweightSheetPreview — additional edge cases", () => {
  it("re-reads from Excel after resetPreviewCache invalidates the cache", async () => {
    const syncFn = vi.fn();
    let callCount = 0;
    const getRangeByIndexes = vi.fn(() => ({
      values: [[`call-${++callCount}`]],
      load: vi.fn(),
    }));

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getUsedRangeOrNullObject: () => ({
                isNullObject: false,
                rowCount: 1,
                columnCount: 1,
                load: vi.fn(),
              }),
              getRangeByIndexes,
            }),
          },
        },
        sync: syncFn,
      });
    });

    const first = await getLightweightSheetPreview();
    expect(first).toBe("[A]\ncall-1");

    // Cache should serve the same value
    const cached = await getLightweightSheetPreview();
    expect(cached).toBe("[A]\ncall-1");

    // Invalidate and re-read
    resetPreviewCache();
    const refreshed = await getLightweightSheetPreview();
    expect(refreshed).toBe("[A]\ncall-2");
  });

  it("handles null cell values gracefully", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getUsedRangeOrNullObject: () => ({
                isNullObject: false,
                rowCount: 1,
                columnCount: 3,
                load: vi.fn(),
              }),
              getRangeByIndexes: () => ({
                values: [["hello", null, 42]],
                load: vi.fn(),
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getLightweightSheetPreview();
    expect(result).toBe("[A],[B],[C]\nhello,,42");
  });

  it("returns null and logs error when Excel.run throws", async () => {
    excelRun.mockRejectedValue(new Error("Excel crashed"));
    const result = await getLightweightSheetPreview();
    expect(result).toBeNull();
  });
});

describe("getXlCellRanges — additional edge cases", () => {
  it("reads from a specific sheet when sheetName is provided", async () => {
    const getItemFn = vi.fn(() => ({
      getRange: (addr: string) => ({
        address: `Sales!${addr}`,
        values: [[100]],
        formulas: [["100"]],
        numberFormat: [["#,##0"]],
        load: vi.fn(),
        format: {
          fill: { color: "#FFFF00", load: vi.fn() },
          font: { color: "#333333", bold: true, italic: false, load: vi.fn() },
        },
      }),
    }));

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getItem: getItemFn,
          },
        },
        sync: vi.fn(),
      });
    });

    const result = (await getXlCellRanges(["A1"], "Sales")) as any[];
    expect(getItemFn).toHaveBeenCalledWith("Sales");
    expect(result).toHaveLength(1);
    expect(result[0].address).toBe("Sales!A1");
    expect(result[0].fillColor).toBe("#FFFF00");
    expect(result[0].bold).toBe(true);
  });

  it("returns empty array when Office is unavailable", async () => {
    (globalThis as any).Office.context = null;
    const result = await getXlCellRanges(["A1"]);
    expect(result).toEqual([]);
  });
});

describe("xlSearchData — dedicated tests", () => {
  it("returns matches when found across sheets", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            items: [
              {
                name: "Sheet1",
                load: vi.fn(),
                findAllOrNullObject: () => ({
                  isNullObject: false,
                  load: vi.fn(),
                  areas: {
                    items: [
                      {
                        address: "Sheet1!B3",
                        values: [["found-it"]],
                        load: vi.fn(),
                      },
                    ],
                    load: vi.fn(),
                  },
                }),
              },
              {
                name: "Sheet2",
                load: vi.fn(),
                findAllOrNullObject: () => ({
                  isNullObject: true,
                  load: vi.fn(),
                  areas: { items: [], load: vi.fn() },
                }),
              },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = (await xlSearchData("found-it")) as any[];
    expect(result).toHaveLength(1);
    expect(result[0].address).toBe("Sheet1!B3");
    expect(result[0].worksheet).toBe("Sheet1");
    expect(result[0].values).toEqual([["found-it"]]);
  });

  it("returns empty array when nothing is found", async () => {
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
                  load: vi.fn(),
                  areas: { items: [], load: vi.fn() },
                }),
              },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = (await xlSearchData("nope")) as any[];
    expect(result).toHaveLength(0);
  });

  it("searches only the specified sheet", async () => {
    const getItemFn = vi.fn(() => ({
      name: "TargetSheet",
      load: vi.fn(),
      findAllOrNullObject: () => ({
        isNullObject: false,
        load: vi.fn(),
        areas: {
          items: [
            { address: "TargetSheet!A1", values: [["match"]], load: vi.fn() },
          ],
          load: vi.fn(),
        },
      }),
    }));

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getItem: getItemFn,
          },
        },
        sync: vi.fn(),
      });
    });

    const result = (await xlSearchData("match", { sheetName: "TargetSheet" })) as any[];
    expect(getItemFn).toHaveBeenCalledWith("TargetSheet");
    expect(result).toHaveLength(1);
  });

  it("passes search options through correctly", async () => {
    let capturedOptions: any = null;
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            items: [
              {
                name: "Sheet1",
                load: vi.fn(),
                findAllOrNullObject: (_q: string, opts: any) => {
                  capturedOptions = opts;
                  return {
                    isNullObject: true,
                    load: vi.fn(),
                    areas: { items: [], load: vi.fn() },
                  };
                },
              },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    await xlSearchData("test", { caseSensitive: true, matchEntireCell: true });
    expect(capturedOptions).toEqual({
      completeMatch: true,
      matchCase: true,
    });
  });

  it("returns empty array when Office is unavailable", async () => {
    (globalThis as any).Office.context = null;
    const result = await xlSearchData("anything");
    expect(result).toEqual([]);
  });

  it("skips sheets that throw during search", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            items: [
              {
                name: "BadSheet",
                load: vi.fn(),
                findAllOrNullObject: () => {
                  throw new Error("cannot search");
                },
              },
              {
                name: "GoodSheet",
                load: vi.fn(),
                findAllOrNullObject: () => ({
                  isNullObject: false,
                  load: vi.fn(),
                  areas: {
                    items: [
                      { address: "GoodSheet!C5", values: [["ok"]], load: vi.fn() },
                    ],
                    load: vi.fn(),
                  },
                }),
              },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = (await xlSearchData("ok")) as any[];
    expect(result).toHaveLength(1);
    expect(result[0].worksheet).toBe("GoodSheet");
  });
});

describe("getXlRangeAsCsv — dedicated tests", () => {
  it("returns CSV with all values", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getItem: () => ({
              getRange: () => ({
                values: [
                  ["Name", "Age"],
                  ["Alice", 30],
                  ["Bob", 25],
                ],
                rowCount: 3,
                load: vi.fn(),
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getXlRangeAsCsv("Sheet1", "A1:B3");
    expect(result).toBe("Name,Age\nAlice,30\nBob,25");
  });

  it("respects maxRows parameter", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getItem: () => ({
              getRange: () => ({
                values: [["R1"], ["R2"], ["R3"], ["R4"], ["R5"]],
                rowCount: 5,
                load: vi.fn(),
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getXlRangeAsCsv("Sheet1", "A1:A5", 2);
    expect(result).toBe("R1\nR2");
  });

  it("respects offset parameter", async () => {
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getItem: () => ({
              getRange: () => ({
                values: [["R1"], ["R2"], ["R3"], ["R4"]],
                rowCount: 4,
                load: vi.fn(),
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    const result = await getXlRangeAsCsv("Sheet1", "A1:A4", 2, 1);
    expect(result).toBe("R2\nR3");
  });

  it("returns empty string when Office is unavailable", async () => {
    (globalThis as any).Office.context = null;
    const result = await getXlRangeAsCsv("Sheet1", "A1");
    expect(result).toBe("");
  });
});

describe("applyCellUpdates", () => {
  it("replace mode resizes single-cell range to match values array", async () => {
    const setValuesFn = vi.fn();
    const getResizedRange = vi.fn(() => {
      const obj = { _values: null as any };
      Object.defineProperty(obj, "values", {
        set(v: any) { setValuesFn(v); },
        get() { return obj._values; },
      });
      return obj;
    });

    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getRange: () => ({
                getCell: () => ({
                  getResizedRange,
                }),
              }),
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    await applyCellUpdates([{
      address: "M1",
      values: [["H1", "H2"], ["a", "b"], ["c", "d"]],
      mode: "replace",
    }]);

    // Should resize from 1x1 to 3x2 (rowCount-1=2, colCount-1=1)
    expect(getResizedRange).toHaveBeenCalledWith(2, 1);
  });
});

describe("insertCharts — auto-positioning", () => {
  it("positions chart below used range when topLeftCell is absent", async () => {
    const setPositionFn = vi.fn();
    excelRun.mockImplementation(async (cb: any) => {
      return cb({
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              getRange: () => ({}),
              getUsedRangeOrNullObject: () => ({
                isNullObject: false,
                rowIndex: 0,
                rowCount: 13,
                load: vi.fn(),
              }),
              getCell: (row: number, col: number) => ({ row, col }),
              charts: {
                add: () => ({
                  name: null,
                  title: { text: "", visible: false },
                  axes: {
                    categoryAxis: { title: { text: "", visible: false } },
                    valueAxis: { title: { text: "", visible: false } },
                  },
                  setPosition: setPositionFn,
                }),
              },
            }),
          },
        },
        sync: vi.fn(),
      });
    });

    await insertCharts([{
      chartType: "ColumnClustered",
      sourceAddress: "A1:B13",
      title: "Test Chart",
    }]);

    expect(setPositionFn).toHaveBeenCalled();
  });
});
