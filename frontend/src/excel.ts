/* global Excel, Office */

import {
  CellSelection,
  CellUpdate,
  ChartInsert,
  FormatUpdate,
  PivotTableInsert,
  PivotTableDataField,
  SheetMetadata,
  UserContext,
  WorkbookMetadata,
  WorkbookToolCall,
  WorkbookToolResult
} from "./types";

type RangeReference = {
  sheet?: string;
  address: string;
};

const CELL_REF_WITH_SHEET =
  /(?:'(?<sheet>[^']+)'|(?<sheet>[A-Za-z0-9_]+))!(?<range>[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?)/gi;
const CELL_REF_NO_SHEET = /[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?/gi;

const isIsolatedMatch = (
  text: string,
  start: number,
  length: number
): boolean => {
  const before = start > 0 ? text[start - 1] : undefined;
  const after = text[start + length];
  const boundary = (char?: string) => !char || !/[A-Za-z0-9_]/.test(char);
  return boundary(before) && boundary(after);
};

const normalizeSheetName = (raw?: string): string | undefined => {
  if (!raw) {
    return undefined;
  }
  let sanitized = raw;
  if (sanitized.startsWith("'") && sanitized.endsWith("'")) {
    sanitized = sanitized.slice(1, -1).replace(/''/g, "'");
  }
  return sanitized;
};

const splitReference = (
  reference: string
): { sheetName?: string; rangeAddress: string } => {
  if (!reference.includes("!")) {
    return { rangeAddress: reference };
  }
  const segments = reference.split("!");
  const rangeAddress = segments.pop() ?? reference;
  const sheetName = normalizeSheetName(segments.join("!"));
  return { sheetName, rangeAddress };
};

const normalizeIdentifier = (value: string): string =>
  value
    .trim()
    .replace(/^xl/i, "")
    .replace(/[^a-z0-9]/gi, "")
    .toLowerCase();

const CHART_TYPE_ALIASES: Record<string, Excel.ChartType> = {
  scatter: Excel.ChartType.xyscatter,
  scatterplot: Excel.ChartType.xyscatter,
  scatterchart: Excel.ChartType.xyscatter,
  scattermarkers: Excel.ChartType.xyscatter,
  xyscattermarkers: Excel.ChartType.xyscatter,
  xyscatter: Excel.ChartType.xyscatter,
  scatterlines: Excel.ChartType.xyscatterLines,
  scatterline: Excel.ChartType.xyscatterLines,
  xyscatterlines: Excel.ChartType.xyscatterLines,
  scatterlinesnomarkers: Excel.ChartType.xyscatterLinesNoMarkers,
  xyscatterlinesnomarkers: Excel.ChartType.xyscatterLinesNoMarkers,
  scatterlinenomarkers: Excel.ChartType.xyscatterLinesNoMarkers,
  bubble: Excel.ChartType.bubble,
  line: Excel.ChartType.lineMarkers,
  column: Excel.ChartType.columnClustered,
  bar: Excel.ChartType.barClustered
};

const resolveChartType = (rawType: string): Excel.ChartType | null => {
  if (!rawType) {
    return null;
  }
  const normalizedInput = normalizeIdentifier(rawType);
  if (normalizedInput in CHART_TYPE_ALIASES) {
    return CHART_TYPE_ALIASES[normalizedInput];
  }
  const chartTypeEntries = Object.entries(
    Excel.ChartType as unknown as Record<string, string>
  );
  for (const [key, value] of chartTypeEntries) {
    if (
      normalizeIdentifier(key) === normalizedInput ||
      normalizeIdentifier(value) === normalizedInput
    ) {
      return value as Excel.ChartType;
    }
  }
  return null;
};

// ---------------------------------------------------------------------------
// Workbook metadata & context (new)
// ---------------------------------------------------------------------------

/**
 * Collect workbook-level metadata: filename and all sheet names/dimensions.
 * Called once on add-in init; result is stored in React state and re-used
 * on every subsequent message send.
 *
 * @returns WorkbookMetadata or a failure object if Office is unavailable.
 */
export async function getWorkbookMetadata(): Promise<WorkbookMetadata> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return {
      success: false,
      fileName: "",
      sheetsMetadata: [],
      totalSheets: 0
    };
  }

  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.load("name");
      const sheets = workbook.worksheets;
      sheets.load("items");
      await context.sync();

      // Load per-sheet properties and used-range counts
      const sheetItems = sheets.items;
      const usedRanges = sheetItems.map((ws) => {
        ws.load(["id", "name", "position"]);
        return ws.getUsedRangeOrNullObject();
      });
      usedRanges.forEach((ur) => ur.load(["isNullObject", "rowCount", "columnCount"]));
      await context.sync();

      // Read first row (headers) for each sheet
      const headerRanges = sheetItems.map((ws, i) => {
        if (usedRanges[i].isNullObject || usedRanges[i].columnCount === 0) {
          return null;
        }
        const r = ws.getRangeByIndexes(0, 0, 1, usedRanges[i].columnCount);
        r.load("values");
        return r;
      });
      await context.sync();

      const sheetsMetadata: SheetMetadata[] = sheetItems.map((ws, i) => {
        const hasData = !usedRanges[i].isNullObject;
        const headers = headerRanges[i]
          ? (headerRanges[i]!.values[0] as (string | number | boolean | null)[]).map(
              (v) => (v == null ? "" : String(v))
            )
          : undefined;
        return {
          id: ws.id,
          name: ws.name,
          index: ws.position,
          maxRows: hasData ? usedRanges[i].rowCount : 0,
          maxColumns: hasData ? usedRanges[i].columnCount : 0,
          columnHeaders: headers
        };
      });

      return {
        success: true,
        fileName: workbook.name || "Workbook",
        sheetsMetadata,
        totalSheets: sheetItems.length
      };
    });
  } catch (error) {
    console.error("getWorkbookMetadata failed", error);
    return { success: false, fileName: "", sheetsMetadata: [], totalSheets: 0 };
  }
}

/**
 * Collect fresh per-request context: the active sheet name and the current
 * selection address string.
 *
 * @returns UserContext with active sheet name and selected range address.
 */
export async function getUserContext(): Promise<UserContext> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return { currentActiveSheetName: "", selectedRanges: "" };
  }
  try {
    return await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      const selection = context.workbook.getSelectedRange();
      selection.load("address");
      await context.sync();
      return {
        currentActiveSheetName: activeSheet.name,
        selectedRanges: selection.address
      };
    });
  } catch (error) {
    console.error("getUserContext failed", error);
    return { currentActiveSheetName: "", selectedRanges: "" };
  }
}

// ---------------------------------------------------------------------------
// Sheet preview cache — invalidated by worksheet change / sheet switch events
// ---------------------------------------------------------------------------

let previewCache: { csv: string } | null = null;
let previewDirty = true;

const markPreviewDirty = async () => {
  previewDirty = true;
  previewCache = null;
};

/** Reset cache state. Exported for testing only. */
export function resetPreviewCache(): void {
  previewDirty = true;
  previewCache = null;
}

/**
 * Register event listeners that invalidate the sheet preview cache when
 * the active sheet's data changes or the user switches sheets.
 * Call once on add-in init (alongside getWorkbookMetadata).
 */
export async function initPreviewCache(): Promise<void> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return;
  }
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(markPreviewDirty);
      context.workbook.worksheets.onActivated.add(markPreviewDirty);
      await context.sync();
    });
  } catch (error) {
    console.warn("initPreviewCache: could not register listeners", error);
  }
}

/**
 * Read the first {@link maxRows} rows of the active sheet as a CSV string.
 * Used as a lightweight "sheet preview" attached to every chat request so the
 * LLM has structural awareness without a full tool call.
 *
 * Returns a cached result when the sheet data has not changed since the last
 * read. The cache is invalidated by event listeners registered via
 * {@link initPreviewCache}.
 *
 * @param maxRows - Maximum number of rows to read (default 50).
 * @returns CSV string, or null when the sheet is empty or on error.
 */
export async function getLightweightSheetPreview(
  maxRows = 50
): Promise<string | null> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return null;
  }
  if (!previewDirty && previewCache) {
    return previewCache.csv;
  }
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load(["isNullObject", "rowCount", "columnCount"]);
      await context.sync();

      if (usedRange.isNullObject || usedRange.rowCount === 0) {
        previewCache = null;
        previewDirty = false;
        return null;
      }

      const rowsToRead = Math.min(maxRows, usedRange.rowCount);
      const previewRange = sheet.getRangeByIndexes(
        0, 0, rowsToRead, usedRange.columnCount
      );
      previewRange.load("values");
      await context.sync();

      const rows = previewRange.values as (string | number | boolean | null)[][];

      // Build column-letter header row so the LLM can map positions to Excel columns
      const colCount = usedRange.columnCount;
      const colLetters: string[] = [];
      for (let i = 0; i < colCount; i++) {
        let letter = "";
        let n = i;
        while (n >= 0) {
          letter = String.fromCharCode(65 + (n % 26)) + letter;
          n = Math.floor(n / 26) - 1;
        }
        colLetters.push(`[${letter}]`);
      }
      const letterRow = colLetters.join(",");

      const csv = letterRow + "\n" + rows
        .map((row) =>
          row
            .map((cell) => {
              const s = cell == null ? "" : String(cell);
              return s.includes(",") || s.includes("\n") ? `"${s.replace(/"/g, '""')}"` : s;
            })
            .join(",")
        )
        .join("\n");

      previewCache = { csv };
      previewDirty = false;
      return csv;
    });
  } catch (error) {
    console.error("getLightweightSheetPreview failed", error);
    return null;
  }
}

// ---------------------------------------------------------------------------
// Excel read tools (called on-demand via tool_call_required)
// ---------------------------------------------------------------------------

/**
 * Read one or more ranges with values, formulas, and basic formatting.
 *
 * @param ranges - Array of range addresses (e.g. ["A1:C10"]).
 * @param sheetName - Worksheet name; defaults to active sheet if omitted.
 * @returns Array of range detail objects.
 */
export async function getXlCellRanges(
  ranges: string[],
  sheetName?: string
): Promise<unknown> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return [];
  }
  return Excel.run(async (context) => {
    const ws =
      sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

    const rangeProxies = ranges.map((addr) => {
      const range = ws.getRange(addr);
      range.load(["address", "values", "formulas", "numberFormat"]);
      range.format.fill.load("color");
      range.format.font.load(["color", "bold", "italic"]);
      return range;
    });
    await context.sync();

    return rangeProxies.map((range) => ({
      address: range.address,
      values: range.values,
      formulas: range.formulas,
      numberFormat: range.numberFormat,
      fillColor: range.format.fill.color,
      fontColor: range.format.font.color,
      bold: range.format.font.bold,
      italic: range.format.font.italic
    }));
  });
}

/**
 * Read a range from a sheet and return its data as a CSV string.
 * Supports row-offset pagination for large sheets.
 *
 * @param sheetName - Worksheet name.
 * @param range - Range address (e.g. "A1:D200").
 * @param maxRows - Maximum rows to return.
 * @param offset - Number of rows to skip from the top of the range.
 * @returns CSV string.
 */
export async function getXlRangeAsCsv(
  sheetName: string,
  range: string,
  maxRows?: number,
  offset?: number
): Promise<string> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return "";
  }
  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(sheetName);
    let r = ws.getRange(range);
    r.load(["values", "rowCount"]);
    await context.sync();

    let allValues = r.values as (string | number | boolean | null)[][];

    const startRow = offset ?? 0;
    const endRow = maxRows != null ? startRow + maxRows : allValues.length;
    const rows = allValues.slice(startRow, endRow);

    return rows
      .map((row) =>
        row
          .map((cell) => {
            const s = cell == null ? "" : String(cell);
            return s.includes(",") || s.includes("\n")
              ? `"${s.replace(/"/g, '""')}"`
              : s;
          })
          .join(",")
      )
      .join("\n");
  });
}

/**
 * Search for text or values across one or all sheets.
 *
 * @param query - The value to search for.
 * @param options - Optional search configuration.
 * @returns Array of matches with address, sheet name, and cell value.
 */
export async function xlSearchData(
  query: string,
  options: {
    sheetName?: string;
    caseSensitive?: boolean;
    matchEntireCell?: boolean;
  } = {}
): Promise<unknown> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return [];
  }
  return Excel.run(async (context) => {
    const matches: unknown[] = [];
    const searchOptions: Excel.SearchCriteria = {
      completeMatch: options.matchEntireCell ?? false,
      matchCase: options.caseSensitive ?? false
    };

    const sheetsToSearch: Excel.Worksheet[] = [];
    if (options.sheetName) {
      sheetsToSearch.push(context.workbook.worksheets.getItem(options.sheetName));
    } else {
      const allSheets = context.workbook.worksheets;
      allSheets.load("items");
      await context.sync();
      sheetsToSearch.push(...allSheets.items);
    }

    for (const ws of sheetsToSearch) {
      ws.load("name");
      try {
        const foundRanges = ws.findAllOrNullObject(query, searchOptions);
        foundRanges.load("isNullObject");
        await context.sync();
        if (!foundRanges.isNullObject) {
          foundRanges.areas.load("items");
          await context.sync();
          for (const area of foundRanges.areas.items) {
            area.load(["address", "values"]);
          }
          await context.sync();
          for (const area of foundRanges.areas.items) {
            matches.push({
              address: area.address,
              worksheet: ws.name,
              values: area.values
            });
          }
        }
      } catch {
        // Sheet may not support search — skip silently
      }
    }
    return matches;
  });
}

/**
 * Discover charts, tables, and pivot tables in the workbook.
 *
 * @param sheetName - Limit results to this sheet (optional).
 * @param objectType - Filter by "chart", "table", or "pivot" (optional).
 * @returns Object inventory structured by type.
 */
export async function getAllXlObjects(
  sheetName?: string,
  objectType?: "chart" | "table" | "pivot"
): Promise<unknown> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return { charts: [], tables: [], pivotTables: [] };
  }
  return Excel.run(async (context) => {
    const sheetsToScan: Excel.Worksheet[] = [];
    if (sheetName) {
      sheetsToScan.push(context.workbook.worksheets.getItem(sheetName));
    } else {
      const allSheets = context.workbook.worksheets;
      allSheets.load("items");
      await context.sync();
      sheetsToScan.push(...allSheets.items);
    }

    const charts: unknown[] = [];
    const tables: unknown[] = [];
    const pivotTables: unknown[] = [];

    // Phase 1: load all collections for all sheets in one sync
    const sheetCollections = sheetsToScan.map((ws) => {
      ws.load("name");
      const c = (!objectType || objectType === "chart") ? ws.charts : null;
      const t = (!objectType || objectType === "table") ? ws.tables : null;
      const p = (!objectType || objectType === "pivot") ? ws.pivotTables : null;
      c?.load("items");
      t?.load("items");
      p?.load("items");
      return { ws, c, t, p };
    });
    await context.sync();

    // Phase 2: load item properties for all items in one sync
    for (const { c, t, p } of sheetCollections) {
      c?.items.forEach((ch) => ch.load(["name", "chartType"]));
      t?.items.forEach((tb) => tb.load(["name", "showTotals"]));
      p?.items.forEach((pv) => pv.load("name"));
    }
    await context.sync();

    // Extract results from populated proxies
    for (const { ws, c, t, p } of sheetCollections) {
      c?.items.forEach((ch) => {
        charts.push({ name: ch.name, type: ch.chartType, sheet: ws.name });
      });
      t?.items.forEach((tb) => {
        tables.push({ name: tb.name, sheet: ws.name });
      });
      p?.items.forEach((pv) => {
        pivotTables.push({ name: pv.name, sheet: ws.name });
      });
    }

    return { charts, tables, pivotTables };
  });
}

/**
 * Execute an Office.js code snippet inside an Excel.run context.
 * The snippet receives ``context`` (Excel.RequestContext) and ``Excel``
 * as parameters and should return a JSON-serializable value.
 *
 * CAUTION: Only use this for read operations from trusted LLM responses.
 *
 * @param code - JavaScript code string that returns a value.
 * @returns The value returned by the snippet, or an error string.
 */
export async function executeXlOfficeJs(code: string): Promise<unknown> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return { error: "Office not available" };
  }
  return Excel.run(async (context) => {
    try {
      // eslint-disable-next-line no-new-func
      const fn = new Function(
        "context",
        "Excel",
        `return (async () => { ${code} })()`
      );
      const result = await fn(context, Excel);
      await context.sync();
      return result;
    } catch (err) {
      return { error: String(err) };
    }
  });
}

/**
 * Dispatch a WorkbookToolCall to the appropriate Office.js function.
 * Wraps execution in try/catch and populates the ``error`` field on failure.
 *
 * @param call - The tool call emitted by the server.
 * @returns A WorkbookToolResult with the execution outcome.
 */
export async function executeWorkbookTool(
  call: WorkbookToolCall
): Promise<WorkbookToolResult> {
  try {
    let result: unknown;
    const a = call.args;
    switch (call.tool) {
      case "get_xl_cell_ranges":
        result = await getXlCellRanges(
          (a.ranges as string[]) || [],
          (a.sheetName as string) || undefined
        );
        break;
      case "get_xl_range_as_csv":
        result = await getXlRangeAsCsv(
          (a.sheetName as string) || "",
          (a.range as string) || "A1",
          a.maxRows != null ? Number(a.maxRows) : undefined,
          a.offset != null ? Number(a.offset) : undefined
        );
        break;
      case "xl_search_data":
        result = await xlSearchData((a.query as string) || "", {
          sheetName: (a.sheetName as string) || undefined,
          caseSensitive: Boolean(a.caseSensitive),
          matchEntireCell: Boolean(a.matchEntireCell)
        });
        break;
      case "get_all_xl_objects":
        result = await getAllXlObjects(
          (a.sheetName as string) || undefined,
          (a.objectType as "chart" | "table" | "pivot") || undefined
        );
        break;
      case "execute_xl_office_js":
        result = await executeXlOfficeJs((a.code as string) || "");
        break;
      default:
        return {
          id: call.id,
          tool: call.tool,
          result: null,
          error: `Unknown tool: ${call.tool}`
        };
    }
    return { id: call.id, tool: call.tool, result };
  } catch (err) {
    return {
      id: call.id,
      tool: call.tool,
      result: null,
      error: err instanceof Error ? err.message : String(err)
    };
  }
}

// ---------------------------------------------------------------------------
// Existing read / write helpers (unchanged)
// ---------------------------------------------------------------------------

export async function getCurrentSelection(): Promise<CellSelection[]> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return [];
  }

  try {
    return await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values"]);
      const worksheet = range.worksheet;
      worksheet.load("name");
      await context.sync();
      return [
        {
          address: range.address,
          values: range.values as (string | number | boolean | null)[][],
          worksheet: worksheet.name
        }
      ];
    });
  } catch (error) {
    console.error("Unable to read selected range", error);
    return [];
  }
}

export async function applyCellUpdates(updates: CellUpdate[]): Promise<void> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return;
  }

  await Excel.run(async (context) => {
    for (const update of updates) {
      const { sheetName: addressSheet, rangeAddress } = splitReference(
        update.address
      );
      const worksheetHint =
        update.worksheet != null ? normalizeSheetName(update.worksheet) : undefined;
      const targetSheet = worksheetHint ?? addressSheet;
      const worksheet =
        targetSheet != null
          ? context.workbook.worksheets.getItem(targetSheet)
          : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(rangeAddress);
      switch (update.mode) {
        case "append": {
          const rowCount = update.values.length;
          const colCount = update.values[0]?.length ?? 1;
          const target = range.getResizedRange(rowCount - 1, colCount - 1);
          target.values = update.values;
          break;
        }
        case "replace":
        default: {
          const rowCount = update.values.length;
          const colCount = update.values[0]?.length ?? 1;
          const topLeft = range.getCell(0, 0);
          const target = topLeft.getResizedRange(rowCount - 1, colCount - 1);
          target.values = update.values;
          break;
        }
      }
    }
    await context.sync();
  });
}

export async function applyFormatUpdates(
  updates: FormatUpdate[]
): Promise<void> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return;
  }
  if (updates.length === 0) {
    return;
  }

  await Excel.run(async (context) => {
    for (const update of updates) {
      try {
        const { sheetName: addressSheet, rangeAddress } = splitReference(
          update.address
        );
        const worksheetHint =
          update.worksheet != null
            ? normalizeSheetName(update.worksheet)
            : undefined;
        const targetSheet = worksheetHint ?? addressSheet;
        const worksheet =
          targetSheet != null
            ? context.workbook.worksheets.getItem(targetSheet)
            : context.workbook.worksheets.getActiveWorksheet();
        const range = worksheet.getRange(rangeAddress);
        if (update.fillColor) {
          range.format.fill.color = update.fillColor;
        }
        if (update.fontColor) {
          range.format.font.color = update.fontColor;
        }
        if (typeof update.bold === "boolean") {
          range.format.font.bold = update.bold;
        }
        if (typeof update.italic === "boolean") {
          range.format.font.italic = update.italic;
        }
        if (update.numberFormat) {
          range.numberFormat = [[update.numberFormat]];
        }
        if (update.borderColor || update.borderStyle || update.borderWeight) {
          const borders = range.format.borders;
          const applyBorder = (index: Excel.BorderIndex) => {
            const border = borders.getItem(index);
            if (update.borderColor) {
              border.color = update.borderColor;
            }
            if (update.borderStyle) {
              border.style = update.borderStyle as Excel.BorderLineStyle;
            }
            if (update.borderWeight) {
              border.weight = update.borderWeight as Excel.BorderWeight;
            }
          };
          applyBorder(Excel.BorderIndex.edgeTop);
          applyBorder(Excel.BorderIndex.edgeBottom);
          applyBorder(Excel.BorderIndex.edgeLeft);
          applyBorder(Excel.BorderIndex.edgeRight);
        }
      } catch (error) {
        console.error("Unable to apply format update", update, error);
      }
    }
    await context.sync();
  });
}

export async function insertCharts(inserts: ChartInsert[]): Promise<void> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return;
  }
  if (inserts.length === 0) {
    return;
  }

  await Excel.run(async (context) => {
    for (const insert of inserts) {
      try {
        const { sheetName: inferredSheet, rangeAddress } = splitReference(
          insert.sourceAddress
        );
        const sourceSheetName =
          normalizeSheetName(insert.sourceWorksheet ?? undefined) ??
          inferredSheet;
        const sourceWorksheet =
          sourceSheetName != null
            ? context.workbook.worksheets.getItem(sourceSheetName)
            : context.workbook.worksheets.getActiveWorksheet();
        // Handle non-contiguous ranges (comma-separated like "A1:A13,G1:G13")
        let sourceData: any;
        if (rangeAddress.includes(",")) {
          try {
            sourceData = (sourceWorksheet as any).getRanges(rangeAddress);
          } catch {
            // Fallback for older API: use bounding-box single range
            sourceData = sourceWorksheet.getRange(rangeAddress.split(",")[0]);
          }
        } else {
          sourceData = sourceWorksheet.getRange(rangeAddress);
        }

        const destinationSheetName = normalizeSheetName(
          insert.destinationWorksheet ?? undefined
        );
        const destinationWorksheet =
          destinationSheetName != null
            ? context.workbook.worksheets.getItem(destinationSheetName)
            : sourceWorksheet;

        const chartType = resolveChartType(insert.chartType);
        if (!chartType) {
          console.warn(
            "Unsupported chart type provided; skipping chart insertion",
            insert.chartType
          );
          continue;
        }
        let seriesBy: Excel.ChartSeriesBy;
        switch ((insert.seriesBy ?? "auto").toLowerCase()) {
          case "rows":
            seriesBy = Excel.ChartSeriesBy.rows;
            break;
          case "columns":
            seriesBy = Excel.ChartSeriesBy.columns;
            break;
          default:
            seriesBy = Excel.ChartSeriesBy.auto;
        }

        const chart = destinationWorksheet.charts.add(
          chartType,
          sourceData,
          seriesBy
        );
        if (insert.name) {
          chart.name = insert.name;
        }
        if (insert.title) {
          chart.title.text = insert.title;
          chart.title.visible = true;
        }
        if (insert.xAxisTitle) {
          chart.axes.categoryAxis.title.text = insert.xAxisTitle;
          chart.axes.categoryAxis.title.visible = true;
        }
        if (insert.yAxisTitle) {
          chart.axes.valueAxis.title.text = insert.yAxisTitle;
          chart.axes.valueAxis.title.visible = true;
        }
        if (insert.topLeftCell) {
          const topLeft = destinationWorksheet.getRange(insert.topLeftCell);
          const bottomRight = insert.bottomRightCell
            ? destinationWorksheet.getRange(insert.bottomRightCell)
            : undefined;
          chart.setPosition(topLeft, bottomRight);
        } else {
          const usedRange = destinationWorksheet.getUsedRangeOrNullObject();
          usedRange.load(["isNullObject", "rowIndex", "rowCount"]);
          await context.sync();
          if (!usedRange.isNullObject) {
            const targetRow = usedRange.rowIndex + usedRange.rowCount + 1;
            const topLeft = destinationWorksheet.getCell(targetRow, 0);
            chart.setPosition(topLeft);
          }
        }
      } catch (error) {
        console.error("Unable to insert chart", insert, error);
      }
    }
    await context.sync();
  });
}

const AGGREGATION_FUNCTION_MAP: Record<string, Excel.AggregationFunction> = {
  sum: Excel.AggregationFunction.sum,
  count: Excel.AggregationFunction.count,
  average: Excel.AggregationFunction.average,
  max: Excel.AggregationFunction.max,
  min: Excel.AggregationFunction.min,
  product: Excel.AggregationFunction.product,
  countNumbers: Excel.AggregationFunction.countNumbers,
  standardDeviation: Excel.AggregationFunction.standardDeviation,
  standardDeviationP: Excel.AggregationFunction.standardDeviationP,
  variance: Excel.AggregationFunction.variance,
  varianceP: Excel.AggregationFunction.varianceP
};

/**
 * Create pivot tables in Excel from the LLM response.
 * Uses the Office.js PivotTable API (ExcelApi 1.8+) to add pivot tables
 * and configure row, column, data, and filter hierarchies.
 *
 * @param inserts - Array of PivotTableInsert definitions from the LLM.
 */
export async function insertPivotTables(
  inserts: PivotTableInsert[]
): Promise<void> {
  if (!Office.context || Office.context.host !== Office.HostType.Excel) {
    return;
  }
  if (inserts.length === 0) {
    return;
  }

  await Excel.run(async (context) => {
    for (const insert of inserts) {
      try {
        // Resolve source range
        const { sheetName: inferredSheet, rangeAddress } = splitReference(
          insert.sourceAddress
        );
        const sourceSheetName =
          normalizeSheetName(insert.sourceWorksheet ?? undefined) ??
          inferredSheet;
        const sourceWorksheet =
          sourceSheetName != null
            ? context.workbook.worksheets.getItem(sourceSheetName)
            : context.workbook.worksheets.getActiveWorksheet();
        const sourceRange = sourceWorksheet.getRange(rangeAddress);

        // Resolve destination address parts (strip sheet prefix like "Sheet2!E1")
        const {
          sheetName: destInferredSheet,
          rangeAddress: destRangeAddress,
        } = insert.destinationAddress
          ? splitReference(insert.destinationAddress)
          : { sheetName: undefined, rangeAddress: undefined };

        // Resolve destination sheet — prefer explicit, then inferred from address, then source sheet
        const destSheetName =
          normalizeSheetName(insert.destinationWorksheet ?? undefined) ??
          destInferredSheet;
        let destWorksheet: Excel.Worksheet;
        if (destSheetName != null) {
          // User/LLM requested a specific sheet — create it if missing
          const existing = context.workbook.worksheets.getItemOrNullObject(destSheetName);
          existing.load("isNullObject");
          await context.sync();
          destWorksheet = existing.isNullObject
            ? context.workbook.worksheets.add(destSheetName)
            : existing;
        } else {
          destWorksheet = sourceWorksheet;
        }

        // Resolve destination cell — if not specified, find empty area to the right of data
        let destRange: Excel.Range;
        if (destRangeAddress) {
          destRange = destWorksheet.getRange(destRangeAddress);
        } else {
          // Place 2 columns to the right of used data to avoid overlap
          const usedRange = destWorksheet.getUsedRangeOrNullObject();
          usedRange.load(["isNullObject", "columnIndex", "columnCount"]);
          await context.sync();
          if (!usedRange.isNullObject) {
            const destCol = usedRange.columnIndex + usedRange.columnCount + 1;
            destRange = destWorksheet.getCell(0, destCol);
          } else {
            destRange = destWorksheet.getRange("A1");
          }
        }

        // Create the pivot table and sync to materialize it before accessing hierarchies
        const pivotTable = destWorksheet.pivotTables.add(
          insert.name,
          sourceRange,
          destRange
        );
        await context.sync();

        // Add row hierarchies
        for (const rowField of insert.rows ?? []) {
          pivotTable.rowHierarchies.add(
            pivotTable.hierarchies.getItem(rowField)
          );
        }
        if (insert.rows?.length) await context.sync();

        // Add column hierarchies
        for (const colField of insert.columns ?? []) {
          pivotTable.columnHierarchies.add(
            pivotTable.hierarchies.getItem(colField)
          );
        }
        if (insert.columns?.length) await context.sync();

        // Add data (values) hierarchies
        for (const dataField of insert.values ?? []) {
          const dataHierarchy = pivotTable.dataHierarchies.add(
            pivotTable.hierarchies.getItem(dataField.name)
          );
          const aggKey = dataField.summarizeBy ?? "sum";
          if (aggKey in AGGREGATION_FUNCTION_MAP) {
            dataHierarchy.summarizeBy = AGGREGATION_FUNCTION_MAP[aggKey];
          }
        }
        if (insert.values?.length) await context.sync();

        // Add filter hierarchies
        for (const filterField of insert.filters ?? []) {
          pivotTable.filterHierarchies.add(
            pivotTable.hierarchies.getItem(filterField)
          );
        }
        if (insert.filters?.length) await context.sync();
      } catch (error) {
        console.error("Unable to insert pivot table", insert, error);
      }
    }
    await context.sync();
  });
}

export function extractRangeReferences(prompt: string): RangeReference[] {
  const references = new Map<string, RangeReference>();

  for (const match of prompt.matchAll(CELL_REF_WITH_SHEET)) {
    const sheetRaw = match.groups?.sheet;
    const rangeAddress = match.groups?.range;
    if (!rangeAddress) {
      continue;
    }
    const start = match.index ?? prompt.indexOf(match[0]);
    if (!isIsolatedMatch(prompt, start, match[0].length)) {
      continue;
    }
    const sheetName = normalizeSheetName(sheetRaw);
    const key = `${sheetName ?? ""}!${rangeAddress}`.toUpperCase();
    references.set(key, { sheet: sheetName, address: rangeAddress });
  }

  for (const match of prompt.matchAll(CELL_REF_NO_SHEET)) {
    const rangeAddress = match[0];
    if (!rangeAddress) {
      continue;
    }
    const start = match.index ?? prompt.indexOf(match[0]);
    if (start > 0 && prompt[start - 1] === "!") {
      continue;
    }
    if (!isIsolatedMatch(prompt, start, match[0].length)) {
      continue;
    }
    const key = `!${rangeAddress}`.toUpperCase();
    if (!references.has(key)) {
      references.set(key, { address: rangeAddress });
    }
  }

  return Array.from(references.values());
}

export async function getSelectionsFromReferences(
  references: RangeReference[]
): Promise<CellSelection[]> {
  if (
    !references.length ||
    !Office.context ||
    Office.context.host !== Office.HostType.Excel
  ) {
    return [];
  }

  try {
    return await Excel.run(async (context) => {
      const items: Array<{
        reference: RangeReference;
        range: Excel.Range;
        worksheet: Excel.Worksheet;
      }> = [];

      for (const reference of references) {
        const { sheetName, rangeAddress } = splitReference(reference.address);
        const targetSheet = reference.sheet ?? sheetName;
        const worksheet =
          targetSheet != null
            ? context.workbook.worksheets.getItem(targetSheet)
            : context.workbook.worksheets.getActiveWorksheet();
        const range = worksheet.getRange(rangeAddress);
        range.load(["address", "values"]);
        worksheet.load("name");
        items.push({ reference, range, worksheet });
      }

      await context.sync();

      return items.map(({ reference, range, worksheet }) => {
        const worksheetName = reference.sheet ?? worksheet.name;
        return {
          address: range.address,
          values: range.values as (string | number | boolean | null)[][],
          worksheet: worksheetName
        };
      });
    });
  } catch (error) {
    console.error("Unable to read referenced ranges", references, error);
    return [];
  }
}

export async function getSelectionsFromPrompt(
  prompt: string
): Promise<CellSelection[]> {
  const references = extractRangeReferences(prompt);
  if (references.length === 0) {
    return [];
  }
  return getSelectionsFromReferences(references);
}
