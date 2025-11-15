/* global Excel, Office */

import {
  CellSelection,
  CellUpdate,
  ChartInsert,
  FormatUpdate
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
          range.values = update.values;
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
        const sourceRange = sourceWorksheet.getRange(rangeAddress);

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
          sourceRange,
          seriesBy
        );
        if (insert.name) {
          chart.name = insert.name;
        }
        if (insert.title) {
          chart.title.text = insert.title;
          chart.title.visible = true;
        }
        if (insert.topLeftCell) {
          const topLeft = destinationWorksheet.getRange(insert.topLeftCell);
          const bottomRight = insert.bottomRightCell
            ? destinationWorksheet.getRange(insert.bottomRightCell)
            : undefined;
          chart.setPosition(topLeft, bottomRight);
        }
      } catch (error) {
        console.error("Unable to insert chart", insert, error);
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

