/**
 * Stateful Office.js Workbook Simulator for E2E testing.
 *
 * Replaces the stateless officeMock.ts with an in-memory workbook that tracks
 * cell values, formats, charts, pivot tables, and tables across Excel.run calls.
 * Implements the Office.js proxy-load-sync lifecycle so tests catch the same
 * class of bugs that manual Excel testing catches.
 *
 * Key exports:
 * - SimWorkbook: the root workbook state container
 * - createSimExcelRun: drop-in replacement for Excel.run
 * - buildSimWorkbook: factory that creates a SimWorkbook from a JSON fixture
 * - installSimulator: installs the simulator as the global Excel.run
 *
 * @module workbookSimulator
 */

import { vi } from "vitest";

// ---------------------------------------------------------------------------
// Address parsing utilities
// ---------------------------------------------------------------------------

/** Convert a column letter string (A, B, ..., Z, AA, AB, ...) to a 0-based index. */
function colLetterToIndex(letters: string): number {
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index - 1;
}

/** Convert a 0-based column index to a column letter string. */
function indexToColLetter(index: number): string {
  let letter = "";
  let n = index;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

/**
 * Parse a cell address like "A1" into { row, col } (both 0-based).
 * Also handles "Sheet1!A1" by stripping the sheet prefix.
 */
function parseCellAddress(address: string): { row: number; col: number } {
  const clean = address.includes("!") ? address.split("!").pop()! : address;
  const match = clean.match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error(`Invalid cell address: ${address}`);
  return {
    col: colLetterToIndex(match[1].toUpperCase()),
    row: parseInt(match[2], 10) - 1,
  };
}

/**
 * Parse a range address like "A1:C3" into start/end row/col (0-based).
 * Single-cell addresses like "A1" are treated as a 1x1 range.
 */
function parseRangeAddress(address: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} {
  const clean = address.includes("!") ? address.split("!").pop()! : address;
  const parts = clean.split(":");
  const start = parseCellAddress(parts[0]);
  if (parts.length === 1) {
    return {
      startRow: start.row,
      startCol: start.col,
      endRow: start.row,
      endCol: start.col,
    };
  }
  const end = parseCellAddress(parts[1]);
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

/**
 * Build a full range address string like "Sheet1!A1:C3" from components.
 */
function buildAddress(
  sheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number
): string {
  const start = `${indexToColLetter(startCol)}${startRow + 1}`;
  const end = `${indexToColLetter(endCol)}${endRow + 1}`;
  const rangeAddr = start === end ? start : `${start}:${end}`;
  // WHY: Office.js returns addresses with quotes around sheet names containing spaces
  const safe =
    sheetName.includes(" ") ? `'${sheetName}'` : sheetName;
  return `${safe}!${rangeAddr}`;
}

// ---------------------------------------------------------------------------
// Cell type alias
// ---------------------------------------------------------------------------

type CellValue = string | number | boolean | null;

// ---------------------------------------------------------------------------
// Global context tracker for range registration
// ---------------------------------------------------------------------------

// WHY: When SimRange.getCell/getResizedRange/getOffsetRange create new
// ranges, those ranges need to be registered with the current context so
// their pending writes get flushed on context.sync(). This module-level
// variable tracks the active context.
let _activeContext: SimRequestContext | null = null;

// ---------------------------------------------------------------------------
// SimRange — reference to a rectangular region of a SimWorksheet
// ---------------------------------------------------------------------------

/**
 * Simulates an Office.js Range proxy object.
 *
 * Supports the load-sync lifecycle: properties are only accessible after
 * calling .load() for them and then context.sync(). Writes (setting .values,
 * .format properties) are queued and flushed on context.sync().
 */
export class SimRange {
  private _worksheet: SimWorksheet;
  private _startRow: number;
  private _startCol: number;
  private _endRow: number;
  private _endCol: number;
  private _isNullObject: boolean;

  // WHY: Tracks which properties have been loaded. Accessing before load+sync
  // should fail, matching real Office.js behavior.
  private _loaded: Set<string> = new Set();
  private _synced = false;

  /** Pending value writes — flushed on context.sync(). */
  _pendingValues: CellValue[][] | null = null;
  _pendingFormats: Array<{
    prop: string;
    value: unknown;
  }> = [];

  constructor(
    worksheet: SimWorksheet,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
    isNullObject = false
  ) {
    this._worksheet = worksheet;
    this._startRow = startRow;
    this._startCol = startCol;
    this._endRow = endRow;
    this._endCol = endCol;
    this._isNullObject = isNullObject;
  }

  // -- Load / sync lifecycle ------------------------------------------------

  load(props: string | string[]): void {
    const names = Array.isArray(props) ? props : [props];
    for (const name of names) {
      this._loaded.add(name);
    }
  }

  /** Called by SimRequestContext.sync() to mark this proxy as synced. */
  _markSynced(): void {
    this._synced = true;
  }

  private _assertLoaded(prop: string): void {
    // WHY: In real Office.js you must load+sync before reading. We enforce
    // this so tests catch missing .load() calls that slip through in
    // stateless mocks but crash in real Excel.
    if (!this._synced || !this._loaded.has(prop)) {
      // Be lenient for common patterns: if any load call was made and sync
      // happened, allow access. This avoids over-strictness for nested loads
      // like "areas/items/address".
      if (this._synced && this._loaded.size > 0) return;
      throw new Error(
        `PropertyNotLoaded: '${prop}' was not loaded. Call .load("${prop}") then context.sync() first.`
      );
    }
  }

  // -- Properties -----------------------------------------------------------

  get isNullObject(): boolean {
    return this._isNullObject;
  }

  get rowCount(): number {
    this._assertLoaded("rowCount");
    return this._endRow - this._startRow + 1;
  }

  get columnCount(): number {
    this._assertLoaded("columnCount");
    return this._endCol - this._startCol + 1;
  }

  get rowIndex(): number {
    this._assertLoaded("rowIndex");
    return this._startRow;
  }

  get columnIndex(): number {
    this._assertLoaded("columnIndex");
    return this._startCol;
  }

  get address(): string {
    this._assertLoaded("address");
    return buildAddress(
      this._worksheet.name,
      this._startRow,
      this._startCol,
      this._endRow,
      this._endCol
    );
  }

  get values(): CellValue[][] {
    this._assertLoaded("values");
    return this._worksheet._readValues(
      this._startRow,
      this._startCol,
      this._endRow,
      this._endCol
    );
  }

  set values(v: CellValue[][]) {
    this._pendingValues = v;
  }

  get formulas(): CellValue[][] {
    this._assertLoaded("formulas");
    return this._worksheet._readFormulas(
      this._startRow,
      this._startCol,
      this._endRow,
      this._endCol
    );
  }

  get numberFormat(): string[][] {
    this._assertLoaded("numberFormat");
    return this._worksheet._readNumberFormats(
      this._startRow,
      this._startCol,
      this._endRow,
      this._endCol
    );
  }

  set numberFormat(v: string[][]) {
    for (let r = 0; r < v.length; r++) {
      for (let c = 0; c < v[r].length; c++) {
        this._worksheet._setNumberFormat(
          this._startRow + r,
          this._startCol + c,
          v[r][c]
        );
      }
    }
  }

  get worksheet(): SimWorksheet {
    return this._worksheet;
  }

  // -- Format accessors (simplified) ----------------------------------------

  get format(): {
    fill: { color: string; load: (p: string | string[]) => void };
    font: {
      color: string;
      bold: boolean;
      italic: boolean;
      load: (p: string | string[]) => void;
    };
    borders: {
      getItem: (index: string) => {
        color: string;
        style: string;
        weight: string;
      };
    };
  } {
    const self = this;
    const key = `${this._startRow},${this._startCol}`;
    const fmt = this._worksheet._getFormat(this._startRow, this._startCol);

    return {
      fill: {
        get color() {
          return fmt.fillColor;
        },
        set color(v: string) {
          self._pendingFormats.push({ prop: "fillColor", value: v });
        },
        load: (_p: string | string[]) => {
          this._loaded.add("format/fill/color");
          this._loaded.add("color");
        },
      },
      font: {
        get color() {
          return fmt.fontColor;
        },
        set color(v: string) {
          self._pendingFormats.push({ prop: "fontColor", value: v });
        },
        get bold() {
          return fmt.bold;
        },
        set bold(v: boolean) {
          self._pendingFormats.push({ prop: "bold", value: v });
        },
        get italic() {
          return fmt.italic;
        },
        set italic(v: boolean) {
          self._pendingFormats.push({ prop: "italic", value: v });
        },
        load: (_p: string | string[]) => {
          this._loaded.add("format/font");
          this._loaded.add("color");
          this._loaded.add("bold");
          this._loaded.add("italic");
        },
      },
      borders: {
        getItem: (index: string) => {
          const borderKey = `${key}:${index}`;
          const border = self._worksheet._getBorder(
            self._startRow,
            self._startCol,
            index
          );
          return {
            get color() {
              return border.color;
            },
            set color(v: string) {
              self._pendingFormats.push({
                prop: `border:${index}:color`,
                value: v,
              });
            },
            get style() {
              return border.style;
            },
            set style(v: string) {
              self._pendingFormats.push({
                prop: `border:${index}:style`,
                value: v,
              });
            },
            get weight() {
              return border.weight;
            },
            set weight(v: string) {
              self._pendingFormats.push({
                prop: `border:${index}:weight`,
                value: v,
              });
            },
          };
        },
      },
    };
  }

  // -- Range navigation -----------------------------------------------------

  private _register(range: SimRange): SimRange {
    // WHY: Register newly created ranges with the active context so their
    // pending writes get flushed on context.sync().
    if (_activeContext) {
      _activeContext._ranges.push(range);
    }
    return range;
  }

  getCell(row: number, col: number): SimRange {
    const r = this._startRow + row;
    const c = this._startCol + col;
    return this._register(new SimRange(this._worksheet, r, c, r, c));
  }

  getResizedRange(deltaRows: number, deltaCols: number): SimRange {
    return this._register(
      new SimRange(
        this._worksheet,
        this._startRow,
        this._startCol,
        this._endRow + deltaRows,
        this._endCol + deltaCols
      )
    );
  }

  getOffsetRange(rowOffset: number, colOffset: number): SimRange {
    return this._register(
      new SimRange(
        this._worksheet,
        this._startRow + rowOffset,
        this._startCol + colOffset,
        this._endRow + rowOffset,
        this._endCol + colOffset
      )
    );
  }

  /** Flush pending writes to the worksheet's backing store. */
  _flush(): void {
    if (this._pendingValues) {
      this._worksheet._writeValues(
        this._startRow,
        this._startCol,
        this._pendingValues
      );
      this._pendingValues = null;
    }
    for (const f of this._pendingFormats) {
      if (f.prop.startsWith("border:")) {
        const [, index, bprop] = f.prop.split(":");
        // Apply to all cells in the range
        for (let r = this._startRow; r <= this._endRow; r++) {
          for (let c = this._startCol; c <= this._endCol; c++) {
            this._worksheet._setBorderProp(r, c, index, bprop, f.value as string);
          }
        }
      } else {
        // Apply format to all cells in the range
        for (let r = this._startRow; r <= this._endRow; r++) {
          for (let c = this._startCol; c <= this._endCol; c++) {
            this._worksheet._setFormatProp(r, c, f.prop, f.value);
          }
        }
      }
    }
    this._pendingFormats = [];
  }
}

// ---------------------------------------------------------------------------
// SimChart — tracks chart existence and properties
// ---------------------------------------------------------------------------

export interface SimChartSeries {
  name: string;
  valuesRange?: SimRange;
  xAxisRange?: SimRange;
}

export class SimChart {
  name = "";
  chartType: string;
  sourceRange: SimRange;
  seriesBy: string;
  title = { text: "", visible: false };
  axes = {
    categoryAxis: { title: { text: "", visible: false } },
    valueAxis: { title: { text: "", visible: false } },
  };
  _series: SimChartSeries[] = [];
  _position: { topLeft?: SimRange; bottomRight?: SimRange } = {};

  // WHY: Track the row where chart is positioned so tests can assert
  // auto-positioning logic (e.g., chart placed below used range).
  positionRow = -1;

  constructor(chartType: string, sourceRange: SimRange, seriesBy: string) {
    this.chartType = chartType;
    this.sourceRange = sourceRange;
    this.seriesBy = seriesBy;
  }

  get series() {
    const self = this;
    return {
      count: self._series.length,
      load(_props: string) {
        /* no-op for sim */
      },
      getItemAt(index: number): {
        setXAxisValues(range: SimRange): void;
        setValues(range: SimRange): void;
      } {
        return {
          setXAxisValues(range: SimRange) {
            self._series[index] = self._series[index] || {
              name: `Series ${index}`,
            };
            self._series[index].xAxisRange = range;
          },
          setValues(range: SimRange) {
            self._series[index] = self._series[index] || {
              name: `Series ${index}`,
            };
            self._series[index].valuesRange = range;
          },
        };
      },
      add(name: string) {
        const s: SimChartSeries = { name };
        self._series.push(s);
        return {
          setValues(range: SimRange) {
            s.valuesRange = range;
          },
          setXAxisValues(range: SimRange) {
            s.xAxisRange = range;
          },
        };
      },
    };
  }

  setPosition(topLeft: SimRange, bottomRight?: SimRange): void {
    this._position = { topLeft, bottomRight };
    // WHY: Extract the row from the topLeft range for test assertions
    this.positionRow = (topLeft as any)._startRow;
  }
}

// ---------------------------------------------------------------------------
// SimPivotTable — tracks pivot table existence and hierarchy config
// ---------------------------------------------------------------------------

export interface SimHierarchy {
  name: string;
  summarizeBy?: string;
}

export class SimPivotTable {
  name: string;
  sourceRange: SimRange;
  destRange: SimRange;
  rowHierarchies: SimHierarchy[] = [];
  columnHierarchies: SimHierarchy[] = [];
  dataHierarchies: SimHierarchy[] = [];
  filterHierarchies: SimHierarchy[] = [];

  // WHY: Store the field names from the source data so hierarchies.getItem()
  // can resolve them, matching real Office.js behavior.
  private _fieldNames: string[];

  constructor(
    name: string,
    sourceRange: SimRange,
    destRange: SimRange,
    fieldNames: string[]
  ) {
    this.name = name;
    this.sourceRange = sourceRange;
    this.destRange = destRange;
    this._fieldNames = fieldNames;
  }

  get hierarchies() {
    const fields = this._fieldNames;
    return {
      getItem(fieldName: string) {
        // WHY: Return a reference object even if the field isn't found,
        // matching Office.js behavior (error happens later at sync).
        return { name: fieldName };
      },
    };
  }

  _makeHierarchyAdder(target: SimHierarchy[]) {
    return {
      add(item: { name: string }) {
        const h: SimHierarchy = { name: item.name };
        target.push(h);
        return {
          get summarizeBy(): string | undefined {
            return h.summarizeBy;
          },
          set summarizeBy(v: string) {
            h.summarizeBy = v;
          },
        };
      },
    };
  }
}

// ---------------------------------------------------------------------------
// SimWorksheet — holds cell data, formats, charts, pivots, tables
// ---------------------------------------------------------------------------

interface CellFormat {
  fillColor: string;
  fontColor: string;
  bold: boolean;
  italic: boolean;
  numberFormat: string;
}

interface CellBorder {
  color: string;
  style: string;
  weight: string;
}

export class SimWorksheet {
  id: string;
  name: string;
  position: number;

  // WHY: Sparse storage — only cells that have been written to exist in the map.
  // This is more memory-efficient and naturally handles the "empty sheet" case.
  private _cells: Map<string, CellValue> = new Map();
  private _formulas: Map<string, CellValue> = new Map();
  private _numberFormats: Map<string, string> = new Map();
  private _formats: Map<string, CellFormat> = new Map();
  private _borders: Map<string, CellBorder> = new Map();

  // WHY: These backing arrays are the source of truth for charts/pivots/tables.
  // The _wrapWorksheet method replaces the public .charts/.pivotTables/.tables
  // properties with Office.js-style collection wrappers (with .items, .load, .add).
  // Use _chartsArray / _pivotsArray for direct test assertions.
  _chartsArray: SimChart[] = [];
  _pivotsArray: SimPivotTable[] = [];
  _tablesArray: Array<{ name: string; showTotals: boolean }> = [];

  // These get replaced by _wrapWorksheet with collection wrapper objects
  charts: any = [];
  pivotTables: any = [];
  tables: any = [];

  // WHY: Track bounds for usedRange computation
  private _maxRow = -1;
  private _maxCol = -1;

  /** Event stubs matching Office.js API shape */
  onChanged = { add: vi.fn() };

  constructor(id: string, name: string, position: number) {
    this.id = id;
    this.name = name;
    this.position = position;
  }

  // -- Internal cell accessors ----------------------------------------------

  private _cellKey(row: number, col: number): string {
    return `${row},${col}`;
  }

  _readValues(
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): CellValue[][] {
    const result: CellValue[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: CellValue[] = [];
      for (let c = startCol; c <= endCol; c++) {
        row.push(this._cells.get(this._cellKey(r, c)) ?? null);
      }
      result.push(row);
    }
    return result;
  }

  _readFormulas(
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): CellValue[][] {
    const result: CellValue[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: CellValue[] = [];
      for (let c = startCol; c <= endCol; c++) {
        row.push(this._formulas.get(this._cellKey(r, c)) ?? null);
      }
      result.push(row);
    }
    return result;
  }

  _readNumberFormats(
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): string[][] {
    const result: string[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: string[] = [];
      for (let c = startCol; c <= endCol; c++) {
        row.push(this._numberFormats.get(this._cellKey(r, c)) ?? "General");
      }
      result.push(row);
    }
    return result;
  }

  _writeValues(
    startRow: number,
    startCol: number,
    values: CellValue[][]
  ): void {
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const row = startRow + r;
        const col = startCol + c;
        this._cells.set(this._cellKey(row, col), values[r][c]);
        if (row > this._maxRow) this._maxRow = row;
        if (col > this._maxCol) this._maxCol = col;
      }
    }
  }

  _setNumberFormat(row: number, col: number, fmt: string): void {
    this._numberFormats.set(this._cellKey(row, col), fmt);
  }

  _getFormat(row: number, col: number): CellFormat {
    const key = this._cellKey(row, col);
    if (!this._formats.has(key)) {
      this._formats.set(key, {
        fillColor: "",
        fontColor: "",
        bold: false,
        italic: false,
        numberFormat: "General",
      });
    }
    return this._formats.get(key)!;
  }

  _setFormatProp(row: number, col: number, prop: string, value: unknown): void {
    const fmt = this._getFormat(row, col);
    (fmt as any)[prop] = value;
  }

  _getBorder(row: number, col: number, index: string): CellBorder {
    const key = `${row},${col}:${index}`;
    if (!this._borders.has(key)) {
      this._borders.set(key, { color: "", style: "", weight: "" });
    }
    return this._borders.get(key)!;
  }

  _setBorderProp(
    row: number,
    col: number,
    index: string,
    prop: string,
    value: string
  ): void {
    const border = this._getBorder(row, col, index);
    (border as any)[prop] = value;
  }

  // -- Public accessors for test assertions ---------------------------------

  /** Get cell value at a given address like "A1". */
  getCellValue(address: string): CellValue {
    const { row, col } = parseCellAddress(address);
    return this._cells.get(this._cellKey(row, col)) ?? null;
  }

  /** Get format for a cell at a given address like "A1". */
  getCellFormat(address: string): CellFormat {
    const { row, col } = parseCellAddress(address);
    return this._getFormat(row, col);
  }

  /** Get border for a cell at a given address like "A1" and border index. */
  getCellBorder(address: string, borderIndex: string): CellBorder {
    const { row, col } = parseCellAddress(address);
    return this._getBorder(row, col, borderIndex);
  }

  // -- Office.js API surface ------------------------------------------------

  /**
   * Compute used range from actual cell data bounds.
   * Returns a null-object SimRange if the sheet is empty.
   */
  getUsedRangeOrNullObject(): SimRange {
    if (this._maxRow < 0 || this._maxCol < 0) {
      const nullRange = new SimRange(this, 0, 0, 0, 0, true);
      // WHY: Auto-load common properties on null objects so tests don't need
      // explicit load calls just to check isNullObject.
      nullRange.load(["isNullObject", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
      nullRange._markSynced();
      return nullRange;
    }
    return new SimRange(this, 0, 0, this._maxRow, this._maxCol);
  }

  getRange(address: string): SimRange {
    const { startRow, startCol, endRow, endCol } = parseRangeAddress(address);
    return new SimRange(this, startRow, startCol, endRow, endCol);
  }

  getRangeByIndexes(
    row: number,
    col: number,
    rowCount: number,
    colCount: number
  ): SimRange {
    return new SimRange(
      this,
      row,
      col,
      row + rowCount - 1,
      col + colCount - 1
    );
  }

  getCell(row: number, col: number): SimRange {
    return new SimRange(this, row, col, row, col);
  }

  /**
   * Simplified search — finds all cells matching the query string.
   */
  findAllOrNullObject(
    query: string,
    _criteria: { completeMatch?: boolean; matchCase?: boolean }
  ): {
    isNullObject: boolean;
    areas: { items: SimRange[]; load: (p: string) => void };
    load: (p: string) => void;
    _markSynced: () => void;
  } {
    const matches: SimRange[] = [];
    const matchCase = _criteria.matchCase ?? false;
    const completeMatch = _criteria.completeMatch ?? false;
    const q = matchCase ? query : query.toLowerCase();

    for (const [key, val] of this._cells.entries()) {
      if (val == null) continue;
      let cellStr = String(val);
      if (!matchCase) cellStr = cellStr.toLowerCase();
      const found = completeMatch ? cellStr === q : cellStr.includes(q);
      if (found) {
        const [r, c] = key.split(",").map(Number);
        matches.push(new SimRange(this, r, c, r, c));
      }
    }

    const isNull = matches.length === 0;
    return {
      isNullObject: isNull,
      areas: {
        items: matches,
        load(_p: string) {
          matches.forEach((m) => {
            m.load(["address", "values"]);
            m._markSynced();
          });
        },
      },
      load(_p: string) {
        /* no-op */
      },
      _markSynced() {
        matches.forEach((m) => m._markSynced());
      },
    };
  }

  /** Get used range row and column bounds for test assertions. */
  get usedRangeRows(): number {
    return this._maxRow < 0 ? 0 : this._maxRow + 1;
  }

  get usedRangeCols(): number {
    return this._maxCol < 0 ? 0 : this._maxCol + 1;
  }

  load(_props: string | string[]): void {
    /* no-op — worksheet properties are always available */
  }
}

// ---------------------------------------------------------------------------
// SimWorksheetCollection — worksheets.getItem(), getActiveWorksheet(), etc.
// ---------------------------------------------------------------------------

class SimWorksheetCollection {
  private _workbook: SimWorkbook;

  onActivated = { add: vi.fn() };

  constructor(workbook: SimWorkbook) {
    this._workbook = workbook;
  }

  get items(): SimWorksheet[] {
    return Array.from(this._workbook._sheets.values());
  }

  load(_props: string | string[]): void {
    /* no-op */
  }

  getItem(name: string): SimWorksheet {
    const ws = this._workbook._sheets.get(name);
    if (!ws) throw new Error(`Worksheet '${name}' not found`);
    return ws;
  }

  getItemOrNullObject(name: string): SimWorksheet & { isNullObject: boolean } {
    const ws = this._workbook._sheets.get(name);
    if (ws) {
      return Object.assign(ws, {
        isNullObject: false,
        load(_p: string | string[]) {
          /* no-op */
        },
      });
    }
    // WHY: Return a proxy-like null object matching Office.js behavior
    const nullSheet = new SimWorksheet("null", name, -1) as SimWorksheet & {
      isNullObject: boolean;
    };
    (nullSheet as any).isNullObject = true;
    return nullSheet;
  }

  getActiveWorksheet(): SimWorksheet {
    return this._workbook.getActiveWorksheet();
  }

  add(name: string): SimWorksheet {
    const position = this._workbook._sheets.size;
    const ws = new SimWorksheet(
      `sheet-${Date.now()}-${position}`,
      name,
      position
    );
    this._workbook._sheets.set(name, ws);
    // WHY: Newly added worksheets must be wrapped by the active context
    // so their range methods register for sync flushing.
    if (_activeContext) {
      _activeContext._wrapWorksheet(ws);
    }
    return ws;
  }
}

// ---------------------------------------------------------------------------
// SimWorkbook — root state container
// ---------------------------------------------------------------------------

export class SimWorkbook {
  name: string;
  _sheets: Map<string, SimWorksheet> = new Map();
  private _activeSheetName: string;
  private _selectedRangeAddress: string;
  worksheets: SimWorksheetCollection;

  // WHY: We need load() to be available on the workbook proxy for
  // workbook.load("name") calls.
  private _loaded: Set<string> = new Set();

  constructor(
    name: string,
    activeSheet: string,
    selectedRange = "A1"
  ) {
    this.name = name;
    this._activeSheetName = activeSheet;
    this._selectedRangeAddress = selectedRange;
    this.worksheets = new SimWorksheetCollection(this);
  }

  load(props: string | string[]): void {
    const names = Array.isArray(props) ? props : [props];
    for (const n of names) this._loaded.add(n);
  }

  getActiveWorksheet(): SimWorksheet {
    const ws = this._sheets.get(this._activeSheetName);
    if (!ws) {
      throw new Error(
        `Active worksheet '${this._activeSheetName}' not found`
      );
    }
    return ws;
  }

  getSelectedRange(): SimRange {
    const ws = this.getActiveWorksheet();
    const { startRow, startCol, endRow, endCol } = parseRangeAddress(
      this._selectedRangeAddress
    );
    const range = new SimRange(ws, startRow, startCol, endRow, endCol);
    range.load(["address", "values"]);
    range._markSynced();
    return range;
  }

  getSelectedRanges(): {
    address: string;
    areas: { items: SimRange[] };
    load: (p: string) => void;
  } {
    const ws = this.getActiveWorksheet();
    const addresses = this._selectedRangeAddress.split(",");
    const ranges = addresses.map((addr) => {
      const { startRow, startCol, endRow, endCol } = parseRangeAddress(
        addr.trim()
      );
      const r = new SimRange(ws, startRow, startCol, endRow, endCol);
      // WHY: Pre-load and sync so the range is immediately readable,
      // matching common usage patterns.
      r.load(["address", "values"]);
      r._markSynced();
      (r as any).worksheet = ws;
      return r;
    });

    const fullAddress = addresses
      .map((a) => `${ws.name}!${a.trim()}`)
      .join(",");

    return {
      address: fullAddress,
      areas: { items: ranges },
      load(_p: string) {
        /* no-op */
      },
    };
  }

  getActiveCell(): SimRange {
    const ws = this.getActiveWorksheet();
    const { startRow, startCol } = parseRangeAddress(
      this._selectedRangeAddress.split(",")[0]
    );
    const r = new SimRange(ws, startRow, startCol, startRow, startCol);
    r.load(["address"]);
    r._markSynced();
    return r;
  }

  /** Add a worksheet with pre-populated cell data. */
  addSheet(
    name: string,
    cells: Record<string, CellValue> = {},
    position = -1
  ): SimWorksheet {
    const pos = position >= 0 ? position : this._sheets.size;
    const ws = new SimWorksheet(`sheet-${pos}`, name, pos);
    for (const [addr, val] of Object.entries(cells)) {
      const { row, col } = parseCellAddress(addr);
      ws._writeValues(row, col, [[val]]);
    }
    this._sheets.set(name, ws);
    return ws;
  }

  /** Get a worksheet by name for test assertions. */
  getSheet(name: string): SimWorksheet {
    const ws = this._sheets.get(name);
    if (!ws) throw new Error(`Sheet '${name}' not found`);
    return ws;
  }
}

// ---------------------------------------------------------------------------
// SimRequestContext — manages sync lifecycle
// ---------------------------------------------------------------------------

/**
 * Simulates Excel.RequestContext. Tracks all created SimRange proxies and
 * flushes their pending writes on sync().
 */
export class SimRequestContext {
  workbook: SimWorkbook;
  _ranges: SimRange[] = [];

  constructor(workbook: SimWorkbook) {
    this.workbook = workbook;

    // WHY: Set this context as the active context so SimRange methods
    // (getCell, getResizedRange, getOffsetRange) can register new ranges.
    _activeContext = this;

    // WHY: Intercept range creation so we can track all proxies for sync().
    // We wrap the worksheet methods to register every range they create.
    const self = this;
    for (const ws of workbook._sheets.values()) {
      self._wrapWorksheet(ws);
    }
  }

  _wrapWorksheet(ws: SimWorksheet): void {
    const self = this;
    const origGetRange = ws.getRange.bind(ws);
    const origGetRangeByIndexes = ws.getRangeByIndexes.bind(ws);
    const origGetCell = ws.getCell.bind(ws);
    const origGetUsedRange = ws.getUsedRangeOrNullObject.bind(ws);

    ws.getRange = (address: string) => {
      const r = origGetRange(address);
      self._ranges.push(r);
      return r;
    };
    ws.getRangeByIndexes = (
      row: number,
      col: number,
      rowCount: number,
      colCount: number
    ) => {
      const r = origGetRangeByIndexes(row, col, rowCount, colCount);
      self._ranges.push(r);
      return r;
    };
    ws.getCell = (row: number, col: number) => {
      const r = origGetCell(row, col);
      self._ranges.push(r);
      return r;
    };
    ws.getUsedRangeOrNullObject = () => {
      const r = origGetUsedRange();
      self._ranges.push(r);
      return r;
    };

    // WHY: Keep references to the backing arrays before overwriting with
    // wrapper objects. The wrappers provide the Office.js collection API
    // (items, load, add) while the backing arrays hold the actual data.
    const chartsArray = ws._chartsArray;
    const pivotsArray = ws._pivotsArray;
    const tablesArray = ws._tablesArray;

    (ws as any).charts = {
      get items() {
        return chartsArray;
      },
      load(_p: string) {
        /* no-op */
      },
      add: (
        chartType: string,
        sourceRange: SimRange,
        seriesBy: string
      ): SimChart => {
        const chart = new SimChart(chartType, sourceRange, seriesBy);
        chartsArray.push(chart);
        return chart;
      },
    };

    (ws as any).pivotTables = {
      get items() {
        return pivotsArray;
      },
      load(_p: string) {
        /* no-op */
      },
      add: (
        name: string,
        sourceRange: SimRange,
        destRange: SimRange
      ): SimPivotTable => {
        // WHY: Extract field names from the first row of source data so
        // hierarchies.getItem() works realistically.
        const vals = sourceRange.worksheet._readValues(
          (sourceRange as any)._startRow,
          (sourceRange as any)._startCol,
          (sourceRange as any)._startRow,
          (sourceRange as any)._endCol
        );
        const fieldNames = vals[0]?.map((v) => String(v ?? "")) ?? [];
        const pt = new SimPivotTable(name, sourceRange, destRange, fieldNames);

        // Attach hierarchy adder methods
        (pt as any).rowHierarchies = pt._makeHierarchyAdder(pt.rowHierarchies);
        (pt as any).columnHierarchies = pt._makeHierarchyAdder(
          pt.columnHierarchies
        );
        (pt as any).dataHierarchies = pt._makeHierarchyAdder(
          pt.dataHierarchies
        );
        (pt as any).filterHierarchies = pt._makeHierarchyAdder(
          pt.filterHierarchies
        );

        pivotsArray.push(pt);
        return pt;
      },
    };

    // Tables collection stub
    (ws as any).tables = {
      get items() {
        return ws.tables;
      },
      load(_p: string) {
        /* no-op */
      },
    };
  }

  /**
   * Flush all pending writes and mark all proxies as synced.
   * Matches Excel.RequestContext.sync() behavior.
   */
  async sync(): Promise<void> {
    // Flush writes from all tracked ranges
    for (const range of this._ranges) {
      range._flush();
      range._markSynced();
    }

    // WHY: Also wrap any newly created worksheets (e.g., from worksheets.add())
    // that weren't wrapped during constructor time.
    for (const ws of this.workbook._sheets.values()) {
      if (!(ws as any)._wrapped) {
        this._wrapWorksheet(ws);
        (ws as any)._wrapped = true;
      }
    }

    // WHY: Keep context active for subsequent sync() calls within the same
    // Excel.run callback (real Office.js supports multiple syncs).
  }
}

// ---------------------------------------------------------------------------
// Factory: build a SimWorkbook from a JSON fixture
// ---------------------------------------------------------------------------

export interface WorkbookFixtureDef {
  name?: string;
  sheets: Array<{
    name: string;
    cells?: Record<string, CellValue>;
  }>;
  activeSheet: string;
  selectedRange?: string;
}

/**
 * Create a SimWorkbook from a test fixture definition.
 *
 * @param def - JSON fixture defining initial workbook state.
 * @returns A fully populated SimWorkbook ready for testing.
 */
export function buildSimWorkbook(def: WorkbookFixtureDef): SimWorkbook {
  const wb = new SimWorkbook(
    def.name ?? "TestWorkbook.xlsx",
    def.activeSheet,
    def.selectedRange ?? "A1"
  );
  for (const sheet of def.sheets) {
    wb.addSheet(sheet.name, sheet.cells ?? {});
  }
  return wb;
}

// ---------------------------------------------------------------------------
// createSimExcelRun — drop-in replacement for Excel.run
// ---------------------------------------------------------------------------

/**
 * Create a simulated Excel.run function bound to a SimWorkbook.
 *
 * WHY: This is the bridge between the real excel.ts code (which calls
 * Excel.run) and the simulator. Tests replace the global Excel.run with
 * this, so the real mutation/read functions execute against sim state.
 *
 * @param workbook - The SimWorkbook to operate on.
 * @returns An async function with the same signature as Excel.run.
 */
export function createSimExcelRun(workbook: SimWorkbook) {
  return async (
    callback: (context: SimRequestContext) => Promise<any>
  ): Promise<any> => {
    const context = new SimRequestContext(workbook);
    try {
      const result = await callback(context as any);
      await context.sync();
      return result;
    } finally {
      _activeContext = null;
    }
  };
}

// ---------------------------------------------------------------------------
// installSimulator — sets up globals for a test
// ---------------------------------------------------------------------------

/**
 * Install the simulator as the global Excel and Office objects.
 * Returns a cleanup function that restores the original globals.
 *
 * @param workbook - The SimWorkbook to use for Excel.run calls.
 * @returns Object with the workbook and a restore() function.
 */
export function installSimulator(workbook: SimWorkbook) {
  const origExcel = (globalThis as any).Excel;
  const origOffice = (globalThis as any).Office;

  (globalThis as any).Excel = {
    ...origExcel,
    run: createSimExcelRun(workbook),
  };

  (globalThis as any).Office = {
    context: {
      host: "Excel",
      diagnostics: { version: "sim" },
    },
    HostType: { Excel: "Excel", Word: "Word" },
    onReady: vi.fn((cb?: () => void) => {
      cb?.();
      return Promise.resolve();
    }),
  };

  return {
    workbook,
    restore() {
      (globalThis as any).Excel = origExcel;
      (globalThis as any).Office = origOffice;
    },
  };
}
