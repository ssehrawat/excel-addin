"""Lightweight in-memory workbook fixture for E2E pipeline testing.

Provides a Python-side workbook data holder that can respond to the same
tool calls the frontend would execute via Office.js. Used by the headless
frontend test driver to simulate tool call round-trips without Excel.

Key exports:
- WorkbookFixture: builds from scenario JSON, answers tool calls
- build_selection: converts fixture workbook + selectedRange into CellSelection
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional


def _col_to_index(letters: str) -> int:
    """Convert column letters (A, B, ..., AA, AB) to 0-based index."""
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _parse_cell(address: str) -> tuple[int, int]:
    """Parse 'A1' into (row=0, col=0)."""
    m = re.match(r"([A-Z]+)(\d+)", address.upper())
    if not m:
        raise ValueError(f"Invalid cell address: {address}")
    return int(m.group(2)) - 1, _col_to_index(m.group(1))


def _parse_range(address: str) -> tuple[int, int, int, int]:
    """Parse 'A1:C3' into (startRow, startCol, endRow, endCol)."""
    if "!" in address:
        address = address.split("!")[-1]
    parts = address.split(":")
    r1, c1 = _parse_cell(parts[0])
    if len(parts) == 1:
        return r1, c1, r1, c1
    r2, c2 = _parse_cell(parts[1])
    return r1, c1, r2, c2


class WorkbookFixture:
    """In-memory workbook that answers tool calls for E2E pipeline tests.

    Args:
        sheets: Mapping of sheet name -> cell address -> value.
                 E.g. {"Sheet1": {"A1": "Name", "A2": "Alice"}}
        active_sheet: Name of the active sheet.
        selected_range: Selected range address (e.g. "Sheet1!A1:B3").
    """

    def __init__(
        self,
        sheets: Dict[str, Dict[str, Any]],
        active_sheet: str = "Sheet1",
        selected_range: str = "A1",
    ):
        self.sheets = sheets
        self.active_sheet = active_sheet
        self.selected_range = selected_range

    def _get_cell(self, sheet_name: str, row: int, col: int) -> Any:
        """Read a single cell value by row/col index."""
        from string import ascii_uppercase

        # WHY: Convert row/col back to address for dict lookup since fixtures
        # store cells by address string.
        def _idx_to_col(idx: int) -> str:
            result = ""
            n = idx
            while n >= 0:
                result = chr(65 + n % 26) + result
                n = n // 26 - 1
            return result

        addr = f"{_idx_to_col(col)}{row + 1}"
        sheet = self.sheets.get(sheet_name, {})
        return sheet.get(addr)

    def _get_range_values(
        self, sheet_name: str, range_addr: str
    ) -> List[List[Any]]:
        """Read a 2D array of values for a range address."""
        r1, c1, r2, c2 = _parse_range(range_addr)
        result = []
        for r in range(r1, r2 + 1):
            row = []
            for c in range(c1, c2 + 1):
                row.append(self._get_cell(sheet_name, r, c))
            result.append(row)
        return result

    def _range_to_csv(self, sheet_name: str, range_addr: str,
                      max_rows: Optional[int] = None,
                      offset: int = 0) -> str:
        """Convert a range to CSV string."""
        values = self._get_range_values(sheet_name, range_addr)
        rows = values[offset:]
        if max_rows is not None:
            rows = rows[:max_rows]
        lines = []
        for row in rows:
            cells = []
            for cell in row:
                s = "" if cell is None else str(cell)
                if "," in s or "\n" in s:
                    s = f'"{s}"'
                cells.append(s)
            lines.append(",".join(cells))
        return "\n".join(lines)

    def execute_tool(self, tool: str, args: Dict[str, Any]) -> Dict[str, Any]:
        """Simulate frontend tool execution.

        Args:
            tool: Tool name (e.g. "get_xl_range_as_csv").
            args: Tool arguments from the backend.

        Returns:
            Dict with id, tool, and result (matching WorkbookToolResult shape).
        """
        tool_id = args.get("id", f"tr-{tool}")

        if tool == "get_xl_range_as_csv":
            sheet = args.get("sheetName", self.active_sheet)
            rng = args.get("range", "A1")
            max_rows = args.get("maxRows")
            offset = args.get("offset", 0)
            result = self._range_to_csv(sheet, rng, max_rows, offset)
            return {"id": tool_id, "tool": tool, "result": result}

        elif tool == "get_xl_cell_ranges":
            sheet = args.get("sheetName", self.active_sheet)
            ranges = args.get("ranges", [])
            results = []
            for addr in ranges:
                values = self._get_range_values(sheet, addr)
                results.append({"address": f"{sheet}!{addr}", "values": values})
            return {"id": tool_id, "tool": tool, "result": results}

        elif tool == "xl_search_data":
            query = args.get("query", "")
            sheet_name = args.get("sheetName")
            matches = []
            sheets_to_search = (
                [sheet_name] if sheet_name else list(self.sheets.keys())
            )
            for sn in sheets_to_search:
                for addr, val in self.sheets.get(sn, {}).items():
                    if query.lower() in str(val).lower():
                        matches.append(
                            {"address": f"{sn}!{addr}", "worksheet": sn,
                             "values": [[val]]}
                        )
            return {"id": tool_id, "tool": tool, "result": matches}

        elif tool == "get_all_xl_objects":
            return {
                "id": tool_id,
                "tool": tool,
                "result": {"charts": [], "tables": [], "pivotTables": []},
            }

        elif tool == "execute_xl_office_js":
            return {
                "id": tool_id,
                "tool": tool,
                "result": {"success": True, "message": "Mock execution."},
            }

        return {"id": tool_id, "tool": tool, "result": None,
                "error": f"Unknown tool: {tool}"}


def _compute_sheet_bounds(cells: Dict[str, Any]) -> tuple[int, int, List[str]]:
    """Compute max rows, max cols, and column headers from a cell dict.

    Returns:
        (max_rows, max_cols, column_headers)
    """
    if not cells:
        return 0, 0, []
    max_row = 0
    max_col = 0
    for addr in cells:
        r, c = _parse_cell(addr)
        if r + 1 > max_row:
            max_row = r + 1
        if c + 1 > max_col:
            max_col = c + 1

    # WHY: Extract headers from row 1 to match what getWorkbookMetadata() does
    def _idx_to_col(idx: int) -> str:
        result = ""
        n = idx
        while n >= 0:
            result = chr(65 + n % 26) + result
            n = n // 26 - 1
        return result

    headers = []
    for c in range(max_col):
        addr = f"{_idx_to_col(c)}1"
        val = cells.get(addr)
        headers.append(str(val) if val is not None else "")
    return max_row, max_col, headers


def build_full_context(
    fixture: WorkbookFixture, max_preview_rows: int = 50
) -> Dict[str, Any]:
    """Build the complete workbook context that handleSend() collects.

    Generates workbookMetadata, userContext, activeSheetPreview, and selection
    from the fixture data — mirroring what the frontend reads from Office.js.

    Args:
        fixture: The WorkbookFixture to extract context from.
        max_preview_rows: Max rows for the active sheet preview CSV.

    Returns:
        Dict with keys: workbookMetadata, userContext, activeSheetPreview,
        selection — ready to merge into a ChatRequest payload.
    """
    # Build workbookMetadata
    sheets_metadata = []
    for i, (name, cells) in enumerate(fixture.sheets.items()):
        max_rows, max_cols, headers = _compute_sheet_bounds(cells)
        sheets_metadata.append({
            "id": f"sheet-{i}",
            "name": name,
            "index": i,
            "maxRows": max_rows,
            "maxColumns": max_cols,
            "columnHeaders": headers if headers else None,
        })

    workbook_metadata = {
        "success": True,
        "fileName": "Financial_Market_Demo.xlsx",
        "sheetsMetadata": sheets_metadata,
        "totalSheets": len(sheets_metadata),
    }

    # Build userContext
    user_context = {
        "currentActiveSheetName": fixture.active_sheet,
        "selectedRanges": fixture.selected_range,
    }

    # Build activeSheetPreview (CSV of first N rows of active sheet)
    active_cells = fixture.sheets.get(fixture.active_sheet, {})
    max_rows, max_cols, _ = _compute_sheet_bounds(active_cells)
    preview_rows = min(max_preview_rows, max_rows)
    preview = None
    if preview_rows > 0 and max_cols > 0:
        # WHY: Match the format in getLightweightSheetPreview() —
        # column letter header row, then data rows as CSV
        def _idx_to_col(idx: int) -> str:
            result = ""
            n = idx
            while n >= 0:
                result = chr(65 + n % 26) + result
                n = n // 26 - 1
            return result

        col_letters = [f"[{_idx_to_col(c)}]" for c in range(max_cols)]
        lines = [",".join(col_letters)]
        for r in range(preview_rows):
            row_cells = []
            for c in range(max_cols):
                val = fixture._get_cell(fixture.active_sheet, r, c)
                s = "" if val is None else str(val)
                if "," in s or "\n" in s:
                    s = f'"{s}"'
                row_cells.append(s)
            lines.append(",".join(row_cells))
        preview = "\n".join(lines)

    # Build selection
    selection = build_selection(fixture)

    return {
        "workbookMetadata": workbook_metadata,
        "userContext": user_context,
        "activeSheetPreview": preview,
        "selection": selection,
    }


def build_selection(fixture: WorkbookFixture) -> List[Dict[str, Any]]:
    """Build CellSelection list from fixture's selected range.

    Returns:
        List of dicts matching the CellSelection schema.
    """
    sel_range = fixture.selected_range
    if not sel_range:
        return []

    # Parse sheet name from "Sheet1!A1:B3" format
    if "!" in sel_range:
        parts = sel_range.split("!")
        sheet_name = parts[0].strip("'")
        range_addr = parts[1]
    else:
        sheet_name = fixture.active_sheet
        range_addr = sel_range

    values = fixture._get_range_values(sheet_name, range_addr)
    return [
        {
            "address": f"{sheet_name}!{range_addr}",
            "values": values,
            "worksheet": sheet_name,
        }
    ]
