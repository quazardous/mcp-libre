"""
CalcService — Calc spreadsheet operations.
"""

import logging
from typing import Any, Dict, Optional, Tuple

logger = logging.getLogger(__name__)


class CalcService:
    """Spreadsheet cell, sheet, and range operations via UNO."""

    def __init__(self, registry):
        self._registry = registry
        self._base = registry.base

    # ==================================================================
    # Helpers
    # ==================================================================

    @staticmethod
    def _parse_cell_address(addr: str) -> Tuple[Optional[str], int, int]:
        """Parse 'A1' or 'Sheet1.A1' -> (sheet_name|None, col, row)."""
        sheet_name = None
        if '.' in addr:
            sheet_name, addr = addr.split('.', 1)
        col_str = ""
        row_str = ""
        for ch in addr:
            if ch.isalpha():
                col_str += ch
            else:
                row_str += ch
        col = 0
        for ch in col_str.upper():
            col = col * 26 + (ord(ch) - ord('A') + 1)
        col -= 1  # 0-based
        row = int(row_str) - 1  # 0-based
        return sheet_name, col, row

    def _get_sheet(self, doc, sheet_name: str = None):
        """Get a sheet by name or the active sheet."""
        sheets = doc.getSheets()
        if sheet_name:
            if not sheets.hasByName(sheet_name):
                raise ValueError(
                    f"Sheet '{sheet_name}' not found. "
                    f"Available: {list(sheets.getElementNames())}")
            return sheets.getByName(sheet_name)
        controller = doc.getCurrentController()
        if controller:
            return controller.getActiveSheet()
        return sheets.getByIndex(0)

    # ==================================================================
    # Public API
    # ==================================================================

    def read_cells(self, range_str: str,
                   sheet_name: str = None,
                   file_path: str = None) -> Dict[str, Any]:
        """Read cell values from a range (e.g. 'A1:D10')."""
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_calc(doc):
                return {"success": False, "error": "Not a Calc document"}

            if ':' in range_str:
                start_addr, end_addr = range_str.split(':', 1)
            else:
                start_addr = end_addr = range_str

            s_sheet, s_col, s_row = self._parse_cell_address(start_addr)
            e_sheet, e_col, e_row = self._parse_cell_address(end_addr)
            resolved_sheet = sheet_name or s_sheet or e_sheet
            sheet = self._get_sheet(doc, resolved_sheet)

            cell_range = sheet.getCellRangeByPosition(
                s_col, s_row, e_col, e_row)
            data = cell_range.getDataArray()
            rows = [list(row) for row in data]

            return {
                "success": True,
                "range": range_str,
                "sheet": sheet.getName(),
                "data": rows,
                "row_count": len(rows),
                "col_count": len(rows[0]) if rows else 0,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def write_cell(self, cell: str, value: str,
                   sheet_name: str = None,
                   file_path: str = None) -> Dict[str, Any]:
        """Write a value to a cell (e.g. 'B3' or 'Sheet1.B3')."""
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_calc(doc):
                return {"success": False, "error": "Not a Calc document"}

            parsed_sheet, col, row = self._parse_cell_address(cell)
            sheet = self._get_sheet(doc, sheet_name or parsed_sheet)
            cell_obj = sheet.getCellByPosition(col, row)

            try:
                num_val = float(value)
                cell_obj.setValue(num_val)
            except ValueError:
                cell_obj.setString(value)

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {
                "success": True,
                "cell": cell,
                "sheet": sheet.getName(),
                "value": value,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_sheets(self, file_path: str = None) -> Dict[str, Any]:
        """List all sheets with names and basic info."""
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_calc(doc):
                return {"success": False, "error": "Not a Calc document"}

            sheets_obj = doc.getSheets()
            count = sheets_obj.getCount()
            sheets = []
            for i in range(count):
                sheet = sheets_obj.getByIndex(i)
                sheets.append({
                    "index": i,
                    "name": sheet.getName(),
                    "is_visible": (sheet.IsVisible
                                   if hasattr(sheet, 'IsVisible')
                                   else True),
                })
            return {"success": True, "sheets": sheets, "count": count}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_sheet_info(self, sheet_name: str = None,
                       file_path: str = None) -> Dict[str, Any]:
        """Get info about a sheet: used range, row/col count."""
        try:
            doc = self._base.resolve_document(file_path)
            if not self._base.is_calc(doc):
                return {"success": False, "error": "Not a Calc document"}

            sheet = self._get_sheet(doc, sheet_name)
            cursor = sheet.createCursor()
            cursor.gotoStartOfUsedArea(False)
            cursor.gotoEndOfUsedArea(True)

            range_addr = cursor.getRangeAddress()
            used_rows = range_addr.EndRow - range_addr.StartRow + 1
            used_cols = range_addr.EndColumn - range_addr.StartColumn + 1

            return {
                "success": True,
                "sheet_name": sheet.getName(),
                "used_rows": used_rows,
                "used_cols": used_cols,
                "start_row": range_addr.StartRow,
                "start_col": range_addr.StartColumn,
                "end_row": range_addr.EndRow,
                "end_col": range_addr.EndColumn,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
