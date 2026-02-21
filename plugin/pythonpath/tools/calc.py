"""Calc tools — read/write cells, list sheets."""

from .base import McpTool


class ReadSpreadsheetCells(McpTool):
    name = "read_cells"
    description = (
        "Read cell values from a Calc spreadsheet. "
        "Prefix range with sheet name and dot for a specific sheet "
        "(e.g. 'Sheet1.A1:D10')."
    )
    parameters = {
        "type": "object",
        "properties": {
            "range_str": {
                "type": "string",
                "description": "Cell range (e.g. 'A1:D10', 'B3'). "
                               "Prefix with sheet name for a specific sheet "
                               "(e.g. 'Sheet1.A1:D10').",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the spreadsheet (optional)",
            },
        },
        "required": ["range_str"],
    }

    def execute(self, range_str, file_path=None, **_):
        return self.services.calc.read_cells(range_str, file_path)


class WriteSpreadsheetCell(McpTool):
    name = "write_cell"
    description = (
        "Write a value to a Calc spreadsheet cell. "
        "Numbers are auto-detected."
    )
    parameters = {
        "type": "object",
        "properties": {
            "cell": {
                "type": "string",
                "description": "Cell address (e.g. 'B3'). "
                               "Prefix with sheet name for a specific sheet "
                               "(e.g. 'Sheet1.B3').",
            },
            "value": {
                "type": "string",
                "description": "Value to write (numbers are auto-detected)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the spreadsheet (optional)",
            },
        },
        "required": ["cell", "value"],
    }

    def execute(self, cell, value, file_path=None, **_):
        return self.services.calc.write_cell(cell, value, file_path)


class ListSpreadsheetSheets(McpTool):
    name = "list_sheets"
    description = "List all sheets in a Calc spreadsheet."
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute path to the spreadsheet (optional)",
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.calc.list_sheets(file_path)


class GetSpreadsheetSheetInfo(McpTool):
    name = "get_sheet_info"
    description = "Get info about a spreadsheet sheet (used range, dimensions)."
    parameters = {
        "type": "object",
        "properties": {
            "sheet_name": {
                "type": "string",
                "description": "Sheet name (optional, defaults to active sheet)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the spreadsheet (optional)",
            },
        },
    }

    def execute(self, sheet_name=None, file_path=None, **_):
        return self.services.calc.get_sheet_info(sheet_name, file_path)
