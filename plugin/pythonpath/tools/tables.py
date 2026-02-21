"""Writer table tools — list, read, write, create tables."""

from .base import McpTool


class ListDocumentTables(McpTool):
    name = "list_tables"
    description = (
        "List all text tables in a Writer document. "
        "Returns table name, row count, and column count for each table."
    )
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
    }

    def execute(self, file_path=None, **_):
        return self.services.tables.list_tables(file_path)


class ReadDocumentTable(McpTool):
    name = "read_table"
    description = (
        "Read all cell contents from a Writer table. "
        "Returns a 2D array of cell values."
    )
    parameters = {
        "type": "object",
        "properties": {
            "table_name": {
                "type": "string",
                "description": "Name of the table (use list_document_tables to find)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["table_name"],
    }

    def execute(self, table_name, file_path=None, **_):
        return self.services.tables.read_table(table_name, file_path)


class WriteDocumentTableCell(McpTool):
    name = "write_table_cell"
    description = (
        "Write to a cell in a Writer table. "
        "Numbers are auto-detected. Use cell addresses like A1, B3, etc."
    )
    parameters = {
        "type": "object",
        "properties": {
            "table_name": {
                "type": "string",
                "description": "Name of the table",
            },
            "cell": {
                "type": "string",
                "description": "Cell address (e.g. 'A1', 'B3')",
            },
            "value": {
                "type": "string",
                "description": "Value to write",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["table_name", "cell", "value"],
    }

    def execute(self, table_name, cell, value, file_path=None, **_):
        return self.services.tables.write_table_cell(
            table_name, cell, value, file_path)


class CreateDocumentTable(McpTool):
    name = "create_table"
    description = (
        "Create a new table at a paragraph position. "
        "The table is inserted after the target paragraph."
    )
    parameters = {
        "type": "object",
        "properties": {
            "rows": {
                "type": "integer",
                "description": "Number of rows",
            },
            "cols": {
                "type": "integer",
                "description": "Number of columns",
            },
            "locator": {
                "type": "string",
                "description": "Unified locator for insertion point",
            },
            "paragraph_index": {
                "type": "integer",
                "description": "Paragraph index (legacy)",
            },
            "file_path": {
                "type": "string",
                "description": "Absolute path to the document (optional)",
            },
        },
        "required": ["rows", "cols"],
    }

    def execute(self, rows, cols, paragraph_index=None, locator=None,
                file_path=None, **_):
        return self.services.tables.create_table(
            rows, cols, paragraph_index, locator, file_path)
