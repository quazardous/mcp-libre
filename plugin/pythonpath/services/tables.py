"""
TableService — Writer table operations.
"""

import logging
from typing import Any, Dict

logger = logging.getLogger(__name__)


class TableService:
    """Writer table operations via UNO."""

    def __init__(self, registry):
        self._registry = registry
        self._base = registry.base

    def list_tables(self, file_path: str = None) -> Dict[str, Any]:
        """List all text tables in the document."""
        try:
            doc = self._base.resolve_document(file_path)
            if not hasattr(doc, 'getTextTables'):
                return {"success": False,
                        "error": "Document does not support text tables"}

            tables_sup = doc.getTextTables()
            tables = []
            for name in tables_sup.getElementNames():
                table = tables_sup.getByName(name)
                rows = table.getRows().getCount()
                cols = table.getColumns().getCount()
                tables.append({
                    "name": name,
                    "rows": rows,
                    "cols": cols,
                })
            return {"success": True, "tables": tables,
                    "count": len(tables)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_table(self, table_name: str,
                   file_path: str = None) -> Dict[str, Any]:
        """Read all cell contents from a Writer table."""
        try:
            doc = self._base.resolve_document(file_path)
            tables_sup = doc.getTextTables()

            if not tables_sup.hasByName(table_name):
                return {"success": False,
                        "error": f"Table '{table_name}' not found",
                        "available": list(tables_sup.getElementNames())}

            table = tables_sup.getByName(table_name)
            rows = table.getRows().getCount()
            cols = table.getColumns().getCount()

            data = []
            for r in range(rows):
                row_data = []
                for c in range(cols):
                    col_letter = (chr(ord('A') + c) if c < 26
                                  else f"A{chr(ord('A') + c - 26)}")
                    cell_name = f"{col_letter}{r + 1}"
                    try:
                        cell = table.getCellByName(cell_name)
                        row_data.append(cell.getString())
                    except Exception:
                        row_data.append("")
                data.append(row_data)

            return {"success": True, "table_name": table_name,
                    "rows": rows, "cols": cols, "data": data}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def write_table_cell(self, table_name: str, cell: str, value: str,
                         file_path: str = None) -> Dict[str, Any]:
        """Write to a cell in a Writer table."""
        try:
            doc = self._base.resolve_document(file_path)
            tables_sup = doc.getTextTables()

            if not tables_sup.hasByName(table_name):
                return {"success": False,
                        "error": f"Table '{table_name}' not found"}

            table = tables_sup.getByName(table_name)
            cell_obj = table.getCellByName(cell)
            if cell_obj is None:
                return {"success": False,
                        "error": f"Cell '{cell}' not found in {table_name}"}

            try:
                cell_obj.setValue(float(value))
            except (ValueError, TypeError):
                cell_obj.setString(value)

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True, "table": table_name,
                    "cell": cell, "value": value}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def create_table(self, rows: int, cols: int,
                     paragraph_index: int = None,
                     locator: str = None,
                     file_path: str = None) -> Dict[str, Any]:
        """Create a new table at a paragraph position."""
        try:
            doc = self._base.resolve_document(file_path)

            if locator is not None and paragraph_index is None:
                resolved = self._base.resolve_locator(doc, locator)
                paragraph_index = resolved.get("para_index")
            if paragraph_index is None:
                return {"success": False,
                        "error": "Provide locator or paragraph_index"}

            target, _ = self._registry.writer.find_paragraph_element(
                doc, paragraph_index)
            if target is None:
                return {"success": False,
                        "error": f"Paragraph {paragraph_index} not found"}

            table = doc.createInstance("com.sun.star.text.TextTable")
            table.initialize(rows, cols)
            doc_text = doc.getText()
            cursor = doc_text.createTextCursorByRange(target.getEnd())
            doc_text.insertTextContent(cursor, table, False)

            table_name = table.getName()

            if doc.hasLocation():
                self._base.store_doc(doc)

            return {"success": True, "table_name": table_name,
                    "rows": rows, "cols": cols}
        except Exception as e:
            return {"success": False, "error": str(e)}
