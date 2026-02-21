"""Document lifecycle tools — open, close, create, save, list."""

from .base import McpTool


class CreateDocument(McpTool):
    name = "create_document"
    description = "Create a new LibreOffice document."
    parameters = {
        "type": "object",
        "properties": {
            "doc_type": {
                "type": "string",
                "enum": ["writer", "calc", "impress", "draw"],
                "description": "Type of document to create (default: writer)",
            },
            "content": {
                "type": "string",
                "description": "Initial content for the document (for writer docs)",
            },
        },
    }

    def execute(self, doc_type="writer", content="", **_):
        doc = self.services.base.create_document(doc_type)
        if doc is None:
            return {"success": False,
                    "error": f"Failed to create {doc_type} document"}
        result = {"success": True,
                  "message": f"Created new {doc_type} document"}
        if content and doc_type == "writer":
            text = doc.getText()
            cursor = text.createTextCursor()
            text.insertString(cursor, content, False)
            result["message"] += " with initial content"
        return result


class OpenDocumentInLibreOffice(McpTool):
    name = "open_document"
    description = "Open a document in LibreOffice GUI for live viewing."
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Path to the document to open",
            },
            "force": {
                "type": "boolean",
                "description": "Force open even if a document with the same "
                               "name is already open (default: False)",
            },
        },
        "required": ["file_path"],
    }

    def execute(self, file_path, force=False, **_):
        result = self.services.base.open_document(file_path, force=force)
        return {k: v for k, v in result.items() if k != "doc"}


class CloseDocument(McpTool):
    name = "close_document"
    description = "Close a document by file path (no save)."
    parameters = {
        "type": "object",
        "properties": {
            "file_path": {
                "type": "string",
                "description": "Absolute file path",
            },
        },
        "required": ["file_path"],
    }

    def execute(self, file_path, **_):
        return self.services.base.close_document(file_path)


class ListOpenDocumentsLive(McpTool):
    name = "list_open_documents"
    description = "List all currently open documents in LibreOffice."
    parameters = {"type": "object", "properties": {}}

    def execute(self, **_):
        try:
            desktop = self.services.desktop
            documents = []
            frames = desktop.getFrames()
            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                controller = frame.getController()
                if controller:
                    doc = controller.getModel()
                    if doc and hasattr(doc, "getURL"):
                        import urllib.parse
                        url = doc.getURL()
                        doc_type = self.services.base.get_document_type(doc)
                        entry = {"type": doc_type}
                        if url:
                            entry["url"] = url
                            try:
                                entry["path"] = urllib.parse.unquote(
                                    url.replace("file:///", "")
                                ).replace("/", "\\")
                            except Exception:
                                pass
                        documents.append(entry)
            return {"success": True, "documents": documents,
                    "count": len(documents)}
        except Exception as e:
            return {"success": False, "error": str(e)}


class SaveActiveDocument(McpTool):
    name = "save_document"
    description = "Save the currently active document to its current location."
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
        try:
            doc = self.services.base.resolve_document(file_path)
            self.services.base.store_doc(doc)
            return {"success": True, "message": "Document saved"}
        except Exception as e:
            return {"success": False, "error": str(e)}


class SaveDocumentCopy(McpTool):
    name = "save_document_as"
    description = "Save/duplicate a document under a new name."
    parameters = {
        "type": "object",
        "properties": {
            "target_path": {
                "type": "string",
                "description": "New file path to save the copy to",
            },
            "file_path": {
                "type": "string",
                "description": "Source document (optional, uses active doc)",
            },
        },
        "required": ["target_path"],
    }

    def execute(self, target_path, file_path=None, **_):
        return self.services.writer.save_document_as(target_path, file_path)


class GetRecentDocuments(McpTool):
    name = "get_recent_documents"
    description = (
        "Get the list of recently opened documents from LO history. "
        "Returns file paths and titles of recently opened documents."
    )
    parameters = {
        "type": "object",
        "properties": {
            "max_count": {
                "type": "integer",
                "description": "Maximum number of documents to return (default: 20)",
            },
        },
    }

    def execute(self, max_count=20, **_):
        return self.services.writer.get_recent_documents(max_count)
