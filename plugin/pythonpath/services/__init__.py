"""
ServiceRegistry — container for all UNO domain services.

Instantiated once by the MCP server and injected into every tool.
"""

import logging

from .base import BaseService

logger = logging.getLogger(__name__)


class ServiceRegistry:
    """Holds every domain service; passed to McpTool constructors."""

    def __init__(self):
        # Core infrastructure (must be first)
        self.base = BaseService()
        self.base._registry = self

        # Domain services — imported lazily to avoid circular deps
        # and to keep startup fast.  Each service receives `self`
        # so it can access other services via the registry.
        from .writer import WriterService
        from .calc import CalcService
        from .impress import ImpressService
        from .images import ImageService
        from .comments import CommentService
        from .tables import TableService
        from .styles import StyleService

        self.writer = WriterService(self)
        self.calc = CalcService(self)
        self.impress = ImpressService(self)
        self.images = ImageService(self)
        self.comments = CommentService(self)
        self.tables = TableService(self)
        self.styles = StyleService(self)

        logger.info("ServiceRegistry ready (%d services)", 8)

    # Convenience shortcuts delegated to base
    @property
    def desktop(self):
        return self.base.desktop

    def resolve_document(self, path=None):
        return self.base.resolve_document(path)

    def resolve_locator(self, doc, locator):
        return self.base.resolve_locator(doc, locator)
