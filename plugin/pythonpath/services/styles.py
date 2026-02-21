"""
StyleService — style listing and introspection.
"""

import logging
from typing import Any, Dict

logger = logging.getLogger(__name__)


class StyleService:
    """Style operations via UNO."""

    def __init__(self, registry):
        self._registry = registry
        self._base = registry.base

    def list_styles(self, family: str = "ParagraphStyles",
                    file_path: str = None) -> Dict[str, Any]:
        """List available styles in a family."""
        try:
            doc = self._base.resolve_document(file_path)
            families = doc.getStyleFamilies()

            if not families.hasByName(family):
                available = list(families.getElementNames())
                return {"success": False,
                        "error": f"Unknown style family: {family}",
                        "available_families": available}

            style_family = families.getByName(family)
            styles = []
            for name in style_family.getElementNames():
                style = style_family.getByName(name)
                entry = {
                    "name": name,
                    "is_user_defined": style.isUserDefined(),
                    "is_in_use": style.isInUse(),
                }
                try:
                    entry["parent_style"] = style.getPropertyValue(
                        "ParentStyle")
                except Exception:
                    pass
                styles.append(entry)

            return {"success": True, "family": family,
                    "styles": styles, "count": len(styles)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_style_info(self, style_name: str,
                       family: str = "ParagraphStyles",
                       file_path: str = None) -> Dict[str, Any]:
        """Get detailed properties of a style."""
        try:
            doc = self._base.resolve_document(file_path)
            families = doc.getStyleFamilies()
            style_family = families.getByName(family)

            if not style_family.hasByName(style_name):
                return {"success": False,
                        "error": f"Style '{style_name}' not found "
                                 f"in {family}"}

            style = style_family.getByName(style_name)
            info = {
                "name": style_name,
                "family": family,
                "is_user_defined": style.isUserDefined(),
                "is_in_use": style.isInUse(),
            }

            props_to_read = {
                "ParagraphStyles": [
                    "ParentStyle", "FollowStyle",
                    "CharFontName", "CharHeight", "CharWeight",
                    "ParaAdjust", "ParaTopMargin", "ParaBottomMargin",
                ],
                "CharacterStyles": [
                    "ParentStyle", "CharFontName", "CharHeight",
                    "CharWeight", "CharPosture", "CharColor",
                ],
            }

            for prop_name in props_to_read.get(family, []):
                try:
                    val = style.getPropertyValue(prop_name)
                    info[prop_name] = val
                except Exception:
                    pass

            return {"success": True, **info}
        except Exception as e:
            return {"success": False, "error": str(e)}
