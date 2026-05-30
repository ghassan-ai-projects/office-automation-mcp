"""
Theme system — loads brand/font definitions from themes.json.

At import time, reads themes.json (cached) and dynamically creates
theme classes.  Backward-compatible: ``Nature``, ``TEMPLATES``,
``TEMPLATE_NAMES`` are still exported as module globals.
"""

from __future__ import annotations

import json
import os
import functools
from typing import Any

from pptx.dml.color import RGBColor


# ── Paths ─────────────────────────────────────────────────────────────────

_THEMES_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "themes.json")


# ── Helpers ───────────────────────────────────────────────────────────────

def _hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert '#RRGGBB' or 'RRGGBB' to RGBColor."""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


@functools.lru_cache(maxsize=1)
def _load_themes_json() -> dict[str, Any]:
    """Read themes.json once (cached for process lifetime)."""
    with open(_THEMES_PATH, encoding="utf-8") as f:
        data = json.load(f)
    return data


def _make_theme_class(theme_key: str, theme_data: dict) -> type:
    """Dynamically build a theme class from a JSON theme entry.

    The returned class has the same attribute names as the original
    hardcoded ``Nature`` / ``LightProfessional`` classes so that all
    existing importers (``slides/*.py``, ``utils.py``) keep working
    without changes.
    """
    colors = theme_data["colors"]

    # Hex convenience attributes (used by as_dict / set_slide_background)
    bg_hex = colors["bg"]
    card_bg_hex = colors["card_bg"]
    card2_hex = colors["card2"]
    accent_hex = colors["accent"]

    attrs: dict[str, Any] = {
        "NAME": theme_data.get("name", theme_key.title()),
        "FONT": theme_data.get("font", "Calibri"),
        # RGBColor attributes (used by all slide builders)
        "BG": _hex_to_rgb(bg_hex),
        "CARD_BG": _hex_to_rgb(card_bg_hex),
        "CARD2": _hex_to_rgb(card2_hex),
        "BORDER": _hex_to_rgb(colors["border"]),
        "ACCENT": _hex_to_rgb(accent_hex),
        "ACCENT2": _hex_to_rgb(colors["accent2"]),
        "SUCCESS": _hex_to_rgb(colors["success"]),
        "TEXT_P": _hex_to_rgb(colors["text_primary"]),
        "TEXT_S": _hex_to_rgb(colors["text_secondary"]),
        "TEXT_M": _hex_to_rgb(colors["text_muted"]),
        "WHITE": _hex_to_rgb(colors["white"]),
        # Hex string attributes (for set_slide_background etc.)
        "BG_HEX": bg_hex,
        "CARD_BG_HEX": card_bg_hex,
        "CARD2_HEX": card2_hex,
        "ACCENT_HEX": accent_hex,
    }

    # Keep raw colors dict for as_dict — use original key names for compat
    def _as_dict(cls) -> dict[str, str]:
        return {
            "name": cls.NAME,
            "font": cls.FONT,
            "background": bg_hex,
            "card_bg": card_bg_hex,
            "card2": card2_hex,
            "border": colors["border"],
            "accent": accent_hex,
            "accent2": colors["accent2"],
            "success": colors["success"],
            "text_primary": colors["text_primary"],
            "text_secondary": colors["text_secondary"],
            "text_muted": colors["text_muted"],
            "white": colors["white"],
        }

    attrs["as_dict"] = classmethod(_as_dict)

    cls = type(theme_data.get("name", theme_key.title()), (), attrs)
    return cls


# ── Public API ────────────────────────────────────────────────────────────

def load_themes() -> dict[str, type]:
    """Return dict mapping theme keys (e.g. ``"nature"``) to theme classes.

    Themes are read from the JSON file on first call and cached thereafter.
    """
    data = _load_themes_json()
    themes = {}
    for key, theme_data in data["themes"].items():
        themes[key] = _make_theme_class(key, theme_data)
    return themes


_THEMES: dict[str, type] | None = None


def _get_themes() -> dict[str, type]:
    global _THEMES
    if _THEMES is None:
        _THEMES = load_themes()
    return _THEMES


def get_theme(name: str, default: str | None = None) -> type:
    """Look up a theme class by key, falling back to *default* (or ``Nature``)."""
    themes = _get_themes()
    theme = themes.get(name)
    if theme is not None:
        return theme
    # Fallback chain
    fallback = default if default else "nature"
    if fallback in themes:
        return themes[fallback]
    # Absolute last resort — raise clear error
    raise KeyError(
        f"Theme {name!r} not found and fallback {fallback!r} also not found. "
        f"Available themes: {list(themes)}"
    )


def load_config() -> dict[str, Any]:
    """Return the ``defaults`` section from themes.json.

    Includes ``slide_width_inches``, ``slide_height_inches``,
    ``default_font``, ``footer_text``, ``auto_save``, etc.
    """
    data = _load_themes_json()
    return dict(data.get("defaults", {}))


def get_theme_names() -> list[str]:
    """Return sorted list of available theme keys."""
    return sorted(_get_themes().keys())


# ── Module-level globals (backward compatibility) ─────────────────────────

_THEMES = _get_themes()
TEMPLATES: dict[str, type] = _THEMES
TEMPLATE_NAMES: list[str] = list(_THEMES.keys())

Nature: type = _THEMES.get("nature")
LightProfessional: type = _THEMES.get("light-professional")
