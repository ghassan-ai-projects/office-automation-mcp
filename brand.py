"""Nature Theme brand colors and font definitions."""

from pptx.dml.color import RGBColor

# ── Nature Theme — Sage White + Forest Green ──────────────────────────────

class Nature:
    NAME = "Nature"
    FONT = "Calibri"

    # Backgrounds
    BG      = RGBColor(0xF5, 0xF7, 0xF5)  # Sage white
    CARD_BG = RGBColor(0xED, 0xF2, 0xED)  # Light sage
    CARD2   = RGBColor(0xE1, 0xEA, 0xE1)  # Sage tint
    BORDER  = RGBColor(0xB0, 0xC4, 0xB0)  # Sage border

    # Accents
    ACCENT  = RGBColor(0x1D, 0x6F, 0x42)  # Forest emerald ★
    ACCENT2 = RGBColor(0x15, 0x55, 0x34)  # Darker green
    SUCCESS = RGBColor(0x0E, 0x74, 0x90)  # Ocean blue

    # Text
    TEXT_P  = RGBColor(0x1A, 0x2E, 0x1A)  # Deep forest
    TEXT_S  = RGBColor(0x4A, 0x67, 0x41)  # Sage green
    TEXT_M  = RGBColor(0x6B, 0x82, 0x68)  # Muted
    WHITE   = RGBColor(0xFF, 0xFF, 0xFF)

    # Background hex strings (for set_slide_background)
    BG_HEX      = "#F5F7F5"
    CARD_BG_HEX = "#EDF2ED"
    CARD2_HEX   = "#E1EAE1"
    ACCENT_HEX  = "#1D6F42"

    @classmethod
    def as_dict(cls):
        return {
            "name": cls.NAME,
            "font": cls.FONT,
            "background": cls.BG_HEX,
            "card_bg": cls.CARD_BG_HEX,
            "card2": cls.CARD2_HEX,
            "border": "#B0C4B0",
            "accent": cls.ACCENT_HEX,
            "accent2": "#155534",
            "success": "#0E7490",
            "text_primary": "#1A2E1A",
            "text_secondary": "#4A6741",
            "text_muted": "#6B8268",
            "white": "#FFFFFF",
        }


# ── Light Professional theme (optional) ───────────────────────────────────

class LightProfessional:
    NAME = "Light Professional"
    FONT = "Calibri"

    BG      = RGBColor(0xFF, 0xFF, 0xFF)
    CARD_BG = RGBColor(0xF3, 0xF4, 0xF6)
    CARD2   = RGBColor(0xE8, 0xEA, 0xED)
    BORDER  = RGBColor(0xD1, 0xD5, 0xDB)
    ACCENT  = RGBColor(0x25, 0x63, 0xEB)
    ACCENT2 = RGBColor(0x1E, 0x40, 0xAF)
    SUCCESS = RGBColor(0x05, 0x9C, 0x69)
    TEXT_P  = RGBColor(0x11, 0x18, 0x27)
    TEXT_S  = RGBColor(0x4B, 0x55, 0x63)
    TEXT_M  = RGBColor(0x6B, 0x72, 0x80)
    WHITE   = RGBColor(0xFF, 0xFF, 0xFF)

    @classmethod
    def as_dict(cls):
        return {
            "name": cls.NAME,
            "font": cls.FONT,
            "background": "#FFFFFF",
            "card_bg": "#F3F4F6",
            "card2": "#E8EAED",
            "border": "#D1D5DB",
            "accent": "#2563EB",
            "accent2": "#1E40AF",
            "success": "#059C69",
            "text_primary": "#111827",
            "text_secondary": "#4B5563",
            "text_muted": "#6B7280",
            "white": "#FFFFFF",
        }


# ── Template registry ─────────────────────────────────────────────────────

TEMPLATES = {
    "nature": Nature,
    "light-professional": LightProfessional,
}

TEMPLATE_NAMES = list(TEMPLATES.keys())
