"""Late Night Commit Scale — fun slide builder."""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from brand import Nature
from utils import add_bg, add_accent_bar, add_textbox, add_shape, SLIDE_W, SLIDE_H


def add_funscale_slide(slide, data):
    """Build a 'Late Night Commit Scale' fun slide.

    Schema: {headline, subtitle, levels: [{time, label, description}]}
    Levels are displayed horizontally as a scale with severity colors.
    """
    # Background
    add_bg(slide, Nature.BG)

    # Top accent bar
    add_accent_bar(slide, left=Inches(0), top=Inches(0), width=SLIDE_W, height=Inches(0.06),
                   color=Nature.ACCENT)

    # Headline
    add_textbox(slide, Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.5),
                text=data.get("headline", ""), font_size=28, font_color=Nature.ACCENT,
                bold=True)

    # Subtitle
    subtitle = data.get("subtitle", "")
    if subtitle:
        add_textbox(slide, Inches(0.8), Inches(0.9), Inches(11.5), Inches(0.35),
                    text=subtitle, font_size=14, font_color=Nature.TEXT_S)

    levels = data.get("levels", [])

    if not levels:
        return

    # Scale colors — gradient from calm green to intense red
    def _level_color(idx, total):
        palette = [
            Nature.SUCCESS,       # Safe
            RGBColor(0x4C, 0xAF, 0x50),  # Green
            RGBColor(0xFF, 0x98, 0x00),  # Orange
            RGBColor(0xFF, 0x57, 0x22),  # Deep orange
            Nature.ACCENT2,       # Dark green
            RGBColor(0x7B, 0x1F, 0xA2),  # Purple
            RGBColor(0xD3, 0x2F, 0x2F),  # Red
            RGBColor(0x00, 0x00, 0x00),  # Black
        ]
        return palette[idx % len(palette)]

    num_levels = min(len(levels), 8)
    bar_w = (SLIDE_W - Inches(1.6)) / num_levels
    bar_h = Inches(2.2)
    bar_y = Inches(1.8)

    for i in range(num_levels):
        lvl = levels[i]
        x = Inches(0.8) + i * bar_w

        # Level bar
        color = _level_color(i, num_levels)
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, bar_y,
                                      bar_w - Inches(0.08), bar_h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()
        bar.adjustments[0] = 0.08

        # Time label
        add_textbox(slide, x, bar_y + Inches(0.15), bar_w - Inches(0.08), Inches(0.35),
                    text=lvl.get("time", ""), font_size=12, font_color=Nature.WHITE,
                    bold=True, alignment=PP_ALIGN.CENTER)

        # Label
        add_textbox(slide, x, bar_y + Inches(0.5), bar_w - Inches(0.08), Inches(0.4),
                    text=lvl.get("label", ""), font_size=11, font_color=Nature.WHITE,
                    bold=False, alignment=PP_ALIGN.CENTER)

        # Description
        add_textbox(slide, x, bar_y + Inches(0.95), bar_w - Inches(0.08), Inches(1.1),
                    text=lvl.get("description", ""), font_size=9, font_color=Nature.WHITE,
                    alignment=PP_ALIGN.CENTER)

    # Bottom accent bar
    add_accent_bar(slide, left=Inches(0), top=SLIDE_H - Inches(0.06), width=SLIDE_W,
                   height=Inches(0.06), color=Nature.ACCENT)
