"""Quote / testimonial slide builder."""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from brand import Nature
from utils import add_bg, add_accent_bar, add_textbox, SLIDE_W, SLIDE_H


def add_quotes_slide(slide, data):
    """Build a quote/testimonial slide.

    Schema: {headline, subtitle, quotes: [{text, attribution}]}
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

    quotes = data.get("quotes", [])

    # Layout quotes vertically or side-by-side
    num_quotes = min(len(quotes), 4)
    if num_quotes == 0:
        return

    if num_quotes == 1:
        _render_quote(slide, quotes[0], Inches(1.5), Inches(1.6), Inches(10), Inches(4.5), big=True)
    else:
        # Side by side
        q_w = (SLIDE_W - Inches(2.2)) / num_quotes
        for i in range(num_quotes):
            left = Inches(0.8) + i * (q_w + Inches(0.3))
            _render_quote(slide, quotes[i], left, Inches(1.6), q_w, Inches(4.5))

    # Bottom accent bar
    add_accent_bar(slide, left=Inches(0), top=SLIDE_H - Inches(0.06), width=SLIDE_W,
                   height=Inches(0.06), color=Nature.ACCENT)


def _render_quote(slide, quote, left, top, width, height, big=False):
    """Render a single quote with opening quotation mark style."""
    text = quote.get("text", "")
    attribution = quote.get("attribution", "")

    # Quote card background
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = Nature.CARD_BG
    card.line.color.rgb = Nature.BORDER
    card.line.width = Pt(1)
    card.adjustments[0] = 0.04

    # Big quotation mark
    add_textbox(slide, left + Inches(0.3), top + Inches(0.15), Inches(1), Inches(0.7),
                text='"', font_size=48 if not big else 60, font_color=Nature.ACCENT,
                bold=True)

    # Quote text
    q_size = 16 if not big else 20
    add_textbox(slide, left + Inches(0.3), top + Inches(0.7), width - Inches(0.6), height - Inches(1.8),
                text=text, font_size=q_size, font_color=Nature.TEXT_P)

    # Attribution
    if attribution:
        add_textbox(slide, left + Inches(0.3), top + height - Inches(0.5), width - Inches(0.6), Inches(0.35),
                    text=f"— {attribution}", font_size=12, font_color=Nature.TEXT_M,
                    alignment=PP_ALIGN.RIGHT)
