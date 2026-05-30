"""Title slide builder."""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from brand import Nature
from utils import add_bg, add_accent_bar, add_textbox, SLIDE_W, SLIDE_H


def add_title_slide(slide, data):
    """Build a title slide from data dict with {headline, subheadline, tagline, attribution, url}."""
    # Background
    add_bg(slide, Nature.BG)

    # Top accent bar
    add_accent_bar(slide, left=Inches(0), top=Inches(0), width=SLIDE_W, height=Inches(0.08),
                   color=Nature.ACCENT)

    # Left accent vertical bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.2),
                                  Inches(0.06), Inches(2.8))
    bar.fill.solid()
    bar.fill.fore_color.rgb = Nature.ACCENT
    bar.line.fill.background()

    # Headline
    add_textbox(slide, Inches(1.2), Inches(2.2), Inches(10.5), Inches(1.2),
                text=data.get("headline", ""), font_size=40, font_color=Nature.TEXT_P,
                bold=True)

    # Subheadline
    add_textbox(slide, Inches(1.2), Inches(3.5), Inches(10.5), Inches(0.6),
                text=data.get("subheadline", ""), font_size=20, font_color=Nature.TEXT_S)

    # Tagline
    tagline = data.get("tagline", "")
    if tagline:
        add_textbox(slide, Inches(1.2), Inches(4.3), Inches(10.5), Inches(0.5),
                    text=tagline, font_size=14, font_color=Nature.TEXT_M)

    # Attribution
    attribution = data.get("attribution", "")
    if attribution:
        add_textbox(slide, Inches(1.2), Inches(6.2), Inches(5), Inches(0.4),
                    text=attribution, font_size=11, font_color=Nature.TEXT_M)

    # URL
    url = data.get("url", "")
    if url:
        add_textbox(slide, Inches(1.2), Inches(6.6), Inches(5), Inches(0.3),
                    text=url, font_size=10, font_color=Nature.ACCENT)

    # Bottom accent bar
    add_accent_bar(slide, left=Inches(0), top=SLIDE_H - Inches(0.06), width=SLIDE_W,
                   height=Inches(0.06), color=Nature.ACCENT)
