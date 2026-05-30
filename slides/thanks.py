"""Thanks / closing slide builder."""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from brand import Nature
from utils import add_bg, add_accent_bar, add_textbox, SLIDE_W, SLIDE_H


def add_thanks_slide(slide, data):
    """Build a Thanks / closing slide.

    Schema: {headline, tagline, url}
    """
    # Background
    add_bg(slide, Nature.BG)

    # Top accent bar (thicker for emphasis)
    add_accent_bar(slide, left=Inches(0), top=Inches(0), width=SLIDE_W, height=Inches(0.08),
                   color=Nature.ACCENT)

    # Large "Thanks" text
    add_textbox(slide, Inches(0.8), Inches(2.0), Inches(11.5), Inches(1.2),
                text=data.get("headline", "Thank You"), font_size=48,
                font_color=Nature.ACCENT, bold=True, alignment=PP_ALIGN.CENTER)

    # Tagline
    tagline = data.get("tagline", "")
    if tagline:
        add_textbox(slide, Inches(0.8), Inches(3.3), Inches(11.5), Inches(0.6),
                    text=tagline, font_size=20, font_color=Nature.TEXT_S,
                    alignment=PP_ALIGN.CENTER)

    # URL
    url = data.get("url", "")
    if url:
        add_textbox(slide, Inches(0.8), Inches(4.2), Inches(11.5), Inches(0.4),
                    text=url, font_size=12, font_color=Nature.ACCENT,
                    alignment=PP_ALIGN.CENTER)

    # Decorative bottom bar
    add_accent_bar(slide, left=SLIDE_W / 2 - Inches(1), top=SLIDE_H - Inches(0.08),
                   width=Inches(2), height=Inches(0.06), color=Nature.ACCENT)
