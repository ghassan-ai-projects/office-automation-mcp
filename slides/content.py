"""Content / bullets slide builder."""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from brand import Nature
from utils import add_bg, add_accent_bar, add_textbox, SLIDE_W, SLIDE_H


def add_content_slide(slide, data):
    """Build a content/bullets slide from data dict.

    Schema: {headline, subtitle, bullets: [str], footer}
    """
    # Background
    add_bg(slide, Nature.BG)

    # Top accent bar
    add_accent_bar(slide, left=Inches(0), top=Inches(0), width=SLIDE_W, height=Inches(0.06),
                   color=Nature.ACCENT)

    # Headline
    add_textbox(slide, Inches(0.8), Inches(0.6), Inches(11.5), Inches(0.7),
                text=data.get("headline", ""), font_size=30, font_color=Nature.ACCENT,
                bold=True)

    # Separator line
    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.35),
                                  Inches(2), Inches(0.03))
    sep.fill.solid()
    sep.fill.fore_color.rgb = Nature.ACCENT
    sep.line.fill.background()

    # Subtitle
    subtitle = data.get("subtitle", "")
    if subtitle:
        add_textbox(slide, Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.4),
                    text=subtitle, font_size=16, font_color=Nature.TEXT_S)

    # Bullets
    bullets = data.get("bullets", [])
    bullet_text = "\n".join(f"•  {b}" for b in bullets) if bullets else ""
    if bullet_text:
        add_textbox(slide, Inches(0.8), Inches(2.1), Inches(11.5), Inches(4.5),
                    text=bullet_text, font_size=16, font_color=Nature.TEXT_P)

    # Footer
    footer = data.get("footer", "")
    if footer:
        add_textbox(slide, Inches(0.8), SLIDE_H - Inches(0.5), Inches(11.5), Inches(0.3),
                    text=footer, font_size=10, font_color=Nature.TEXT_M)

    # Bottom accent bar
    add_accent_bar(slide, left=Inches(0), top=SLIDE_H - Inches(0.06), width=SLIDE_W,
                   height=Inches(0.06), color=Nature.ACCENT)
