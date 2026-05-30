"""2×2 Bento card grid slide builder."""

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from brand import Nature
from utils import add_bg, add_accent_bar, add_textbox, add_card, SLIDE_W, SLIDE_H


def add_cards_slide(slide, data):
    """Build a 2×2 bento card grid slide from data dict.

    Schema: {headline, subtitle, cards: [{title, description, tags: [str]}, ...]}
    Cards can be 2 (1 row) or 4 (2×2).
    """
    # Background
    add_bg(slide, Nature.BG)

    # Top accent bar
    add_accent_bar(slide, left=Inches(0), top=Inches(0), width=SLIDE_W, height=Inches(0.06),
                   color=Nature.ACCENT)

    # Headline
    add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(0.5),
                text=data.get("headline", ""), font_size=28, font_color=Nature.ACCENT,
                bold=True)

    # Subtitle
    subtitle = data.get("subtitle", "")
    if subtitle:
        add_textbox(slide, Inches(0.8), Inches(1.05), Inches(11.5), Inches(0.35),
                    text=subtitle, font_size=14, font_color=Nature.TEXT_S)

    cards = data.get("cards", [])

    # Card dimensions
    gap = Inches(0.2)
    card_w = (SLIDE_W - Inches(1.6) - gap) / 2 if len(cards) >= 2 else (SLIDE_W - Inches(1.6))
    card_h = Inches(2.5)
    start_y = Inches(1.6)

    if len(cards) <= 2:
        # Single row
        for i, c in enumerate(cards):
            left = Inches(0.8) + i * (card_w + gap)
            add_card(slide, left, start_y, card_w, card_h,
                     title=c.get("title", ""),
                     description=c.get("description", ""),
                     tags=c.get("tags", []))
    else:
        # 2×2 grid (up to 4)
        row_h = card_h + gap
        for i, c in enumerate(cards[:4]):
            col = i % 2
            row = i // 2
            left = Inches(0.8) + col * (card_w + gap)
            top = start_y + row * row_h
            add_card(slide, left, top, card_w, card_h,
                     title=c.get("title", ""),
                     description=c.get("description", ""),
                     tags=c.get("tags", []))

    # Bottom accent bar
    add_accent_bar(slide, left=Inches(0), top=SLIDE_H - Inches(0.06), width=SLIDE_W,
                   height=Inches(0.06), color=Nature.ACCENT)
