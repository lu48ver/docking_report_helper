import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from ui.constants import (
    PAGE_PAD_X,
    PAGE_PAD_Y,
    SECTION_GAP,
    SECTION_PAD,
    COMPACT_SECTION_PAD,
)


def make_page(parent, padding=None, style="Page.TFrame") -> ttk.Frame:
    return ttk.Frame(parent, style=style, padding=padding or (PAGE_PAD_X, PAGE_PAD_Y))


def make_section(
    parent,
    title: str,
    description: str = "",
    style: str = "Surface.TFrame",
    compact: bool = False,
) -> ttk.Frame:
    wrapper = ttk.Frame(
        parent,
        style=style,
        padding=COMPACT_SECTION_PAD if compact else SECTION_PAD,
    )
    wrapper.pack(fill=X, pady=(0, SECTION_GAP))

    header = ttk.Frame(wrapper, style=style)
    header.pack(fill=X, pady=(0, 12))
    ttk.Label(header, text=title, style="SectionTitle.TLabel").pack(side=LEFT)
    ttk.Separator(header, orient="horizontal").pack(
        side=LEFT, fill=X, expand=True, padx=(10, 0)
    )

    if description:
        ttk.Label(
            wrapper,
            text=description,
            style="SectionDesc.TLabel",
            justify=LEFT,
            wraplength=900,
        ).pack(fill=X, pady=(0, 10))

    body = ttk.Frame(wrapper, style=style)
    body.pack(fill=BOTH, expand=True)
    return body


def make_card(parent, style="AccentCard.TFrame", padding=(16, 14), **kwargs) -> ttk.Frame:
    return ttk.Frame(parent, style=style, padding=padding, **kwargs)


def make_stat_card(
    parent,
    title: str,
    value: str = "",
    note: str = "",
    style="AccentCard.TFrame",
    eyebrow_style="CardEyebrow.TLabel",
    value_style="CardValue.TLabel",
    note_style="CardNote.TLabel",
):
    card = make_card(parent, style=style)
    ttk.Label(card, text=title, style=eyebrow_style).pack(anchor=W)
    value_label = ttk.Label(card, text=value, style=value_style)
    value_label.pack(anchor=W, pady=(6, 2))
    note_label = ttk.Label(card, text=note, style=note_style)
    note_label.pack(anchor=W)
    return card, value_label, note_label
