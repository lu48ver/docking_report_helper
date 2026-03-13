"""ui/widgets.py — 共用 UI 輔助函式 (v3)

三個頁面都 import 這裡的建構器，確保 section header / card 樣式完全一致。
"""
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from ui.constants import PX, PY_S, PY_B, SECTION_FONT


def make_section(parent, title: str, last: bool = False) -> ttk.Frame:
    """
    建立扁平 section header + body frame，回傳 body。

    格式：  TITLE ──────────────────
            [body frame]

    body 已設定 columnconfigure(1, weight=1)，適合 3 欄 grid 表單。

    Parameters
    ----------
    parent : widget
        Section 要放入的父容器（通常是 ScrolledFrame 或 Frame）。
    title : str
        Section 標題文字。
    last : bool
        True → 使用 PY_B padding（最後一個 section 多留底部空間）。
    """
    pady = PY_B if last else PY_S

    wrapper = ttk.Frame(parent)
    wrapper.pack(fill=X, padx=PX, pady=pady)

    # Header row: bold label + expanding separator
    hdr = ttk.Frame(wrapper)
    hdr.pack(fill=X, pady=(0, 10))
    ttk.Label(hdr, text=title,
              font=SECTION_FONT, bootstyle="secondary").pack(side=LEFT)
    ttk.Separator(hdr, orient="horizontal").pack(
        side=LEFT, fill=X, expand=True, padx=(8, 0))

    # Body frame — callers put widgets here via grid()
    body = ttk.Frame(wrapper)
    body.pack(fill=X, padx=4)
    body.columnconfigure(1, weight=1)
    return body


def make_card(parent, **kwargs) -> ttk.Frame:
    """
    建立 DayCard 樣式的 info card frame（淡藍底）。

    必須在 ttk.Style 已注冊 "DayCard.TFrame" 之後呼叫
    （由 MainWindow._configure_styles() 完成）。

    Parameters
    ----------
    parent : widget
        Card 的父容器。
    **kwargs
        額外傳給 ttk.Frame 的參數（例如 padding）。
    """
    padding = kwargs.pop("padding", (16, 10))
    card = ttk.Frame(parent, style="DayCard.TFrame", padding=padding, **kwargs)
    return card
