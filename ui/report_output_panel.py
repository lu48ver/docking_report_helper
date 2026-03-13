import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
import queue

from ui.constants import (
    PX, PY_S, PY_B, PYW,
    SECTION_FONT, HINT_FONT, LBL_W,
    BTN_GENERATE,
)
from ui.widgets import make_section, make_card


class ReportOutputPanel(ttk.Frame):
    def __init__(self, parent, template_manager, settings=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.template_manager = template_manager
        self.on_generate = None  # callback: () -> None
        self._log_queue = queue.Queue()
        self._sample_tasks = ["SY-001 - 30%", "SY-002 - 50%", "ER-001 - 完成"]

        # Restore saved layout preference
        default_layout = "span"
        if settings and hasattr(settings, "table_layout") and settings.table_layout:
            default_layout = settings.table_layout
        self._initial_layout = default_layout

        self._build_ui()

    # ── Canvas palette (litera theme) ─────────────────────────────────────
    _C_BG      = "#f8f9fa"
    _C_BORDER  = "#ced4da"
    _C_HEADER  = "#4582ec"
    _C_HDR_TXT = "#ffffff"
    _C_PHOTO   = "#e8f0fe"
    _C_PHO_TXT = "#6c757d"

    # ── UI construction ───────────────────────────────────────────────────

    def _build_ui(self):
        # ── A. Template selector ─────────────────────────────────────────
        # TemplateSelector is now a plain ttk.Frame with its own flat header.
        from ui.template_selector import TemplateSelector
        self.template_selector = TemplateSelector(self, self.template_manager)
        self.template_selector.pack(fill=X, padx=PX, pady=(20, 0))

        ttk.Separator(self, orient="horizontal").pack(
            fill=X, padx=PX, pady=(12, 0))

        # ── B+C. Middle row: layout+canvas (LEFT) | output summary (RIGHT) ─
        mid = ttk.Frame(self, padding=(PX, 12, PX, 0))
        mid.pack(fill=X)

        # LEFT pane — table layout selection + preview canvas
        left_pane = ttk.Frame(mid)
        left_pane.pack(side=LEFT, anchor=N, padx=(0, 24))

        self._inline_section_header(left_pane, "表格版面")

        # Radio buttons in one horizontal row
        radio_row = ttk.Frame(left_pane)
        radio_row.pack(fill=X, pady=(0, 10))
        self.layout_var = ttk.StringVar(value=self._initial_layout)
        for text, value in [
            ("跨欄置中", "span"),
            ("單欄排列", "single"),
            ("雙欄並排", "grid"),
        ]:
            ttk.Radiobutton(
                radio_row, text=text, value=value,
                variable=self.layout_var,
                command=self._on_layout_changed,
            ).pack(side=LEFT, padx=(0, 20))

        # Preview canvas
        self.preview_canvas = tk.Canvas(
            left_pane,
            width=340, height=170,
            bg=self._C_BG,
            highlightthickness=1,
            highlightbackground=self._C_BORDER,
        )
        self.preview_canvas.pack()
        self.after(100, self._draw_preview)

        # RIGHT pane — output summary card
        right_pane = ttk.Frame(mid)
        right_pane.pack(side=LEFT, anchor=N, fill=X, expand=True)

        self._inline_section_header(right_pane, "輸出摘要")

        # Summary card (淡藍底色)
        card = make_card(right_pane, padding=(14, 12))
        card.pack(fill=X, pady=(0, 4))
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text="輸出檔名",
                  width=LBL_W, anchor=E,
                  font=HINT_FONT, style="DayCard.TLabel").grid(
            row=0, column=0, sticky=E, pady=(0, 8))
        self.filename_label = ttk.Label(
            card,
            text="— 請先設定日期與部門 —",
            style="DayCard.TLabel",
            font=("", 10, "bold"),
        )
        self.filename_label.grid(
            row=0, column=1, sticky=W, padx=(10, 0), pady=(0, 8))

        ttk.Label(card, text="資料筆數",
                  width=LBL_W, anchor=E,
                  font=HINT_FONT, style="DayCard.TLabel").grid(
            row=1, column=0, sticky=E)
        self.preview_label = ttk.Label(
            card, text="尚未載入資料",
            style="DayCard.TLabel",
            font=("", 9),
        )
        self.preview_label.grid(row=1, column=1, sticky=W, padx=(10, 0))

        # ── D. Generate button — the single prominent CTA ─────────────────
        ttk.Separator(self, orient="horizontal").pack(
            fill=X, padx=PX, pady=(14, 0))

        btn_frame = ttk.Frame(self, padding=(0, 14))
        btn_frame.pack(fill=X)
        self.generate_btn = ttk.Button(
            btn_frame,
            text="  ▶  產生報告  ",
            command=self._on_generate_click,
            bootstyle=BTN_GENERATE,
            width=34,
        )
        self.generate_btn.pack(anchor=CENTER)

        ttk.Separator(self, orient="horizontal").pack(
            fill=X, padx=PX, pady=(0, 4))

        # ── E. Execution log ──────────────────────────────────────────────
        log_outer = ttk.Frame(self, padding=(PX, 0, PX, 16))
        log_outer.pack(fill=BOTH, expand=True)

        self._inline_section_header(log_outer, "執行記錄")

        self.log_text = ttk.Text(log_outer, height=7,
                                 state="disabled", wrap="word",
                                 font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_outer, orient="vertical",
                                  command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.pack(fill=BOTH, expand=True)

    @staticmethod
    def _inline_section_header(parent, title: str):
        """Inline flat section header (matches make_section style)."""
        hdr = ttk.Frame(parent)
        hdr.pack(fill=X, pady=(0, 8))
        ttk.Label(hdr, text=title,
                  font=SECTION_FONT, bootstyle="secondary").pack(side=LEFT)
        ttk.Separator(hdr, orient="horizontal").pack(
            side=LEFT, fill=X, expand=True, padx=(8, 0))

    # ── Canvas preview ────────────────────────────────────────────────────

    def _on_layout_changed(self):
        self._draw_preview()

    def _draw_preview(self):
        c = self.preview_canvas
        c.delete("all")
        layout = self.layout_var.get()
        tasks = self._sample_tasks[:2]

        if layout == "span":
            self._draw_span_preview(c, tasks)
        elif layout == "single":
            self._draw_single_preview(c, tasks)
        else:
            self._draw_grid_preview(c, tasks)

    def _draw_span_preview(self, c, tasks):
        x0, y = 20, 15
        w = 300
        half = w // 2
        header_h, photo_h = 28, 45

        for task in tasks:
            c.create_rectangle(x0, y, x0 + w, y + header_h,
                                fill=self._C_HEADER, outline=self._C_BORDER)
            c.create_text(x0 + w // 2, y + header_h // 2, text=task,
                          font=("", 9, "bold"), fill=self._C_HDR_TXT)
            y += header_h
            c.create_rectangle(x0, y, x0 + half, y + photo_h,
                                fill=self._C_PHOTO, outline=self._C_BORDER)
            c.create_rectangle(x0 + half, y, x0 + w, y + photo_h,
                                fill=self._C_PHOTO, outline=self._C_BORDER)
            c.create_text(x0 + half // 2, y + photo_h // 2,
                          text="照片", font=("", 8), fill=self._C_PHO_TXT)
            c.create_text(x0 + half + half // 2, y + photo_h // 2,
                          text="照片", font=("", 8), fill=self._C_PHO_TXT)
            y += photo_h + 4

    def _draw_single_preview(self, c, tasks):
        x0, y = 20, 15
        w = 300
        content_h, photo_h = 28, 45

        for task in tasks:
            c.create_rectangle(x0, y, x0 + w, y + content_h,
                                fill=self._C_HEADER, outline=self._C_BORDER)
            c.create_text(x0 + 10, y + content_h // 2, text=task,
                          font=("", 9, "bold"), fill=self._C_HDR_TXT, anchor=W)
            y += content_h
            c.create_rectangle(x0, y, x0 + w, y + photo_h,
                                fill=self._C_PHOTO, outline=self._C_BORDER)
            c.create_text(x0 + w // 2, y + photo_h // 2,
                          text="照片", font=("", 8), fill=self._C_PHO_TXT)
            y += photo_h + 4

    def _draw_grid_preview(self, c, tasks):
        x0, y = 20, 15
        w = 300
        half = w // 2
        content_h, blank_h = 28, 45

        c.create_rectangle(x0, y, x0 + half, y + content_h,
                            fill=self._C_HEADER, outline=self._C_BORDER)
        c.create_rectangle(x0 + half, y, x0 + w, y + content_h,
                            fill=self._C_HEADER, outline=self._C_BORDER)
        if len(tasks) > 0:
            c.create_text(x0 + 8, y + content_h // 2, text=tasks[0],
                          font=("", 9, "bold"), fill=self._C_HDR_TXT, anchor=W)
        if len(tasks) > 1:
            c.create_text(x0 + half + 8, y + content_h // 2, text=tasks[1],
                          font=("", 9, "bold"), fill=self._C_HDR_TXT, anchor=W)
        y += content_h
        c.create_rectangle(x0, y, x0 + half, y + blank_h,
                            fill=self._C_PHOTO, outline=self._C_BORDER)
        c.create_rectangle(x0 + half, y, x0 + w, y + blank_h,
                            fill=self._C_PHOTO, outline=self._C_BORDER)
        c.create_text(x0 + half // 2, y + blank_h // 2,
                      text="照片", font=("", 8), fill=self._C_PHO_TXT)
        c.create_text(x0 + half + half // 2, y + blank_h // 2,
                      text="照片", font=("", 8), fill=self._C_PHO_TXT)

    def set_sample_tasks(self, tasks: list):
        self._sample_tasks = (
            [t[1] for t in tasks[:3]] if tasks else ["Task A", "Task B"])
        self._draw_preview()

    # ── Public API ────────────────────────────────────────────────────────

    def get_table_layout(self) -> str:
        return self.layout_var.get()

    def update_filename(
        self,
        day_date_str: str,
        day_str: str,
        dept_title: str = "機艙塢修報告",
        prefix: str = "SY-",
    ):
        filename = f"{prefix}{dept_title}-{day_date_str}-{day_str}.docx"
        self.filename_label.config(text=filename)

    def update_preview(self, shipyard_count: int, engine_count: int):
        self.preview_label.config(
            text=f"船廠工程: {shipyard_count} 筆  /  自修: {engine_count} 筆"
        )

    def log(self, message: str):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"{now}  {message}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _on_generate_click(self):
        if self.on_generate:
            self.on_generate()

    def get_selected_template_path(self) -> str:
        return self.template_selector.get_selected_template_path()

    def set_generating(self, is_generating: bool):
        if is_generating:
            self.generate_btn.config(state="disabled", text="⏳  產生中…")
        else:
            self.generate_btn.config(state="normal", text="  ▶  產生報告  ")
