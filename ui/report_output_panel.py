import tkinter as tk
from datetime import datetime
import queue

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame

from ui.constants import BTN_GENERATE
from ui.widgets import make_page, make_section, make_stat_card


class ReportOutputPanel(ttk.Frame):
    _C_BG = "#f8f9fa"
    _C_BORDER = "#ced4da"
    _C_HEADER = "#4582ec"
    _C_HDR_TXT = "#ffffff"
    _C_PHOTO = "#e8f0fe"
    _C_PHO_TXT = "#6c757d"

    def __init__(self, parent, template_manager, settings=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.template_manager = template_manager
        self.on_generate = None
        self._log_queue = queue.Queue()
        self._sample_tasks = ["SY-001 - 30%", "SY-002 - 50%", "ER-001 - 完成"]
        self._initial_layout = (
            settings.table_layout
            if settings and hasattr(settings, "table_layout") and settings.table_layout
            else "span"
        )
        self.preview_cards = {}

        self._build_ui()

    def _build_ui(self):
        scroller = ScrolledFrame(self, autohide=True)
        scroller.pack(fill=BOTH, expand=True)
        page = make_page(scroller, padding=(20, 18))
        page.pack(fill=BOTH, expand=True)

        from ui.template_selector import TemplateSelector

        template_section = make_section(
            page,
            "模板選擇",
            "先選擇 Word 模板，再繼續設定版面與輸出內容。",
        )
        template_section.columnconfigure(0, weight=1)
        self.template_selector = TemplateSelector(template_section, self.template_manager)
        self.template_selector.grid(row=0, column=0, sticky=EW)

        mid = ttk.Frame(page, style="Page.TFrame")
        mid.pack(fill=X, pady=(0, 22))
        mid.columnconfigure(0, weight=1, uniform="report_cols")
        mid.columnconfigure(1, weight=1, uniform="report_cols")
        left_col = ttk.Frame(mid, style="Page.TFrame")
        left_col.grid(row=0, column=0, sticky=NSEW, padx=(0, 10))
        right_col = ttk.Frame(mid, style="Page.TFrame")
        right_col.grid(row=0, column=1, sticky=NSEW, padx=(10, 0))

        layout_section = make_section(
            left_col,
            "版面選擇",
            "選擇這次報告中工作項目與照片的排列方式。",
        )
        self.layout_section_wrapper = layout_section.master
        self.layout_var = ttk.StringVar(value=self._initial_layout)
        options = ttk.Frame(layout_section, style="Surface.TFrame")
        options.pack(fill=X)
        for idx, (text, value) in enumerate([
            ("跨欄置中", "span"),
            ("單欄排列", "single"),
            ("雙欄並排", "grid"),
        ]):
            options.columnconfigure(idx, weight=1)
            card = ttk.Frame(options, style="OptionCard.TFrame", padding=(12, 12))
            card.grid(row=0, column=idx, sticky=NSEW, padx=(0, 10 if idx < 2 else 0))
            ttk.Radiobutton(
                card,
                text=text,
                value=value,
                variable=self.layout_var,
                command=self._on_layout_changed,
            ).pack(anchor=W)
            canvas = tk.Canvas(
                card,
                width=220,
                height=150,
                bg=self._C_BG,
                highlightthickness=1,
                highlightbackground=self._C_BORDER,
            )
            canvas.pack(fill=X, pady=(10, 0))
            card.bind("<Button-1>", lambda event, layout=value: self._select_layout(layout))
            canvas.bind("<Button-1>", lambda event, layout=value: self._select_layout(layout))
            self.preview_cards[value] = (card, canvas)

        summary_section = make_section(
            right_col,
            "輸出資訊摘要",
            "最後確認這次即將輸出的檔名與資料筆數。",
        )
        self.summary_section_wrapper = summary_section.master
        filename_card, self.filename_label, self.filename_note = make_stat_card(
            summary_section,
            "輸出檔名",
            "— 請先設定日期與部門 —",
            "檔名會依設定頁的日期、部門與前綴自動組合。",
            style="AccentCard.TFrame",
        )
        filename_card.pack(fill=X, pady=(0, 12))

        count_card, self.preview_label, self.preview_note = make_stat_card(
            summary_section,
            "資料筆數摘要",
            "尚未載入資料",
            "載入 Excel 後會顯示船廠工程與自修筆數。",
            style="SummaryCard.TFrame",
            eyebrow_style="SummaryEyebrow.TLabel",
            value_style="SummaryValue.TLabel",
            note_style="SummaryNote.TLabel",
        )
        count_card.pack(fill=X)

        btn_frame = ttk.Frame(page, style="Page.TFrame")
        btn_frame.pack(fill=X, pady=(2, 16))
        self.generate_btn = ttk.Button(
            btn_frame,
            text="產生報告",
            command=self._on_generate_click,
            bootstyle=BTN_GENERATE,
            width=24,
        )
        self.generate_btn.pack(anchor=CENTER)

        log_section = make_section(
            page,
            "執行記錄",
            "產生報告過程與回饋資訊會顯示在這裡。",
            compact=True,
        )
        log_section.configure(height=220)
        log_section.pack_propagate(False)
        self.log_text = ttk.Text(
            log_section,
            height=10,
            state="disabled",
            wrap="word",
            font=("Consolas", 9),
        )
        scrollbar = ttk.Scrollbar(log_section, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.pack(fill=BOTH, expand=True)

        self.after(100, self._draw_preview)
        self.after(120, self._refresh_option_cards)
        self.after(140, self._balance_middle_sections)

    def _select_layout(self, value: str):
        self.layout_var.set(value)
        self._on_layout_changed()

    def _on_layout_changed(self):
        self._draw_preview()
        self._refresh_option_cards()

    def _draw_preview(self):
        tasks = self._sample_tasks[:2]
        for layout, (_, canvas) in self.preview_cards.items():
            canvas.delete("all")
            if layout == "span":
                self._draw_span_preview(canvas, tasks)
            elif layout == "single":
                self._draw_single_preview(canvas, tasks)
            else:
                self._draw_grid_preview(canvas, tasks)

    def _refresh_option_cards(self):
        selected = self.layout_var.get()
        for value, (card, _) in self.preview_cards.items():
            card.configure(style="SelectedOptionCard.TFrame" if value == selected else "OptionCard.TFrame")

    def _balance_middle_sections(self):
        self.update_idletasks()
        max_height = max(
            self.layout_section_wrapper.winfo_reqheight(),
            self.summary_section_wrapper.winfo_reqheight(),
        )
        for wrapper in (self.layout_section_wrapper, self.summary_section_wrapper):
            wrapper.configure(height=max_height)
            wrapper.pack_propagate(False)

    def _draw_span_preview(self, canvas, tasks):
        x0, y = 16, 12
        width = 188
        half = width // 2
        header_h, photo_h = 26, 40
        for task in tasks:
            canvas.create_rectangle(x0, y, x0 + width, y + header_h, fill=self._C_HEADER, outline=self._C_BORDER)
            canvas.create_text(x0 + width // 2, y + header_h // 2, text=task, font=("", 9, "bold"), fill=self._C_HDR_TXT)
            y += header_h
            canvas.create_rectangle(x0, y, x0 + half, y + photo_h, fill=self._C_PHOTO, outline=self._C_BORDER)
            canvas.create_rectangle(x0 + half, y, x0 + width, y + photo_h, fill=self._C_PHOTO, outline=self._C_BORDER)
            canvas.create_text(x0 + half // 2, y + photo_h // 2, text="照片", font=("", 8), fill=self._C_PHO_TXT)
            canvas.create_text(x0 + half + half // 2, y + photo_h // 2, text="照片", font=("", 8), fill=self._C_PHO_TXT)
            y += photo_h + 4

    def _draw_single_preview(self, canvas, tasks):
        x0, y = 16, 12
        width = 188
        header_h, photo_h = 26, 40
        for task in tasks:
            canvas.create_rectangle(x0, y, x0 + width, y + header_h, fill=self._C_HEADER, outline=self._C_BORDER)
            canvas.create_text(x0 + 8, y + header_h // 2, text=task, font=("", 9, "bold"), fill=self._C_HDR_TXT, anchor=W)
            y += header_h
            canvas.create_rectangle(x0, y, x0 + width, y + photo_h, fill=self._C_PHOTO, outline=self._C_BORDER)
            canvas.create_text(x0 + width // 2, y + photo_h // 2, text="照片", font=("", 8), fill=self._C_PHO_TXT)
            y += photo_h + 4

    def _draw_grid_preview(self, canvas, tasks):
        x0, y = 16, 28
        width = 188
        half = width // 2
        header_h, photo_h = 26, 44
        canvas.create_rectangle(x0, y, x0 + half, y + header_h, fill=self._C_HEADER, outline=self._C_BORDER)
        canvas.create_rectangle(x0 + half, y, x0 + width, y + header_h, fill=self._C_HEADER, outline=self._C_BORDER)
        if len(tasks) > 0:
            canvas.create_text(x0 + 6, y + header_h // 2, text=tasks[0], font=("", 8, "bold"), fill=self._C_HDR_TXT, anchor=W)
        if len(tasks) > 1:
            canvas.create_text(x0 + half + 6, y + header_h // 2, text=tasks[1], font=("", 8, "bold"), fill=self._C_HDR_TXT, anchor=W)
        y += header_h
        canvas.create_rectangle(x0, y, x0 + half, y + photo_h, fill=self._C_PHOTO, outline=self._C_BORDER)
        canvas.create_rectangle(x0 + half, y, x0 + width, y + photo_h, fill=self._C_PHOTO, outline=self._C_BORDER)
        canvas.create_text(x0 + half // 2, y + photo_h // 2, text="照片", font=("", 8), fill=self._C_PHO_TXT)
        canvas.create_text(x0 + half + half // 2, y + photo_h // 2, text="照片", font=("", 8), fill=self._C_PHO_TXT)

    def set_sample_tasks(self, tasks: list):
        self._sample_tasks = [task[1] for task in tasks[:3]] if tasks else ["Task A", "Task B"]
        self._draw_preview()

    def get_table_layout(self) -> str:
        return self.layout_var.get()

    def update_filename(self, day_date_str: str, day_str: str, dept_title: str = "機艙塢修報告", prefix: str = "SY-"):
        filename = f"{prefix}{dept_title}-{day_date_str}-{day_str}.docx"
        self.filename_label.config(text=filename)

    def update_preview(self, shipyard_count: int, engine_count: int):
        total = shipyard_count + engine_count
        self.preview_label.config(text=f"{total} 筆資料")
        self.preview_note.config(text=f"船廠工程 {shipyard_count} 筆  |  自修 {engine_count} 筆")

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
            self.generate_btn.config(state="disabled", text="產生中…")
        else:
            self.generate_btn.config(state="normal", text="產生報告")
