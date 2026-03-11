import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
import threading
import queue


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

    def _build_ui(self):
        # Template selector
        from ui.template_selector import TemplateSelector
        self.template_selector = TemplateSelector(self, self.template_manager)
        self.template_selector.pack(fill=BOTH, padx=10, pady=(10, 5), expand=False)

        # ── Table layout selection + preview ────────────────
        layout_frame = ttk.LabelFrame(self, text="表格版面")
        layout_frame.pack(fill=X, padx=10, pady=5)

        # Left side: radio buttons
        radio_frame = ttk.Frame(layout_frame)
        radio_frame.pack(side=LEFT, padx=10, pady=10, anchor=N)

        self.layout_var = ttk.StringVar(value=self._initial_layout)
        layouts = [
            ("跨欄置中", "span"),
            ("單欄排列", "single"),
            ("雙欄並排", "grid"),
        ]
        for text, value in layouts:
            rb = ttk.Radiobutton(
                radio_frame, text=text, value=value,
                variable=self.layout_var,
                command=self._on_layout_changed,
            )
            rb.pack(anchor=W, pady=3)

        # Right side: canvas preview
        self.preview_canvas = tk.Canvas(
            layout_frame, width=340, height=180,
            bg="white", highlightthickness=1, highlightbackground="#cccccc",
        )
        self.preview_canvas.pack(side=LEFT, padx=10, pady=10)

        # Draw initial preview
        self.after(100, self._draw_preview)

        # ── Output info ─────────────────────────────────────
        output_frame = ttk.LabelFrame(self, text="輸出資訊")
        output_frame.pack(fill=X, padx=10, pady=5)

        ttk.Label(output_frame, text="輸出檔名:").grid(row=0, column=0, sticky=W, pady=2)
        self.filename_label = ttk.Label(output_frame, text="(尚未計算)", bootstyle="info")
        self.filename_label.grid(row=0, column=1, sticky=W, padx=5, pady=2)

        # Preview counts
        ttk.Label(output_frame, text="資料筆數:").grid(row=1, column=0, sticky=W, pady=2)
        self.preview_label = ttk.Label(output_frame, text="尚未載入資料", bootstyle="secondary")
        self.preview_label.grid(row=1, column=1, sticky=W, padx=5, pady=2)

        # Generate button
        self.generate_btn = ttk.Button(
            self,
            text="產生報告",
            command=self._on_generate_click,
            bootstyle="success",
            width=30,
        )
        self.generate_btn.pack(pady=10)

        # Log area
        log_frame = ttk.LabelFrame(self, text="執行記錄")
        log_frame.pack(fill=BOTH, expand=True, padx=10, pady=(5, 10))

        self.log_text = ttk.Text(log_frame, height=8, state="disabled", wrap="word")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.pack(fill=BOTH, expand=True)

    # ── Canvas preview drawing ──────────────────────────────

    def _on_layout_changed(self):
        self._draw_preview()

    def _draw_preview(self):
        c = self.preview_canvas
        c.delete("all")
        layout = self.layout_var.get()
        tasks = self._sample_tasks[:2]  # Show 2 sample tasks

        if layout == "span":
            self._draw_span_preview(c, tasks)
        elif layout == "single":
            self._draw_single_preview(c, tasks)
        else:
            self._draw_grid_preview(c, tasks)

    def _draw_span_preview(self, c, tasks):
        """Draw span layout: merged header + 2-col blank row."""
        x0, y = 20, 15
        w = 300
        half = w // 2
        header_h = 28
        photo_h = 45

        for task in tasks:
            # Merged header row
            c.create_rectangle(x0, y, x0 + w, y + header_h, outline="#333")
            c.create_text(x0 + w // 2, y + header_h // 2, text=task,
                          font=("", 9), fill="#333")
            y += header_h

            # Photo row — 2 cells
            c.create_rectangle(x0, y, x0 + half, y + photo_h, outline="#999")
            c.create_rectangle(x0 + half, y, x0 + w, y + photo_h, outline="#999")
            c.create_text(x0 + half // 2, y + photo_h // 2,
                          text="(照片)", font=("", 8), fill="#bbb")
            c.create_text(x0 + half + half // 2, y + photo_h // 2,
                          text="(照片)", font=("", 8), fill="#bbb")
            y += photo_h + 2

    def _draw_single_preview(self, c, tasks):
        """Draw single column layout."""
        x0, y = 20, 15
        w = 300
        content_h = 28
        photo_h = 45

        for task in tasks:
            # Content row
            c.create_rectangle(x0, y, x0 + w, y + content_h, outline="#333")
            c.create_text(x0 + 10, y + content_h // 2, text=task,
                          font=("", 9), fill="#333", anchor=W)
            y += content_h

            # Photo row
            c.create_rectangle(x0, y, x0 + w, y + photo_h, outline="#999")
            c.create_text(x0 + w // 2, y + photo_h // 2,
                          text="(照片)", font=("", 8), fill="#bbb")
            y += photo_h + 2

    def _draw_grid_preview(self, c, tasks):
        """Draw 2-column grid layout."""
        x0, y = 20, 15
        w = 300
        half = w // 2
        content_h = 28
        blank_h = 45

        # Content row
        c.create_rectangle(x0, y, x0 + half, y + content_h, outline="#333")
        c.create_rectangle(x0 + half, y, x0 + w, y + content_h, outline="#333")
        if len(tasks) > 0:
            c.create_text(x0 + 8, y + content_h // 2, text=tasks[0],
                          font=("", 9), fill="#333", anchor=W)
        if len(tasks) > 1:
            c.create_text(x0 + half + 8, y + content_h // 2, text=tasks[1],
                          font=("", 9), fill="#333", anchor=W)
        y += content_h

        # Blank row
        c.create_rectangle(x0, y, x0 + half, y + blank_h, outline="#999")
        c.create_rectangle(x0 + half, y, x0 + w, y + blank_h, outline="#999")
        c.create_text(x0 + half // 2, y + blank_h // 2,
                      text="(照片)", font=("", 8), fill="#bbb")
        c.create_text(x0 + half + half // 2, y + blank_h // 2,
                      text="(照片)", font=("", 8), fill="#bbb")

    def set_sample_tasks(self, tasks: list):
        """Update sample tasks for preview (first 2-3 items)."""
        self._sample_tasks = [t[1] for t in tasks[:3]] if tasks else ["Task A", "Task B"]
        self._draw_preview()

    # ── Public API ──────────────────────────────────────────

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
            self.generate_btn.config(state="disabled", text="產生中...")
        else:
            self.generate_btn.config(state="normal", text="產生報告")
