import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
from datetime import datetime
import openpyxl
import threading

from core.weather_service import WeatherService


class ConfigPanel(ttk.Frame):
    def __init__(self, parent, settings, **kwargs):
        super().__init__(parent, **kwargs)
        self.settings = settings
        self.on_excel_loaded = None  # callback: (path, sheet_name) -> None
        self.on_date_changed = None  # callback: () -> None

        self._weather_data    = None   # cached WeatherData
        self._search_results  = []     # list[CityResult] from last search

        self._build_ui()
        self._restore_from_settings()

    # ──────────────────────────────────────────────────────────────────────
    # UI construction
    # ──────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        PF = dict(padx=14, pady=(8, 4))   # pack kwargs for LabelFrames
        GX = dict(padx=(0, 8), pady=5)    # grid x-padding for widgets

        # === File Paths ===
        file_frame = ttk.LabelFrame(self, text=" 📁  檔案路徑")
        file_frame.pack(fill=X, **PF)

        ttk.Label(file_frame, text="Excel 檔案:").grid(row=0, column=0, sticky=W, pady=4)
        self.excel_var = ttk.StringVar()
        ttk.Entry(file_frame, textvariable=self.excel_var).grid(
            row=0, column=1, sticky=EW, padx=8, pady=4)
        ttk.Button(file_frame, text="瀏覽…", command=self._browse_excel,
                   bootstyle="secondary-outline", width=8).grid(row=0, column=2, pady=4)

        ttk.Label(file_frame, text="輸出資料夾:").grid(row=1, column=0, sticky=W, pady=4)
        self.output_var = ttk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_var).grid(
            row=1, column=1, sticky=EW, padx=8, pady=4)
        ttk.Button(file_frame, text="瀏覽…", command=self._browse_output,
                   bootstyle="secondary-outline", width=8).grid(row=1, column=2, pady=4)
        file_frame.columnconfigure(1, weight=1)

        # === Date Configuration ===
        date_frame = ttk.LabelFrame(self, text=" 📅  日期設定")
        date_frame.pack(fill=X, **PF)

        ttk.Label(date_frame, text="專案起始日:").grid(row=0, column=0, sticky=W, **GX)
        self.start_date_entry = ttk.DateEntry(date_frame, dateformat="%Y-%m-%d", width=14)
        self.start_date_entry.grid(row=0, column=1, sticky=W, padx=8, pady=5)
        ttk.Label(date_frame, text="用於計算 D 天數", bootstyle="secondary").grid(
            row=0, column=2, sticky=W, padx=4)

        ttk.Label(date_frame, text="報告日期:").grid(row=1, column=0, sticky=W, **GX)
        self.report_date_entry = ttk.DateEntry(date_frame, dateformat="%Y-%m-%d", width=14)
        self.report_date_entry.grid(row=1, column=1, sticky=W, padx=8, pady=5)
        ttk.Button(date_frame, text="今天", command=self._set_today,
                   bootstyle="info-outline", width=6).grid(row=1, column=2, sticky=W, padx=4)

        ttk.Label(date_frame, text="計算結果:").grid(row=2, column=0, sticky=W, **GX)
        self.computed_label = ttk.Label(
            date_frame, text="", bootstyle="primary",
            font=("", 11, "bold"),
        )
        self.computed_label.grid(row=2, column=1, columnspan=2, sticky=W, padx=8, pady=5)

        self.start_date_entry.entry.bind("<<DateEntrySelected>>", lambda e: self._update_computed())
        self.report_date_entry.entry.bind("<<DateEntrySelected>>", lambda e: self._update_computed())
        self.start_date_entry.entry.bind("<FocusOut>", lambda e: self._update_computed())
        self.report_date_entry.entry.bind("<FocusOut>", lambda e: self._update_computed())

        # === Excel Configuration ===
        excel_frame = ttk.LabelFrame(self, text=" 📊  Excel 設定")
        excel_frame.pack(fill=X, **PF)

        ttk.Label(excel_frame, text="工作表名稱:").grid(row=0, column=0, sticky=W, **GX)
        self.sheet_var = ttk.StringVar(value="總表")
        self.sheet_combo = ttk.Combobox(excel_frame, textvariable=self.sheet_var,
                                        width=22, state="readonly")
        self.sheet_combo.grid(row=0, column=1, sticky=W, padx=8, pady=5)
        ttk.Button(excel_frame, text="載入 Excel", command=self._load_excel,
                   bootstyle="primary", width=12).grid(row=0, column=2, padx=4)

        # === Department Selection ===
        dept_frame = ttk.LabelFrame(self, text=" 🚢  部門選擇")
        dept_frame.pack(fill=X, **PF)

        self.dept_var = ttk.StringVar(value=self.settings.department or "engine")
        radio_frame = ttk.Frame(dept_frame)
        radio_frame.grid(row=0, column=0, columnspan=4, sticky=W, pady=(2, 6))
        for idx, (label, val) in enumerate([
            ("機艙  (Eeg Dept.)", "engine"),
            ("甲板  (Deck Dept.)", "deck"),
            ("不選擇", "none"),
        ]):
            ttk.Radiobutton(
                radio_frame, text=label, value=val,
                variable=self.dept_var, command=self._on_dept_changed,
            ).pack(side=LEFT, padx=(0, 16))

        ttk.Label(dept_frame, text="檔名前綴:").grid(row=1, column=0, sticky=W, pady=(0, 6))
        self.prefix_var = ttk.StringVar(value=self.settings.filename_prefix or "SY-")
        ttk.Entry(dept_frame, textvariable=self.prefix_var, width=14).grid(
            row=1, column=1, sticky=W, padx=8, pady=(0, 6))
        ttk.Label(dept_frame, text="例: SY-、MV-、留空", bootstyle="secondary").grid(
            row=1, column=2, sticky=W, padx=4, pady=(0, 6))
        self.prefix_var.trace_add("write", lambda *_: self._on_dept_changed())

        # === Weather Settings ===
        self._build_weather_ui()

    def _build_weather_ui(self):
        wf = ttk.LabelFrame(self, text=" 🌤  天氣資訊")
        wf.pack(fill=X, padx=14, pady=(8, 12))
        wf.columnconfigure(1, weight=1)

        # ── Enable toggle ──────────────────────────────────────────────
        self.weather_enabled_var = ttk.BooleanVar(value=self.settings.show_weather)
        ttk.Checkbutton(
            wf, text="在報告中顯示天氣資訊",
            variable=self.weather_enabled_var,
            bootstyle="success",
        ).grid(row=0, column=0, columnspan=3, sticky=W, pady=(4, 6))

        # ── 地點搜尋（單一框） ─────────────────────────────────────────
        ttk.Label(wf, text="地點:").grid(row=1, column=0, sticky=W, padx=(4, 0), pady=5)
        self.city_search_var = ttk.StringVar(
            value=getattr(self.settings, "weather_city", "") or "")
        self.city_entry = ttk.Entry(wf, textvariable=self.city_search_var)
        self.city_entry.grid(row=1, column=1, sticky=EW, padx=8, pady=5)
        self.city_entry.bind("<Return>", lambda e: self._search_city())
        ttk.Button(wf, text="搜尋", command=self._search_city,
                   bootstyle="info-outline", width=6).grid(row=1, column=2, padx=(0, 8), pady=5)

        # ── Combo for selecting from results ──────────────────────────
        self.city_result_var = ttk.StringVar()
        self.city_combo = ttk.Combobox(wf, textvariable=self.city_result_var,
                                       state="readonly")
        self.city_combo.grid(row=2, column=0, columnspan=3, sticky=EW,
                             padx=8, pady=(0, 4))
        self.city_combo.grid_remove()   # hidden until results arrive
        self.city_combo.bind("<<ComboboxSelected>>", self._on_result_selected)

        # ── Status labels ──────────────────────────────────────────────
        self.selected_label = ttk.Label(wf, text="尚未選擇地點", bootstyle="secondary")
        self.selected_label.grid(row=3, column=0, columnspan=3, sticky=W,
                                 padx=8, pady=(2, 0))

        self.weather_result_label = ttk.Label(wf, text="", bootstyle="success",
                                              font=("", 10))
        self.weather_result_label.grid(row=4, column=0, columnspan=3, sticky=W,
                                       padx=8, pady=(0, 6))

    # ──────────────────────────────────────────────────────────────────────
    # Weather: search logic
    # ──────────────────────────────────────────────────────────────────────

    def _search_city(self):
        query = self.city_search_var.get().strip()
        if not query:
            return
        self._search_results = []
        self._weather_data   = None
        self.weather_result_label.config(text="")
        self.selected_label.config(text="搜尋中…", bootstyle="warning")
        self.city_combo.grid_remove()

        def _worker():
            try:
                results = WeatherService.search_city(query)
            except Exception:
                results = []
            self.after(0, lambda: self._on_search_done(results))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_search_done(self, results: list):
        self._search_results = results
        if not results:
            self.selected_label.config(text="找不到符合的地點", bootstyle="danger")
            self.city_combo.grid_remove()
            return
        names = [r.display_name() for r in results]
        self.city_combo["values"] = names
        self.city_combo.set(names[0])
        if len(results) > 1:
            # show picker only when multiple results
            self.city_combo.grid()
        else:
            self.city_combo.grid_remove()
        self._confirm_result(0)

    def _on_result_selected(self, event=None):
        idx = self.city_combo.current()
        if 0 <= idx < len(self._search_results):
            self._confirm_result(idx)

    def _confirm_result(self, idx: int):
        if idx >= len(self._search_results):
            return
        self._fetch_weather_for(self._search_results[idx])

    def _fetch_weather_for(self, location):
        lat, lon = location.latitude, location.longitude
        ns = "N" if lat >= 0 else "S"
        ew = "E" if lon >= 0 else "W"
        coord = f"{abs(lat):.2f}°{ns}, {abs(lon):.2f}°{ew}"
        self.selected_label.config(
            text=f"✓  {location.display_name()}  ({coord})",
            bootstyle="success",
        )
        self.weather_result_label.config(text="正在取得天氣資料…", bootstyle="warning")

        def _worker():
            try:
                report_date = self.get_report_date()
                w = WeatherService.fetch_weather(lat, lon, report_date)
            except Exception:
                w = None
            if w:
                self._weather_data = w
                self.after(0, lambda: self.weather_result_label.config(
                    text=f"⛅  {w.format_line()}", bootstyle="success",
                ))
            else:
                self.after(0, lambda: self.weather_result_label.config(
                    text="無法取得天氣資料（此日期可能超出範圍）", bootstyle="danger",
                ))

        threading.Thread(target=_worker, daemon=True).start()

    # ──────────────────────────────────────────────────────────────────────
    # Department / prefix
    # ──────────────────────────────────────────────────────────────────────

    def _on_dept_changed(self):
        if self.on_date_changed:
            self.on_date_changed()

    # ──────────────────────────────────────────────────────────────────────
    # Settings restore / file helpers
    # ──────────────────────────────────────────────────────────────────────

    def _restore_from_settings(self):
        if self.settings.excel_path:
            self.excel_var.set(self.settings.excel_path)
        if self.settings.output_dir:
            self.output_var.set(self.settings.output_dir)
        if self.settings.sheet_name:
            self.sheet_var.set(self.settings.sheet_name)

        try:
            sd = datetime.strptime(self.settings.start_date, "%Y-%m-%d")
            self.start_date_entry.entry.delete(0, "end")
            self.start_date_entry.entry.insert(0, sd.strftime("%Y-%m-%d"))
        except (ValueError, TypeError):
            pass

        self._set_today()
        self._update_computed()

        if self.settings.excel_path:
            self._populate_sheet_names(self.settings.excel_path)

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_var.set(path)
            self.settings.excel_path = path
            self._populate_sheet_names(path)

    def _browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)
            self.settings.output_dir = path

    def _populate_sheet_names(self, path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            names = wb.sheetnames
            wb.close()
            self.sheet_combo["values"] = names
            if self.sheet_var.get() not in names and names:
                self.sheet_var.set("總表" if "總表" in names else names[0])
        except Exception:
            pass

    def _set_today(self):
        today = datetime.today()
        self.report_date_entry.entry.delete(0, "end")
        self.report_date_entry.entry.insert(0, today.strftime("%Y-%m-%d"))
        self._update_computed()

    def _update_computed(self):
        try:
            start  = datetime.strptime(self.start_date_entry.entry.get(), "%Y-%m-%d")
            report = datetime.strptime(self.report_date_entry.entry.get(), "%Y-%m-%d")
            day_n  = (report - start).days
            self.computed_label.config(
                text=f"D{day_n}  /  {report.strftime('%Y%m%d')}"
            )
        except (ValueError, TypeError):
            self.computed_label.config(text="(日期格式錯誤)")

        if self.on_date_changed:
            self.on_date_changed()

    def _load_excel(self):
        path  = self.excel_var.get()
        sheet = self.sheet_var.get()
        if not path:
            from ttkbootstrap.dialogs import Messagebox
            Messagebox.show_error("請先選擇 Excel 檔案", title="錯誤")
            return
        self.settings.excel_path = path
        self.settings.sheet_name = sheet
        if self.on_excel_loaded:
            self.on_excel_loaded(path, sheet)

    # ──────────────────────────────────────────────────────────────────────
    # Public getters
    # ──────────────────────────────────────────────────────────────────────

    def get_start_date(self) -> datetime:
        return datetime.strptime(self.start_date_entry.entry.get(), "%Y-%m-%d")

    def get_report_date(self) -> datetime:
        return datetime.strptime(self.report_date_entry.entry.get(), "%Y-%m-%d")

    def get_excel_path(self) -> str:
        return self.excel_var.get()

    def get_output_dir(self) -> str:
        return self.output_var.get()

    def get_sheet_name(self) -> str:
        return self.sheet_var.get()

    def get_department(self) -> str:
        return self.dept_var.get()

    def get_filename_prefix(self) -> str:
        return self.prefix_var.get()

    def get_weather_city(self) -> str:
        # If a result was confirmed, use its canonical name; else use typed text
        if self._search_results:
            idx = self.city_combo.current()
            if 0 <= idx < len(self._search_results):
                return self._search_results[idx].name
        return self.city_search_var.get()

    def get_weather_text(self) -> str:
        """Return the weather line to embed in the report, or '' if disabled / not fetched."""
        if self.weather_enabled_var.get() and self._weather_data:
            return self._weather_data.format_line()
        return ""

    # ──────────────────────────────────────────────────────────────────────
    # Persist settings
    # ──────────────────────────────────────────────────────────────────────

    def save_to_settings(self):
        self.settings.excel_path      = self.excel_var.get()
        self.settings.output_dir      = self.output_var.get()
        self.settings.sheet_name      = self.sheet_var.get()
        self.settings.department      = self.dept_var.get()
        self.settings.filename_prefix = self.prefix_var.get()
        self.settings.show_weather    = self.weather_enabled_var.get()
        self.settings.weather_city    = self.city_search_var.get()
        try:
            self.settings.start_date = self.start_date_entry.entry.get()
        except (ValueError, TypeError):
            pass
