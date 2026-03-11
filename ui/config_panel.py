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
        self._weather_data = None    # cached WeatherData
        self._city_results = []      # list[CityResult] from last search
        self._selected_city = None   # currently selected CityResult

        self._build_ui()
        self._restore_from_settings()

    # ──────────────────────────────────────────────────────────────────────
    # UI construction
    # ──────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # === File Paths ===
        file_frame = ttk.LabelFrame(self, text="檔案路徑")
        file_frame.pack(fill=X, padx=10, pady=(10, 5))

        ttk.Label(file_frame, text="Excel 檔案:").grid(row=0, column=0, sticky=W, pady=2)
        self.excel_var = ttk.StringVar()
        ttk.Entry(file_frame, textvariable=self.excel_var, width=60).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(file_frame, text="瀏覽", command=self._browse_excel, bootstyle="outline").grid(row=0, column=2, pady=2)

        ttk.Label(file_frame, text="輸出資料夾:").grid(row=1, column=0, sticky=W, pady=2)
        self.output_var = ttk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_var, width=60).grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(file_frame, text="瀏覽", command=self._browse_output, bootstyle="outline").grid(row=1, column=2, pady=2)
        file_frame.columnconfigure(1, weight=1)

        # === Date Configuration ===
        date_frame = ttk.LabelFrame(self, text="日期設定")
        date_frame.pack(fill=X, padx=10, pady=5)

        ttk.Label(date_frame, text="專案起始日:").grid(row=0, column=0, sticky=W, pady=2)
        self.start_date_entry = ttk.DateEntry(date_frame, dateformat="%Y-%m-%d", width=15)
        self.start_date_entry.grid(row=0, column=1, sticky=W, padx=5, pady=2)
        ttk.Label(date_frame, text="(用於計算 D 天數)", bootstyle="secondary").grid(row=0, column=2, sticky=W, padx=5)

        ttk.Label(date_frame, text="報告日期:").grid(row=1, column=0, sticky=W, pady=2)
        self.report_date_entry = ttk.DateEntry(date_frame, dateformat="%Y-%m-%d", width=15)
        self.report_date_entry.grid(row=1, column=1, sticky=W, padx=5, pady=2)
        ttk.Button(date_frame, text="今天", command=self._set_today, bootstyle="info-outline").grid(row=1, column=2, sticky=W, padx=5)

        ttk.Label(date_frame, text="計算結果:").grid(row=2, column=0, sticky=W, pady=2)
        self.computed_label = ttk.Label(date_frame, text="", bootstyle="success", font=("", 11, "bold"))
        self.computed_label.grid(row=2, column=1, columnspan=2, sticky=W, padx=5, pady=2)

        self.start_date_entry.entry.bind("<<DateEntrySelected>>", lambda e: self._update_computed())
        self.report_date_entry.entry.bind("<<DateEntrySelected>>", lambda e: self._update_computed())
        self.start_date_entry.entry.bind("<FocusOut>", lambda e: self._update_computed())
        self.report_date_entry.entry.bind("<FocusOut>", lambda e: self._update_computed())

        # === Excel Configuration ===
        excel_frame = ttk.LabelFrame(self, text="Excel 設定")
        excel_frame.pack(fill=X, padx=10, pady=5)

        ttk.Label(excel_frame, text="工作表名稱:").grid(row=0, column=0, sticky=W, pady=2)
        self.sheet_var = ttk.StringVar(value="總表")
        self.sheet_combo = ttk.Combobox(excel_frame, textvariable=self.sheet_var, width=20, state="readonly")
        self.sheet_combo.grid(row=0, column=1, sticky=W, padx=5, pady=2)
        ttk.Button(excel_frame, text="載入 Excel", command=self._load_excel, bootstyle="success").grid(row=0, column=2, padx=5)

        # === Department Selection ===
        dept_frame = ttk.LabelFrame(self, text="部門選擇")
        dept_frame.pack(fill=X, padx=10, pady=5)

        self.dept_var = ttk.StringVar(value=self.settings.department or "engine")
        for idx, (label, val) in enumerate([
            ("機艙 (Eeg Dept.)", "engine"),
            ("甲板 (Deck Dept.)", "deck"),
            ("不選擇", "none"),
        ]):
            ttk.Radiobutton(
                dept_frame, text=label, value=val,
                variable=self.dept_var, command=self._on_dept_changed,
            ).grid(row=0, column=idx, padx=10, pady=5, sticky=W)

        ttk.Label(dept_frame, text="檔名前綴:").grid(row=1, column=0, sticky=W, padx=10, pady=(0, 5))
        self.prefix_var = ttk.StringVar(value=self.settings.filename_prefix or "SY-")
        ttk.Entry(dept_frame, textvariable=self.prefix_var, width=15).grid(
            row=1, column=1, sticky=W, padx=5, pady=(0, 5))
        self.prefix_var.trace_add("write", lambda *_: self._on_dept_changed())

        # === Weather Settings ===
        self._build_weather_ui()

    def _build_weather_ui(self):
        weather_frame = ttk.LabelFrame(self, text="天氣資訊")
        weather_frame.pack(fill=X, padx=10, pady=5)
        weather_frame.columnconfigure(1, weight=1)

        # ── Enable toggle ──
        self.weather_enabled_var = ttk.BooleanVar(value=self.settings.show_weather)
        ttk.Checkbutton(
            weather_frame, text="在報告中顯示天氣",
            variable=self.weather_enabled_var,
        ).grid(row=0, column=0, columnspan=3, sticky=W, padx=10, pady=(6, 2))

        # ── Search row ──
        ttk.Label(weather_frame, text="地點搜尋:").grid(row=1, column=0, sticky=W, padx=10, pady=2)

        search_row = ttk.Frame(weather_frame)
        search_row.grid(row=1, column=1, columnspan=2, sticky=EW, padx=(0, 10), pady=2)
        search_row.columnconfigure(0, weight=1)

        self.city_search_var = ttk.StringVar(value=self.settings.weather_city or "Kaohsiung")
        self.city_entry = ttk.Entry(search_row, textvariable=self.city_search_var)
        self.city_entry.grid(row=0, column=0, sticky=EW, padx=(0, 6))
        self.city_entry.bind("<Return>", lambda e: self._search_city())

        ttk.Button(
            search_row, text="🔍 搜尋", command=self._search_city,
            bootstyle="info-outline", width=8,
        ).grid(row=0, column=1)

        # ── Results listbox ──
        lb_frame = ttk.Frame(weather_frame)
        lb_frame.grid(row=2, column=0, columnspan=3, sticky=EW, padx=10, pady=(2, 0))
        lb_frame.columnconfigure(0, weight=1)

        self.city_listbox = tk.Listbox(
            lb_frame, height=4, selectmode=tk.SINGLE,
            font=("Consolas", 9),
            bg="#f8f8f8", selectbackground="#0d6efd", selectforeground="white",
            activestyle="none", relief="flat", borderwidth=1,
            highlightthickness=1, highlightbackground="#cccccc",
        )
        lb_scroll = ttk.Scrollbar(lb_frame, orient=VERTICAL, command=self.city_listbox.yview)
        self.city_listbox.configure(yscrollcommand=lb_scroll.set)
        lb_scroll.grid(row=0, column=1, sticky=NS)
        self.city_listbox.grid(row=0, column=0, sticky=EW)
        self.city_listbox.bind("<<ListboxSelect>>", self._on_city_selected)

        # ── Selected + weather labels ──
        self.selected_label = ttk.Label(
            weather_frame, text="尚未選擇地點", bootstyle="secondary",
        )
        self.selected_label.grid(row=3, column=0, columnspan=3, sticky=W, padx=10, pady=(4, 0))

        self.weather_result_label = ttk.Label(
            weather_frame, text="", bootstyle="secondary",
        )
        self.weather_result_label.grid(row=4, column=0, columnspan=3, sticky=W, padx=10, pady=(0, 6))

    # ──────────────────────────────────────────────────────────────────────
    # Weather search / selection logic
    # ──────────────────────────────────────────────────────────────────────

    def _search_city(self):
        query = self.city_search_var.get().strip()
        if not query:
            return
        self.city_listbox.delete(0, tk.END)
        self.city_listbox.insert(tk.END, "搜尋中…")
        self.city_listbox.config(state=tk.DISABLED)

        def _worker():
            try:
                results = WeatherService.search_city(query)
            except Exception:
                results = []
            self.after(0, lambda: self._populate_city_results(results))

        threading.Thread(target=_worker, daemon=True).start()

    def _populate_city_results(self, results):
        self.city_listbox.config(state=tk.NORMAL)
        self.city_listbox.delete(0, tk.END)
        self._city_results = results
        if not results:
            self.city_listbox.insert(tk.END, "  （找不到符合的地點）")
            return
        for city in results:
            self.city_listbox.insert(tk.END, f"  {city.display_name()}")

    def _on_city_selected(self, event=None):
        sel = self.city_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self._city_results):
            return
        city = self._city_results[idx]
        self._selected_city = city
        # Keep the search field in sync (used by save_to_settings)
        self.city_search_var.set(city.name)

        coord = f"{abs(city.latitude):.2f}°{'N' if city.latitude >= 0 else 'S'}, " \
                f"{abs(city.longitude):.2f}°{'E' if city.longitude >= 0 else 'W'}"
        self.selected_label.config(
            text=f"✓  {city.display_name()}  ({coord})",
            bootstyle="success",
        )
        self.weather_result_label.config(text="正在取得天氣資料…", bootstyle="warning")

        def _worker():
            try:
                report_date = self.get_report_date()
                w = WeatherService.fetch_weather(city.latitude, city.longitude, report_date)
            except Exception:
                w = None
            if w:
                self._weather_data = w
                self.after(0, lambda: self.weather_result_label.config(
                    text=f"⛅  {w.format_line()}",
                    bootstyle="success",
                ))
            else:
                self.after(0, lambda: self.weather_result_label.config(
                    text="無法取得天氣資料（此日期可能超出範圍）",
                    bootstyle="danger",
                ))

        threading.Thread(target=_worker, daemon=True).start()

    # ──────────────────────────────────────────────────────────────────────
    # Department / prefix
    # ──────────────────────────────────────────────────────────────────────

    def _on_dept_changed(self):
        if self.on_date_changed:
            self.on_date_changed()

    # ──────────────────────────────────────────────────────────────────────
    # Settings restore
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

    # ──────────────────────────────────────────────────────────────────────
    # File / date helpers
    # ──────────────────────────────────────────────────────────────────────

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
            start = datetime.strptime(self.start_date_entry.entry.get(), "%Y-%m-%d")
            report = datetime.strptime(self.report_date_entry.entry.get(), "%Y-%m-%d")
            day_number = (report - start).days
            self.computed_label.config(
                text=f"D{day_number}  /  {report.strftime('%Y%m%d')}"
            )
        except (ValueError, TypeError):
            self.computed_label.config(text="(日期格式錯誤)")

        if self.on_date_changed:
            self.on_date_changed()

    def _load_excel(self):
        path = self.excel_var.get()
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
        return self.city_search_var.get()

    def get_weather_text(self) -> str:
        """Return weather line for report, or empty if disabled / not fetched."""
        if self.weather_enabled_var.get() and self._weather_data:
            return self._weather_data.format_line()
        return ""

    # ──────────────────────────────────────────────────────────────────────
    # Persist settings
    # ──────────────────────────────────────────────────────────────────────

    def save_to_settings(self):
        self.settings.excel_path = self.excel_var.get()
        self.settings.output_dir = self.output_var.get()
        self.settings.sheet_name = self.sheet_var.get()
        self.settings.department = self.dept_var.get()
        self.settings.filename_prefix = self.prefix_var.get()
        self.settings.show_weather = self.weather_enabled_var.get()
        self.settings.weather_city = self.city_search_var.get()
        try:
            self.settings.start_date = self.start_date_entry.entry.get()
        except (ValueError, TypeError):
            pass
