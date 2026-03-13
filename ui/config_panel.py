import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
from datetime import datetime
import openpyxl
import threading

from core.weather_service import WeatherService, COMMON_COUNTRIES


class ConfigPanel(ttk.Frame):
    def __init__(self, parent, settings, **kwargs):
        super().__init__(parent, **kwargs)
        self.settings = settings
        self.on_excel_loaded = None  # callback: (path, sheet_name) -> None
        self.on_date_changed = None  # callback: () -> None

        self._weather_data = None    # cached WeatherData
        self._city_results  = []     # list[CityResult] from last city search
        self._town_results  = []     # list[CityResult] from last town search
        self._selected_city = None   # currently confirmed CityResult (city level)
        self._selected_town = None   # currently confirmed CityResult (town level)

        # Build country display strings and code lookup
        self._country_display = [f"{name} ({code})" for code, name in COMMON_COUNTRIES]
        self._code_to_display  = {code: f"{name} ({code})" for code, name in COMMON_COUNTRIES}
        self._display_to_code  = {f"{name} ({code})": code for code, name in COMMON_COUNTRIES}

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
        wf = ttk.LabelFrame(self, text="天氣資訊")
        wf.pack(fill=X, padx=10, pady=5)
        wf.columnconfigure(1, weight=1)

        PX = (0, 10)   # right-side padding

        # ── Enable toggle ──────────────────────────────────────────────
        self.weather_enabled_var = ttk.BooleanVar(value=self.settings.show_weather)
        ttk.Checkbutton(
            wf, text="在報告中顯示天氣",
            variable=self.weather_enabled_var,
        ).grid(row=0, column=0, columnspan=3, sticky=W, padx=10, pady=(6, 4))

        def _lbl(row, text):
            ttk.Label(wf, text=text).grid(row=row, column=0, sticky=W, padx=10, pady=3)

        def _search_row(parent, var, search_cmd, extra_label=None):
            """Return a frame with: Combobox[weight=1]  [🔍 button]  [opt label]"""
            f = ttk.Frame(parent)
            f.columnconfigure(0, weight=1)
            combo = ttk.Combobox(f, textvariable=var, state="normal", width=38)
            combo.grid(row=0, column=0, sticky=EW, padx=(0, 4))
            combo.bind("<Return>", lambda e: search_cmd())
            ttk.Button(f, text="🔍", command=search_cmd,
                       bootstyle="info-outline", width=3).grid(row=0, column=1)
            if extra_label:
                ttk.Label(f, text=extra_label, bootstyle="secondary").grid(
                    row=0, column=2, padx=(6, 0))
            return f, combo

        # ── 國家 ── readonly combobox (pre-populated list) ─────────────
        _lbl(1, "國家:")
        self.country_var = ttk.StringVar()
        self.country_combo = ttk.Combobox(
            wf, textvariable=self.country_var,
            values=self._country_display,
            state="readonly", width=32,
        )
        self.country_combo.grid(row=1, column=1, columnspan=2, sticky=W, padx=PX, pady=3)
        self.country_combo.bind("<<ComboboxSelected>>", self._on_country_changed)
        saved_code = getattr(self.settings, "weather_country", "TW") or "TW"
        self.country_var.set(self._code_to_display.get(saved_code, self._country_display[0]))

        # ── 城市 ── editable combobox + search button ──────────────────
        _lbl(2, "城市:")
        self.city_search_var = ttk.StringVar(
            value=getattr(self.settings, "weather_city", "") or "")
        city_frame, self.city_combo = _search_row(wf, self.city_search_var, self._search_city)
        city_frame.grid(row=2, column=1, columnspan=2, sticky=EW, padx=PX, pady=3)
        self.city_combo.bind("<<ComboboxSelected>>", self._on_city_selected)

        # ── 鄉鎮 ── editable combobox + search button (選填) ───────────
        _lbl(3, "鄉鎮:")
        self.town_search_var = ttk.StringVar(
            value=getattr(self.settings, "weather_town", "") or "")
        town_frame, self.town_combo = _search_row(
            wf, self.town_search_var, self._search_town, extra_label="(選填)")
        town_frame.grid(row=3, column=1, columnspan=2, sticky=EW, padx=PX, pady=3)
        self.town_combo.bind("<<ComboboxSelected>>", self._on_town_selected)

        # ── Status labels ───────────────────────────────────────────────
        self.selected_label = ttk.Label(wf, text="尚未選擇地點", bootstyle="secondary")
        self.selected_label.grid(row=4, column=0, columnspan=3, sticky=W, padx=10, pady=(4, 0))

        self.weather_result_label = ttk.Label(wf, text="", bootstyle="secondary")
        self.weather_result_label.grid(row=5, column=0, columnspan=3, sticky=W,
                                       padx=10, pady=(0, 6))

    # ──────────────────────────────────────────────────────────────────────
    # Weather: cascading search logic
    # ──────────────────────────────────────────────────────────────────────

    def _get_country_code(self) -> str:
        return self._display_to_code.get(self.country_var.get(), "")

    def _on_country_changed(self, event=None):
        """Clear city / town results when country changes."""
        self._city_results = []
        self._town_results = []
        self._selected_city = None
        self._selected_town = None
        self._weather_data  = None
        self.city_combo.set("")
        self.city_combo["values"] = []
        self.town_combo.set("")
        self.town_combo["values"] = []
        self.selected_label.config(text="已更換國家，請重新搜尋城市", bootstyle="warning")
        self.weather_result_label.config(text="")

    # ── shared helper ──────────────────────────────────────────────────

    def _do_search(self, query: str, combo: ttk.Combobox,
                   result_store: str, on_done):
        """Generic search: query API → fill combo → call on_done(results)."""
        if not query:
            return
        combo.set("搜尋中…")
        combo["values"] = []
        country_code = self._get_country_code()

        def _worker():
            try:
                results = WeatherService.search_city(query, country_code=country_code)
            except Exception:
                results = []
            self.after(0, lambda: self._fill_combo(combo, results, result_store, on_done))

        threading.Thread(target=_worker, daemon=True).start()

    def _fill_combo(self, combo: ttk.Combobox, results: list,
                    result_attr: str, on_first_select):
        """Populate combo with results, then open the dropdown."""
        setattr(self, result_attr, results)
        if not results:
            combo.set("（找不到符合的地點）")
            combo["values"] = ["（找不到符合的地點）"]
            self.selected_label.config(text="找不到符合的地點", bootstyle="danger")
            return
        names = [r.display_name() for r in results]
        combo["values"] = names
        combo.set(names[0])
        # Open the dropdown so the user can see all choices
        combo.event_generate("<Button-1>")
        combo.focus_set()
        combo.after(50, combo.event_generate, "<<ComboboxSelected>>")
        on_first_select(0)

    # ── 城市 ───────────────────────────────────────────────────────────

    def _search_city(self):
        query = self.city_search_var.get().strip()
        self._city_results  = []
        self._town_results  = []
        self._selected_city = None
        self._selected_town = None
        self._weather_data  = None
        self.town_combo.set("")
        self.town_combo["values"] = []
        self.weather_result_label.config(text="")
        self.selected_label.config(text="搜尋城市中…", bootstyle="warning")
        self._do_search(query, self.city_combo, "_city_results", self._confirm_city)

    def _on_city_selected(self, event=None):
        idx = self.city_combo.current()
        if 0 <= idx < len(self._city_results):
            self._confirm_city(idx)

    def _confirm_city(self, idx: int):
        if idx >= len(self._city_results):
            return
        city = self._city_results[idx]
        self._selected_city = city
        # Clear town when city changes
        self._selected_town = None
        self._town_results  = []
        self.town_combo.set("")
        self.town_combo["values"] = []
        self._fetch_weather_for(city)

    # ── 鄉鎮 ───────────────────────────────────────────────────────────

    def _search_town(self):
        query = self.town_search_var.get().strip()
        self._town_results  = []
        self._selected_town = None
        self.weather_result_label.config(text="")
        self._do_search(query, self.town_combo, "_town_results", self._confirm_town)

    def _on_town_selected(self, event=None):
        idx = self.town_combo.current()
        if 0 <= idx < len(self._town_results):
            self._confirm_town(idx)

    def _confirm_town(self, idx: int):
        if idx >= len(self._town_results):
            return
        town = self._town_results[idx]
        self._selected_town = town
        self._fetch_weather_for(town)

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
        self.settings.weather_country = self._get_country_code()
        self.settings.weather_city    = self.city_search_var.get()
        self.settings.weather_town    = self.town_search_var.get()
        try:
            self.settings.start_date = self.start_date_entry.entry.get()
        except (ValueError, TypeError):
            pass
