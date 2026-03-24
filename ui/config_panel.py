import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
from tkinter import filedialog
from datetime import datetime
import openpyxl
import threading

from core.weather_service import WeatherService
from ui.constants import (
    FIELD_GAP_X,
    FIELD_GAP_Y,
    LABEL_INPUT_GAP,
    SHORT_INPUT_W,
    MEDIUM_INPUT_W,
    BTN_BROWSE,
    BTN_ACTION,
    BTN_PRIMARY,
)
from ui.widgets import make_page, make_section, make_stat_card


class ConfigPanel(ttk.Frame):
    def __init__(self, parent, settings, **kwargs):
        super().__init__(parent, **kwargs)
        self.settings = settings
        self.on_excel_loaded = None
        self.on_date_changed = None
        self._weather_data = None
        self._search_results = []

        self._build_ui()
        self._restore_from_settings()

    def _build_ui(self):
        sf = ScrolledFrame(self, autohide=True)
        sf.pack(fill=BOTH, expand=True)
        page = make_page(sf, padding=(20, 18))
        page.pack(fill=BOTH, expand=True)
        self._sf_container = page

        self._build_source_section(page)
        self._build_date_section(page)
        self._build_report_section(page)
        self._build_weather_section(page)

        self.start_date_entry.entry.bind("<<DateEntrySelected>>", lambda e: self._update_computed())
        self.report_date_entry.entry.bind("<<DateEntrySelected>>", lambda e: self._update_computed())
        self.start_date_entry.entry.bind("<FocusOut>", lambda e: self._update_computed())
        self.report_date_entry.entry.bind("<FocusOut>", lambda e: self._update_computed())

    def _build_source_section(self, parent):
        body = make_section(
            parent,
            "資料來源",
            "先指定來源資料與輸出位置，再選擇工作表載入。",
            compact=True,
        )
        body.columnconfigure(0, weight=1)

        self.excel_var = ttk.StringVar()
        self.output_var = ttk.StringVar()
        self.sheet_var = ttk.StringVar(value="總表")

        self._make_path_row(body, 0, "Excel 檔案", self.excel_var, self._browse_excel)
        self._make_path_row(body, 1, "輸出資料夾", self.output_var, self._browse_output)

        row = ttk.Frame(body, style="Surface.TFrame")
        row.grid(row=2, column=0, sticky=EW, pady=(FIELD_GAP_Y + 2, 0))
        row.columnconfigure(1, weight=1)

        ttk.Label(row, text="工作表", style="FieldLabel.TLabel").grid(row=0, column=0, sticky=W)
        self.sheet_combo = ttk.Combobox(
            row,
            textvariable=self.sheet_var,
            width=MEDIUM_INPUT_W,
            state="readonly",
        )
        self.sheet_combo.grid(row=0, column=1, sticky=W, padx=(LABEL_INPUT_GAP, FIELD_GAP_X))
        ttk.Button(
            row,
            text="載入 Excel",
            command=self._load_excel,
            bootstyle=BTN_PRIMARY,
            width=12,
        ).grid(row=0, column=2, sticky=W)
        ttk.Label(
            row,
            text="確認工作表後載入，資料會同步到 Excel 編輯器與報告預覽。",
            style="Hint.TLabel",
        ).grid(row=1, column=1, columnspan=2, sticky=W, pady=(6, 0))

    def _build_date_section(self, parent):
        body = make_section(
            parent,
            "日期與報告資訊",
            "設定報告日期與基準日，並即時查看 D-Day 摘要。",
            compact=True,
        )
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)

        left = ttk.Frame(body, style="Surface.TFrame")
        left.grid(row=0, column=0, sticky=NSEW, padx=(0, 16))
        left.columnconfigure(1, weight=1)

        ttk.Label(left, text="起始日", style="FieldLabel.TLabel").grid(row=0, column=0, sticky=W, pady=(0, FIELD_GAP_Y))
        self.start_date_entry = ttk.DateEntry(left, dateformat="%Y-%m-%d", width=SHORT_INPUT_W)
        self.start_date_entry.grid(row=0, column=1, sticky=W, padx=(LABEL_INPUT_GAP, 0), pady=(0, FIELD_GAP_Y))
        ttk.Label(left, text="用於計算 D 天數", style="Hint.TLabel").grid(row=1, column=1, sticky=W, pady=(0, FIELD_GAP_Y))

        ttk.Label(left, text="報告日期", style="FieldLabel.TLabel").grid(row=2, column=0, sticky=W)
        self.report_date_entry = ttk.DateEntry(left, dateformat="%Y-%m-%d", width=SHORT_INPUT_W)
        self.report_date_entry.grid(row=2, column=1, sticky=W, padx=(LABEL_INPUT_GAP, 0))

        right = ttk.Frame(body, style="Surface.TFrame")
        right.grid(row=0, column=1, sticky=NSEW)
        ttk.Button(right, text="今天", command=self._set_today, bootstyle=BTN_ACTION, width=8).pack(anchor=W)
        ttk.Label(
            right,
            text="快速帶入今天日期。",
            style="Hint.TLabel",
            justify=LEFT,
            wraplength=260,
        ).pack(anchor=W, pady=(8, 12))

        card, self.computed_label, self.dday_note_label = make_stat_card(
            right,
            "報告時間摘要",
            "",
            "",
            style="AccentCard.TFrame",
        )
        card.pack(fill=X)

    def _build_report_section(self, parent):
        body = make_section(
            parent,
            "報告設定",
            "設定部門與檔名前綴。",
            compact=True,
        )
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        self.dept_var = ttk.StringVar(value=self.settings.department or "engine")
        dept_col = ttk.Frame(body, style="Surface.TFrame")
        dept_col.grid(row=0, column=0, sticky=NW, padx=(0, 20))
        ttk.Label(dept_col, text="部門", style="FieldLabel.TLabel").pack(anchor=W)
        radio_row = ttk.Frame(dept_col, style="Surface.TFrame")
        radio_row.pack(anchor=W, pady=(8, 0))
        for idx, (label, value) in enumerate([
            ("機艙 (Eng Dept.)", "engine"),
            ("甲板 (Deck Dept.)", "deck"),
            ("不選擇", "none"),
        ]):
            ttk.Radiobutton(
                radio_row,
                text=label,
                value=value,
                variable=self.dept_var,
                command=self._on_dept_changed,
            ).grid(row=0, column=idx, sticky=W, padx=(0, 18))

        prefix_col = ttk.Frame(body, style="Surface.TFrame")
        prefix_col.grid(row=0, column=1, sticky=NW)
        ttk.Label(prefix_col, text="檔名前綴", style="FieldLabel.TLabel").pack(anchor=W)
        self.prefix_var = ttk.StringVar(value=self.settings.filename_prefix or "SY-")
        ttk.Entry(prefix_col, textvariable=self.prefix_var, width=SHORT_INPUT_W).pack(anchor=W, pady=(8, 0))
        ttk.Label(prefix_col, text="例: SY-、MV-、留空", style="Hint.TLabel").pack(anchor=W, pady=(6, 0))
        self.prefix_var.trace_add("write", lambda *_: self._on_dept_changed())

    def _build_weather_section(self, parent):
        body = make_section(
            parent,
            "天氣資訊",
            "附加選項，需要時再設定即可。",
            style="MutedSurface.TFrame",
            compact=True,
        )
        body.columnconfigure(1, weight=1)

        self.weather_enabled_var = ttk.BooleanVar(value=self.settings.show_weather)
        ttk.Checkbutton(
            body,
            text="在報告中顯示天氣資訊",
            variable=self.weather_enabled_var,
            bootstyle="success",
        ).grid(row=0, column=0, columnspan=3, sticky=W, pady=(0, FIELD_GAP_Y))

        ttk.Label(body, text="地點", style="FieldLabel.TLabel").grid(row=1, column=0, sticky=W)
        self.city_search_var = ttk.StringVar(value=getattr(self.settings, "weather_city", "") or "")
        self.city_entry = ttk.Entry(body, textvariable=self.city_search_var)
        self.city_entry.grid(row=1, column=1, sticky=EW, padx=(LABEL_INPUT_GAP, FIELD_GAP_X))
        self.city_entry.bind("<Return>", lambda e: self._search_city())
        ttk.Button(body, text="搜尋", command=self._search_city, bootstyle=BTN_ACTION, width=6).grid(
            row=1, column=2, sticky=W
        )

        self.city_result_var = ttk.StringVar()
        self.city_combo = ttk.Combobox(body, textvariable=self.city_result_var, state="readonly")
        self.city_combo.grid(row=2, column=0, columnspan=3, sticky=EW, pady=(10, 0))
        self.city_combo.grid_remove()
        self.city_combo.bind("<<ComboboxSelected>>", self._on_result_selected)

        self.selected_label = ttk.Label(body, text="輸入地點名稱後點「搜尋」", style="Hint.TLabel")
        self.selected_label.grid(row=3, column=0, columnspan=3, sticky=W, pady=(10, 0))

        self.weather_result_label = ttk.Label(body, text="", style="WeatherResult.TLabel")
        self.weather_result_label.grid(row=4, column=0, columnspan=3, sticky=W, pady=(6, 0))

    @staticmethod
    def _format_weather_error(exc: Exception) -> str:
        message = str(exc).strip()
        return message if message else "天氣服務暫時無法使用"

    def _make_path_row(self, parent, row, label, variable, command):
        wrap = ttk.Frame(parent, style="Surface.TFrame")
        wrap.grid(row=row, column=0, sticky=EW, pady=(0, FIELD_GAP_Y))
        wrap.columnconfigure(1, weight=1)

        ttk.Label(wrap, text=label, style="FieldLabel.TLabel").grid(row=0, column=0, sticky=W)
        ttk.Entry(wrap, textvariable=variable).grid(
            row=0,
            column=1,
            sticky=EW,
            padx=(LABEL_INPUT_GAP, FIELD_GAP_X),
        )
        ttk.Button(wrap, text="瀏覽…", command=command, bootstyle=BTN_BROWSE, width=8).grid(
            row=0, column=2, sticky=E
        )

    def _search_city(self):
        query = self.city_search_var.get().strip()
        if not query:
            return
        self._search_results = []
        self._weather_data = None
        self.weather_result_label.config(text="", style="WeatherResult.TLabel")
        self.selected_label.config(text="搜尋中…", style="WarningHint.TLabel")
        self.city_combo.grid_remove()

        def _worker():
            try:
                results = WeatherService.search_city(query)
            except Exception as exc:
                self.after(0, lambda err=exc: self._on_search_failed(err))
                return
            self.after(0, lambda: self._on_search_done(results))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_search_failed(self, exc: Exception):
        self._search_results = []
        self._weather_data = None
        self.city_combo.grid_remove()
        self.selected_label.config(
            text=f"查詢失敗：{self._format_weather_error(exc)}",
            style="DangerHint.TLabel",
        )

    def _on_search_done(self, results: list):
        self._search_results = results
        if not results:
            self.selected_label.config(text="找不到符合的地點", style="DangerHint.TLabel")
            self.city_combo.grid_remove()
            return
        names = [result.display_name() for result in results]
        self.city_combo["values"] = names
        self.city_combo.set(names[0])
        if len(results) > 1:
            self.city_combo.grid()
        else:
            self.city_combo.grid_remove()
        self._confirm_result(0)

    def _on_result_selected(self, event=None):
        idx = self.city_combo.current()
        if 0 <= idx < len(self._search_results):
            self._confirm_result(idx)

    def _confirm_result(self, idx: int):
        if idx < len(self._search_results):
            self._fetch_weather_for(self._search_results[idx])

    def _fetch_weather_for(self, location):
        lat, lon = location.latitude, location.longitude
        ns = "N" if lat >= 0 else "S"
        ew = "E" if lon >= 0 else "W"
        coord = f"{abs(lat):.2f}°{ns}, {abs(lon):.2f}°{ew}"
        self.selected_label.config(
            text=f"✓  {location.display_name()}  ({coord})",
            style="SuccessHint.TLabel",
        )
        self.weather_result_label.config(text="正在取得天氣資料…", style="WarningHint.TLabel")

        def _worker():
            try:
                report_date = self.get_report_date()
                weather = WeatherService.fetch_weather(lat, lon, report_date)
            except Exception as exc:
                self.after(0, lambda err=exc: self._on_weather_fetch_failed(err))
                return
            if weather:
                self._weather_data = weather
                self.after(
                    0,
                    lambda: self.weather_result_label.config(
                        text=f"天氣資訊: {weather.format_line()}",
                        style="WeatherResult.TLabel",
                    ),
                )
            else:
                self._weather_data = None
                self.after(
                    0,
                    lambda: self.weather_result_label.config(
                        text="無法取得天氣資料（此日期可能超出範圍）",
                        style="DangerHint.TLabel",
                    ),
                )

        threading.Thread(target=_worker, daemon=True).start()

    def _on_weather_fetch_failed(self, exc: Exception):
        self._weather_data = None
        self.weather_result_label.config(
            text=f"天氣查詢失敗：{self._format_weather_error(exc)}",
            style="DangerHint.TLabel",
        )

    def _on_dept_changed(self):
        if self.on_date_changed:
            self.on_date_changed()

    def _restore_from_settings(self):
        if self.settings.excel_path:
            self.excel_var.set(self.settings.excel_path)
        if self.settings.output_dir:
            self.output_var.set(self.settings.output_dir)
        if self.settings.sheet_name:
            self.sheet_var.set(self.settings.sheet_name)

        try:
            start_date = datetime.strptime(self.settings.start_date, "%Y-%m-%d")
            self.start_date_entry.entry.delete(0, "end")
            self.start_date_entry.entry.insert(0, start_date.strftime("%Y-%m-%d"))
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
            workbook = openpyxl.load_workbook(path, read_only=True)
            names = workbook.sheetnames
            workbook.close()
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
            self.computed_label.config(text=f"D{day_number}")
            self.dday_note_label.config(
                text=f"報告日期 {report.strftime('%Y-%m-%d')}  |  起始日 {start.strftime('%Y-%m-%d')}"
            )
        except (ValueError, TypeError):
            self.computed_label.config(text="日期格式錯誤")
            self.dday_note_label.config(text="請確認起始日與報告日期格式")

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
        if self._search_results:
            idx = self.city_combo.current()
            if 0 <= idx < len(self._search_results):
                return self._search_results[idx].name
        return self.city_search_var.get()

    def get_weather_text(self) -> str:
        if self.weather_enabled_var.get() and self._weather_data:
            return self._weather_data.format_line()
        return ""

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
