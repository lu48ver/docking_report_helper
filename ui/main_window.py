import os
import threading

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox

from app_meta import APP_NAME, APP_VERSION
from config.config_manager import ConfigManager
from core.template_manager import TemplateManager
from core.excel_processor import ExcelProcessor
from core.report_generator import ReportGenerator
from core.folder_creator import FolderCreator
from core.weather_service import WeatherService
from core.path_utils import get_app_path
from ui.config_panel import ConfigPanel
from ui.excel_editor_panel import ExcelEditorPanel
from ui.report_output_panel import ReportOutputPanel
from ui.constants import (
    APP_BG,
    SURFACE_BG,
    SURFACE_MUTED_BG,
    SURFACE_SUBTLE_BG,
    TEXT_COLOR,
    TEXT_MUTED,
    ACCENT_BG,
    ACCENT_FG,
    SUCCESS_BG,
    SUCCESS_FG,
    UI_FONT,
)


class MainWindow(ttk.Window):
    def __init__(self):
        self.config_manager = ConfigManager()
        self.settings = self.config_manager.load()
        self.template_manager = TemplateManager(get_app_path("templates"))

        super().__init__(title=APP_NAME, themename="litera")
        self.geometry(self.settings.window_geometry)
        self.minsize(980, 700)

        self._configure_styles()
        self._build_ui()
        self._wire_events()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_styles(self):
        style = ttk.Style()
        self.configure(bg=APP_BG)

        style.configure(".", font=(UI_FONT, 10))
        style.configure("TFrame", background=APP_BG)
        style.configure("Page.TFrame", background=APP_BG)
        style.configure("Surface.TFrame", background=SURFACE_BG)
        style.configure("MutedSurface.TFrame", background=SURFACE_MUTED_BG)
        style.configure("StatusBar.TFrame", background=SURFACE_SUBTLE_BG)
        style.configure("AccentCard.TFrame", background=ACCENT_BG)
        style.configure("SummaryCard.TFrame", background=SUCCESS_BG)
        style.configure("OptionCard.TFrame", background=SURFACE_MUTED_BG)
        style.configure("SelectedOptionCard.TFrame", background=ACCENT_BG)

        style.configure("TLabel", background=APP_BG, foreground=TEXT_COLOR)
        style.configure("SectionTitle.TLabel", background=SURFACE_BG, foreground=TEXT_COLOR, font=(UI_FONT, 11, "bold"))
        style.configure("SectionDesc.TLabel", background=SURFACE_BG, foreground=TEXT_MUTED, font=(UI_FONT, 9))
        style.configure("FieldLabel.TLabel", background=SURFACE_BG, foreground=TEXT_MUTED, font=(UI_FONT, 9, "bold"))
        style.configure("Hint.TLabel", background=SURFACE_BG, foreground=TEXT_MUTED, font=(UI_FONT, 9))
        style.configure("SuccessHint.TLabel", background=SURFACE_BG, foreground=SUCCESS_FG, font=(UI_FONT, 9))
        style.configure("WarningHint.TLabel", background=SURFACE_BG, foreground="#9a6c11", font=(UI_FONT, 9))
        style.configure("DangerHint.TLabel", background=SURFACE_BG, foreground="#a13a3a", font=(UI_FONT, 9))
        style.configure("WeatherResult.TLabel", background=SURFACE_MUTED_BG, foreground=SUCCESS_FG, font=(UI_FONT, 10, "bold"))
        style.configure("CardEyebrow.TLabel", background=ACCENT_BG, foreground=ACCENT_FG, font=(UI_FONT, 9, "bold"))
        style.configure("CardValue.TLabel", background=ACCENT_BG, foreground=TEXT_COLOR, font=(UI_FONT, 18, "bold"))
        style.configure("CardNote.TLabel", background=ACCENT_BG, foreground=TEXT_MUTED, font=(UI_FONT, 9))
        style.configure("SummaryEyebrow.TLabel", background=SUCCESS_BG, foreground=SUCCESS_FG, font=(UI_FONT, 9, "bold"))
        style.configure("SummaryValue.TLabel", background=SUCCESS_BG, foreground=TEXT_COLOR, font=(UI_FONT, 18, "bold"))
        style.configure("SummaryNote.TLabel", background=SUCCESS_BG, foreground=TEXT_MUTED, font=(UI_FONT, 9))
        style.configure("ToolbarCaption.TLabel", background=SURFACE_BG, foreground=TEXT_MUTED, font=(UI_FONT, 9, "bold"))
        style.configure("ToolbarMeta.TLabel", background=SURFACE_BG, foreground=TEXT_COLOR, font=(UI_FONT, 10))
        style.configure("StatusCaption.TLabel", background=SURFACE_SUBTLE_BG, foreground=TEXT_MUTED, font=(UI_FONT, 9, "bold"))
        style.configure("StatusValue.TLabel", background=SURFACE_SUBTLE_BG, foreground=TEXT_COLOR, font=(UI_FONT, 9))

        style.configure("TNotebook", background=APP_BG, borderwidth=0, tabmargins=(12, 12, 12, 0))
        style.configure("TNotebook.Tab", padding=(18, 10), font=(UI_FONT, 10, "bold"))
        style.map(
            "TNotebook.Tab",
            background=[("selected", SURFACE_BG), ("!selected", "#dde4ec")],
            foreground=[("selected", TEXT_COLOR), ("!selected", TEXT_MUTED)],
        )
        style.configure("Treeview.Heading", font=(UI_FONT, 9, "bold"))

    def _build_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=(10, 0))

        self.config_panel = ConfigPanel(self.notebook, self.settings)
        self.notebook.add(self.config_panel, text="設定")

        self.editor_panel = ExcelEditorPanel(self.notebook)
        self.notebook.add(self.editor_panel, text="Excel 編輯器")

        self.output_panel = ReportOutputPanel(self.notebook, self.template_manager, settings=self.settings)
        self.notebook.add(self.output_panel, text="產生報告")

        ttk.Separator(self, orient="horizontal").pack(fill=X, side=BOTTOM)
        status_frame = ttk.Frame(self, padding=(16, 5))
        status_frame.pack(fill=X, side=BOTTOM)
        self.status_label = ttk.Label(status_frame, text="就緒", bootstyle="secondary")
        self.status_label.pack(side=LEFT)
        ttk.Label(status_frame, text=f"{APP_NAME} v{APP_VERSION}", bootstyle="secondary").pack(side=RIGHT)

    def _wire_events(self):
        self.config_panel.on_excel_loaded = self._handle_excel_loaded
        self.config_panel.on_date_changed = self._update_preview
        self.output_panel.on_generate = self._handle_generate
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        if self.settings.excel_path and os.path.exists(self.settings.excel_path):
            self.after(500, lambda: self._handle_excel_loaded(self.settings.excel_path, self.settings.sheet_name))

    def _on_tab_changed(self, event):
        try:
            if self.notebook.index(self.notebook.select()) == 2:
                self._update_preview()
        except Exception:
            pass

    def _handle_excel_loaded(self, path: str, sheet_name: str):
        self.status_label.config(text=f"載入中: {path}")
        self.editor_panel.load_excel(path, sheet_name)
        self._update_preview()
        self.status_label.config(text="就緒")

    def _get_dept_info(self):
        dept = self.config_panel.get_department()
        return ReportGenerator.DEPT_MAP.get(dept, ReportGenerator.DEPT_MAP["engine"])

    def _update_preview(self):
        try:
            report_date = self.config_panel.get_report_date()
            start_date = self.config_panel.get_start_date()
            day_number = (report_date - start_date).days
            day_str = f"D{day_number}"
            day_date_str = report_date.strftime("%Y%m%d")
            dept_title, _ = self._get_dept_info()
            prefix = self.config_panel.get_filename_prefix()
            self.output_panel.update_filename(day_date_str, day_str, dept_title, prefix)

            excel_path = self.config_panel.get_excel_path()
            sheet_name = self.config_panel.get_sheet_name()
            if excel_path and os.path.exists(excel_path):
                processor = ExcelProcessor(excel_path, sheet_name)
                processor.load()
                try:
                    shipyard, engine = processor.extract_tasks(report_date)
                    self.output_panel.update_preview(len(shipyard), len(engine))
                    all_tasks = shipyard + engine
                    if all_tasks:
                        self.output_panel.set_sample_tasks(all_tasks)
                except ValueError:
                    self.output_panel.update_preview(0, 0)
        except (ValueError, TypeError):
            pass

    def _handle_generate(self):
        excel_path = self.config_panel.get_excel_path()
        output_dir = self.config_panel.get_output_dir()
        template_path = self.output_panel.get_selected_template_path()

        if not excel_path or not os.path.exists(excel_path):
            Messagebox.show_error("請先選擇有效的 Excel 檔案", title="錯誤")
            return
        if not output_dir:
            Messagebox.show_error("請先選擇輸出資料夾", title="錯誤")
            return
        if not template_path or not os.path.exists(template_path):
            Messagebox.show_error("請先選擇模板", title="錯誤")
            return

        if not self.editor_panel.check_unsaved():
            return

        self.output_panel.set_generating(True)
        self.output_panel.clear_log()
        threading.Thread(target=self._generate_worker, daemon=True).start()

    def _generate_worker(self):
        try:
            self._log("開始產生報告...")

            report_date = self.config_panel.get_report_date()
            start_date = self.config_panel.get_start_date()
            excel_path = self.config_panel.get_excel_path()
            sheet_name = self.config_panel.get_sheet_name()
            output_dir = self.config_panel.get_output_dir()
            template_path = self.output_panel.get_selected_template_path()
            department = self.config_panel.get_department()
            prefix = self.config_panel.get_filename_prefix()
            dept_title, dept_eng = self._get_dept_info()

            day_number = (report_date - start_date).days
            day_str = f"D{day_number}"
            day_date_str = report_date.strftime("%Y%m%d")
            output_path = os.path.join(output_dir, f"{prefix}{dept_title}-{day_date_str}-{day_str}.docx")

            weather_text = ""
            if self.config_panel.weather_enabled_var.get():
                weather_text = self.config_panel.get_weather_text()
                if not weather_text:
                    self._log("正在抓取天氣資料...")
                    city = self.config_panel.get_weather_city()
                    if city:
                        try:
                            cities = WeatherService.search_city(city)
                            if cities:
                                weather = WeatherService.fetch_weather(
                                    cities[0].latitude,
                                    cities[0].longitude,
                                    report_date,
                                )
                                if weather:
                                    weather_text = weather.format_line()
                        except Exception as exc:
                            self._log(f"天氣抓取失敗: {exc}")

            self._log("載入 Excel 資料...")
            processor = ExcelProcessor(excel_path, sheet_name)
            processor.load()
            shipyard_tasks, engine_room_tasks = processor.extract_tasks(report_date)
            self._log(f"找到船廠工程 {len(shipyard_tasks)} 筆，自修 {len(engine_room_tasks)} 筆")

            self._log(f"使用模板: {os.path.basename(template_path)}")
            generator = ReportGenerator(template_path)
            table_layout = self.output_panel.get_table_layout()
            self._log(f"部門: {dept_title} / 表格版面: {table_layout}")
            if weather_text:
                self._log(f"天氣資訊: {weather_text}")

            generator.generate(
                target_date=report_date,
                start_date=start_date,
                shipyard_tasks=shipyard_tasks,
                engine_room_tasks=engine_room_tasks,
                output_path=output_path,
                table_layout=table_layout,
                department=department,
                weather_text=weather_text,
            )
            self._log(f"報告已儲存: {output_path}")

            self._log("建立照片資料夾...")
            new_shipyard = generator._filter_new_tasks(shipyard_tasks) if generator.doc else shipyard_tasks
            new_engine = generator._filter_new_tasks(engine_room_tasks) if generator.doc else engine_room_tasks
            created = FolderCreator.create_photo_folders(
                photo_root=output_dir,
                day_date_str=day_date_str,
                new_shipyard_tasks=new_shipyard,
                new_engine_room_tasks=new_engine,
            )
            self._log(f"已建立 {len(created)} 個資料夾")

            self.config_panel.save_to_settings()
            self.settings.template_path = template_path
            self.settings.table_layout = table_layout
            self.config_manager.save(self.settings)

            self._log("完成！")
            self.after(0, lambda: self.output_panel.update_preview(len(shipyard_tasks), len(engine_room_tasks)))
            self.after(0, lambda: self.output_panel.set_generating(False))
            self.after(
                0,
                lambda: Messagebox.show_info(
                    f"報告已儲存至:\n{output_path}\n\n照片資料夾也已建立。",
                    title="完成",
                ),
            )
        except Exception as exc:
            self._log(f"錯誤: {exc}")
            self.after(0, lambda: self.output_panel.set_generating(False))
            self.after(0, lambda: Messagebox.show_error(f"產生報告失敗:\n{exc}", title="錯誤"))

    def _log(self, message: str):
        self.after(0, lambda: self.output_panel.log(message))

    def _on_close(self):
        if not self.editor_panel.check_unsaved():
            return
        self.config_panel.save_to_settings()
        self.settings.window_geometry = self.geometry()
        self.settings.table_layout = self.output_panel.get_table_layout()
        self.config_manager.save(self.settings)
        self.destroy()
