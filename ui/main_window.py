import os
import threading
from datetime import datetime
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox

from config.config_manager import ConfigManager
from config.app_settings import AppSettings
from core.template_manager import TemplateManager
from core.excel_processor import ExcelProcessor
from core.report_generator import ReportGenerator
from core.folder_creator import FolderCreator
from core.weather_service import WeatherService
from ui.config_panel import ConfigPanel
from ui.excel_editor_panel import ExcelEditorPanel
from ui.report_output_panel import ReportOutputPanel


class MainWindow(ttk.Window):
    def __init__(self):
        self.config_manager = ConfigManager()
        self.settings = self.config_manager.load()
        self.template_manager = TemplateManager("templates")

        super().__init__(title="塢修報告產生器", themename="litera")
        self.geometry(self.settings.window_geometry)
        self.minsize(980, 700)

        # Register custom ttk styles BEFORE building any widgets
        self._configure_styles()
        self._build_ui()
        self._wire_events()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Custom style registration ─────────────────────────────────────────

    def _configure_styles(self):
        """Register custom ttk styles used by v3 UI components."""
        style = ttk.Style()

        # D-day highlight card + output summary card
        style.configure("DayCard.TFrame",
                        background="#eaf0fb")
        style.configure("DayCard.TLabel",
                        background="#eaf0fb",
                        foreground="#2c5fad")

        # Excel editor toolbar — extremely light grey to distinguish from sheet
        style.configure("Toolbar.TFrame",
                        background="#f4f6f9")
        style.configure("Toolbar.TLabel",
                        background="#f4f6f9")
        style.configure("Toolbar.TButton",
                        background="#f4f6f9")

    # ── UI construction ───────────────────────────────────────────────────

    def _build_ui(self):
        # Notebook with three tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=BOTH, expand=True, padx=8, pady=(8, 0))

        # Tab 1: Settings
        self.config_panel = ConfigPanel(self.notebook, self.settings)
        self.notebook.add(self.config_panel, text=" ⚙  設定 ")

        # Tab 2: Excel Editor
        self.editor_panel = ExcelEditorPanel(self.notebook)
        self.notebook.add(self.editor_panel, text=" 📊  Excel 編輯器 ")

        # Tab 3: Generate Report
        self.output_panel = ReportOutputPanel(
            self.notebook, self.template_manager, settings=self.settings)
        self.notebook.add(self.output_panel, text=" 📄  產生報告 ")

        # Status bar
        ttk.Separator(self, orient="horizontal").pack(fill=X, side=BOTTOM)
        status_frame = ttk.Frame(self, padding=(16, 5))
        status_frame.pack(fill=X, side=BOTTOM)
        self.status_label = ttk.Label(
            status_frame, text="就緒", bootstyle="secondary")
        self.status_label.pack(side=LEFT)
        ttk.Label(status_frame, text="塢修報告產生器  v3.0",
                  bootstyle="secondary").pack(side=RIGHT)

    def _wire_events(self):
        # Config panel -> load excel into editor
        self.config_panel.on_excel_loaded = self._handle_excel_loaded

        # Config panel -> date changed -> refresh preview
        self.config_panel.on_date_changed = self._update_preview

        # Generate button
        self.output_panel.on_generate = self._handle_generate

        # Refresh preview when switching to "產生報告" tab
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # Auto-load if settings have a valid excel path
        if self.settings.excel_path and os.path.exists(self.settings.excel_path):
            self.after(500, lambda: self._handle_excel_loaded(
                self.settings.excel_path, self.settings.sheet_name
            ))

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

        thread = threading.Thread(target=self._generate_worker, daemon=True)
        thread.start()

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
            output_path = os.path.join(
                output_dir,
                f"{prefix}{dept_title}-{day_date_str}-{day_str}.docx")

            # Fetch weather if enabled
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
                                w = WeatherService.fetch_weather(
                                    cities[0].latitude, cities[0].longitude,
                                    report_date)
                                if w:
                                    weather_text = w.format_line()
                        except Exception as we:
                            self._log(f"天氣抓取失敗: {we}")

            # Extract tasks
            self._log("載入 Excel 資料...")
            processor = ExcelProcessor(excel_path, sheet_name)
            processor.load()
            shipyard_tasks, engine_room_tasks = processor.extract_tasks(report_date)
            self._log(
                f"找到船廠工程 {len(shipyard_tasks)} 筆，自修 {len(engine_room_tasks)} 筆")

            # Generate report
            self._log(f"使用模板: {os.path.basename(template_path)}")
            generator = ReportGenerator(template_path)
            table_layout = self.output_panel.get_table_layout()
            self._log(f"部門: {dept_title}  /  表格版面: {table_layout}")
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

            # Create photo folders
            self._log("建立照片資料夾...")
            new_shipyard = (generator._filter_new_tasks(shipyard_tasks)
                            if generator.doc else shipyard_tasks)
            new_engine = (generator._filter_new_tasks(engine_room_tasks)
                          if generator.doc else engine_room_tasks)

            created = FolderCreator.create_photo_folders(
                photo_root=output_dir,
                day_date_str=day_date_str,
                new_shipyard_tasks=new_shipyard,
                new_engine_room_tasks=new_engine,
            )
            self._log(f"已建立 {len(created)} 個資料夾")

            # Save config
            self.config_panel.save_to_settings()
            self.settings.template_path = template_path
            self.settings.table_layout = table_layout
            self.config_manager.save(self.settings)

            self._log("完成！")
            self.after(0, lambda: self.output_panel.update_preview(
                len(shipyard_tasks), len(engine_room_tasks)))
            self.after(0, lambda: self.output_panel.set_generating(False))
            self.after(0, lambda: Messagebox.show_info(
                f"報告已儲存至:\n{output_path}\n\n照片資料夾也已建立。",
                title="完成",
            ))

        except Exception as e:
            self._log(f"錯誤: {e}")
            self.after(0, lambda: self.output_panel.set_generating(False))
            self.after(0, lambda: Messagebox.show_error(
                f"產生報告失敗:\n{e}", title="錯誤"))

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
