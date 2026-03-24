import os
from datetime import datetime

import tksheet
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from ui.constants import BTN_SAVE, BTN_UTILITY


class ExcelEditorPanel(ttk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.is_modified = False
        self._excel_path = None
        self._sheet_name = None
        self._original_columns = []
        self._processor = None

        self._build_ui()

    def _build_ui(self):
        page = ttk.Frame(self, style="Page.TFrame", padding=(22, 22, 22, 16))
        page.pack(fill=BOTH, expand=True)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(page, style="Surface.TFrame", padding=(18, 14))
        toolbar.grid(row=0, column=0, sticky=EW, pady=(0, 14))
        toolbar.columnconfigure(2, weight=1)

        action_row = ttk.Frame(toolbar, style="Surface.TFrame")
        action_row.grid(row=0, column=0, sticky=W)
        ttk.Button(
            action_row,
            text="儲存變更",
            command=self._save_changes,
            bootstyle=BTN_SAVE,
            width=10,
        ).pack(side=LEFT, padx=(0, 8))
        ttk.Button(
            action_row,
            text="重新載入",
            command=self._reload,
            bootstyle=BTN_UTILITY,
            width=10,
        ).pack(side=LEFT)

        ttk.Separator(toolbar, orient="vertical").grid(row=0, column=1, sticky=NS, padx=16)

        meta = ttk.Frame(toolbar, style="Surface.TFrame")
        meta.grid(row=0, column=2, sticky=E)
        ttk.Label(meta, text="目前檔案", style="ToolbarCaption.TLabel").grid(row=0, column=0, sticky=E)
        self.file_label = ttk.Label(meta, text="尚未載入 Excel", style="ToolbarMeta.TLabel")
        self.file_label.grid(row=0, column=1, sticky=W, padx=(8, 0))
        ttk.Label(meta, text="資料規模", style="ToolbarCaption.TLabel").grid(
            row=0, column=2, sticky=E, padx=(18, 0)
        )
        self.dimension_label = ttk.Label(meta, text="0 列 × 0 欄", style="ToolbarMeta.TLabel")
        self.dimension_label.grid(row=0, column=3, sticky=W, padx=(8, 0))

        self.modified_label = ttk.Label(toolbar, text="", style="WarningHint.TLabel")
        self.modified_label.grid(row=0, column=3, sticky=E, padx=(18, 0))

        sheet_frame = ttk.Frame(page, style="Surface.TFrame")
        sheet_frame.grid(row=1, column=0, sticky=NSEW)
        sheet_frame.columnconfigure(0, weight=1)
        sheet_frame.rowconfigure(0, weight=1)

        self.sheet = tksheet.Sheet(
            sheet_frame,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
        )
        self.sheet.grid(row=0, column=0, sticky=NSEW)
        self.sheet.enable_bindings()
        self.sheet.extra_bindings([
            ("end_edit_cell", self._on_cell_edit),
            ("end_paste", self._on_cell_edit),
            ("end_delete", self._on_cell_edit),
            ("end_insert_rows", self._on_cell_edit),
            ("end_delete_rows", self._on_cell_edit),
        ])

        status = ttk.Frame(page, style="StatusBar.TFrame", padding=(14, 8))
        status.grid(row=2, column=0, sticky=EW, pady=(12, 0))
        status.columnconfigure(1, weight=1)
        ttk.Label(status, text="狀態", style="StatusCaption.TLabel").grid(row=0, column=0, sticky=W)
        self.status_label = ttk.Label(status, text="尚未載入檔案", style="StatusValue.TLabel")
        self.status_label.grid(row=0, column=1, sticky=W, padx=(8, 0))

    def load_excel(self, path: str, sheet_name: str):
        from core.excel_processor import ExcelProcessor

        self.file_label.config(text="載入中…")
        self.update_idletasks()
        try:
            self._processor = ExcelProcessor(path, sheet_name)
            self._processor.load()
            self._excel_path = path
            self._sheet_name = sheet_name
            self._original_columns = list(self._processor._original_columns)

            headers = self._processor.get_display_headers()
            data = self._processor.get_display_data()

            self.sheet.headers(headers)
            self.sheet.set_sheet_data(data)
            self.sheet.set_all_column_widths(120)

            row_count = len(data)
            col_count = len(headers)
            filename = os.path.basename(path)
            self.file_label.config(text=f"{filename} / {sheet_name}")
            self.dimension_label.config(text=f"{row_count} 列 × {col_count} 欄")
            self.status_label.config(text=f"已載入: {path}")
            self.is_modified = False
            self.modified_label.config(text="")
        except Exception as exc:
            from ttkbootstrap.dialogs import Messagebox

            Messagebox.show_error(f"載入 Excel 失敗:\n{exc}", title="錯誤")

    def _on_cell_edit(self, event):
        self.is_modified = True
        self.modified_label.config(text="未儲存變更")

    def _save_changes(self):
        if not self._excel_path or not self._processor:
            from ttkbootstrap.dialogs import Messagebox

            Messagebox.show_error("尚未載入 Excel 檔案", title="錯誤")
            return
        try:
            data = self.sheet.get_sheet_data()
            headers = self.sheet.headers()
            self._processor.save_from_table_data(data, headers)
            self.is_modified = False
            self.modified_label.config(text="")
            now = datetime.now().strftime("%H:%M:%S")
            self.status_label.config(text=f"已儲存 ({now})")
        except Exception as exc:
            from ttkbootstrap.dialogs import Messagebox

            Messagebox.show_error(f"儲存失敗:\n{exc}", title="錯誤")

    def _reload(self):
        if self._excel_path and self._sheet_name:
            if self.is_modified:
                from ttkbootstrap.dialogs import Messagebox

                result = Messagebox.yesno("有未儲存的變更，確定要重新載入嗎？", title="確認")
                if result != "Yes":
                    return
            self.load_excel(self._excel_path, self._sheet_name)

    def check_unsaved(self) -> bool:
        if not self.is_modified:
            return True
        from ttkbootstrap.dialogs import Messagebox

        result = Messagebox.yesnocancel("有未儲存的 Excel 變更，要先儲存嗎？", title="未儲存的變更")
        if result == "Yes":
            self._save_changes()
            return True
        if result == "No":
            return True
        return False
