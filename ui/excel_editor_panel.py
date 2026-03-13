import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
import tksheet

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
        # ── Toolbar ──────────────────────────────────────────
        toolbar = ttk.Frame(self, padding=(10, 6))
        toolbar.pack(fill=X)

        # Action group
        ttk.Button(toolbar, text="儲存變更", command=self._save_changes,
                   bootstyle=BTN_SAVE, width=10).pack(side=LEFT, padx=(0, 4))
        ttk.Button(toolbar, text="重新載入", command=self._reload,
                   bootstyle=BTN_UTILITY, width=9).pack(side=LEFT)

        # Vertical separator between actions and file info
        ttk.Separator(toolbar, orient="vertical").pack(
            side=LEFT, fill=Y, padx=(12, 12))

        # File info: filename + sheet + row/col counts
        self.info_label = ttk.Label(toolbar, text="尚未載入 Excel",
                                    bootstyle="secondary", font=("", 9))
        self.info_label.pack(side=LEFT)

        # Unsaved indicator — right side
        self.modified_label = ttk.Label(toolbar, text="",
                                        bootstyle="warning",
                                        font=("", 9, "bold"))
        self.modified_label.pack(side=RIGHT, padx=(8, 2))

        # ── Separator ────────────────────────────────────────
        ttk.Separator(self, orient="horizontal").pack(fill=X)

        # ── Spreadsheet (fill all remaining space) ───────────
        self.sheet = tksheet.Sheet(
            self,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
        )
        self.sheet.pack(fill=BOTH, expand=True)

        # Enable all interactive features
        self.sheet.enable_bindings()

        # Bind edit events to track modifications
        self.sheet.extra_bindings([
            ("end_edit_cell",   self._on_cell_edit),
            ("end_paste",       self._on_cell_edit),
            ("end_delete",      self._on_cell_edit),
            ("end_insert_rows", self._on_cell_edit),
            ("end_delete_rows", self._on_cell_edit),
        ])

        # ── Status bar ────────────────────────────────────────
        ttk.Separator(self, orient="horizontal").pack(fill=X)
        status = ttk.Frame(self, padding=(10, 3))
        status.pack(fill=X)
        self.status_label = ttk.Label(status, text="",
                                      bootstyle="secondary", font=("", 8))
        self.status_label.pack(side=LEFT)

    # ── Data loading ─────────────────────────────────────────

    def load_excel(self, path: str, sheet_name: str):
        from core.excel_processor import ExcelProcessor
        self.info_label.config(text="載入中…")
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

            # Auto-resize columns
            self.sheet.set_all_column_widths(120)

            row_count = len(data)
            col_count = len(headers)
            fname = os.path.basename(path)
            self.info_label.config(
                text=f"{fname}  [{sheet_name}]   {row_count} 列 × {col_count} 欄")
            self.status_label.config(text=f"已載入: {path}")
            self.is_modified = False
            self.modified_label.config(text="")

        except Exception as e:
            from ttkbootstrap.dialogs import Messagebox
            Messagebox.show_error(f"載入 Excel 失敗:\n{e}", title="錯誤")

    # ── Internal callbacks ────────────────────────────────────

    def _on_cell_edit(self, event):
        self.is_modified = True
        self.modified_label.config(text="● 未儲存")

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
        except Exception as e:
            from ttkbootstrap.dialogs import Messagebox
            Messagebox.show_error(f"儲存失敗:\n{e}", title="錯誤")

    def _reload(self):
        if self._excel_path and self._sheet_name:
            if self.is_modified:
                from ttkbootstrap.dialogs import Messagebox
                result = Messagebox.yesno(
                    "有未儲存的變更，確定要重新載入嗎？", title="確認")
                if result != "Yes":
                    return
            self.load_excel(self._excel_path, self._sheet_name)

    def check_unsaved(self) -> bool:
        """Returns True if it's OK to proceed, False if user cancelled."""
        if not self.is_modified:
            return True
        from ttkbootstrap.dialogs import Messagebox
        result = Messagebox.yesnocancel(
            "有未儲存的 Excel 變更，要先儲存嗎？", title="未儲存的變更")
        if result == "Yes":
            self._save_changes()
            return True
        elif result == "No":
            return True
        else:
            return False
