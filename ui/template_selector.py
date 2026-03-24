import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog

from ui.constants import BTN_ADD, BTN_BROWSE, BTN_UTILITY


class TemplateSelector(ttk.Frame):
    def __init__(self, parent, template_manager, **kwargs):
        super().__init__(parent, **kwargs)
        self.template_manager = template_manager
        self.selected_path = ttk.StringVar()

        self._build_ui()
        self.refresh_templates()

    def _build_ui(self):
        columns = ("name", "modified")
        self.tree = ttk.Treeview(
            self,
            columns=columns,
            show="headings",
            height=5,
            selectmode="browse",
        )
        self.tree.heading("name", text="模板名稱")
        self.tree.heading("modified", text="修改日期")
        self.tree.column("name", width=420, stretch=True)
        self.tree.column("modified", width=170, stretch=False)
        self.tree.pack(fill=BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=X, pady=(12, 0))
        ttk.Button(
            btn_frame,
            text="新增模板…",
            command=self._add_template,
            bootstyle=BTN_ADD,
            width=12,
        ).pack(side=LEFT, padx=(0, 6))
        ttk.Button(
            btn_frame,
            text="開啟資料夾",
            command=self._open_folder,
            bootstyle=BTN_BROWSE,
            width=12,
        ).pack(side=LEFT, padx=(0, 6))
        ttk.Button(
            btn_frame,
            text="重新整理",
            command=self.refresh_templates,
            bootstyle=BTN_UTILITY,
            width=10,
        ).pack(side=LEFT)

    def refresh_templates(self):
        self.tree.delete(*self.tree.get_children())
        templates = self.template_manager.discover_templates()
        if not templates:
            self.tree.insert("", "end", values=("（尚未有模板，請點「新增模板」）", "—"))
            return
        for template in templates:
            self.tree.insert(
                "",
                "end",
                values=(template.name, template.modified_date),
                tags=(template.path,),
            )
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self._on_select(None)

    def _on_select(self, event):
        selection = self.tree.selection()
        if not selection:
            return
        item = self.tree.item(selection[0])
        tags = item.get("tags", [])
        if tags:
            self.selected_path.set(tags[0])

    def _add_template(self):
        path = filedialog.askopenfilename(
            title="選擇 Word 模板檔案",
            filetypes=[("Word files", "*.docx")],
        )
        if path:
            self.template_manager.add_template(path)
            self.refresh_templates()

    def _open_folder(self):
        os.startfile(self.template_manager.get_templates_dir())

    def get_selected_template_path(self) -> str:
        return self.selected_path.get()
