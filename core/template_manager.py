import os
from dataclasses import dataclass
from datetime import datetime


@dataclass
class TemplateInfo:
    name: str
    path: str
    modified_date: str


class TemplateManager:
    def __init__(self, templates_dir: str = "templates"):
        self.templates_dir = templates_dir

    def discover_templates(self) -> list:
        templates = []
        if not os.path.exists(self.templates_dir):
            os.makedirs(self.templates_dir, exist_ok=True)
            return templates

        # Check subfolders first
        for entry in os.listdir(self.templates_dir):
            entry_path = os.path.join(self.templates_dir, entry)

            if os.path.isdir(entry_path):
                # Look for .docx inside subfolder
                for f in os.listdir(entry_path):
                    if f.endswith(".docx") and not f.startswith("~$"):
                        full_path = os.path.join(entry_path, f)
                        mod_time = os.path.getmtime(full_path)
                        mod_str = datetime.fromtimestamp(mod_time).strftime("%Y-%m-%d %H:%M")
                        templates.append(TemplateInfo(
                            name=f"{entry}/{f}",
                            path=full_path,
                            modified_date=mod_str,
                        ))

            elif entry.endswith(".docx") and not entry.startswith("~$"):
                # Flat .docx file directly in templates/
                mod_time = os.path.getmtime(entry_path)
                mod_str = datetime.fromtimestamp(mod_time).strftime("%Y-%m-%d %H:%M")
                templates.append(TemplateInfo(
                    name=entry,
                    path=entry_path,
                    modified_date=mod_str,
                ))

        return templates

    def add_template(self, source_path: str, name: str = None) -> TemplateInfo:
        if name is None:
            name = os.path.splitext(os.path.basename(source_path))[0]

        dest_dir = os.path.join(self.templates_dir, name)
        os.makedirs(dest_dir, exist_ok=True)

        dest_file = os.path.join(dest_dir, os.path.basename(source_path))
        import shutil
        shutil.copy2(source_path, dest_file)

        mod_time = os.path.getmtime(dest_file)
        mod_str = datetime.fromtimestamp(mod_time).strftime("%Y-%m-%d %H:%M")
        return TemplateInfo(
            name=f"{name}/{os.path.basename(source_path)}",
            path=dest_file,
            modified_date=mod_str,
        )

    def get_templates_dir(self) -> str:
        return os.path.abspath(self.templates_dir)
