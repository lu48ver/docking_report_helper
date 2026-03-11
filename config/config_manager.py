import json
import os
from config.app_settings import AppSettings


class ConfigManager:
    def __init__(self, config_path: str = "last_paths.json"):
        self.config_path = config_path

    def load(self) -> AppSettings:
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return AppSettings.from_dict(data)
            except (json.JSONDecodeError, TypeError):
                return AppSettings()
        return AppSettings()

    def save(self, settings: AppSettings) -> None:
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(settings.to_dict(), f, ensure_ascii=False, indent=2)
