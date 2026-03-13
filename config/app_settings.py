from dataclasses import dataclass, field, asdict


@dataclass
class AppSettings:
    excel_path: str = ""
    template_path: str = ""
    output_dir: str = ""
    start_date: str = "2025-04-24"
    sheet_name: str = "總表"
    selected_template: str = ""
    table_layout: str = "span"
    window_geometry: str = "1200x800"
    department: str = "engine"       # "engine" / "deck" / "none"
    filename_prefix: str = "SY-"    # 使用者自訂檔名前綴
    show_weather: bool = True         # 是否顯示天氣
    weather_country: str = "TW"      # 國家 ISO 代碼
    weather_city: str = "Kaohsiung"  # 城市搜尋關鍵字
    weather_town: str = ""           # 鄉鎮搜尋關鍵字（選填）

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "AppSettings":
        valid_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in valid_fields}
        return cls(**filtered)
