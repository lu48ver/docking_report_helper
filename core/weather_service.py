import json
import os
import re
import ssl
import urllib.request
import urllib.parse
from urllib.error import HTTPError, URLError
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Optional

try:
    import certifi
except ImportError:  # pragma: no cover - optional at runtime
    certifi = None


# ── CJK city name → English fallback map ──────────────────────────────────
# GeoNames (used by Open-Meteo) has inconsistent CJK coverage.
# When a Chinese/Japanese/Korean query returns no results, the service
# tries the mapped English name automatically.
_CJK_ALIAS: dict[str, str] = {
    # Taiwan
    "台北": "Taipei",     "臺北": "Taipei",
    "台中": "Taichung",   "臺中": "Taichung",
    "台南": "Tainan",     "臺南": "Tainan",
    "高雄": "Kaohsiung",
    "基隆": "Keelung",
    "桃園": "Taoyuan",
    "新竹": "Hsinchu",
    "苗栗": "Miaoli",
    "彰化": "Changhua",
    "南投": "Nantou",
    "雲林": "Yunlin",
    "嘉義": "Chiayi",
    "屏東": "Pingtung",
    "宜蘭": "Yilan",
    "花蓮": "Hualien",
    "台東": "Taitung",    "臺東": "Taitung",
    "澎湖": "Penghu",
    "金門": "Kinmen",
    "馬祖": "Matsu",
    "汐止": "Xizhi",
    "板橋": "Banqiao",
    "三重": "Sanchong",
    "中和": "Zhonghe",
    "新莊": "Xinzhuang",
    "左營": "Zuoying",
    "鳳山": "Fengshan",
    "安平": "Anping",
    # China (common ports / cities)
    "上海": "Shanghai",
    "北京": "Beijing",
    "廣州": "Guangzhou",
    "深圳": "Shenzhen",
    "天津": "Tianjin",
    "青島": "Qingdao",
    "廈門": "Xiamen",
    "寧波": "Ningbo",
    "大連": "Dalian",
    "武漢": "Wuhan",
    "成都": "Chengdu",
    "重慶": "Chongqing",
    "南京": "Nanjing",
    "杭州": "Hangzhou",
    "福州": "Fuzhou",
    "海口": "Haikou",
    # Japan (common)
    "東京": "Tokyo",
    "大阪": "Osaka",
    "横浜": "Yokohama",
    "神戸": "Kobe",
    "名古屋": "Nagoya",
    "福岡": "Fukuoka",
    "札幌": "Sapporo",
    "那覇": "Naha",
    # Korea
    "서울": "Seoul",
    "부산": "Busan",
    "인천": "Incheon",
}


# ISO-3166-1 alpha-2 country code  →  English name
# Sorted A→Z by name for display in dropdowns
COMMON_COUNTRIES: list[tuple[str, str]] = sorted([
    ("AF", "Afghanistan"),    ("AL", "Albania"),      ("DZ", "Algeria"),
    ("AR", "Argentina"),      ("AU", "Australia"),    ("AT", "Austria"),
    ("BH", "Bahrain"),        ("BD", "Bangladesh"),   ("BE", "Belgium"),
    ("BZ", "Belize"),         ("BO", "Bolivia"),      ("BA", "Bosnia and Herzegovina"),
    ("BR", "Brazil"),         ("BG", "Bulgaria"),     ("KH", "Cambodia"),
    ("CA", "Canada"),         ("CL", "Chile"),        ("CN", "China"),
    ("CO", "Colombia"),       ("CR", "Costa Rica"),   ("HR", "Croatia"),
    ("CU", "Cuba"),           ("CY", "Cyprus"),       ("CZ", "Czech Republic"),
    ("DK", "Denmark"),        ("DJ", "Djibouti"),     ("DO", "Dominican Republic"),
    ("EC", "Ecuador"),        ("EG", "Egypt"),        ("SV", "El Salvador"),
    ("ET", "Ethiopia"),       ("FI", "Finland"),      ("FR", "France"),
    ("GH", "Ghana"),          ("GR", "Greece"),       ("GT", "Guatemala"),
    ("GY", "Guyana"),         ("HT", "Haiti"),        ("HN", "Honduras"),
    ("HK", "Hong Kong"),      ("HU", "Hungary"),      ("IS", "Iceland"),
    ("IN", "India"),          ("ID", "Indonesia"),    ("IR", "Iran"),
    ("IQ", "Iraq"),           ("IE", "Ireland"),      ("IL", "Israel"),
    ("IT", "Italy"),          ("JM", "Jamaica"),      ("JP", "Japan"),
    ("JO", "Jordan"),         ("KE", "Kenya"),        ("KW", "Kuwait"),
    ("LB", "Lebanon"),        ("LY", "Libya"),        ("LT", "Lithuania"),
    ("MG", "Madagascar"),     ("MY", "Malaysia"),     ("MV", "Maldives"),
    ("MT", "Malta"),          ("MX", "Mexico"),       ("MA", "Morocco"),
    ("MZ", "Mozambique"),     ("MM", "Myanmar"),      ("NL", "Netherlands"),
    ("NZ", "New Zealand"),    ("NI", "Nicaragua"),    ("NG", "Nigeria"),
    ("NO", "Norway"),         ("OM", "Oman"),         ("PK", "Pakistan"),
    ("PA", "Panama"),         ("PY", "Paraguay"),     ("PE", "Peru"),
    ("PH", "Philippines"),    ("PL", "Poland"),       ("PT", "Portugal"),
    ("QA", "Qatar"),          ("RO", "Romania"),      ("RU", "Russia"),
    ("SA", "Saudi Arabia"),   ("RS", "Serbia"),       ("SG", "Singapore"),
    ("SO", "Somalia"),        ("ZA", "South Africa"), ("KR", "South Korea"),
    ("ES", "Spain"),          ("LK", "Sri Lanka"),    ("SD", "Sudan"),
    ("SR", "Suriname"),       ("SE", "Sweden"),       ("CH", "Switzerland"),
    ("SY", "Syria"),          ("TW", "Taiwan"),       ("TZ", "Tanzania"),
    ("TH", "Thailand"),       ("TT", "Trinidad and Tobago"),
    ("TN", "Tunisia"),        ("TR", "Turkey"),       ("UA", "Ukraine"),
    ("AE", "United Arab Emirates"),
    ("GB", "United Kingdom"), ("US", "United States"),
    ("UY", "Uruguay"),        ("VE", "Venezuela"),    ("VN", "Vietnam"),
    ("YE", "Yemen"),
], key=lambda x: x[1])


WMO_WEATHER_CODES = {
    0: ("Clear", "晴"),
    1: ("Mainly Clear", "大致晴朗"),
    2: ("Partly Cloudy", "多雲"),
    3: ("Overcast", "陰天"),
    45: ("Fog", "霧"),
    48: ("Depositing Rime Fog", "霧淞"),
    51: ("Light Drizzle", "小毛毛雨"),
    53: ("Moderate Drizzle", "毛毛雨"),
    55: ("Dense Drizzle", "大毛毛雨"),
    61: ("Slight Rain", "小雨"),
    63: ("Moderate Rain", "中雨"),
    65: ("Heavy Rain", "大雨"),
    71: ("Slight Snow", "小雪"),
    73: ("Moderate Snow", "中雪"),
    75: ("Heavy Snow", "大雪"),
    80: ("Slight Rain Showers", "小陣雨"),
    81: ("Moderate Rain Showers", "陣雨"),
    82: ("Violent Rain Showers", "暴陣雨"),
    85: ("Slight Snow Showers", "小雪陣"),
    86: ("Heavy Snow Showers", "大雪陣"),
    95: ("Thunderstorm", "雷陣雨"),
    96: ("Thunderstorm with Slight Hail", "雷陣雨伴小冰雹"),
    99: ("Thunderstorm with Heavy Hail", "雷陣雨伴大冰雹"),
}


@dataclass
class CityResult:
    name: str
    country: str
    latitude: float
    longitude: float
    admin1: str = ""

    def display_name(self) -> str:
        parts = [self.name]
        if self.admin1:
            parts.append(self.admin1)
        parts.append(self.country)
        return ", ".join(parts)


@dataclass
class WeatherData:
    temperature: float
    humidity: float
    weather_code: int

    @property
    def weather_en(self) -> str:
        return WMO_WEATHER_CODES.get(self.weather_code, ("Unknown", "未知"))[0]

    @property
    def weather_zh(self) -> str:
        return WMO_WEATHER_CODES.get(self.weather_code, ("Unknown", "未知"))[1]

    def format_line(self) -> str:
        return f"Weather: {self.weather_en}, {self.temperature:.0f}\u00b0C / {self.humidity:.0f}% RH"


class WeatherService:
    GEOCODING_URL = "https://geocoding-api.open-meteo.com/v1/search"
    FORECAST_URL = "https://api.open-meteo.com/v1/forecast"
    HISTORICAL_FORECAST_URL = "https://historical-forecast-api.open-meteo.com/v1/forecast"
    ARCHIVE_URL = "https://archive-api.open-meteo.com/v1/archive"

    @staticmethod
    def _build_ssl_context() -> ssl.SSLContext:
        """Prefer a bundled CA file so frozen builds can verify HTTPS reliably."""
        if certifi is not None:
            try:
                cafile = certifi.where()
                if cafile and os.path.exists(cafile):
                    return ssl.create_default_context(cafile=cafile)
            except Exception:
                pass
        return ssl.create_default_context()

    @classmethod
    def _fetch_json(cls, url: str) -> dict:
        req = urllib.request.Request(url, headers={"User-Agent": "ReportHelper/2.0"})
        try:
            with urllib.request.urlopen(req, timeout=10, context=cls._build_ssl_context()) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except HTTPError as exc:
            raise RuntimeError(f"天氣服務回應錯誤（HTTP {exc.code}）") from exc
        except ssl.SSLError as exc:
            raise RuntimeError(f"天氣服務 SSL 驗證失敗：{exc}") from exc
        except URLError as exc:
            reason = getattr(exc, "reason", exc)
            raise RuntimeError(f"無法連線到天氣服務：{reason}") from exc
        except json.JSONDecodeError as exc:
            raise RuntimeError("天氣服務回傳資料格式錯誤") from exc

    @classmethod
    def search_city(cls, name: str, country_code: str = "") -> list:
        """Search for a place by name.

        Strategy (to maximise recall especially for non-Latin scripts):
        1. Try exact query WITH country_code filter.
        2. If zero results, retry WITHOUT country_code (broader search).
        Note: 'language' param is intentionally omitted — it only controls
        the display language of results and can reduce recall for CJK input.
        """

        def _query(nm: str, cc: str = "") -> list:
            params: dict = {"name": nm, "count": 10}
            if cc:
                params["country_code"] = cc.upper()
            data = cls._fetch_json(
                f"{cls.GEOCODING_URL}?{urllib.parse.urlencode(params)}"
            )
            out = []
            for item in data.get("results", []):
                out.append(CityResult(
                    name=item.get("name", ""),
                    country=item.get("country", ""),
                    latitude=item.get("latitude", 0),
                    longitude=item.get("longitude", 0),
                    admin1=item.get("admin1", ""),
                ))
            return out

        # Pass 1: with country filter
        results = _query(name, country_code)

        # Pass 2 fallback: drop country filter if nothing came back
        if not results and country_code:
            results = _query(name)

        # Pass 3 fallback: if input is CJK, try mapped English name
        if not results and re.search(r'[\u4e00-\u9fff\uac00-\ud7af\u3040-\u30ff]', name):
            en_name = _CJK_ALIAS.get(name.strip())
            if en_name:
                results = _query(en_name, country_code)
                if not results and country_code:
                    results = _query(en_name)

        return results

    @classmethod
    def fetch_weather(cls, lat: float, lon: float, date: datetime) -> Optional[WeatherData]:
        today = datetime.today().date()
        target = date.date() if isinstance(date, datetime) else date
        diff = (target - today).days

        date_str = target.strftime("%Y-%m-%d")

        if diff >= -1 and diff <= 6:
            # Current forecast
            base = cls.FORECAST_URL
        elif diff >= -5:
            # Historical forecast (recent days)
            base = cls.HISTORICAL_FORECAST_URL
        else:
            # Archive (older data)
            base = cls.ARCHIVE_URL

        params = urllib.parse.urlencode({
            "latitude": lat,
            "longitude": lon,
            "daily": "temperature_2m_mean,relative_humidity_2m_mean,weather_code",
            "start_date": date_str,
            "end_date": date_str,
            "timezone": "auto",
        })
        url = f"{base}?{params}"

        data = cls._fetch_json(url)
        daily = data.get("daily", {})
        temp_list = daily.get("temperature_2m_mean", [])
        hum_list = daily.get("relative_humidity_2m_mean", [])
        code_list = daily.get("weather_code", [])

        if temp_list and hum_list and code_list:
            return WeatherData(
                temperature=temp_list[0],
                humidity=hum_list[0],
                weather_code=int(code_list[0]),
            )

        return None
