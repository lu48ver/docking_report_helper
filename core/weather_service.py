import json
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Optional


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
    def _fetch_json(url: str) -> dict:
        req = urllib.request.Request(url, headers={"User-Agent": "ReportHelper/2.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))

    @classmethod
    def search_city(cls, name: str) -> list:
        params = urllib.parse.urlencode({
            "name": name,
            "count": 5,
            "language": "zh",
        })
        url = f"{cls.GEOCODING_URL}?{params}"
        data = cls._fetch_json(url)

        results = []
        for item in data.get("results", []):
            results.append(CityResult(
                name=item.get("name", ""),
                country=item.get("country", ""),
                latitude=item.get("latitude", 0),
                longitude=item.get("longitude", 0),
                admin1=item.get("admin1", ""),
            ))
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

        try:
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
        except Exception:
            pass

        return None
