import openmeteo_requests
import pandas as pd
import requests_cache
from retry_requests import retry
import time

cache_session = requests_cache.CachedSession('.cache', expire_after=-1)
retry_session = retry(cache_session, retries=5, backoff_factor=0.2)
openmeteo = openmeteo_requests.Client(session=retry_session)

url = "https://archive-api.open-meteo.com/v1/archive"

# Chia theo từng năm
date_ranges = [
    ("2020-01-01", "2020-12-31"),
    ("2021-01-01", "2021-12-31"),
    ("2022-01-01", "2022-12-31"),
    ("2023-01-01", "2023-12-31"),
    ("2024-01-01", "2024-12-31"),
    ("2025-01-01", "2025-12-31"),
    ("2026-01-01", "2026-05-23"),
]

all_hourly = []
all_daily = []

for start, end in date_ranges:
    print(f"Đang crawl: {start} → {end}")
    params = {
        "latitude": 20.4737,
        "longitude": 106.0229,
        "start_date": start,
        "end_date": end,
        "daily": ["weather_code", "temperature_2m_mean", "temperature_2m_max", "temperature_2m_min", "apparent_temperature_mean", "apparent_temperature_max", "apparent_temperature_min", "rain_sum", "precipitation_sum", "precipitation_hours", "sunshine_duration", "daylight_duration"],
        "hourly": ["temperature_2m", "relative_humidity_2m", "apparent_temperature", "precipitation", "rain", "weather_code", "cloud_cover", "wind_gusts_10m"],
        "timezone": "Asia/Bangkok",
    }

    try:
        responses = openmeteo.weather_api(url, params=params)
        response = responses[0]

        # Hourly
        hourly = response.Hourly()
        hourly_data = {
            "date": pd.date_range(
                start=pd.to_datetime(hourly.Time(), unit="s", utc=True),
                end=pd.to_datetime(hourly.TimeEnd(), unit="s", utc=True),
                freq=pd.Timedelta(seconds=hourly.Interval()),
                inclusive="left"
            ).tz_convert(response.Timezone().decode())
        }
        hourly_data["temperature_2m"] = hourly.Variables(0).ValuesAsNumpy()
        hourly_data["relative_humidity_2m"] = hourly.Variables(1).ValuesAsNumpy()
        hourly_data["apparent_temperature"] = hourly.Variables(2).ValuesAsNumpy()
        hourly_data["precipitation"] = hourly.Variables(3).ValuesAsNumpy()
        hourly_data["rain"] = hourly.Variables(4).ValuesAsNumpy()
        hourly_data["weather_code"] = hourly.Variables(5).ValuesAsNumpy()
        hourly_data["cloud_cover"] = hourly.Variables(6).ValuesAsNumpy()
        hourly_data["wind_gusts_10m"] = hourly.Variables(7).ValuesAsNumpy()
        all_hourly.append(pd.DataFrame(hourly_data))

        # Daily
        daily = response.Daily()
        daily_data = {
            "date": pd.date_range(
                start=pd.to_datetime(daily.Time(), unit="s", utc=True),
                end=pd.to_datetime(daily.TimeEnd(), unit="s", utc=True),
                freq=pd.Timedelta(seconds=daily.Interval()),
                inclusive="left"
            ).tz_convert(response.Timezone().decode())
        }
        daily_data["weather_code"] = daily.Variables(0).ValuesAsNumpy()
        daily_data["temperature_2m_mean"] = daily.Variables(1).ValuesAsNumpy()
        daily_data["temperature_2m_max"] = daily.Variables(2).ValuesAsNumpy()
        daily_data["temperature_2m_min"] = daily.Variables(3).ValuesAsNumpy()
        daily_data["apparent_temperature_mean"] = daily.Variables(4).ValuesAsNumpy()
        daily_data["apparent_temperature_max"] = daily.Variables(5).ValuesAsNumpy()
        daily_data["apparent_temperature_min"] = daily.Variables(6).ValuesAsNumpy()
        daily_data["rain_sum"] = daily.Variables(7).ValuesAsNumpy()
        daily_data["precipitation_sum"] = daily.Variables(8).ValuesAsNumpy()
        daily_data["precipitation_hours"] = daily.Variables(9).ValuesAsNumpy()
        daily_data["sunshine_duration"] = daily.Variables(10).ValuesAsNumpy()
        daily_data["daylight_duration"] = daily.Variables(11).ValuesAsNumpy()
        all_daily.append(pd.DataFrame(daily_data))

        print(f"✅ Xong {start} → {end}")
        time.sleep(1)  # Nghỉ 1 giây tránh spam server

    except Exception as e:
        print(f"❌ Lỗi {start} → {end}: {e}")

# Ghép tất cả lại
hourly_dataframe = pd.concat(all_hourly, ignore_index=True)
daily_dataframe = pd.concat(all_daily, ignore_index=True)

# Fix timezone
hourly_dataframe["date"] = hourly_dataframe["date"].dt.tz_localize(None)
daily_dataframe["date"] = daily_dataframe["date"].dt.tz_localize(None)

# Xuất CSV
hourly_dataframe.to_csv("HN_hourly.csv", index=False, encoding="utf-8-sig")
daily_dataframe.to_csv("HN_daily.csv", index=False, encoding="utf-8-sig")

# Xuất Excel
try:
    hourly_dataframe.to_excel("HN_hourly.xlsx", index=False)
    daily_dataframe.to_excel("HN_daily.xlsx", index=False)
    print("🎉 HOÀN THÀNH! Đã tạo thành công file Excel và CSV.")
except ImportError:
    print("⚠️ Thiếu thư viện openpyxl. Chạy: pip install openpyxl")