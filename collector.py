import os
import math
import requests
import pandas as pd

from datetime import datetime
from zoneinfo import ZoneInfo


# ============================================================
# 基本設定
# ============================================================

CITY = "YunlinCounty"

MASTER_FILE = "master/douliu_30_station_master.xlsx"
OUTPUT_DIR = "output"

# 沿用目前研究主檔設定的斗六火車站座標
DOULIU_RAIL_LAT = 23.711286
DOULIU_RAIL_LON = 120.540981

os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================
# 讀取 GitHub Secrets
# ============================================================

CLIENT_ID = os.getenv("TDX_CLIENT_ID")
CLIENT_SECRET = os.getenv("TDX_CLIENT_SECRET")

if not CLIENT_ID or not CLIENT_SECRET:
    raise RuntimeError(
        "找不到 TDX_CLIENT_ID 或 TDX_CLIENT_SECRET"
    )


# ============================================================
# 取得 TDX Token
# ============================================================

TOKEN_URL = (
    "https://tdx.transportdata.tw/auth/realms/"
    "TDXConnect/protocol/openid-connect/token"
)

token_response = requests.post(
    TOKEN_URL,
    data={
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET
    },
    timeout=30
)

print("Token HTTP Status:", token_response.status_code)

if token_response.status_code != 200:
    print(token_response.text[:1000])
    raise RuntimeError("TDX Token 取得失敗")

access_token = token_response.json()["access_token"]

print("TDX Token 取得成功")


# ============================================================
# TDX API
# ============================================================

STATION_URL = (
    "https://tdx.transportdata.tw/api/basic/v2/"
    f"Bike/Station/City/{CITY}"
)

AVAILABILITY_URL = (
    "https://tdx.transportdata.tw/api/basic/v2/"
    f"Bike/Availability/City/{CITY}"
)

headers = {
    "Authorization": f"Bearer {access_token}",
    "Accept": "application/json",
    "Accept-Encoding": "gzip"
}

params = {
    "$top": 1000,
    "$format": "JSON"
}


# ============================================================
# 抓 Station
# ============================================================

station_response = requests.get(
    STATION_URL,
    headers=headers,
    params=params,
    timeout=60
)

print("Station HTTP Status:", station_response.status_code)

if station_response.status_code != 200:
    print(station_response.text[:1000])
    raise RuntimeError("Station API 失敗")

station_json = station_response.json()

print("雲林 Station 筆數:", len(station_json))


# ============================================================
# 抓 Availability
# ============================================================

availability_response = requests.get(
    AVAILABILITY_URL,
    headers=headers,
    params=params,
    timeout=60
)

print(
    "Availability HTTP Status:",
    availability_response.status_code
)

if availability_response.status_code != 200:
    print(availability_response.text[:1000])
    raise RuntimeError("Availability API 失敗")

availability_json = availability_response.json()

print(
    "雲林 Availability 筆數:",
    len(availability_json)
)


# ============================================================
# 真正成功抓到資料的時間
# ============================================================

now_taipei = datetime.now(
    ZoneInfo("Asia/Taipei")
)


# ============================================================
# 整理 Station
# ============================================================

station_rows = []

for x in station_json:

    station_name = x.get("StationName") or {}
    station_address = x.get("StationAddress") or {}
    position = x.get("StationPosition") or {}

    station_rows.append({
        "StationUID": x.get("StationUID"),
        "StationID": x.get("StationID"),

        "StationName":
            station_name.get("Zh_tw"),

        "StationAddress":
            station_address.get("Zh_tw"),

        "longitude":
            position.get("PositionLon"),

        "latitude":
            position.get("PositionLat"),

        "BikesCapacity":
            x.get("BikesCapacity"),

        "StationSrcUpdateTime":
            x.get("SrcUpdateTime")
    })

station_df = pd.DataFrame(station_rows)


# ============================================================
# 篩選斗六市
# ============================================================

douliu_station_df = station_df[
    station_df["StationAddress"]
    .fillna("")
    .str.contains("斗六市", regex=False)
].copy()

print(
    "斗六 Station 筆數:",
    len(douliu_station_df)
)

if douliu_station_df.empty:
    raise RuntimeError(
        "沒有找到斗六市站點，停止執行"
    )


# ============================================================
# 整理 Availability
# ============================================================

availability_rows = []

for x in availability_json:

    availability_rows.append({
        "StationUID":
            x.get("StationUID"),

        "ServiceStatus":
            x.get("ServiceStatus"),

        "AvailableRentBikes":
            x.get("AvailableRentBikes"),

        "AvailableReturnBikes":
            x.get("AvailableReturnBikes"),

        "AvailabilitySrcUpdateTime":
            x.get("SrcUpdateTime"),

        "AvailabilityUpdateTime":
            x.get("UpdateTime")
    })

availability_df = pd.DataFrame(
    availability_rows
)

# 保險：若來源意外重複，只保留最後一筆
availability_df = (
    availability_df
    .drop_duplicates(
        subset=["StationUID"],
        keep="last"
    )
)


# ============================================================
# Station + Availability
# ============================================================

df = douliu_station_df.merge(
    availability_df,
    on="StationUID",
    how="left"
)


# ============================================================
# 時間變項
# ============================================================

df["collected_at"] = (
    now_taipei.strftime("%Y-%m-%d %H:%M:%S")
)

df["capture_time"] = (
    now_taipei.strftime("%H:%M")
)

weekday_map = {
    0: "星期一",
    1: "星期二",
    2: "星期三",
    3: "星期四",
    4: "星期五",
    5: "星期六",
    6: "星期日"
}

weekday_num = now_taipei.weekday()

df["weekday"] = weekday_map[weekday_num]

df["weekend"] = (
    1 if weekday_num >= 5 else 0
)

minutes = (
    now_taipei.hour * 60
    + now_taipei.minute
)

is_peak = (
    (7 * 60 <= minutes < 9 * 60)
    or
    (17 * 60 <= minutes < 19 * 60)
)

df["peak"] = int(is_peak)


# ============================================================
# 缺車狀態
# ============================================================

rent_bikes = pd.to_numeric(
    df["AvailableRentBikes"],
    errors="coerce"
)

df["stockout"] = pd.Series(
    pd.NA,
    index=df.index,
    dtype="Int64"
)

valid_mask = rent_bikes.notna()

df.loc[
    valid_mask,
    "stockout"
] = (
    rent_bikes[valid_mask] == 0
).astype(int)


# ============================================================
# 距離計算
# ============================================================

def haversine_km(lat1, lon1, lat2, lon2):

    if pd.isna(lat1) or pd.isna(lon1):
        return pd.NA

    lat1 = math.radians(float(lat1))
    lon1 = math.radians(float(lon1))
    lat2 = math.radians(float(lat2))
    lon2 = math.radians(float(lon2))

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    a = (
        math.sin(dlat / 2) ** 2
        +
        math.cos(lat1)
        * math.cos(lat2)
        * math.sin(dlon / 2) ** 2
    )

    c = 2 * math.asin(math.sqrt(a))

    return 6371.0 * c


df["distance_to_douliu_station_km"] = (
    df.apply(
        lambda row: haversine_km(
            row["latitude"],
            row["longitude"],
            DOULIU_RAIL_LAT,
            DOULIU_RAIL_LON
        ),
        axis=1
    )
)

df["distance_to_douliu_station_km"] = (
    pd.to_numeric(
        df["distance_to_douliu_station_km"],
        errors="coerce"
    )
    .round(3)
)


# ============================================================
# 距離分類
# ============================================================

def distance_band(x):

    if pd.isna(x):
        return pd.NA

    if x <= 0.75:
        return "<=0.75km"

    if x <= 1.5:
        return "0.75-1.5km"

    return ">1.5km"


df["distance_band_from_douliu_station"] = (
    df["distance_to_douliu_station_km"]
    .apply(distance_band)
)


# ============================================================
# 站點服務類型
# ============================================================

if os.path.exists(MASTER_FILE):

    print("找到站點主檔")

    master_df = pd.read_excel(
        MASTER_FILE,
        sheet_name="Station_Master_30"
    )

    master_df["station_name"] = (
        master_df["station_name"]
        .astype(str)
        .str.strip()
    )

    df["StationName"] = (
        df["StationName"]
        .astype(str)
        .str.strip()
    )

    master_type = (
        master_df[
            [
                "station_name",
                "model_station_group"
            ]
        ]
        .drop_duplicates(
            subset=["station_name"]
        )
        .rename(
            columns={
                "station_name":
                    "StationName",

                "model_station_group":
                    "station_service_type"
            }
        )
    )

    df = df.merge(
        master_type,
        on="StationName",
        how="left"
    )

    df["station_service_type"] = (
        df["station_service_type"]
        .fillna("Unclassified")
    )

else:

    print(
        "WARNING: 找不到站點主檔，"
        "所有站點暫標 Unclassified"
    )

    df["station_service_type"] = (
        "Unclassified"
    )


# ============================================================
# 氣象先留空
# ============================================================

df["rainfall_mm"] = pd.NA
df["temperature_c"] = pd.NA
df["wind_speed_ms"] = pd.NA


# ============================================================
# 欄位排序
# ============================================================

columns = [
    "collected_at",

    "StationUID",
    "StationID",
    "StationName",
    "StationAddress",

    "AvailableRentBikes",
    "AvailableReturnBikes",
    "ServiceStatus",
    "stockout",

    "capture_time",
    "weekday",
    "weekend",
    "peak",

    "longitude",
    "latitude",
    "distance_to_douliu_station_km",
    "distance_band_from_douliu_station",
    "station_service_type",

    "rainfall_mm",
    "temperature_c",
    "wind_speed_ms",

    "BikesCapacity",
    "AvailabilitySrcUpdateTime",
    "AvailabilityUpdateTime"
]

df = df[
    [c for c in columns if c in df.columns]
]


# ============================================================
# 品質檢查
# ============================================================

print("\n==============================")
print("斗六即時資料抓取完成")
print("==============================")

print("抓取時間:", now_taipei)
print("斗六站點:", len(df))

print(
    "有車位資料:",
    df["AvailableRentBikes"]
    .notna()
    .sum()
)

print(
    "缺車站點:",
    (df["stockout"] == 1)
    .sum()
)

print(
    "未分類服務類型:",
    (df["station_service_type"] == "Unclassified")
    .sum()
)

print("\n前 10 筆：")

print(
    df[
        [
            "StationName",
            "AvailableRentBikes",
            "stockout",
            "distance_to_douliu_station_km",
            "station_service_type"
        ]
    ].head(10).to_string(index=False)
)


# ============================================================
# 正式資料儲存位置
# data/raw/YYYY/MM/DD/
# 每一次抓取建立一個獨立快照 CSV
# ============================================================

year = now_taipei.strftime("%Y")
month = now_taipei.strftime("%m")
day = now_taipei.strftime("%d")

output_dir = os.path.join(
    "data",
    "raw",
    year,
    month,
    day
)

os.makedirs(
    output_dir,
    exist_ok=True
)

timestamp = now_taipei.strftime(
    "%Y%m%d_%H%M%S"
)

output_file = os.path.join(
    output_dir,
    f"douliu_realtime_{timestamp}.csv"
)

df.to_csv(
    output_file,
    index=False,
    encoding="utf-8-sig"
)

print("\n==============================")
print("正式資料已儲存")
print("==============================")
print(output_file)
