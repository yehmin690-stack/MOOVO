# ============================================================
# MOOVO / TDX 斗六市 98 站正式即時資料蒐集程式
#
# 功能：
# 1. 第一次執行：
#       - 抓雲林 Station API
#       - 依 MOOVO 官方斗六 98 站名單辨識
#       - 建立 master/douliu_station_master_full.csv
#
# 2. 後續執行：
#       - 直接讀取 98 站完整 master
#       - 只呼叫 Availability API
#
# 3. 每次建立：
#       data/raw/YYYY/MM/DD/
#       douliu_realtime_YYYYMMDD_HHMMSS.csv
#
# 4. 欄位：
#       缺車狀態
#       抓取時間
#       星期
#       週末
#       尖峰
#       經緯度
#       與斗六火車站距離
#       站點設施服務類型
#       氣象欄位預留
#
# GitHub Secrets：
#       TDX_CLIENT_ID
#       TDX_CLIENT_SECRET
# ============================================================

import os
import math
import unicodedata
import requests
import pandas as pd

from datetime import datetime
from zoneinfo import ZoneInfo


# ============================================================
# 0. 基本設定
# ============================================================

CITY = "YunlinCounty"

ORIGINAL_MASTER_FILE = (
    "master/douliu_30_station_master.xlsx"
)

FULL_MASTER_FILE = (
    "master/douliu_station_master_full.csv"
)

EXPECTED_DOULIU_STATIONS = 98

# 第一次建立 master 時要求完整 98 站。
# 如果名稱對不到，就停止，避免正式資料只收部分站點。
STRICT_98_STATIONS = True


# ============================================================
# 1. 斗六火車站座標
# ============================================================

DOULIU_RAIL_LAT = 23.711286
DOULIU_RAIL_LON = 120.540981


# ============================================================
# 2. MOOVO 官方斗六 98 站
# ============================================================

OFFICIAL_DOULIU_STATIONS = [

    "斗六棒球場",
    "斗六國中",
    "斗六國小",
    "斗六水岸公園",
    "雲林科技大學",
    "雲林縣體育場",
    "中山紀念公園",
    "行啟紀念館",
    "雲林縣政府",
    "南聖宮",

    "鎮西國小",
    "公誠公園",
    "斗六民生公園",
    "斗六高中",
    "成大醫院",
    "斗六家商",
    "斗六家樂福",
    "雲林科技大學(龍潭路)",
    "斗六火車站前站",
    "斗六火車站後站",

    "龍潭社區活動中心",
    "雲林溪美食廣場",
    "水資源二號公園",
    "斗六觀光夜市",
    "嘉東里",
    "勞工育樂中心",
    "石榴車站",
    "福天吊橋",
    "咬狗籃球場",
    "石榴國中",

    "石榴國小",
    "長和宮",
    "崙峰里",
    "久安里(三殿宮)",
    "正心中學",
    "經濟部斗六產業園區服務中心",
    "雲林國中",
    "斗六大智公園",
    "波麗士公園",
    "菜公庄",

    "斗六承天宮",
    "斗六鎮東公園",
    "斗六文德公園",
    "斗六永樂公園",
    "斗六消防局第一大隊斗六消防分局",
    "明德里社區多功能活動中心",
    "萬年庄",
    "中南里",
    "十三里社區多功能活動中心",
    "榴中福德宮",

    "德安社區",
    "福興路",
    "斗六玄太宮(將爺廟)",
    "斗六德龍宮",
    "南京西路",
    "明德北路與保源六街路口",
    "斗六保清宮",
    "保長國小",
    "龍潭鄰里公園",
    "古坑綠色隧道公園",

    "長安派出所",
    "長平里",
    "虎溪豐泰",
    "虎尾溪福龍宮",
    "崙仔集會所",
    "保竹李天宮",
    "保長路",
    "九老爺(久安南路)",
    "江厝路",
    "瓦厝玄安宮",

    "頂柴福德宮",
    "東市場停車場",
    "溝壩里",
    "雲林縣政府(文化觀光處)",
    "福懋二廠",
    "福懋一廠",
    "斗六名門",
    "斗六家商(擴站)",
    "棒球場停車場",
    "福德街停車場",

    "湖山里(玉當山廟前)",
    "公誠國小",
    "明德北路一段",
    "溪州里",
    "三公園區",
    "璞樸莊園",
    "十三里(聖德宮)",
    "十三里(頂水尾口福德宮)",
    "十三里(武德宮)",
    "湖山水庫梅南路",

    "八德里(文化路646巷)",
    "文殊寺",
    "警察局斗六分局(台1丁)",
    "成功里(台3線)",
    "寧濟御庭",
    "石榴班振王宮",
    "榴南公園",
    "林頭北天宮"
]


if len(OFFICIAL_DOULIU_STATIONS) != 98:
    raise RuntimeError(
        "程式內的斗六官方站點名單不是98站，請停止執行。"
    )


# ============================================================
# 3. 站名正規化
# ============================================================

def normalize_name(value):

    if pd.isna(value):
        return ""

    value = str(value)

    # 全形 / 半形統一
    value = unicodedata.normalize(
        "NFKC",
        value
    )

    # 去除空白
    value = "".join(
        value.split()
    )

    return value.strip()


OFFICIAL_NAME_MAP = {
    normalize_name(name): name
    for name in OFFICIAL_DOULIU_STATIONS
}

OFFICIAL_NAME_SET = set(
    OFFICIAL_NAME_MAP.keys()
)


# ============================================================
# 4. Haversine 距離
# ============================================================

def haversine_km(
    lat1,
    lon1,
    lat2,
    lon2
):

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

    c = 2 * math.asin(
        math.sqrt(a)
    )

    return 6371.0 * c


def distance_band(x):

    if pd.isna(x):
        return pd.NA

    if x <= 0.75:
        return "<=0.75km"

    if x <= 1.5:
        return "0.75-1.5km"

    return ">1.5km"


# ============================================================
# 5. 讀取 GitHub Secrets
# ============================================================

CLIENT_ID = os.getenv(
    "TDX_CLIENT_ID"
)

CLIENT_SECRET = os.getenv(
    "TDX_CLIENT_SECRET"
)

if not CLIENT_ID or not CLIENT_SECRET:

    raise RuntimeError(
        "找不到 TDX_CLIENT_ID 或 "
        "TDX_CLIENT_SECRET"
    )


# ============================================================
# 6. 取得 TDX Token
# ============================================================

TOKEN_URL = (
    "https://tdx.transportdata.tw/"
    "auth/realms/TDXConnect/"
    "protocol/openid-connect/token"
)

token_response = requests.post(
    TOKEN_URL,
    data={
        "grant_type":
            "client_credentials",

        "client_id":
            CLIENT_ID,

        "client_secret":
            CLIENT_SECRET
    },
    timeout=30
)

print(
    "Token HTTP Status:",
    token_response.status_code
)

if token_response.status_code != 200:

    print(
        token_response.text[:1000]
    )

    raise RuntimeError(
        "TDX Token 取得失敗"
    )

access_token = (
    token_response
    .json()["access_token"]
)

print(
    "TDX Token 取得成功"
)


# ============================================================
# 7. API 設定
# ============================================================

STATION_URL = (
    "https://tdx.transportdata.tw/"
    "api/basic/v2/"
    f"Bike/Station/City/{CITY}"
)

AVAILABILITY_URL = (
    "https://tdx.transportdata.tw/"
    "api/basic/v2/"
    f"Bike/Availability/City/{CITY}"
)

headers = {

    "Authorization":
        f"Bearer {access_token}",

    "Accept":
        "application/json",

    "Accept-Encoding":
        "gzip"
}

params = {

    "$top": 1000,

    "$format": "JSON"
}


# ============================================================
# 8. TDX GET
# ============================================================

def get_tdx_json(
    url,
    label
):

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=60
    )

    print(
        f"{label} HTTP Status:",
        response.status_code
    )

    if response.status_code != 200:

        print(
            response.text[:1500]
        )

        raise RuntimeError(
            f"{label} API 取得失敗"
        )

    data = response.json()

    print(
        f"{label} 筆數:",
        len(data)
    )

    return data


# ============================================================
# 9. 建立斗六98站完整主檔
#
# 第一次才執行 Station API
# ============================================================

def build_full_station_master():

    print(
        "\n"
        "====================================="
    )

    print(
        "第一次執行：建立斗六98站完整 master"
    )

    print(
        "====================================="
    )


    # --------------------------------------------------------
    # 抓雲林 Station
    # --------------------------------------------------------

    station_json = get_tdx_json(
        STATION_URL,
        "Station"
    )


    # --------------------------------------------------------
    # 整理 Station
    # --------------------------------------------------------

    rows = []

    for x in station_json:

        station_name = (
            x.get("StationName")
            or {}
        )

        station_address = (
            x.get("StationAddress")
            or {}
        )

        position = (
            x.get("StationPosition")
            or {}
        )

        name = station_name.get(
            "Zh_tw"
        )

        rows.append({

            "StationUID":
                x.get("StationUID"),

            "StationID":
                x.get("StationID"),

            "StationName":
                name,

            "_normalized_name":
                normalize_name(name),

            "StationAddress":
                station_address.get(
                    "Zh_tw"
                ),

            "longitude":
                position.get(
                    "PositionLon"
                ),

            "latitude":
                position.get(
                    "PositionLat"
                ),

            "BikesCapacity":
                x.get(
                    "BikesCapacity"
                ),

            "StationSrcUpdateTime":
                x.get(
                    "SrcUpdateTime"
                ),

            "StationUpdateTime":
                x.get(
                    "UpdateTime"
                )
        })


    station_df = pd.DataFrame(
        rows
    )


    # --------------------------------------------------------
    # 用官方98站站名篩選
    # 不再看地址
    # --------------------------------------------------------

    douliu_df = station_df[
        station_df[
            "_normalized_name"
        ].isin(
            OFFICIAL_NAME_SET
        )
    ].copy()


    # --------------------------------------------------------
    # 官方站名重新帶回
    # --------------------------------------------------------

    douliu_df[
        "OfficialStationName"
    ] = douliu_df[
        "_normalized_name"
    ].map(
        OFFICIAL_NAME_MAP
    )


    # --------------------------------------------------------
    # 防止重複
    # --------------------------------------------------------

    duplicate_names = (
        douliu_df[
            "_normalized_name"
        ]
        .value_counts()
    )

    duplicate_names = (
        duplicate_names[
            duplicate_names > 1
        ]
    )

    if not duplicate_names.empty:

        print(
            "\n⚠️ 發現重複 StationName："
        )

        print(
            duplicate_names
        )


    douliu_df = (
        douliu_df
        .drop_duplicates(
            subset=[
                "_normalized_name"
            ],
            keep="last"
        )
    )


    # --------------------------------------------------------
    # 找出沒有對到 TDX 的官方站
    # --------------------------------------------------------

    found_names = set(
        douliu_df[
            "_normalized_name"
        ]
        .tolist()
    )

    missing_names = [

        OFFICIAL_NAME_MAP[x]

        for x in OFFICIAL_NAME_SET

        if x not in found_names
    ]


    print(
        "\nMOOVO 官方應有：",
        EXPECTED_DOULIU_STATIONS
    )

    print(
        "TDX 成功對應：",
        len(douliu_df)
    )


    if missing_names:

        print(
            "\n⚠️ 以下官方站名沒有在TDX Station中找到："
        )

        for name in sorted(
            missing_names
        ):

            print(
                " -",
                name
            )


    # --------------------------------------------------------
    # 正式模式：首次建立 master 必須對到98站
    # --------------------------------------------------------

    if (
        STRICT_98_STATIONS
        and len(douliu_df)
        != EXPECTED_DOULIU_STATIONS
    ):

        print(
            "\n目前 TDX 的所有站名如下，"
            "可用來檢查站名差異："
        )

        for name in sorted(
            station_df[
                "StationName"
            ]
            .dropna()
            .astype(str)
            .tolist()
        ):

            print(
                name
            )

        raise RuntimeError(
            "斗六官方98站沒有全部與TDX對應。"
            "請先修正站名，不要開始正式蒐集。"
        )


    # --------------------------------------------------------
    # 計算距離
    # --------------------------------------------------------

    douliu_df[
        "distance_to_douliu_station_km"
    ] = douliu_df.apply(

        lambda row:
            haversine_km(

                row[
                    "latitude"
                ],

                row[
                    "longitude"
                ],

                DOULIU_RAIL_LAT,

                DOULIU_RAIL_LON
            ),

        axis=1
    )


    douliu_df[
        "distance_to_douliu_station_km"
    ] = pd.to_numeric(

        douliu_df[
            "distance_to_douliu_station_km"
        ],

        errors="coerce"

    ).round(3)


    douliu_df[
        "distance_band_from_douliu_station"
    ] = douliu_df[
        "distance_to_douliu_station_km"
    ].apply(
        distance_band
    )


    # ========================================================
    # 10. 合併原本30站人工服務類型
    # ========================================================

    douliu_df[
        "station_service_type"
    ] = "Unclassified"


    if os.path.exists(
        ORIGINAL_MASTER_FILE
    ):

        print(
            "\n找到原本30站分類主檔"
        )

        old_master = pd.read_excel(
            ORIGINAL_MASTER_FILE,
            sheet_name="Station_Master_30"
        )


        if (
            "station_name"
            in old_master.columns
            and
            "model_station_group"
            in old_master.columns
        ):

            old_master[
                "_normalized_name"
            ] = old_master[
                "station_name"
            ].apply(
                normalize_name
            )


            category_map = dict(
                zip(

                    old_master[
                        "_normalized_name"
                    ],

                    old_master[
                        "model_station_group"
                    ]
                )
            )


            mapped_category = (
                douliu_df[
                    "_normalized_name"
                ]
                .map(
                    category_map
                )
            )


            douliu_df[
                "station_service_type"
            ] = (
                mapped_category
                .fillna(
                    "Unclassified"
                )
            )

        else:

            print(
                "⚠️ 原30站主檔缺少 "
                "station_name 或 "
                "model_station_group 欄位"
            )

    else:

        print(
            "\n⚠️ 找不到原30站分類主檔"
        )


    # --------------------------------------------------------
    # 依官方順序 1~98 排列
    # --------------------------------------------------------

    official_order = {

        normalize_name(name):
            idx + 1

        for idx, name
        in enumerate(
            OFFICIAL_DOULIU_STATIONS
        )
    }


    douliu_df[
        "official_station_order"
    ] = douliu_df[
        "_normalized_name"
    ].map(
        official_order
    )


    douliu_df = (
        douliu_df
        .sort_values(
            "official_station_order"
        )
        .reset_index(
            drop=True
        )
    )


    # --------------------------------------------------------
    # 主檔欄位
    # --------------------------------------------------------

    master_columns = [

        "official_station_order",

        "StationUID",
        "StationID",

        "OfficialStationName",
        "StationName",
        "StationAddress",

        "longitude",
        "latitude",

        "distance_to_douliu_station_km",
        "distance_band_from_douliu_station",

        "station_service_type",

        "BikesCapacity",

        "StationSrcUpdateTime",
        "StationUpdateTime"
    ]


    douliu_df = douliu_df[
        master_columns
    ]


    # --------------------------------------------------------
    # 建立 master 資料夾
    # --------------------------------------------------------

    os.makedirs(
        "master",
        exist_ok=True
    )


    # --------------------------------------------------------
    # 寫入完整98站 master
    # --------------------------------------------------------

    douliu_df.to_csv(
        FULL_MASTER_FILE,
        index=False,
        encoding="utf-8-sig"
    )


    print(
        "\n✅ 已建立：",
        FULL_MASTER_FILE
    )

    print(
        "完整主檔站數：",
        len(douliu_df)
    )


    return douliu_df


# ============================================================
# 11. 讀取 / 建立斗六完整 master
# ============================================================

if os.path.exists(
    FULL_MASTER_FILE
):

    print(
        "\n找到斗六完整站點主檔："
    )

    print(
        FULL_MASTER_FILE
    )

    master_df = pd.read_csv(
        FULL_MASTER_FILE
    )


    print(
        "完整主檔站數：",
        len(master_df)
    )


    if len(master_df) != 98:

        raise RuntimeError(
            "douliu_station_master_full.csv "
            "不是98站，請檢查主檔。"
        )

else:

    master_df = (
        build_full_station_master()
    )


# ============================================================
# 12. 取得即時 Availability
#
# 正常正式執行時只需要呼叫這個 API
# ============================================================

availability_json = get_tdx_json(
    AVAILABILITY_URL,
    "Availability"
)


# ============================================================
# 13. 成功取得 API 後立即記錄抓取時間
# ============================================================

now_taipei = datetime.now(
    ZoneInfo(
        "Asia/Taipei"
    )
)


# ============================================================
# 14. 整理 Availability
# ============================================================

availability_rows = []

for x in availability_json:

    availability_rows.append({

        "StationUID":
            x.get("StationUID"),

        "ServiceStatus":
            x.get("ServiceStatus"),

        "AvailableRentBikes":
            x.get(
                "AvailableRentBikes"
            ),

        "AvailableReturnBikes":
            x.get(
                "AvailableReturnBikes"
            ),

        "AvailabilitySrcUpdateTime":
            x.get(
                "SrcUpdateTime"
            ),

        "AvailabilityUpdateTime":
            x.get(
                "UpdateTime"
            )
    })


availability_df = pd.DataFrame(
    availability_rows
)


availability_df = (
    availability_df
    .drop_duplicates(
        subset=[
            "StationUID"
        ],
        keep="last"
    )
)


# ============================================================
# 15. 以98站 master 為左表
#
# 非常重要：
# 即使某站即時 API 暫時沒回，
# 仍保留這一站。
# ============================================================

df = master_df.merge(

    availability_df,

    on="StationUID",

    how="left",

    validate="one_to_one"
)


# ============================================================
# 16. 確認永遠是98列
# ============================================================

if len(df) != 98:

    raise RuntimeError(
        f"合併後不是98站，而是 {len(df)} 站"
    )


# ============================================================
# 17. 時間變項
# ============================================================

df[
    "collected_at"
] = now_taipei.strftime(
    "%Y-%m-%d %H:%M:%S"
)


df[
    "capture_time"
] = now_taipei.strftime(
    "%H:%M"
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


weekday_num = (
    now_taipei.weekday()
)


df[
    "weekday"
] = weekday_map[
    weekday_num
]


df[
    "weekend"
] = (
    1
    if weekday_num >= 5
    else 0
)


# ============================================================
# 18. 尖峰變項
#
# 報告設定：
# 07:00~09:00
# 17:00~19:00
# ============================================================

minutes = (
    now_taipei.hour * 60
    +
    now_taipei.minute
)


morning_peak = (
    7 * 60
    <= minutes
    < 9 * 60
)


evening_peak = (
    17 * 60
    <= minutes
    < 19 * 60
)


df[
    "peak"
] = int(
    morning_peak
    or
    evening_peak
)


# ============================================================
# 19. 缺車狀態
#
# AvailableRentBikes = 0 → 1
# AvailableRentBikes > 0 → 0
# 缺資料 → NA
# ============================================================

rent_bikes = pd.to_numeric(

    df[
        "AvailableRentBikes"
    ],

    errors="coerce"
)


df[
    "stockout"
] = pd.Series(

    pd.NA,

    index=df.index,

    dtype="Int64"
)


valid_mask = (
    rent_bikes.notna()
)


df.loc[
    valid_mask,
    "stockout"
] = (
    rent_bikes[
        valid_mask
    ]
    == 0
).astype(int)


# ============================================================
# 20. 氣象欄位先預留
#
# 後續再接中央氣象署資料
# ============================================================

df[
    "rainfall_mm"
] = pd.NA


df[
    "temperature_c"
] = pd.NA


df[
    "wind_speed_ms"
] = pd.NA


# ============================================================
# 21. 正式輸出欄位
# ============================================================

column_order = [

    # -------- 抓取識別 --------
    "collected_at",

    # -------- 站點識別 --------
    "official_station_order",

    "StationUID",
    "StationID",

    "OfficialStationName",
    "StationName",
    "StationAddress",

    # -------- 車位資料 --------
    "AvailableRentBikes",
    "AvailableReturnBikes",

    "ServiceStatus",

    # -------- 依變項 --------
    "stockout",

    # -------- 時間變項 --------
    "capture_time",
    "weekday",
    "weekend",
    "peak",

    # -------- 空間變項 --------
    "longitude",
    "latitude",

    "distance_to_douliu_station_km",

    "distance_band_from_douliu_station",

    "station_service_type",

    # -------- 氣象變項 --------
    "rainfall_mm",
    "temperature_c",
    "wind_speed_ms",

    # -------- 原始檢查欄位 --------
    "BikesCapacity",

    "AvailabilitySrcUpdateTime",
    "AvailabilityUpdateTime"
]


df = df[
    [
        col
        for col in column_order
        if col in df.columns
    ]
]


# ============================================================
# 22. 資料品質檢查
# ============================================================

print(
    "\n"
    "========================================"
)

print(
    "斗六98站即時資料抓取完成"
)

print(
    "========================================"
)


print(
    "抓取時間：",
    now_taipei.isoformat()
)


print(
    "斗六正式站點數：",
    len(df)
)


print(
    "有車位資料：",
    df[
        "AvailableRentBikes"
    ]
    .notna()
    .sum()
)


print(
    "缺少車位資料：",
    df[
        "AvailableRentBikes"
    ]
    .isna()
    .sum()
)


print(
    "缺車站點：",
    (
        df[
            "stockout"
        ]
        == 1
    ).sum()
)


print(
    "未分類服務類型：",
    (
        df[
            "station_service_type"
        ]
        == "Unclassified"
    ).sum()
)


# ============================================================
# 23. 如果有站點沒有即時 Availability
# ============================================================

missing_availability = df[
    df[
        "AvailableRentBikes"
    ].isna()
]


if not missing_availability.empty:

    print(
        "\n⚠️ 以下斗六站本次沒有即時車位資料："
    )

    print(

        missing_availability[
            [
                "official_station_order",
                "StationUID",
                "StationName"
            ]
        ]
        .to_string(
            index=False
        )
    )


# ============================================================
# 24. 顯示前10站
# ============================================================

print(
    "\n===== 前10站 ====="
)


print(

    df[
        [
            "official_station_order",
            "StationName",
            "AvailableRentBikes",
            "AvailableReturnBikes",
            "stockout",
            "distance_to_douliu_station_km",
            "station_service_type"
        ]
    ]
    .head(10)
    .to_string(
        index=False
    )
)


# ============================================================
# 25. 正式資料儲存
#
# data/raw/YYYY/MM/DD/
# ============================================================

year = now_taipei.strftime(
    "%Y"
)

month = now_taipei.strftime(
    "%m"
)

day = now_taipei.strftime(
    "%d"
)


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


print(
    "\n✅ 正式資料已儲存："
)

print(
    output_file
)
