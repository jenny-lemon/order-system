import os
import re
import json
import time
from datetime import datetime, timedelta, timezone
from collections import defaultdict

import requests
import pandas as pd
from bs4 import BeautifulSoup

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from accounts import ACCOUNTS
from env import (
    ENV,
    BASE_URL_DEV,
    BASE_URL_PROD,
    GOOGLE_SHEET_ID,
    ENABLE_GCAL_COLOR_SYNC,
    GOOGLE_CALENDAR_MAP,
    GOOGLE_SERVICE_ACCOUNT_FILE,
    COLOR_PURPLE,
    COLOR_YELLOW,
    REQUEST_DELAY,
    ORDER_PREFIX_DEV,
    ORDER_PREFIX_PROD,
)

try:
    import streamlit as st
except Exception:
    st = None

try:
    from env import GOOGLE_MAPS_API_KEY
except Exception:
    GOOGLE_MAPS_API_KEY = ""


# =========================
# 環境
# =========================
if ENV == "dev":
    BASE_URL = BASE_URL_DEV
    ORDER_PREFIX = ORDER_PREFIX_DEV
else:
    BASE_URL = BASE_URL_PROD
    ORDER_PREFIX = ORDER_PREFIX_PROD

LOGIN_URL = f"{BASE_URL}/login"
BOOKING_URL = f"{BASE_URL}/booking/stored_value_routine"
PURCHASE_URL = f"{BASE_URL}/purchase"
GET_MEMBER_URL = f"{BASE_URL}/ajax/get_member"
CHECK_CONTAIN_URL = f"{BASE_URL}/ajax/check_contain"
CALCULATE_HOUR_URL = f"{BASE_URL}/ajax/calculate_hour"
GET_SECTION_URL = f"{BASE_URL}/ajax/get_section"
MAIL_SUCCESS_URL = f"{BASE_URL}/purchase/mail_success/{{order_no}}"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
}

MAIL_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "User-Agent": "Mozilla/5.0",
    "Referer": PURCHASE_URL,
}

CLEAN_TYPE_MAP = {
    "居家清潔": "1",
    "辦公室清潔": "2",
    "裝修細清": "3",
}

ORDER_NO_REGEX = r"(LC|TT)\d+"

STANDARD_SLOTS = [
    "08:30-12:30",
    "09:00-11:00",
    "09:00-12:00",
    "14:00-16:00",
    "14:00-17:00",
    "14:00-18:00",
]

KNOWN_SERVICE_STATUS = [
    "已處理",
    "未處理",
    "處理中",
    "已完成",
    "已取消",
    "待處理",
]

print("=== 儲值金系統設定.py 版本：2026-04-23-final-v5 ===")


# =========================
# 基本工具
# =========================
def is_blank(value):
    return str(value).strip() in ("", "nan", "None")


def normalize_phone(phone_value):
    phone = str(phone_value).strip()
    phone = phone.replace(".0", "")
    phone = re.sub(r"\D", "", phone)
    if len(phone) == 9:
        phone = "0" + phone
    return phone


def normalize_text_for_parse(text):
    return re.sub(r"\s+", "", str(text or ""))


def normalize_addr_for_match(addr):
    return re.sub(r"\s+", "", str(addr or "")).strip()


def parse_date_value(date_value):
    if isinstance(date_value, pd.Timestamp):
        return date_value.to_pydatetime()

    text = str(date_value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass

    raise Exception(f"無法解析日期: {date_value}")


def normalize_sheet_date(date_value):
    return parse_date_value(date_value).strftime("%Y-%m-%d")


def parse_time_slot(start_time_str, end_time_str):
    def to_hm(t):
        parts = str(t).strip().split(":")
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return h, m

    sh, sm = to_hm(start_time_str)
    eh, em = to_hm(end_time_str)
    return sh, sm, eh, em


def calc_hours_from_time(start_time_str, end_time_str):
    """
    規則：
    10:00-12:00 -> 2 小時
    09:00-16:00 -> 7 - 1 = 6 小時
    09:00-18:00 -> 9 - 1 = 8 小時
    只要跨過 12:00~13:00 午休就扣 1 小時
    """
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)

    start_minutes = sh * 60 + sm
    end_minutes = eh * 60 + em

    total_hours = (end_minutes - start_minutes) / 60.0
    if total_hours <= 0:
        return None

    lunch_start = 12 * 60
    lunch_end = 13 * 60

    crosses_lunch = start_minutes < lunch_start and end_minutes > lunch_end
    if crosses_lunch:
        total_hours -= 1

    return total_hours


def normalize_period_text(start_time_str, end_time_str):
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)
    return f"{sh:02d}:{sm:02d}-{eh:02d}:{em:02d}"


def display_period_text(start_time_str, end_time_str):
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)
    return f"{sh:02d}:{sm:02d} - {eh:02d}:{em:02d}"


def normalize_sheet_period(start_time_str, end_time_str):
    return normalize_period_text(start_time_str, end_time_str)


def build_target_slot_from_row(row):
    date_part = normalize_sheet_date(row["日期"])
    period_part = normalize_sheet_period(row["開始時間"], row["結束時間"])
    return f"{date_part}_{period_part}"


def slot_duration_hours(slot_text):
    start_text, end_text = slot_text.split("-")
    return calc_hours_from_time(start_text, end_text)


def slot_start_hour(slot_text):
    start_text = slot_text.split("-")[0]
    return int(start_text.split(":")[0])


def is_morning_slot(slot_text):
    return slot_start_hour(slot_text) < 12


def map_to_system_slot(start_time_str, end_time_str):
    original_slot = normalize_period_text(start_time_str, end_time_str)
    actual_hours = calc_hours_from_time(start_time_str, end_time_str)

    if actual_hours is None:
        raise Exception(f"無法解析服務時段: {start_time_str}-{end_time_str}")

    # 🔥 強制規則：10-12 → 09-11
    if original_slot == "10:00-12:00":
        return {
            "original_slot": original_slot,
            "system_slot": "09:00-11:00",
            "need_note": True,
            "sms_time": original_slot,
            "customer_time_note": f"服務時間：{original_slot}",
        }

    # 標準時段
    if original_slot in STANDARD_SLOTS:
        return {
            "original_slot": original_slot,
            "system_slot": original_slot,
            "need_note": False,
            "sms_time": "",
            "customer_time_note": "",
        }

    # 非標準 → 找最接近
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)

    matched_slot = None
    for slot in STANDARD_SLOTS:
        if abs(slot_duration_hours(slot) - actual_hours) < 1e-9:
            matched_slot = slot
            break

    if not matched_slot:
        raise Exception(f"找不到可對應的系統時段：{original_slot}")

    return {
        "original_slot": original_slot,
        "system_slot": matched_slot,
        "need_note": True,
        "sms_time": original_slot,
        "customer_time_note": f"服務時間：{original_slot}",
    }


def parse_service_human_hour(cell_value, start_time_str=None, end_time_str=None):
    default_people = 2

    if pd.isna(cell_value) or str(cell_value).strip() == "":
        people = default_people
        hours = calc_hours_from_time(start_time_str, end_time_str)
        return people, hours

    text = str(cell_value).strip()
    people_match = re.search(r"(\d+)\s*人", text)
    hour_match = re.search(r"(\d+(?:\.\d+)?)\s*小時", text)

    people = int(people_match.group(1)) if people_match else default_people
    hours = float(hour_match.group(1)) if hour_match else calc_hours_from_time(start_time_str, end_time_str)

    return people, hours


def normalize_hours_text(cell_value, start_time_str=None, end_time_str=None):
    people, hours = parse_service_human_hour(cell_value, start_time_str, end_time_str)
    if hours is None:
        return f"{people}人"
    if float(hours).is_integer():
        htxt = f"{int(hours)}小時"
    else:
        htxt = f"{hours}小時"
    return f"{people}人{htxt}"


def build_group_key(row):
    normalized_human_hour = normalize_hours_text(
        row["服務人時"],
        row["開始時間"],
        row["結束時間"],
    )
    return (
        str(row["姓名"]).strip(),
        normalize_phone(row["電話"]),
        str(row["地址"]).strip(),
        str(row["購買項目"]).strip(),
        normalize_period_text(row["開始時間"], row["結束時間"]),
        normalized_human_hour,
        str(row["備註"]).strip(),
    )


def get_region_by_address(address, accounts_config):
    for region, config in accounts_config.items():
        keywords = config.get("address_keywords", [])
        if keywords:
            for kw in keywords:
                if kw in address:
                    return region
        else:
            if region == "台北" and ("台北市" in address or "新北市" in address):
                return region
            if region == "台中" and "台中市" in address:
                return region
            if region == "桃園" and "桃園" in address:
                return region
            if region == "新竹" and ("新竹市" in address or "新竹縣" in address):
                return region
            if region == "高雄" and ("高雄市" in address or "台南市" in address):
                return region
    return None


def should_process_row(row):
    status = str(row.get("狀態", "")).strip()
    order_no = row.get("訂單編號", "")
    return status == "未安排" and is_blank(order_no)


def should_create_order(row):
    status = str(row.get("狀態", "")).strip()
    order_no = row.get("訂單編號", "")
    return status == "未安排" and is_blank(order_no)


# =========================
# XYZ / 回填模板
# =========================
def finalize_xyz(meta=None, fallback_fare="0"):
    meta = meta or {}

    staff = str(meta.get("服務人員", "") or "").strip()
    status = str(meta.get("服務狀態", "") or "").strip()
    fare = str(meta.get("車馬費", "") or "").strip()

    if not staff:
        staff = "無人力"
    if not status:
        status = "未處理"
    if not fare:
        fare = str(fallback_fare or "0").strip() or "0"

    return {
        "服務人員": staff,
        "服務狀態": status,
        "車馬費": fare,
    }


def build_row_result(
    order_no="",
    result="失敗",
    reason="",
    no_slot_date="",
    insufficient_date="",
    sms_time="",
    customer_note="",
    confirm_mail="",
    calendar_result="",
    calendar_reason="",
    calendar_old="",
    calendar_new="",
    staff="無人力",
    service_status="未處理",
    fare="0",
):
    xyz = finalize_xyz(
        {
            "服務人員": staff,
            "服務狀態": service_status,
            "車馬費": fare,
        },
        fallback_fare="0",
    )

    return {
        "訂單編號": order_no,
        "結果": result,
        "原因": reason,
        "沒班表日期": no_slot_date,
        "餘額不足未送": insufficient_date,
        "簡訊實際服務時間": sms_time,
        "客人備註": customer_note,
        "確認信": confirm_mail,
        "日曆改色結果": calendar_result,
        "日曆改色原因": calendar_reason,
        "日曆原色": calendar_old,
        "日曆新色": calendar_new,
        "服務人員": xyz["服務人員"],
        "服務狀態": xyz["服務狀態"],
        "車馬費": xyz["車馬費"],
    }


# =========================
# Google 憑證 / Sheet
# =========================
def get_service_account_info():
    if st is not None:
        try:
            if "gcp_service_account" in st.secrets:
                print("✅ 使用 Streamlit secrets[gcp_service_account]")
                return dict(st.secrets["gcp_service_account"])
            if "GOOGLE_SERVICE_ACCOUNT" in st.secrets:
                print("✅ 使用 Streamlit secrets[GOOGLE_SERVICE_ACCOUNT]")
                return dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
        except Exception as e:
            print(f"⚠️ 讀取 Streamlit secrets 失敗: {e}")

    raw_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_json:
        try:
            print("✅ 使用環境變數 GOOGLE_SERVICE_ACCOUNT_JSON")
            return json.loads(raw_json)
        except Exception as e:
            raise Exception(f"GOOGLE_SERVICE_ACCOUNT_JSON 不是合法 JSON：{e}")

    candidate_files = []
    if GOOGLE_SERVICE_ACCOUNT_FILE:
        candidate_files.append(GOOGLE_SERVICE_ACCOUNT_FILE)
    candidate_files.append("google_service_account.json")

    for fp in candidate_files:
        if fp and os.path.exists(fp):
            print(f"✅ 使用本機憑證檔：{fp}")
            with open(fp, "r", encoding="utf-8") as f:
                return json.load(f)

    raise FileNotFoundError(
        "找不到 Google 憑證。請在 Streamlit Secrets 設定 gcp_service_account 或 GOOGLE_SERVICE_ACCOUNT，"
        "或提供 GOOGLE_SERVICE_ACCOUNT_JSON，或放置 google_service_account.json。"
    )


def build_gsheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    service_account_info = get_service_account_info()
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    return gspread.authorize(creds)


def load_worksheet(sheet_name):
    client = build_gsheet_client()
    sh = client.open_by_key(GOOGLE_SHEET_ID)
    ws = sh.worksheet(sheet_name)

    values = ws.get_all_values()
    if not values:
        raise Exception(f"工作表 {sheet_name} 沒有資料")

    headers = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=headers)
    df["__sheet_row__"] = range(2, len(df) + 2)
    return ws, df


def ensure_columns_in_sheet(ws):
    headers = ws.row_values(1)

    required = [
        "簡訊實際服務時間",
        "客人備註",
        "訂單編號",
        "結果",
        "原因",
        "沒班表日期",
        "餘額不足未送",
        "確認信",
        "日曆改色結果",
        "日曆改色原因",
        "日曆原色",
        "日曆新色",
        "服務人員",
        "服務狀態",
        "車馬費",
    ]

    changed = False
    for col in required:
        if col not in headers:
            headers.append(col)
            changed = True

    if changed:
        ws.resize(rows=max(ws.row_count, 1), cols=len(headers))
        ws.update("A1", [headers])

    return headers


def update_sheet_rows(ws, row_results):
    headers = ensure_columns_in_sheet(ws)
    header_index = {h: i + 1 for i, h in enumerate(headers)}

    updates = []
    for row_num, info in row_results.items():
        xyz = finalize_xyz(
            {
                "服務人員": info.get("服務人員", ""),
                "服務狀態": info.get("服務狀態", ""),
                "車馬費": info.get("車馬費", ""),
            },
            fallback_fare="0",
        )
        info["服務人員"] = xyz["服務人員"]
        info["服務狀態"] = xyz["服務狀態"]
        info["車馬費"] = xyz["車馬費"]

        print(f"[DEBUG] write row={row_num} XYZ={xyz}")

        for key, value in info.items():
            if key not in header_index:
                continue
            updates.append({
                "range": gspread.utils.rowcol_to_a1(row_num, header_index[key]),
                "values": [[("" if value is None else str(value))]],
            })

    if updates:
        ws.batch_update(updates)


# =========================
# 後台 API
# =========================
def login(session, email, password):
    resp = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        print(f"取得登入頁失敗: {resp.status_code}")
        return False

    soup = BeautifulSoup(resp.text, "html.parser")
    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        print("登入失敗：找不到 _token")
        return False

    token = token_input.get("value", "").strip()
    if not token:
        print("登入失敗：_token 為空")
        return False

    data = {
        "_token": token,
        "email": email,
        "password": password,
    }

    resp = session.post(LOGIN_URL, data=data, headers=HEADERS, allow_redirects=True)
    if resp.status_code == 200 and "login" not in resp.url.lower():
        return True

    print(f"登入失敗 ({email}): status={resp.status_code}, url={resp.url}")
    return False


def get_csrf_token(session):
    resp = session.get(BOOKING_URL, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        raise Exception(f"取得儲值金訂單頁失敗: {resp.status_code}")

    soup = BeautifulSoup(resp.text, "html.parser")
    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        raise Exception("無法從儲值金訂單頁提取 _token")

    token = token_input.get("value", "").strip()
    if not token:
        raise Exception("_token 為空")

    return token


def get_member(session, phone, token, clean_type_id):
    data = {
        "phone": phone,
        "_token": token,
        "clean_type_id": clean_type_id,
    }

    resp = session.post(GET_MEMBER_URL, data=data, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return None

    try:
        result = resp.json()
    except Exception:
        return None

    if isinstance(result, dict) and result.get("return_code") == "0000" and result.get("member"):
        return result

    return None


def pick_best_address_info(member_payload, target_address):
    member = member_payload.get("member", {}) if isinstance(member_payload, dict) else {}
    member_address_list = member.get("memberAddressList", []) if isinstance(member, dict) else []

    target_norm = normalize_addr_for_match(target_address)

    for item in member_address_list:
        item_addr = str(item.get("address", "")).strip()
        if normalize_addr_for_match(item_addr) == target_norm:
            return {
                "addressId": str(item.get("id", "")).strip(),
                "country_id": item.get("countryId", ""),
                "area_id": item.get("areaId", ""),
                "address": item_addr,
                "lat": item.get("lat", ""),
                "lng": item.get("lng", ""),
                "company_id": item.get("companyId", 1),
                "purchase": item.get("purchase", {}) if isinstance(item.get("purchase"), dict) else {},
            }

    return None


def geocode_address(address):
    if not GOOGLE_MAPS_API_KEY:
        return None, None

    try:
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {
            "address": address,
            "language": "zh-TW",
            "key": GOOGLE_MAPS_API_KEY,
        }
        resp = requests.get(url, params=params, timeout=15)
        if resp.status_code != 200:
            return None, None

        data = resp.json()
        results = data.get("results", [])
        if not results:
            return None, None

        location = results[0].get("geometry", {}).get("location", {})
        lat = location.get("lat")
        lng = location.get("lng")
        if lat is None or lng is None:
            return None, None

        return str(lat), str(lng)
    except Exception:
        return None, None


def check_contain(session, member_id, address, lat, lng, token, clean_type_id):
    data = {
        "memberId": member_id,
        "cleanTypeId": clean_type_id,
        "address": address,
        "lat": lat or "",
        "lng": lng or "",
        "_token": token,
    }

    resp = session.post(CHECK_CONTAIN_URL, data=data, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return None

    try:
        return resp.json()
    except Exception:
        return None


def calculate_hour(session, order_data, token):
    data = order_data.copy()
    data["_token"] = token

    resp = session.post(CALCULATE_HOUR_URL, data=data, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return None

    try:
        return resp.json()
    except Exception:
        return None


def extract_calc_fields(calc_result, fallback_hours=None, fallback_fare=None):
    if not isinstance(calc_result, dict):
        return {
            "hour": str(fallback_hours or ""),
            "price": "0",
            "price_vvip": "0",
            "fare": str(fallback_fare or "0"),
        }

    def pick(*keys, default=""):
        for k in keys:
            v = calc_result.get(k)
            if v not in (None, ""):
                return str(v)
        return str(default)

    return {
        "hour": pick("hour", "clean_hour", default=fallback_hours or ""),
        "price": pick("price", "total_price", "service_price", default="0"),
        "price_vvip": pick("price_vvip", "vvip_price", default="0"),
        "fare": pick("fare", "car_fare", default=fallback_fare or "0"),
    }


def get_section_raw(session, order_data, token, date_slot):
    data = order_data.copy()
    data["_token"] = token
    data["date_list[]"] = date_slot

    resp = session.post(GET_SECTION_URL, data=data, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return ""

    return resp.text


def extract_available_slots_from_section(raw_text):
    """
    優先抓 checkbox value：
    value="2026-05-12_08:30-12:30"
    抓不到時才退回畫面文字解析。
    """
    if not raw_text:
        return []

    html = str(raw_text)
    slots = []

    value_matches = re.findall(
        r'value=["\'](\d{4}-\d{2}-\d{2}_\d{2}:\d{2}-\d{2}:\d{2})["\']',
        html
    )
    for slot in value_matches:
        if slot not in slots:
            slots.append(slot)

    if slots:
        return slots

    normalized = re.sub(r"\s+", "", html)
    text_matches = re.findall(
        r'(\d{4}-\d{2}-\d{2}).{0,12}?(\d{2}:\d{2}-\d{2}:\d{2})',
        normalized
    )
    for date_part, period_part in text_matches:
        slot = f"{date_part}_{period_part}"
        if slot not in slots:
            slots.append(slot)

    return slots


def slot_exists_in_section_response(raw_text, target_slot):
    available_slots = extract_available_slots_from_section(raw_text)
    return target_slot in available_slots


def validate_available_slots(session, order_data, token, date_slots):
    valid_slots = []
    invalid_slots = []

    for slot in date_slots:
        raw = get_section_raw(session, order_data, token, slot)
        available_slots = extract_available_slots_from_section(raw)
        ok = slot in available_slots

        print("[DEBUG] validate slot =", slot)
        print("[DEBUG] available slots =", available_slots[:20])
        print("[DEBUG] matched =", ok)
        print("[DEBUG] raw snippet =", str(raw)[:1000])

        if ok:
            valid_slots.append(slot)
        else:
            invalid_slots.append(slot)

    return valid_slots, invalid_slots


# =========================
# Purchase 頁解析
# =========================
def extract_order_cards_from_purchase_html(html):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    blocks = []
    current = None

    for line in lines:
        if re.fullmatch(ORDER_NO_REGEX, line):
            if current:
                current["joined_text"] = "\n".join(current["lines"])
                blocks.append(current)
            current = {"order_no": line, "lines": [line]}
        elif current:
            current["lines"].append(line)

    if current:
        current["joined_text"] = "\n".join(current["lines"])
        blocks.append(current)

    return blocks


def match_order_from_purchase_page(html, target_date, target_period):
    blocks = extract_order_cards_from_purchase_html(html)
    for block in blocks:
        joined = block.get("joined_text", "\n".join(block["lines"]))
        normalized = normalize_text_for_parse(joined)
        if target_date in normalized and normalize_text_for_parse(target_period) in normalized:
            return block["order_no"]
    return None


def fetch_order_no_by_date_and_period(session, target_date, target_period):
    resp = session.get(PURCHASE_URL, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return None
    return match_order_from_purchase_page(resp.text, target_date, target_period)


def _extract_staff_line(lines):
    joined = "\n".join(lines)
    normalized = normalize_text_for_parse(joined)

    m = re.search(r'([\u4e00-\u9fffA-Za-z0-9]+\(\d+\))X([\u4e00-\u9fffA-Za-z0-9]+\(\d+\))', normalized)
    if m:
        return f"{m.group(1)} X {m.group(2)}"

    for i, line in enumerate(lines):
        t = normalize_text_for_parse(line)
        if re.search(r'\(\d+\)X$', t):
            first = re.sub(r'X$', '', t)
            if i + 1 < len(lines):
                second = normalize_text_for_parse(lines[i + 1])
                if re.search(r'\(\d+\)$', second):
                    return f"{first} X {second}"

    return "無人力"


def _extract_status_line(lines):
    joined = "\n".join(lines)
    normalized = normalize_text_for_parse(joined)

    for status in KNOWN_SERVICE_STATUS:
        if status in normalized:
            return status

    if "未處理" in normalized or ("未" in normalized and "處" in normalized and "理" in normalized):
        return "未處理"
    if "已處理" in normalized or ("已" in normalized and "處" in normalized and "理" in normalized):
        return "已處理"

    return "未處理"


def _extract_fare_line(lines):
    joined = "\n".join(lines)
    normalized = normalize_text_for_parse(joined)

    m = re.search(r'車馬費[：:]?(\d+)', normalized)
    if m:
        return m.group(1)

    return "0"


def _extract_service_date_time(lines):
    service_date = ""
    service_time = ""

    for idx, line in enumerate(lines):
        text = line.strip()
        if re.match(r"\d{4}-\d{2}-\d{2}", text):
            service_date = text[:10]

            for j in range(idx + 1, min(idx + 5, len(lines))):
                nxt = lines[j].strip()
                nxt2 = nxt.replace(" ", "")
                if re.match(r"\d{2}:\d{2}-\d{2}:\d{2}", nxt2):
                    service_time = nxt2
                    break
            break

    return service_date, service_time


def fetch_order_meta_by_order_no(session, order_no):
    resp = session.get(PURCHASE_URL, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return {
            "服務人員": "無人力",
            "服務狀態": "未處理",
            "車馬費": "0",
            "服務日期": "",
            "服務時間": "",
        }

    blocks = extract_order_cards_from_purchase_html(resp.text)
    for block in blocks:
        if block["order_no"] == order_no:
            lines = block.get("lines", [])
            service_date, service_time = _extract_service_date_time(lines)
            staff = _extract_staff_line(lines)
            status = _extract_status_line(lines)
            fare = _extract_fare_line(lines)

            return {
                "服務人員": staff if staff else "無人力",
                "服務狀態": status if status else "未處理",
                "車馬費": fare if fare else "0",
                "服務日期": service_date,
                "服務時間": service_time,
            }

    return {
        "服務人員": "無人力",
        "服務狀態": "未處理",
        "車馬費": "0",
        "服務日期": "",
        "服務時間": "",
    }


def send_confirmation_mail(session, order_no):
    url = MAIL_SUCCESS_URL.format(order_no=order_no)
    resp = session.get(url, headers=MAIL_HEADERS, allow_redirects=True)

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"

    try:
        result = resp.json()
        return True, str(result)
    except Exception:
        return True, resp.text[:200]


# =========================
# Google Calendar
# =========================
def build_gcal_service():
    if not ENABLE_GCAL_COLOR_SYNC:
        return None

    scopes = ["https://www.googleapis.com/auth/calendar"]
    service_account_info = get_service_account_info()
    credentials = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    return build("calendar", "v3", credentials=credentials)


def parse_event_time(dt_str):
    if not dt_str:
        return None
    try:
        return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except Exception:
        try:
            return datetime.strptime(dt_str, "%Y-%m-%d")
        except Exception:
            return None


def color_name_from_id(color_id):
    mapping = {
        "1": "薰衣草紫",
        "2": "鼠尾草綠",
        "3": "葡萄紫",
        "4": "火鶴紅",
        "5": "香蕉黃",
        "6": "橘子橙",
        "7": "孔雀藍",
        "8": "石墨灰",
        "9": "藍莓藍",
        "10": "羅勒綠",
        "11": "番茄紅",
    }
    return mapping.get(str(color_id), f"未知({color_id})")


def find_matching_calendar_event(service, calendar_id, address, target_date, start_time_str, end_time_str):
    target_date_obj = parse_date_value(target_date)
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)

    tz = timezone(timedelta(hours=8))
    day_start = datetime(
        target_date_obj.year,
        target_date_obj.month,
        target_date_obj.day,
        0, 0, 0,
        tzinfo=tz,
    )
    day_end = day_start + timedelta(days=1)

    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=day_start.isoformat(),
        timeMax=day_end.isoformat(),
        singleEvents=True,
        orderBy="startTime",
    ).execute()

    events = events_result.get("items", [])
    target_addr = normalize_addr_for_match(address)

    for event in events:
        start_raw = event.get("start", {}).get("dateTime") or event.get("start", {}).get("date")
        end_raw = event.get("end", {}).get("dateTime") or event.get("end", {}).get("date")

        start_dt = parse_event_time(start_raw)
        end_dt = parse_event_time(end_raw)
        if not start_dt or not end_dt:
            continue

        same_date = start_dt.date() == target_date_obj.date()
        same_start = (start_dt.hour, start_dt.minute) == (sh, sm)
        same_end = (end_dt.hour, end_dt.minute) == (eh, em)

        location = event.get("location", "") or ""
        description = event.get("description", "") or ""
        summary = event.get("summary", "") or ""

        text_blob = normalize_addr_for_match(location + " " + description + " " + summary)
        addr_match = target_addr and target_addr in text_blob

        if same_date and same_start and same_end and addr_match:
            return event

    return None


def sync_calendar_color_for_row(service, calendar_id, address, date_value, start_time_str, end_time_str):
    if not ENABLE_GCAL_COLOR_SYNC or service is None:
        return {
            "日曆改色結果": "未執行",
            "日曆改色原因": "未啟用日曆改色",
            "日曆原色": "",
            "日曆新色": "",
        }

    try:
        event = find_matching_calendar_event(
            service=service,
            calendar_id=calendar_id,
            address=address,
            target_date=date_value,
            start_time_str=start_time_str,
            end_time_str=end_time_str,
        )
    except HttpError as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": f"Calendar API 錯誤: {e}",
            "日曆原色": "",
            "日曆新色": "",
        }
    except Exception as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": f"Calendar 例外: {e}",
            "日曆原色": "",
            "日曆新色": "",
        }

    if not event:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": "找不到對應日曆事件",
            "日曆原色": "",
            "日曆新色": "",
        }

    event_id = event.get("id")
    old_color = str(event.get("colorId", ""))
    old_color_name = color_name_from_id(old_color)

    if old_color != COLOR_PURPLE:
        return {
            "日曆改色結果": "未改",
            "日曆改色原因": f"需求有異動（原色：{old_color_name}）",
            "日曆原色": old_color_name,
            "日曆新色": old_color_name,
        }

    try:
        service.events().patch(
            calendarId=calendar_id,
            eventId=event_id,
            body={"colorId": COLOR_YELLOW},
        ).execute()
    except HttpError as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": f"改色 API 錯誤: {e}",
            "日曆原色": old_color_name,
            "日曆新色": old_color_name,
        }
    except Exception as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": f"改色例外: {e}",
            "日曆原色": old_color_name,
            "日曆新色": old_color_name,
        }

    return {
        "日曆改色結果": "成功",
        "日曆改色原因": "葡萄紫 → 香蕉黃",
        "日曆原色": old_color_name,
        "日曆新色": color_name_from_id(COLOR_YELLOW),
    }


# =========================
# 各階段
# =========================
def prepare_base_order_data(
    row,
    member_payload,
    address_info,
    clean_type_id,
    people,
    hours,
    system_period,
    note_info,
):
    member = member_payload.get("member", {}) if isinstance(member_payload, dict) else {}
    last_purchase = member_payload.get("lastPurchase", {}) if isinstance(member_payload, dict) else {}
    old_purchase = address_info.get("purchase", {}) if isinstance(address_info, dict) else {}

    def pick(key, default=""):
        if old_purchase.get(key) not in (None, ""):
            return old_purchase.get(key)
        if last_purchase.get(key) not in (None, ""):
            return last_purchase.get(key)
        return default

    base_memo = ""
    if note_info["need_note"]:
        base_memo = note_info["customer_time_note"]

    return {
        "clean_type_id": clean_type_id,
        "phone": normalize_phone(row["電話"]),
        "name": str(member.get("name") or row["姓名"]).strip(),
        "email": str(member.get("email") or "").strip(),
        "tel": str(member.get("tel") or ""),
        "line": str(member.get("line") or ""),
        "fbName": str(member.get("fb_name") or ""),
        "fb": str(member.get("fb") or ""),
        "memoProcess": str(member.get("memo_process") or ""),
        "memoFinance": str(member.get("memo_finance") or ""),
        "addressId": str(address_info.get("addressId") or ""),
        "country_id": str(address_info.get("country_id") or pick("country_id", "12")),
        "address": str(address_info.get("address") or row["地址"]).strip(),
        "ping": str(pick("ping", "4")),
        "room": str(pick("room", "0")),
        "bathroom": str(pick("bathroom", "0")),
        "balcony": str(pick("balcony", "0")),
        "livingroom": str(pick("livingroom", "0")),
        "kitchen": str(pick("kitchen", "0")),
        "window": str(pick("window", "")),
        "shutter": str(pick("shutter", "")),
        "clothes": str(pick("clothes", "0")),
        "dyson": str(pick("dyson", "0")),
        "refrigerator": str(pick("refrigerator", "0")),
        "disinfection": str(pick("disinfection", "0")),
        "go_abord": str(pick("go_abord", "0")),
        "home_move": str(pick("home_move", "0")),
        "storage": str(pick("storage", "0")),
        "cabinet": str(pick("cabinet", "0")),
        "quintuple": str(pick("quintuple", "0")),
        "hour": "",
        "price": "",
        "price_vvip": "",
        "person": str(int(people)),
        "date_s": "",
        "period_s": system_period,
        "cycle": "1",
        "fare": "",
        "period": note_info["sms_time"] if note_info["need_note"] else "",
        "memo": base_memo,
        "notice": str(address_info.get("notice") or pick("notice", "")),
        "discount_code": "",
        "payway": "4",
        "is_backend": "477",
        "member_id": str(member.get("member_id") or ""),
        "company_id": str(address_info.get("company_id") or pick("company_id", "1")),
        "area_id": str(address_info.get("area_id") or pick("area_id", "25")),
        "lat": str(address_info.get("lat") or pick("lat", "")),
        "lng": str(address_info.get("lng") or pick("lng", "")),
    }


def stage_send_confirmation(order_no, session):
    result = {"確認信": ""}

    if not order_no:
        return result

    try:
        ok, mail_msg = send_confirmation_mail(session, order_no)
        result["確認信"] = "已發送" if ok else f"發送失敗: {mail_msg}"
    except Exception as e:
        result["確認信"] = f"發送失敗: {e}"

    return result


def stage_calendar_color(row, gcal_service, region):
    calendar_id = GOOGLE_CALENDAR_MAP.get(region)
    if not calendar_id:
        return {
            "日曆改色結果": "未執行",
            "日曆改色原因": f"找不到區域 {region} 的日曆設定",
            "日曆原色": "",
            "日曆新色": "",
        }

    try:
        return sync_calendar_color_for_row(
            service=gcal_service,
            calendar_id=calendar_id,
            address=str(row["地址"]).strip(),
            date_value=row["日期"],
            start_time_str=str(row["開始時間"]).strip(),
            end_time_str=str(row["結束時間"]).strip(),
        )
    except Exception as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": str(e),
            "日曆原色": "",
            "日曆新色": "",
        }


def stage_update_status(order_no, calendar_info):
    if order_no and calendar_info.get("日曆改色結果") == "成功":
        return {"狀態": "已安排"}
    return {}


def has_action(selected_actions, action_name):
    if not selected_actions:
        return True
    return action_name in selected_actions


def filter_dates_by_balance(date_slots, date_prices, stored_value):
    selected_slots = []
    selected_prices = []
    total = 0

    for slot, price in zip(date_slots, date_prices):
        if total + price <= stored_value:
            selected_slots.append(slot)
            selected_prices.append(price)
            total += price

    return selected_slots, selected_prices, total


def process_existing_order_only(row, gcal_service, region, session, selected_actions=None):
    order_no = str(row.get("訂單編號", "")).strip()

    if not order_no:
        return build_row_result(
            result="失敗",
            reason="無訂單編號",
            staff="無人力",
            service_status="未處理",
            fare="0",
        )

    meta = fetch_order_meta_by_order_no(session, order_no)

    result = build_row_result(
        order_no=order_no,
        result="跳過",
        reason="",
        staff=meta.get("服務人員", "無人力"),
        service_status=meta.get("服務狀態", "未處理"),
        fare=meta.get("車馬費", "0"),
    )

    did_anything = False

    if has_action(selected_actions, "寄確認信"):
        result.update(stage_send_confirmation(order_no, session))
        did_anything = True

    if has_action(selected_actions, "改 Google 日曆"):
        calendar_info = stage_calendar_color(row, gcal_service, region)
        result.update(calendar_info)
        result.update(stage_update_status(order_no, calendar_info))
        did_anything = True

    if did_anything:
        result["結果"] = "成功"

    return result


def process_one_group(
    session,
    rows_with_idx,
    token,
    gcal_service,
    region,
    backend_user_id=None,
    selected_actions=None,
):
    row_num0, row0 = rows_with_idx[0]

    purchase_item = str(row0["購買項目"]).strip()
    clean_type_id = CLEAN_TYPE_MAP.get(purchase_item)
    if not clean_type_id:
        raise Exception(f"未知購買項目: {purchase_item}")

    mapped = map_to_system_slot(row0["開始時間"], row0["結束時間"])
    system_period = mapped["system_slot"]

    people, hours = parse_service_human_hour(
        row0["服務人時"],
        row0["開始時間"],
        row0["結束時間"],
    )
    if hours is None:
        raise Exception("無法判斷服務時數")

    phone = normalize_phone(row0["電話"])
    member_payload = get_member(session, phone, token, clean_type_id)
    if not member_payload:
        raise Exception(f"會員不存在: {phone}")

    member = member_payload.get("member", {})
    stored_value = int(float(member_payload.get("storedValue", 0) or 0))

    target_address = str(row0["地址"]).strip().split(",")[0]
    best_addr = pick_best_address_info(member_payload, target_address)
    print("[DEBUG] best_addr =", best_addr)

    if not best_addr:
        raise Exception(f"下拉選單找不到對應地址：{target_address}")
    if not str(best_addr.get("addressId", "")).strip():
        raise Exception(f"地址存在但缺少 addressId：{target_address}")

    selected_address = str(best_addr.get("address", "")).strip()

    geo_lat, geo_lng = geocode_address(selected_address)
    if geo_lat and geo_lng:
        best_addr["lat"] = geo_lat
        best_addr["lng"] = geo_lng

    addr_check = check_contain(
        session=session,
        member_id=member.get("member_id", ""),
        address=selected_address,
        lat=best_addr.get("lat", ""),
        lng=best_addr.get("lng", ""),
        token=token,
        clean_type_id=clean_type_id,
    )
    print("[DEBUG] addr_check =", addr_check)

    if not addr_check:
        raise Exception(f"查詢地區失敗：{selected_address}")

    area_data = addr_check.get("area", {}) if isinstance(addr_check.get("area"), dict) else {}
    purchase_data = addr_check.get("purchase", {}) if isinstance(addr_check.get("purchase"), dict) else {}

    best_addr["country_id"] = area_data.get("country_id", best_addr.get("country_id"))
    best_addr["area_id"] = area_data.get("area_id", best_addr.get("area_id"))
    best_addr["company_id"] = area_data.get("company_id", best_addr.get("company_id"))

    if purchase_data:
        best_addr["purchase"] = purchase_data

    best_addr["fare"] = (
        purchase_data.get("fare")
        or purchase_data.get("car_fare")
        or area_data.get("fare")
        or "0"
    )
    best_addr["notice"] = (
        purchase_data.get("notice")
        or purchase_data.get("service_notice")
        or area_data.get("notice")
        or ""
    )

    base_data = prepare_base_order_data(
        row=row0,
        member_payload=member_payload,
        address_info=best_addr,
        clean_type_id=clean_type_id,
        people=people,
        hours=hours,
        system_period=system_period,
        note_info=mapped,
    )

    calc_result = calculate_hour(session, base_data, token)
    if not calc_result:
        raise Exception("計算時數失敗")

    calc_fields = extract_calc_fields(
        calc_result=calc_result,
        fallback_hours=int(float(hours)),
        fallback_fare=best_addr.get("fare", "0"),
    )

    if calc_fields["hour"]:
        base_data["hour"] = calc_fields["hour"]
    if calc_fields["price"]:
        base_data["price"] = calc_fields["price"]
    if calc_fields["price_vvip"]:
        base_data["price_vvip"] = calc_fields["price_vvip"]
    if calc_fields["fare"] not in ("", None):
        base_data["fare"] = calc_fields["fare"]

    row_details = []
    for row_num, row in rows_with_idx:
        target_slot = build_target_slot_from_row(row)
        price = int(float(base_data.get("price") or 0))

        row_details.append({
            "row_num": row_num,
            "date": normalize_sheet_date(row["日期"]),
            "slot": target_slot,
            "price": price,
            "display_period": normalize_sheet_period(row["開始時間"], row["結束時間"]),
            "row": row,
        })

    need_create_order = has_action(selected_actions, "建單")
    row_results = {}

    if not need_create_order:
        for detail in row_details:
            existing_order_no = str(detail["row"].get("訂單編號", "")).strip()
            meta = fetch_order_meta_by_order_no(session, existing_order_no) if existing_order_no else {
                "服務人員": "無人力",
                "服務狀態": "未處理",
                "車馬費": "0",
            }

            row_results[detail["row_num"]] = build_row_result(
                order_no=existing_order_no,
                result="成功" if existing_order_no else "失敗",
                reason="" if existing_order_no else "無訂單編號",
                sms_time=base_data.get("period", ""),
                customer_note=base_data.get("memo", ""),
                staff=meta.get("服務人員", "無人力"),
                service_status=meta.get("服務狀態", "未處理"),
                fare=meta.get("車馬費", "0"),
            )
        return row_results

    raw_slots = [x["slot"] for x in row_details]
    valid_slots, invalid_slots = validate_available_slots(session, base_data, token, raw_slots)

    valid_details = [d for d in row_details if d["slot"] in valid_slots]
    invalid_details = [d for d in row_details if d["slot"] in invalid_slots]

    for detail in invalid_details:
        row_results[detail["row_num"]] = build_row_result(
            result="失敗",
            reason="查詢班表時沒有相應時段資料，依系統規範不可送出",
            no_slot_date=detail["date"],
            sms_time=base_data.get("period", ""),
            customer_note=base_data.get("memo", ""),
            staff="無人力",
            service_status="未處理",
            fare="0",
        )

    if not valid_details:
        return row_results

    valid_slots_for_balance = [x["slot"] for x in valid_details]
    valid_prices_for_balance = [x["price"] for x in valid_details]
    send_slots, _, _ = filter_dates_by_balance(
        valid_slots_for_balance,
        valid_prices_for_balance,
        stored_value,
    )

    send_details = [d for d in valid_details if d["slot"] in send_slots]
    no_balance_details = [d for d in valid_details if d["slot"] not in send_slots]

    for detail in no_balance_details:
        row_results[detail["row_num"]] = build_row_result(
            result="未送",
            reason="餘額不足",
            insufficient_date=detail["date"],
            sms_time=base_data.get("period", ""),
            customer_note=base_data.get("memo", ""),
            staff="無人力",
            service_status="未處理",
            fare=str(base_data.get("fare", "0")),
        )

    if not send_details:
        return row_results

    for detail in send_details:
        payload = base_data.copy()
        payload["price"] = str(detail["price"])
        payload["price_vvip"] = "0"
        payload["fare"] = str(base_data.get("fare") or best_addr.get("fare") or "0")
        payload["notice"] = str(base_data.get("notice") or best_addr.get("notice") or "")
        payload["area_id"] = str(base_data.get("area_id") or best_addr.get("area_id") or "")
        payload["company_id"] = str(base_data.get("company_id") or best_addr.get("company_id") or "")
        payload["addressId"] = str(base_data.get("addressId") or best_addr.get("addressId") or "")

        target_slot = detail["slot"]

        print("[DEBUG] final booking payload =", {
            "row_num": detail["row_num"],
            "target_slot": target_slot,
            "addressId": payload.get("addressId"),
            "fare": payload.get("fare"),
            "notice": payload.get("notice"),
            "area_id": payload.get("area_id"),
            "company_id": payload.get("company_id"),
        })

        session.post(
            BOOKING_URL,
            data={**payload, "_token": token, "date_list[]": [target_slot]},
            headers=HEADERS,
            allow_redirects=True,
        )

        time.sleep(1)

        order_no = fetch_order_no_by_date_and_period(
            session=session,
            target_date=detail["date"],
            target_period=detail["display_period"],
        )

        if not order_no:
            stage_result = build_row_result(
                result="失敗",
                reason=f"送單後抓不到訂單編號（目標時段：{target_slot}）",
                sms_time=base_data.get("period", ""),
                customer_note=base_data.get("memo", ""),
                staff="無人力",
                service_status="未處理",
                fare=str(base_data.get("fare", "0")),
            )
            print("[DEBUG] stage_result =", stage_result)
            row_results[detail["row_num"]] = stage_result
            continue

        meta = fetch_order_meta_by_order_no(session, order_no)

        stage_result = build_row_result(
            order_no=order_no,
            result="成功",
            reason="",
            sms_time=base_data.get("period", ""),
            customer_note=base_data.get("memo", ""),
            staff=meta.get("服務人員", "無人力"),
            service_status=meta.get("服務狀態", "未處理"),
            fare=meta.get("車馬費", "0") or str(base_data.get("fare", "0")),
        )

        if has_action(selected_actions, "寄確認信"):
            stage_result.update(stage_send_confirmation(order_no, session))

        if has_action(selected_actions, "改 Google 日曆"):
            calendar_info = stage_calendar_color(detail["row"], gcal_service, region)
            stage_result.update(calendar_info)
            stage_result.update(stage_update_status(order_no, calendar_info))

        print("[DEBUG] stage_result =", stage_result)
        row_results[detail["row_num"]] = stage_result

    return row_results


# =========================
# 主執行
# =========================
def run_process(sheet_name, start_row, end_row, env_name_from_ui=None):
    print(f"目前環境：{ENV}")
    print(f"BASE_URL：{BASE_URL}")
    print(f"執行工作表：{sheet_name}")
    print(f"執行列範圍：{start_row} ~ {end_row}")

    ws, df = load_worksheet(sheet_name)

    required_cols = [
        "服務人時",
        "備註",
        "姓名",
        "電話",
        "地址",
        "日期",
        "開始時間",
        "結束時間",
        "狀態",
        "購買項目",
        "訂單編號",
    ]
    for col in required_cols:
        if col not in df.columns:
            raise Exception(f"工作表缺少必要欄位: {col}")

    df = df[(df["__sheet_row__"] >= start_row) & (df["__sheet_row__"] <= end_row)]
    df = df[df.apply(should_process_row, axis=1)]

    if df.empty:
        print("沒有符合條件的資料可執行。")
        return

    gcal_service = None
    if ENABLE_GCAL_COLOR_SYNC:
        try:
            gcal_service = build_gcal_service()
            print("Google Calendar 已啟用")
        except Exception as e:
            print(f"Google Calendar 初始化失敗：{e}")
            gcal_service = None

    grouped_orders = defaultdict(list)

    for _, row in df.iterrows():
        region = get_region_by_address(str(row["地址"]), ACCOUNTS)
        if not region:
            continue
        if not should_create_order(row):
            continue

        key = (region, build_group_key(row))
        grouped_orders[key].append((int(row["__sheet_row__"]), row))

    all_row_results = {}

    region_groups = defaultdict(list)
    for (region, group_key), items in grouped_orders.items():
        region_groups[region].append((group_key, items))

    for region, group_items in region_groups.items():
        config = ACCOUNTS.get(region)
        if not config:
            continue

        email = config["email"]
        password = config["password"]

        print(f"\n===== 開始處理區域：{region} ({email}) =====")

        session = requests.Session()
        if not login(session, email, password):
            print("登入失敗，略過該區域")
            continue

        for group_no, (_, rows_with_idx) in enumerate(group_items, start=1):
            _, first_row = rows_with_idx[0]
            print(f"\n--- 處理第 {group_no} 組：{first_row['姓名']}，共 {len(rows_with_idx)} 筆 ---")

            try:
                token = get_csrf_token(session)
                row_results = process_one_group(
                    session=session,
                    rows_with_idx=rows_with_idx,
                    token=token,
                    gcal_service=gcal_service,
                    region=region,
                    backend_user_id=None,
                    selected_actions=["建單", "寄確認信", "改 Google 日曆"],
                )
                all_row_results.update(row_results)
            except Exception as e:
                print(f"❌ 整組失敗：{e}")
                for row_num, _ in rows_with_idx:
                    all_row_results[row_num] = build_row_result(
                        result="失敗",
                        reason=str(e),
                        staff="無人力",
                        service_status="未處理",
                        fare="0",
                    )

            time.sleep(REQUEST_DELAY)

    print("[DEBUG] row_results keys =", list(all_row_results.keys()))
    for row_num, info in all_row_results.items():
        print("[DEBUG] final row", row_num, info.get("服務人員"), info.get("服務狀態"), info.get("車馬費"))

    update_sheet_rows(ws, all_row_results)
    print("已回填 Google Sheet。")


def get_runtime_config(env_name: str):
    if env_name == "dev":
        return {
            "BASE_URL": BASE_URL_DEV,
            "ORDER_PREFIX": ORDER_PREFIX_DEV,
        }
    return {
        "BASE_URL": BASE_URL_PROD,
        "ORDER_PREFIX": ORDER_PREFIX_PROD,
    }


def run_process_web(
    env_name,
    region,
    backend_email,
    backend_password,
    sheet_name,
    start_row,
    end_row,
    selected_actions=None,
    logger=print,
):
    logger("=== run_process_web 已進入：2026-04-23-final-v5 ===")

    global BASE_URL, ORDER_PREFIX
    if env_name == "dev":
        BASE_URL = BASE_URL_DEV
        ORDER_PREFIX = ORDER_PREFIX_DEV
    else:
        BASE_URL = BASE_URL_PROD
        ORDER_PREFIX = ORDER_PREFIX_PROD

    global LOGIN_URL, BOOKING_URL, PURCHASE_URL, GET_MEMBER_URL
    global CHECK_CONTAIN_URL, CALCULATE_HOUR_URL, GET_SECTION_URL, MAIL_SUCCESS_URL

    LOGIN_URL = f"{BASE_URL}/login"
    BOOKING_URL = f"{BASE_URL}/booking/stored_value_routine"
    PURCHASE_URL = f"{BASE_URL}/purchase"
    GET_MEMBER_URL = f"{BASE_URL}/ajax/get_member"
    CHECK_CONTAIN_URL = f"{BASE_URL}/ajax/check_contain"
    CALCULATE_HOUR_URL = f"{BASE_URL}/ajax/calculate_hour"
    GET_SECTION_URL = f"{BASE_URL}/ajax/get_section"
    MAIL_SUCCESS_URL = f"{BASE_URL}/purchase/mail_success/{{order_no}}"

    logger(f"目前環境：{env_name}")
    logger(f"BASE_URL：{BASE_URL}")
    logger(f"執行區域：{region}")
    logger(f"執行工作表：{sheet_name}")
    logger(f"執行列範圍：{start_row} ~ {end_row}")

    if selected_actions is None:
        selected_actions = ["建單", "寄確認信", "改 Google 日曆"]

    ws, df = load_worksheet(sheet_name)

    required_cols = [
        "服務人時",
        "備註",
        "姓名",
        "電話",
        "地址",
        "日期",
        "開始時間",
        "結束時間",
        "狀態",
        "購買項目",
        "訂單編號",
    ]
    for col in required_cols:
        if col not in df.columns:
            raise Exception(f"工作表缺少必要欄位: {col}")

    df = df[(df["__sheet_row__"] >= start_row) & (df["__sheet_row__"] <= end_row)]
    df = df[df.apply(should_process_row, axis=1)]

    if df.empty:
        logger("沒有符合條件的資料可執行。")
        return {"success": True, "message": "沒有符合條件的資料"}

    filtered_rows = []
    for _, row in df.iterrows():
        row_region = get_region_by_address(str(row["地址"]), ACCOUNTS)
        if row_region == region:
            filtered_rows.append(row)

    if not filtered_rows:
        logger(f"沒有 {region} 區域的資料可執行。")
        return {"success": True, "message": f"沒有 {region} 區域資料"}

    df = pd.DataFrame(filtered_rows)
    if "__sheet_row__" not in df.columns:
        raise Exception("資料缺少 __sheet_row__")

    gcal_service = None
    if ENABLE_GCAL_COLOR_SYNC:
        try:
            gcal_service = build_gcal_service()
            logger("Google Calendar 已啟用")
        except Exception as e:
            logger(f"Google Calendar 初始化失敗：{e}")
            gcal_service = None

    session = requests.Session()
    login_ok = login(session, backend_email, backend_password)
    if not login_ok:
        raise Exception("後台登入失敗，請確認帳號密碼")

    grouped_orders = defaultdict(list)
    existing_order_rows = []

    for _, row in df.iterrows():
        row_num = int(row["__sheet_row__"])

        if not has_action(selected_actions, "建單"):
            existing_order_rows.append((row_num, row))
            continue

        if not should_create_order(row):
            existing_order_rows.append((row_num, row))
            continue

        key = build_group_key(row)
        grouped_orders[key].append((row_num, row))

    all_row_results = {}

    for row_num, row in existing_order_rows:
        try:
            result = process_existing_order_only(
                row=row,
                gcal_service=gcal_service,
                region=region,
                session=session,
                selected_actions=selected_actions,
            )
            all_row_results[row_num] = result
        except Exception as e:
            all_row_results[row_num] = build_row_result(
                result="失敗",
                reason=f"補處理失敗: {e}",
                staff="無人力",
                service_status="未處理",
                fare="0",
            )

    for group_no, (_, rows_with_idx) in enumerate(grouped_orders.items(), start=1):
        _, first_row = rows_with_idx[0]
        logger(f"處理第 {group_no} 組：{first_row['姓名']}，共 {len(rows_with_idx)} 筆")

        try:
            token = get_csrf_token(session)
            row_results = process_one_group(
                session=session,
                rows_with_idx=rows_with_idx,
                token=token,
                gcal_service=gcal_service,
                region=region,
                backend_user_id=None,
                selected_actions=selected_actions,
            )
            all_row_results.update(row_results)
        except Exception as e:
            logger(f"整組失敗：{e}")
            for row_num, _ in rows_with_idx:
                all_row_results[row_num] = build_row_result(
                    result="失敗",
                    reason=str(e),
                    staff="無人力",
                    service_status="未處理",
                    fare="0",
                )

        time.sleep(REQUEST_DELAY)

    logger(f"[DEBUG] row_results keys = {list(all_row_results.keys())}")
    for row_num, info in all_row_results.items():
        logger(f"[DEBUG] write row {row_num} {info.get('服務人員')} {info.get('服務狀態')} {info.get('車馬費')}")

    update_sheet_rows(ws, all_row_results)
    logger("已回填 Google Sheet。")

    success_count = sum(1 for v in all_row_results.values() if v.get("結果") == "成功")
    fail_count = sum(1 for v in all_row_results.values() if v.get("結果") == "失敗")

    return {
        "success": True,
        "sheet_name": sheet_name,
        "region": region,
        "env": env_name,
        "success_count": success_count,
        "fail_count": fail_count,
        "total_processed": len(all_row_results),
    }
