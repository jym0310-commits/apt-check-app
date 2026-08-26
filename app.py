import base64
import glob
import html
import io
import json
import os
import uuid
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt
from PIL import Image, ImageOps
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound

# -------------------------
# 기본 설정
# -------------------------
st.set_page_config(layout="wide", page_title="해링턴 하자 관리", page_icon="✓", initial_sidebar_state="collapsed")

WEB_STATUS_OPTIONS = ["미완료", "확인완료"]
LEGACY_INCOMPLETE_STATUSES = {"미확인", "재확인필요", "미완료", ""}
WEB_STATUS_SHEET_NAME = "web_status"
WEB_ITEMS_SHEET_NAME = "web_items"
WEB_IMAGES_SHEET_NAME = "web_images"
MAX_UPLOAD_IMAGES = 5
IMAGE_CHUNK_SIZE = 40000
LEGACY_BUILDING = "204동"
LEGACY_UNIT = "4503호"
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

st.markdown(
    """
    <style>
    :root {
        --ink: #0a0a0a;
        --body: #737373;
        --mute: #a3a3a3;
        --canvas: #ffffff;
        --soft: #fafafa;
        --line: #e5e5e5;
        --line-strong: #d4d4d4;
        --dark: #171717;
    }

    html, body, [class*="css"] {
        font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        color: var(--ink);
    }
    .stApp { background: var(--canvas); }
    .block-container { max-width: 1180px; padding-top: 2.1rem; padding-bottom: 5rem; }
    header[data-testid="stHeader"] { background: rgba(255,255,255,.96); border-bottom: 1px solid var(--line); }
    [data-testid="stSidebar"] { background: var(--soft); border-right: 1px solid var(--line); }

    h1, h2, h3 { letter-spacing: -0.025em; color: var(--ink); }
    h1 { font-size: 2.15rem !important; font-weight: 600 !important; }
    h2 { font-size: 1.55rem !important; font-weight: 600 !important; }
    h3 { font-size: 1.05rem !important; font-weight: 600 !important; }
    p, .stCaption { color: var(--body); }

    .app-nav {
        display:flex; align-items:center; justify-content:space-between; gap:16px;
        padding: 4px 0 28px 0; margin-bottom: 22px; border-bottom:1px solid var(--line);
    }
    .brand-wrap { display:flex; align-items:center; gap:12px; }
    .brand-mark {
        width:34px; height:34px; border-radius:999px; background:var(--ink); color:white;
        display:flex; align-items:center; justify-content:center; font-size:17px; font-weight:700;
    }
    .brand-title { font-size:15px; font-weight:650; color:var(--ink); line-height:1.2; }
    .brand-sub { font-size:12px; color:var(--mute); margin-top:2px; }
    .unit-pill {
        display:inline-flex; align-items:center; gap:7px; padding:8px 14px; border:1px solid var(--line);
        border-radius:999px; background:white; font-size:13px; font-weight:600; color:var(--ink);
    }

    .hero { padding: 28px 0 34px 0; }
    .eyebrow { font-size:13px; color:var(--body); margin-bottom:10px; font-weight:600; }
    .hero-title { font-size:36px; line-height:1.16; font-weight:600; letter-spacing:-0.035em; margin:0; color:var(--ink); }
    .hero-copy { max-width:680px; margin-top:12px; font-size:15px; line-height:1.65; color:var(--body); }

    .login-shell { max-width:680px; margin:8vh auto 0 auto; text-align:center; }
    .login-logo {
        width:54px; height:54px; border-radius:999px; background:var(--ink); color:white;
        display:flex; align-items:center; justify-content:center; margin:0 auto 24px auto; font-size:25px; font-weight:700;
    }
    .login-title { font-size:36px; font-weight:600; line-height:1.12; letter-spacing:-0.04em; margin-bottom:12px; }
    .login-copy { color:var(--body); font-size:15px; margin:0 auto 28px auto; max-width:520px; line-height:1.65; }

    .section-label { font-size:20px; font-weight:600; letter-spacing:-.025em; margin:42px 0 14px 0; }
    .section-caption { color:var(--body); font-size:13px; margin-top:-7px; margin-bottom:16px; }

    .kpi-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin:18px 0 12px 0; }
    .kpi-card { border:1px solid var(--line); border-radius:12px; padding:18px 18px 16px; background:white; }
    .kpi-label { font-size:12px; color:var(--body); margin-bottom:8px; }
    .kpi-value { font-size:27px; line-height:1; font-weight:650; letter-spacing:-.035em; color:var(--ink); }
    .kpi-meta { font-size:11px; color:var(--mute); margin-top:8px; }
    .progress-shell { width:100%; height:8px; background:#f0f0f0; border-radius:999px; overflow:hidden; margin-top:14px; }
    .progress-fill { height:100%; background:var(--ink); border-radius:999px; }

    .issue-card {
        background:white; padding:20px; border-radius:12px; border:1px solid var(--line); margin-bottom:10px; color:var(--ink);
    }
    .issue-head { display:flex; align-items:flex-start; justify-content:space-between; gap:10px; }
    .issue-id { font-size:11px; color:var(--mute); font-family:ui-monospace, SFMono-Regular, Menlo, monospace; margin-bottom:6px; }
    .issue-title { font-size:17px; font-weight:650; letter-spacing:-.02em; margin:0; }
    .issue-detail { font-size:13px; line-height:1.6; color:var(--body); margin:12px 0 0; }
    .status-badge { display:inline-flex; align-items:center; padding:6px 10px; border-radius:999px; font-size:11px; font-weight:650; white-space:nowrap; }
    .status-incomplete { background:#f5f5f5; color:#525252; border:1px solid #e5e5e5; }
    .status-checked { background:#171717; color:#fff; border:1px solid #171717; }
    .status-meta { font-size:11px; color:var(--mute); margin-top:10px; }

    div[data-testid="stForm"] { border:1px solid var(--line) !important; border-radius:12px !important; padding:22px !important; background:white; }
    div[data-testid="stExpander"] { border:1px solid var(--line) !important; border-radius:12px !important; background:white !important; box-shadow:none !important; }
    div[data-testid="stExpander"] details summary { font-weight:600; }
    div[data-baseweb="select"] > div, .stTextInput input, .stTextArea textarea {
        border-color:var(--line-strong) !important; border-radius:12px !important; box-shadow:none !important;
    }
    .stTextInput input { min-height:42px; }
    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
        border-radius:999px !important; border:1px solid var(--line-strong) !important; background:white !important; color:var(--ink) !important;
        font-weight:600 !important; min-height:40px; box-shadow:none !important;
    }
    .stButton > button[kind="primary"], .stFormSubmitButton > button { background:var(--ink) !important; color:white !important; border-color:var(--ink) !important; }
    .stDownloadButton > button { background:var(--ink) !important; color:white !important; border-color:var(--ink) !important; }
    .stButton > button:disabled { background:var(--soft) !important; color:var(--mute) !important; border-color:var(--line) !important; }
    [data-testid="stFileUploaderDropzone"] { border:1px dashed var(--line-strong) !important; border-radius:12px !important; background:var(--soft) !important; }
    [data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:12px; overflow:hidden; }
    [data-testid="stAlert"] { border-radius:12px !important; box-shadow:none !important; }
    hr { border:none !important; border-top:1px solid var(--line) !important; margin:40px 0 !important; }

    @media (max-width: 760px) {
        .block-container { padding: 1rem 1rem 4rem 1rem; }
        .app-nav { padding-bottom:18px; margin-bottom:10px; }
        .hero { padding:20px 0 24px 0; }
        .hero-title, .login-title { font-size:28px; }
        .kpi-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
        .unit-pill { font-size:12px; padding:7px 11px; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -------------------------
# 세대 로그인 (동/호수)
# -------------------------
def normalize_building(value):
    value = str(value or "").strip()
    if not value:
        return ""
    return value if value.endswith("동") else f"{value}동"


def normalize_unit(value):
    value = str(value or "").strip()
    if not value:
        return ""
    return value if value.endswith("호") else f"{value}호"


if "logged_in_unit" not in st.session_state:
    st.session_state.logged_in_unit = False

if not st.session_state.logged_in_unit:
    st.markdown(
        """
        <div class="login-shell">
          <div class="login-logo">✓</div>
          <div class="login-title">우리 집 하자를<br>한 곳에서 관리하세요.</div>
          <div class="login-copy">동과 호수만 입력하면 해당 세대의 하자 목록, 진행상태, 사진과 A/S 신청서를 한 화면에서 관리할 수 있습니다.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.form("unit_login_form"):
        lc1, lc2 = st.columns(2)
        with lc1:
            login_building_raw = st.text_input("동", placeholder="예: 204 또는 204동")
        with lc2:
            login_unit_raw = st.text_input("호수", placeholder="예: 4503 또는 4503호")
        login_submit = st.form_submit_button("세대 관리 시작", use_container_width=True, type="primary")

    if login_submit:
        login_building = normalize_building(login_building_raw)
        login_unit = normalize_unit(login_unit_raw)
        if not login_building or not login_unit:
            st.error("동과 호수를 모두 입력해 주세요.")
        else:
            st.session_state.logged_in_unit = True
            st.session_state.login_building = login_building
            st.session_state.login_unit = login_unit
            st.rerun()
    st.stop()

CURRENT_BUILDING = st.session_state.get("login_building", LEGACY_BUILDING)
CURRENT_UNIT = st.session_state.get("login_unit", LEGACY_UNIT)

with st.sidebar:
    st.markdown("### 현재 세대")
    st.markdown(f"**{CURRENT_BUILDING} {CURRENT_UNIT}**")
    st.caption("이 세대의 데이터만 표시됩니다.")
    if st.button("로그아웃", use_container_width=True):
        for key in ["logged_in_unit", "login_building", "login_unit"]:
            st.session_state.pop(key, None)
        st.rerun()

st.markdown(
    f"""
    <div class="app-nav">
      <div class="brand-wrap">
        <div class="brand-mark">✓</div>
        <div><div class="brand-title">해링턴 하자관리</div><div class="brand-sub">Home inspection workspace</div></div>
      </div>
      <div class="unit-pill">{html.escape(CURRENT_BUILDING)} · {html.escape(CURRENT_UNIT)}</div>
    </div>
    <div class="hero">
      <div class="eyebrow">세대 하자 관리</div>
      <div class="hero-title">{html.escape(CURRENT_BUILDING)} {html.escape(CURRENT_UNIT)}</div>
      <div class="hero-copy">하자 등록 · 진행현황 · A/S 신청</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# -------------------------
# 기존 Excel 데이터 로드
# -------------------------
excel_candidates = glob.glob("*.xlsx")
if not excel_candidates:
    st.error("저장소 안에서 엑셀(.xlsx) 파일을 찾지 못했습니다. GitHub에 파일이 실제로 존재하는지 확인해 주세요.")
    st.stop()
EXCEL_PATH = excel_candidates[0]


@st.cache_data
def load_data(mtime, path):
    return pd.read_excel(path, sheet_name="Sheet1")


mtime = os.path.getmtime(EXCEL_PATH)
df = load_data(mtime, EXCEL_PATH)
df.columns = [str(c).strip() for c in df.columns]

# 진행현황 공백 정리 (원본 Excel 파일 자체는 수정하지 않음)
if "진행현황" in df.columns:
    df["진행현황_표시"] = df["진행현황"].fillna("미지정").astype(str).str.strip()
else:
    df["진행현황_표시"] = "미지정"

# 새 하자 등록 목록박스는 원본 Excel의 실제 공간/부위/유형 조합을 기준으로 사용한다.
REFERENCE_DF = df.copy()
for _col in ["공간", "부위", "유형"]:
    if _col not in REFERENCE_DF.columns:
        REFERENCE_DF[_col] = ""
    REFERENCE_DF[_col] = REFERENCE_DF[_col].fillna("").astype(str).str.strip()

REFERENCE_DF = REFERENCE_DF[
    (REFERENCE_DF["공간"] != "") | (REFERENCE_DF["부위"] != "") | (REFERENCE_DF["유형"] != "")
].copy()
REFERENCE_SPACES = sorted([x for x in REFERENCE_DF["공간"].dropna().unique().tolist() if str(x).strip()])
REFERENCE_PARTS = sorted([x for x in REFERENCE_DF["부위"].dropna().unique().tolist() if str(x).strip()])
REFERENCE_TYPES = sorted([x for x in REFERENCE_DF["유형"].dropna().unique().tolist() if str(x).strip()])

# -------------------------
# 웹 전용 상태: Google Sheet
# -------------------------
def get_secret(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


@st.cache_resource
def get_google_spreadsheet():
    """Streamlit Secrets/credentials.json으로 Google Spreadsheet 연결."""
    sheet_id = get_secret("WEB_STATUS_SHEET_ID") or os.getenv("WEB_STATUS_SHEET_ID")
    if not sheet_id:
        raise RuntimeError("WEB_STATUS_SHEET_ID가 설정되지 않았습니다.")

    service_account_info = get_secret("gcp_service_account")
    if service_account_info:
        service_account_info = dict(service_account_info)
        if "json_key" in service_account_info:
            raw_json = service_account_info["json_key"]
            if isinstance(raw_json, str):
                service_account_info = json.loads(raw_json)
            elif isinstance(raw_json, dict):
                service_account_info = dict(raw_json)
        credentials = Credentials.from_service_account_info(service_account_info, scopes=GOOGLE_SCOPES)
    elif os.path.exists("credentials.json"):
        credentials = Credentials.from_service_account_file("credentials.json", scopes=GOOGLE_SCOPES)
    else:
        raise RuntimeError("Google 서비스 계정 인증정보가 없습니다.")

    return gspread.authorize(credentials).open_by_key(sheet_id)


def get_or_create_worksheet(name, headers, rows=500, cols=10):
    spreadsheet = get_google_spreadsheet()
    try:
        ws = spreadsheet.worksheet(name)
    except WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=name, rows=rows, cols=max(cols, len(headers)))
        ws.append_row(headers)
    values = ws.get_all_values()
    if not values:
        ws.append_row(headers)
    else:
        # 기존 시트 스키마를 보존하면서 새 컬럼은 오른쪽에만 추가한다.
        current_headers = [str(x).strip() for x in values[0]]
        for header in headers:
            if header not in current_headers:
                current_headers.append(header)
        if current_headers != [str(x).strip() for x in values[0]]:
            ws.update([current_headers], f"A1:{gspread.utils.rowcol_to_a1(1, len(current_headers))}")
    return ws


@st.cache_resource
def get_web_status_worksheet():
    return get_or_create_worksheet(
        WEB_STATUS_SHEET_NAME,
        ["item_id", "web_status", "updated_at", "item_label"],
        rows=500,
        cols=4,
    )


@st.cache_data(ttl=30)
def load_web_status():
    worksheet = get_web_status_worksheet()
    records = worksheet.get_all_records()
    result = {}
    for record in records:
        item_id = str(record.get("item_id", "")).strip()
        raw_status = str(record.get("web_status", "")).strip()
        # 이전 버전의 `미확인`/`재확인필요` 기록은 삭제하지 않고 모두 `미완료`로 읽는다.
        if raw_status == "확인완료":
            status = "확인완료"
        elif raw_status in LEGACY_INCOMPLETE_STATUSES:
            status = "미완료"
        else:
            status = "미완료"
        if item_id:
            result[item_id] = {
                "status": status,
                "updated_at": str(record.get("updated_at", "")).strip(),
            }
    return result


def save_web_status(item_id, new_status, item_label):
    """item_id 기준으로 진행상태를 insert/update한다. Excel은 건드리지 않는다."""
    if new_status not in WEB_STATUS_OPTIONS:
        raise ValueError(f"지원하지 않는 진행상태입니다: {new_status}")
    worksheet = get_web_status_worksheet()
    item_id = str(item_id).strip()
    now = datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y-%m-%d %H:%M:%S")

    values = worksheet.get_all_values()
    target_row = None
    for row_num, row in enumerate(values[1:], start=2):
        if row and str(row[0]).strip() == item_id:
            target_row = row_num
            break

    row_data = [[item_id, new_status, now, item_label]]
    if target_row:
        worksheet.update(row_data, f"A{target_row}:D{target_row}")
    else:
        worksheet.append_row(row_data[0])

    load_web_status.clear()


@st.cache_data(ttl=30)
def load_web_items():
    ws = get_or_create_worksheet(
        WEB_ITEMS_SHEET_NAME,
        ["item_id", "공간", "부위", "유형", "상세내용", "created_at", "active", "동", "호"],
        rows=500,
        cols=9,
    )
    records = ws.get_all_records()
    rows = []
    for r in records:
        if str(r.get("active", "TRUE")).strip().upper() in {"FALSE", "0", "N", "NO"}:
            continue
        item_id = str(r.get("item_id", "")).strip()
        if not item_id:
            continue
        building = normalize_building(r.get("동", "")) or LEGACY_BUILDING
        unit = normalize_unit(r.get("호", "")) or LEGACY_UNIT
        rows.append({
            "번호": item_id,
            "공간": str(r.get("공간", "")).strip(),
            "부위": str(r.get("부위", "")).strip(),
            "유형": str(r.get("유형", "")).strip(),
            "상세내용": str(r.get("상세내용", "")).strip(),
            "진행현황_표시": "웹등록",
            "저장된사진파일명": "",
            "데이터출처": "웹등록",
            "등록일시": str(r.get("created_at", "")).strip(),
            "동": building,
            "호": unit,
        })
    return pd.DataFrame(rows)


@st.cache_data(ttl=30)
def load_web_images():
    ws = get_or_create_worksheet(
        WEB_IMAGES_SHEET_NAME,
        ["image_id", "item_id", "sort_order", "chunk_index", "mime_type", "data_chunk"],
        rows=5000,
        cols=6,
    )
    values = ws.get_all_values()
    grouped = {}
    for row in values[1:]:
        if len(row) < 6:
            row = row + [""] * (6 - len(row))
        image_id, item_id, sort_order, chunk_index, mime_type, data_chunk = row[:6]
        if not image_id or not item_id:
            continue
        key = (item_id, image_id)
        g = grouped.setdefault(key, {"sort_order": int(sort_order or 0), "mime_type": mime_type or "image/jpeg", "chunks": []})
        try:
            idx = int(chunk_index or 0)
        except Exception:
            idx = 0
        g["chunks"].append((idx, data_chunk))

    result = {}
    for (item_id, image_id), info in grouped.items():
        encoded = "".join(chunk for _, chunk in sorted(info["chunks"], key=lambda x: x[0]))
        try:
            data = base64.b64decode(encoded)
        except Exception:
            continue
        result.setdefault(item_id, []).append({
            "image_id": image_id,
            "sort_order": info["sort_order"],
            "mime_type": info["mime_type"],
            "data": data,
        })
    for item_id in result:
        result[item_id].sort(key=lambda x: x["sort_order"])
    return result


def compress_image(uploaded_file):
    uploaded_file.seek(0)
    image = Image.open(uploaded_file)
    image = ImageOps.exif_transpose(image).convert("RGB")
    image.thumbnail((1400, 1400))
    out = io.BytesIO()
    image.save(out, format="JPEG", quality=76, optimize=True)
    return out.getvalue(), "image/jpeg"


def save_web_images(item_id, uploaded_files):
    if not uploaded_files:
        return
    ws = get_or_create_worksheet(
        WEB_IMAGES_SHEET_NAME,
        ["image_id", "item_id", "sort_order", "chunk_index", "mime_type", "data_chunk"],
        rows=5000,
        cols=6,
    )
    append_rows = []
    for sort_order, uploaded in enumerate(uploaded_files[:MAX_UPLOAD_IMAGES], start=1):
        data, mime = compress_image(uploaded)
        encoded = base64.b64encode(data).decode("ascii")
        image_id = f"IMG-{uuid.uuid4().hex[:12]}"
        chunks = [encoded[i:i+IMAGE_CHUNK_SIZE] for i in range(0, len(encoded), IMAGE_CHUNK_SIZE)]
        for chunk_index, chunk in enumerate(chunks):
            append_rows.append([image_id, item_id, sort_order, chunk_index, mime, chunk])
    if append_rows:
        ws.append_rows(append_rows, value_input_option="RAW")
    load_web_images.clear()


def create_web_item(space, part, issue_type, detail, initial_status, uploaded_files, building, unit):
    ws = get_or_create_worksheet(
        WEB_ITEMS_SHEET_NAME,
        ["item_id", "공간", "부위", "유형", "상세내용", "created_at", "active", "동", "호"],
        rows=500,
        cols=9,
    )
    now = datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y-%m-%d %H:%M:%S")
    item_id = f"WEB-{datetime.now(ZoneInfo('Asia/Seoul')).strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"
    ws.append_row([item_id, space, part, issue_type, detail, now, "TRUE", building, unit])
    item_label = f"[{item_id}] {space} - {part}"
    save_web_status(item_id, initial_status, item_label)
    save_web_images(item_id, uploaded_files)
    load_web_items.clear()
    load_web_status.clear()
    return item_id


def delete_web_item(item_id):
    """웹에서 새로 등록한 항목만 soft delete(active=FALSE)한다. Excel 원본은 절대 수정하지 않는다."""
    item_id = str(item_id).strip()
    if not item_id.startswith("WEB-"):
        raise ValueError("Excel 원본 항목은 웹에서 삭제할 수 없습니다.")

    ws = get_or_create_worksheet(
        WEB_ITEMS_SHEET_NAME,
        ["item_id", "공간", "부위", "유형", "상세내용", "created_at", "active", "동", "호"],
        rows=500,
        cols=9,
    )
    values = ws.get_all_values()
    target_row = None
    for row_num, row in enumerate(values[1:], start=2):
        if row and str(row[0]).strip() == item_id:
            target_row = row_num
            break

    if not target_row:
        raise ValueError("삭제할 웹 등록 항목을 찾지 못했습니다.")

    # 삭제 이력 복구 가능성을 위해 행/사진은 보존하고 active만 FALSE로 변경한다.
    ws.update([["FALSE"]], f"G{target_row}")
    load_web_items.clear()
    load_web_status.clear()
    load_web_images.clear()


web_status_enabled = True
web_status_error = None
try:
    web_status_map = load_web_status()
except Exception as exc:
    web_status_enabled = False
    web_status_error = str(exc)
    web_status_map = {}

if not web_status_enabled:
    st.warning(
        "진행상태 저장 기능이 아직 연결되지 않았습니다. Google Sheet 연결 설정을 확인해 주세요.\n\n"
        f"설정 내용: {web_status_error}"
    )

# 웹에서 새로 등록한 항목도 함께 불러온다.
try:
    web_items_df = load_web_items() if web_status_enabled else pd.DataFrame()
    web_images_map = load_web_images() if web_status_enabled else {}
except Exception as exc:
    web_items_df = pd.DataFrame()
    web_images_map = {}
    st.warning(f"웹 신규등록 데이터 조회 실패: {type(exc).__name__}: {exc}")

# Excel 원본 데이터는 기존 세대(204동 4503호)의 초기 데이터로 귀속한다.
df["데이터출처"] = "Excel"
df["등록일시"] = ""
df["동"] = LEGACY_BUILDING
df["호"] = LEGACY_UNIT

# 웹 등록 데이터를 병합한다. 기존에 동/호 컬럼 없이 저장된 웹 데이터도 204동 4503호로 자동 귀속된다.
if not web_items_df.empty:
    for col in df.columns:
        if col not in web_items_df.columns:
            web_items_df[col] = ""
    for col in web_items_df.columns:
        if col not in df.columns:
            df[col] = ""
    df = pd.concat([df, web_items_df[df.columns]], ignore_index=True)

# 로그인한 세대의 데이터만 남긴다.
df = df[(df["동"] == CURRENT_BUILDING) & (df["호"] == CURRENT_UNIT)].copy()

# 각 행에 진행상태를 붙인다. Excel 번호 또는 WEB-... ID를 키로 사용한다.
df["진행상태"] = df["번호"].apply(
    lambda x: web_status_map.get(str(x).strip(), {}).get("status", "미완료")
)

# -------------------------
# AS 신청서 생성 (진행상태 기준 / 기존 양식 재현)
# -------------------------
def infer_trade(part_text, type_text=""):
    """기존 AS 신청서의 공종 표기 방식에 맞춰 부위 중심으로 공종을 추정한다."""
    part = str(part_text or "").strip()
    type_value = str(type_text or "").strip()
    combined = f"{part} {type_value}"

    trade_rules = [
        (["벽도배", "천정도배", "도배"], "도배"),
        (["벽도장", "문틀도장", "도장"], "도장"),
        (["마루"], "마루"),
        (["코킹"], "코킹"),
        (["바닥타일", "벽타일", "타일"], "타일"),
        (["스위치", "콘센트", "조명", "전등", "전기"], "전기"),
        (["설비", "배수", "하수구", "수전", "양변기", "세면기"], "설비"),
        (["목문", "목문틀", "문틀", "창문", "창", "신발장", "수납장", "하부장", "상부장", "서랍장", "냉장고장", "가구", "몰딩"], "목공"),
    ]
    for keywords, trade in trade_rules:
        if any(keyword in combined for keyword in keywords):
            return trade
    return "기타"


def _set_cell_text(cell, text, bold=False, font_size=9, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text or ""))
    run.bold = bold
    run.font.size = Pt(font_size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return p


def _set_cell_width(cell, width_cm):
    cell.width = Cm(width_cm)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(Cm(width_cm).emu / 635)))
    tc_w.set(qn("w:type"), "dxa")




def _set_table_fixed_widths(table, widths_cm):
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    grid_cols = table._tbl.tblGrid.gridCol_lst
    for idx, width_cm in enumerate(widths_cm):
        width_twips = int(Cm(width_cm).emu / 635)
        if idx < len(grid_cols):
            grid_cols[idx].set(qn("w:w"), str(width_twips))
        if idx < len(table.columns):
            table.columns[idx].width = Cm(width_cm)
        for row in table.rows:
            if idx < len(row.cells):
                _set_cell_width(row.cells[idx], width_cm)

def _prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def _repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def _collect_item_images(row, item_id):
    images = []
    file_name = str(row.get("저장된사진파일명", "") or "").strip()
    if file_name:
        candidates = [file_name, os.path.basename(file_name)]
        for candidate in candidates:
            if candidate and os.path.exists(candidate):
                try:
                    with open(candidate, "rb") as f:
                        images.append(f.read())
                    break
                except Exception:
                    pass

    for web_img in web_images_map.get(item_id, []):
        data = web_img.get("data")
        if data:
            images.append(data)
    return images


def build_as_request_docx(
    source_df,
    selected_statuses,
    as_date="",
    manager_name="",
    message="",
    building="204동",
    unit="4503호",
    phone="",
):
    """첨부된 기존 A/S 신청서와 유사한 표 형식으로 Word 신청서를 생성한다."""
    selected = source_df[source_df["진행상태"].isin(selected_statuses)].copy()

    doc = Document()
    section = doc.sections[0]
    # 원본 신청서: Letter 용지, 약 1.59cm 여백
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Cm(1.59)
    section.bottom_margin = Cm(1.59)
    section.left_margin = Cm(1.59)
    section.right_margin = Cm(1.59)

    normal = doc.styles["Normal"]
    normal.font.name = "맑은 고딕"
    normal.font.size = Pt(9)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(8)
    title_run = title.add_run("A/S 신청서")
    title_run.bold = True
    title_run.font.size = Pt(18)

    meta_table = doc.add_table(rows=2, cols=3)
    meta_table.style = "Table Grid"
    meta_widths = [3.88, 3.88, 10.65]
    _set_table_fixed_widths(meta_table, meta_widths)
    for idx, label in enumerate(["날짜", "매니저명", "전달사항"]):
        _set_cell_text(meta_table.cell(0, idx), label, bold=True)
    default_msg = f"{', '.join(selected_statuses)} 하자 {len(selected)}건 A/S 요청드립니다."
    _set_cell_text(meta_table.cell(1, 0), as_date or datetime.now(ZoneInfo("Asia/Seoul")).strftime("%m/%d"))
    _set_cell_text(meta_table.cell(1, 1), manager_name)
    _set_cell_text(meta_table.cell(1, 2), message or default_msg)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    info_table = doc.add_table(rows=2, cols=6)
    info_table.style = "Table Grid"
    info_widths = [3.88, 1.86, 1.33, 1.33, 1.33, 8.75]
    labels = ["동", "호", "임시\n키불", "키불출", "입주", "전화"]
    values = [building, unit, "", "", "", phone]
    _set_table_fixed_widths(info_table, info_widths)
    for idx, label in enumerate(labels):
        _set_cell_text(info_table.cell(0, idx), label, bold=True, font_size=8 if idx in [2, 3, 4] else 9)
        _set_cell_text(info_table.cell(1, idx), values[idx])

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    defect_table = doc.add_table(rows=1, cols=4)
    defect_table.style = "Table Grid"
    defect_widths = [1.23, 2.47, 12.24, 2.47]
    _set_table_fixed_widths(defect_table, defect_widths)
    header = defect_table.rows[0]
    _repeat_table_header(header)
    for idx, label in enumerate(["NO", "실", "내용", "공종"]):
        _set_cell_text(header.cells[idx], label, bold=True)

    if selected.empty:
        row_cells = defect_table.add_row().cells
        for c, w in zip(row_cells, defect_widths):
            _set_cell_width(c, w)
        row_cells[0].merge(row_cells[3])
        _set_cell_text(row_cells[0], "선택한 진행상태에 해당하는 하자 항목이 없습니다.")
    else:
        for seq, (_, row) in enumerate(selected.iterrows(), start=1):
            item_id = str(row.get("번호", "") or "").strip()
            space_text = str(row.get("공간", "") or "").strip()
            part_text = str(row.get("부위", "") or "").strip()
            type_text = str(row.get("유형", "") or "").strip()
            detail_text = str(row.get("상세내용", "") or "").strip()
            trade = infer_trade(part_text, type_text)

            defect_row = defect_table.add_row()
            _prevent_row_split(defect_row)
            cells = defect_row.cells
            for c, w in zip(cells, defect_widths):
                _set_cell_width(c, w)
            _set_cell_text(cells[0], seq)
            _set_cell_text(cells[1], space_text)
            _set_cell_text(cells[3], trade)

            cells[2].text = ""
            cells[2].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            text_p = cells[2].paragraphs[0]
            text_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            text_p.paragraph_format.space_after = Pt(2)
            text_run = text_p.add_run(f"{space_text} {part_text} / {type_text} - {detail_text}")
            text_run.font.size = Pt(9)

            images = _collect_item_images(row, item_id)
            if images:
                for img_bytes in images:
                    try:
                        img_p = cells[2].add_paragraph()
                        img_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        img_p.paragraph_format.space_before = Pt(1)
                        img_p.paragraph_format.space_after = Pt(1)
                        img_p.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(8.15))
                    except Exception:
                        pass
            else:
                no_img = cells[2].add_paragraph()
                no_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = no_img.add_run("(사진 없음)")
                r.font.size = Pt(8)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue(), len(selected)


# -------------------------
# 진행상태 대시보드
# -------------------------
total = len(df)
done = int((df["진행상태"] == "확인완료").sum())
incomplete = int((df["진행상태"] == "미완료").sum())
progress = int((done / total) * 100) if total > 0 else 0

st.markdown(
    f"""
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">총 하자</div><div class="kpi-value">{total}</div><div class="kpi-meta">현재 세대 전체 항목</div></div>
      <div class="kpi-card"><div class="kpi-label">확인완료</div><div class="kpi-value">{done}</div><div class="kpi-meta">처리 확인된 항목</div></div>
      <div class="kpi-card"><div class="kpi-label">미완료</div><div class="kpi-value">{incomplete}</div><div class="kpi-meta">추가 조치가 필요한 항목</div></div>
      <div class="kpi-card"><div class="kpi-label">진행률</div><div class="kpi-value">{progress}%</div><div class="progress-shell"><div class="progress-fill" style="width:{progress}%"></div></div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

web_added_count = int((df["데이터출처"] == "웹등록").sum())
st.caption(f"웹에서 추가한 하자 {web_added_count}건 · 진행상태는 미완료 / 확인완료 두 단계로 관리합니다.")

st.markdown('<div class="section-label">공간별 진행현황</div><div class="section-caption">미완료 항목이 많은 공간부터 표시합니다.</div>', unsafe_allow_html=True)
if df.empty:
    st.info("표시할 하자 데이터가 없습니다.")
else:
    space_summary = (
        df.groupby("공간", dropna=False)
        .agg(
            전체=("번호", "count"),
            확인완료=("진행상태", lambda x: int((x == "확인완료").sum())),
            미완료=("진행상태", lambda x: int((x == "미완료").sum())),
        )
        .reset_index()
    )
    space_summary["공간"] = space_summary["공간"].fillna("미지정").replace("", "미지정")
    space_summary["진행률값"] = (
        (space_summary["확인완료"] / space_summary["전체"] * 100)
        .fillna(0).round(0).astype(int)
    )
    # 미완료가 많은 공간을 먼저 확인할 수 있도록 미완료 내림차순, 진행률 오름차순으로 정렬
    space_summary = space_summary.sort_values(["미완료", "진행률값", "전체"], ascending=[False, True, False])
    space_summary["진행률"] = space_summary["진행률값"].astype(str) + "%"
    st.dataframe(
        space_summary[["공간", "전체", "확인완료", "미완료", "진행률"]],
        use_container_width=True,
        hide_index=True,
    )

# -------------------------
# 웹 신규 하자 등록
# -------------------------
with st.expander("새 하자 등록", expanded=False):
    st.caption(
        "새 항목은 현재 로그인한 세대에만 저장되며 Excel을 수정하지 않고 Google Sheet에 별도로 저장됩니다. "
        "공간·부위·유형은 원본 Excel에 있는 전체 목록에서 각각 독립적으로 선택하며 사진은 최대 5장까지 첨부할 수 있습니다."
    )

    if not REFERENCE_SPACES or not REFERENCE_PARTS or not REFERENCE_TYPES:
        st.error("원본 Excel에서 공간/부위/유형 기준값을 읽지 못했습니다.")
        new_space = new_part = new_type = ""
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            new_space = st.selectbox("공간 *", REFERENCE_SPACES, key="new_issue_space")
        with c2:
            new_part = st.selectbox("부위 *", REFERENCE_PARTS, key="new_issue_part")
        with c3:
            new_type = st.selectbox("유형 *", REFERENCE_TYPES, key="new_issue_type")

    c_status, c_info = st.columns([1, 2])
    with c_status:
        new_initial_status = st.selectbox(
            "등록 후 진행상태",
            WEB_STATUS_OPTIONS,
            index=0,
            key="new_issue_initial_status",
        )
    with c_info:
        st.info(
            f"선택: {new_space or '-'} / {new_part or '-'} / {new_type or '-'}",
            icon="📌",
        )

    new_detail = st.text_area(
        "상세내용 *",
        placeholder="하자 위치와 증상을 자세히 입력해 주세요.",
        height=110,
        key="new_issue_detail",
    )
    new_photos = st.file_uploader(
        "사진 첨부 (최대 5장)",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        help="사진은 Google Sheet에 압축하여 별도 저장됩니다.",
        key="new_issue_photos",
    )

    if new_photos:
        preview_cols = st.columns(min(len(new_photos), 5))
        for p_idx, uploaded in enumerate(new_photos[:MAX_UPLOAD_IMAGES]):
            with preview_cols[p_idx % len(preview_cols)]:
                st.image(uploaded, caption=f"사진 {p_idx + 1}", use_container_width=True)

    submitted = st.button(
        "새 하자 등록",
        use_container_width=True,
        disabled=not web_status_enabled,
        key="submit_new_issue",
        type="primary",
    )

    if submitted:
        if not new_space or not new_part or not new_type or not new_detail.strip():
            st.error("공간, 부위, 유형, 상세내용은 모두 필수입니다.")
        elif len(new_photos or []) > MAX_UPLOAD_IMAGES:
            st.error(f"사진은 최대 {MAX_UPLOAD_IMAGES}장까지 첨부할 수 있습니다.")
        else:
            try:
                new_id = create_web_item(
                    new_space,
                    new_part,
                    new_type,
                    new_detail.strip(),
                    new_initial_status,
                    new_photos or [],
                    CURRENT_BUILDING,
                    CURRENT_UNIT,
                )
                st.success(f"새 하자가 등록되었습니다. ID: {new_id}")
                # 다음 등록을 위해 입력값 일부 초기화
                for _key in ["new_issue_detail", "new_issue_photos"]:
                    if _key in st.session_state:
                        del st.session_state[_key]
                st.rerun()
            except Exception as exc:
                st.error(f"새 하자 등록 실패: {type(exc).__name__}: {exc}")

# AS 신청서 출력: 진행상태 기준
with st.expander("A/S 신청서 출력", expanded=False):
    st.caption(
        "첨부해주신 기존 A/S 신청서처럼 `날짜/매니저명/전달사항`, 세대정보, "
        "`NO/실/내용/공종` 표와 사진이 함께 출력됩니다."
    )

    as_statuses = st.multiselect(
        "출력할 진행상태",
        WEB_STATUS_OPTIONS,
        default=["미완료"],
        key="as_web_statuses",
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        as_date = st.text_input(
            "날짜",
            value=datetime.now(ZoneInfo("Asia/Seoul")).strftime("%m/%d"),
            key="as_date",
        )
        as_building = st.text_input("동", value=CURRENT_BUILDING, disabled=True, key="as_building")
    with c2:
        as_manager = st.text_input("매니저명", value="", key="as_manager")
        as_unit = st.text_input("호", value=CURRENT_UNIT, disabled=True, key="as_unit")
    with c3:
        as_phone = st.text_input("전화", value="", key="as_phone")

    as_target_count = int(df["진행상태"].isin(as_statuses).sum()) if as_statuses else 0
    default_message = f"{', '.join(as_statuses)} 하자 {as_target_count}건 A/S 요청드립니다." if as_statuses else ""
    as_message = st.text_input("전달사항", value=default_message, key="as_message")

    if as_statuses:
        st.info(f"현재 AS 신청서 출력 대상: {as_target_count}건")
        as_docx_bytes, _ = build_as_request_docx(
            df,
            as_statuses,
            as_date=as_date,
            manager_name=as_manager,
            message=as_message,
            building=as_building,
            unit=as_unit,
            phone=as_phone,
        )
        st.download_button(
            "A/S 신청서 다운로드 (.docx)",
            data=as_docx_bytes,
            file_name=f"AS신청서_{datetime.now(ZoneInfo('Asia/Seoul')).strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    else:
        st.warning("출력할 진행상태를 하나 이상 선택해 주세요.")

st.markdown("---")

st.markdown('<div class="section-label">하자 목록</div><div class="section-caption">공간과 진행상태를 선택해 필요한 항목만 빠르게 확인하세요.</div>', unsafe_allow_html=True)

# -------------------------
# 평면도 / 필터
# -------------------------
with st.expander("공간 위치 확인 (평면도)"):
    if os.path.exists("image_5012c2.jpg"):
        st.image("image_5012c2.jpg", use_container_width=True)
    else:
        st.info("평면도 이미지(image_5012c2.jpg)를 찾을 수 없습니다.")

filter_col1, filter_col2 = st.columns(2)

with filter_col1:
    spaces = sorted([str(x) for x in df["공간"].dropna().unique().tolist()])
    space = st.selectbox("공간별 필터링", ["전체"] + spaces)

with filter_col2:
    progress_filter = st.selectbox("진행상태", ["전체"] + WEB_STATUS_OPTIONS)

target_df = df.copy()
if space != "전체":
    target_df = target_df[target_df["공간"] == space]
if progress_filter != "전체":
    target_df = target_df[target_df["진행상태"] == progress_filter]

st.caption(f"현재 조건에 {len(target_df)}건이 표시됩니다.")

# -------------------------
# 하자 카드
# -------------------------
cols = st.columns(2)
for i, (index, row) in enumerate(target_df.iterrows()):
    with cols[i % 2]:
        item_id = str(row["번호"]).strip()
        current_web_status = str(row.get("진행상태", "미완료") or "미완료")
        updated_at = web_status_map.get(item_id, {}).get("updated_at", "")

        badge_class = {
            "미완료": "status-incomplete",
            "확인완료": "status-checked",
        }.get(current_web_status, "status-incomplete")
        status_color = {
            "확인완료": "#27ae60",
            "미완료": "#ef4444",
        }.get(current_web_status, "#ef4444")

        space_text = str(row.get("공간", ""))
        part_text = str(row.get("부위", ""))
        type_text = str(row.get("유형", ""))
        detail_text = str(row.get("상세내용", ""))
        item_label = f"[{item_id}] {space_text} - {part_text}"

        safe_id = html.escape(item_id)
        safe_title = html.escape(f"{space_text} · {part_text}")
        safe_type = html.escape(type_text)
        safe_detail = html.escape(detail_text)
        safe_updated = html.escape(str(updated_at))
        st.markdown(
            f"""
            <div class="issue-card">
              <div class="issue-head">
                <div>
                  <div class="issue-id">#{safe_id}</div>
                  <div class="issue-title">{safe_title}</div>
                </div>
                <span class="status-badge {badge_class}">{html.escape(current_web_status)}</span>
              </div>
              <div class="issue-detail"><b>{safe_type}</b> · {safe_detail}</div>
              {f'<div class="status-meta">마지막 변경 · {safe_updated}</div>' if updated_at else ''}
            </div>
            """,
            unsafe_allow_html=True,
        )

        # 버튼을 누르는 즉시 Google Sheet에만 저장한다.
        b1, b2 = st.columns(2)
        with b1:
            if st.button(
                "❌ 미완료",
                key=f"incomplete_{item_id}",
                use_container_width=True,
                disabled=not web_status_enabled or current_web_status == "미완료",
            ):
                try:
                    save_web_status(item_id, "미완료", item_label)
                    st.toast(f"{item_label} → 미완료", icon="❌")
                    st.rerun()
                except Exception as exc:
                    st.error(f"진행상태 저장 실패: {exc}")

        with b2:
            if st.button(
                "✅ 확인완료",
                key=f"checked_{item_id}",
                use_container_width=True,
                disabled=not web_status_enabled or current_web_status == "확인완료",
            ):
                try:
                    save_web_status(item_id, "확인완료", item_label)
                    st.toast(f"{item_label} → 확인완료", icon="✅")
                    st.rerun()
                except Exception as exc:
                    st.error(f"진행상태 저장 실패: {exc}")

        # 삭제는 웹에서 신규 등록한 항목에만 제공한다. Excel 원본 항목은 보호한다.
        if str(row.get("데이터출처", "")) == "웹등록":
            confirm_key = f"delete_confirm_{item_id}"
            if st.session_state.get(confirm_key, False):
                st.warning("이 웹 등록 하자를 목록에서 삭제할까요? 원본 Excel에는 영향이 없습니다.")
                dc1, dc2 = st.columns(2)
                with dc1:
                    if st.button(
                        "🗑️ 정말 삭제",
                        key=f"delete_yes_{item_id}",
                        type="primary",
                        use_container_width=True,
                    ):
                        try:
                            delete_web_item(item_id)
                            st.session_state.pop(confirm_key, None)
                            st.toast(f"{item_label} 삭제 완료", icon="🗑️")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"삭제 실패: {type(exc).__name__}: {exc}")
                with dc2:
                    if st.button(
                        "취소",
                        key=f"delete_no_{item_id}",
                        use_container_width=True,
                    ):
                        st.session_state.pop(confirm_key, None)
                        st.rerun()
            else:
                if st.button(
                    "🗑️ 신규 등록 항목 삭제",
                    key=f"delete_start_{item_id}",
                    use_container_width=True,
                ):
                    st.session_state[confirm_key] = True
                    st.rerun()

        file_name = str(row.get("저장된사진파일명", "")).strip()
        if file_name and os.path.exists(file_name):
            st.image(file_name, use_container_width=True)

        web_imgs = web_images_map.get(item_id, [])
        if web_imgs:
            if len(web_imgs) == 1:
                st.image(web_imgs[0]["data"], use_container_width=True)
            else:
                img_cols = st.columns(min(len(web_imgs), 3))
                for img_idx, web_img in enumerate(web_imgs):
                    with img_cols[img_idx % len(img_cols)]:
                        st.image(web_img["data"], use_container_width=True)
