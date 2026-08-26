import base64
import glob
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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from PIL import Image, ImageOps
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound

# -------------------------
# 기본 설정
# -------------------------
st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

WEB_STATUS_OPTIONS = ["미확인", "확인완료", "재확인필요"]
WEB_STATUS_SHEET_NAME = "web_status"
WEB_ITEMS_SHEET_NAME = "web_items"
WEB_IMAGES_SHEET_NAME = "web_images"
MAX_UPLOAD_IMAGES = 5
IMAGE_CHUNK_SIZE = 40000
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

st.markdown(
    """
    <style>
    .summary-card {
        background: #1B2845;
        color: #ffffff;
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 20px;
    }
    .card {
        background: #ffffff;
        padding: 25px;
        border-radius: 15px;
        border-left: 10px solid #1B2845;
        box-shadow: 0 4px 10px rgba(0,0,0,0.08);
        margin-bottom: 10px;
        color: #333333;
    }
    .card h3 { color: #1B2845; }
    .status-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        font-size: 0.9rem;
        font-weight: 700;
        margin-left: 6px;
    }
    .status-unchecked { background: #eef1f5; color: #475569; }
    .status-checked { background: #dcfce7; color: #166534; }
    .status-recheck { background: #fef3c7; color: #92400e; }
    .web-status-box {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 10px 12px;
        margin: 0 0 14px 0;
        background: #fafafa;
    }
    .stSelectbox { border: 1px solid #1B2845; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🏢 해링턴 플레이스 하자 관리 대시보드")

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
        ws = spreadsheet.add_worksheet(title=name, rows=rows, cols=cols)
        ws.append_row(headers)
    if not ws.get_all_values():
        ws.append_row(headers)
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
        status = str(record.get("web_status", "")).strip()
        if item_id and status in WEB_STATUS_OPTIONS:
            result[item_id] = {
                "status": status,
                "updated_at": str(record.get("updated_at", "")).strip(),
            }
    return result


def save_web_status(item_id, new_status, item_label):
    """item_id 기준으로 웹 상태를 insert/update한다. Excel은 건드리지 않는다."""
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
        ["item_id", "공간", "부위", "유형", "상세내용", "created_at", "active"],
        rows=500,
        cols=7,
    )
    records = ws.get_all_records()
    rows = []
    for r in records:
        if str(r.get("active", "TRUE")).strip().upper() in {"FALSE", "0", "N", "NO"}:
            continue
        item_id = str(r.get("item_id", "")).strip()
        if not item_id:
            continue
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


def create_web_item(space, part, issue_type, detail, initial_status, uploaded_files):
    ws = get_or_create_worksheet(
        WEB_ITEMS_SHEET_NAME,
        ["item_id", "공간", "부위", "유형", "상세내용", "created_at", "active"],
        rows=500,
        cols=7,
    )
    now = datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y-%m-%d %H:%M:%S")
    item_id = f"WEB-{datetime.now(ZoneInfo('Asia/Seoul')).strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"
    ws.append_row([item_id, space, part, issue_type, detail, now, "TRUE"])
    item_label = f"[{item_id}] {space} - {part}"
    save_web_status(item_id, initial_status, item_label)
    save_web_images(item_id, uploaded_files)
    load_web_items.clear()
    load_web_status.clear()
    return item_id


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
        "웹 확인상태 저장 기능이 아직 연결되지 않았습니다. 기존 Excel 진행현황 조회는 정상 동작합니다.\n\n"
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

# Excel 데이터에 출처 표시 후 웹 등록 데이터를 병합한다.
df["데이터출처"] = "Excel"
df["등록일시"] = ""
if not web_items_df.empty:
    for col in df.columns:
        if col not in web_items_df.columns:
            web_items_df[col] = ""
    for col in web_items_df.columns:
        if col not in df.columns:
            df[col] = ""
    df = pd.concat([df, web_items_df[df.columns]], ignore_index=True)

# 각 행에 웹 상태를 붙인다. Excel 번호 또는 WEB-... ID를 키로 사용한다.
df["웹확인상태"] = df["번호"].apply(
    lambda x: web_status_map.get(str(x).strip(), {}).get("status", "미확인")
)

# -------------------------
# AS 신청서 생성 (웹 확인상태 기준)
# -------------------------
def build_as_request_docx(source_df, selected_web_statuses):
    """선택한 웹 확인상태의 항목으로 AS 신청서 Word 파일을 만든다.

    Excel의 진행현황은 출력 참고정보로만 표시하며 대상 선정에는 사용하지 않는다.
    """
    selected = source_df[source_df["웹확인상태"].isin(selected_web_statuses)].copy()

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("해링턴 플레이스 AS 신청서")
    run.bold = True
    run.font.size = Pt(18)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        f"출력 기준: 웹 확인상태 ({', '.join(selected_web_statuses)})  |  "
        f"대상 {len(selected)}건  |  생성일 {datetime.now(ZoneInfo('Asia/Seoul')).strftime('%Y-%m-%d %H:%M')}"
    )

    doc.add_paragraph(
        "※ AS 신청 대상은 웹 확인상태를 기준으로 선정되었습니다. "
        "Excel 진행현황은 참고용으로만 표시됩니다."
    )

    if selected.empty:
        doc.add_paragraph("선택한 웹 확인상태에 해당하는 하자 항목이 없습니다.")
    else:
        for seq, (_, row) in enumerate(selected.iterrows(), start=1):
            item_id = str(row.get("번호", "")).strip()
            space_text = str(row.get("공간", "") or "")
            part_text = str(row.get("부위", "") or "")
            type_text = str(row.get("유형", "") or "")
            detail_text = str(row.get("상세내용", "") or "")
            excel_status = str(row.get("진행현황_표시", "") or "")
            web_status = str(row.get("웹확인상태", "") or "")

            heading = doc.add_paragraph()
            r = heading.add_run(f"{seq}. [{item_id}] {space_text} - {part_text}")
            r.bold = True
            r.font.size = Pt(12)

            table = doc.add_table(rows=4, cols=2)
            table.style = "Table Grid"
            fields = [
                ("웹 확인상태", web_status),
                ("Excel 진행현황", excel_status),
                ("유형", type_text),
                ("상세내용", detail_text),
            ]
            for idx, (label, value) in enumerate(fields):
                table.cell(idx, 0).text = label
                table.cell(idx, 1).text = value

            # 기존 Excel 사진 + 웹 신규등록 사진 모두 출력
            file_name = str(row.get("저장된사진파일명", "") or "").strip()
            if file_name and os.path.exists(file_name):
                try:
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(file_name, width=Inches(4.8))
                except Exception:
                    doc.add_paragraph(f"사진 파일: {file_name}")

            for web_img in web_images_map.get(item_id, []):
                try:
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(io.BytesIO(web_img["data"]), width=Inches(4.8))
                except Exception:
                    pass

            if seq != len(selected):
                doc.add_paragraph("―" * 35)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue(), len(selected)


# -------------------------
# 기존 Excel 진행현황 요약
# -------------------------
excel_only_df = df[df["데이터출처"] == "Excel"].copy()
total = len(excel_only_df)
status_counts = excel_only_df["진행현황_표시"].value_counts()
done = status_counts.get("완료", 0)
todo = total - done
progress = int((done / total) * 100) if total > 0 else 0

st.markdown(
    f"""
    <div class='summary-card'>
        <h3>전체 하자 처리 현황 (Excel 기준)</h3>
        <div style='font-size: 40px; font-weight: bold;'>{done} / {total} 건</div>
        <p>전체 진행률 {progress}%</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.progress(progress / 100)

d1, d2, d3 = st.columns(3)
d1.metric("총 하자 건수", f"{total}건")
d2.metric("완료 건수", f"{done}건", delta=f"{progress}%")
d3.metric("미조치 건수", f"{todo}건", delta_color="inverse")

# 웹 상태 요약은 Excel 상태와 완전히 별도로 표시
web_unchecked = int((df["웹확인상태"] == "미확인").sum())
web_checked = int((df["웹확인상태"] == "확인완료").sum())
web_recheck = int((df["웹확인상태"] == "재확인필요").sum())

st.subheader("웹 확인 현황")
w1, w2, w3 = st.columns(3)
w1.metric("미확인", f"{web_unchecked}건")
w2.metric("확인완료", f"{web_checked}건")
w3.metric("재확인필요", f"{web_recheck}건")

web_added_count = int((df["데이터출처"] == "웹등록").sum())
st.caption(f"웹에서 새로 등록한 하자: {web_added_count}건")

# -------------------------
# 웹 신규 하자 등록
# -------------------------
with st.expander("➕ 새 하자 등록", expanded=False):
    st.caption("새 항목은 Excel을 수정하지 않고 Google Sheet에 별도로 저장됩니다. 사진은 최대 5장까지 첨부할 수 있습니다.")
    with st.form("new_issue_form", clear_on_submit=True):
        f1, f2 = st.columns(2)
        with f1:
            new_space = st.text_input("공간 *", placeholder="예: 거실, 침실1, 주방")
            new_type = st.text_input("유형", placeholder="예: 파손, 오염, 들뜸")
        with f2:
            new_part = st.text_input("부위 *", placeholder="예: 벽지, 문틀, 바닥")
            new_initial_status = st.selectbox("등록 후 웹 확인상태", WEB_STATUS_OPTIONS, index=2)
        new_detail = st.text_area("상세내용 *", placeholder="하자 위치와 증상을 자세히 입력해 주세요.", height=110)
        new_photos = st.file_uploader(
            "사진 첨부 (최대 5장)",
            type=["jpg", "jpeg", "png", "webp"],
            accept_multiple_files=True,
            help="사진은 Google Sheet에 압축하여 별도 저장됩니다.",
        )
        submitted = st.form_submit_button("💾 새 하자 등록", use_container_width=True, disabled=not web_status_enabled)

    if submitted:
        if not new_space.strip() or not new_part.strip() or not new_detail.strip():
            st.error("공간, 부위, 상세내용은 필수입니다.")
        elif len(new_photos or []) > MAX_UPLOAD_IMAGES:
            st.error(f"사진은 최대 {MAX_UPLOAD_IMAGES}장까지 첨부할 수 있습니다.")
        else:
            try:
                new_id = create_web_item(
                    new_space.strip(),
                    new_part.strip(),
                    new_type.strip(),
                    new_detail.strip(),
                    new_initial_status,
                    new_photos or [],
                )
                st.success(f"새 하자가 등록되었습니다. ID: {new_id}")
                st.rerun()
            except Exception as exc:
                st.error(f"새 하자 등록 실패: {type(exc).__name__}: {exc}")

# AS 신청서 출력: Excel 상태가 아니라 웹 확인상태를 기준으로 대상 선정
with st.expander("📝 AS 신청서 출력", expanded=False):
    st.caption("AS 신청 대상은 아래에서 선택한 **웹 확인상태** 기준으로 생성됩니다.")
    as_statuses = st.multiselect(
        "출력할 웹 확인상태",
        WEB_STATUS_OPTIONS,
        default=["재확인필요"],
        key="as_web_statuses",
    )

    if as_statuses:
        as_target_count = int(df["웹확인상태"].isin(as_statuses).sum())
        st.info(f"현재 AS 신청서 출력 대상: {as_target_count}건")
        as_docx_bytes, _ = build_as_request_docx(df, as_statuses)
        st.download_button(
            "📄 AS 신청서 다운로드 (.docx)",
            data=as_docx_bytes,
            file_name=f"AS신청서_웹상태_{datetime.now(ZoneInfo('Asia/Seoul')).strftime('%Y%m%d_%H%M')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    else:
        st.warning("출력할 웹 확인상태를 하나 이상 선택해 주세요.")

st.markdown("---")

# -------------------------
# 평면도 / 필터
# -------------------------
with st.expander("🗺️ 공간 위치 확인 (평면도)"):
    if os.path.exists("image_5012c2.jpg"):
        st.image("image_5012c2.jpg", use_container_width=True)
    else:
        st.info("평면도 이미지(image_5012c2.jpg)를 찾을 수 없습니다.")

filter_col1, filter_col2, filter_col3 = st.columns(3)

with filter_col1:
    spaces = [x for x in df["공간"].dropna().unique().tolist()]
    space = st.selectbox("공간별 필터링", ["전체"] + spaces)

with filter_col2:
    status_filter = st.selectbox("Excel 진행상태", ["전체", "완료", "미완료", "웹등록"])

with filter_col3:
    web_filter = st.selectbox(
        "웹 확인상태",
        ["전체"] + WEB_STATUS_OPTIONS,
    )

target_df = df.copy()

if space != "전체":
    target_df = target_df[target_df["공간"] == space]

if status_filter == "완료":
    target_df = target_df[target_df["진행현황_표시"] == "완료"]
elif status_filter == "미완료":
    target_df = target_df[(target_df["데이터출처"] == "Excel") & (target_df["진행현황_표시"] != "완료")]
elif status_filter == "웹등록":
    target_df = target_df[target_df["데이터출처"] == "웹등록"]

if web_filter != "전체":
    target_df = target_df[target_df["웹확인상태"] == web_filter]

st.caption(f"현재 조건에 {len(target_df)}건이 표시됩니다.")

# -------------------------
# 하자 카드
# -------------------------
cols = st.columns(2)
for i, (index, row) in enumerate(target_df.iterrows()):
    with cols[i % 2]:
        excel_status = row["진행현황_표시"]
        status_color = "#27ae60" if excel_status == "완료" else "#e74c3c"

        item_id = str(row["번호"]).strip()
        current_web_status = web_status_map.get(item_id, {}).get("status", "미확인")
        updated_at = web_status_map.get(item_id, {}).get("updated_at", "")

        badge_class = {
            "미확인": "status-unchecked",
            "확인완료": "status-checked",
            "재확인필요": "status-recheck",
        }.get(current_web_status, "status-unchecked")

        space_text = str(row.get("공간", ""))
        part_text = str(row.get("부위", ""))
        type_text = str(row.get("유형", ""))
        detail_text = str(row.get("상세내용", ""))
        item_label = f"[{item_id}] {space_text} - {part_text}"

        st.markdown(
            f"""
            <div class='card' style='border-left-color: {status_color};'>
                <h3>{item_label}</h3>
                <p><b>원본 상태:</b> {excel_status} <span style="color:#64748b;">({row.get("데이터출처", "Excel")})</span></p>
                <p><b>상세:</b> {type_text} / {detail_text}</p>
            </div>
            <div class='web-status-box'>
                <b>웹 확인상태:</b>
                <span class='status-badge {badge_class}'>{current_web_status}</span>
                {f"<div style='font-size:0.78rem;color:#64748b;margin-top:5px;'>마지막 변경: {updated_at}</div>" if updated_at else ""}
            </div>
            """,
            unsafe_allow_html=True,
        )

        # 버튼을 누르는 즉시 Google Sheet에만 저장한다.
        b1, b2, b3 = st.columns(3)
        with b1:
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
                    st.error(f"웹 상태 저장 실패: {exc}")

        with b2:
            if st.button(
                "⚠️ 재확인",
                key=f"recheck_{item_id}",
                use_container_width=True,
                disabled=not web_status_enabled or current_web_status == "재확인필요",
            ):
                try:
                    save_web_status(item_id, "재확인필요", item_label)
                    st.toast(f"{item_label} → 재확인필요", icon="⚠️")
                    st.rerun()
                except Exception as exc:
                    st.error(f"웹 상태 저장 실패: {exc}")

        with b3:
            if st.button(
                "↩ 미확인",
                key=f"unchecked_{item_id}",
                use_container_width=True,
                disabled=not web_status_enabled or current_web_status == "미확인",
            ):
                try:
                    save_web_status(item_id, "미확인", item_label)
                    st.toast(f"{item_label} → 미확인", icon="↩️")
                    st.rerun()
                except Exception as exc:
                    st.error(f"웹 상태 저장 실패: {exc}")

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
