import streamlit as st
import pandas as pd
import os
import glob
import io
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 페이지 설정
st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# CSS: 다크 모드 대응 및 전문가적인 디자인
st.markdown("""
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
        margin-bottom: 20px; 
        color: #333333;
    }
    .card h3 { color: #1B2845; }
    .stSelectbox { border: 1px solid #1B2845; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏢 해링턴 플레이스 하자 관리 대시보드")

# 엑셀 파일 자동 탐색
excel_candidates = glob.glob("*.xlsx")
if not excel_candidates:
    st.error("저장소 안에서 엑셀(.xlsx) 파일을 찾지 못했습니다.")
    st.stop()
EXCEL_PATH = excel_candidates[0]

@st.cache_data
def load_data(mtime, path):
    return pd.read_excel(path, sheet_name='Sheet1')

mtime = os.path.getmtime(EXCEL_PATH)
df = load_data(mtime, EXCEL_PATH)
df.columns = [str(c).strip() for c in df.columns]

# --- 대시보드 요약 ---
total = len(df)
status_counts = df['진행현황'].value_counts() if '진행현황' in df.columns else pd.Series()
done = status_counts.get('완료', 0)
todo = total - done
progress = int((done / total) * 100) if total > 0 else 0

st.markdown(f"""
    <div class='summary-card'>
        <h3>전체 하자 처리 현황</h3>
        <div style='font-size: 40px; font-weight: bold;'>{done} / {total} 건</div>
        <p>전체 진행률 {progress}%</p>
    </div>
""", unsafe_allow_html=True)

st.progress(progress / 100)

d1, d2, d3 = st.columns(3)
d1.metric("총 하자 건수", f"{total}건")
d2.metric("완료 건수", f"{done}건", delta=f"{progress}%")
d3.metric("미조치 건수", f"{todo}건", delta_color="inverse")

st.markdown("---")

# --- AS신청서 출력 기능 ---

def set_cell_background(cell, color_hex):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shd)

def style_header_cell(cell, text):
    cell.text = text
    set_cell_background(cell, 'F2F2F2')
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True

def gongjong(bu):
    bu = str(bu)
    if bu in ['목문', '창문', '신발장', '하부장', '상부장', '서랍장', '냉장고장', '걸레받이', '수납장', '목문틀']:
        return '목공'
    if bu in ['벽도배', '천정도배']:
        return '도배'
    if bu in ['벽도장', '문틀도장']:
        return '도장'
    if bu == '바닥타일':
        return '타일'
    if bu == '코킹':
        return '코킹'
    if bu == '스위치':
        return '전기'
    if bu == '설비':
        return '설비'
    if bu == '마루':
        return '마루'
    return '기타'

def build_as_docx(target_df):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.59)
    section.page_height = Cm(27.94)
    section.left_margin = section.right_margin = Cm(1.5)
    section.top_margin = section.bottom_margin = Cm(1.5)

    title = doc.add_heading('A/S 신청서', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 상단 정보 테이블
    t1 = doc.add_table(rows=2, cols=3)
    t1.style = 'Table Grid'
    style_header_cell(t1.rows[0].cells[0], '날짜')
    style_header_cell(t1.rows[0].cells[1], '매니저명')
    style_header_cell(t1.rows[0].cells[2], '전달사항')
    t1.rows[1].cells[2].text = f'미완료 하자 {len(target_df)}건 A/S 요청드립니다.'

    doc.add_paragraph('')

    # 동/호/임시키불/키불출/입주/전화 (각각 별도 컬럼으로 분리)
    t2 = doc.add_table(rows=2, cols=6)
    t2.style = 'Table Grid'
    for i, label in enumerate(['동', '호', '임시키불', '키불출', '입주', '전화']):
        style_header_cell(t2.rows[0].cells[i], label)

    doc.add_paragraph('')

    # 메인 리스트: NO / 실 / 내용(+사진) / 공종
    t3 = doc.add_table(rows=1, cols=4)
    t3.style = 'Table Grid'
    for i, label in enumerate(['NO', '실', '내용', '공종']):
        style_header_cell(t3.rows[0].cells[i], label)

    widths = [Cm(1.3), Cm(2.5), Cm(11.0), Cm(2.3)]
    for idx, w in enumerate(widths):
        t3.rows[0].cells[idx].width = w

    for _, row in target_df.iterrows():
        cells = t3.add_row().cells
        cells[0].text = str(row['번호'])
        cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cells[1].text = str(row['공간'])
        cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        content_cell = cells[2]
        content_cell.paragraphs[0].text = f"{row['공간']} {row['부위']} / {row['유형']} - {row['상세내용']}"

        photo_name = str(row.get('저장된사진파일명', '')).strip()
        if photo_name and os.path.exists(photo_name):
            p = content_cell.add_paragraph()
            run = p.add_run()
            try:
                run.add_picture(photo_name, width=Cm(7.5))
            except Exception:
                p2 = content_cell.add_paragraph()
                p2.add_run('(사진 삽입 실패)').italic = True
        else:
            p = content_cell.add_paragraph()
            p.add_run('(사진 없음)').italic = True

        cells[3].text = gongjong(row['부위'])
        cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        for idx, w in enumerate(widths):
            cells[idx].width = w

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

col_btn1, col_btn2 = st.columns([1, 3])
with col_btn1:
    generate_clicked = st.button("📄 AS신청서 출력")

if generate_clicked:
    todo_df = df[df['진행현황'].astype(str).str.strip() != '완료']
    with st.spinner(f"미완료 {len(todo_df)}건으로 문서 생성 중..."):
        docx_buf = build_as_docx(todo_df)
    st.success(f"미완료 {len(todo_df)}건이 담긴 AS신청서가 준비됐습니다.")
    st.download_button(
        label="⬇️ AS신청서 다운로드 (.docx)",
        data=docx_buf,
        file_name="AS신청서_미완료.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

st.markdown("---")

# --- 공간/상태 필터 및 리스트 ---
with st.expander("🗺️ 공간 위치 확인 (평면도)"):
    if os.path.exists("image_5012c2.jpg"):
        st.image("image_5012c2.jpg", use_container_width=True)
    else:
        st.info("평면도 이미지(image_5012c2.jpg)를 찾을 수 없습니다.")

filter_col1, filter_col2 = st.columns(2)
with filter_col1:
    space = st.selectbox("공간별 필터링", ["전체"] + list(df['공간'].unique()))
with filter_col2:
    status_filter = st.selectbox("진행상태 필터링", ["전체", "완료", "미완료"])

target_df = df if space == "전체" else df[df['공간'] == space]
if status_filter == "완료":
    target_df = target_df[target_df['진행현황'].astype(str).str.strip() == '완료']
elif status_filter == "미완료":
    target_df = target_df[target_df['진행현황'].astype(str).str.strip() != '완료']

cols = st.columns(2)
for i, (index, row) in enumerate(target_df.iterrows()):
    with cols[i % 2]:
        status = str(row.get('진행현황', '미지정')).replace('nan', '미지정')
        status_color = "#27ae60" if status == '완료' else "#e74c3c"

        st.markdown(f"""
            <div class='card' style='border-left-color: {status_color};'>
                <h3>[{row['번호']}] {row['공간']} - {row['부위']}</h3>
                <p><b>상태:</b> {status}</p>
                <p><b>상세:</b> {row.get('유형', '')} / {row.get('상세내용', '')}</p>
            </div>
        """, unsafe_allow_html=True)

        file_name = str(row.get('저장된사진파일명', '')).strip()
        if os.path.exists(file_name):
            st.image(file_name, use_container_width=True)
