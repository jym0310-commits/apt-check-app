import streamlit as st
import pandas as pd
import gspread
import json
from google.oauth2.service_account import Credentials
import os

# 페이지 설정
st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# CSS 스타일링
st.markdown("""
    <style>
    .card { background: #ffffff; padding: 20px; border-radius: 15px; border-left: 10px solid #1B2845; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
    .title-box { background-color: #1B2845; padding: 20px; border-radius: 10px; color: white; text-align: center; }
    .stButton>button { width: 100%; border-radius: 5px; background-color: #1B2845; color: white; }
    </style>
    """, unsafe_allow_html=True)

# 헤더
st.markdown("<div class='title-box'><h1>🏢 해링턴 플레이스 하자 관리 대시보드</h1></div>", unsafe_allow_html=True)

# 1. 구글 시트 연결 (Secrets 사용)
@st.cache_resource
def get_sheet():
    # Secrets에서 json 텍스트를 가져와서 딕셔너리로 변환
    secrets_dict = json.loads(st.secrets["gcp_service_account"]["json"])
    creds = Credentials.from_service_account_info(secrets_dict)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_with_scope = creds.with_scopes(scope)
    client = gspread.authorize(creds_with_scope)
    
    # 요청하신 파일명으로 정확히 연결
    return client.open("해링턴_사전점검_사진포함_최종_260629").sheet1

sheet = get_sheet()

# 2. 데이터 가져오기
data = sheet.get_all_records()
df = pd.DataFrame(data)

# 3. 대시보드 통계
# '진행현황' 열이 있다고 가정
if '진행현황' in df.columns:
    status_counts = df['진행현황'].value_counts()
    done = status_counts.get('완료', 0)
    todo = len(df) - done
else:
    done, todo = 0, len(df)

col1, col2, col3 = st.columns(3)
col1.metric("총 하자 건수", f"{len(df)}건")
col2.metric("완료 건수", f"{done}건")
col3.metric("미조치 건수", f"{todo}건")

# 4. 평면도 섹션
with st.expander("🗺️ 공간 위치 확인 (평면도)"):
    if os.path.exists("image_5012c2.jpg"):
        st.image("image_5012c2.jpg", use_container_width=True)
    else:
        st.warning("평면도 이미지가 없습니다.")

st.markdown("---")

# 5. 하자 리스트 출력 및 수정 버튼
space = st.selectbox("공간별 필터링", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

cols = st.columns(2)
for i, row in target_df.iterrows():
    with cols[i % 2]:
        status = row.get('진행현황', '미지정')
        status_color = "#27ae60" if status == '완료' else "#e74c3c"
        
        st.markdown(f"<div class='card' style='border-left-color: {status_color};'>", unsafe_allow_html=True)
        st.subheader(f"🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
        st.write(f"**상태:** {status} | **내용:** {row['상세내용']}")
        
        # 완료 버튼 로직
        if status != '완료':
            if st.button(f"✅ 완료 처리 ({row['번호']})", key=f"btn_{row['번호']}"):
                # 구글 시트 업데이트 (데이터가 2행부터 시작한다고 가정)
                col_index = df.columns.get_loc('진행현황') + 1
                sheet.update_cell(i + 2, col_index, '완료')
                st.rerun()
        
        # 사진 표시
        file_name = str(row['저장된사진파일명']).strip()
        if os.path.exists(file_name):
            st.image(file_name, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
