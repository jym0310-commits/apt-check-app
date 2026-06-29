import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os

# 페이지 설정
st.set_page_config(layout="wide", page_title="하자 관리 시스템")

# 전문가적인 웹 스타일 CSS 적용
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stApp { max-width: 1200px; margin: 0 auto; }
    .card { background: white; padding: 20px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
    .metric-card { background: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #3498db; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    h3 { color: #2c3e50; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏢 해링턴 플레이스 사전점검 리스트")
st.markdown("---")

# 데이터 로드 함수
@st.cache_resource
def get_sheet():
    # 이 부분은 형님의 환경에 맞춰 다시 구성되어야 합니다.
    # 이전 단계에서 성공했던 인증 방식(Secrets 사용)을 유지하세요.
    secrets = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(dict(secrets))
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_with_scope = creds.with_scopes(scope)
    client = gspread.authorize(creds_with_scope)
    return client.open("해링턴_사전점검_사진포함_최종_260629").sheet1

# 데이터 가져오기
sheet = get_sheet()
data = sheet.get_all_records()
df = pd.DataFrame(data)

# 데이터 전처리
df.columns = [str(c).strip() for c in df.columns]
df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

# --- [상단 요약 대시보드] ---
st.subheader("📊 전체 하자 요약 현황")
m1, m2, m3 = st.columns(3)
m1.metric("전체 하자 건수", f"{len(df)}건")
m2.metric("공간 수", f"{len(df['공간'].unique())}곳")
m3.metric("주요 유형", df['유형'].mode()[0])
st.markdown("---")

# 공간 필터링
col1, col2 = st.columns([1, 4])
with col1:
    space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))

target_df = df if space == "전체" else df[df['공간'] == space]

# 그리드 출력
cols = st.columns(2)
for i, (index, row) in enumerate(target_df.iterrows()):
    with cols[i % 2]:
        with st.container():
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader(f"[{row['번호']}] {row['공간']} - {row['부위']}")
            st.write(f"**상세내용:** {row['상세내용']}")
            
            file_name = str(row['저장된사진파일명']).strip()
            if os.path.exists(file_name):
                st.image(file_name, use_container_width=True)
            else:
                st.info(f"이미지 준비 중: {file_name}")
            st.markdown("</div>", unsafe_allow_html=True)
