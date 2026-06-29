import streamlit as st
import pandas as pd
import os

st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# CSS: 더 세련된 카드 디자인
st.markdown("""
    <style>
    .metric-row { display: flex; gap: 20px; }
    .card { background: #ffffff; padding: 25px; border-radius: 15px; border-left: 10px solid #1B2845; box-shadow: 0 4px 10px rgba(0,0,0,0.08); margin-bottom: 20px; }
    .stSelectbox { border: 1px solid #1B2845; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏢 해링턴 플레이스 하자 관리 대시보드")

# 1. 데이터 로드 (실시간 반영을 위해 구글 시트 연결 추천)
@st.cache_data(ttl=600) # 10분마다 새로고침
def load_data():
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

df = load_data()
df.columns = [str(c).strip() for c in df.columns]

# 2. 대시보드 요약
col_a, col_b, col_c = st.columns(3)
col_a.metric("총 하자 건수", f"{len(df)}건")
# '진행현황' 열이 있다고 가정 (미조치/조치중/완료)
status_counts = df['진행현황'].value_counts() if '진행현황' in df.columns else None
col_b.metric("완료된 하자", f"{status_counts.get('완료', 0)}건")
col_c.metric("미조치 건수", f"{status_counts.get('미조치', 0)}건")

# 3. 평면도 및 필터
with st.expander("🗺️ 공간 위치 확인 (평면도)"):
    st.image("image_5012c2.jpg", use_container_width=True)

space = st.selectbox("공간별 필터링", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

# 4. 리스트 출력 (상태에 따라 카드 색상 변경)
cols = st.columns(2)
for i, (index, row) in enumerate(target_df.iterrows()):
    with cols[i % 2]:
        status_color = "#27ae60" if row.get('진행현황') == '완료' else "#e74c3c"
        st.markdown(f"""
            <div class='card' style='border-left-color: {status_color};'>
                <h3>[{row['번호']}] {row['공간']} - {row['부위']}</h3>
                <p><b>상태:</b> {row.get('진행현황', '미지정')}</p>
                <p><b>상세:</b> {row['유형']} / {row['상세내용']}</p>
            </div>
        """, unsafe_allow_html=True)
        
        file_name = str(row['저장된사진파일명']).strip()
        if os.path.exists(file_name):
            st.image(file_name, use_container_width=True)
