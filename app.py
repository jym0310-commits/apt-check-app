import streamlit as st
import pandas as pd
import os

# 페이지 설정
st.set_page_config(layout="wide", page_title="하자 관리 시스템")

# 전문가적인 웹 스타일 CSS 적용
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stApp { max-width: 1200px; margin: 0 auto; }
    .card { 
        background: white; padding: 20px; border-radius: 15px; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px;
    }
    h3 { color: #2c3e50; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏢 해링턴 플레이스 사전점검 리스트")
st.markdown("---")

# 데이터 로드
@st.cache_data
def load_data():
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

df = load_data()
df.columns = [str(c).strip() for c in df.columns]
df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

# 레이아웃: 필터를 상단 가로 배치
col1, col2 = st.columns([1, 4])
with col1:
    space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))

target_df = df if space == "전체" else df[df['공간'] == space]

# 전문가적인 그리드 출력 (PC 2열, 모바일 1열 자동 전환)
cols = st.columns(2) # PC에서는 2열로 배치
for i, (index, row) in enumerate(target_df.iterrows()):
    with cols[i % 2]:
        with st.container():
            st.markdown(f"<div class='card'>", unsafe_allow_html=True)
            st.markdown(f"### 🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
            st.markdown(f"**상세내용:** {row['유형']} / {row['상세내용']}")
            
            file_name = str(row['저장된사진파일명']).strip()
            if os.path.exists(file_name):
                st.image(file_name, use_container_width=True)
            else:
                st.error("이미지를 찾을 수 없습니다.")
            st.markdown("</div>", unsafe_allow_html=True)
