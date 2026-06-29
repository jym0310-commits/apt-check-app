import streamlit as st
import pandas as pd
import os

# 페이지 설정
st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# 해링턴 플레이스 테마 적용 (CSS)
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stApp { max-width: 1200px; margin: 0 auto; }
    .card { background: white; padding: 20px; border-radius: 12px; border: 1px solid #e0e0e0; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .title-box { background-color: #1B2845; padding: 20px; border-radius: 10px; color: white; text-align: center; margin-bottom: 20px; }
    h3 { color: #1B2845; }
    .highlight { color: #C5A059; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# 헤더
st.markdown("<div class='title-box'><h1>🏢 해링턴 플레이스 사전점검 리스트</h1></div>", unsafe_allow_html=True)

# 1. 평면도 섹션
with st.expander("📍 위치 확인하기 (평면도 보기)", expanded=True):
    # 'image_5012c2.jpg' 파일이 깃허브 루트에 있다고 가정
    if os.path.exists("image_5012c2.jpg"):
        st.image("image_5012c2.jpg", caption="세대 평면도", use_container_width=True)
    else:
        st.error("평면도 파일(image_5012c2.jpg)을 깃허브에 업로드해주세요.")

# 데이터 로드
@st.cache_data
def load_data():
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

df = load_data()
df.columns = [str(c).strip() for c in df.columns]
df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

# 통계 요약
m1, m2, m3 = st.columns(3)
m1.metric("전체 하자 건수", f"{len(df)}건")
m2.metric("점검 공간", len(df['공간'].unique()))
m3.metric("최다 유형", df['유형'].mode()[0])

st.markdown("---")

# 공간 필터링
space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

# 리스트 출력
cols = st.columns(2)
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
                st.warning(f"이미지 없음: {file_name}")
            st.markdown("</div>", unsafe_allow_html=True)
