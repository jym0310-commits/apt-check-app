import streamlit as st
import pandas as pd
import os

st.set_page_config(layout="wide")
st.title("🏢 해링턴 플레이스 사전점검 리스트")

EXCEL_PATH = "해링턴_사전점검_사진포함_최종.xlsx"

@st.cache_data
def load_data(mtime):
    return pd.read_excel(EXCEL_PATH, sheet_name='Sheet1')

mtime = os.path.getmtime(EXCEL_PATH)
df = load_data(mtime)

df.columns = [str(c).strip() for c in df.columns]

# 번호가 있는 행만 추출
df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

for _, row in target_df.iterrows():
    with st.container():
        st.markdown(f"### 🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
        st.markdown(f"**상세내용:** {row['유형']} / {row['상세내용']}")
        
        img_filename = str(row['저장된사진파일명']).strip()
        
        if img_filename and os.path.exists(img_filename):
            st.image(img_filename, caption=img_filename, use_container_width=True)
        else:
            st.warning(f"이미지 파일을 찾을 수 없습니다: {img_filename}")
        
        st.divider()
