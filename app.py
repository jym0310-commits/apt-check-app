import streamlit as st
import pandas as pd
import os
import glob

st.set_page_config(layout="wide")
st.title("🏢 해링턴 플레이스 사전점검 리스트")

# 저장소 안에서 xlsx 파일을 자동으로 찾음 (파일명 오타/인코딩 문제 회피)
excel_candidates = glob.glob("*.xlsx")

if not excel_candidates:
    st.error("저장소 안에서 엑셀(.xlsx) 파일을 찾지 못했습니다. GitHub에 파일이 실제로 존재하는지 확인해 주세요.")
    st.stop()

EXCEL_PATH = excel_candidates[0]
st.caption(f"불러온 파일: {EXCEL_PATH}")  # 확인용, 나중에 지워도 됨

@st.cache_data
def load_data(mtime, path):
    return pd.read_excel(path, sheet_name='Sheet1')

mtime = os.path.getmtime(EXCEL_PATH)
df = load_data(mtime, EXCEL_PATH)

df.columns = [str(c).strip() for c in df.columns]

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
