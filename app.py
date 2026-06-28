import streamlit as st
import pandas as pd
import os

# 페이지 설정
st.set_page_config(layout="wide", page_title="하자 관리 시스템")

st.title("🏢 해링턴 플레이스 사전점검 리스트")

# 데이터 로드
@st.cache_data
def load_data():
    # 엑셀 파일은 깃허브 저장소 루트에 있다고 가정
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

try:
    df = load_data()
    df.columns = [str(c).strip() for c in df.columns]

    # 번호 정제
    df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

    # 공간 필터링
    space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))
    target_df = df if space == "전체" else df[df['공간'] == space]

    # 리스트 출력
    for _, row in target_df.iterrows():
        with st.container():
            st.markdown(f"### 🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
            st.markdown(f"**상세내용:** {row['유형']} / {row['상세내용']}")
            
            # 사진 불러오기 (깃허브의 photos 폴더에 사진들이 있다고 가정)
            file_name = str(row['저장된사진파일명']).strip()
            img_path = f"photos/{file_name}"
            
            if os.path.exists(img_path):
                st.image(img_path, use_container_width=True)
            else:
                st.warning(f"사진을 찾을 수 없습니다: {file_name}")
            
            st.divider()

except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")
