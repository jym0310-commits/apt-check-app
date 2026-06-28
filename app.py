import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")
st.title("🏢 해링턴 플레이스 사전점검 리스트")

# 구글 드라이브 폴더의 고유 ID (형님의 공유 폴더 주소에서 가져와야 함)
# 예: https://drive.google.com/drive/folders/1ABC-123DEF... -> 1ABC-123DEF... 부분을 아래에 복사
GOOGLE_DRIVE_FOLDER_ID = "여기에_폴더_ID를_붙여넣으세요"

@st.cache_data
def load_data():
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

df = load_data()
df.columns = [str(c).strip() for c in df.columns]

# 번호가 있는 행만 추출
df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

for _, row in target_df.iterrows():
    with st.container():
        st.markdown(f"### 🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
        st.markdown(f"**상세내용:** {row['유형']} / {row['상세내용']}")
        
        # 구글 드라이브 링크 조합 (직접 이미지 경로 활용)
        # ※ 구글 드라이브 이미지를 웹에 띄우는 표준 방식 사용
        img_id = row['저장된사진파일명']
        st.image(f"https://lh3.googleusercontent.com/d/{img_id}", caption=img_id, use_container_width=True)
        
        st.divider()
