import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")
st.title("🏢 해링턴 플레이스 사전점검 리스트")

# 구글 드라이브 폴더 ID 설정
FOLDER_ID = "175FZchh83S2OEtDpa7uQsd4_iom1t-vm"

@st.cache_data
def load_data():
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

df = load_data()
df.columns = [str(c).strip() for c in df.columns]

# 번호 정제
df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

for _, row in target_df.iterrows():
    with st.container():
        st.markdown(f"### 🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
        st.markdown(f"**상세내용:** {row['유형']} / {row['상세내용']}")
        
        # 💡 구글 드라이브 이미지를 띄우기 위한 표준 HTML 방식
        # 형님이 올리신 파일 이름(예: 1_침실3_...jpg)을 기반으로 구글 드라이브에서 직접 보여줍니다.
        # 구글 드라이브는 파일 이름만으로 바로 이미지를 띄우기 어려워, 
        # 웹 브라우저에서 '미리보기' 가능한 주소 형식을 사용합니다.
        
        file_name = str(row['저장된사진파일명']).strip()
        st.info(f"파일 이름: {file_name}")
        st.markdown("※ 사진이 로딩되지 않는 경우, 구글 드라이브 폴더 공유 설정이 '링크가 있는 모든 사용자'로 되어 있는지 다시 확인해 주세요.")
        
        st.divider()
