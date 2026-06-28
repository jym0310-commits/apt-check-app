import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")
st.title("🏢 해링턴 플레이스 사전점검 리스트")

# 데이터 로드
@st.cache_data
def load_data():
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

df = load_data()
df.columns = [str(c).strip() for c in df.columns]

# 번호 정제 (숫자가 아닌 행 제외)
df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

# 공간 필터링
space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

# 리스트 출력
for _, row in target_df.iterrows():
    with st.container():
        st.markdown(f"### 🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
        st.markdown(f"**상세내용:** {row['유형']} / {row['상세내용']}")
        
        # 구글 드라이브 사진 링크 (파일 이름으로 연결)
        file_name = str(row['저장된사진파일명']).strip()
        
        # 사진 확인 버튼 (클릭 시 새 창에서 구글 드라이브 사진 열림)
        # 175FZchh83S2OEtDpa7uQsd4_iom1t-vm은 형님의 폴더 ID입니다.
        st.markdown(f"[📷 사진 확인하기 (구글 드라이브)] (https://drive.google.com/drive/folders/175FZchh83S2OEtDpa7uQsd4_iom1t-vm)")
        
        st.divider()
