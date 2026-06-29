import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# CSS 스타일링
st.markdown("""
    <style>
    .card { background: #ffffff; padding: 20px; border-radius: 15px; border-left: 10px solid #1B2845; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
    .title-box { background-color: #1B2845; padding: 20px; border-radius: 10px; color: white; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<div class='title-box'><h1>🏢 해링턴 플레이스 하자 관리 대시보드</h1></div>", unsafe_allow_html=True)

# 구글 시트 연결
@st.cache_data(ttl=60) # 1분마다 자동 새로고침
def load_data():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open("하자관리시트이름").sheet1 # 여기서 '하자관리시트이름'을 형님 시트 이름으로 바꾸세요!
    data = sheet.get_all_records()
    return pd.DataFrame(data)

df = load_data()

# 대시보드 요약
status_counts = df['진행현황'].value_counts()
done = status_counts.get('완료', 0)
todo = len(df) - done

col1, col2, col3 = st.columns(3)
col1.metric("총 하자", f"{len(df)}건")
col2.metric("완료", f"{done}건")
col3.metric("미조치", f"{todo}건")

st.markdown("---")

# 리스트 출력
space = st.selectbox("공간별 필터링", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

cols = st.columns(2)
for i, (index, row) in enumerate(target_df.iterrows()):
    with cols[i % 2]:
        status = row.get('진행현황', '미지정')
        status_color = "#27ae60" if status == '완료' else "#e74c3c"
        
        st.markdown(f"""
            <div class='card' style='border-left-color: {status_color};'>
                <h3>🔴 [{row['번호']}] {row['공간']} - {row['부위']}</h3>
                <p><b>상태:</b> {status}</p>
            </div>
        """, unsafe_allow_html=True)
        
        file_name = str(row['저장된사진파일명']).strip()
        if os.path.exists(file_name):
            st.image(file_name, use_container_width=True)
