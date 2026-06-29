import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# 구글 시트 연결
@st.cache_data(ttl=10) # 10초마다 데이터 갱신
def get_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    return client.open("하자관리시트이름").sheet1

sheet = get_sheet()
data = sheet.get_all_records()
df = pd.DataFrame(data)

st.title("🏢 해링턴 플레이스 하자 관리 대시보드")

# 상태 변경 로직
space = st.selectbox("공간 선택", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

cols = st.columns(2)
for i, row in target_df.iterrows():
    with cols[i % 2]:
        status_color = "#27ae60" if row['진행현황'] == '완료' else "#e74c3c"
        
        with st.container():
            st.markdown(f"<div class='card' style='border-left: 10px solid {status_color}; padding: 20px; background: white;'>", unsafe_allow_html=True)
            st.write(f"### [{row['번호']}] {row['공간']} - {row['부위']}")
            st.write(f"상태: **{row['진행현황']}**")
            
            # 여기서 바로 수정!
            if row['진행현황'] != '완료':
                if st.button(f"✅ 완료 처리 ({row['번호']})", key=f"btn_{row['번호']}"):
                    # 구글 시트의 특정 행(번호+1)의 '진행현황' 열(예: 4번 열)을 '완료'로 수정
                    sheet.update_cell(i + 2, df.columns.get_loc('진행현황') + 1, '완료')
                    st.rerun() # 화면 새로고침
            
            st.markdown("</div>", unsafe_allow_html=True)
