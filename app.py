import streamlit as st
import pandas as pd
import os

# 페이지 설정
st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# 스타일 설정
st.markdown("""
    <style>
    .card { background: #ffffff; padding: 20px; border-radius: 15px; border-left: 10px solid #1B2845; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
    .title-box { background-color: #1B2845; padding: 20px; border-radius: 10px; color: white; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<div class='title-box'><h1>🏢 해링턴 플레이스 하자 관리 대시보드</h1></div>", unsafe_allow_html=True)

@st.cache_data
def load_data():
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

df = load_data()
df.columns = [str(c).strip() for c in df.columns]
df = df[pd.to_numeric(df['번호'], errors='coerce').notnull()]

# --- 요약 대시보드 (오류 방지 코드 포함) ---
st.subheader("📊 전체 하자 요약 현황")
# '진행현황' 열이 있으면 통계 내고, 없으면 0으로 처리
if '진행현황' in df.columns:
    status_counts = df['진행현황'].value_counts()
    done = status_counts.get('완료', 0)
    todo = status_counts.get('미조치', 0)
else:
    done, todo = 0, len(df)

m1, m2, m3 = st.columns(3)
m1.metric("총 하자 건수", f"{len(df)}건")
m2.metric("완료 건수", f"{done}건")
m3.metric("미조치 건수", f"{todo}건")

# --- 평면도 섹션 ---
with st.expander("🗺️ 공간 위치 확인 (평면도)"):
    if os.path.exists("image_5012c2.jpg"):
        st.image("image_5012c2.jpg", caption="세대 평면도", use_container_width=True)
    else:
        st.warning("평면도 이미지(image_5012c2.jpg)를 깃허브에 업로드해주세요.")

# --- 하자 리스트 출력 ---
st.markdown("---")
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
                <p><b>상세:</b> {row['유형']} / {row['상세내용']}</p>
            </div>
        """, unsafe_allow_html=True)
        
        file_name = str(row['저장된사진파일명']).strip()
        if os.path.exists(file_name):
            st.image(file_name, use_container_width=True)
        else:
            st.info("📷 사진 준비 중")
