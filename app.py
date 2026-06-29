import streamlit as st
import pandas as pd
import os

st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# CSS: 대시보드 전용 스타일링 추가
st.markdown("""
    <style>
    .summary-card { background: #1B2845; color: white; padding: 20px; border-radius: 15px; text-align: center; margin-bottom: 20px; }
    .metric-value { font-size: 24px; font-weight: bold; }
    .card { background: #ffffff; padding: 20px; border-radius: 15px; border-left: 10px solid #1B2845; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏢 해링턴 플레이스 하자 관리 대시보드")

@st.cache_data(ttl=600)
def load_data():
    return pd.read_excel("해링턴_사전점검_사진포함_최종.xlsx", sheet_name='Sheet1')

df = load_data()
df.columns = [str(c).strip() for c in df.columns]

# --- [전문가적인 대시보드 섹션] ---
total = len(df)
status_counts = df['진행현황'].value_counts() if '진행현황' in df.columns else pd.Series()
done = status_counts.get('완료', 0)
todo = total - done
progress = int((done / total) * 100) if total > 0 else 0

st.markdown(f"""
    <div class='summary-card'>
        <h3>전체 하자 처리 현황</h3>
        <div style='font-size: 40px;'>{done} / {total} 건</div>
        <p>전체 진행률 {progress}%</p>
    </div>
""", unsafe_allow_html=True)

# 시각적 프로그레스 바
st.progress(progress / 100)

# 통계 상세 지표
d1, d2, d3 = st.columns(3)
d1.metric("총 하자 건수", f"{total}건")
d2.metric("완료 건수", f"{done}건", delta=f"{progress}%")
d3.metric("미조치 건수", f"{todo}건", delta_color="inverse")

st.markdown("---")

# 3. 평면도 및 필터
with st.expander("🗺️ 공간 위치 확인 (평면도)"):
    st.image("image_5012c2.jpg", use_container_width=True)

space = st.selectbox("공간별 필터링", ["전체"] + list(df['공간'].unique()))
target_df = df if space == "전체" else df[df['공간'] == space]

# 4. 리스트 출력
cols = st.columns(2)
for i, (index, row) in enumerate(target_df.iterrows()):
    with cols[i % 2]:
        status_color = "#27ae60" if row.get('진행현황') == '완료' else "#e74c3c"
        st.markdown(f"""
            <div class='card' style='border-left-color: {status_color};'>
                <h3>[{row['번호']}] {row['공간']} - {row['부위']}</h3>
                <p><b>상태:</b> {row.get('진행현황', '미지정')}</p>
                <p><b>상세:</b> {row['유형']} / {row['상세내용']}</p>
            </div>
        """, unsafe_allow_html=True)
        
        file_name = str(row['저장된사진파일명']).strip()
        if os.path.exists(file_name):
            st.image(file_name, use_container_width=True)
