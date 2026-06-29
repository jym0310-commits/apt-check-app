import streamlit as st
import pandas as pd
import os

# 페이지 설정
st.set_page_config(layout="wide", page_title="해링턴 하자 관리 시스템")

# CSS 스타일링
st.markdown("""
    <style>
    .card { background: #ffffff; padding: 20px; border-radius: 15px; border-left: 10px solid #1B2845; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }
    .title-box { background-color: #1B2845; padding: 20px; border-radius: 10px; color: white; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<div class='title-box'><h1>🏢 해링턴 플레이스 하자 관리 대시보드</h1></div>", unsafe_allow_html=True)

# 1. 데이터 파일 업로드 기능 추가
st.sidebar.header("데이터 업로드")
uploaded_file = st.sidebar.file_uploader("하자 리스트 엑셀 파일을 업로드하세요", type=["xlsx", "csv"])

if uploaded_file is not None:
    # 업로드된 파일 읽기
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    st.success("파일이 성공적으로 로드되었습니다!")

    # 2. 대시보드 통계
    if '진행현황' in df.columns:
        status_counts = df['진행현황'].value_counts()
        done = status_counts.get('완료', 0)
        todo = len(df) - done
    else:
        done, todo = 0, len(df)

    col1, col2, col3 = st.columns(3)
    col1.metric("총 하자 건수", f"{len(df)}건")
    col2.metric("완료 건수", f"{done}건")
    col3.metric("미조치 건수", f"{todo}건")

    # 3. 공간 필터링
    space = st.selectbox("공간별 필터링", ["전체"] + list(df['공간'].unique()))
    target_df = df if space == "전체" else df[df['공간'] == space]

    # 4. 하자 리스트 출력
    cols = st.columns(2)
    for i, row in target_df.iterrows():
        with cols[i % 2]:
            status = row.get('진행현황', '미지정')
            status_color = "#27ae60" if status == '완료' else "#e74c3c"
            
            st.markdown(f"<div class='card' style='border-left-color: {status_color};'>", unsafe_allow_html=True)
            st.subheader(f"🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
            st.write(f"**상태:** {status} | **내용:** {row['상세내용']}")
            
            # 사진 표시 (파일명이 데이터에 포함되어 있다고 가정)
            file_name = str(row.get('저장된사진파일명', '')).strip()
            if file_name and os.path.exists(file_name):
                st.image(file_name, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
else:
    st.info("왼쪽 사이드바에서 하자 리스트 엑셀 파일을 업로드해 주세요.")
