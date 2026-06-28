import os
import pandas as pd
import streamlit as st
import glob

# 페이지 기본 설정
st.set_page_config(page_title="해링턴 플레이스 사전점검 리스트", layout="wide")

# 경로 설정
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "apt_check")
excel_path = os.path.join(desktop_path, "해링턴_사전점검_사진포함_최종.xlsx")

st.title("🏢 해링턴 플레이스 테크노폴리스 사전점검")
st.subheader("실시간 하자 보수 신청 내역 모니터링")
st.markdown("---")

if not os.path.exists(excel_path):
    st.error("지정된 엑셀 파일(해링턴_사전점검_사진포함_최종.xlsx)을 찾을 수 없습니다.")
else:
    try:
        # 🌟 [핵심 수술] 피벗 시트(Sheet2)를 무시하고 진짜 데이터가 있는 'Sheet1'을 강제로 지정해서 읽습니다.
        df = pd.read_excel(excel_path, sheet_name='Sheet1')
        
        # 컬럼명 좌우 공백 정제
        df.columns = [str(c).strip() for c in df.columns]
        
        parsed_data = []
        
        for idx, row in df.iterrows():
            # 번호가 비어있거나 nan이면 패스
            if pd.isna(row['번호']) or str(row['번호']).strip() == 'nan':
                continue
                
            val_no = str(row['번호']).split('.')[0].strip()   # 소수점 제거 (.0)
            val_space = str(row['공간']).strip()
            val_part = str(row['부위']).strip()
            val_type = str(row['유형']).strip()
            val_desc = str(row['상세내용']).strip()
            val_date = str(row['일시']).strip()
            val_img = str(row['저장된사진파일명']).strip()

            # 글자 정제
            if val_type == 'nan': val_type = ""
            if val_desc == 'nan': val_desc = ""
            if val_date == 'nan': val_date = ""
            if val_img == 'nan' or val_img == '사진없음': val_img = ""

            # 메인 노출 상세내용 조합
            # 유형과 상세내용에 둘 다 글자가 있으면 같이 보여주고, 하나만 있으면 있는 것만 보여줍니다.
            if val_type and val_desc:
                full_content = f"[{val_type}] {val_desc}"
            else:
                full_content = val_type if val_type else val_desc

            # 사진 파일명이 진짜 비어있는 경우에만 자동 규칙 보완
            if not val_img:
                safe_text = "".join([c for c in val_type if c.isalnum() or c in "._- "])[:12].strip()
                val_img = f"{val_no}_{val_space}_{val_part}_{safe_text}.jpg"

            parsed_data.append({
                '번호': int(val_no),
                '공간': val_space,
                '부위': val_part,
                '상세내용': full_content,
                '일시': val_date,
                '저장된사진파일명': val_img
            })

        final_df = pd.DataFrame(parsed_data)
        
        if final_df.empty:
            st.warning("Sheet1에서 하자 내역 데이터를 읽어오지 못했습니다.")
        else:
            # 1. 상단 대시보드 통계 (진짜 100건이 넘는 수량이 명확하게 찍힙니다)
            total_count = len(final_df)
            space_counts = final_df['공간'].value_counts()
            
            col1, col2 = st.columns([1, 3])
            with col1:
                st.metric(label="총 발견된 하자 건수", value=f"{total_count} 건")
            with col2:
                st.write("📍 **공간별 현황:** " + ", ".join([f"{k}({v}건)" for k, v in space_counts.items() if k]))
                
            st.markdown("---")
            
            # 2. 공간 선택 필터
            unique_spaces = [s for s in final_df['공간'].unique() if s]
            spaces = ["전체"] + list(unique_spaces)
            selected_space = st.selectbox("🔍 공간별 선택", spaces)
            
            filtered_df = final_df if selected_space == "전체" else final_df[final_df['공간'] == selected_space]
            
            # 이미지 파일 딕셔너리 빌드
            jpg_files = glob.glob(os.path.join(desktop_path, "*.[jJ][pP][gG]")) + glob.glob(os.path.join(desktop_path, "*.[pP][nN][gG]"))
            img_dict = {os.path.basename(f): f for f in jpg_files}

            # 3. 데이터 출력 리스트 UI
            for index, row in filtered_df.iterrows():
                with st.container():
                    c1, c2 = st.columns([2, 1])
                    
                    with c1:
                        st.markdown(f"### 🔴 [{row['번호']}] {row['공간']} - {row['부위']}")
                        st.markdown(f"**• 상세 내용:** {row['상세내용']}")
                        if row['일시']:
                            st.caption(f"📅 등록 일시: {row['일시']}")
                    
                    with c2:
                        matched_img_path = None
                        target_img = row['저장된사진파일명']
                        
                        # 1순위 완전 매칭
                        if target_img in img_dict:
                            matched_img_path = img_dict[target_img]
                        # 2순위 번호 접두사 매칭
                        else:
                            for fname, fpath in img_dict.items():
                                if fname.startswith(f"{row['번호']}_"):
                                    matched_img_path = fpath
                                    break
                        
                        if matched_img_path and os.path.exists(matched_img_path):
                            st.image(matched_img_path, use_container_width=True)
                        else:
                            st.caption("📷 사진 원본 없음")
                            
                st.markdown("<br><hr style='border:1px dashed #eee'><br>", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"데이터를 읽는 도중 오류가 발생했습니다: {e}")