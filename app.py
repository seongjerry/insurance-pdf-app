import streamlit as st
import fitz  # PyMuPDF
import io
import os
import json
import datetime

# ==========================================
# ⚙️ 관리자 세팅 구역 (서버 고정값)
# ==========================================
FONT_DIR = "fonts"
os.makedirs(FONT_DIR, exist_ok=True)
DEFAULT_FONT_NAME = "nanum"
DEFAULT_FONT_PATH = os.path.join(FONT_DIR, "NanumGothic.ttf")

TEMPLATES_CONFIG = {
    "실적입력동의서": {
        "file_name": "template.pdf",
        "coords": {
            "보험사명": {"x": 140, "y": 188},
            "상품명": {"x": 273, "y": 188, "fontsize": 9},
            "증권번호": {"x": 480, "y": 188},
            "신청일자_년": {"x": 230, "y": 632},
            "신청일자_월": {"x": 285, "y": 632},
            "신청일자_일": {"x": 330, "y": 632},
            "계약자명": {"x": 160, "y": 673},
            "피보험자명(선택)": {"x": 160, "y": 718}
        }
    },
    "완전판매확인서": {
        "file_name": "template2.pdf",
        "coords": {
            "상품명": {"x": 134, "y": 622, "fontsize": 9},
            "증권번호": {"x": 465, "y": 77},
            "신청일자_년": {"x": 380, "y": 650},
            "신청일자_월": {"x": 445, "y": 650},
            "신청일자_일": {"x": 490, "y": 650},
            "계약자명": {"x": 360, "y": 677},
            "피보험자명(선택)": {"x": 365, "y": 701},
            "신청일자_년2": {"x": 380, "y": 787},
            "신청일자_월2": {"x": 445, "y": 787},
            "신청일자_일2": {"x": 490, "y": 787},
            "설계사명": {"x": 360, "y": 813}
        }
    },
    "금소법FA고지의무확인서": {
        "file_name": "template3.pdf",
        "coords": {
            "증권번호": {"x": 390, "y": 41},
            "계약자명": {"x": 415, "y": 744},
            "계약자명2": {"x": 415, "y": 805},
            "설계사명": {"x": 105, "y": 744},
            "신청일자_년": {"x": 170, "y": 805},
            "신청일자_월": {"x": 215, "y": 805},
            "신청일자_일": {"x": 250, "y": 805}
        }
    }
}
# ==========================================

def get_font():
    if os.path.exists(DEFAULT_FONT_PATH):
        return DEFAULT_FONT_PATH
    win_font = "C:\\Windows\\Fonts\\malgun.ttf"
    if os.path.exists(win_font):
        return win_font
    return None

def clear_form():
    """모든 입력된 값과 생성된 세션을 초기화하는 콜백 함수"""
    # 텍스트 위젯은 빈 문자열로, 날짜는 오늘 날짜로 덮어씌워서 즉시 초기화
    st.session_state["k_company"] = ""
    st.session_state["k_contractor"] = ""
    st.session_state["k_product"] = ""
    st.session_state["k_policy"] = ""
    st.session_state["k_insured"] = ""
    st.session_state["k_agent"] = ""
    st.session_state["k_date"] = datetime.date.today()
    
    # 다운로드 버튼 및 결과물 삭제
    keys_to_delete = ["result_pdfs", "generated", "contractor_name"]
    for k in keys_to_delete:
        if k in st.session_state:
            del st.session_state[k]

def process_selected_pdfs(selected_templates_keys, input_data, templates_config):
    """
    선택된 여러 템플릿에 데이터를 기입하고, 개별 PDF 바이트 데이터들을 Dictionary 형태로 반환합니다.
    """
    font_path = get_font()
    
    # 쪼개진 년/월/일 포맷팅
    year_str = str(input_data["date_val"].year)[-2:]
    month_str = f'{input_data["date_val"].month:02d}'
    day_str = f'{input_data["date_val"].day:02d}'
    
    # 공통 입력값 사전 (매핑의 편의를 위함)
    data_values = {
        "보험사명": input_data.get("company_name", ""),
        "상품명": input_data.get("product_name", ""),
        "증권번호": input_data.get("policy_number", ""),
        "계약자명": input_data.get("contractor_name", ""),
        "계약자명2": input_data.get("contractor_name", ""),
        "피보험자명(선택)": input_data.get("insured_name", ""),
        "설계사명": input_data.get("agent_name", ""),
        "신청일자_년": year_str,
        "신청일자_월": month_str,
        "신청일자_일": day_str,
        "신청일자_년2": year_str,
        "신청일자_월2": month_str,
        "신청일자_일2": day_str,
        "신청일자_통합": f"20{year_str}년 {month_str}월 {day_str}일"
    }

    result_pdfs = {}
    
    for tmpl_key in selected_templates_keys:
        t_config = templates_config[tmpl_key]
        file_path = t_config["file_name"]
        
        # 파일이 로컬에 없으면 스킵
        if not os.path.exists(file_path):
            st.warning(f"경고: [{tmpl_key}] 원본 PDF 파일({file_path})을 찾을 수 없어 건너뜁니다.")
            continue
            
        doc = fitz.open(file_path)
        coords = t_config["coords"]
        
        for i in range(len(doc)):
            page = doc[i]
            if font_path:
                page.insert_font(fontname=DEFAULT_FONT_NAME, fontfile=font_path)
                
            for field_name, val in data_values.items():
                if field_name in coords and val:  # 값이 비어있지 않고 좌표가 등록되어 있으면
                    cinfo = coords[field_name]
                    x = cinfo.get('x', 0)
                    y = cinfo.get('y', 0)
                    f_size = cinfo.get('fontsize', 11)
                    
                    if font_path:
                        page.insert_text(fitz.Point(x, y), str(val), fontname=DEFAULT_FONT_NAME, fontsize=f_size, color=(0,0,0))
                    else:
                        page.insert_text(fitz.Point(x, y), str(val), fontsize=f_size, color=(0,0,0))
                        
        output_stream = io.BytesIO()
        doc.save(output_stream)
        doc.close()
        
        result_pdfs[tmpl_key] = output_stream.getvalue()
        
    if not result_pdfs:
        raise Exception("출력할 수 있는 문서가 하나도 없습니다. 템플릿 파일을 확인해주세요.")
    
    return result_pdfs

def main():
    st.set_page_config(page_title="통합 서류 생성기", page_icon="📄", layout="centered")
    st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)
    
    st.title("📄 통합 서류 생성기")
    st.markdown("입력하신 한 번의 정보로 필요한 서류들을 모두 작성하여 하나로 묶어드립니다.")
    
    with st.container():
        st.subheader("📝 고객 및 상품 정보")
        col1, col2 = st.columns(2)
        with col1:
            company_name = st.text_input("보험사명", key="k_company")
            contractor_name = st.text_input("계약자명", key="k_contractor")
            date_val = st.date_input("계약체결일자", key="k_date")
            
            # 초기화 버튼을 계약체결일자 아래(왼쪽 열)에 배치하고 우측 입력칸 라벨 높이와 맞춤
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            st.button("🔄 초기화", on_click=clear_form, use_container_width=True)

        with col2:
            product_name = st.text_input("상품명", key="k_product")
            policy_number = st.text_input("증권번호", key="k_policy")
            insured_name = st.text_input("피보험자명(선택)", key="k_insured")
            agent_name = st.text_input("모집설계사 성명", key="k_agent")
            
    st.markdown("---")
    st.subheader("🖨️ 생성 및 다운로드할 서류 선택")
    st.markdown("👉 **아래에서 체크한 서류들에 대해서만 [문서 생성] 및 [다운로드 버튼]이 나타납니다.**")
    
    # 📌 세 개의 체크박스를 깔끔하게 배치
    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        use_doc1 = st.checkbox("실적입력동의서", value=True)
    with col_c2:
        use_doc2 = st.checkbox("완전판매확인서", value=True)
    with col_c3:
        use_doc3 = st.checkbox("금소법FA고지의무확인서", value=True)
        
    selected_docs = []
    if use_doc1: selected_docs.append("실적입력동의서")
    if use_doc2: selected_docs.append("완전판매확인서")
    if use_doc3: selected_docs.append("금소법FA고지의무확인서")

    submit_btn = st.button("🚀 위 정보로 문서 생성하기", type="primary", use_container_width=True)

    with st.expander("⚙️ [관리자용] 좌표 미세 조정 (다중 템플릿 통합 메뉴)"):
        st.warning("이곳의 좌표를 수정하면 각 문서의 글씨 위치가 바뀝니다. (완전적용하려면 Github 코드 수정 필요)")
        coords_json = st.text_area("JSON 다중 좌표 매핑", value=json.dumps(TEMPLATES_CONFIG, ensure_ascii=False, indent=2), height=300, key="coords_json_v11")
        try:
            runtime_config = json.loads(coords_json)
        except json.JSONDecodeError:
            st.error("좌표 형식이 올바르지 않습니다.")
            runtime_config = TEMPLATES_CONFIG

    if submit_btn:
        if not company_name or not contractor_name:
            st.error("빈칸이 있습니다. 필수 정보(보험사명, 계약자명)를 입력해주세요!")
        elif len(selected_docs) == 0:
            st.error("최소 1장 이상의 서류를 체크해 주세요!")
        else:
            input_data = {
                "company_name": company_name,
                "contractor_name": contractor_name,
                "date_val": date_val,
                "product_name": product_name,
                "policy_number": policy_number,
                "insured_name": insured_name,
                "agent_name": agent_name
            }
            
            with st.spinner("선택하신 개별 양식들을 각각 작성하고 있습니다..."):
                try:
                    # PDF 생성 후 세션 스테이트(서버 메모리)에 저장하여 날아가지 않게 보존
                    result_pdfs_dict = process_selected_pdfs(selected_docs, input_data, runtime_config)
                    st.session_state["result_pdfs"] = result_pdfs_dict
                    st.session_state["contractor_name"] = contractor_name
                    st.session_state["generated"] = True
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")

    # 세션 스테이트에 파일이 존재하면 언제나 다운로드 버튼들을 화면에 유지시킴 (Streamlit 리프레쉬 방어)
    if st.session_state.get("generated", False) and "result_pdfs" in st.session_state:
        st.success(f"✅ 파일 처리가 완료되었습니다! (총 {len(st.session_state['result_pdfs'])}장 생성)")
        st.markdown("👇 **다운로드를 눌러도 창이 초기화되지 않습니다. 필요한 파일을 차례대로 모두 다운로드하세요!**")
        
        # 반복문으로 여러 장의 개별 다운로드 버튼을 렌더링
        for doc_name, pdf_bytes in st.session_state["result_pdfs"].items():
            st.download_button(
                label=f"⬇️ [{st.session_state['contractor_name']}] {doc_name}.pdf 다운로드",
                data=pdf_bytes,
                file_name=f"[{st.session_state['contractor_name']}]_{doc_name}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"dl_btn_{doc_name}" # 필수: 여러 버튼이 충돌하지 않도록 고유 키 부여
            )

if __name__ == "__main__":
    main()
