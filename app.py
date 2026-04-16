import io
import os
import json
import datetime
import fitz
import streamlit as st
import pandas as pd
import zipfile
from streamlit_drawable_canvas import st_canvas
from PIL import Image

# ==========================================
# ⚙️ 시스템 기준 폴더 및 세팅 구역
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_FONT_PATH = os.path.join(BASE_DIR, "fonts", "NanumGothic.ttf")
DEFAULT_FONT_NAME = "nanum"

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
            "피보험자명(선택)": {"x": 160, "y": 718},
            "서명_계약자": {"x": 210, "y": 655, "type": "image", "w": 60, "h": 25},
            "서명_피보험자": {"x": 210, "y": 700, "type": "image", "w": 60, "h": 25}
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
            "설계사명": {"x": 360, "y": 813},
            "서명_계약자": {"x": 470, "y": 660, "type": "image", "w": 60, "h": 25},
            "서명_피보험자": {"x": 470, "y": 685, "type": "image", "w": 60, "h": 25},
            "서명_모집자": {"x": 470, "y": 796, "type": "image", "w": 60, "h": 25}
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
            "신청일자_일": {"x": 250, "y": 805},
            "서명_계약자": {"x": 490, "y": 790, "type": "image", "w": 60, "h": 25}
        }
    }
}

# ==========================================
# 🛠️ 로직 구역
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
    st.session_state["k_company"] = ""
    st.session_state["k_contractor"] = ""
    st.session_state["k_product"] = ""
    st.session_state["k_policy"] = ""
    st.session_state["k_insured"] = ""
    st.session_state["k_agent"] = ""
    st.session_state["k_date"] = datetime.date.today()
    
    keys_to_delete = ["result_pdfs", "generated", "contractor_name"]
    for k in keys_to_delete:
        if k in st.session_state:
            del st.session_state[k]
            
    # 서명 캔버스 데이터 초기화를 위한 랜덤 키 변경 (화면 갱신용)
    st.session_state["canvas_key_suffix"] = str(datetime.datetime.now().timestamp())

def convert_canvas_to_bytes(canvas_result):
    """캔버스 데이터를 배경 투명한 PNG 바이트로 변환"""
    if canvas_result is not None and canvas_result.image_data is not None:
        objects = canvas_result.json_data.get("objects") if canvas_result.json_data else []
        if len(objects) > 0:
            img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
            data = img.getdata()
            new_data = []
            for item in data:
                if item[0] > 240 and item[1] > 240 and item[2] > 240:
                    new_data.append((255, 255, 255, 0))
                else:
                    new_data.append(item)
            img.putdata(new_data)
            b = io.BytesIO()
            img.save(b, format="PNG")
            return b.getvalue()
    return None

def process_selected_pdfs(selected_templates_keys, input_data, templates_config, signatures=None):
    font_path = get_font()
    
    year_str = str(input_data["date_val"].year)[-2:]
    month_str = f'{input_data["date_val"].month:02d}'
    day_str = f'{input_data["date_val"].day:02d}'
    
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

    if signatures:
        data_values.update(signatures)

    result_pdfs = {}
    
    for tmpl_key in selected_templates_keys:
        t_config = templates_config[tmpl_key]
        file_path = os.path.join(BASE_DIR, t_config["file_name"])
        
        if not os.path.exists(file_path):
            st.warning(f"경고: [{tmpl_key}] 원본 PDF 파일({t_config['file_name']})을 찾을 수 없어 건너뜁니다.")
            continue
            
        doc = fitz.open(file_path)
        coords = t_config["coords"]
        
        for i in range(len(doc)):
            page = doc[i]
            if font_path:
                page.insert_font(fontname=DEFAULT_FONT_NAME, fontfile=font_path)
                
            for field_name, val in data_values.items():
                if field_name in coords and val:  
                    cinfo = coords[field_name]
                    if cinfo.get("type") == "image":
                        # val은 투명 배경 처리된 PNG 바이너리
                        x, y = cinfo.get("x", 0), cinfo.get("y", 0)
                        w, h = cinfo.get("w", 60), cinfo.get("h", 25)
                        rect = fitz.Rect(x, y, x+w, y+h)
                        page.insert_image(rect, stream=val)
                    else:
                        x, y = cinfo.get('x', 0), cinfo.get('y', 0)
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

def get_excel_template():
    df = pd.DataFrame(columns=["보험사명", "상품명", "증권번호", "계약자명", "피보험자명(선택)", "모집설계사 성명", "계약체결일자"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def main():
    st.set_page_config(page_title="통합 서류 생성기", page_icon="📄", layout="centered")
    st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)
    
    if "canvas_key_suffix" not in st.session_state:
        st.session_state["canvas_key_suffix"] = "initial"
        
    st.title("📄 통합 서류 생성기")
    st.markdown("단건의 상세 서류 묶음을 제작하거나, 엑셀을 통해 한 번에 수백 장을 찍어낼 수 있습니다.")
    
    tab1, tab2 = st.tabs(["✍️ 단건 생성 모드 (서명 지원)", "📁 엑셀 대량 생성 모드 (ZIP 다운)"])
    
    with tab1:
        st.markdown("입력하신 한 분의 정보를 기반으로 서류들을 생성합니다.")
        with st.container():
            st.subheader("📝 고객 및 상품 정보")
            col1, col2 = st.columns(2)
            with col1:
                company_name = st.text_input("보험사명", key="k_company")
                contractor_name = st.text_input("계약자명", key="k_contractor")
                date_val = st.date_input("계약체결일자", key="k_date")
                
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                st.button("🔄 내용을 모두 비우기 (초기화)", on_click=clear_form, use_container_width=True)

            with col2:
                product_name = st.text_input("상품명", key="k_product")
                policy_number = st.text_input("증권번호", key="k_policy")
                insured_name = st.text_input("피보험자명(선택)", key="k_insured")
                agent_name = st.text_input("모집설계사 성명", key="k_agent")

        st.markdown("---")
        st.subheader("🖋️ 전자 서명 (터치패드 및 마우스 기입)")
        c1, c2, c3 = st.columns(3)
        # 캔버스의 key를 가변으로 두어 초기화 시 캔버스도 완전히 리프레시 되도록 함
        ckey_suffix = st.session_state["canvas_key_suffix"]
        with c1:
            st.write("🧑‍💼 계약자 서명")
            sign_c = st_canvas(stroke_width=2, stroke_color="#000", background_color="#fff", height=60, width=200, drawing_mode="freedraw", key=f"s_c_{ckey_suffix}")
        with c2:
            st.write("🧑 피보험자 서명 (선택)")
            sign_i = st_canvas(stroke_width=2, stroke_color="#000", background_color="#fff", height=60, width=200, drawing_mode="freedraw", key=f"s_i_{ckey_suffix}")
        with c3:
            st.write("💼 모집자 서명 (완전판매용)")
            sign_a = st_canvas(stroke_width=2, stroke_color="#000", background_color="#fff", height=60, width=200, drawing_mode="freedraw", key=f"s_a_{ckey_suffix}")
            
        st.markdown("---")
        st.subheader("🖨️ 생성 및 다운로드할 서류 선택")
        st.markdown("👉 **아래에서 체크한 서류들에 대해서만 [문서 생성] 및 [다운로드 버튼]이 나타납니다.**")
        
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

        submit_btn = st.button("🚀 위 정보로 단건 문서 생성하기", type="primary", use_container_width=True)

        with st.expander("⚙️ [관리자용] 좌표 미세 조정 (다중 템플릿 통합 메뉴)"):
            st.warning("이곳의 좌표를 수정하면 각 문서의 글씨 위치가 바뀝니다. (완전적용하려면 Github 코드 수정 필요)")
            coords_json = st.text_area("JSON 다중 좌표 매핑", value=json.dumps(TEMPLATES_CONFIG, ensure_ascii=False, indent=2), height=300, key="coords_json_v13")
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
                
                sig_dict = {}
                # 변환된 이미지 바이트가 None이 아니면 (즉, 서명을 했으면) 사전에 추가
                b_c = convert_canvas_to_bytes(sign_c)
                if b_c: sig_dict["서명_계약자"] = b_c
                
                b_i = convert_canvas_to_bytes(sign_i)
                if b_i: sig_dict["서명_피보험자"] = b_i
                
                b_a = convert_canvas_to_bytes(sign_a)
                if b_a: sig_dict["서명_모집자"] = b_a
                
                with st.spinner("선택하신 개별 양식들을 각각 작성하고 있습니다..."):
                    try:
                        result_pdfs_dict = process_selected_pdfs(selected_docs, input_data, runtime_config, sig_dict)
                        st.session_state["result_pdfs"] = result_pdfs_dict
                        st.session_state["contractor_name"] = contractor_name
                        st.session_state["generated"] = True
                    except Exception as e:
                        st.error(f"오류가 발생했습니다: {e}")

        if st.session_state.get("generated", False) and "result_pdfs" in st.session_state:
            st.success(f"✅ 파일 처리가 완료되었습니다! (총 {len(st.session_state['result_pdfs'])}장 생성)")
            st.markdown("👇 **다운로드를 눌러도 창이 초기화되지 않습니다. 필요한 파일을 차례대로 모두 다운로드하세요!**")
            
            for doc_name, pdf_bytes in st.session_state["result_pdfs"].items():
                st.download_button(
                    label=f"⬇️ [{st.session_state['contractor_name']}] {doc_name}.pdf 다운로드",
                    data=pdf_bytes,
                    file_name=f"[{st.session_state['contractor_name']}]_{doc_name}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"dl_btn_{doc_name}"
                )

    with tab2:
        st.markdown("⚠️ **안내:** 대량 엑셀 생성 기능 이용 시, 한 번에 수백 장의 PDF가 개별 고객 정보로 작성되어 하나의 압축(.zip) 파일로 제공됩니다. **(단, 전자 서명란은 공란으로 빈칸 출력됩니다)**")
        
        st.download_button(
            label="📥 엑셀 업로드용 기본 양식(템플릿) 다운로드",
            data=get_excel_template(),
            file_name="고객정보_업로드양식.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")
        st.subheader("🖨️ 대량 생성 시 적용할 서류 선택")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            use_m1 = st.checkbox("실적입력동의서", value=True, key="m1")
        with col_m2:
            use_m2 = st.checkbox("완전판매확인서", value=True, key="m2")
        with col_m3:
            use_m3 = st.checkbox("금소법FA고지의무확인서", value=True, key="m3")
            
        selected_m_docs = []
        if use_m1: selected_m_docs.append("실적입력동의서")
        if use_m2: selected_m_docs.append("완전판매확인서")
        if use_m3: selected_m_docs.append("금소법FA고지의무확인서")
        
        uploaded_file = st.file_uploader("📂 데이터가 입력된 엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"])
        
        if uploaded_file and len(selected_m_docs) > 0:
            if st.button("🚀 업로드한 엑셀 데이터로 전체 PDF 일괄 생성 (ZIP 저장)", type="primary"):
                with st.spinner("대량 데이터를 분석하여 PDF를 생성하고 있습니다. (잠시만 기다려주세요...)"):
                    try:
                        df = pd.read_excel(uploaded_file)
                        df = df.fillna("")
                        
                        pdf_dict_list = []
                        for idx, row in df.iterrows():
                            date_val_raw = row.get("계약체결일자", "")
                            if pd.isna(date_val_raw) or date_val_raw == "":
                                date_val_p = datetime.date.today()
                            elif isinstance(date_val_raw, pd.Timestamp) or isinstance(date_val_raw, datetime.datetime):
                                date_val_p = date_val_raw.date()
                            else:
                                try:
                                    date_val_p = pd.to_datetime(date_val_raw).date()
                                except Exception:
                                    date_val_p = datetime.date.today()
                                    
                            input_data_p = {
                                "company_name": str(row.get("보험사명", "")),
                                "product_name": str(row.get("상품명", "")),
                                "policy_number": str(row.get("증권번호", "")),
                                "contractor_name": str(row.get("계약자명", "")),
                                "insured_name": str(row.get("피보험자명(선택)", "")),
                                "agent_name": str(row.get("모집설계사 성명", "")),
                                "date_val": date_val_p
                            }
                            
                            # 대량 모드에서는 서명을 None(공란)으로 처리
                            res = process_selected_pdfs(selected_m_docs, input_data_p, TEMPLATES_CONFIG, signatures=None)
                            if res:
                                # 구분하기 쉽게 파일명에 순번과 이름을 붙임
                                mapped_res = {f"{idx+1}_[{input_data_p['contractor_name']}]_{k}.pdf": v for k, v in res.items()}
                                pdf_dict_list.append(mapped_res)
                        
                        if len(pdf_dict_list) > 0:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                for res in pdf_dict_list:
                                    for filename, byte_val in res.items():
                                        zip_file.writestr(filename, byte_val)
                                        
                            st.success(f"✅ 총 {len(df)}명의 고객에 대한 {len(df)*len(selected_m_docs)}장 PDF 생성이 압축 완료되었습니다!")
                            st.download_button(
                                label="📦 일괄 완성된 ZIP 파일 다운로드", 
                                data=zip_buffer.getvalue(), 
                                file_name="대량데이터_자동생성.zip", 
                                mime="application/zip",
                                use_container_width=True
                            )
                    except Exception as e:
                        st.error(f"대량 생성 중 오류가 발생했습니다. (엑셀 형식을 확인해주세요): {e}")

if __name__ == "__main__":
    main()
