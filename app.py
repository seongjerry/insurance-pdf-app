import streamlit as st
import fitz  # PyMuPDF
import io
import os
import json

# ==========================================
# ⚙️ 관리자 세팅 구역 (서버 고정값)
# ==========================================
FONT_DIR = "fonts"
os.makedirs(FONT_DIR, exist_ok=True)
DEFAULT_FONT_NAME = "nanum"
DEFAULT_FONT_PATH = os.path.join(FONT_DIR, "NanumGothic.ttf")

# 서버에 미리 준비해둘 백지 동의서 원본 파일명
TEMPLATE_PDF_PATH = "template.pdf"

# 서버에 미리 맞춰둔 고정 좌표값 (나중에 양식이 바뀌면 이 숫자만 바꾸면 됩니다)
DEFAULT_COORDINATES = {
    "보험사명": {"x": 120, "y": 187},
    "상품명": {"x": 248, "y": 187, "fontsize": 9},
    "증권번호": {"x": 455, "y": 187},
    "신청일자_년": {"x": 225, "y": 632},
    "신청일자_월": {"x": 270, "y": 632},
    "신청일자_일": {"x": 330, "y": 632},
    "계약자명": {"x": 160, "y": 673},
    "피보험자명(선택)": {"x": 160, "y": 718}
}
# ==========================================

def get_font():
    """한글 폰트 로드 (없으면 윈도우 기본 폰트로 대체)"""
    if os.path.exists(DEFAULT_FONT_PATH):
        return DEFAULT_FONT_PATH
    win_font = "C:\\Windows\\Fonts\\malgun.ttf"
    if os.path.exists(win_font):
        return win_font
    return None

def create_dummy_template():
    """사용자가 아직 진짜 template.pdf를 넣지 않았을 때 에러 방지용으로 임시 파일을 만듭니다."""
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text(fitz.Point(50, 50), "[시스템 안내]", fontsize=20, color=(1,0,0))
    page.insert_text(fitz.Point(50, 80), "여기에 사용할 실제 동의서 PDF 파일의 이름을", fontsize=15)
    page.insert_text(fitz.Point(50, 100), "'template.pdf' 로 바꾸어 이 폴더에 넣어주세요.", fontsize=15)
    doc.save(TEMPLATE_PDF_PATH)
    doc.close()

def process_pdf(data_mapping):
    """
    미리 서버에 저장된 template.pdf를 열어서 데이터를 기입하고 바이트 형태로 반환합니다.
    """
    # 원본 파일이 없으면 안내용 임시 백지 PDF를 하나 생성합니다.
    if not os.path.exists(TEMPLATE_PDF_PATH):
        create_dummy_template()
        
    # 매번 업로드 받는 것이 아니라, 로컬에 깔린 template.pdf를 읽습니다.
    doc = fitz.open(TEMPLATE_PDF_PATH)
    font_path = get_font()
    
    for i in range(len(doc)):
        page = doc[i]
        
        if font_path:
            page.insert_font(fontname=DEFAULT_FONT_NAME, fontfile=font_path)
            
        for key, info in data_mapping.items():
            text = info.get('value', '')
            x = info.get('x', 0)
            y = info.get('y', 0)
            f_size = info.get('fontsize', 11)  # 개별적으로 폰트 사이즈가 지정되어 있으면 적용
            
            if not text:
                continue
                
            if font_path:
                page.insert_text(fitz.Point(x, y), str(text), fontname=DEFAULT_FONT_NAME, fontsize=f_size, color=(0,0,0))
            else:
                page.insert_text(fitz.Point(x, y), str(text), fontsize=f_size, color=(0,0,0))
                
    output_stream = io.BytesIO()
    doc.save(output_stream)
    doc.close()
    
    return output_stream.getvalue()

def main():
    # 화면을 가운데 정렬(centered)로 바꾸어 진짜 웹 서비스처럼 집중도 있게 구성
    st.set_page_config(page_title="실적입력동의서생성기", page_icon="📄", layout="centered")
    
    # 크롬 등 브라우저의 자동 번역 기능이 한국어를 이상하게 오역하는 것을 방지합니다.
    st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)
    
    st.title("📄 실적입력동의서생성기")
    st.markdown("아래 고객 정보를 입력하시면 지정된 양식에 맞추어 즉시 문서를 생성합니다.")
    
    # 폼 영역 (사용자는 여기만 신경 쓰면 됩니다)
    with st.container():
        st.subheader("📝 고객 및 상품 정보")
        col1, col2 = st.columns(2)
        with col1:
            company_name = st.text_input("보험사명", "")
            contractor_name = st.text_input("계약자명", "")
            date_val = st.date_input("신청일자")
        with col2:
            product_name = st.text_input("상품명", "")
            policy_number = st.text_input("증권번호", "")
            insured_name = st.text_input("피보험자명(선택)", "")
            
        submit_btn = st.button("🚀 위 정보로 문서 생성하기", type="primary", use_container_width=True)

    # 좌표 관리 영역 (일반 사용자에겐 숨김 처리. 관리자만 열어볼 수 있도록 토글 형태)
    with st.expander("⚙️ [관리자용] 좌표 미세 조정 (양식이 바뀔 때만 열어서 수정)"):
        st.warning("이곳의 좌표를 수정하면 새로 뽑히는 PDF의 글씨 위치가 바뀝니다.")
        # key값을 부여하여 Streamlit 내부 캐시를 초기화하고 항상 최신 명칭을 불러오도록 강제합니다.
        coords_json = st.text_area("JSON 좌표 매핑", value=json.dumps(DEFAULT_COORDINATES, ensure_ascii=False, indent=2), height=150, key="coords_json_v7")
        try:
            coords = json.loads(coords_json)
        except json.JSONDecodeError:
            st.error("좌표 형식이 올바르지 않습니다.")
            coords = DEFAULT_COORDINATES

    # 생성 로직
    if submit_btn:
        if not company_name or not contractor_name:
            st.error("빈칸이 있습니다. 필수 정보(보험사명, 계약자명)를 입력해주세요!")
            return
            
        data_mapping = {}
        # 각 필드 정보에서 x, y 외에 fontsize 같은 추가 옵션이 있으면 그대로 덮어쓰기 위해 병합
        if "보험사명" in coords: data_mapping["보험사명"] = {**coords["보험사명"], "value": company_name}
        if "상품명" in coords: data_mapping["상품명"] = {**coords["상품명"], "value": product_name}
        if "증권번호" in coords: data_mapping["증권번호"] = {**coords["증권번호"], "value": policy_number}
        if "계약자명" in coords: data_mapping["계약자명"] = {**coords["계약자명"], "value": contractor_name}
        if "피보험자명(선택)" in coords: data_mapping["피보험자명(선택)"] = {**coords["피보험자명(선택)"], "value": insured_name}
        
        # 날짜를 세 개(년, 월, 일)로 쪼개어서 독립적인 좌표로 뿌립니다.
        # 기존 양식에 "20  년" 이 프린트되어 있으므로 "26"과 같이 뒤 2자리만 사용합니다.
        year_str = str(date_val.year)[-2:]
        month_str = f"{date_val.month:02d}"
        day_str = f"{date_val.day:02d}"
        
        if "신청일자_년" in coords: data_mapping["신청일자_년"] = {**coords["신청일자_년"], "value": year_str}
        if "신청일자_월" in coords: data_mapping["신청일자_월"] = {**coords["신청일자_월"], "value": month_str}
        if "신청일자_일" in coords: data_mapping["신청일자_일"] = {**coords["신청일자_일"], "value": day_str}
        
        with st.spinner("선택된 양식에 정보를 합성 중입니다..."):
            try:
                result_pdf_bytes = process_pdf(data_mapping)
                st.success("✅ 파일 처리가 완료되었습니다. 아래 다운로드 버튼을 누르세요.")
                
                # 다운로드 버튼
                st.download_button(
                    label=f"⬇️ [{contractor_name}] 실적입력동의서 다운로드",
                    data=result_pdf_bytes,
                    file_name=f"{contractor_name}_실적입력동의서.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
