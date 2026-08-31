import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import json
import io
from PIL import Image
import datetime

# 고해상도 스캔 이미지 용량 제한 해제
Image.MAX_IMAGE_PIXELS = None

# ==========================================
# 0. 사내 전용 비밀번호 설정
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 사내 전용 시스템")
        st.text_input("접속 비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 사내 전용 시스템")
        st.text_input("접속 비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        st.error("비밀번호가 틀렸습니다.")
        return False
    return True

if not check_password():
    st.stop()

# ==========================================
# 1. 환경 설정 및 API 키 입력
# ==========================================
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)

model = genai.GenerativeModel('gemini-3.6-flash')

# ==========================================
# 2. 핵심 함수 정의
# ==========================================
def extract_info_from_image(image):
    prompt = """
    You are an expert at extracting information from identity documents (passports, Canadian visas, study/work permits).
    Analyze this document and extract the following information. 
    Return ONLY a valid JSON object without markdown tags.
    
    Format required:
    {
      "surname": "Family name or Surname in English uppercase",
      "given_name": "Given names in English uppercase",
      "dob": "Date of birth in YYYY-MM-DD format",
      "uci": "UCI (Unique Client Identifier) if present, formatted as digits only (e.g., 1234567890). If not found, return an empty string."
    }
    """
    try:
        response = model.generate_content([prompt, image])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"정보 추출 중 오류가 발생했습니다: {e}")
        return None

def fill_imm5476(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    
    target_data = {
        "surname": data.get("surname", ""),
        "given": data.get("given_name", ""),
        "dob": data.get("dob", ""),  
        "email": data.get("email", ""),
        "uci": data.get("uci", "").replace("-", ""),  
        "signDate": data.get("signDate", "")  
    }
    
    flags = {
        "surname": False, "given": False, "dob": False, "email": False, "uci": False
    }
    date_counter = 0
    fields_found = False
    
    for page in doc:
        for widget in page.widgets():
            field_name = widget.field_name
            if not field_name: 
                continue
            
            fields_found = True
            fname_lower = field_name.lower()
            
            if "family name" in fname_lower and not flags["surname"]:
                widget.field_value = target_data["surname"]
                widget.update()
                flags["surname"] = True
                
            elif "given name" in fname_lower and not flags["given"]:
                widget.field_value = target_data["given"]
                widget.update()
                flags["given"] = True
                
            elif "date of birth" in fname_lower and not flags["dob"]:
                widget.field_value = target_data["dob"]
                widget.update()
                flags["dob"] = True
                
            elif "email" in fname_lower and not flags["email"]:
                widget.field_value = target_data["email"]
                widget.update()
                flags["email"] = True
                
            elif ("uci" in fname_lower or "unique client identifier" in fname_lower) and not flags["uci"]:
                widget.field_value = target_data["uci"]
                widget.update()
                flags["uci"] = True
                
            elif "date" in fname_lower and "birth" not in fname_lower:
                date_counter += 1
                if date_counter == 1:
                    widget.field_value = target_data["signDate"]
                    widget.update()

    if not fields_found:
        raise Exception("PDF에서 입력 칸을 찾을 수 없습니다. (입력 칸이 살아있는 원본 PDF를 올려주세요.)")

    output_pdf = io.BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    
    return output_pdf

# ==========================================
# 3. Streamlit 웹 UI 구성
# ==========================================
st.set_page_config(page_title="IMM5476 자동 작성 도구", layout="centered")

st.title("🍁 IMM5476 자동 작성 도구 (사내전용)")
st.write("여권이나 퍼밋 서류를 올리면 최신 AI가 데이터를 완벽하게 읽어냅니다.")

if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = None

st.subheader("1. IMM5476 템플릿 업로드")
st.info("💡 지정된 IMM5476 서식을 올려주세요.")
template_file = st.file_uploader("IMM5476 템플릿 PDF 선택", type=['pdf'], key="template")

st.subheader("2. 손님 여권 또는 퍼밋 (이미지 또는 PDF)")
client_file = st.file_uploader("여권/퍼밋 선택 (JPG, PNG, PDF)", type=['jpg', 'jpeg', 'png', 'pdf'], key="client")

if client_file is not None:
    image_to_process = None
    
    if client_file.type == "application/pdf":
        doc = fitz.open(stream=client_file.read(), filetype="pdf")
        page = doc.load_page(0)
        # 메모리 절약을 위해 해상도 렌더링 최적화
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        image_to_process = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    else:
        image_to_process = Image.open(client_file)
    
    # 서버 메모리 초과 방지를 위한 스마트 리사이징 (최대 가로 1800px)
    if image_to_process.width > 1800:
        ratio = 1800 / image_to_process.width
        new_size = (1800, int(image_to_process.height * ratio))
        image_to_process = image_to_process.resize(new_size)

    st.image(image_to_process, caption="업로드 문서 미리보기", use_container_width=True)

    if st.button("🚀 AI로 정보 자동 추출하기"):
        with st.spinner("AI가 문서를 꼼꼼히 읽고 있습니다... (약 3~5초 소요)"):
            extracted = extract_info_from_image(image_to_process)
            if extracted:
                st.session_state.extracted_data = extracted
                st.success("정보 추출 성공!")

if st.session_state.extracted_data is not None:
    st.subheader("3. 정보 확인 및 추가 입력")
    data = st.session_state.extracted_data
    
    col1, col2 = st.columns(2)
    with col1:
        surname = st.text_input("성 (Surname)", value=data.get("surname", ""))
        dob = st.text_input("생년월일 (YYYY-MM-DD)", value=data.get("dob", ""))
        email = st.text_input("이메일", value="")
    with col2:
        given = st.text_input("이름 (Given Name)", value=data.get("given_name", ""))
        uci = st.text_input("UCI (있는 경우)", value=data.get("uci", ""))
        sign_date = st.date_input("서명란 날짜", value=datetime.date.today())

    st.subheader("4. 최종 PDF 생성")
    if st.button("문서 생성 및 다운로드", type="primary"):
        if template_file is None:
            st.error("먼저 1번 단계에서 템플릿 PDF를 업로드해주세요!")
        elif not email:
            st.warning("이메일을 입력해주세요.")
        else:
            final_data = {
                "surname": surname,
                "given_name": given,
                "dob": dob,
                "uci": uci,
                "email": email,
                "signDate": sign_date.strftime("%Y-%m-%d")
            }
            
            try:
                filled_pdf = fill_imm5476(template_file.getvalue(), final_data)
                st.download_button(
                    label="📥 완성된 IMM5476 다운로드",
                    data=filled_pdf,
                    file_name=f"IMM5476_{surname}_{given}.pdf",
                    mime="application/pdf"
                )
                st.balloons()
            except Exception as e:
                st.error(f"PDF 생성 중 오류가 발생했습니다: {e}")
