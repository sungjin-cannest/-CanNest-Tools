import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import json
import io
import os
from PIL import Image
import datetime

# 고해상도 스캔 이미지 용량 제한 해제
Image.MAX_IMAGE_PIXELS = None

# ==========================================
# 0. Secrets 안전 검사 및 보안 비밀번호 설정
# ==========================================
if "APP_PASSWORD" not in st.secrets or "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Cloud의 Secrets 설정이 필요합니다.")
    st.info("우측 하단 [Manage app] -> [Settings] -> [Secrets]에 GEMINI_API_KEY와 APP_PASSWORD를 입력해 주세요.")
    st.stop()

def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 CanNest 통합 업무 시스템")
        st.text_input("접속 비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 CanNest 통합 업무 시스템")
        st.text_input("접속 비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        st.error("비밀번호가 틀렸습니다.")
        return False
    return True

if not check_password():
    st.stop()

# ==========================================
# 1. API 키 및 모델 설정
# ==========================================
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-1.5-flash') # 가장 빠르고 최신인 1.5-flash 적용

# ==========================================
# 2. 공통 및 AI 데이터 추출 함수
# ==========================================
def process_uploaded_file_to_image(file_obj):
    """여권 등 이미지 분석 시 선명도 유지 및 경량화"""
    if file_obj.type == "application/pdf":
        doc = fitz.open(stream=file_obj.read(), filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    else:
        img = Image.open(file_obj)
        if img.mode != "RGB":
            img = img.convert("RGB")
    
    if img.width > 1400:
        ratio = 1400 / img.width
        new_size = (1400, int(img.height * ratio))
        img = img.resize(new_size, Image.Resampling.LANCZOS)
    
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    buf.seek(0)
    return Image.open(buf)

def format_full_name(surname, given_name):
    s = surname.strip()
    g = given_name.strip()
    if not s and not g: return ""
    if not s: return g
    if not g: return s
    return f"{g} {s}"

# ⚡ [신규 추가] PDF 텍스트 초고속 추출 로직
def prepare_document_for_gemini(file_bytes, mime_type, file_name=""):
    """PDF에서 텍스트만 뽑아 속도를 10배 이상 향상시킵니다."""
    if "pdf" in mime_type.lower():
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            for page in doc:
                text += page.get_text("text") + "\n"
            
            # 텍스트가 정상적으로 추출되었다면 가벼운 텍스트로 넘김
            if len(text.strip()) > 100:
                return f"\n--- [Document: {file_name}] ---\n{text}\n"
        except:
            pass
    
    # 텍스트 추출이 불가능한 스캔본/이미지면 기존처럼 바이너리로 넘김
    return {"mime_type": mime_type, "data": file_bytes}

def extract_imm5476_info(image):
    prompt = """
    Analyze this identity document (passport/permit/visa) carefully.
    Extract the following details into exact JSON structure:
    - surname: Family name in English uppercase
    - given_name: Given names in English uppercase
    - dob: Date of birth in YYYY-MM-DD format
    - uci: UCI numbers only if present (10 digits or 8 digits), else empty string
    Return ONLY raw valid JSON object without markdown or code formatting.
    """
    try:
        response = model.generate_content([prompt, image])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"정보 추출 오류: {e}")
        return None

def extract_all_passports_batch(has_non_acc, images):
    prompt = f"""
    You are an expert OCR system specialized in international passports.
    I am providing {len(images)} passport image(s) in exact order.
    Order structure:
    - {'Image 1 is the non-accompanying parent passport.' if has_non_acc else 'There is no non-accompanying parent passport provided.'}
    - {'Remaining images (Image ' + ('2' if has_non_acc else '1') + f' to {len(images)}) are accompanying family members (parents or children).' if (len(images) > (1 if has_non_acc else 0)) else 'No family passports provided.'}

    For EACH passport image, extract:
    - surname: Surname / Family name in English uppercase
    - given_name: Given name(s) in English uppercase
    - dob: Date of birth in YYYY-MM-DD format
    - passport_number: Passport number in uppercase alphanumeric
    - gender: Sex of the person, strictly "F" or "M"

    Return ONLY a raw valid JSON object:
    {{
      "non_accompanying_parent": {{
        "surname": "...", "given_name": "...", "dob": "YYYY-MM-DD", "passport_number": "...", "gender": "M"
      }} or null,
      "family_members": [
        {{ "surname": "...", "given_name": "...", "dob": "YYYY-MM-DD", "passport_number": "...", "gender": "F" }}
      ]
    }}
    """
    contents = [prompt] + images
    try:
        response = model.generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"여권 일괄 추출 오류: {e}")
        return None

def extract_case_prep_info(tmpl_bytes, client_files):
    prompt = """
    You are helping immigration case-prep staff. 
    The FIRST document attached is a BLANK reference immigration form (Canadian IRCC "IMM" series) — read it to find every field that requires client-specific data entry.
    The REMAINING attached documents are the client's own materials (intake questionnaire, passport, permit, etc.) — use them as your source of truth for values.

    Return ONLY a raw JSON object in this exact shape:
    {
      "sections": [
        {
          "section": "short section name from the form",
          "fields": [
            { "field": "field label", "value": "value found or empty", "source": "source document or empty" }
          ]
        }
      ]
    }
    Rules: Keep values exact. If missing, set value/source to empty string. Maintain strict order.
    """
    
    # 초고속 텍스트 변환 로직 탑재 🚀
    contents = [prompt]
    contents.append(prepare_document_for_gemini(tmpl_bytes, "application/pdf", "Blank_IMM_Form.pdf"))
    
    for f in client_files:
        mime = f.type if f.type else "application/pdf"
        contents.append(prepare_document_for_gemini(f.getvalue(), mime, f.name))

    try:
        response = model.generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('
