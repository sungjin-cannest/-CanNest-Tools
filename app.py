import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
from PIL import Image
import pillow_heif  # HEIC 지원 라이브러리
import io
import zipfile
import json
import os
import datetime
import time
import uuid
from concurrent.futures import ThreadPoolExecutor

# HEIC 이미지 지원 등록 및 고해상도 제한 해제
pillow_heif.register_heif_opener()
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

def safe_generate_content(contents):
    candidate_models = ['gemini-3.6-flash']
    last_error = None
    for model_name in candidate_models:
        try:
            mod = genai.GenerativeModel(model_name)
            response = mod.generate_content(contents)
            return response
        except Exception as e:
            last_error = e
            if "404" in str(e) or "not found" in str(e).lower():
                continue
            else:
                raise e
    raise last_error

# ==========================================
# 2. 공통 캐싱 및 스마트 글자 크기 조절 함수
# ==========================================
@st.cache_data(ttl=3600, show_spinner=False)
def load_pdf_bytes_cached(file_path):
    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            return f.read()
    return None

def get_preloaded_file_bytes(file_names):
    for fn in file_names:
        data = load_pdf_bytes_cached(fn)
        if data:
            return data
    return None

def process_uploaded_file_to_image(file_obj):
    if file_obj.type == "application/pdf":
        doc = fitz.open(stream=file_obj.read(), filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    else:
        img = Image.open(file_obj)
        if img.mode != "RGB":
            img = img.convert("RGB")
    
    if img.width > 1500:
        ratio = 1500 / img.width
        new_size = (1500, int(img.height * ratio))
        img = img.resize(new_size, Image.Resampling.LANCZOS)
    
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=70)
    buf.seek(0)
    return Image.open(buf)

def format_full_name(surname, given_name):
    s = surname.strip()
    g = given_name.strip()
    if not s and not g: return ""
    if not s: return g
    if not g: return s
    return f"{g} {s}"

def set_smart_widget_value(widget, value, default_fontsize=10, min_fontsize=5.5):
    val_str = str(value) if value is not None else ""
    widget.field_value = val_str
    
    if hasattr(widget, "field_flags") and widget.field_flags:
        widget.field_flags &= ~1 
        
    if val_str and hasattr(widget, "rect"):
        box_width = widget.rect.width - 4 
        if box_width > 0:
            try:
                font = fitz.Font("helv")
                len_at_default = font.text_length(val_str, fontsize=default_fontsize)
                if len_at_default > box_width:
                    len_at_1 = font.text_length(val_str, fontsize=1)
                    if len_at_1 > 0:
                        scaled_size = box_width / len_at_1
                        widget.text_fontsize = max(min_fontsize, min(default_fontsize, scaled_size))
                    else:
                        widget.text_fontsize = default_fontsize
                else:
                    widget.text_fontsize = default_fontsize
            except Exception:
                widget.text_fontsize = default_fontsize
        else:
            widget.text_fontsize = default_fontsize
    else:
        widget.text_fontsize = default_fontsize
        
    widget.update()

def prepare_document_for_gemini(file_bytes, mime_type, file_name=""):
    if "pdf" in mime_type.lower():
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            for page in doc:
                text += page.get_text("text") + "\n"
            
            if len(text.strip()) > 100:
                return [f"\n--- [Document: {file_name}] ---\n{text[:20000]}\n"]
            else:
                images = []
                for page_num in range(min(len(doc), 10)):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=70)
                    buf.seek(0)
                    images.append({"mime_type": "image/jpeg", "data": buf.getvalue()})
                return images
        except:
            pass
    return [{"mime_type": mime_type, "data": file_bytes}]

def batch_process_client_files(client_files):
    def worker(f):
        mime = f.type if f.type else "application/pdf"
        return prepare_document_for_gemini(f.getvalue(), mime, f.name)

    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(executor.map(worker, client_files))

    flat_contents = []
    for res in results:
        flat_contents.extend(res)
    return flat_contents

def is_minor(dob_str):
    try:
        birth_date = datetime.datetime.strptime(dob_str, "%Y-%m-%d").date()
        today = datetime.date.today()
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        return age < 19
    except:
        return True

# ==========================================
# 3. AI 분석 및 PDF 채우기 로직 (1, 2, 3번 툴)
# ==========================================
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
        response = safe_generate_content([prompt, image])
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
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"여권 일괄 추출 오류: {e}")
        return None

def extract_case_prep_info(tmpl_bytes, client_files):
    prompt = """
    You are an expert Canadian immigration case prep assistant.
    The FIRST document is a BLANK reference IRCC IMM form template.
    The REMAINING attached documents are client materials (intake questionnaire, passport, work/study permit, visitor record, resume, WES, etc.).

    CRITICAL RULE FOR "Current country or territory of residence" (Section 7 in Personal Details):
    - Status, From date, and To date in this section MUST be extracted directly from the client's CURRENT PERMIT / VISA document (Work Permit, Study Permit, Visitor Record).
      * Country/Territory: Canada
      * Status: Worker / Student / Visitor (based on current permit type)
      * From: Permit Issue / Effective Date
      * To: Permit Expiry Date

    CRITICAL RULE FOR CROSS-DOCUMENT MISMATCH DETECTION:
    Compare information across ALL client documents carefully.
    If there is a mismatch between documents (e.g. Permit issue date vs Questionnaire entry date, or Resume employment date vs Questionnaire):
       Set the "value" field strictly as:
       "⚠️ 정보 불일치 (재확인 필요): [DocType A] ValueA vs [DocType B] ValueB"
       Example: "⚠️ 정보 불일치 (재확인 필요): [질문지] 2024-06-09 vs [퍼밋] 2024-06-08"

    Rules for extracting field values:
    1. Standard match: Provide exact value and cite source document name.
    2. Mismatch: Format as warning string above.
    3. Missing: Set as empty string "".

    Return ONLY a raw JSON object in this exact shape:
    {
      "sections": [
        {
          "section": "short section name from the form",
          "fields": [
            { "field": "field label", "value": "extracted value or mismatch warning", "source": "source document name(s)" }
          ]
        }
      ]
    }
    Maintain strict order of fields as requested by the template form.
    """
    
    contents = [prompt]
    contents.extend(prepare_document_for_gemini(tmpl_bytes, "application/pdf", "Blank_IMM_Form.pdf"))
    contents.extend(batch_process_client_files(client_files))

    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"서류 정리 오류: {e}")
        return None

def fill_imm5476(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    target_data = {
        "surname": data.get("surname", ""), "given": data.get("given_name", ""),
        "dob": data.get("dob", ""), "email": data.get("email", ""),
        "uci": data.get("uci", "").replace("-", ""), "signDate": data.get("signDate", "")  
    }
    flags = {"surname": False, "given": False, "dob": False, "email": False, "uci": False}
    date_counter = 0
    
    for page in doc:
        for widget in page.widgets():
            field_name = widget.field_name
            if not field_name: continue
            fname_lower = field_name.lower()
            
            val_to_set = None
            if "family name" in fname_lower and not flags["surname"]:
                val_to_set = target_data["surname"]; flags["surname"] = True
            elif "given name" in fname_lower and not flags["given"]:
                val_to_set = target_data["given"]; flags["given"] = True
            elif "date of birth" in fname_lower and not flags["dob"]:
                val_to_set = target_data["dob"]; flags["dob"] = True
            elif "email" in fname_lower and not flags["email"]:
                val_to_set = target_data["email"]; flags["email"] = True
            elif ("uci" in fname_lower or "unique client identifier" in fname_lower) and not flags["uci"]:
                val_to_set = target_data["uci"]; flags["uci"] = True
            elif "date" in fname_lower and "birth" not in fname_lower:
                date_counter += 1
                if date_counter == 1: val_to_set = target_data["signDate"]
                
            if val_to_set is not None:
                set_smart_widget_value(widget, val_to_set, default_fontsize=10)

    output_pdf = io.BytesIO()
    doc.save(output_pdf); doc.close(); output_pdf.seek(0)
    return output_pdf

def fill_consent_letter(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    children = data.get("children", [])
    num_pages_needed = max(1, (len(children) + 2) // 3)
    for _ in range(num_pages_needed - 1): doc.insert_pdf(doc, from_page=0, to_page=0)
        
    for page_num in range(num_pages_needed):
        page = doc[page_num]
        page_children = children[page_num * 3 : (page_num + 1) * 3]
        child_widgets = [("Information about travelling children", "yyyymmdd"), ("1_2", "2_2"), ("1_3", "2_3")]
        
        for widget in page.widgets():
            fname = widget.field_name.strip() if widget.field_name else ""
            if not fname: continue
            
            val_to_set = None
            if fname == "1": val_to_set = data.get("non_acc_name", "")
            elif fname == "2": val_to_set = data.get("non_acc_address", "")
            elif fname == "3": val_to_set = data.get("non_acc_phone", "")
            elif fname == "email": val_to_set = data.get("non_acc_email", "")
            elif fname == "Check Box1": val_to_set = "0"
            elif fname == "This child or these children hashave my or our consent to travel with": val_to_set = data.get("acc_name", "")
            elif fname == "Relationship with Children 1": val_to_set = data.get("acc_relationship", "")
            elif fname == "Relationship with Children 2": val_to_set = data.get("acc_passport", "")
            elif fname == "I give my consent for this child to travel to": val_to_set = "Canada"
            elif fname == "2_4": val_to_set = data.get("acc_name", "")
            elif fname == "At the following addresses 1": val_to_set = data.get("trip_address", "")
            elif fname == "At the following addresses 2": val_to_set = data.get("trip_phone", "")
            elif fname == "email_2": val_to_set = data.get("trip_email", "")
            elif fname == "yyyymmdd_2": val_to_set = data.get("sign_date", "")
            else:
                for idx, (name_key, dob_key) in enumerate(child_widgets):
                    if idx < len(page_children):
                        if fname == name_key: val_to_set = page_children[idx].get("name", "")
                        elif fname == dob_key: val_to_set = page_children[idx].get("dob", "")
            
            if val_to_set is not None:
                set_smart_widget_value(widget, val_to_set, default_fontsize=10)

    output_pdf = io.BytesIO()
    doc.save(output_pdf); doc.close(); output_pdf.seek(0)
    return output_pdf

# ==========================================
# 4. CRM 서류 판별 및 자동 분할(Split) 엔진 (4번 툴)
# ==========================================
def split_pdf_bytes(file_bytes, start_page, end_page):
    """1-indexed 페이지 번호를 받아 PDF를 자르고 반환합니다."""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    new_doc = fitz.open()
    
    start_idx = max(0, start_page - 1)
    end_idx = min(len(doc) - 1, end_page - 1)
    
    for i in range(start_idx, end_idx + 1):
        new_doc.insert_pdf(doc, from_page=i, to_page=i)
        
    output_pdf = io.BytesIO()
    new_doc.save(output_pdf)
    new_doc.close()
    doc.close()
    output_pdf.seek(0)
    return output_pdf.getvalue()

def analyze_and_split_crm_document(file_bytes, mime_type, original_filename):
    prompt = """
    You are an expert AI document classifier for a Canadian immigration firm.
    The attached file may contain a SINGLE document, or it may be a MERGED file containing MULTIPLE different documents (e.g., Page 1 is a Passport, Pages 2-3 are a Study Permit).
    
    Analyze ALL pages. Group consecutive pages that belong to the SAME document.
    For EACH distinct document you identify, generate an EXACT filename according to our STRICT internal CRM rules.

    [CRITICAL NAMING RULES]
    1. Client Name:
       - Korean client: Full Name in Korean with NO SPACES (e.g., 홍길동, 김영미).
       - Non-Korean client: STRICTLY the VERY FIRST WORD of their Given Name / First Name in Title Case. (e.g., If name is "Janani SARATH BABU", extract "Janani").
       - ONLY if the document has ZERO text (like a blank Digital Photo), use the word "NAME".
    2. Dates: YYYY.MM.DD (e.g., 2022.01.02). Year ONLY where specified in the manual.
    3. Delimiter: Always use underscore '_'
    4. Unlisted Documents: If it doesn't match the list below, use its exact title in English.

    [MANUAL CATEGORY & FORMAT SPECIFICATIONS]
    - Digital Photo / Passport Photo: {Name}_Digital Photo.jpg
    - Passport: {Name}_PP_{ExpiryDate YYYY.MM.DD}
    - Work Permit: {Name}_WP_{ExpiryDate YYYY.MM.DD}
    - Study Permit: {Name}_SP_{ExpiryDate YYYY.MM.DD}
    - Visitor Record: {Name}_VR_{ExpiryDate YYYY.MM.DD}
    - Questionnaire: {Name}_QA_{Type e.g. WP/EE/PR Card/PNP/Spouse WP/VR}_{ReceivedDate YYYY.MM.DD}
    - Police Certificate: {Name}_Police Cert_{CountryNameInEnglish}
    - LOE / Employment Letter: {Name}_LOE_{CompanyInEnglish}_{한글/영문/공증}
    - Paystub: {Name}_Paystub_{CompanyInEnglish}_{StartDate YYYY.MM.DD-EndDate YYYY.MM.DD}
    - Degree/Diploma: {Name}_{Diploma/Bachelor/Highschool/Master/Certificate}_{SchoolName}
    - WES: {Name}_WES
    - Certificate of Income: {Name}_COI_{Year YYYY}
    - Language Test: {Name}_{IELTS/CELPIP}_{TestDate YYYY.MM.DD}
    - Resume: {Name}_Resume_{ReceivedDate YYYY.MM.DD}
    - Medical Exam (eMedical Information Sheet): {Name}_Emedical_{Year YYYY}
    - Marriage Cert: {Name}_Marriage Cert_{IssueDate YYYY.MM.DD}
    - Transcript: {Name}_Transcript_{SchoolName}
    - Bank Statement: {Name}_Bank Statement_{Year YYYY}
    - Family Cert: {Name}_Family Cert_{IssueDate YYYY.MM.DD}
    - Basic Cert: {Name}_Basic Cert
    - Birth Cert: {Name}_Birth Cert
    - Travel Consent: {ChildName}_Travel Consent
    - LOA (Letter of Acceptance): {Name}_LOA_{SchoolName}
    - Tuition Receipt: {Name}_Tuition Receipt_{SchoolName}
    - Confirmation of Enrollment: {Name}_Confirmation of Enrollment_{SchoolName}

    Return ONLY a raw JSON object with an array of documents:
    {
        "documents": [
            {
                "client_name": "Extracted name",
                "doc_category": "Category or exact title",
                "suggested_filename": "Full_Generated_Filename_With_Extension",
                "start_page": 1,
                "end_page": 2
            }
        ]
    }
    NOTE: start_page and end_page are 1-indexed. If it's a 1-page document, start_page and end_page are both the same number.
    If you see a Passport on Page 1, and a Bank Statement on Pages 2-4, return TWO distinct objects in the array.
    """
    
    contents = [prompt]
    num_pages = 1
    
    if "pdf" in mime_type.lower():
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            num_pages = len(doc)
            max_pages = min(num_pages, 30) # AI 부하 방지용 최대 30페이지 제한
            
            for page_num in range(max_pages):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=70)
                
                contents.append(f"--- Page {page_num + 1} ---")
                contents.append({"mime_type": "image/jpeg", "data": buf.getvalue()})
            doc.close()
        except Exception:
            pass
    else:
        try:
            img = Image.open(io.BytesIO(file_bytes))
            if img.mode != "RGB": img = img.convert("RGB")
            buf = io.BytesIO()
            if img.width > 2000:
                ratio = 2000 / img.width
                img = img.resize((2000, int(img.height * ratio)), Image.Resampling.LANCZOS)
            img.save(buf, format="JPEG", quality=75)
            contents.append("--- Page 1 ---")
            contents.append({"mime_type": "image/jpeg", "data": buf.getvalue()})
        except Exception:
            contents.append({"mime_type": mime_type, "data": file_bytes})

    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('
