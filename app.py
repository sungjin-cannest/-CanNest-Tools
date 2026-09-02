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
# 2. 공통 캐싱 및 이미지 처리 함수
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
            widget.text_fontsize = 0
            if hasattr(widget, "field_flags") and widget.field_flags:
                widget.field_flags &= ~1
                
            field_name = widget.field_name
            if not field_name: continue
            fname_lower = field_name.lower()
            
            if "family name" in fname_lower and not flags["surname"]:
                widget.field_value = target_data["surname"]; widget.update(); flags["surname"] = True
            elif "given name" in fname_lower and not flags["given"]:
                widget.field_value = target_data["given"]; widget.update(); flags["given"] = True
            elif "date of birth" in fname_lower and not flags["dob"]:
                widget.field_value = target_data["dob"]; widget.update(); flags["dob"] = True
            elif "email" in fname_lower and not flags["email"]:
                widget.field_value = target_data["email"]; widget.update(); flags["email"] = True
            elif ("uci" in fname_lower or "unique client identifier" in fname_lower) and not flags["uci"]:
                widget.field_value = target_data["uci"]; widget.update(); flags["uci"] = True
            elif "date" in fname_lower and "birth" not in fname_lower:
                date_counter += 1
                if date_counter == 1: widget.field_value = target_data["signDate"]; widget.update()

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
            widget.text_fontsize = 0
            if hasattr(widget, "field_flags") and widget.field_flags:
                widget.field_flags &= ~1
            
            fname = widget.field_name.strip() if widget.field_name else ""
            if not fname: continue
            if fname == "1": widget.field_value = data.get("non_acc_name", ""); widget.update()
            elif fname == "2": widget.field_value = data.get("non_acc_address", ""); widget.update()
            elif fname == "3": widget.field_value = data.get("non_acc_phone", ""); widget.update()
            elif fname == "email": widget.field_value = data.get("non_acc_email", ""); widget.update()
            elif fname == "Check Box1": widget.field_value = "0"; widget.update()
            elif fname == "This child or these children hashave my or our consent to travel with": widget.field_value = data.get("acc_name", ""); widget.update()
            elif fname == "Relationship with Children 1": widget.field_value = data.get("acc_relationship", ""); widget.update()
            elif fname == "Relationship with Children 2": widget.field_value = data.get("acc_passport", ""); widget.update()
            elif fname == "I give my consent for this child to travel to": widget.field_value = "Canada"; widget.update()
            elif fname == "2_4": widget.field_value = data.get("acc_name", ""); widget.update()
            elif fname == "At the following addresses 1": widget.field_value = data.get("trip_address", ""); widget.update()
            elif fname == "At the following addresses 2": widget.field_value = data.get("trip_phone", ""); widget.update()
            elif fname == "email_2": widget.field_value = data.get("trip_email", ""); widget.update()
            elif fname == "yyyymmdd_2": widget.field_value = data.get("sign_date", ""); widget.update()
                
            for idx, (name_key, dob_key) in enumerate(child_widgets):
                if idx < len(page_children):
                    if fname == name_key: widget.field_value = page_children[idx].get("name", ""); widget.update()
                    elif fname == dob_key: widget.field_value = page_children[idx].get("dob", ""); widget.update()

    output_pdf = io.BytesIO()
    doc.save(output_pdf); doc.close(); output_pdf.seek(0)
    return output_pdf

# ==========================================
# 4. CRM 서류 판별 및 압축 변환 엔진 (4번 툴)
# ==========================================
def analyze_document_with_crm_rules(file_bytes, mime_type, original_filename):
    prompt = """
    You are an expert AI document classifier for a Canadian immigration firm.
    Read the provided document carefully and generate an EXACT filename according to our STRICT internal CRM rules.

    [CRITICAL NAMING RULES]
    1. Client Name:
       - Korean client: Full Name in Korean with NO SPACES (e.g., 홍길동, 김영미).
       - Non-Korean client: STRICTLY the VERY FIRST WORD of their Given Name / First Name in Title Case. (e.g., If the name is "Janani SARATH BABU", extract "Janani"). Look very closely at fields like "Client personal details", "Given names", or "Applicant".
       - ONLY if the document has absolutely ZERO text (like a blank Digital Photo), use the word "NAME". Otherwise, try your best to find the name.
    2. Dates:
       - Format MUST be YYYY.MM.DD (e.g., 2022.01.02).
       - Year ONLY where specified in the manual (e.g., COI, Emedical, Bank Statement). For Emedical, look for the examination year (e.g., 2026).
    3. Delimiter: Always use underscore '_' between Name, Category, Details, and Dates.
    4. Unlisted Documents: If the document does NOT match any category in the manual list below, look at the TOP of the document to find its exact title in English.

    [MANUAL CATEGORY & FORMAT SPECIFICATIONS]
    - Digital Photo / Passport Photo: {Name}_Digital Photo.jpg (MUST use .jpg extension)
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

    Return ONLY a raw JSON object with this format:
    {
        "client_name": "Extracted name",
        "doc_category": "Category from manual OR exact title",
        "suggested_filename": "Full_Generated_Filename_With_Correct_Extension"
    }
    """
    
    img_data = None
    if "pdf" in mime_type.lower():
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=70)
            img_data = {"mime_type": "image/jpeg", "data": buf.getvalue()}
            doc.close()
        except Exception:
            return {"client_name": "NAME", "doc_category": "기타", "suggested_filename": f"NAME_미분류_{original_filename}"}
    else:
        try:
            img = Image.open(io.BytesIO(file_bytes))
            if img.mode != "RGB":
                img = img.convert("RGB")
            buf = io.BytesIO()
            if img.width > 2000:
                ratio = 2000 / img.width
                img = img.resize((2000, int(img.height * ratio)), Image.Resampling.LANCZOS)
            img.save(buf, format="JPEG", quality=75)
            img_data = {"mime_type": "image/jpeg", "data": buf.getvalue()}
        except Exception:
            img_data = {"mime_type": mime_type, "data": file_bytes}

    try:
        response = safe_generate_content([prompt, img_data])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        data = json.loads(clean_text)
        
        filename = data.get("suggested_filename", "NAME_미분류_서류.pdf")
        if not (filename.lower().endswith(".pdf") or filename.lower().endswith(".jpg") or filename.lower().endswith(".jpeg")):
            filename += ".pdf"
            
        data["suggested_filename"] = filename
        return data
        
    except Exception as e:
        st.error(f"서류 분석 중 오류가 발생했습니다 ({original_filename}): {e}")
        base_name = os.path.splitext(original_filename)[0]
        return {"client_name": "NAME", "doc_category": "기타", "suggested_filename": f"NAME_{base_name}.pdf"}

def process_and_compress_file(file_bytes, mime_type, target_filename):
    is_jpeg = target_filename.lower().endswith(('.jpg', '.jpeg'))
    
    if is_jpeg:
        img = Image.open(io.BytesIO(file_bytes))
        if img.mode != "RGB":
            img = img.convert("RGB")
            
        if img.width > 2400:
            ratio = 2400 / img.width
            img = img.resize((2400, int(img.height * ratio)), Image.Resampling.LANCZOS)
            
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85, optimize=True)
        buf.seek(0)
        return buf.getvalue(), "image/jpeg"
        
    else:
        output_pdf = io.BytesIO()
        target_dpi = 150
        quality = 65
        
        if "pdf" in mime_type.lower():
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            new_doc = fitz.open()
            
            for page in doc:
                zoom = target_dpi / 72.0
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                img_buf = io.BytesIO()
                img.save(img_buf, format="JPEG", quality=quality, optimize=True)
                img_buf.seek(0)
                
                pdf_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
                pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())
                
            new_doc.save(output_pdf, deflate=True, garbage=4)
            new_doc.close()
            doc.close()
        else:
            img = Image.open(io.BytesIO(file_bytes))
            if img.mode != "RGB":
                img = img.convert("RGB")
                
            if img.width > 1800:
                ratio = 1800 / img.width
                img = img.resize((1800, int(img.height * ratio)), Image.Resampling.LANCZOS)
                
            img_buf = io.BytesIO()
            img.save(img_buf, format="JPEG", quality=quality, optimize=True)
            img_buf.seek(0)
            
            new_doc = fitz.open()
            page_width = img.width * 72 / target_dpi
            page_height = img.height * 72 / target_dpi
            pdf_page = new_doc.new_page(width=page_width, height=page_height)
            
            pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())
            
            new_doc.save(output_pdf)
            new_doc.close()
            
        output_pdf.seek(0)
        return output_pdf.getvalue(), "application/pdf"

# ==========================================
# 5. Streamlit 네비게이션 및 UI 구성 (4개 툴 통합)
# ==========================================
st.set_page_config(page_title="CanNest 통합 업무 시스템", layout="wide")

st.sidebar.title("🦅 CanNest Tool")
app_mode = st.sidebar.radio("원하시는 업무 도구를 선택하세요", [
    "🍁 IMM5476 자동 작성", 
    "✈️ 한부모 동의서 자동 작성",
    "📋 이민서류 정보 정리 (Case File Prep)",
    "🏷️ CRM 파일명 자동 생성 및 최적화"
])

# ------------------------------------------
# 메뉴 1: IMM5476 자동 작성
# ------------------------------------------
if app_mode == "🍁 IMM5476 자동 작성":
    st.title("🍁 IMM5476 자동 작성 도구")
    if "extracted_5476" not in st.session_state: st.session_state.extracted_5476 = None

    template_5476_bytes = get_preloaded_file_bytes(["imm5476_template.pdf", "imm5476_template.pdf.pdf"])
    
    if template_5476_bytes:
        st.success("✅ 사내 표준 'IMM5476' 양식이 자동으로 로드되었습니다.")
    else:
        st.error("⚠️ GitHub에 'imm5476_template.pdf' 파일이 없습니다. 수동으로 업로드해 주세요.")
        template_file = st.file_uploader("IMM5476 템플릿 PDF 선택", type=['pdf'], key="template_5476")
        if template_file: template_5476_bytes = template_file.getvalue()

    st.markdown("---")
    client_file = st.file_uploader("1. 손님 여권 또는 퍼밋", type=['jpg', 'jpeg', 'png', 'pdf'], key="client_5476")

    if client_file and st.button("정보 추출하기", use_container_width=True):
        with st.spinner("서류 분석 중입니다. 잠시만 기다려 주세요..."):
            extracted = extract_imm5476_info(process_uploaded_file_to_image(client_file))
            if extracted: st.session_state.extracted_5476 = extracted; st.success("정보 추출이 완료되었습니다.")

    if st.session_state.extracted_5476:
        data = st.session_state.extracted_5476
        c1, c2 = st.columns(2)
        with c1:
            surname = st.text_input("성 (Surname)", data.get("surname", ""))
            dob = st.text_input("생년월일", data.get("dob", ""))
            email = st.text_input("이메일", "")
        with c2:
            given = st.text_input("이름 (Given Name)", data.get("given_name", ""))
            uci = st.text_input("UCI", data.get("uci", ""))
            sign_date = st.date_input("서명날짜", datetime.date.today())

        if st.button("문서 생성 및 다운로드", type="primary"):
            if not template_5476_bytes:
                st.error("템플릿 파일이 없습니다.")
            else:
                final_data = {"surname": surname, "given_name": given, "dob": dob, "uci": uci, "email": email, "signDate": sign_date.strftime("%Y-%m-%d")}
                pdf_out = fill_imm5476(template_5476_bytes, final_data)
                st.download_button("📥 다운로드", pdf_out, file_name=f"IMM5476_{surname}_{given}.pdf", mime="application/pdf")

# ------------------------------------------
# 메뉴 2: 한부모 동의서 자동 작성
# ------------------------------------------
elif app_mode == "✈️ 한부모 동의서 자동 작성":
    st.title("✈️ 한부모 동의서 자동 작성 도구")
    if "consent_non_acc" not in st.session_state: st.session_state.consent_non_acc = {}
    if "consent_family" not in st.session_state: st.session_state.consent_family = []

    consent_template_bytes = get_preloaded_file_bytes(["consent_template.pdf", "consent_template.pdf.pdf"])

    if consent_template_bytes:
        st.success("✅ 사내 표준 '한부모 동의서' 양식이 자동으로 로드되었습니다.")
    else:
        st.error("⚠️ GitHub에 'consent_template.pdf' 파일이 없습니다. 수동으로 업로드해 주세요.")
        consent_template = st.file_uploader("동의서 양식 선택", type=['pdf'])
        if consent_template: consent_template_bytes = consent_template.getvalue()

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1: non_acc_file = st.file_uploader("비동반 부모님 여권 (1장)", type=['jpg', 'jpeg', 'png', 'pdf'])
    with c2: family_files = st.file_uploader("동반 부모/자녀 여권", type=['jpg', 'jpeg', 'png', 'pdf'], accept_multiple_files=True)

    if st.button("여권 정보 추출하기", type="primary", use_container_width=True):
        images = []
        has_non_acc = bool(non_acc_file)
        if non_acc_file: images.append(process_uploaded_file_to_image(non_acc_file))
        if family_files: images.extend([process_uploaded_file_to_image(f) for f in family_files])
        
        if images:
            with st.spinner("여권 정보를 분석 중입니다. 잠시만 기다려 주세요..."):
                res = extract_all_passports_batch(has_non_acc, images)
                if res:
                    st.session_state.consent_non_acc = res.get("non_accompanying_parent", {}) or {}
                    st.session_state.consent_family = res.get("family_members", []) or []
                    st.success("여권 정보 추출이 완료되었습니다.")

    non_acc_data = st.session_state.consent_non_acc
    non_acc_name = st.text_input("비동반 부모 성명", format_full_name(non_acc_data.get('surname',''), non_acc_data.get('given_name','')))
    non_acc_address = st.text_input("주소")
    ca, cb = st.columns(2)
    with ca: non_acc_phone = st.text_input("전화번호")
    with cb: non_acc_email = st.text_input("이메일")

    acc_parents = [p for p in st.session_state.consent_family if not is_minor(p.get("dob", ""))]
    children_list = [p for p in st.session_state.consent_family if is_minor(p.get("dob", ""))]
    
    acc_name, acc_passport, acc_rel = "", "", "Mother"
    if acc_parents:
        p = acc_parents[0]
        acc_name = format_full_name(p.get('surname',''), p.get('given_name',''))
        acc_passport = p.get("passport_number", "")
        acc_rel = "Mother" if p.get("gender") == "F" else "Father"

    cp1, cp2, cp3 = st.columns(3)
    with cp1: acc_name = st.text_input("동반 부모 성명", acc_name)
    with cp2: acc_rel = st.selectbox("관계", ["Mother", "Father"], index=0 if acc_rel=="Mother" else 1)
    with cp3: acc_passport = st.text_input("여권번호", acc_passport)

    final_children = []
    for idx, c in enumerate(children_list):
        cc1, cc2 = st.columns(2)
        with cc1: n = st.text_input(f"자녀{idx+1} 성명", format_full_name(c.get('surname',''), c.get('given_name','')))
        with cc2: d = st.text_input(f"자녀{idx+1} 생일", c.get("dob", "").replace("-", "/"))
        final_children.append({"name": n, "dob": d})
    
    if not children_list:
        cc1, cc2 = st.columns(2)
        with cc1: n = st.text_input("자녀 성명")
        with cc2: d = st.text_input("자녀 생일")
        if n: final_children.append({"name": n, "dob": d})

    trip_address = st.text_input("현지 주소")
    ct1, ct2 = st.columns(2)
    with ct1: trip_phone = st.text_input("현지 전화")
    with ct2: trip_email = st.text_input("현지 이메일")
    sign_date_str = st.date_input("서명일", datetime.date.today()).strftime("%Y/%m/%d")

    if st.button("문서 생성 및 다운로드", type="primary"):
        if not consent_template_bytes:
            st.error("양식 파일이 없습니다.")
        else:
            data_consent = {
                "non_acc_name": non_acc_name, "non_acc_address": non_acc_address, "non_acc_phone": non_acc_phone, "non_acc_email": non_acc_email,
                "children": final_children, "acc_name": acc_name, "acc_relationship": acc_rel, "acc_passport": acc_passport,
                "trip_address": trip_address, "trip_phone": trip_phone, "trip_email": trip_email, "sign_date": sign_date_str
            }
            pdf_out = fill_consent_letter(consent_template_bytes, data_consent)
            st.download_button("📥 다운로드", pdf_out, f"Consent_{non_acc_name}.pdf", "application/pdf")

# ------------------------------------------
# 메뉴 3: 이민서류 정보 정리 (Case File Prep)
# ------------------------------------------
elif app_mode == "📋 이민서류 정보 정리 (Case File Prep)":
    st.title("📋 이민서류 정보 정리 도구")
    
    if "prep_result" not in st.session_state:
        st.session_state.prep_result = None

    st.subheader("1. 대상 서식 (IMM PDF)")
    
    form_map = {
        "직접 파일 업로드 (기타 서식)": None,
        "IMM1294 (SP-OUTSIDE)": "imm1294.pdf",
        "IMM1295 (WP-OUTSIDE)": "imm1295.pdf",
        "IMM5708 (VR-INSIDE)": "imm5708.pdf",
        "IMM5709 (SP-INSIDE)": "imm5709.pdf",
        "IMM5710 (WP-INSIDE)": "imm5710.pdf"
    }
    
    selected_form = st.selectbox("📌 템플릿 서식 선택 (GitHub 박제본)", list(form_map.keys()))
    
    tmpl_bytes = None
    
    if form_map[selected_form] is not None:
        file_path = form_map[selected_form]
        tmpl_bytes = load_pdf_bytes_cached(file_path)
        if tmpl_bytes:
            st.success(f"✅ '{selected_form}' 양식이 자동으로 로드되었습니다.")
        else:
            st.error(f"⚠️ {file_path} 파일이 서버에 없습니다. 파일 업로드 상태를 확인해 주세요.")
    else:
        tmpl_prep_file = st.file_uploader("빈 IMM 서식 (반드시 Print to PDF로 평탄화된 파일)", type=['pdf'], key="case_tmpl")
        if tmpl_prep_file:
            tmpl_bytes = tmpl_prep_file.getvalue()

    st.markdown("---")
    st.subheader("2. 손님 제출 서류 (복수 선택 가능)")
    client_prep_files = st.file_uploader("질문지, 여권, 퍼밋 등 서류 선택", type=['jpg', 'jpeg', 'png', 'pdf'], accept_multiple_files=True, key="case_client_docs")

    if st.button("서류 정보 정리하기", type="primary", use_container_width=True):
        if tmpl_bytes is None:
            st.warning("1번 단계에서 서식이 정상적으로 선택되거나 업로드되지 않았습니다.")
        elif not client_prep_files:
            st.warning("2번 단계에서 손님 서류를 1개 이상 올려주세요.")
        else:
            with st.spinner("서류를 대조하여 정보 및 불일치 항목을 확인 중입니다. 잠시만 기다려 주세요..."):
                res = extract_case_prep_info(tmpl_bytes, client_prep_files)
                if res:
                    st.session_state.prep_result = res
                    st.success("서류 정보 정리가 완료되었습니다.")

    if st.session_state.prep_result:
        st.markdown("---")
        st.subheader("3. 정리된 정보 결과")

        parsed = st.session_state.prep_result
        sections = parsed.get("sections", [])

        if not sections:
            st.error("서식에서 분석할 항목을 찾지 못했습니다. (파일이 평탄화된 PDF인지 확인해 주세요)")
        else:
            full_text_list = []
            for sec in sections:
                sec_name = sec.get("section", "기타 항목")
                st.write(f"### 📌 {sec_name}")
                full_text_list.append(f"[{sec_name}]")

                table_data = []
                for f in sec.get("fields", []):
                    field_lbl = f.get("field", "")
                    val = f.get("value", "")
                    src = f.get("source", "")

                    if not val:
                        display_val = "⚠️ 확인 필요 (미발견)"
                        full_text_list.append(f"{field_lbl}: (확인 필요)")
                    else:
                        display_val = val
                        full_text_list.append(f"{field_lbl}: {val}")

                    table_data.append({
                        "항목 (Field)": field_lbl,
                        "추출값 (Value)": display_val,
                        "출처 (Source)": src if src else "-"
                    })

                st.table(table_data)
                full_text_list.append("")

            st.markdown("#### 📋 한눈에 복사하기")
            st.text_area("아래 텍스트를 복사하여 서식에 옮겨 적으세요", value="\n".join(full_text_list), height=250)
            
            if st.button("＋ 새 케이스 정리하기"):
                st.session_state.prep_result = None
                st.rerun()

# ------------------------------------------
# 메뉴 4: CRM 파일명 자동 생성 및 최적화
# ------------------------------------------
elif app_mode == "🏷️ CRM 파일명 자동 생성 및 최적화":
    st.title("🏷️ CRM 파일명 자동 생성 및 최적화 도구")
    st.caption("고객 서류 업로드 시 AI가 파일명을 규칙에 맞게 자동 생성하고 즉시 압축 변환합니다.")

    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = str(uuid.uuid4())
    if "analysis_results" not in st.session_state:
        st.session_state.analysis_results = None

    uploaded_files = st.file_uploader(
        "서류 업로드 (복수 선택 가능)", 
        type=['jpg', 'jpeg', 'png', 'pdf', 'heic'], 
        accept_multiple_files=True,
        key=st.session_state.uploader_key
    )

    if uploaded_files:
        if st.button("서류 분석 및 최적화 시작", type="primary", use_container_width=True):
            results = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"서류 분석 및 압축 중 ({idx+1}/{len(uploaded_files)}): {file.name} - 잠시만 기다려 주세요...")
                file_bytes = file.getvalue()
                mime_type = file.type if file.type else "application/pdf"
                
                analysis = analyze_document_with_crm_rules(file_bytes, mime_type, file.name)
                final_name = analysis.get("suggested_filename", file.name)
                
                compressed_bytes, out_mime = process_and_compress_file(file_bytes, mime_type, final_name)
                
                orig_kb = len(file_bytes) / 1024
                comp_kb = len(compressed_bytes) / 1024
                
                results.append({
                    "original_name": file.name,
                    "suggested_filename": final_name,
                    "category": analysis.get("doc_category", "기타"),
                    "client_name": analysis.get("client_name", ""),
                    "mime": out_mime,
                    "orig_kb": orig_kb,
                    "comp_kb": comp_kb,
                    "bytes": compressed_bytes
                })
                
                progress_bar.progress((idx + 1) / len(uploaded_files))

                if idx < len(uploaded_files) - 1:
                    time.sleep(1)
                
            status_text.success("모든 서류의 분석 및 최적화가 완료되었습니다.")
            st.session_state.analysis_results = results

    if st.session_state.analysis_results:
        st.markdown("---")
        st.subheader("변환 완료된 서류 다운로드")
        st.info("💡 파일명이 잘못 설정된 경우, 다운로드 버튼 옆 텍스트 입력창에서 직접 수정한 후 다운로드할 수 있습니다.")
        
        zip_buffer = io.BytesIO()
        final_downloads = []
        
        for idx, item in enumerate(st.session_state.analysis_results):
            col1, col2, col3 = st.columns([3, 3, 2])
            
            with col1:
                st.write(f"**원본**: `{item['original_name']}`")
                st.caption(f"{item['orig_kb']:.1f} KB ➡️ **{item['comp_kb']:.1f} KB**")
                
            with col2:
                user_edited_name = st.text_input(
                    "파일명", 
                    value=item['suggested_filename'], 
                    key=f"edit_{idx}_{item['original_name']}",
                    label_visibility="collapsed"
                )
                final_downloads.append((user_edited_name, item['bytes'], item['mime']))
                
            with col3:
                st.download_button("⬇️ 개별 다운로드", data=item['bytes'], file_name=user_edited_name, mime=item['mime'], key=f"dl_btn_{idx}")
                
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for fname, fbytes, _ in final_downloads:
                zip_file.writestr(fname, fbytes)
                
        zip_buffer.seek(0)
        today_str = datetime.date.today().strftime("%Y%m%d")
        
        st.markdown("---")
        st.download_button(
            "📦 전체 서류 ZIP 다운로드",
            data=zip_buffer,
            file_name=f"CRM_Documents_{today_str}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

        st.markdown("---")
        if st.button("🔄 전체 리셋", type="secondary", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.session_state.uploader_key = str(uuid.uuid4())
            st.rerun()
