import streamlit as st

# 📌 Streamlit 페이지 설정
st.set_page_config(page_title="CanNest 잡오퍼 DOCX 생성기 (DEV)", layout="wide")

import google.generativeai as genai
import fitz  # PyMuPDF
from PIL import Image, ImageOps
import pillow_heif  # HEIC 지원
import io
import json
import datetime
import re
import urllib.request
import os

# 순수 DOCX 생성 라이브러리 (docxtpl 사용 안 함)
try:
    import docx
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    st.error("⚠️ 'python-docx' 라이브러리가 필요합니다.")

pillow_heif.register_heif_opener()
Image.MAX_IMAGE_PIXELS = None

# ==========================================
# 0. Secrets 안전 검사 및 보안 비밀번호 설정
# ==========================================
if "APP_PASSWORD" not in st.secrets or "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Cloud의 Secrets 설정이 필요합니다.")
    st.stop()

def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.title("🔒 CanNest 잡오퍼 생성기 (DEV)")
        st.text_input("접속 비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title("🔒 CanNest 잡오퍼 생성기 (DEV)")
        st.text_input("접속 비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        st.error("비밀번호가 틀렸습니다.")
        return False
    return True

if not check_password():
    st.stop()

# ==========================================
# 1. API 키 및 가장 안전한 모델 호출 로직
# ==========================================
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

def safe_generate_content(contents):
    # 에러가 발생하지 않는 가장 범용적이고 안정적인 모델명만 고정
    models_to_try = [
        'gemini-1.5-flash',
        'gemini-1.5-pro',
        'gemini-1.0-pro',
        'gemini-pro'
    ]
    
    last_err = None
    for m in models_to_try:
        try:
            model = genai.GenerativeModel(m)
            response = model.generate_content(contents)
            if response.text:
                return response
        except Exception as e:
            last_err = e
            continue  # 실패하면 즉시 다음 모델로 넘어감
            
    raise Exception(f"AI 모델 호출 실패. 마지막 발생 에러: {last_err}")

# ==========================================
# 2. 내장 헬퍼 함수
# ==========================================
def process_uploaded_file_to_image(file_obj):
    if file_obj.type == "application/pdf":
        doc = fitz.open(stream=file_obj.read(), filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    else:
        img = Image.open(file_obj)
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
    
    max_dim = max(img.width, img.height)
    if max_dim > 1800:
        ratio = 1800.0 / float(max_dim)
        new_size = (int(img.width * ratio), int(img.height * ratio))
        img = img.resize(new_size, Image.Resampling.LANCZOS)
    
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=75)
    buf.seek(0)
    return Image.open(buf)

def format_full_name(surname, given_name):
    s = str(surname).strip() if surname else ""
    g = str(given_name).strip() if given_name else ""
    if not s and not g: return ""
    if not s: return g
    if not g: return s
    return f"{g} {s}"

def extract_imm5476_info(image):
    prompt = """
    Analyze this identity document carefully.
    Extract surname, given_name, dob (YYYY-MM-DD), uci into exact JSON structure:
    {"surname": "...", "given_name": "...", "dob": "YYYY-MM-DD", "uci": "..."}
    Return ONLY raw valid JSON object.
    """
    try:
        response = safe_generate_content([prompt, image])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception:
        return None

def prepare_document_for_gemini(file_bytes, mime_type, file_name=""):
    ext = os.path.splitext(file_name)[1].lower() if file_name else ""
    if "word" in mime_type.lower() or "doc" in mime_type.lower() or ext in ['.doc', '.docx']:
        try:
            doc_obj = docx.Document(io.BytesIO(file_bytes))
            text_list = [p.text for p in doc_obj.paragraphs if p.text.strip()]
            for table in doc_obj.tables:
                for row in table.rows:
                    row_txt = " | ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
                    if row_txt: text_list.append(row_txt)
            full_text = "\n".join(text_list)
            if full_text.strip():
                return [f"\n--- [Word Document: {file_name}] ---\n{full_text[:20000]}\n"]
        except Exception: pass

    if "pdf" in mime_type.lower():
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            for page in doc: text += page.get_text("text") + "\n"
            if len(text.strip()) > 100:
                return [f"\n--- [Document: {file_name}] ---\n{text[:20000]}\n"]
        except Exception: pass
    return [{"mime_type": mime_type, "data": file_bytes}]

# ==========================================
# 3. 데이터 파싱 및 가공 함수
# ==========================================
@st.cache_data(ttl=86400, show_spinner=False)
def get_live_esdc_median_wages():
    url = "https://www.canada.ca/en/employment-social-development/services/foreign-workers/median-wage.html"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            html = response.read().decode('utf-8', errors='ignore')
            wages = {}
            prov_regex = r'(Alberta|British Columbia|Manitoba|New Brunswick|Newfoundland and Labrador|Northwest Territories|Nova Scotia|Nunavut|Ontario|Prince Edward Island|Quebec|Saskatchewan|Yukon)'
            
            rows = re.findall(r'<tr\b[^>]*>(.*?)</tr>', html, re.DOTALL | re.IGNORECASE)
            for row in rows:
                p_match = re.search(prov_regex, row, re.IGNORECASE)
                if p_match:
                    p_name = p_match.group(1).title()
                    dollar_matches = re.findall(r'\$\s*(\d+(?:\.\d+)?)', row)
                    if dollar_matches:
                        latest_wage = float(dollar_matches[-1])
                        code_map = {
                            "British Columbia": "BC", "Alberta": "AB", "Ontario": "ON", "Saskatchewan": "SK",
                            "Manitoba": "MB", "New Brunswick": "NB", "Nova Scotia": "NS", "Prince Edward Island": "PE",
                            "Newfoundland And Labrador": "NL", "Yukon": "YT", "Northwest Territories": "NT",
                            "Nunavut": "NU", "Quebec": "QC"
                        }
                        if p_name in code_map:
                            wages[code_map[p_name]] = latest_wage
            if wages:
                return wages, "Canada.ca ESDC 실시간 데이터"
    except Exception:
        pass
        
    return {
        "BC": 38.40, "AB": 37.50, "ON": 36.92, "SK": 34.62, "MB": 31.33,
        "NB": 31.73, "NS": 31.96, "PE": 31.20, "NL": 33.60, "YT": 45.60,
        "NT": 48.00, "NU": 45.00, "QC": 36.00
    }, "ESDC 백업 기준"

def calculate_employment_term(wage_val, address_text):
    try:
        wage_match = re.search(r'(\d+(?:\.\d+)?)', str(wage_val))
        if not wage_match:
            return "3 years", 38.40, "기본값 적용"
        wage = float(wage_match.group(1))

        addr_upper = str(address_text).upper()
        detected_prov = "BC"
        
        prov_map = {
            "BC": ["BC", "BRITISH COLUMBIA"], "AB": ["AB", "ALBERTA"], "ON": ["ON", "ONTARIO"]
        }

        for code, keywords in prov_map.items():
            if any(re.search(r'\b' + re.escape(kw) + r'\b', addr_upper) for kw in keywords):
                detected_prov = code
                break

        live_wages, source_tag = get_live_esdc_median_wages()
        median_wage = live_wages.get(detected_prov, 38.40)

        if wage >= median_wage:
            return "3 years", median_wage, f"{detected_prov} 중위임금(${median_wage:.2f}) 이상 ➔ High-Wage Stream (3년 오퍼)"
        else:
            return "1 year", median_wage, f"{detected_prov} 중위임금(${median_wage:.2f}) 미만 ➔ Low-Wage Stream (1년 오퍼)"
    except Exception:
        return "3 years", 38.40, "기본값 적용"

def fetch_url_content(url):
    target_url = url.strip()
    if not target_url.startswith("http://") and not target_url.startswith("https://"):
        target_url = "https://" + target_url
    try:
        req = urllib.request.Request(target_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=12) as response:
            html = response.read().decode('utf-8', errors='ignore')
            text = re.sub(r'<script\b[^<]*(?:(?!</script>)<[^<]*)*</script>', '', html, flags=re.IGNORECASE)
            text = re.sub(r'<style\b[^<]*(?:(?!</style>)<[^<]*)*</style>', '', html, flags=re.IGNORECASE)
            text = re.sub(r'<[^>]+>', ' ', text)
            text = re.sub(r'\s+', ' ', text).strip()
            return text[:15000]
    except Exception:
        return ""

def get_provincial_overtime_clause(address_text):
    text = str(address_text).upper()
    if "BC" in text or "BRITISH COLUMBIA" in text:
        return ("1 and ½ times the employee’s regular rate of pay for hours in excess of 8 hours/day or 40 hours/week\n"
                "2 times the employee’s regular rate of pay for hours in excess of 12 hours/day")
    elif "AB" in text or "ALBERTA" in text:
        return "1.5 times the employee's regular rate of pay for hours in excess of 8 hours/day or 44 hours/week"
    elif "ON" in text or "ONTARIO" in text:
        return "1.5 times the employee's regular rate of pay for hours worked in excess of 44 hours per week"
    else:
        return "1.5 times the employee's regular rate of pay for hours worked over 8 hours/day or 40 hours/week"

# ==========================================
# 4. 순수 python-docx 양식 100% 동일 구현 엔진
# ==========================================
def generate_job_offer_docx(data, selected_style="Style A"):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    emp_name = data.get('employer_name', '')
    client_name = data.get('client_name', '')

    # 로고 이미지 삽입
    logo_bytes = data.get('logo_bytes')
    if logo_bytes:
        try:
            p_logo = doc.add_paragraph()
            p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_logo.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(2.0))
            doc.add_paragraph()
        except Exception:
            pass

    # 상단 날짜 및 수신자
    doc.add_paragraph(data.get('offer_date', datetime.date.today().strftime("%B %d, %Y")))
    
    p_dear = doc.add_paragraph()
    p_dear.add_run("Dear ")
    p_dear.add_run(f"{client_name},").bold = True

    # ----------------------------------------------------
    # Style A (Everfresh 스타일)
    # ----------------------------------------------------
    if "Style A" in selected_style:
        intro_p = doc.add_paragraph()
        intro_p.add_run("We are pleased to offer you a full-time position as ")
        intro_p.add_run(f"{data.get('job_title', '')}").bold = True
        intro_p.add_run(f" for a {data.get('employment_term', '3 years')} term with ")
        intro_p.add_run(f"{emp_name}").bold = True
        intro_p.add_run(" based on the following terms and conditions:")

        p = doc.add_paragraph(); p.add_run("Job Title: ").bold = True; p.add_run(f"{data.get('job_title', '')}")

        duties_list = data.get('job_duties', [])
        if isinstance(duties_list, str):
            duties_list = [d.strip() for d in duties_list.split('\n') if d.strip()]
        if duties_list:
            p_d = doc.add_paragraph(); p_d.add_run("Job Duties:").bold = True
            for duty in duties_list:
                clean_duty = re.sub(r'^[•\-\*]\s*', '', duty)
                doc.add_paragraph(clean_duty, style='List Bullet')

        p_comp = doc.add_paragraph()
        p_comp.add_run("Compensation:\n").bold = True
        p_comp.add_run("Hourly wage and hours\n")
        p_comp.add_run(f"The employee will be paid ${data.get('wage', '0.00')} per hour, based on a minimum of {data.get('hours', '30')} hours per week.")

        p = doc.add_paragraph(); p.add_run("Terms of Employment:\n").bold = True; p.add_run(f"This is a full-time, {data.get('employment_term', '3 years')} employment term starting from the date agreed upon by the employer and employee.")
        p = doc.add_paragraph(); p.add_run("Job Location:\n").bold = True; p.add_run(f"{data.get('job_location', '')}")
        p = doc.add_paragraph(); p.add_run("Benefits:\n").bold = True; p.add_run(f"{data.get('benefits', '4% vacation pay')}")
        
        p = doc.add_paragraph(); p.add_run("Confidentiality:\n").bold = True
        p.add_run(f"By accepting the terms of this offer, the employee agrees to keep all confidential information obtained during their employment with {emp_name} strictly confidential. The employee further agrees that, upon termination of employment for any reason, they will return all physical and digital property belonging to or originating from {emp_name} within five days of receiving notice of termination.")

        p = doc.add_paragraph(); p.add_run("Start Date:\n").bold = True; p.add_run(f"{data.get('start_date', 'As soon as possible upon obtaining a valid Work Permit')}")
        
        p = doc.add_paragraph(); p.add_run("Overtime:\n").bold = True; p.add_run(f"{data.get('overtime_clause', '')}")

        doc.add_paragraph(f"We are pleased to extend this offer of employment to you on behalf of {emp_name}. We are confident that you will make a valuable contribution to our company, and we look forward to working with you.")

        doc.add_paragraph("Sincerely,\t\t\t\t\tI accept the terms of this offer:")
        doc.add_paragraph("X\t\t\t\t\t\t______")

        p_sig = doc.add_paragraph()
        p_sig.add_run(f"{data.get('signer_name', '')}\n").bold = True
        p_sig.add_run(f"{data.get('signer_title', '')}\n{emp_name}\n")
        if data.get('employer_phone'): p_sig.add_run(f"T. {data.get('employer_phone')}\n")

        p_client_sig = doc.add_paragraph()
        p_client_sig.add_run(f"{client_name}\n").bold = True
        if data.get('client_dob'): p_client_sig.add_run(f"(DOB: {data.get('client_dob')})")

    # ----------------------------------------------------
    # Style B (Sushiwood 스타일)
    # ----------------------------------------------------
    else:
        intro_p = doc.add_paragraph()
        intro_p.add_run(f"{emp_name}").bold = True
        intro_p.add_run(" is pleased to offer you the position of ")
        intro_p.add_run(f"{data.get('job_title', '')}").bold = True
        intro_p.add_run(". We are confident that your skills will be a great asset to our business. The details of your employment are outlined below:")

        def add_kv(label, val):
            p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
            p.add_run(f"{label}: ").bold = True; p.add_run(str(val))

        add_kv("Job Title", data.get('job_title', ''))
        add_kv("Employment Type", "Full-time")
        add_kv("Terms of Employment", data.get('employment_term', '3 years'))
        add_kv("Start Date", data.get('start_date', 'As soon as possible upon obtaining a valid Work Permit'))
        add_kv("Hourly wage", f"{data.get('wage', '0.00')} CAD per hour")

        p_ot = doc.add_paragraph(); p_ot.paragraph_format.space_after = Pt(2)
        p_ot.add_run("Overtime:\n").bold = True
        p_ot.add_run(f"{data.get('overtime_clause', '')}")

        add_kv("Hours of Work", f"{data.get('hours', '30')} hours per week")
        add_kv("Work Location", data.get('job_location', ''))
        add_kv("Vacation Pay", data.get('benefits', '4% gross earnings'))

        duties_list = data.get('job_duties', [])
        if isinstance(duties_list, str):
            duties_list = [d.strip() for d in duties_list.split('\n') if d.strip()]
        if duties_list:
            p_d = doc.add_paragraph(); p_d.add_run("Job Duties:").bold = True
            for duty in duties_list:
                clean_duty = re.sub(r'^[•\-\*]\s*', '', duty)
                doc.add_paragraph(clean_duty, style='List Bullet')

        p_c = doc.add_paragraph(); p_c.add_run("Confidentiality Agreement:\n").bold = True
        p_c.add_run(f"By agreeing to the terms of this offer, the employee further agrees that they will hold all confidential information with which they are entrusted while employed by {emp_name} in strict confidence. The employee also agrees that upon termination of employment for any reason, they will return all physical and digital property which belongs to or originates at {emp_name} within 5 days of notice of termination of employment.")

        doc.add_paragraph("We look forward to your acceptance of this offer.\nPlease sign below to confirm your agreement.")
        doc.add_paragraph("Yours truly,\nX_____________________________________________")

        p_sig = doc.add_paragraph()
        p_sig.add_run(f"{data.get('signer_name', '')}\n").bold = True
        p_sig.add_run(f"{data.get('signer_title', '')}\n{emp_name}\n")
        if data.get('employer_address'): p_sig.add_run(f"{data.get('employer_address')}\n")
        if data.get('employer_phone'): p_sig.add_run(f"T. {data.get('employer_phone')}\n")
        if data.get('employer_email'): p_sig.add_run(f"E. {data.get('employer_email')}")

        doc.add_paragraph("\nI accept the terms of this offer:\n___________________________________________")

        p_acc = doc.add_paragraph()
        p_acc.add_run(f"{client_name}\n").bold = True
        if data.get('client_dob'): p_acc.add_run(f"(DOB: {data.get('client_dob')})")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

def parse_existing_job_offer(file_bytes, mime_type):
    prompt = """
    Analyze this Job Offer document carefully. Extract into valid JSON:
    - client_name
    - client_dob
    - employer_name
    - signer_name
    - signer_title
    - employer_address
    - employer_phone
    - employer_email
    - job_title
    - wage
    - hours
    - job_location
    - benefits
    - job_duties (Array of strings)
    """
    contents = prepare_document_for_gemini(file_bytes, mime_type, "Existing_Job_Offer.pdf")
    contents.insert(0, prompt)
    try:
        response = safe_generate_content(contents)
        return json.loads(response.text.strip().replace('```json', '').replace('```', ''))
    except Exception: return {}

# ==========================================
# 5. Streamlit 메인 UI
# ==========================================
st.title("📄 잡오퍼 DOCX 생성기 (DEV)")
st.caption("실시간 ESDC Median Wage 파싱 및 사내 표준 양식 기반 잡오퍼(.docx) 자동 생성기입니다.")

if "job_offer_data" not in st.session_state:
    st.session_state.job_offer_data = {}

doc_mode = st.radio("작성 모드를 선택하세요", ["🆕 신규 잡오퍼 생성", "🔄 기존 잡오퍼 연장/업데이트"])

st.markdown("---")
st.subheader("1. 손님 정보 (여권 업로드)")
passport_file = st.file_uploader("손님 여권 이미지 또는 PDF", type=['jpg', 'jpeg', 'png', 'pdf', 'heic'], key="jo_passport")

if doc_mode == "🔄 기존 잡오퍼 연장/업데이트":
    st.subheader("2. 기존 잡오퍼 서류")
    old_jo_file = st.file_uploader("기존 잡오퍼 (DOCX 또는 PDF)", type=['pdf', 'docx', 'doc'], key="jo_old_file")
    if old_jo_file and st.button("기존 잡오퍼 분석하여 정보 가져오기"):
        with st.spinner("기존 잡오퍼 분석 중..."):
            existing_parsed = parse_existing_job_offer(old_jo_file.getvalue(), old_jo_file.type if old_jo_file.type else "application/pdf")
            if existing_parsed:
                st.session_state.job_offer_data.update(existing_parsed)
                st.success("기존 잡오퍼 정보를 불러왔습니다.")

st.subheader("3. 채용 공고")
job_posting_text = st.text_area("채용 공고 링크(URL) 또는 텍스트", height=100)
logo_file = st.file_uploader("회사 로고 이미지 (선택 사항)", type=['jpg', 'jpeg', 'png'], key="jo_logo")

if st.button("AI 채용공고 & 여권 실시간 분석 시작", type="primary", use_container_width=True):
    with st.spinner("정보 추출 및 분석 중..."):
        extracted_info = {}
        if logo_file: extracted_info['logo_bytes'] = logo_file.getvalue()

        if passport_file:
            pass_img = process_uploaded_file_to_image(passport_file)
            pass_data = extract_imm5476_info(pass_img)
            if pass_data:
                extracted_info['client_name'] = format_full_name(pass_data.get('surname', ''), pass_data.get('given_name', ''))
                extracted_info['client_dob'] = pass_data.get('dob', '')
        
        if job_posting_text.strip():
            raw_input = job_posting_text.strip()
            is_url = bool(re.search(r'https?://[^\s]+|www\.[^\s]+', raw_input))
            
            if is_url:
                url_match = re.search(r'(https?://[^\s]+|www\.[^\s]+)', raw_input).group(0)
                fetched_text = fetch_url_content(url_match)
                text_to_analyze = fetched_text if fetched_text else raw_input
            else: 
                text_to_analyze = raw_input
                
            prompt_job = f"""
            Analyze this job posting:
            {text_to_analyze}

            Extract into JSON:
            - employer_name
            - job_title
            - wage (numeric string)
            - hours
            - job_location
            - employer_address
            - employer_phone
            - employer_email
            - benefits
            - job_duties (Array of strings)
            Return ONLY raw valid JSON object.
            """
            try:
                resp = safe_generate_content([prompt_job])
                clean = resp.text.strip().replace('```json', '').replace('```', '')
                extracted_info.update(json.loads(clean))
            except Exception as e: 
                st.warning(f"분석 중 오류 발생: {e}")
                
        st.session_state.job_offer_data.update(extracted_info)
        st.success("실시간 정보 수집 및 분석이 완료되었습니다.")

st.markdown("---")
st.subheader("4. 최종 잡오퍼 정보 확인 및 수정")

jo_data = st.session_state.job_offer_data
temp_wage = jo_data.get('wage', '20.15')
temp_loc = jo_data.get('job_location', jo_data.get('employer_address', ''))
calc_term, calc_median, calc_reason = calculate_employment_term(temp_wage, temp_loc)

col_c1, col_c2 = st.columns(2)
with col_c1:
    c_name = st.text_input("손님 영문 성명", value=jo_data.get('client_name', ''))
    offer_dt = st.date_input("오퍼 작성일", datetime.date.today()).strftime("%B %d, %Y")
    term_str = st.text_input("계약 기간", value=jo_data.get('employment_term', calc_term))
    st.info(f"🌐 **ESDC 중위 임금 계산:** {calc_reason}")
with col_c2:
    c_dob = st.text_input("손님 생년월일", value=jo_data.get('client_dob', ''))
    start_dt_str = st.text_input("근무 시작일", value=jo_data.get('start_date', 'As soon as possible upon obtaining a valid Work Permit'))

col_e1, col_e2 = st.columns(2)
with col_e1:
    emp_name = st.text_input("회사명", value=jo_data.get('employer_name', ''))
    signer_n = st.text_input("대표자 성명", value=jo_data.get('signer_name', ''))
    signer_t = st.text_input("대표자 직책", value=jo_data.get('signer_title', 'Owner'))
with col_e2:
    emp_addr = st.text_input("회사 주소", value=jo_data.get('employer_address', ''))
    emp_phone = st.text_input("회사 전화번호", value=jo_data.get('employer_phone', ''))
    emp_email = st.text_input("회사 이메일", value=jo_data.get('employer_email', ''))

selected_layout = st.selectbox("잡오퍼 문서 양식", ["Style A (서두 서술형 / 하단 개별 서명란)", "Style B (상단 조건 리스트형 / 서명 블록)"])

col_j1, col_j2 = st.columns(2)
with col_j1:
    j_title = st.text_input("직책", value=jo_data.get('job_title', ''))
    j_wage = st.text_input("시급", value=str(jo_data.get('wage', '20.15')))
    j_hours = st.text_input("주당 근무시간", value=str(jo_data.get('hours', '30')))
with col_j2:
    j_loc = st.text_input("근무지 주소", value=jo_data.get('job_location', emp_addr))
    j_benefits = st.text_input("혜택", value=jo_data.get('benefits', '4% vacation pay'))

auto_ot_clause = get_provincial_overtime_clause(j_loc if j_loc else emp_addr)
j_ot = st.text_area("오버타임 조항", value=auto_ot_clause, height=80)

duties_input_str = jo_data.get('job_duties', [])
if isinstance(duties_input_str, list): duties_input_str = "\n".join(duties_input_str)
j_duties_text = st.text_area("주요 직무", value=duties_input_str, height=150)

st.markdown("---")
if st.button("📄 MS Word (.docx) 다운로드", type="primary", use_container_width=True):
    if not c_name or not emp_name or not j_title:
        st.error("손님 성명, 회사명, 직책은 필수 입력 항목입니다.")
    else:
        final_jo_dict = {
            "client_name": c_name, "client_dob": c_dob, "offer_date": offer_dt, "employment_term": term_str,
            "start_date": start_dt_str, "employer_name": emp_name, "signer_name": signer_n, "signer_title": signer_t,
            "employer_address": emp_addr, "employer_phone": emp_phone, "employer_email": emp_email,
            "job_title": j_title, "wage": j_wage, "hours": j_hours, "job_location": j_loc,
            "benefits": j_benefits, "overtime_clause": j_ot, "job_duties": j_duties_text,
            "logo_bytes": jo_data.get('logo_bytes')
        }
        
        docx_bytes = generate_job_offer_docx(final_jo_dict, selected_style=selected_layout)
        out_filename = f"[Job Offer]_{c_name.strip().split()[0].capitalize() if c_name.strip() else 'Client'}.docx"
        
        st.success("문서 생성이 완료되었습니다!")
        st.download_button("📥 파일 다운로드", data=docx_bytes, file_name=out_filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
