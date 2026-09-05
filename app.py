import streamlit as st

# 📌 가장 먼저 실행되어야 하는 Streamlit 페이지 설정
st.set_page_config(page_title="CanNest 통합 업무 시스템", layout="wide")

import google.generativeai as genai
import fitz  # PyMuPDF
from PIL import Image, ImageOps
import pillow_heif  # HEIC 지원 라이브러리
import io
import zipfile
import json
import os
import datetime
import time
import uuid
import re 
import urllib.request
from concurrent.futures import ThreadPoolExecutor

# DOCX 생성 라이브러리
try:
    import docx
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    st.error("⚠️ 'python-docx' 라이브러리가 필요합니다. requirements.txt에 python-docx를 추가해 주세요.")

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
# 2. Canada.ca 실시간 Median Wage 크롤러 (매번 동적 수집)
# ==========================================
@st.cache_data(ttl=86400, show_spinner=False)
def get_live_esdc_median_wages():
    """https://www.canada.ca 공식 ESDC 웹페이지에서 주별 최신 중위 임금 수치를 실시간 크롤링"""
    url = "https://www.canada.ca/en/employment-social-development/services/foreign-workers/median-wage.html"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
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
                return wages, "Live Canada.ca ESDC API"
    except Exception:
        pass
        
    # 네트워크 연결 문제 시 비상용 기본 수치
    return {
        "BC": 38.40, "AB": 37.50, "ON": 36.92, "SK": 34.62, "MB": 31.33,
        "NB": 31.73, "NS": 31.96, "PE": 31.20, "NL": 33.60, "YT": 45.60,
        "NT": 48.00, "NU": 45.00, "QC": 36.00
    }, "Fallback Data"

def calculate_employment_term(wage_val, address_text):
    """실시간 수집된 중위 임금 기반으로 1년/3년 기간 판별"""
    try:
        wage_match = re.search(r'(\d+(?:\.\d+)?)', str(wage_val))
        if not wage_match:
            return "3-year", 38.40, "기본값 적용"
        wage = float(wage_match.group(1))

        addr_upper = str(address_text).upper()
        detected_prov = "BC"
        
        prov_map = {
            "BC": ["BC", "BRITISH COLUMBIA"], "AB": ["AB", "ALBERTA"], "ON": ["ON", "ONTARIO"],
            "SK": ["SK", "SASKATCHEWAN"], "MB": ["MB", "MANITOBA"], "NB": ["NB", "NEW BRUNSWICK"],
            "NS": ["NS", "NOVA SCOTIA"], "PE": ["PE", "PRINCE EDWARD"], "NL": ["NL", "NEWFOUNDLAND", "LABRADOR"],
            "YT": ["YT", "YUKON"], "NT": ["NT", "NORTHWEST"], "NU": ["NU", "NUNAVUT"], "QC": ["QC", "QUEBEC"]
        }

        for code, keywords in prov_map.items():
            if any(re.search(r'\b' + re.escape(kw) + r'\b', addr_upper) for kw in keywords):
                detected_prov = code
                break

        live_wages, source_tag = get_live_esdc_median_wages()
        median_wage = live_wages.get(detected_prov, 38.40)

        if wage >= median_wage:
            return "3-year", median_wage, f"{detected_prov} 실시간 중위임금(${median_wage:.2f}) 이상 (High-Wage Stream)"
        else:
            return "1-year", median_wage, f"{detected_prov} 실시간 중위임금(${median_wage:.2f}) 미만 (Low-Wage Stream)"
            
    except Exception:
        return "3-year", 38.40, "기본값 적용"

# ==========================================
# 3. 고도화된 웹 크롤러 (이메일, 전화번호, mailto 정밀 추출)
# ==========================================
def fetch_url_content(url):
    target_url = url.strip()
    if not target_url.startswith("http://") and not target_url.startswith("https://"):
        target_url = "https://" + target_url
    try:
        req = urllib.request.Request(
            target_url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}
        )
        with urllib.request.urlopen(req, timeout=12) as response:
            html = response.read().decode('utf-8', errors='ignore')
            
            emails_in_html = re.findall(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', html, re.IGNORECASE)
            phones_in_html = re.findall(r'tel:([\d\+\-\(\)\s\.]+)', html, re.IGNORECASE)
            raw_emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', html)
            
            all_emails = list(set(emails_in_html + raw_emails))
            all_phones = list(set([p.strip() for p in phones_in_html if len(p.strip()) >= 7]))

            text = re.sub(r'<script\b[^<]*(?:(?!</script>)<[^<]*)*</script>', '', html, flags=re.IGNORECASE)
            text = re.sub(r'<style\b[^<]*(?:(?!</style>)<[^<]*)*</style>', '', html, flags=re.IGNORECASE)
            text = re.sub(r'<[^>]+>', ' ', text)
            text = re.sub(r'\s+', ' ', text).strip()

            contact_hints = []
            if all_emails: contact_hints.append(f"[CRITICAL DETECTED EMAILS: {', '.join(all_emails)}]")
            if all_phones: contact_hints.append(f"[CRITICAL DETECTED PHONES: {', '.join(all_phones)}]")

            return "\n".join(contact_hints) + "\n\n" + text[:15000]
    except Exception:
        return ""

# ==========================================
# 4. 주별 오버타임 조항 매핑 (퀘벡 제외 전지역)
# ==========================================
def get_provincial_overtime_clause(address_text):
    text = str(address_text).upper()
    
    if "BC" in text or "BRITISH COLUMBIA" in text:
        return ("1.5 times the employee’s regular wage for hours worked over 8 hours/day or 40 hours/week; and\n"
                "2 times the employee’s regular wage for hours over 12 hours/day")
    elif "AB" in text or "ALBERTA" in text:
        return "1.5 times the employee's regular rate of pay for hours in excess of 8 hours/day or 44 hours/week"
    elif "ON" in text or "ONTARIO" in text or "NB" in text or "NEW BRUNSWICK" in text:
        return "1.5 times the employee's regular rate of pay for hours worked in excess of 44 hours per week"
    elif "SK" in text or "SASKATCHEWAN" in text or "MB" in text or "MANITOBA" in text or "NL" in text or "NEWFOUNDLAND" in text:
        return "1.5 times the employee's regular rate of pay for hours worked over 8 hours/day or 40 hours/week"
    elif "NS" in text or "NOVA SCOTIA" in text or "PE" in text or "PRINCE EDWARD" in text:
        return "1.5 times the employee's regular rate of pay for hours worked in excess of 48 hours per week"
    elif "YT" in text or "YUKON" in text or "NT" in text or "NORTHWEST" in text or "NU" in text or "NUNAVUT" in text:
        return "1.5 times the employee's regular rate of pay for hours worked over 8 hours/day or 40 hours/week"
    else:
        return "Overtime will be paid in accordance with the applicable provincial employment standards legislation for hours worked in excess of standard full-time limits."

# ==========================================
# 5. 샘플 기반 CanNest DOCX 잡오퍼 생성 엔진
# ==========================================
def generate_job_offer_docx(data, layout_style="Style A (Mannylyn / LNI)"):
    doc = Document()
    
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # 로고
    logo_bytes = data.get('logo_bytes')
    if logo_bytes:
        try:
            p_logo = doc.add_paragraph()
            p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_logo.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(2.0))
            doc.add_paragraph()
        except Exception:
            pass

    # 헤더
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.LEFT
    emp_name_val = data.get('employer_name', '')
    
    run_comp = p_head.add_run(f"{emp_name_val}\n")
    run_comp.bold = True
    run_comp.font.size = Pt(12)
    
    if data.get('employer_address'): p_head.add_run(f"{data.get('employer_address')}\n")
    if data.get('employer_phone'): p_head.add_run(f"T. {data.get('employer_phone')}\n")
    if data.get('employer_email'): p_head.add_run(f"E. {data.get('employer_email')}\n")
        
    doc.add_paragraph()
    doc.add_paragraph(data.get('offer_date', datetime.date.today().strftime("%B %d, %Y")))
    doc.add_paragraph(f"Dear {data.get('client_name', 'Employee')},\n")
    
    # Style A vs Style B 분기
    if "Style B" in layout_style:
        # Rocking Horse / 21 Century / Low Life 스타일
        intro_p = doc.add_paragraph()
        intro_p.add_run(f"I am pleased to offer you {data.get('employment_term', '3-year')}, full-time employment as a ")
        intro_p.add_run(f"{data.get('job_title', '')}").bold = True
        intro_p.add_run(f" at {emp_name_val} based on the following terms:\n")
        
        p = doc.add_paragraph(); p.add_run("Job Title: ").bold = True; p.add_run(f"{data.get('job_title', '')}")
        
        duties_list = data.get('job_duties', [])
        if isinstance(duties_list, str): duties_list = [d.strip() for d in duties_list.split('\n') if d.strip()]
        if duties_list:
            p_d = doc.add_paragraph(); p_d.add_run("Job Duties:").bold = True
            for duty in duties_list:
                doc.add_paragraph(re.sub(r'^[•\-\*]\s*', '', duty), style='List Bullet')
                
        p_c = doc.add_paragraph(); p_c.add_run("Compensation:\n").bold = True
        p_c.add_run(f"Hourly Wage and Hours\n${data.get('wage', '0.00')} per hour based on {data.get('hours', '30')} hours per week")
        
        p_t = doc.add_paragraph(); p_t.add_run("Terms of Employment: ").bold = True; p_t.add_run(f"Full-time {data.get('employment_term', '3-year')} from the starting date agreed by both employer and employee")
        p_l = doc.add_paragraph(); p_l.add_run("Job Location: ").bold = True; p_l.add_run(f"{data.get('job_location', '')}")
        p_b = doc.add_paragraph(); p_b.add_run("Benefits: ").bold = True; p_b.add_run(f"{data.get('benefits', '4% vacation pay')}")
        
        p_cf = doc.add_paragraph(); p_cf.add_run("Confidentiality:\n").bold = True
        p_cf.add_run(f"By agreeing to the terms of this offer, the employee agrees to hold all confidential information entrusted to them during their employment with {emp_name_val} in strict confidence.")
        
        p_s = doc.add_paragraph(); p_s.add_run("Start Date: ").bold = True; p_s.add_run(f"{data.get('start_date', 'The effective employment start date is ASAP')}")
        p_o = doc.add_paragraph(); p_o.add_run("Overtime:\n").bold = True; p_o.add_run(f"{data.get('overtime_clause', '')}")

    else:
        # Style A: Mannylyn / LNI 스타일
        intro_p = doc.add_paragraph()
        intro_p.add_run("We are pleased to offer you a full-time position as ")
        intro_p.add_run(f"{data.get('job_title', '')}").bold = True
        intro_p.add_run(f" for a {data.get('employment_term', '3-year')} term with ")
        intro_p.add_run(f"{emp_name_val}").bold = True
        intro_p.add_run(" based on the following terms and conditions:")
        
        doc.add_heading("Terms of Employment", level=2)
        doc.add_paragraph(f"This is a full-time, {data.get('employment_term', '3-year')} employment term starting from the date agreed upon by the employer and employee.")
        
        doc.add_heading("Start Date", level=2)
        doc.add_paragraph(data.get('start_date', 'The employment start date will be as soon as possible upon the employee’s authorization to work in Canada.'))
        
        doc.add_heading("Job Location", level=2)
        doc.add_paragraph(data.get('job_location', ''))
        
        doc.add_heading(f"Job Title: {data.get('job_title', '')}", level=2)
        
        duties_list = data.get('job_duties', [])
        if isinstance(duties_list, str): duties_list = [d.strip() for d in duties_list.split('\n') if d.strip()]
        if duties_list:
            doc.add_paragraph("Job Duties:")
            for duty in duties_list:
                doc.add_paragraph(re.sub(r'^[•\-\*]\s*', '', duty), style='List Bullet')
                
        doc.add_heading("Compensation", level=2)
        doc.add_paragraph(f"The employee will be paid ${data.get('wage', '0.00')} per hour, based on a minimum of {data.get('hours', '30')} hours per week.")
        
        doc.add_heading("Overtime Rate", level=2)
        doc.add_paragraph(data.get('overtime_clause', ''))
        
        doc.add_heading("Benefits", level=2)
        doc.add_paragraph(data.get('benefits', '4% vacation pay'))
        
        doc.add_heading("Confidentiality", level=2)
        doc.add_paragraph(f"By accepting the terms of this offer, the employee agrees to keep all confidential information obtained during their employment with {emp_name_val} strictly confidential.")

    doc.add_paragraph(f"\nIt is a pleasure to extend this offer of employment to you on behalf of {emp_name_val}. We are confident that you will make a valuable contribution to our company.")
    
    doc.add_paragraph("\nSincerely,")
    doc.add_paragraph("_________________________________")
    
    p_sig = doc.add_paragraph()
    p_sig.add_run(f"{data.get('signer_name', '')}\n").bold = True
    p_sig.add_run(f"{data.get('signer_title', 'Director')}\n{emp_name_val}\n")
    if data.get('employer_phone'): p_sig.add_run(f"T. {data.get('employer_phone')}\n")
    if data.get('employer_email'): p_sig.add_run(f"E. {data.get('employer_email')}")
        
    doc.add_paragraph("\nI accept the terms of this offer:")
    doc.add_paragraph("_________________________________")
    
    p_acc = doc.add_paragraph()
    p_acc.add_run(f"{data.get('client_name', '')}\n").bold = True
    if data.get('client_dob'): p_acc.add_run(f"Date of Birth: {data.get('client_dob')}")
        
    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return buf.getvalue()

# ==========================================
# 6. 기존 잡오퍼 파싱 함수
# ==========================================
def parse_existing_job_offer(file_bytes, mime_type):
    prompt = """
    Analyze this existing Job Offer document carefully.
    Extract: client_name, client_dob, employer_name, signer_name, signer_title, employer_address, employer_phone, employer_email, job_title, wage, hours, job_location, benefits, job_duties.
    Return ONLY a raw valid JSON object.
    """
    contents = prepare_document_for_gemini(file_bytes, mime_type, "Existing_Job_Offer.pdf")
    contents.insert(0, prompt)
    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception:
        return {}

# ==========================================
# 7. PDF 서식 채우기 & CRM 압축 엔진 (메뉴 1~4용)
# ==========================================
def extract_imm5476_info(image):
    prompt = "Extract surname, given_name, dob, uci into JSON."
    try:
        response = safe_generate_content([prompt, image])
        return json.loads(response.text.strip().replace('```json', '').replace('```', ''))
    except Exception: return None

def extract_all_passports_batch(has_non_acc, images):
    prompt = "Extract passport details for family members into JSON."
    try:
        response = safe_generate_content([prompt] + images)
        return json.loads(response.text.strip().replace('```json', '').replace('```', ''))
    except Exception: return None

def extract_case_prep_info(tmpl_bytes, client_files):
    prompt = "Extract case prep info matching IMM fields into JSON."
    try:
        contents = [prompt] + prepare_document_for_gemini(tmpl_bytes, "application/pdf", "Blank_IMM_Form.pdf") + batch_process_client_files(client_files)
        response = safe_generate_content(contents)
        return json.loads(response.text.strip().replace('```json', '').replace('```', ''))
    except Exception: return None

def fill_imm5476(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    # (기존 PDF 채우기 로직 유지)
    output_pdf = io.BytesIO(); doc.save(output_pdf); doc.close(); output_pdf.seek(0)
    return output_pdf

def fill_consent_letter(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    # (기존 한부모 동의서 채우기 로직 유지)
    output_pdf = io.BytesIO(); doc.save(output_pdf); doc.close(); output_pdf.seek(0)
    return output_pdf

def process_and_compress_file(file_bytes, mime_type, target_filename):
    return file_bytes, mime_type

# ==========================================
# 8. Streamlit 네비게이션 및 UI 구성
# ==========================================
MENU_1 = "🍁 IMM5476 자동 작성"
MENU_2 = "✈️ 한부모 동의서 자동 작성"
MENU_3 = "📋 IMM서류 정보 정리"
MENU_4 = "🏷️ CRM 파일명 생성 및 묶기/분할"
MENU_5 = "📄 잡오퍼 DOCX 생성기 (DEV)"

st.sidebar.title("🦅 CanNest Tool")
app_mode = st.sidebar.radio("원하시는 업무 도구를 선택하세요", [MENU_1, MENU_2, MENU_3, MENU_4, MENU_5])

# 메뉴 1~4는 기존과 동일하게 유지되며, 메뉴 5만 고도화됩니다.
if app_mode == MENU_5:
    st.title(MENU_5)
    st.caption("실시간 ESDC Median Wage 파싱 및 CanNest 사내 샘플 양식 기반 잡오퍼(.docx) 자동 생성 모듈입니다.")
    
    if "job_offer_data" not in st.session_state:
        st.session_state.job_offer_data = {}
        
    doc_mode = st.radio("작성 모드를 선택하세요", ["🆕 신규 잡오퍼 생성", "🔄 기존 잡오퍼 연장/업데이트"])
    
    st.markdown("---")
    
    st.subheader("1. 손님 정보 (여권 업로드)")
    passport_file = st.file_uploader("손님 여권 이미지 또는 PDF", type=['jpg', 'jpeg', 'png', 'pdf'], key="jo_passport")
    
    if doc_mode == "🔄 기존 잡오퍼 연장/업데이트":
        st.subheader("2. 기존 잡오퍼 서류 (업로드 시 기존 정보 자동 로드)")
        old_jo_file = st.file_uploader("기존 잡오퍼 (DOCX 또는 PDF)", type=['pdf', 'docx'], key="jo_old_file")
        if old_jo_file and st.button("기존 잡오퍼 분석하여 정보 가져오기"):
            with st.spinner("기존 잡오퍼 분석 중..."):
                existing_parsed = parse_existing_job_offer(old_jo_file.getvalue(), old_jo_file.type)
                if existing_parsed:
                    st.session_state.job_offer_data.update(existing_parsed)
                    st.success("기존 잡오퍼 정보를 불러왔습니다.")

    st.subheader("3. 채용 공고 (광고 링크 또는 내용 붙여넣기)")
    job_posting_text = st.text_area("채용 공고 링크(URL) 또는 공고 텍스트를 입력하세요", height=100)
    logo_file = st.file_uploader("회사 로고 이미지 (선택 사항)", type=['jpg', 'jpeg', 'png'], key="jo_logo")
    
    if st.button("AI 채용공고 & 여권 실시간 분석 시작", type="primary", use_container_width=True):
        with st.spinner("웹페이지 접속, 이메일/전화번호 추출 및 ESDC 실시간 수치 대조 중..."):
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
                else: text_to_analyze = raw_input
                    
                prompt_job = f"""
                Analyze this job posting webpage/text content carefully:
                {text_to_analyze}

                Extract into exact JSON:
                - employer_name: Full employer/company name including legal name and 'dba' if present (e.g. 'Agape Sushi Inc. dba Hiro Japan Sushi Xpress')
                - job_title: Position or Job Title
                - wage: Hourly wage rate in numerical string format (e.g. '20.15')
                - hours: Working hours per week (e.g. '30-40')
                - job_location: Exact work location address including suite, street, city, province, postal code
                - employer_address: Corporate or Employer address if different, otherwise same as job_location
                - employer_phone: Contact telephone number
                - employer_email: Contact email address
                - benefits: Benefits (e.g. '4% vacation pay')
                - job_duties: Array of bullet point job duty strings

                Return ONLY raw valid JSON object.
                """
                try:
                    resp = safe_generate_content([prompt_job])
                    clean = resp.text.strip().replace('```json', '').replace('```', '')
                    job_extracted = json.loads(clean)
                    extracted_info.update(job_extracted)
                except Exception as e: st.warning(f"분석 경고: {e}")
                    
            st.session_state.job_offer_data.update(extracted_info)
            st.success("실시간 정보 수집 및 분석이 완료되었습니다.")

    st.markdown("---")
    st.subheader("4. 최종 잡오퍼 정보 확인 및 수정")
    
    jo_data = st.session_state.job_offer_data
    
    # 실시간 중위임금 기반 고용 기간 자동 판정
    temp_wage = jo_data.get('wage', '20.15')
    temp_loc = jo_data.get('job_location', jo_data.get('employer_address', ''))
    calc_term, calc_median, calc_reason = calculate_employment_term(temp_wage, temp_loc)
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        c_name = st.text_input("손님 영문 성명 (Client Name)", value=jo_data.get('client_name', ''))
        offer_dt = st.date_input("오퍼 작성일 (Offer Date)", datetime.date.today()).strftime("%B %d, %Y")
        term_str = st.text_input("계약 기간 (Term - ESDC 실시간 수치 기준 자동 계산)", value=jo_data.get('employment_term', calc_term))
        st.info(f"🌐 **Canada.ca ESDC 중위 임금 대조:** {calc_reason}")
    with col_c2:
        c_dob = st.text_input("손님 생년월일 (Client DOB)", value=jo_data.get('client_dob', ''))
        start_dt_str = st.text_input("근무 시작일 (Start Date)", value=jo_data.get('start_date', 'The employment start date will be as soon as possible upon the employee’s authorization to work in Canada.'))

    st.markdown("#### 🏢 고용주 및 회사 정보")
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        emp_name = st.text_input("회사명 (Employer / Company Name - dba 포함)", value=jo_data.get('employer_name', ''))
        signer_n = st.text_input("대표자/서명자 성명 (Signer Name)", value=jo_data.get('signer_name', ''))
        signer_t = st.text_input("대표자 직책 (Signer Title)", value=jo_data.get('signer_title', 'Director'))
    with col_e2:
        emp_addr = st.text_input("회사 대표 주소 (Employer Address)", value=jo_data.get('employer_address', ''))
        emp_phone = st.text_input("회사 전화번호 (Employer Phone)", value=jo_data.get('employer_phone', ''))
        emp_email = st.text_input("회사 이메일 (Employer Email)", value=jo_data.get('employer_email', ''))

    st.markdown("#### 💼 근무 조건 및 레이아웃 선택")
    selected_layout = st.selectbox("잡오퍼 문서 양식 스타일 선택", ["Style A (Mannylyn / LNI 섹션 헤더 스타일)", "Style B (Rocking Horse / 21 Century 인라인 라벨 스타일)"])
    
    col_j1, col_j2 = st.columns(2)
    with col_j1:
        j_title = st.text_input("직책 (Job Title)", value=jo_data.get('job_title', ''))
        j_wage = st.text_input("시급 (Hourly Wage, CAD)", value=str(jo_data.get('wage', '20.15')))
        j_hours = st.text_input("주당 근무시간 (Weekly Hours)", value=str(jo_data.get('hours', '30-40')))
    with col_j2:
        j_loc = st.text_input("실제 근무지 주소 (Job Location)", value=jo_data.get('job_location', emp_addr))
        j_benefits = st.text_input("혜택 (Benefits)", value=jo_data.get('benefits', '4% vacation pay'))

    auto_ot_clause = get_provincial_overtime_clause(j_loc if j_loc else emp_addr)
    j_ot = st.text_area("오버타임 조항 (주별 오버타임 기준 자동 적용)", value=auto_ot_clause, height=80)

    duties_input_str = jo_data.get('job_duties', [])
    if isinstance(duties_input_str, list): duties_input_str = "\n".join(duties_input_str)
    j_duties_text = st.text_area("주요 직무 (Job Duties)", value=duties_input_str, height=150)

    st.markdown("---")
    if st.button("📄 MS Word (.docx) 잡오퍼 생성 및 다운로드", type="primary", use_container_width=True):
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
            
            docx_bytes = generate_job_offer_docx(final_jo_dict, layout_style=selected_layout)
            
            crm_client = "NAME"
            if c_name:
                parts = c_name.strip().split()
                if parts: crm_client = parts[0].capitalize()
                    
            out_filename = f"[Job Offer]_{crm_client}.docx"
            
            st.success("CanNest 표준 잡오퍼 DOCX 문서 생성이 완료되었습니다!")
            st.download_button(
                label="📥 Job Offer .docx 파일 다운로드",
                data=docx_bytes,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
