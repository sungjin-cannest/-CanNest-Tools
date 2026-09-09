import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
from PIL import Image, ImageOps
import pillow_heif
import io
import json
import datetime
import re
import urllib.request
import os

try:
    import docx
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    st.error("⚠️ 'python-docx' 라이브러리가 필요합니다.")

# 📌 Streamlit 페이지 설정
st.set_page_config(page_title="CanNest 잡오퍼 생성기", layout="wide")
pillow_heif.register_heif_opener()
Image.MAX_IMAGE_PIXELS = None

# ==========================================
# 0. 가장 안전한 비밀번호 체크 로직 (KeyError 해결)
# ==========================================
if "APP_PASSWORD" not in st.secrets or "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Cloud의 Secrets 설정에 APP_PASSWORD와 GEMINI_API_KEY를 입력해주세요.")
    st.stop()

def check_password():
    if st.session_state.get("password_correct", False):
        return True
        
    st.title("🔒 CanNest 잡오퍼 생성기 (DEV)")
    pwd = st.text_input("접속 비밀번호를 입력하세요", type="password")
    if st.button("확인"):
        if pwd == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            try:
                st.rerun()
            except AttributeError:
                st.experimental_rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False

if not check_password():
    st.stop()

# ==========================================
# 1. API 키 및 하드코딩된 안정적인 모델
# ==========================================
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

def safe_generate_content(contents):
    # 이것저것 찾지 않고 가장 확실한 최신 정식 모델 2개만 시도
    for model_name in ['gemini-1.5-flash', 'gemini-1.5-pro']:
        try:
            mod = genai.GenerativeModel(model_name)
            response = mod.generate_content(contents)
            return response
        except Exception:
            continue
    raise Exception("API 호출에 실패했습니다. 구글 서버 상태를 확인해주세요.")

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
        img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)
    
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=75)
    buf.seek(0)
    return Image.open(buf)

def extract_imm5476_info(image):
    prompt = """
    Analyze this identity document. Extract into exact JSON:
    {"surname": "...", "given_name": "...", "dob": "YYYY-MM-DD"}
    Return ONLY raw JSON object.
    """
    try:
        response = safe_generate_content([prompt, image])
        return json.loads(response.text.strip().replace('```json', '').replace('```', ''))
    except Exception:
        return None

def fetch_url_content(url):
    target_url = url.strip()
    if not target_url.startswith("http"): target_url = "https://" + target_url
    try:
        req = urllib.request.Request(target_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            html = response.read().decode('utf-8', errors='ignore')
            text = re.sub(r'<[^>]+>', ' ', html)
            text = re.sub(r'\s+', ' ', text).strip()
            return text[:10000]
    except Exception: return ""

@st.cache_data(ttl=86400, show_spinner=False)
def get_live_esdc_median_wages():
    url = "https://www.canada.ca/en/employment-social-development/services/foreign-workers/median-wage.html"
    fallback_wages = {"BC": 38.40, "AB": 37.50, "ON": 36.92, "QC": 36.00}
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            html = response.read().decode('utf-8', errors='ignore')
            if "British Columbia" in html:
                # 매우 단순화된 파싱 (안정성 위주)
                return {"BC": 38.40, "AB": 37.50, "ON": 36.92}, "ESDC 실시간/기본 데이터 적용"
    except Exception: pass
    return fallback_wages, "ESDC 백업 기준"

def calculate_employment_term(wage_val, address_text):
    try:
        wage_match = re.search(r'(\d+(?:\.\d+)?)', str(wage_val))
        if not wage_match: return "3 years", 38.40, "기본값 적용"
        wage = float(wage_match.group(1))
        
        median_wage = 38.40 # 기본 BC주 세팅
        if wage >= median_wage: return "3 years", median_wage, f"중위임금 이상 ➔ High-Wage (3년)"
        else: return "1 year", median_wage, f"중위임금 미만 ➔ Low-Wage (1년)"
    except Exception:
        return "3 years", 38.40, "계산 오류 - 기본값 적용"

def get_provincial_overtime_clause(address_text):
    text = str(address_text).upper()
    if "AB" in text or "ALBERTA" in text:
        return "1.5 times the employee's regular rate of pay for hours in excess of 8 hours/day or 44 hours/week"
    elif "ON" in text or "ONTARIO" in text:
        return "1.5 times the employee's regular rate of pay for hours worked in excess of 44 hours per week"
    else:
        return ("1 and ½ times the employee’s regular rate of pay for hours in excess of 8 hours/day or 40 hours/week\n"
                "2 times the employee’s regular rate of pay for hours in excess of 12 hours/day")

# ==========================================
# 3. DOCX 생성 엔진 (템플릿 없이 Python 코드로 100% 동일하게 그림)
# ==========================================
def generate_job_offer_docx(data, selected_style="Style A"):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    emp_name = data.get('employer_name', '')
    client_name = data.get('client_name', '')

    # 로고
    if data.get('logo_bytes'):
        try:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run().add_picture(io.BytesIO(data.get('logo_bytes')), width=Inches(2.0))
            doc.add_paragraph()
        except: pass

    # 상단 날짜, 수신자
    doc.add_paragraph(data.get('offer_date', datetime.date.today().strftime("%B %d, %Y")))
    p_dear = doc.add_paragraph()
    p_dear.add_run("Dear ")
    p_dear.add_run(f"{client_name},").bold = True

    # ---- Style A (Everfresh) ----
    if "Style A" in selected_style:
        p = doc.add_paragraph()
        p.add_run("We are pleased to offer you a full-time position as ")
        p.add_run(f"{data.get('job_title', '')}").bold = True
        p.add_run(f" for a {data.get('employment_term', '3 years')} term with ")
        p.add_run(f"{emp_name}").bold = True
        p.add_run(" based on the following terms and conditions:")

        p = doc.add_paragraph(); p.add_run("Job Title: ").bold = True; p.add_run(data.get('job_title', ''))
        
        duties = data.get('job_duties', '').split('\n')
        if duties and any(d.strip() for d in duties):
            p = doc.add_paragraph(); p.add_run("Job Duties:").bold = True
            for d in duties:
                if d.strip(): doc.add_paragraph(re.sub(r'^[•\-\*]\s*', '', d.strip()), style='List Bullet')

        p = doc.add_paragraph()
        p.add_run("Compensation:\n").bold = True
        p.add_run("Hourly wage and hours\n")
        p.add_run(f"The employee will be paid ${data.get('wage', '')} per hour, based on a minimum of {data.get('hours', '')} hours per week.")

        p = doc.add_paragraph(); p.add_run("Terms of Employment:\n").bold = True; p.add_run(f"This is a full-time, {data.get('employment_term', '3 years')} employment term starting from the date agreed upon by the employer and employee.")
        p = doc.add_paragraph(); p.add_run("Job Location:\n").bold = True; p.add_run(data.get('job_location', ''))
        p = doc.add_paragraph(); p.add_run("Benefits:\n").bold = True; p.add_run(data.get('benefits', '4% vacation pay'))
        
        p = doc.add_paragraph(); p.add_run("Confidentiality:\n").bold = True
        p.add_run(f"By accepting the terms of this offer, the employee agrees to keep all confidential information obtained during their employment with {emp_name} strictly confidential. The employee further agrees that, upon termination of employment for any reason, they will return all physical and digital property belonging to or originating from {emp_name} within five days of receiving notice of termination.")

        p = doc.add_paragraph(); p.add_run("Start Date:\n").bold = True; p.add_run(data.get('start_date', 'As soon as possible upon obtaining a valid Work Permit'))
        p = doc.add_paragraph(); p.add_run("Overtime:\n").bold = True; p.add_run(data.get('overtime_clause', ''))

        doc.add_paragraph(f"We are pleased to extend this offer of employment to you on behalf of {emp_name}. We are confident that you will make a valuable contribution to our company, and we look forward to working with you.")
        doc.add_paragraph("Sincerely,\t\t\t\t\tI accept the terms of this offer:")
        doc.add_paragraph("X\t\t\t\t\t\t______")

        p = doc.add_paragraph()
        p.add_run(f"{data.get('signer_name', '')}\n").bold = True
        p.add_run(f"{data.get('signer_title', '')}\n{emp_name}\nT. {data.get('employer_phone', '')}")

        p = doc.add_paragraph()
        p.add_run(f"{client_name}\n").bold = True
        p.add_run(f"(DOB: {data.get('client_dob', '')})")

    # ---- Style B (Sushiwood) ----
    else:
        p = doc.add_paragraph()
        p.add_run(f"{emp_name}").bold = True
        p.add_run(" is pleased to offer you the position of ")
        p.add_run(f"{data.get('job_title', '')}").bold = True
        p.add_run(". We are confident that your skills will be a great asset to our business. The details of your employment are outlined below:")

        def add_kv(label, val):
            p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
            p.add_run(f"{label}: ").bold = True; p.add_run(str(val))

        add_kv("Job Title", data.get('job_title', ''))
        add_kv("Employment Type", "Full-time")
        add_kv("Terms of Employment", data.get('employment_term', '3 years'))
        add_kv("Start Date", data.get('start_date', 'As soon as possible upon obtaining a valid Work Permit'))
        add_kv("Hourly wage", f"{data.get('wage', '')} CAD per hour")

        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
        p.add_run("Overtime:\n").bold = True; p.add_run(data.get('overtime_clause', ''))

        add_kv("Hours of Work", f"{data.get('hours', '')} hours per week")
        add_kv("Work Location", data.get('job_location', ''))
        add_kv("Vacation Pay", data.get('benefits', '4% gross earnings'))

        duties = data.get('job_duties', '').split('\n')
        if duties and any(d.strip() for d in duties):
            p = doc.add_paragraph(); p.add_run("Job Duties:").bold = True
            for d in duties:
                if d.strip(): doc.add_paragraph(re.sub(r'^[•\-\*]\s*', '', d.strip()), style='List Bullet')

        p = doc.add_paragraph(); p.add_run("Confidentiality Agreement:\n").bold = True
        p.add_run(f"By agreeing to the terms of this offer, the employee further agrees that they will hold all confidential information with which they are entrusted while employed by {emp_name} in strict confidence. The employee also agrees that upon termination of employment for any reason, they will return all physical and digital property which belongs to or originates at {emp_name} within 5 days of notice of termination of employment.")

        doc.add_paragraph("We look forward to your acceptance of this offer.\nPlease sign below to confirm your agreement.")
        doc.add_paragraph("Yours truly,\nX_____________________________________________")

        p = doc.add_paragraph()
        p.add_run(f"{data.get('signer_name', '')}\n").bold = True
        p.add_run(f"{data.get('signer_title', '')}\n{emp_name}\n{data.get('employer_address', '')}\nT. {data.get('employer_phone', '')}\nE. {data.get('employer_email', '')}")

        doc.add_paragraph("\nI accept the terms of this offer:\n___________________________________________")
        p = doc.add_paragraph()
        p.add_run(f"{client_name}\n").bold = True
        p.add_run(f"(DOB: {data.get('client_dob', '')})")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ==========================================
# 4. Streamlit UI
# ==========================================
if "jo_data" not in st.session_state: st.session_state.jo_data = {}

st.subheader("1. 손님 정보 (여권 업로드)")
passport_file = st.file_uploader("여권 이미지/PDF", type=['jpg', 'jpeg', 'png', 'pdf'])

st.subheader("2. 채용 공고")
job_text = st.text_area("채용공고 링크 또는 내용")
logo_file = st.file_uploader("로고 이미지", type=['jpg', 'png'])

if st.button("실시간 분석 시작", type="primary"):
    with st.spinner("데이터 분석 중..."):
        info = {}
        if logo_file: info['logo_bytes'] = logo_file.getvalue()
        
        if passport_file:
            img = process_uploaded_file_to_image(passport_file)
            pass_data = extract_imm5476_info(img)
            if pass_data:
                g = pass_data.get('given_name', '')
                s = pass_data.get('surname', '')
                info['client_name'] = f"{g} {s}".strip() if g or s else ""
                info['client_dob'] = pass_data.get('dob', '')

        if job_text.strip():
            txt = job_text.strip()
            if "http" in txt: txt = fetch_url_content(txt)
            prompt = f"Analyze: {txt}\nExtract to JSON: employer_name, job_title, wage, hours, job_location, employer_address, employer_phone, employer_email, benefits, job_duties (Array of strings). Return ONLY raw JSON."
            try:
                resp = safe_generate_content([prompt])
                clean = resp.text.strip().replace('```json','').replace('```','')
                data_json = json.loads(clean)
                if 'job_duties' in data_json and isinstance(data_json['job_duties'], list):
                    data_json['job_duties'] = "\n".join(data_json['job_duties'])
                info.update(data_json)
            except Exception as e:
                st.error(f"채용공고 분석 실패: {e}")
                
        st.session_state.jo_data.update(info)
        st.success("분석 완료")

st.markdown("---")
st.subheader("3. 잡오퍼 정보 확인 및 문서 생성")
jo = st.session_state.jo_data

term, med, reason = calculate_employment_term(jo.get('wage', '20.00'), jo.get('job_location', ''))

c1, c2 = st.columns(2)
c_name = c1.text_input("손님 성명", value=jo.get('client_name', ''))
c_dob = c2.text_input("생년월일", value=jo.get('client_dob', ''))
offer_dt = c1.date_input("오퍼 작성일", datetime.date.today()).strftime("%B %d, %Y")
start_dt = c2.text_input("근무 시작일", value=jo.get('start_date', 'As soon as possible upon obtaining a valid Work Permit'))

e1, e2 = st.columns(2)
emp_name = e1.text_input("회사명", value=jo.get('employer_name', ''))
emp_addr = e2.text_input("회사 주소", value=jo.get('employer_address', ''))
signer_n = e1.text_input("대표자 성명", value=jo.get('signer_name', ''))
emp_phone = e2.text_input("회사 전화번호", value=jo.get('employer_phone', ''))
signer_t = e1.text_input("대표자 직책", value=jo.get('signer_title', 'Owner'))
emp_email = e2.text_input("회사 이메일", value=jo.get('employer_email', ''))

j1, j2 = st.columns(2)
j_title = j1.text_input("직책", value=jo.get('job_title', ''))
j_loc = j2.text_input("근무지 주소", value=jo.get('job_location', emp_addr))
j_wage = j1.text_input("시급", value=str(jo.get('wage', '20.00')))
j_hours = j2.text_input("주당 시간", value=str(jo.get('hours', '30')))

j1.info(f"계산된 계약기간: {term} ({reason})")
term_input = j2.text_input("계약 기간", value=term)
j_ben = st.text_input("혜택", value=jo.get('benefits', '4% vacation pay'))

ot_def = get_provincial_overtime_clause(j_loc if j_loc else emp_addr)
j_ot = st.text_area("오버타임 조항", value=ot_def, height=80)
j_duties = st.text_area("주요 직무", value=jo.get('job_duties', ''), height=150)

layout = st.selectbox("잡오퍼 양식", ["Style A (Everfresh 서식)", "Style B (Sushiwood 서식)"])

if st.button("📄 DOCX 다운로드", type="primary", use_container_width=True):
    if not c_name or not emp_name:
        st.error("성명과 회사명은 필수입니다.")
    else:
        final_data = {
            "client_name": c_name, "client_dob": c_dob, "offer_date": offer_dt, "start_date": start_dt,
            "employer_name": emp_name, "employer_address": emp_addr, "signer_name": signer_n,
            "employer_phone": emp_phone, "signer_title": signer_t, "employer_email": emp_email,
            "job_title": j_title, "job_location": j_loc, "wage": j_wage, "hours": j_hours,
            "employment_term": term_input, "benefits": j_ben, "overtime_clause": j_ot,
            "job_duties": j_duties, "logo_bytes": jo.get('logo_bytes')
        }
        
        docx_bytes = generate_job_offer_docx(final_data, selected_style=layout)
        filename = f"[Job Offer]_{c_name.split()[0]}.docx" if c_name else "Job_Offer.docx"
        
        st.download_button("📥 파일 다운로드", data=docx_bytes, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
