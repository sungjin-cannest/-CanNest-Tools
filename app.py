"""
CanNest 잡오퍼 DOCX 생성기 (v3)
----------------------------------
v2 대비 변경점:
  1) 기존 잡오퍼 업로드 시, 빈 템플릿이 아니라 "업로드한 문서 그 자체"를 열어
     바뀌는 항목(받는사람/직책/직무/급여/기간/근무지/overtime/시작일)만 교체.
     로고·서명·회사헤더·Confidentiality 문구는 원본 그대로 유지됩니다.
  2) 여권 + 기존잡오퍼 + 채용공고 분석을 버튼 1번("AI 분석 시작")으로 통합.
  3) 채용공고 URL에 Cloudflare 이메일 난독화([email protected])가 걸려있을 때
     실제 이메일 주소로 디코딩해서 추출하도록 수정.
  4) overtime 조항 중복 삽입 버그, "$$..wage.. per hour" 단위 중복 버그 수정.
"""

import streamlit as st

st.set_page_config(page_title="CanNest 잡오퍼 DOCX 생성기 (v3)", layout="wide")

import os
import io
import re
import copy
import json
import socket
import ipaddress
import datetime
import urllib.request
import urllib.parse
import urllib.error

import fitz  # PyMuPDF
from PIL import Image, ImageOps
import pillow_heif
import pandas as pd

import docx
from docx.oxml.ns import qn

# 신규 Gemini SDK (google-generativeai 아님!)
from google import genai
from google.genai import types

pillow_heif.register_heif_opener()
Image.MAX_IMAGE_PIXELS = None

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "Job_Offer_Template.docx")

# ==========================================
# 0. 인증
# ==========================================
if "APP_PASSWORD" not in st.secrets or "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Cloud의 Secrets 설정이 필요합니다.")
    st.stop()


def check_password():
    if st.session_state.get("password_correct", False):
        return True
    st.title("🔒 CanNest 잡오퍼 생성기")
    pwd = st.text_input("접속 비밀번호를 입력하세요", type="password")
    if st.button("확인"):
        if pwd == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False


if not check_password():
    st.stop()

# ==========================================
# 1. Gemini 클라이언트 (신규 SDK) — 구조화 출력 + 모델 폴백
# ==========================================
_client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

MODEL_CANDIDATES = [
    "gemini-2.5-flash",
    "gemini-3.7-flash",
    "gemini-2.5-pro",
]


def call_gemini(parts, response_schema=None):
    contents = []
    for p in parts:
        if isinstance(p, str):
            contents.append(p)
        elif isinstance(p, Image.Image):
            buf = io.BytesIO()
            p.save(buf, format="JPEG")
            contents.append(types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"))
        else:
            contents.append(p)

    config = {}
    if response_schema is not None:
        config["response_mime_type"] = "application/json"
        config["response_schema"] = response_schema

    last_error = None
    for model_name in MODEL_CANDIDATES:
        try:
            response = _client.models.generate_content(
                model=model_name, contents=contents, config=config or None
            )
            return response
        except Exception as e:
            last_error = e
            continue
    raise Exception(f"Gemini API 호출 실패 (모든 모델 시도함): {last_error}")


def call_gemini_json(parts, response_schema):
    try:
        response = call_gemini(parts, response_schema=response_schema)
        return json.loads(response.text)
    except Exception as e:
        st.warning(f"AI 분석 중 경고: {e}")
        return None


PASSPORT_SCHEMA = {
    "type": "object",
    "properties": {
        "surname": {"type": "string"},
        "given_name": {"type": "string"},
        "dob": {"type": "string", "description": "YYYY-MM-DD"},
        "uci": {"type": "string"},
    },
    "required": ["surname", "given_name"],
}

JOB_POSTING_SCHEMA = {
    "type": "object",
    "properties": {
        "employer_name": {"type": "string"},
        "job_title": {"type": "string"},
        "noc_code": {"type": "string"},
        "wage": {"type": "string"},
        "hours": {"type": "string"},
        "job_location": {"type": "string"},
        "employer_address": {"type": "string"},
        "employer_phone": {"type": "string"},
        "employer_email": {"type": "string"},
        "benefits": {"type": "string"},
        "job_duties": {"type": "array", "items": {"type": "string"}},
    },
}

EXISTING_OFFER_SCHEMA = {
    "type": "object",
    "properties": {
        "client_name": {"type": "string"},
        "client_dob": {"type": "string"},
        "employer_name": {"type": "string"},
        "signer_name": {"type": "string"},
        "signer_title": {"type": "string"},
        "employer_address": {"type": "string"},
        "employer_phone": {"type": "string"},
        "employer_email": {"type": "string"},
        "job_title": {"type": "string"},
        "wage": {"type": "string"},
        "hours": {"type": "string"},
        "job_location": {"type": "string"},
        "benefits": {"type": "string"},
        "job_duties": {"type": "array", "items": {"type": "string"}},
    },
}

# ==========================================
# 2. 이미지 / 파일 헬퍼
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


def format_full_name(surname, given_name):
    s = str(surname).strip() if surname else ""
    g = str(given_name).strip() if given_name else ""
    if not s and not g:
        return ""
    if not s:
        return g
    if not g:
        return s
    return f"{g} {s}"


def prepare_document_for_gemini(file_bytes, mime_type, file_name=""):
    ext = os.path.splitext(file_name)[1].lower() if file_name else ""
    if "word" in mime_type.lower() or "doc" in mime_type.lower() or ext in [".doc", ".docx"]:
        try:
            doc_obj = docx.Document(io.BytesIO(file_bytes))
            text_list = [p.text for p in doc_obj.paragraphs if p.text.strip()]
            for table in doc_obj.tables:
                for row in table.rows:
                    row_txt = " | ".join([c.text.strip() for c in row.cells if c.text.strip()])
                    if row_txt:
                        text_list.append(row_txt)
            full_text = "\n".join(text_list)
            if full_text.strip():
                return [f"\n--- [Word Document: {file_name}] ---\n{full_text[:20000]}\n"]
        except Exception:
            pass

    if "pdf" in mime_type.lower():
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = "".join(page.get_text("text") + "\n" for page in doc)
            if len(text.strip()) > 100:
                return [f"\n--- [Document: {file_name}] ---\n{text[:20000]}\n"]
        except Exception:
            pass
    return [types.Part.from_bytes(data=file_bytes, mime_type=mime_type)]


def extract_imm5476_info(image):
    prompt = "Analyze this identity document and extract the requested fields exactly."
    return call_gemini_json([prompt, image], PASSPORT_SCHEMA)


def parse_existing_job_offer(file_bytes, mime_type):
    prompt = "Analyze this existing Job Offer document and extract the requested fields exactly."
    contents = prepare_document_for_gemini(file_bytes, mime_type, "Existing_Job_Offer.pdf")
    return call_gemini_json([prompt] + contents, EXISTING_OFFER_SCHEMA) or {}


# ==========================================
# 3. URL 보안 강화 fetch (SSRF 방지) + Cloudflare 이메일 디코딩
# ==========================================
def _is_hostname_safe(hostname: str) -> bool:
    try:
        infos = socket.getaddrinfo(hostname, None)
    except socket.gaierror:
        return False
    if not infos:
        return False
    for info in infos:
        ip_str = info[4][0]
        try:
            ip = ipaddress.ip_address(ip_str)
        except ValueError:
            return False
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast or ip.is_unspecified:
            return False
    return True


class _SafeRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parsed = urllib.parse.urlparse(newurl)
        if parsed.scheme not in ("http", "https") or not parsed.hostname or not _is_hostname_safe(parsed.hostname):
            raise urllib.error.URLError(f"차단된 리다이렉트 대상: {newurl}")
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _decode_cf_email(cfemail_hex: str) -> str:
    """Cloudflare 이메일 난독화(data-cfemail) XOR 디코딩."""
    try:
        r = int(cfemail_hex[:2], 16)
        return "".join(
            chr(int(cfemail_hex[i:i + 2], 16) ^ r)
            for i in range(2, len(cfemail_hex), 2)
        )
    except Exception:
        return ""


_CF_EMAIL_RE = re.compile(
    r'data-cfemail="([0-9a-fA-F]+)"[^>]*>\s*\[email\s*protected\]\s*<',
    re.IGNORECASE,
)


def _decloak_cloudflare_emails(html: str) -> str:
    """채용공고 페이지가 Cloudflare 이메일 보호를 쓰는 경우,
    화면에 [email protected]로 표시되는 부분을 실제 주소로 되돌립니다."""
    return _CF_EMAIL_RE.sub(lambda m: _decode_cf_email(m.group(1)) + "<", html)


def fetch_url_content_safe(url: str, timeout: int = 12, max_bytes: int = 2_000_000):
    """(text, error_message) 튜플 반환. error_message가 None이면 성공."""
    target_url = url.strip()
    if not re.match(r"^https?://", target_url, re.IGNORECASE):
        target_url = "https://" + target_url

    parsed = urllib.parse.urlparse(target_url)
    if parsed.scheme not in ("http", "https"):
        return "", "허용되지 않는 URL 스킴입니다."
    if not parsed.hostname or not _is_hostname_safe(parsed.hostname):
        return "", "내부망/사설 IP로 확인되어 요청을 차단했습니다."

    opener = urllib.request.build_opener(_SafeRedirectHandler)
    req = urllib.request.Request(target_url, headers={"User-Agent": "Mozilla/5.0 (CanNestJobOfferBot)"})
    try:
        with opener.open(req, timeout=timeout) as resp:
            raw = resp.read(max_bytes + 1)
            if len(raw) > max_bytes:
                raw = raw[:max_bytes]
            html = raw.decode("utf-8", errors="ignore")
    except Exception as e:
        return "", f"URL 요청 실패: {e}"

    html = _decloak_cloudflare_emails(html)  # ← Cloudflare 이메일 디코딩

    text = re.sub(r"<script\b[^<]*(?:(?!</script>)<[^<]*)*</script>", "", html, flags=re.IGNORECASE)
    text = re.sub(r"<style\b[^<]*(?:(?!</style>)<[^<]*)*</style>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:15000], None


# ==========================================
# 4. 전체 13개 주/준주 데이터 (중위임금 백업 + 초과근무 조항)
#    ⚠️ 아래 수치/조항은 참고용 백업값이며 실제 발급 전 최신 법령 재확인 필수
#    v3: PROVINCES[code]["overtime"]의 첫 줄(인트로 문장)은 신규 템플릿 생성용으로만 쓰고,
#    기존문서 편집 시에는 get_provincial_overtime_bullets()로 인트로를 뺀 불릿만 씁니다
#    (문서에 이미 공통 인트로 문장이 있어 중복되는 걸 방지).
# ==========================================
BACKUP_WAGE_AS_OF = "2026-01"

PROVINCES = {
    "BC": {
        "names": ["British Columbia", "BC"],
        "backup_wage": 38.40,
        "overtime": (
            "Overtime will be paid in accordance with BC employment standards legislation:\n"
            "1.5 times regular wage for hours over 8/day or 40/week; and\n"
            "2 times regular wage for hours over 12/day."
        ),
    },
    "AB": {
        "names": ["Alberta", "AB"],
        "backup_wage": 37.50,
        "overtime": "Overtime will be paid in accordance with AB employment standards legislation:\n1.5 times regular rate of pay for hours in excess of 8 hours/day or 44 hours/week.",
    },
    "ON": {
        "names": ["Ontario", "ON"],
        "backup_wage": 36.92,
        "overtime": "Overtime will be paid in accordance with ON employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 44 hours per week.",
    },
    "SK": {
        "names": ["Saskatchewan", "SK"],
        "backup_wage": 34.62,
        "overtime": "Overtime will be paid in accordance with SK employment standards legislation:\n1.5 times regular rate of pay for hours over 8 hours/day or 40 hours/week.",
    },
    "MB": {
        "names": ["Manitoba", "MB"],
        "backup_wage": 31.33,
        "overtime": "Overtime will be paid in accordance with MB employment standards legislation:\n1.5 times regular rate of pay for hours over 8 hours/day or 40 hours/week.",
    },
    "NB": {
        "names": ["New Brunswick", "NB"],
        "backup_wage": 31.73,
        "overtime": "Overtime will be paid in accordance with NB employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 44 hours per week.",
    },
    "NS": {
        "names": ["Nova Scotia", "NS"],
        "backup_wage": 31.96,
        "overtime": "Overtime will be paid in accordance with NS employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 48 hours per week.",
    },
    "PE": {
        "names": ["Prince Edward Island"],
        "backup_wage": 31.20,
        "overtime": "Overtime will be paid in accordance with PE employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 48 hours per week.",
    },
    "NL": {
        "names": ["Newfoundland and Labrador", "NL"],
        "backup_wage": 33.60,
        "overtime": "Overtime will be paid in accordance with NL employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 40 hours per week.",
    },
    "YT": {
        "names": ["Yukon", "YT"],
        "backup_wage": 45.60,
        "overtime": "Overtime will be paid in accordance with YT employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 8 hours/day or 40 hours/week.",
    },
    "NT": {
        "names": ["Northwest Territories", "NT"],
        "backup_wage": 48.00,
        "overtime": "Overtime will be paid in accordance with NT employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 8 hours/day or 40 hours/week.",
    },
    "NU": {
        "names": ["Nunavut", "NU"],
        "backup_wage": 45.00,
        "overtime": "Overtime will be paid in accordance with NU employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 8 hours/day or 40 hours/week.",
    },
    "QC": {
        "names": ["Quebec", "Québec", "QC"],
        "backup_wage": 36.00,
        "overtime": "Overtime will be paid in accordance with QC employment standards legislation:\n1.5 times regular rate of pay for hours worked in excess of 40 hours per week.",
    },
}

BACKUP_WAGES = {code: v["backup_wage"] for code, v in PROVINCES.items()}


@st.cache_data(ttl=86400, show_spinner=False)
def get_live_esdc_median_wages():
    url = "https://www.canada.ca/en/employment-social-development/services/foreign-workers/median-wage.html"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        tables = pd.read_html(io.StringIO(html))
        wages = {}
        name_to_code = {n.upper(): code for code, v in PROVINCES.items() for n in v["names"]}
        for table in tables:
            table.columns = [str(c) for c in table.columns]
            for _, row in table.iterrows():
                row_text = " ".join(str(x) for x in row.values)
                for name_upper, code in name_to_code.items():
                    if name_upper in row_text.upper():
                        dollar_matches = re.findall(r"\$\s*(\d+(?:\.\d+)?)", row_text)
                        if dollar_matches:
                            wages[code] = float(dollar_matches[-1])
        if wages:
            today_tag = datetime.date.today().strftime("%Y-%m-%d")
            return wages, "Canada.ca ESDC 실시간 데이터", today_tag
    except Exception:
        pass

    return dict(BACKUP_WAGES), "ESDC 백업 기준표 (네트워크 조회 실패)", BACKUP_WAGE_AS_OF


def detect_province(address_text: str) -> str:
    addr_upper = str(address_text).upper()
    for code, v in PROVINCES.items():
        for name in v["names"]:
            if re.search(r"\b" + re.escape(name.upper()) + r"\b", addr_upper):
                return code
    return "BC"


def calculate_employment_term(wage_val, address_text):
    try:
        wage_match = re.search(r"(\d+(?:\.\d+)?)", str(wage_val))
        wage = float(wage_match.group(1)) if wage_match else 0.0
        detected_prov = detect_province(address_text)
        live_wages, source_tag, as_of = get_live_esdc_median_wages()
        median_wage = live_wages.get(detected_prov, BACKUP_WAGES[detected_prov])
        stream = "High-Wage Stream (3년 오퍼)" if wage >= median_wage else "Low-Wage Stream (1년 오퍼)"
        term = "3-year" if wage >= median_wage else "1-year"
        reason = f"{detected_prov} 중위임금 ${median_wage:.2f} ({source_tag}, 기준일 {as_of}) → {stream}"
        return term, median_wage, reason, detected_prov
    except Exception:
        return "3-year", 38.40, "기본값 적용 (계산 실패)", "BC"


def get_provincial_overtime_clause(address_text):
    """신규(빈 템플릿) 생성용 — 인트로 문장 포함 전체 조항."""
    return PROVINCES[detect_province(address_text)]["overtime"]


def get_provincial_overtime_bullets(address_text):
    """기존문서 편집용 — 인트로 문장 제외, 요율 불릿 줄만."""
    full_clause = PROVINCES[detect_province(address_text)]["overtime"]
    lines = [l.strip() for l in full_clause.split("\n") if l.strip()]
    if lines and lines[0].lower().startswith("overtime will be paid"):
        lines = lines[1:]
    return lines


# ==========================================
# 5. DOCX 엔진 (공용 헬퍼)
# ==========================================
def _iter_paragraphs_in_cell(cell):
    for p in cell.paragraphs:
        yield p
    for table in cell.tables:
        for row in table.rows:
            for c in row.cells:
                yield from _iter_paragraphs_in_cell(c)


def iter_all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_paragraphs_in_cell(cell)
    for section in doc.sections:
        for hf in (section.header, section.footer):
            for p in hf.paragraphs:
                yield p


def set_paragraph_text(paragraph, new_text):
    """단락의 첫 번째 run 서식을 유지하면서 전체 텍스트를 교체."""
    if not paragraph.runs:
        paragraph.add_run(new_text)
        return
    base_rpr = paragraph.runs[0]._r.find(qn("w:rPr"))
    for r in list(paragraph.runs):
        r._r.getparent().remove(r._r)
    run = paragraph.add_run()
    if base_rpr is not None:
        run._r.insert(0, copy.deepcopy(base_rpr))
    run.text = new_text


def replace_placeholders(doc, mapping: dict):
    pattern = re.compile("|".join(re.escape(k) for k in mapping.keys()))
    for p in iter_all_paragraphs(doc):
        full_text = "".join(r.text for r in p.runs)
        if not full_text or not pattern.search(full_text):
            continue
        new_text = pattern.sub(lambda m: str(mapping[m.group(0)]), full_text)
        if new_text != full_text:
            set_paragraph_text(p, new_text)


def find_paragraph_starting_with(doc, prefix):
    for p in doc.paragraphs:
        if p.text.strip().lower().startswith(prefix.lower()):
            return p
    return None


def insert_bullets_after(anchor_paragraph, texts, template_bullet_xml):
    ref = anchor_paragraph._p
    for text in texts:
        new_p = copy.deepcopy(template_bullet_xml)
        ref.addnext(new_p)
        ref = new_p
        wrapped = docx.text.paragraph.Paragraph(new_p, anchor_paragraph._parent)
        set_paragraph_text(wrapped, text)


def replace_overtime_block(doc, overtime_lines):
    """(신규 템플릿용) 'Overtime will be paid...' 바로 다음의 불릿들을 교체."""
    intro = find_paragraph_starting_with(doc, "Overtime will be paid")
    if intro is None:
        return
    bullets = []
    node = intro._p.getnext()
    while node is not None and node.tag == qn("w:p"):
        wrapped = docx.text.paragraph.Paragraph(node, intro._parent)
        if wrapped.style.name == "List Paragraph":
            bullets.append(node)
            node = node.getnext()
        else:
            break
    if not bullets:
        return
    template_xml = copy.deepcopy(bullets[0])
    anchor = bullets[0].getprevious()
    for b in bullets:
        b.getparent().remove(b)
    ref = anchor
    for line in overtime_lines:
        new_p = copy.deepcopy(template_xml)
        ref.addnext(new_p)
        ref = new_p
        wrapped = docx.text.paragraph.Paragraph(new_p, intro._parent)
        set_paragraph_text(wrapped, line)


def insert_logo(doc, logo_bytes):
    for p in doc.paragraphs:
        if "[Company Logo]" in p.text:
            set_paragraph_text(p, "")
            if logo_bytes:
                run = p.add_run()
                try:
                    run.add_picture(io.BytesIO(logo_bytes), width=docx.shared.Inches(2.0))
                except Exception:
                    pass
            return


def clean_wage(value) -> str:
    m = re.search(r"\d+(?:\.\d+)?", str(value))
    return m.group(0) if m else str(value).strip()


def clean_hours(value) -> str:
    text = str(value)
    text = re.sub(r"(?i)\b(hours?|hrs?)\b", "", text)
    text = re.sub(r"(?i)(per\s*week|/\s*week)", "", text)
    return text.strip(" -/\t")


def build_wage_sentence(wage, hours) -> str:
    """(기존문서 편집용) 매번 새로 조립해서 '$$..per hour' 같은 단위 중복 방지."""
    return f"The employee will be paid ${clean_wage(wage)}/hour, based on a minimum of {clean_hours(hours)} hours per week."


def build_term_sentence(term_years: str) -> str:
    return f"This is a full-time, {term_years} employment term starting from the date agreed upon by the employer and employee."


# ==========================================
# 5a. 신규(빈 템플릿) 생성 엔진
# ==========================================
def generate_job_offer_docx(data: dict) -> bytes:
    doc = docx.Document(TEMPLATE_PATH)

    duties = data.get("job_duties", [])
    if isinstance(duties, str):
        duties = [re.sub(r"^[•\-\*]\s*", "", d.strip()) for d in duties.split("\n") if d.strip()]

    insert_logo(doc, data.get("logo_bytes"))

    for p in iter_all_paragraphs(doc):
        full_text = "".join(r.text for r in p.runs)
        if "[City, Province, Postal Code]" in full_text:
            set_paragraph_text(p, data.get("job_location", ""))
            break

    noc_code = str(data.get("noc_code", "")).strip()
    if not noc_code:
        for p in iter_all_paragraphs(doc):
            full_text = "".join(r.text for r in p.runs)
            if "(NOC [Code])" in full_text:
                set_paragraph_text(p, full_text.replace(" (NOC [Code])", ""))
                break

    mapping = {
        "[Company Name]": data.get("employer_name", ""),
        "[Company Address]": data.get("employer_address", ""),
        "[Company Number]": data.get("employer_phone", ""),
        "[Date]": data.get("offer_date", datetime.date.today().strftime("%B %d, %Y")),
        "[Employee Name]": data.get("client_name", ""),
        "[Job Title]": data.get("job_title", ""),
        "[Employment Term]": data.get("employment_term", ""),
        "[Code]": noc_code,
        "[Hourly Wage]": clean_wage(data.get("wage", "")),
        "[Minimum Weekly Hours]": clean_hours(data.get("hours", "")),
        "[Vacation Pay / Benefits]": data.get("benefits", ""),
        "[Authorized Representative Name]": data.get("signer_name", ""),
        "[Representative Title]": data.get("signer_title", ""),
        "[Representative Phone Number]": data.get("employer_phone", ""),
        "[Representative email address]": data.get("employer_email", ""),
        "[Date of Birth]": data.get("client_dob", ""),
    }
    replace_placeholders(doc, mapping)

    duties_anchor = find_paragraph_starting_with(doc, "Job Duties:")
    if duties_anchor is not None and duties:
        list_bullets = [p for p in doc.paragraphs if p.style.name == "List Paragraph"]
        if list_bullets:
            template_xml = copy.deepcopy(list_bullets[0]._p)
            insert_bullets_after(duties_anchor, duties, template_xml)

    overtime_text = data.get("overtime_clause", "")
    overtime_lines = [l.strip() for l in overtime_text.split("\n") if l.strip()]
    if overtime_lines:
        replace_overtime_block(doc, overtime_lines)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ==========================================
# 5b. 기존 문서 편집 엔진 (v3 신규)
#    업로드된 기존 잡오퍼 docx를 그대로 열어 바뀌는 항목만 교체.
#    로고 / 서명 / 회사 헤더 / Confidentiality 문구는 건드리지 않음.
# ==========================================
def _replace_label_line(doc, label, new_value):
    """'Job Title: Chef' 처럼 라벨+내용이 한 줄인 경우, 내용만 교체."""
    p = find_paragraph_starting_with(doc, label)
    if p is None:
        return False
    set_paragraph_text(p, f"{label} {new_value}")
    return True


def _replace_next_paragraph(doc, label, new_value):
    """'Job Location:' 처럼 라벨 단독 줄 다음에 내용이 오는 경우,
    라벨 줄은 유지, 바로 다음 문단만 통째로 교체."""
    p = find_paragraph_starting_with(doc, label)
    if p is None:
        return False
    nxt = p._p.getnext()
    if nxt is None or nxt.tag != qn("w:p"):
        return False
    wrapped = docx.text.paragraph.Paragraph(nxt, p._parent)
    set_paragraph_text(wrapped, new_value)
    return True


def _replace_bullets_after(doc, label, new_lines):
    """label 문단 뒤 List Paragraph 형제들을 new_lines로 교체.
    label과 첫 불릿 사이 인트로 문장은 건드리지 않음."""
    anchor = find_paragraph_starting_with(doc, label)
    if anchor is None or not new_lines:
        return False

    node = anchor._p.getnext()
    while node is not None and node.tag == qn("w:p"):
        wrapped = docx.text.paragraph.Paragraph(node, anchor._parent)
        if wrapped.style.name == "List Paragraph":
            break
        node = node.getnext()

    bullets = []
    cur = node
    while cur is not None and cur.tag == qn("w:p"):
        wrapped = docx.text.paragraph.Paragraph(cur, anchor._parent)
        if wrapped.style.name != "List Paragraph":
            break
        bullets.append(cur)
        cur = cur.getnext()

    if not bullets:
        return False

    template_xml = copy.deepcopy(bullets[0])
    ref = bullets[0].getprevious()
    for b in bullets:
        b.getparent().remove(b)
    for line in new_lines:
        new_p = copy.deepcopy(template_xml)
        ref.addnext(new_p)
        ref = new_p
        wrapped = docx.text.paragraph.Paragraph(new_p, anchor._parent)
        set_paragraph_text(wrapped, line)
    return True


def _update_salutation(doc, new_name):
    p = find_paragraph_starting_with(doc, "Dear ")
    if p is None:
        return False
    new_text = re.sub(r"^Dear .+?,", f"Dear {new_name},", p.text)
    set_paragraph_text(p, new_text)
    return True


def _update_job_title_line(doc, job_title, noc_code):
    p = find_paragraph_starting_with(doc, "Job Title:")
    if p is None:
        return False
    if noc_code and str(noc_code).strip():
        new_text = f"Job Title: {job_title} (NOC {noc_code})"
    else:
        new_text = f"Job Title: {job_title}"
    set_paragraph_text(p, new_text)
    return True


def generate_job_offer_from_existing(existing_file_bytes: bytes, data: dict) -> bytes:
    """업로드된 '기존 잡오퍼' docx를 그대로 열어서, 바뀌는 항목만 교체.
    로고 / 서명란 / 회사 헤더 / Confidentiality 문구는 건드리지 않음."""
    doc = docx.Document(io.BytesIO(existing_file_bytes))

    # 상단 날짜 (예: 'April 4, 2025')
    if data.get("offer_date"):
        for p in doc.paragraphs[:5]:
            if re.match(r"^[A-Z][a-z]+ \d{1,2},? \d{4}$", p.text.strip()):
                set_paragraph_text(p, data["offer_date"])
                break

    if data.get("client_name"):
        _update_salutation(doc, data["client_name"])

    if data.get("job_title"):
        _update_job_title_line(doc, data["job_title"], data.get("noc_code", ""))

    duties = data.get("job_duties", [])
    if isinstance(duties, str):
        duties = [re.sub(r"^[•\-\*]\s*", "", d.strip()) for d in duties.split("\n") if d.strip()]
    if duties:
        _replace_bullets_after(doc, "Job Duties:", duties)

    if data.get("wage") and data.get("hours"):
        sentence = build_wage_sentence(data["wage"], data["hours"])
        if not _replace_next_paragraph(doc, "Hourly wage and hours", sentence):
            _replace_next_paragraph(doc, "Hourly Wage and Hours", sentence)

    if data.get("employment_term"):
        _replace_next_paragraph(doc, "Terms of Employment:", build_term_sentence(data["employment_term"]))

    if data.get("job_location"):
        _replace_next_paragraph(doc, "Job Location:", data["job_location"])

    bullets = get_provincial_overtime_bullets(data.get("job_location", ""))
    if bullets:
        _replace_bullets_after(doc, "Overtime:", bullets)

    if data.get("start_date"):
        _replace_next_paragraph(doc, "Start Date:", data["start_date"])

    # 로고 / 서명 이미지 / 회사명·주소·전화 / Confidentiality 문구는 손대지 않음 → 원본 유지

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ==========================================
# 6. Streamlit UI
# ==========================================
st.title("📄 잡오퍼 DOCX 생성기 (v3)")
st.caption("기존 문서 그대로 편집 · 1단계 통합 분석 · 13개 주/준주 커버리지 · Cloudflare 이메일 디코딩")

if "job_offer_data" not in st.session_state:
    st.session_state.job_offer_data = {}

col_reset, _ = st.columns([1, 5])
with col_reset:
    if st.button("🔄 새 케이스 시작 (초기화)"):
        st.session_state.job_offer_data = {}
        st.session_state.pop("_old_jo_bytes", None)
        for k in ("jo_passport", "jo_old_file", "jo_logo"):
            st.session_state.pop(k, None)
        st.rerun()

st.markdown("---")
st.subheader("1. 서류 업로드")

passport_file = st.file_uploader(
    "손님 여권 이미지 또는 PDF",
    type=["jpg", "jpeg", "png", "pdf", "heic", "HEIC"],
    key="jo_passport",
)

old_jo_file = st.file_uploader(
    "기준으로 사용할 기존 잡오퍼 (있으면 이 문서를 그대로 편집합니다)",
    type=["docx"],
    key="jo_old_file",
)

job_posting_text = st.text_area(
    "채용 공고 링크(URL) 또는 공고 텍스트",
    height=100,
)

logo_file = st.file_uploader(
    "회사 로고 이미지 (기존 잡오퍼 편집 모드에서는 무시됩니다 — 원본 로고 그대로 유지)",
    type=["jpg", "jpeg", "png"],
    key="jo_logo",
)

if st.button("AI 분석 시작", type="primary", use_container_width=True):
    with st.spinner("정보 추출 중..."):
        extracted_info = {}

        if logo_file and not old_jo_file:
            extracted_info["logo_bytes"] = logo_file.getvalue()

        if passport_file:
            pass_img = process_uploaded_file_to_image(passport_file)
            pass_data = extract_imm5476_info(pass_img)
            if pass_data:
                extracted_info["client_name"] = format_full_name(
                    pass_data.get("surname", ""), pass_data.get("given_name", "")
                )
                extracted_info["client_dob"] = pass_data.get("dob", "")

        if old_jo_file:
            existing_parsed = parse_existing_job_offer(
                old_jo_file.getvalue(),
                old_jo_file.type or "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            if existing_parsed:
                extracted_info.update(existing_parsed)
            st.session_state["_old_jo_bytes"] = old_jo_file.getvalue()
        else:
            st.session_state.pop("_old_jo_bytes", None)

        if job_posting_text.strip():
            raw_input = job_posting_text.strip()
            is_url = bool(re.search(r"https?://[^\s]+|www\.[^\s]+", raw_input))
            if is_url:
                url_match = re.search(r"(https?://[^\s]+|www\.[^\s]+)", raw_input).group(0)
                fetched_text, err = fetch_url_content_safe(url_match)
                if err:
                    st.warning(f"URL 조회 실패: {err} (입력하신 텍스트 자체를 분석합니다)")
                text_to_analyze = fetched_text if fetched_text else raw_input
            else:
                text_to_analyze = raw_input

            job_extracted = call_gemini_json(
                [f"Analyze this job posting content and extract the requested fields:\n\n{text_to_analyze}"],
                JOB_POSTING_SCHEMA,
            )
            if job_extracted:
                extracted_info.update(job_extracted)

        st.session_state.job_offer_data.update(extracted_info)
        st.success("분석이 완료되었습니다.")

st.markdown("---")
st.subheader("2. 최종 잡오퍼 정보 확인 및 수정")

jo_data = st.session_state.job_offer_data
temp_wage = jo_data.get("wage", "20.15")
temp_loc = jo_data.get("job_location", jo_data.get("employer_address", ""))
calc_term, calc_median, calc_reason, calc_prov = calculate_employment_term(temp_wage, temp_loc)

col_c1, col_c2 = st.columns(2)
with col_c1:
    c_name = st.text_input("손님 성명", value=jo_data.get("client_name", ""))
    offer_dt = st.date_input("오퍼 작성일", datetime.date.today()).strftime("%B %d, %Y")
    term_str = st.text_input("계약 기간", value=jo_data.get("employment_term", calc_term))
    st.info(f"🌐 **판정 근거 (주: {calc_prov}):** {calc_reason}")
with col_c2:
    c_dob = st.text_input("손님 생년월일", value=jo_data.get("client_dob", ""))
    start_dt_str = st.text_input(
        "근무 시작일",
        value=jo_data.get("start_date", "The employment start date will be as soon as possible upon the employee’s authorization to work in Canada."),
    )

st.markdown("#### 🏢 고용주 및 회사 정보")
col_e1, col_e2 = st.columns(2)
with col_e1:
    emp_name = st.text_input("회사명", value=jo_data.get("employer_name", ""))
    signer_n = st.text_input("대표자 성명", value=jo_data.get("signer_name", ""))
    signer_t = st.text_input("대표자 직책", value=jo_data.get("signer_title", "Director"))
with col_e2:
    emp_addr = st.text_input("회사 대표 주소", value=jo_data.get("employer_address", ""))
    emp_phone = st.text_input("회사 전화번호", value=jo_data.get("employer_phone", ""))
    emp_email = st.text_input("회사 이메일", value=jo_data.get("employer_email", ""))

st.markdown("#### 💼 근무 조건")
col_j1, col_j2 = st.columns(2)
with col_j1:
    j_title = st.text_input("직책", value=jo_data.get("job_title", ""))
    j_wage = st.text_input(
        "시급 (CAD)", value=str(jo_data.get("wage", "20.15")),
        help="숫자만 입력해도 되고, '$38.50/hour'처럼 붙여넣어도 자동으로 숫자만 추출됩니다.",
    )
    j_hours = st.text_input(
        "주당 근무시간", value=str(jo_data.get("hours", "30-40")),
        help="'30-40 hours/week'처럼 붙여넣어도 단위는 자동으로 제거됩니다.",
    )
with col_j2:
    j_loc = st.text_input("근무지 주소", value=jo_data.get("job_location", emp_addr))
    j_benefits = st.text_input("혜택", value=jo_data.get("benefits", "4% vacation pay"))
    j_noc = st.text_input("NOC 코드", value=jo_data.get("noc_code", ""))
    if not j_noc.strip():
        st.warning("⚠️ NOC 코드가 비어 있습니다. LMIA 관련 서류는 보통 NOC 코드가 필수이니 확인 후 입력해주세요.")

auto_ot_clause = get_provincial_overtime_clause(j_loc or emp_addr)
j_ot = st.text_area("오버타임 조항 (주별 기준 자동 적용, 신규 생성 시에만 사용됨)", value=auto_ot_clause, height=100)

duties_input_str = jo_data.get("job_duties", [])
if isinstance(duties_input_str, list):
    duties_input_str = "\n".join(duties_input_str)
j_duties_text = st.text_area("주요 직무 (한 줄에 하나씩)", value=duties_input_str, height=150)

if st.session_state.get("_old_jo_bytes"):
    st.info("📎 기존 잡오퍼 문서를 그대로 편집합니다. 로고/서명/회사 헤더는 원본 그대로 유지됩니다.")

st.markdown("---")
if st.button("📄 MS Word (.docx) 생성 및 다운로드", type="primary", use_container_width=True):
    if not c_name or not emp_name or not j_title:
        st.error("손님 성명, 회사명, 직책은 필수 입력 항목입니다.")
    else:
        final_jo_dict = {
            "client_name": c_name, "client_dob": c_dob, "offer_date": offer_dt, "employment_term": term_str,
            "start_date": start_dt_str, "employer_name": emp_name, "signer_name": signer_n, "signer_title": signer_t,
            "employer_address": emp_addr, "employer_phone": emp_phone, "employer_email": emp_email,
            "job_title": j_title, "noc_code": j_noc, "wage": j_wage, "hours": j_hours, "job_location": j_loc,
            "benefits": j_benefits, "overtime_clause": j_ot, "job_duties": j_duties_text,
            "logo_bytes": jo_data.get("logo_bytes"),
        }

        old_bytes = st.session_state.get("_old_jo_bytes")
        if old_bytes:
            docx_bytes = generate_job_offer_from_existing(old_bytes, final_jo_dict)
        else:
            if not os.path.exists(TEMPLATE_PATH):
                st.error(f"템플릿 파일을 찾을 수 없습니다: {TEMPLATE_PATH} (저장소에 Job_Offer_Template.docx를 app.py와 같은 폴더에 커밋하세요)")
                st.stop()
            docx_bytes = generate_job_offer_docx(final_jo_dict)

        crm_client = "NAME"
        if c_name:
            parts = c_name.strip().split()
            if parts:
                crm_client = parts[0].capitalize()
        out_filename = f"[Job Offer]_{crm_client}.docx"

        st.success("잡오퍼 DOCX 문서 생성이 완료되었습니다!")
        st.download_button(
            label="📥 Job Offer 다운로드",
            data=docx_bytes,
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
