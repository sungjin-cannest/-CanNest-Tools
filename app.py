from concurrent.futures import ThreadPoolExecutor
import datetime
import hashlib
import io
import inspect as _inspect
import json
import os
import re
import time
import uuid
import zipfile
from zoneinfo import ZoneInfo

try:
    import pymupdf as fitz  # PyMuPDF 최신 버전
except ImportError:
    import fitz  # 구버전 호환
# [수정] 새 Gemini 라이브러리(google-genai) 우선 사용 — 구버전(google-generativeai)은 지원 종료되었고
# '생각 단계(thinking)' 길이를 조절할 수 없어 Gemini 3.x에서 느림
try:
    from google import genai as gg
    from google.genai import types as gtypes
    NEW_SDK = True
except ImportError:
    import google.generativeai as genai
    NEW_SDK = False

from PIL import Image, ImageDraw, ImageFont, ImageOps
import pillow_heif  # HEIC 지원 라이브러리
import streamlit as st

# 📌 가장 먼저 실행되어야 하는 Streamlit 페이지 설정
st.set_page_config(page_title="CanNest 통합 업무 시스템", layout="wide")

# DOCX 파싱 지원
try:
    import docx
    from docx.shared import RGBColor
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# HEIC 이미지 지원 등록 및 고해상도 제한 해제
pillow_heif.register_heif_opener()
Image.MAX_IMAGE_PIXELS = None

LOCAL_TZ = ZoneInfo("America/Vancouver")

# Streamlit 버전에 따라 '가로 꽉 채우기' 옵션 이름이 다름 (신버전: width="stretch")
try:
    _WIDE = {"width": "stretch"} if "width" in _inspect.signature(st.download_button).parameters else {"use_container_width": True}
except Exception:
    _WIDE = {"use_container_width": True}

HANGUL_RE = re.compile(r"[\u3131-\u318E\uAC00-\uD7A3]")
IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.heic', '.heif', '.webp')

# ==========================================
# 0. Secrets 안전 검사 및 보안 비밀번호 설정
# ==========================================
if "APP_PASSWORD" not in st.secrets or "GEMINI_API_KEY" not in st.secrets:
  st.error("⚠️ Streamlit Cloud의 Secrets 설정이 필요합니다.")
  st.info(
      "우측 하단 [Manage app] -> [Settings] -> [Secrets]에 GEMINI_API_KEY와"
      " APP_PASSWORD를 입력해 주세요."
  )
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
    st.text_input(
        "접속 비밀번호를 입력하세요",
        type="password",
        on_change=password_entered,
        key="password",
    )
    return False
  elif not st.session_state["password_correct"]:
    st.title("🔒 CanNest 통합 업무 시스템")
    st.text_input(
        "접속 비밀번호를 입력하세요",
        type="password",
        on_change=password_entered,
        key="password",
    )
    st.error("비밀번호가 틀렸습니다.")
    return False
  return True


if not check_password():
  st.stop()

# ==========================================
# 1. API 키 및 모델 설정 (공통 - 모든 메뉴)
# ==========================================
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]

@st.cache_resource(show_spinner=False)
def get_gemini_client(api_key):
    return gg.Client(api_key=api_key)

if NEW_SDK:
    gclient = get_gemini_client(GEMINI_API_KEY)
else:
    genai.configure(api_key=GEMINI_API_KEY)

# 생각 단계 길이: low = 빠름 (질문지 칸 채우기는 깊은 추론보다 정확한 옮겨 적기가 중요)
THINKING_LEVEL = "low"

# [수정] Google이 구모델을 계속 종료하므로(2.5/3.5/3.6 → 신규 사용자 차단),
# 최신 모델 우선 + 모두 막히면 이 API 키로 쓸 수 있는 최신 Flash 모델을 자동으로 찾아 사용
CANDIDATE_MODELS = ['gemini-3.8-flash', 'gemini-flash-latest', 'gemini-3.7-flash']

def _is_model_gone(msg):
    m = msg.lower()
    return ("404" in m or "not found" in m or "no longer available" in m
            or "not supported" in m or "deprecated" in m)

@st.cache_data(ttl=6 * 3600, show_spinner=False)
def discover_flash_models():
    """API 키로 사용 가능한 모델 목록에서 텍스트 생성용 Flash 모델을 최신 버전 순으로 반환."""
    found = []
    try:
        models = gclient.models.list() if NEW_SDK else genai.list_models()
        for m in models:
            name = m.name.split("/")[-1]
            methods = (getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", None) or [])
            if "generateContent" not in methods or "flash" not in name:
                continue
            if any(x in name for x in ("lite", "image", "tts", "live", "audio", "embedding", "exp", "preview")):
                continue
            ver = re.findall(r"gemini-(\d+(?:\.\d+)?)", name)
            found.append((float(ver[0]) if ver else 0.0, name))
    except Exception:
        return []
    found.sort(reverse=True)
    return [n for _, n in found]

def clean_text(s):
    """
    [수정] PDF/Word에서 읽은 글자 중 깨진 특수문자(쪼개진 이모지 등 'surrogate')를 정리.
    이게 남아 있으면 Gemini 전송 시 "'utf-8' codec can't encode ... surrogates not allowed" 오류가 납니다.
    - 짝이 맞는 조각은 원래 글자로 복원, 짝 없는 조각은 제거
    """
    if not isinstance(s, str):
        return s
    try:
        s.encode("utf-8")
        return s  # 문제 없으면 그대로
    except UnicodeEncodeError:
        pass
    fixed = s.encode("utf-16", "surrogatepass").decode("utf-16", "replace")
    return fixed.replace("\ufffd", "")

def _clean_contents(contents):
    out = []
    for c in contents:
        if isinstance(c, str):
            out.append(clean_text(c))
        elif isinstance(c, dict) and isinstance(c.get("data"), str):
            out.append({**c, "data": clean_text(c["data"])})
        else:
            out.append(c)
    return out

def _to_new_sdk_parts(contents):
    parts = []
    for c in contents:
        if isinstance(c, dict) and "data" in c:
            data = c["data"] if isinstance(c["data"], bytes) else str(c["data"]).encode("utf-8")
            parts.append(gtypes.Part.from_bytes(data=data, mime_type=c.get("mime_type", "application/octet-stream")))
        else:
            parts.append(c)
    return parts

def _call_model(model_name, contents, json_mode, thinking_level):
    if NEW_SDK:
        cfg = {}
        if json_mode:
            cfg["response_mime_type"] = "application/json"
        if thinking_level:
            cfg["thinking_config"] = gtypes.ThinkingConfig(thinking_level=thinking_level)
        if json_mode and not model_name.startswith("gemini-3") and "latest" not in model_name:
            cfg["temperature"] = 0  # 메뉴 5 + 구모델에만 (Gemini 3는 temperature 무시)
        return gclient.models.generate_content(
            model=model_name, contents=_to_new_sdk_parts(contents),
            config=gtypes.GenerateContentConfig(**cfg))
    gen_config = {}
    if json_mode:
        gen_config = {"temperature": 0, "response_mime_type": "application/json"}
    return genai.GenerativeModel(model_name, generation_config=gen_config).generate_content(contents)

def safe_generate_content(contents, json_mode=False, thinking=None):
    """
    모든 메뉴 공통 AI 호출.
    - thinking: None이면 모델 기본값 (메뉴 1~4), "low"면 빠르게 (메뉴 5)
    """
    """모델 폴백 + 일시적 오류(429/500/503) 재시도. 걸린 시간과 모델을 session_state에 기록."""
    contents = _clean_contents(contents)
    last_error = None
    # 지난번에 성공한 모델을 먼저 시도 → 매번 종료된 모델부터 두드리지 않음
    working = st.session_state.get("gemini_working_model")
    order = ([working] if working else []) + CANDIDATE_MODELS
    tried = set()

    def model_iter():
        for n in order:
            yield n
        for n in discover_flash_models():  # 위 후보가 모두 막혔을 때만 목록 조회
            yield n

    for model_name in model_iter():
        if model_name in tried:
            continue
        tried.add(model_name)
        use_thinking = thinking if NEW_SDK else None
        attempt = 0
        while attempt < 3:
            t0 = time.time()
            try:
                resp = _call_model(model_name, contents, json_mode, use_thinking)
                st.session_state["gemini_working_model"] = model_name
                st.session_state["gemini_last_run"] = {"model": model_name, "seconds": round(time.time() - t0, 1)}
                return resp
            except Exception as e:
                last_error = e
                msg = str(e).lower()
                if any(k in msg for k in ("credits are depleted", "prepayment", "billing", "api key not valid",
                                          "api_key_invalid", "permission denied", "quota exceeded for quota metric")):
                    # 결제/키 문제는 재시도해도 소용없음 → 바로 안내
                    raise RuntimeError(
                        "Gemini API 결제 또는 API 키 문제로 분석할 수 없습니다. "
                        "AI Studio(https://ai.studio/projects)에서 크레딧 잔액·결제 설정·API 키를 확인해 주세요. "
                        f"(원본 오류: {e})") from e
                if use_thinking and "thinking" in msg:
                    use_thinking = None  # 생각 단계 설정을 지원하지 않는 모델 → 설정 없이 다시
                    continue
                if _is_model_gone(msg):
                    break  # 다음 모델로
                if any(k in msg for k in ["429", "500", "503", "unavailable", "overloaded", "deadline", "resource exhausted"]):
                    attempt += 1
                    time.sleep(2 * attempt)
                    continue  # 같은 모델 재시도
                raise
        # 재시도 소진 또는 404 → 다음 모델
    if last_error is not None and _is_model_gone(str(last_error)):
        raise RuntimeError(
            f"사용 가능한 Gemini 모델을 찾지 못했습니다 (시도: {', '.join(tried)}). "
            f"코드의 CANDIDATE_MODELS에 Google이 안내하는 최신 모델명을 추가해 주세요. 원본 오류: {last_error}")
    raise last_error

def parse_json_response(text):
    """```json 펜스, 앞뒤 설명문이 붙어도 JSON만 추출."""
    t = (text or "").strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t, flags=re.IGNORECASE)
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        start, end = t.find("{"), t.rfind("}")
        if start != -1 and end > start:
            return json.loads(t[start:end + 1])
        raise

# ==========================================
# 2. MS Word (.doc / .docx) 텍스트 추출 엔진
# ==========================================
def read_word_document_text(file_bytes, file_name=""):
  try:
    doc_obj = docx.Document(io.BytesIO(file_bytes))
    text_list = [p.text for p in doc_obj.paragraphs if p.text.strip()]
    for table in doc_obj.tables:
      for row in table.rows:
        row_txt = " | ".join(
            [cell.text.strip() for cell in row.cells if cell.text.strip()]
        )
        if row_txt:
          text_list.append(row_txt)
    full_text = "\n".join(text_list)
    if len(full_text.strip()) > 10:
      return full_text
  except Exception:
    pass

  try:
    raw_bytes = file_bytes
    text_utf16 = raw_bytes.decode("utf-16le", errors="ignore")
    cleaned = re.sub(r"[^\w\s\.,\-\:\/@\(\)\?\!\'\"]+", " ", text_utf16)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if len(cleaned) > 30:
      return cleaned
  except Exception:
    pass

  try:
    text_ascii = file_bytes.decode("latin1", errors="ignore")
    cleaned = re.sub(r"[^\w\s\.,\-\:\/@\(\)\?\!\'\"]+", " ", text_ascii)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned
  except Exception:
    return ""


# ==========================================
# 3. 공통 캐싱 및 스마트 글자 크기 조절 함수
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
  file_bytes = file_obj.getvalue()
  fname = file_obj.name.lower()

  if fname.endswith((".doc", ".docx")):
    text = read_word_document_text(file_bytes, file_obj.name)
    img = Image.new("RGB", (800, 1000), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    display_str = (
        f"[Word Document: {file_obj.name}]\n\n"
        + (text[:800] if text else "Word Document Content")
    )
    draw.text((40, 40), display_str, fill=(0, 0, 0))
    return img
  elif file_obj.type == "application/pdf" or fname.endswith(".pdf"):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
  else:
    img = Image.open(io.BytesIO(file_bytes))
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
  s = surname.strip()
  g = given_name.strip()
  if not s and not g:
    return ""
  if not s:
    return g
  if not g:
    return s
  return f"{g} {s}"


def set_smart_widget_value(
    widget, value, default_fontsize=11, min_fontsize=5.5
):
  val_str = str(value) if value is not None else ""
  widget.field_value = val_str

  if hasattr(widget, "field_flags") and widget.field_flags:
    widget.field_flags &= ~1

  try:
    widget.text_font = "Cour"
  except Exception:
    try:
      widget.text_font = "Courier"
    except Exception:
      pass

  if val_str and hasattr(widget, "rect"):
    box_width = widget.rect.width - 4
    if box_width > 0:
      try:
        if os.path.exists("cour.ttf"):
          font = fitz.Font(fontfile="cour.ttf")
        else:
          font = fitz.Font("courier")

        len_at_default = font.text_length(val_str, fontsize=default_fontsize)
        if len_at_default > box_width:
          len_at_1 = font.text_length(val_str, fontsize=1)
          if len_at_1 > 0:
            scaled_size = box_width / len_at_1
            widget.text_fontsize = max(
                min_fontsize, min(default_fontsize, scaled_size)
            )
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


def embed_courier_in_doc(doc):
  if not os.path.exists("cour.ttf"):
    return

  try:
    page = doc[0]
    font_xrefs = {}
    for alias in ["Cour", "Courier", "CoNi", "CourierNew", "CourierNewPSMT"]:
      try:
        xref = page.insert_font(fontfile="cour.ttf", fontname=alias)
        font_xrefs[alias] = xref
      except Exception:
        pass

    catalog_xref = doc.pdf_catalog()
    acroform_res = doc.xref_get_key(catalog_xref, "AcroForm")

    acroform_xref = None
    if acroform_res[0] == "xr":
      acroform_xref = int(acroform_res[1].split()[0])
    elif acroform_res[0] == "dict":
      acroform_xref = catalog_xref

    if acroform_xref and font_xrefs:
      font_dict_entries = " ".join(
          [f"/{alias} {xref} 0 R" for alias, xref in font_xrefs.items()]
      )
      doc.xref_set_key(acroform_xref, "DR/Font", f"<< {font_dict_entries} >>")
  except Exception:
    pass


def page_to_image_pdf_page(new_doc, src_doc, page_idx, rotation=0, dpi=200):
  page = src_doc.load_page(page_idx)
  zoom = dpi / 72.0
  pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
  img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

  if rotation % 360 != 0:
    img = img.rotate(-rotation, expand=True)

  buf = io.BytesIO()
  img.save(buf, format="JPEG", quality=92, optimize=True)
  buf.seek(0)

  page_w = img.width * 72.0 / dpi
  page_h = img.height * 72.0 / dpi
  pdf_page = new_doc.new_page(width=page_w, height=page_h)
  pdf_page.insert_image(pdf_page.rect, stream=buf.getvalue())


def sanitize_and_unlock_pdf(pdf_bytes):
  try:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
      for widget in list(page.widgets()):
        field_type_str = getattr(widget, "field_type_string", "").lower()
        if (
            "sig" in field_type_str
            or widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE
        ):
          page.delete_widget(widget)
        else:
          if hasattr(widget, "field_flags"):
            widget.field_flags &= ~1
            widget.update()
            
    doc.need_appearances(True)
    doc.set_metadata({})

    out_buf = io.BytesIO()
    doc.save(out_buf, deflate=True)
    doc.close()
    return out_buf.getvalue()
  except Exception:
    return pdf_bytes


def prepare_document_for_gemini(file_bytes, mime_type, file_name=""):
  ext = os.path.splitext(file_name)[1].lower() if file_name else ""
  if (
      "word" in mime_type.lower()
      or "doc" in mime_type.lower()
      or ext in [".doc", ".docx"]
  ):
    word_text = read_word_document_text(file_bytes, file_name)
    if word_text:
      return [
          f"\n--- [Word Document: {file_name}] ---\n{word_text[:20000]}\n"
      ]

  if "pdf" in mime_type.lower():
    try:
      doc = fitz.open(stream=file_bytes, filetype="pdf")
      text = ""
      for page in doc:
        text += page.get_text("text") + "\n"
        for widget in page.widgets():
            if widget.field_value:
                text += f"{widget.field_name}: {widget.field_value}\n"

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
    except Exception:
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
    age = (
        today.year
        - birth_date.year
        - ((today.month, today.day) < (birth_date.month, birth_date.day))
    )
    return age < 19
  except Exception:
    return True


# ==========================================
# 4. PDF 서식 채우기 로직
# ==========================================
def extract_imm5476_info(image):
  # 💡 AI 프롬프트 업데이트: 이메일과 주소/전화번호 추출 지시 추가
  prompt = """
    Analyze this document (which may be an identity document or an existing application form like IMM 5476) carefully.
    Extract the following details into exact JSON structure:
    - surname: Family name converted to Title Case (First letter capitalized, e.g., 'KIM' -> 'Kim')
    - given_name: Given names converted to Title Case (e.g., 'EUN SUN' -> 'Eun Sun')
    - dob: Date of birth in YYYY-MM-DD format (if available, else "")
    - uci: UCI numbers only if present (10 digits or 8 digits), else empty string
    - email: Email address if present, else empty string
    - address_phone: Residential/Mailing address or telephone number if present, else empty string
    Return ONLY raw valid JSON object without markdown or code formatting.
    """
  try:
    response = safe_generate_content([prompt, image])
    clean_text = (
        response.text.strip().replace("```json", "").replace("```", "")
    )
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
    - {'Remaining images are accompanying family members.' if (len(images) > (1 if has_non_acc else 0)) else 'No family passports provided.'}

    For EACH passport image, extract:
    - surname: Surname / Family name converted to Title Case
    - given_name: Given name(s) converted to Title Case
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
    clean_text = (
        response.text.strip().replace("```json", "").replace("```", "")
    )
    return json.loads(clean_text)
  except Exception as e:
    st.error(f"여권 일괄 추출 오류: {e}")
    return None


def extract_case_prep_info(tmpl_bytes, client_files):
  prompt = """
    You are an expert Canadian immigration case prep assistant.
    The FIRST document is a BLANK reference IRCC IMM form template.
    The REMAINING attached documents are client materials.

    Carefully scan ALL attached client documents to extract all required information matching the IMM form fields.

    [CRITICAL DETAILED EXTRACTION RULES]

    1. NAMES & GENERAL:
       - Convert ALL CAPS names to Title Case (e.g., 'Eun Sun Kim', 'Alexis Antonio').

    2. CONTACT INFORMATION:
       - Extract Telephone Number thoroughly: Type (Cell/Home/Business), Country Code, Area Code & Number, Extension.
       - Extract Mailing Address & Residential Address completely.

    3. MARITAL STATUS (Current & Previous):
       - Current Marital Status: Status, Date of Marriage/Common-Law (YYYY-MM-DD), Current Spouse Full Name, Is Spouse Canadian Citizen/PR?
       - Previous Marital Status: Has client been previously married/common-law? (Yes/No). If Yes, extract Previous Spouse Full Name, Relationship Type, From Date (YYYY-MM-DD), To Date (YYYY-MM-DD), Previous Spouse Date of Birth (YYYY-MM-DD).

    4. EDUCATION (ALL Post-Secondary Entries):
       - DO NOT extract only the highest level. Extract ALL post-secondary education history found in questionnaires, WES, transcripts, or resumes.
       - For EACH education entry, format fields with entry labels like 'Education Entry 1 - Dates', 'Education Entry 1 - Field & Level of Study', 'Education Entry 1 - School Name', 'Education Entry 1 - Location'.
       - For EACH entry, extract:
         * From Date (YYYY-MM) & To Date (YYYY-MM)
         * Field and Level of Study
         * School / Facility Name
         * City/Town, Country or Territory, Province/State (if applicable)

    5. EMPLOYMENT (10-Year History with Full Location):
       - Extract ALL employment/activity entries for the past 10 years.
       - For EACH entry, format fields with entry labels like 'Employment Entry 1 - Dates', 'Employment Entry 1 - Job Title', 'Employment Entry 1 - Company Name', 'Employment Entry 1 - Location'.
       - For EACH entry, extract:
         * From Date (YYYY-MM) & To Date (YYYY-MM)
         * Activity / Occupation / Job Title
         * Company / Employer Name
         * City/Town, Country or Territory, Province/State

    6. BACKGROUND INFORMATION (Refusals, Criminality, Military, Medical):
       - Refusal / Visa History (Q2): Has client ever overstayed status, been refused a visa/permit, or been denied entry/ordered to leave Canada or any other country? (Yes/No). Provide full details/explanation if Yes.
       - Criminal History (Q3): Has client ever committed, been arrested for, charged with, or convicted of any criminal offense? (Yes/No). Provide full details if Yes.
       - Military / Police / Defense Service (Q4): Did client serve in military, militia, police, or civil defense unit? (Yes/No). If Yes, provide exact dates of service and countries served.
       - Medical History (Q1): Any TB history or physical/mental disorder requiring social/health services? (Yes/No) + details.

    7. CURRENT RESIDENCE PERMIT DATES:
       - Status, From date (YYYY-MM-DD), and To date (YYYY-MM-DD) MUST be extracted directly from current permit/visa documents.

    8. MISMATCH DETECTION:
       - If there is a mismatch across documents for any field, set value strictly as:
         "⚠️ 정보 불일치 (재확인 필요): [Doc A] ValueA vs [Doc B] ValueB"

    Return ONLY a raw valid JSON object in this exact shape:
    {
      "sections": [
        {
          "section": "Section Name (e.g. Personal Details, Contact Information, Education, Employment, Background Information)",
          "fields": [
            { "field": "Field Label", "value": "Extracted Value or Mismatch Warning", "source": "Source Doc Name(s)" }
          ]
        }
      ]
    }
    """

  contents = [prompt]
  contents.extend(
      prepare_document_for_gemini(
          tmpl_bytes, "application/pdf", "Blank_IMM_Form.pdf"
      )
  )
  contents.extend(batch_process_client_files(client_files))

  try:
    response = safe_generate_content(contents)
    clean_text = (
        response.text.strip().replace("```json", "").replace("```", "")
    )
    return json.loads(clean_text)
  except Exception as e:
    st.error(f"서류 정리 오류: {e}")
    return None


def fill_imm5476(template_bytes, data):
  doc = fitz.open(stream=template_bytes, filetype="pdf")
  target_data = {
      "surname": data.get("surname", ""),
      "given": data.get("given_name", ""),
      "dob": data.get("dob", ""),
      "email": data.get("email", ""),
      "address_phone": data.get("address_phone", ""),
      "uci": data.get("uci", "").replace("-", ""),
      "signDate": data.get("signDate", ""),
  }
  flags = {
      "surname": False,
      "given": False,
      "dob": False,
      "email": False,
      "address_phone": False,
      "uci": False,
  }

  page_date_counters = {}

  for page_idx, page in enumerate(doc):
    page_date_counters[page_idx] = 0
    for widget in page.widgets():
      field_name = widget.field_name
      if not field_name:
        continue
      fname_lower = field_name.lower()

      val_to_set = None
      if "family name" in fname_lower and not flags["surname"]:
        val_to_set = target_data["surname"]
        flags["surname"] = True
      elif "given name" in fname_lower and not flags["given"]:
        val_to_set = target_data["given"]
        flags["given"] = True
      elif "date of birth" in fname_lower and not flags["dob"]:
        val_to_set = target_data["dob"]
        flags["dob"] = True
      elif "email" in fname_lower and not flags["email"]:
        val_to_set = target_data["email"]
        flags["email"] = True
      elif (
          ("telephone" in fname_lower or "address" in fname_lower)
          and page_idx == 0
          and not flags["address_phone"]
      ):
        val_to_set = target_data["address_phone"]
        flags["address_phone"] = True
      elif (
          "uci" in fname_lower or "unique client identifier" in fname_lower
      ) and not flags["uci"]:
        val_to_set = target_data["uci"]
        flags["uci"] = True
      elif "date" in fname_lower and "birth" not in fname_lower:
        page_date_counters[page_idx] += 1
        if page_idx == 2 and page_date_counters[page_idx] == 1:
          val_to_set = target_data["signDate"]
        elif page_idx == 3 and page_date_counters[page_idx] == 1:
          val_to_set = target_data["signDate"]

      if val_to_set is not None:
        set_smart_widget_value(widget, val_to_set, default_fontsize=9)

  embed_courier_in_doc(doc)
  doc.need_appearances(True)

  output_pdf = io.BytesIO()
  doc.save(output_pdf, deflate=True)
  doc.close()
  output_pdf.seek(0)
  return output_pdf


def fill_consent_letter(template_bytes, data):
  doc = fitz.open(stream=template_bytes, filetype="pdf")
  children = data.get("children", [])
  num_pages_needed = max(1, (len(children) + 2) // 3)
  for _ in range(num_pages_needed - 1):
    doc.insert_pdf(doc, from_page=0, to_page=0)

  for page_num in range(num_pages_needed):
    page = doc[page_num]
    page_children = children[page_num * 3 : (page_num + 1) * 3]
    child_widgets = [
        ("Information about travelling children", "yyyymmdd"),
        ("1_2", "2_2"),
        ("1_3", "2_3"),
    ]

    for widget in page.widgets():
      fname = widget.field_name.strip() if widget.field_name else ""
      if not fname:
        continue
      fname_lower = fname.lower()

      field_type_str = getattr(widget, "field_type_string", "").lower()
      if (
          "check" in field_type_str
          or "radio" in field_type_str
          or "check box" in fname_lower
          or "checkbox" in fname_lower
          or "alone" in fname_lower
      ):
        try:
          widget.field_value = "Off"
          widget.update()
        except Exception:
          pass
        continue

      val_to_set = None
      if fname == "1":
        val_to_set = data.get("non_acc_name", "")
      elif fname == "2":
        val_to_set = data.get("non_acc_address", "")
      elif fname == "3":
        val_to_set = data.get("non_acc_phone", "")
      elif fname == "email":
        val_to_set = data.get("non_acc_email", "")
      elif (
          fname
          == "This child or these children hashave my or our consent to"
          " travel with"
      ):
        val_to_set = data.get("acc_name", "")
      elif fname == "Relationship with Children 1":
        val_to_set = data.get("acc_relationship", "")
      elif fname == "Relationship with Children 2":
        val_to_set = data.get("acc_passport", "")
      elif fname == "I give my consent for this child to travel to":
        val_to_set = "Canada"
      elif fname == "2_4" or fname_lower == "to stay with":
        val_to_set = data.get("acc_name", "")
      elif fname == "At the following addresses 1":
        val_to_set = data.get("trip_address", "")
      elif fname == "At the following addresses 2":
        val_to_set = data.get("trip_phone", "")
      elif fname == "email_2":
        val_to_set = data.get("trip_email", "")
      elif fname == "yyyymmdd_2":
        val_to_set = data.get("sign_date", "")
      elif (
          fname == "1_4"
          or "travel date" in fname_lower
          or fname_lower == "date"
          or "from" in fname_lower
          or "departure" in fname_lower
          or "start date" in fname_lower
      ):
        val_to_set = data.get("trip_date", "")
      elif (
          fname_lower == "to"
          or "return" in fname_lower
          or "end date" in fname_lower
      ):
        val_to_set = ""
      else:
        for idx, (name_key, dob_key) in enumerate(child_widgets):
          if idx < len(page_children):
            if fname == name_key:
              val_to_set = page_children[idx].get("name", "")
            elif fname == dob_key:
              val_to_set = page_children[idx].get("dob", "")

      if val_to_set is not None:
        set_smart_widget_value(widget, val_to_set, default_fontsize=11)

  embed_courier_in_doc(doc)
  doc.need_appearances(True)

  output_pdf = io.BytesIO()
  doc.save(output_pdf, deflate=True)
  doc.close()
  output_pdf.seek(0)
  return output_pdf


# ==========================================
# 5. CRM 스마트 압축 및 PDF 안전 병합 엔진
# ==========================================
def process_and_compress_file(
    file_bytes, mime_type, target_filename, already_sanitized: bool = False
):
  is_jpeg = target_filename.lower().endswith((".jpg", ".jpeg"))

  if is_jpeg:
    img = Image.open(io.BytesIO(file_bytes))
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
      img = img.convert("RGB")

    max_dim = max(img.width, img.height)
    if max_dim > 2000:
      ratio = 2000.0 / float(max_dim)
      img = img.resize(
          (int(img.width * ratio), int(img.height * ratio)),
          Image.Resampling.LANCZOS,
      )

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=80, optimize=True)
    buf.seek(0)
    return buf.getvalue(), "image/jpeg"

  else:
    if "pdf" in mime_type.lower():
      doc = fitz.open(stream=file_bytes, filetype="pdf")

      total_text_len = 0
      has_widgets = False

      for page_idx in range(min(len(doc), 5)):
        page = doc.load_page(page_idx)
        text = page.get_text("text").strip()
        total_text_len += len(text)
        if list(page.widgets()):
          has_widgets = True
        if total_text_len > 50:
          break

      if total_text_len > 50 or has_widgets:
        doc.close()
        if already_sanitized:
          return file_bytes, "application/pdf"
        else:
          return sanitize_and_unlock_pdf(file_bytes), "application/pdf"

      new_doc = fitz.open()
      target_dpi = 150
      quality = 65

      for page in doc:
        rot = page.rotation % 360
        if rot in (90, 270):
          eff_width = page.rect.height
          eff_height = page.rect.width
        else:
          eff_width = page.rect.width
          eff_height = page.rect.height

        zoom = target_dpi / 72.0
        if max(eff_width, eff_height) > 2000:
          zoom = 1.0

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        img_buf = io.BytesIO()
        img.save(img_buf, format="JPEG", quality=quality, optimize=True)
        img_buf.seek(0)

        new_page_width = pix.width / zoom
        new_page_height = pix.height / zoom
        pdf_page = new_doc.new_page(
            width=new_page_width, height=new_page_height
        )
        pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())

      output_pdf = io.BytesIO()
      new_doc.save(output_pdf, deflate=True, garbage=4)
      new_doc.close()
      doc.close()

      compressed_bytes = output_pdf.getvalue()
      final_bytes = (
          compressed_bytes
          if len(compressed_bytes) < len(file_bytes)
          else file_bytes
      )
      if already_sanitized:
        return final_bytes, "application/pdf"
      else:
        return sanitize_and_unlock_pdf(final_bytes), "application/pdf"

    else:
      target_dpi = 150
      quality = 70
      img = Image.open(io.BytesIO(file_bytes))
      img = ImageOps.exif_transpose(img)
      if img.mode != "RGB":
        img = img.convert("RGB")

      max_dim = max(img.width, img.height)
      if max_dim > 1800:
        ratio = 1800.0 / float(max_dim)
        img = img.resize(
            (int(img.width * ratio), int(img.height * ratio)),
            Image.Resampling.LANCZOS,
        )

      img_buf = io.BytesIO()
      img.save(img_buf, format="JPEG", quality=quality, optimize=True)
      img_buf.seek(0)

      new_doc = fitz.open()
      page_width = img.width * 72.0 / target_dpi
      page_height = img.height * 72.0 / target_dpi
      pdf_page = new_doc.new_page(width=page_width, height=page_height)

      pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())

      output_pdf = io.BytesIO()
      new_doc.save(output_pdf)
      new_doc.close()

      compressed_bytes = output_pdf.getvalue()
      if already_sanitized:
        return compressed_bytes, "application/pdf"
      else:
        return sanitize_and_unlock_pdf(compressed_bytes), "application/pdf"



# ==========================================
# 5-0. [메뉴 4] 질문지 파일명 코드 고정
# ==========================================
# [수정] 질문지 파일명이 'QA_Work Permit Extension'처럼 제각각 나오던 문제 →
# AI 결과를 한 번 더 검사해서 정해진 코드(WP / OWP / SP / EEBC / BCPNP)로 바꿈
# 순서 중요: 더 구체적인 것부터 (EEBC → EE, Working Holiday/OWP → WP, Visitor Record → TRV)
QA_CODE_RULES = [
    ("EEBC", r"eebc|express\s*entry\s*bc"),
    ("BCPNP", r"bc\s*-?\s*pnp|provincial\s*nominee|pnp"),
    ("EE", r"\bee\b|express\s*entry|익스프레스"),
    ("WH", r"\bwh\b|working\s*holiday|\biec\b|워킹\s*홀리데이|워홀"),
    ("OWP", r"\bowp\b|open\s*work"),
    ("VR", r"\bvr\b|visitor\s*record|visitor\s*extension|방문\s*연장|방문자\s*기록"),
    ("TRV", r"\btrv\b|temporary\s*resident\s*visa|visitor\s*visa|tourist\s*visa|방문\s*비자|관광\s*비자"),
    ("SP", r"\bsp\b|study|student|학생|학업"),
    ("WP", r"\bwp\b|work|pgwp|lmia|취업|워크"),
]

def normalize_qa_filename(filename):
    """'홍길동_QA_Work Permit Extension_2025.01.14.pdf' → '홍길동_QA_WP_2025.01.14.pdf'"""
    if not filename or not re.search(r"_QA_", filename, re.I):
        return filename
    m = re.match(r"^(.*?)_QA_(.*)$", filename, re.I)
    if not m:
        return filename
    name, rest = m.group(1), m.group(2)
    ext = ""
    em = re.search(r"\.(pdf|jpe?g|docx?|png)$", rest, re.I)
    if em:
        ext, rest = em.group(0), rest[:em.start()]
    dm = re.search(r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}", rest)
    date = ""
    if dm:
        y, mo, d = re.split(r"[.\-/]", dm.group(0))
        date = f"{y}.{int(mo):02d}.{int(d):02d}"
        rest = (rest[:dm.start()] + rest[dm.end():])
    type_text = rest.replace("_", " ").strip()
    code = next((c for c, pat in QA_CODE_RULES if re.search(pat, type_text, re.I)), None)
    if code is None:
        code = "확인필요"  # 종류를 알 수 없으면 화면에 '규칙 미확인 서류' 경고가 뜨도록
    return f"{name}_QA_{code}" + (f"_{date}" if date else "") + ext


# ==========================================
# 5-1. [메뉴 4] 페이지 수 제한 완화 도우미
# ==========================================
# [수정] 기존 40장 제한 → 200장.
# 모든 페이지를 AI에 '한 번에' 보내야 서류 묶기(여권+스탬프, 은행내역 여러 장 등)가 정확하므로
# 요청을 나누지 않고, 장수가 많을수록 미리보기 이미지를 작게 만들어 요청 용량 한도 안에 맞춤.
CRM_MAX_PAGES = 200
CRM_PAYLOAD_LIMIT = 15 * 1024 * 1024  # AI 요청 한 번의 이미지 총량 목표 (한도 약 20MB보다 여유 있게)

def count_total_pages(files):
    total = 0
    for f in files:
        name = f.name.lower()
        if name.endswith(".pdf") or "pdf" in (f.type or "").lower():
            try:
                d = fitz.open(stream=f.getvalue(), filetype="pdf")
                total += len(d)
                d.close()
            except Exception:
                total += 1
        else:
            total += 1
    return total

def crm_preview_settings(total_pages):
    """(PDF 확대 배율, JPEG 품질, 이미지 최대 변 길이) — 장수가 많을수록 작게."""
    if total_pages <= 40:
        return 1.2, 60, 1200   # 기존과 동일
    if total_pages <= 80:
        return 1.1, 55, 1100
    if total_pages <= 140:
        return 1.0, 50, 1000
    return 0.9, 45, 900

def shrink_previews_to_fit(pages, limit=CRM_PAYLOAD_LIMIT):
    """미리보기 총 용량이 한도를 넘으면 전체를 비율대로 줄여서 다시 저장."""
    total = sum(len(p["preview_bytes"]) for p in pages)
    tries = 0
    while total > limit and tries < 3:
        scale = (limit / total) ** 0.5 * 0.9
        for p in pages:
            img = Image.open(io.BytesIO(p["preview_bytes"]))
            img = img.resize((max(1, int(img.width * scale)), max(1, int(img.height * scale))), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format="JPEG", quality=45)
            p["preview_bytes"] = buf.getvalue()
        total = sum(len(p["preview_bytes"]) for p in pages)
        tries += 1
    return total


# ==========================================
# 5-2. [메뉴 5] 질문지 기존 정보 기입 엔진
# ==========================================
def qa_read_word_document_text(file_bytes, file_name=""):
    if not DOCX_AVAILABLE:
        return ""
    try:
        doc_obj = docx.Document(io.BytesIO(file_bytes))
        text_list = [p.text for p in doc_obj.paragraphs if p.text.strip()]
        for table in doc_obj.tables:
            for row in table.rows:
                seen = set()
                cells = []
                for cell in row.cells:  # 병합 셀 중복 제거
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    if cell.text.strip():
                        cells.append(cell.text.strip())
                if cells:
                    text_list.append(" | ".join(cells))
        return clean_text("\n".join(text_list))
    except Exception:
        return ""


def image_bytes_to_jpeg_part(file_bytes, max_dim=2000):
    """[수정] HEIC/PNG 등 이미지를 Gemini가 확실히 읽는 JPEG로 변환 (EXIF 회전 보정)."""
    img = Image.open(io.BytesIO(file_bytes))
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    if max(img.size) > max_dim:
        ratio = max_dim / float(max(img.size))
        img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return {"mime_type": "image/jpeg", "data": buf.getvalue()}


# ------------------------------------------
# 1-1. 대상 PDF 폼 필드 목록 추출 (고유 키 부여)
# ------------------------------------------
BUTTON_TYPES = (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON)
CHOICE_TYPES = (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX)

def _safe_on_state(w):
    try:
        s = w.on_state()
        if s and str(s).lower() != "off":
            return str(s)
    except Exception:
        pass
    return None

# 질문지 안의 '구역 제목' 패턴
SECTION_HEADER_RE = re.compile(
    r"(개인\s*정보|가족\s*관계\s*정보|personal\s+(information|details)|family\s+(information|members)|background\s+information)",
    re.IGNORECASE)
# '누구에 대한 구역인지'를 정하는 제목: 사람을 가리키는 단어로 시작해야 함
# (예: '[ SPOUSE - BACKGROUND INFORMATION ]', '배우자의 개인 정보', '자녀(1)의 개인 정보', '[ DEPENDANT CHILD 1 ]')
# '[ Required for Applicants Aged 18 or Older ]'처럼 중간에 단어가 섞인 제목은 사람 구역으로 보지 않음
PERSON_HEADER_RE = re.compile(
    r"^[\[\(\s]*(main\s+applicant|principal\s+applicant|applicant|spouse|common[\s-]*law\s+partner|partner|"
    r"dependa?e?nt\s+child|child|신청자|주\s*신청자|배우자|자녀)", re.IGNORECASE)

def _person_from_text(t):
    t = (t or "").lower()
    m = re.search(r"자녀\s*\(?(\d+)\)?|child\s*\(?(\d+)\)?", t)
    if m:
        return f"CHILD {m.group(1) or m.group(2)}"
    if re.match(r"^[\[\(\s]*(자녀|dependa?e?nt|child)", t):
        return "CHILD"
    if re.match(r"^[\[\(\s]*(배우자|spouse|common[\s-]*law|partner)", t):
        return "SPOUSE"
    return "MAIN APPLICANT"

def detect_sections(doc):
    """
    구역 제목 위치 목록 [(page, y, 제목, 사람 또는 None)].
    - 사람 제목(신청자/배우자/자녀)은 이후 칸들의 '누구' 를 바꿈
    - 일반 제목([ TRAVEL HISTORY ], [ FAMILY INFORMATION ] 등)은 구역 이름만 바꾸고 '누구'는 유지
    """
    headers = []
    for p_idx, page in enumerate(doc):
        for b in page.get_text("blocks"):
            first_line = clean_text((b[4] or "").strip().split("\n")[0].strip())
            if not first_line or len(first_line) > 80:
                continue
            is_person = bool(PERSON_HEADER_RE.search(first_line)) and (
                SECTION_HEADER_RE.search(first_line) or first_line.lstrip().startswith("[")
                or re.search(r"정보|information|details", first_line, re.I))
            is_general = bool(SECTION_HEADER_RE.search(first_line)) or bool(re.match(r"^\[\s*[A-Z][A-Z &/-]{3,}\s*\]", first_line))
            if is_person or is_general:
                title = re.split(r"\s+[-–]\s+If applicable|\(\*|\s+-\s+모든", first_line)[0].strip()
                headers.append((p_idx, b[1], title, _person_from_text(first_line) if is_person else None))
    headers.sort(key=lambda h: (h[0], h[1]))
    return headers

def person_for_section(title):
    """(구버전 호환) 제목 하나로 사람 판단."""
    t = (title or "")
    m0 = SECTION_HEADER_RE.search(t)
    if m0:
        t = t[:m0.start()]
    return _person_from_text(t)

def person_for(headers, page_idx, y):
    person = "MAIN APPLICANT"
    for hp, hy, title, who in headers:
        if hp < page_idx or (hp == page_idx and hy <= y + 2):
            if who:
                person = who
        else:
            break
    return person

def section_for(headers, page_idx, y):
    current = ""
    for hp, hy, title, who in headers:
        if hp < page_idx or (hp == page_idx and hy <= y + 2):
            current = title
        else:
            break
    return current

# ------------------------------------------
# 1-1a. '현재 정보' 칸 (손님이 직접 최신으로 작성) → 항상 비워둠
# ------------------------------------------
CURRENT_ADDRESS_RE = re.compile(
    r"current\s+(residential\s+|home\s+)?address|residential\s+address|present\s+address|home\s+address",
    re.IGNORECASE)
# 한글 라벨은 글자 사이가 띄어진 PDF가 있어 공백을 모두 지우고 비교 (예: '현재 거 주지 주소')
CURRENT_ADDRESS_KO_RE = re.compile(r"현재(거주지|집주소|거주하는곳|사는곳)")

def _is_current_address_label(text):
    if not text:
        return False
    if "사망" in text:  # 가족표 '현재 주소 (사망 시…)'는 제외
        return False
    return bool(CURRENT_ADDRESS_RE.search(text) or CURRENT_ADDRESS_KO_RE.search(re.sub(r"\s+", "", text)))
RECENT_ENTRY_PHRASES = ["가장 최근 입국", "최근 입국", "가장 최근에 들어온", "최근에 들어온", "최근 들어온",
                        "most recent entry", "last entry", "latest entry", "most recent arrival"]
FIRST_ENTRY_PHRASES = ["최초 입국", "처음 입국", "처음 들어온", "최초로 들어온",
                       "first entry", "original entry", "first arrival"]

def mark_always_blank(doc, fields):
    """
    현재 거주지 주소, 캐나다 가장 최근 입국 정보 칸에 always_blank 표시.
    - 주소: 칸 왼쪽 라벨/툴팁이 '현재 거주지 주소' 등인 칸 (가족표의 '현재 주소'는 제외)
    - 최근 입국: 칸 위쪽 100pt 안의 제목 중 가로로 가장 가까운 것이 '최근 입국'인 칸
    """
    phrase_cache = {}
    def phrases_on(p_idx):
        if p_idx not in phrase_cache:
            page = doc[p_idx]
            found = []
            # 단어 단위로 문구 찾기 (일부 PDF는 글자 검색 위치가 틀어져 있어 search_for를 쓰지 않음)
            words = page.get_text("words")
            lines = {}
            for w in words:
                lines.setdefault((w[5], w[6]), []).append(w)  # (block, line)
            for kind, plist in (("recent", RECENT_ENTRY_PHRASES), ("first", FIRST_ENTRY_PHRASES)):
                seen_rects = []
                targets = [re.sub(r"\s+", "", ph.lower()) for ph in plist]
                for lw in lines.values():
                    lw.sort(key=lambda w: w[0])
                    for i in range(len(lw)):
                        acc = ""
                        for j in range(i, len(lw)):
                            acc += lw[j][4].lower()
                            hit = next((t for t in targets if acc.startswith(t) or acc == t), None)
                            if hit:
                                r = fitz.Rect(lw[i][:4])
                                for k in range(i, j + 1):
                                    r |= fitz.Rect(lw[k][:4])
                                if not any(r.intersects(x) for x in seen_rects):
                                    seen_rects.append(r)
                                    found.append((kind, r))
                                break
                            if not any(t.startswith(acc) for t in targets):
                                break
            phrase_cache[p_idx] = found
        return phrase_cache[p_idx]

    for f in fields:
        if f["type"] != fitz.PDF_WIDGET_TYPE_TEXT:
            continue
        page = doc[f["page"]]
        x0, y0, x1, y1 = f["rect"]
        left = " ".join(page.get_text("text", clip=fitz.Rect(max(0, x0 - 220), y0 - 1, x0, y1 + 1)).split())
        if (_is_current_address_label(f.get("printed_label") or "") or _is_current_address_label(left)
                or _is_current_address_label(f.get("label") or "")):
            f["always_blank"] = "현재 거주지 주소"
            continue
        cx = (x0 + x1) / 2
        above = [(kind, r) for kind, r in phrases_on(f["page"]) if r.y0 <= y0 + 2 and y0 - r.y1 < 100]  # 위쪽 또는 같은 높이 라벨
        if not above:
            continue
        # 바로 위 제목 줄을 먼저 고르고, 그 줄 안에서 가로로 가장 가까운 제목을 선택
        closest_y = max(r.y1 for _, r in above)
        same_row = [(abs(cx - (r.x0 + r.x1) / 2), kind) for kind, r in above if closest_y - r.y1 <= 6]
        if min(same_row)[1] == "recent":
            f["always_blank"] = "가장 최근 입국 정보"

    # 현재 주소 바로 옆/아래의 '우편번호'만 있는 칸도 현재 주소의 일부 → 비움
    addr = [f for f in fields if f.get("always_blank") == "현재 거주지 주소"]
    for f in fields:
        if f.get("always_blank") or f["type"] != fitz.PDF_WIDGET_TYPE_TEXT:
            continue
        lab = " ".join(x for x in (f.get("printed_label"), f.get("label")) if x)
        if not re.search(r"우\s*편\s*번\s*호|postal|zip", lab, re.I):
            continue
        for a in addr:
            if a["page"] != f["page"]:
                continue
            ax0, ay0, ax1, ay1 = a["rect"]
            fx0, fy0, fx1, fy1 = f["rect"]
            same_row = fy0 < ay1 - 2 and fy1 > ay0 + 2
            just_below = 0 <= fy0 - ay1 < 30 and fx0 < ax1 and fx1 > ax0
            if same_row or just_below:
                f["always_blank"] = "현재 거주지 주소"
                break

def compute_printed_labels(doc, fields):
    """
    [수정] 각 칸에 '인쇄된 질문 라벨'을 정확히 하나 찾음.
    - 같은 줄 왼쪽 라벨 (단, 왼쪽에 다른 칸이 있으면 그 칸 오른쪽 글자만)
    - 왼쪽에 없으면 바로 위 라벨 (위에 다른 칸이 있으면 그 칸 아래 글자만)
    칸 내부 이름(예: 'Telephone')이 잘못 붙어 있어도, 실제 화면의 라벨로 판단하게 하기 위함.
    """
    by_page = {}
    for f in fields:
        by_page.setdefault(f["page"], []).append(f)
    for p_idx, flist in by_page.items():
        page = doc[p_idx]
        words = page.get_text("words")
        rects = [f["rect"] for f in flist]
        for f in flist:
            if f["type"] in BUTTON_TYPES:
                f["printed_label"] = ""
                f["label_side"] = ""
                continue
            x0, y0, x1, y1 = f["rect"]
            # 왼쪽 막는 칸
            block_x = max([r[2] for r in rects if r[2] <= x0 + 1 and r[1] < y1 - 2 and r[3] > y0 + 2 and r != f["rect"]] + [x0 - 260])
            left = [w for w in words if w[2] <= x0 + 2 and w[0] >= block_x - 1
                    and y0 - 2 <= (w[1] + w[3]) / 2 <= y1 + 2]
            label = " ".join(w[4] for w in sorted(left, key=lambda w: (round(w[1] / 3), w[0])))
            f["label_side"] = "left" if label.strip() else ""
            if not label.strip():
                block_y = max([r[3] for r in rects if r[3] <= y0 + 1 and r[0] < x1 - 2 and r[2] > x0 + 2 and r != f["rect"]] + [y0 - 40])
                above = [w for w in words if w[3] <= y0 + 2 and w[1] >= block_y - 1
                         and x0 - 4 <= (w[0] + w[2]) / 2 <= x1 + 4]
                label = " ".join(w[4] for w in sorted(above, key=lambda w: (round(w[1] / 3), w[0])))
            f["printed_label"] = clean_text(label.strip())[:120]

def extract_pdf_form_fields(pdf_bytes):
    """
    [핵심 수정] 같은 이름을 공유하는 라디오/체크박스(예: Yes·No 한 그룹)는
    이름만으로는 어느 버튼인지 구분할 수 없어, 기존 코드는 그룹의 '모든' 버튼을 켰고
    결국 마지막 버튼(주로 'No')이 선택되는 반대 결과가 나왔습니다.
    → 버튼마다 '이름 [상태값]' 형태의 고유 키를 부여해 AI가 정확히 하나만 고르게 합니다.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    headers = detect_sections(doc)
    raw = []
    name_counts = {}
    for p_idx, page in enumerate(doc):
        for w in page.widgets():
            if w.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
                continue
            r = w.rect
            clip = fitz.Rect(max(0, r.x0 - 120), max(0, r.y0 - 20),
                             min(page.rect.width, r.x1 + 60), min(page.rect.height, r.y1 + 20))
            nearby = clean_text(" ".join(page.get_text("text", clip=clip).split())[:160])
            if w.field_type in BUTTON_TYPES:
                # 버튼 바로 오른쪽/왼쪽 글자 = 해당 선택지 이름 (Yes/No 구분용)
                right = clean_text(" ".join(page.get_text("text", clip=fitz.Rect(r.x1, r.y0 - 2, r.x1 + 45, r.y1 + 2)).split()))
                left = clean_text(" ".join(page.get_text("text", clip=fitz.Rect(max(0, r.x0 - 30), r.y0 - 2, r.x0, r.y1 + 2)).split()))
                right_word = right.split()[0] if right.split() else ""
                left_word = left.split()[-1] if left.split() else ""
                nearby = f"{nearby} || THIS OPTION'S LABEL (text right after the box): {right_word or '?'} (text right before: {left_word or '?'})"
            info = {
                "name": w.field_name or "",
                "label": clean_text((w.field_label or "").strip()),
                "type": w.field_type,
                "type_str": w.field_type_string,
                "on_state": _safe_on_state(w) if w.field_type in BUTTON_TYPES else None,
                "choices": [clean_text(c if isinstance(c, str) else c[-1]) for c in (w.choice_values or [])] if w.field_type in CHOICE_TYPES else [],
                "nearby": nearby,
                "section": section_for(headers, p_idx, r.y0),
                "person": person_for(headers, p_idx, r.y0),
                "page": p_idx,
                "xref": w.xref,
                "rect": (r.x0, r.y0, r.x1, r.y1),
                "current": w.field_value,
            }
            raw.append(info)
            if w.field_type in BUTTON_TYPES:
                name_counts[info["name"]] = name_counts.get(info["name"], 0) + 1

    # 고유 키 부여
    fields, used = [], set()
    btn_index = {}
    for f in raw:
        # 칸 이름에 깨진 글자가 있어도 AI가 돌려준 키와 맞도록 정리된 이름을 키로 사용 (채우기는 xref로 함)
        key = clean_text(f["name"])
        if f["type"] in BUTTON_TYPES and name_counts.get(f["name"], 0) > 1:
            btn_index[f["name"]] = btn_index.get(f["name"], 0) + 1
            suffix = clean_text(f["on_state"] or f"#{btn_index[f['name']]}")
            key = f"{key} [{suffix}]"
        if f["type"] not in BUTTON_TYPES and key in used:
            # 같은 이름의 텍스트 필드는 값이 공유되므로 한 번만 노출
            f["key"] = key
            f["duplicate"] = True
            fields.append(f)
            continue
        base, n = key, 2
        while key in used:
            key = f"{base} ({n})"
            n += 1
        used.add(key)
        f["key"] = key
        f["duplicate"] = False
        fields.append(f)

    compute_printed_labels(doc, fields)
    mark_always_blank(doc, fields)

    xfa = False
    try:
        cat = doc.pdf_catalog()
        af = doc.xref_get_key(cat, "AcroForm")
        if af[0] == "xref":
            af_xref = int(af[1].split()[0])
            xfa = doc.xref_get_key(af_xref, "XFA")[0] != "null"
        elif af[0] == "dict":
            xfa = "/XFA" in af[1]
    except Exception:
        pass
    doc.close()
    return fields, xfa

# ------------------------------------------
# 1-1b. 표(경력·학력·주소·여행 등) 자동 인식
# ------------------------------------------
def detect_tables(pdf_bytes, fields):
    """
    위젯 위치로 표를 찾아 행/열 구조로 정리.
    - 같은 높이의 텍스트 칸 2개 이상 = 한 행, 열 위치가 같은 행이 3개 이상 이어지면 = 표 (페이지를 넘어가도 이어짐)
    - 행 안에 'Current'가 인쇄되어 있으면 그 행을 '현재 행'으로 표시
    → AI는 행 데이터만 돌려주고, 어느 칸에 넣을지는 코드가 정확히 계산합니다.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    cand = [f for f in fields if not f["duplicate"] and f["type"] == fitz.PDF_WIDGET_TYPE_TEXT]
    cand.sort(key=lambda f: (f["page"], f["rect"][1], f["rect"][0]))

    # 1) 행 묶기
    rows = []
    for f in cand:
        if rows and rows[-1]["page"] == f["page"] and abs(rows[-1]["y0"] - f["rect"][1]) <= 4:
            r = rows[-1]
            r["fields"].append(f)
            r["y1"] = max(r["y1"], f["rect"][3])
        else:
            rows.append({"page": f["page"], "y0": f["rect"][1], "y1": f["rect"][3], "fields": [f]})
    for r in rows:
        r["fields"].sort(key=lambda f: f["rect"][0])
        r["xs"] = [f["rect"][0] for f in r["fields"]]

    def compatible(a, b):
        matches = sum(1 for x in a["xs"] if any(abs(x - y) <= 8 for y in b["xs"]))
        return matches >= max(2, min(len(a["xs"]), len(b["xs"])) - 1)

    # 2) 연속 행 → 표
    groups, cur = [], []
    for i, r in enumerate(rows):
        ok = len(r["fields"]) >= 2
        if ok and cur:
            prev = cur[-1]
            same_page_next = prev["page"] == r["page"] and r["y0"] - prev["y1"] < 15
            next_page_top = (r["page"] == prev["page"] + 1 and rows[i - 1] is prev)
            if not ((same_page_next or next_page_top) and compatible(prev, r)):
                groups.append(cur)
                cur = []
        if ok:
            cur.append(r)
        else:
            if cur:
                groups.append(cur)
            cur = []
    if cur:
        groups.append(cur)

    tables = []
    for g in groups:
        if len(g) < 3:
            continue
        # [수정] '라벨 | 칸 | 라벨 | 칸' 형태의 개인정보 칸 묶음은 표가 아님
        # (예: City of Birth [ ] Country of Birth [ ]) → 칸마다 왼쪽에 자기 라벨이 있으면 제외
        cells = [f for r in g for f in r["fields"]]
        labeled = sum(1 for f in cells if f.get("label_side") == "left"
                      and not re.fullmatch(r"\s*current\s*", f.get("printed_label") or "", re.I))
        if labeled >= len(cells) * 0.5:
            continue
        # 열 위치 모으기
        col_xs = []
        for r in g:
            for x in r["xs"]:
                if not any(abs(x - c) <= 8 for c in col_xs):
                    col_xs.append(x)
        col_xs.sort()
        first = g[0]
        page = doc[first["page"]]
        top = first["y0"]
        # 열 제목 (표 바로 위 32pt)
        col_names = []
        for ci, cx in enumerate(col_xs):
            x_end = col_xs[ci + 1] if ci + 1 < len(col_xs) else page.rect.width
            words = [w for w in page.get_text("words", clip=fitz.Rect(cx - 2, max(0, top - 34), x_end - 1, top - 1))]
            name = " ".join(w[4] for w in sorted(words, key=lambda w: (round(w[1]), w[0])))
            col_names.append(clean_text(name.strip()) or f"Column {ci + 1}")
        title_txt = clean_text(" ".join(page.get_text("text", clip=fitz.Rect(0, max(0, top - 110), page.rect.width, top - 34)).split()))
        # 행 구성 + Current 행 찾기
        t_rows, current_idx = [], None
        for ri, r in enumerate(g):
            cells = {}
            for f in r["fields"]:
                ci = min(range(len(col_xs)), key=lambda c: abs(col_xs[c] - f["rect"][0]))
                cells[ci] = f["key"]
            pg = doc[r["page"]]
            band = fitz.Rect(0, r["y0"] - 3, pg.rect.width, r["y1"] + 3)
            words = [w[4].lower() for w in pg.get_text("words", clip=band)]
            if current_idx is None and any(w.startswith("current") for w in words):
                current_idx = ri
            t_rows.append(cells)
        to_col = next((i for i, n in enumerate(col_names) if re.match(r"\s*(to\b|까지)", n, re.I)), None)
        tables.append({
            "id": f"T{len(tables) + 1}",
            "page": first["page"],
            "person": first["fields"][0].get("person", "MAIN APPLICANT"),
            "section": first["fields"][0].get("section", ""),
            "title": title_txt[-160:],
            "columns": col_names,
            "rows": t_rows,
            "current_row": current_idx,
            "to_col": to_col,
        })
        for cells in t_rows:
            for k in cells.values():
                for f in fields:
                    if f["key"] == k:
                        f["table"] = tables[-1]["id"]
    doc.close()
    return tables

def format_table_catalog(tables):
    out = []
    for t in tables:
        cols = ", ".join(f'"c{i + 1}" = {n}' for i, n in enumerate(t["columns"]))
        cur = ("YES - the form has a pre-printed 'Current' row for the present job/residence"
               if t["current_row"] is not None else "no")
        out.append(f"TABLE {t['id']} | PERSON: {t['person']} | SECTION: {t['section'] or '-'} | PAGE: {t['page'] + 1} | "
                   f"ROWS: {len(t['rows'])} | CURRENT ROW: {cur}\n  TITLE/CONTEXT: {t['title']}\n  COLUMNS: {cols}")
    return "\n".join(out)

PRESENT_WORDS = {"present", "current", "now", "현재", "재직중", "재직 중", "ongoing", "to date", "-"}

def _is_present(v):
    return str(v or "").strip().lower() in PRESENT_WORDS

def place_table_records(tables, ai_tables):
    """
    AI가 준 표 데이터(최근→과거 순 레코드 목록)를 실제 칸에 배치.
    - 'Current' 행이 있는 표: Current 행은 항상 비워둠 (손님이 최신 정보로 직접 작성).
      이전 자료에서 '~ Current/Present'였던 레코드는 Current 바로 아래 행에 넣고 To 칸은 비움.
    - Current 행이 없는 표도 진행 중 레코드의 To 칸은 비움.
    - 중복 레코드 제거, 칸 수를 넘으면 넘친 개수 보고.
    반환: ({key: value}, {key: "표 설명"}, [경고])
    """
    values, labels, warns = {}, {}, []
    by_id = {t["id"]: t for t in tables}
    for tid, records in (ai_tables or {}).items():
        t = by_id.get(tid)
        if t is None or not isinstance(records, list):
            continue
        # 정리 + 중복 제거
        clean, seen = [], set()
        for rec in records:
            if not isinstance(rec, dict):
                continue
            vals = {}
            for ci in range(len(t["columns"])):
                v = rec.get(f"c{ci + 1}")
                if v is not None and str(v).strip():
                    vals[ci] = str(v).strip()
            if not vals:
                continue
            sig = tuple(sorted((k, v.lower()) for k, v in vals.items()))
            if sig in seen:
                continue
            seen.add(sig)
            is_cur = bool(rec.get("is_current")) or (t["to_col"] is not None and _is_present(vals.get(t["to_col"])))
            clean.append((is_cur, vals))

        n_rows = len(t["rows"])
        ci_row = t["current_row"]
        order = []  # (row_index, vals)
        # 진행 중 레코드는 To 칸 비우기 ('Present', 'Current' 등 제거)
        prepared = []
        for is_cur, vals in clean:
            if t["to_col"] is not None and (is_cur or _is_present(vals.get(t["to_col"]))):
                vals = {ci: v for ci, v in vals.items() if ci != t["to_col"]}
            prepared.append((is_cur, vals))
        # 진행 중이던 레코드가 맨 위 (최근 순)
        prepared.sort(key=lambda x: not x[0])
        start = (ci_row + 1) if ci_row is not None else 0
        order = [(start + i, v) for i, (_, v) in enumerate(prepared)]

        overflow = 0
        for ri, vals in order:
            if ri >= n_rows:
                overflow += 1
                continue
            cells = t["rows"][ri]
            for ci, v in vals.items():
                key = cells.get(ci)
                if key is None:
                    # Current 행의 'To' 칸처럼 칸 자체가 없는 경우 (인쇄된 Current) → 건너뜀
                    continue
                values[key] = v
                labels[key] = f"[{t['id']}] {ri + 1}행 · {t['columns'][ci][:20]}"
        if overflow:
            warns.append(f"표 {t['id']} ({t['title'][-30:]}): 칸이 부족해 {overflow}개 항목을 넣지 못했습니다. 별도 Word 파일로 정리해 주세요.")
    return values, labels, warns

def format_field_catalog(fields):
    lines = []
    for f in fields:
        if f["duplicate"] or f.get("table") or f.get("always_blank"):
            continue
        kind = {fitz.PDF_WIDGET_TYPE_CHECKBOX: "CHECKBOX",
                fitz.PDF_WIDGET_TYPE_RADIOBUTTON: "RADIO"}.get(f["type"], f["type_str"].upper())
        line = f"KEY: {json.dumps(f['key'], ensure_ascii=False)} | PERSON: {f.get('person', 'MAIN APPLICANT')} | SECTION: {f.get('section') or '-'} | TYPE: {kind} | PAGE: {f['page'] + 1}"
        if f.get("printed_label"):
            line += f" | PRINTED LABEL: {f['printed_label']}"
        if f["label"]:
            line += f" | TOOLTIP: {f['label'][:120]}"
        line += f" | NEARBY TEXT: {f['nearby']}"
        if f["choices"]:
            line += f" | ALLOWED VALUES: {f['choices']}"
        lines.append(line)
    return "\n".join(lines)

# ------------------------------------------
# 1-2. 손님 자료 → Gemini 입력 변환
# ------------------------------------------
def qa_prepare_document_for_gemini(file_bytes, mime_type, file_name="", max_chars=35000):
    ext = os.path.splitext(file_name)[1].lower() if file_name else ""
    mime_type = (mime_type or "").lower()
    header = f"\n--- [Document: {file_name}] ---\n"

    if ext == ".doc" or mime_type == "application/msword":
        # [수정] 구형 .doc는 python-docx로 못 읽고, Gemini도 받지 못해 전체 분석이 실패했음
        return [], f"{file_name}: 구형 .doc 형식은 읽을 수 없습니다. Word에서 .docx 또는 PDF로 저장 후 올려주세요."

    if "word" in mime_type or ext == ".docx":
        word_text = qa_read_word_document_text(file_bytes, file_name)
        if word_text:
            return [header + word_text[:max_chars] + "\n"], None
        return [], f"{file_name}: Word 파일에서 텍스트를 읽지 못했습니다."

    if ext in IMAGE_EXTS or mime_type.startswith("image/"):
        try:
            return [header + "(image below)", image_bytes_to_jpeg_part(file_bytes)], None
        except Exception as e:
            return [], f"{file_name}: 이미지를 열 수 없습니다 ({e})."

    if "pdf" in mime_type or ext == ".pdf":
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            for page in doc:
                text += page.get_text("text") + "\n"
                for w in page.widgets():
                    v = w.field_value
                    if v and str(v).lower() != "off":
                        label = w.field_label or w.field_name
                        text += f"[Filled field: {label}] {v}\n"

            # [수정] 텍스트가 조금만 있는 스캔본(예: 여권 스캔 + 워터마크 문구)도 이미지로 함께 보냄
            page_count = len(doc)
            avg_chars = len(text.strip()) / max(page_count, 1)
            parts = []
            if text.strip():
                parts.append(header + text[:max_chars] + "\n")
            if avg_chars < 200:
                if not parts:
                    parts.append(header + "(scanned pages below)")
                for p_idx in range(min(page_count, 10)):
                    pix = doc.load_page(p_idx).get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=80)
                    parts.append({"mime_type": "image/jpeg", "data": buf.getvalue()})
            doc.close()
            return parts, None
        except Exception as e:
            return [], f"{file_name}: PDF를 열 수 없습니다 ({e})."

    return [], f"{file_name}: 지원하지 않는 형식입니다."

def qa_batch_process_client_files(files, owner=None):
    """owner: 'MAIN APPLICANT' 또는 'FAMILY MEMBER' — 업로드 칸 기준으로 서류 주인을 표시."""
    def worker(f):
        label = f"{owner} | {f.name}" if owner else f.name
        return qa_prepare_document_for_gemini(f.getvalue(), f.type, label)
    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(executor.map(worker, files))
    flat, warnings = [], []
    for parts, warn in results:
        flat.extend(parts)
        if warn:
            warnings.append(warn)
    return flat, warnings

# ==========================================
# 2. AI 질문지 분석 및 매핑
# ==========================================
QA_RULES = """
[CRITICAL RULES - MUST OBEY]
1. UNKNOWN INFO = OMIT: If the client documents do NOT explicitly answer Tuberculosis, Medical issues, Crime, Visa Refusals, or Biometrics, OMIT those fields entirely. NEVER guess "No".
2. ENGLISH TESTS: Only if the client explicitly has an IELTS, CELPIP, or Duolingo score, select the English-test YES option and fill in the scores. Otherwise leave it blank.
3. CANADA VISA HISTORY: If the client currently holds a Canadian permit (Visitor Record, Work Permit, Study Permit), record it in the Canada visa application history field (e.g., "Work Permit (Valid until YYYY-MM-DD)").
4. CURRENT JOB LOGIC: Today is {today}. A job whose end date is before this month is NOT current (is_current = false). Only a job the person still holds today is current.
5. TABLE ORDER: All history tables are most recent first, with no duplicates.
6. OVERSEAS RESIDENCE: If overseas residence is inferred only from overseas work, fill City/Country but leave From/To dates blank.
7. FAMILY INFO: Fill spouse and children information if available.
8. SIGNATURE / DATE OF SIGNING: Leave blank.
9. LANGUAGE: All values MUST be in English (Latin letters, numbers). Romanize Korean names exactly as they appear in the passport. NO KOREAN characters in values.
10. DATES: Use YYYY-MM-DD unless the field's nearby text shows another format (e.g., DD/MM/YYYY, or separate Year/Month boxes).
11. CONFLICTS: If documents disagree, prefer passport > permit/official letters > previous questionnaire > resume.
"""

PEOPLE_RULES = """
[WHO IS WHO - VERY IMPORTANT]
- Each client document header says whose document it is:
  * "[Document: MAIN APPLICANT | ...]" = a document of the MAIN APPLICANT (주 신청자).
  * "[Document: FAMILY MEMBER | ...]" = a document of a family member (spouse, common-law partner, or child).
- For each FAMILY MEMBER document, decide who that person is relative to the main applicant, using the documents themselves
  (marriage certificate, family relation certificate, birth date/age, previous questionnaire, names) and the questionnaire:
  * married spouse or common-law partner → SPOUSE sections (choose 기혼/married vs 사실혼/common-law from the evidence)
  * child → CHILD sections, oldest child = CHILD 1
- A FAMILY MEMBER document may be that person's OWN previous questionnaire. In it, the '신청자/Applicant' sections describe
  THAT FAMILY MEMBER (not our main applicant). Use it as that person's own data, and use its spouse/family sections to work out
  the relationship (e.g. if it lists our main applicant as its spouse → this person is the SPOUSE; if it lists our main applicant
  as a parent → this person is a CHILD).
- A MAIN APPLICANT document may still mention family members (e.g. spouse name in a previous questionnaire) — use that for the
  main applicant's spouse/family questions, but the main applicant's OWN data always comes from MAIN APPLICANT documents.
- Every field has a PERSON tag (MAIN APPLICANT / SPOUSE / CHILD n) and a SECTION tag. Fill each field ONLY with the data of that PERSON:
  * Sections titled '신청자 ...' / 'Applicant ...' (or unprefixed sections) = MAIN APPLICANT's own data.
  * Sections titled '배우자 ...' / 'Spouse ...' = SPOUSE's own data (spouse's address, education, work, travel, etc.).
  * Sections titled '자녀(1)' / '자녀(2)' / 'Child' = that child's data.
- Inside the MAIN APPLICANT's personal section, the questions "배우자 이름 / Spouse name / 배우자 생년월일" are about the SPOUSE.
  Inside the SPOUSE's personal section, the same questions are about the MAIN APPLICANT.
- '신청자 가족관계 정보' (applicant family table) = the MAIN APPLICANT's parents, spouse, children, siblings.
  '배우자 가족관계 정보' (spouse family table) = the SPOUSE's parents, siblings, and also the MAIN APPLICANT as 'Spouse' and their children.
- NEVER copy one person's data (address history, schools, jobs, travel, birth place, height, eye colour) into another person's section.
- If there are NO family member documents, fill only the main applicant's sections (plus spouse/child names and dates if the main applicant's documents state them).
"""

def analyze_qa_for_fill(main_files, family_files, target_file):
    """반환: (result_dict | None, warnings, context)"""
    today = datetime.datetime.now(LOCAL_TZ).strftime("%B %d, %Y")
    people_rules = PEOPLE_RULES
    target_bytes = target_file.getvalue()
    is_docx = target_file.name.lower().endswith(".docx")

    source_parts, warnings = qa_batch_process_client_files(main_files or [], "MAIN APPLICANT")
    if family_files:
        fam_parts, fam_warn = qa_batch_process_client_files(family_files, "FAMILY MEMBER")
        source_parts += fam_parts
        warnings += fam_warn
    if not source_parts:
        st.error("읽을 수 있는 손님 자료가 없습니다.")
        return None, warnings, None

    if is_docx:
        form_text = qa_read_word_document_text(target_bytes, target_file.name)
        prompt = f"""
You are an expert Canadian immigration assistant filling out a blank questionnaire (Word document) using the client's existing documents.
{QA_RULES.format(today=today)}
{people_rules}
OUTPUT: Return ONLY a JSON object: {{"docx_fields": {{"<exact question/label text copied from the form>": "<answer>"}}}}
- Keys MUST be copied exactly (character for character) from the form text below, e.g. the Korean label "출생 도시" or "Date of Birth".
- Use the most specific label (e.g. "Father's Name", not just "Name").
- Only include fields you can answer from the client materials.

===== BLANK QUESTIONNAIRE (TARGET) =====
{form_text[:60000]}
===== END OF TARGET =====

===== CLIENT SOURCE MATERIALS =====
"""
        context = {"kind": "docx"}
    else:
        fields, has_xfa = extract_pdf_form_fields(target_bytes)
        if not fields:
            st.error("이 PDF에는 입력 가능한 칸(대화형 폼 필드)이 없습니다. 칸이 있는 PDF 또는 .docx 질문지를 올려주세요.")
            return None, warnings, None
        if has_xfa:
            warnings.append("이 PDF는 XFA(동적) 양식입니다. Adobe Acrobat에서는 채운 값이 보이지 않을 수 있으니 결과를 꼭 열어서 확인하세요.")
        tables = detect_tables(target_bytes, fields)
        page_text = ""
        try:
            d = fitz.open(stream=target_bytes, filetype="pdf")
            page_text = "\n".join(f"[Page {i + 1}]\n{p.get_text('text')}" for i, p in enumerate(d))
            d.close()
        except Exception:
            pass
        prompt = f"""
You are an expert Canadian immigration assistant filling out a blank fillable PDF questionnaire using the client's existing documents.
{QA_RULES.format(today=today)}
{people_rules}
HOW TO FILL:
- Each fillable field has a unique KEY listed below. Use KEYS EXACTLY as written (including any " [..]" suffix).
- TEXT / COMBOBOX: value is the answer text. For a COMBOBOX use one of its ALLOWED VALUES.
- CHECKBOX / RADIO: to select an option, set that option's KEY to "Yes". For a Yes/No question, select ONLY the one correct option key (e.g. "Q3 [Yes]": "Yes"). Never set "No" or "Off" — simply omit options that should not be selected.
- Use PAGE and NEARBY TEXT to understand which question each field belongs to.
- Decide what a box asks by its PRINTED LABEL first (the question printed right next to / above the box), then TOOLTIP, then NEARBY TEXT.
  KEY names are internal names that are often copied or wrong (e.g. a box named "Telephone" may actually be the bank-balance box) — never decide by the KEY name alone.
- Put a value only into a box whose label asks for exactly that kind of information. A phone number goes only into a phone/contact box, money only into a balance/amount box, etc.
- Current residential address and the most recent entry to Canada are intentionally NOT listed — the client fills them in. Do not try to fill them.
TABLES (work history, education, addresses, travel, relatives, etc.):
- Do NOT fill table cells by key. Instead return the table's records under "tables" -> table id, as a list of objects using the column ids ("c1", "c2", ...).
- List records MOST RECENT FIRST, each real record ONLY ONCE (never repeat the same job/school/address).
- Add "is_current": true for a record that was ongoing in the latest client document (e.g. "2020-04 ~ Present/Current" in a previous questionnaire). Leave its To column empty.
- Do NOT leave empty records to skip rows — the program places records into the right rows automatically
  (the pre-printed 'Current' row is always left blank for the client to update; all records start on the row below it).
- Fill each table with the data of that table's PERSON only.
OUTPUT: Return ONLY a JSON object:
{{"pdf_fields": {{"<KEY>": "<value>"}},
  "tables": {{"<TABLE ID>": [{{"is_current": true, "c1": "2025-09", "c2": "Present", "c3": "..."}}, {{"is_current": false, "c1": "...", "c2": "..."}}]}}}}

===== TABLES OF THE TARGET FORM =====
{format_table_catalog(tables)}

===== OTHER FILLABLE FIELDS OF THE TARGET FORM =====
{format_field_catalog(fields)}
===== FORM PAGE TEXT (for context) =====
{page_text[:40000]}
===== END OF TARGET =====

===== CLIENT SOURCE MATERIALS =====
"""
        context = {"kind": "pdf", "fields": fields, "tables": tables}

    try:
        response = safe_generate_content([prompt] + source_parts, json_mode=True, thinking=THINKING_LEVEL)
        data = parse_json_response(response.text)
        if not isinstance(data, dict):
            raise ValueError("AI 응답 형식이 올바르지 않습니다.")
        return data, warnings, context
    except Exception as e:
        st.error(f"서류 분석 중 오류가 발생했습니다: {e}")
        return None, warnings, None

PHONE_RE = re.compile(r"^\+?\(?\d{1,4}\)?[\s.-]?\d{2,4}[\s.-]\d{3,4}[\s.-]?\d{3,4}$")
MONEY_LABEL_RE = re.compile(r"balance|amount|잔고|잔액|금액|salary|income|wage|급여|연봉|\$|cad|usd|krw", re.I)
PHONE_LABEL_RE = re.compile(r"phone|telephone|tel\b|mobile|cell|연락처|전화|휴대", re.I)
EMAIL_LABEL_RE = re.compile(r"e-?mail|이메일", re.I)
DATE_LABEL_RE = re.compile(r"yyyy|yy-mm|date|날짜|생년월일|from|to\b", re.I)

def check_value_fits_label(f, v):
    """값이 칸 라벨과 성격이 맞는지 간단 점검. 안 맞으면 경고 문구 반환 (해당 칸은 기본 '기입 안 함')."""
    if f["type"] != fitz.PDF_WIDGET_TYPE_TEXT:
        return ""
    label = " ".join(x for x in (f.get("printed_label"), f.get("label")) if x)
    if not label:
        return ""
    is_phone_val = bool(PHONE_RE.match(v.strip())) or bool(re.match(r"^(\+?82[\s-]?)?0?1[016789][\s-]?\d{3,4}[\s-]?\d{4}$", v.strip()))
    if MONEY_LABEL_RE.search(label) and not PHONE_LABEL_RE.search(label) and is_phone_val:
        return "⚠️ 금액 칸에 전화번호가 들어감 (기입 해제됨)"
    if EMAIL_LABEL_RE.search(label) and not PHONE_LABEL_RE.search(label) and "@" not in v:
        return "⚠️ 이메일 칸인데 이메일 형식이 아님 (기입 해제됨)"
    if "@" in v and not EMAIL_LABEL_RE.search(label):
        return "⚠️ 이메일 칸이 아닌데 이메일이 들어감 (기입 해제됨)"
    if PHONE_LABEL_RE.search(label) and not EMAIL_LABEL_RE.search(label) and not re.search(r"\d{3}", v) and not re.search(r"^(n/?a|none|없음)$", v, re.I):
        return "⚠️ 전화번호 칸인데 숫자가 없음 (기입 해제됨)"
    return ""

def build_review_rows(data, context):
    """AI 결과를 검토용 표로 변환 (없는 키·한글 값 표시)."""
    rows = []
    if context["kind"] == "pdf":
        by_key = {f["key"]: f for f in context["fields"] if not f["duplicate"]}
        t_values, t_labels, t_warns = place_table_records(context.get("tables", []), data.get("tables"))
        context["table_warnings"] = t_warns
        plain = {k: v for k, v in (data.get("pdf_fields") or {}).items()
                 if not (by_key.get(k) or {}).get("table") and not (by_key.get(k) or {}).get("always_blank")}  # 표 칸은 표 로직 결과만 사용
        merged = list(t_values.items()) + list(plain.items())
        for k, v in merged:
            if v is None or str(v).strip() == "":
                continue
            f = by_key.get(k)
            if isinstance(v, bool):
                v = "Yes" if v else ""
                if not v:
                    continue
            v = str(v).strip()
            if f is not None and f["type"] in BUTTON_TYPES and v.lower() not in TRUE_WORDS:
                continue  # "No"/"Off" 지정은 체크하지 않으므로 표에서도 제외
            note = ""
            mismatch = check_value_fits_label(f, v) if f is not None else ""
            if f is None:
                note = "⚠️ 서식에 없는 칸 (기입 안 됨)"
            elif mismatch:
                note = mismatch
            elif HANGUL_RE.search(v):
                note = "⚠️ 한글 포함"
            elif f["choices"] and v.lower() not in [c.lower() for c in f["choices"]]:
                note = "⚠️ 선택지에 없는 값"
            rows.append({"기입": f is not None and not mismatch, "구역": ({"MAIN APPLICANT": "신청자", "SPOUSE": "배우자"}.get(f.get("person"), (f.get("person") or "").replace("CHILD", "자녀")) if f else ""),
                         "칸 (Key)": k,
                         "주변 문구": (t_labels.get(k) or (f.get("printed_label") if f else "") or (f["nearby"] if f else ""))[:80],
                         "값": v, "확인 필요": note})
    else:
        for k, v in (data.get("docx_fields") or {}).items():
            if v is None or str(v).strip() == "":
                continue
            v = str(v).strip()
            rows.append({"기입": True, "구역": "", "칸 (Key)": k, "주변 문구": "", "값": v,
                         "확인 필요": "⚠️ 한글 포함" if HANGUL_RE.search(v) else ""})
    return rows

# ==========================================
# 3. PDF 대화형 폼(Widget) 채우기 엔진
# ==========================================
TRUE_WORDS = {'yes', 'true', 'v', 'x', 'on', '1', 'checked', 'y'}

def _fit_fontsize(text, rect_width, base=10.0, min_size=6.0):
    size = base
    while size > min_size and fitz.get_text_length(text, fontname="cour", fontsize=size) > rect_width - 4:
        size -= 0.5
    return size

def _fit_multiline(text, width, height, max_size=9.0, min_size=5.0):
    """줄바꿈했을 때 칸 높이에 들어가는 가장 큰 글자 크기 (없으면 None)."""
    size = max_size
    words = text.split()
    while size >= min_size:
        lines, cur = 1, ""
        for wd in words:
            trial = (cur + " " + wd).strip()
            if fitz.get_text_length(trial, fontname="cour", fontsize=size) <= width - 4:
                cur = trial
            else:
                lines += 1
                cur = wd
        if lines * size * 1.3 + 3 <= height:
            return size
        size -= 0.5
    return None

def ensure_acroform_font(doc, alias="Cour", base_font="Courier"):
    """
    [수정] 질문지 PDF의 폼 글꼴 목록(DR)에 Courier가 없으면, Adobe가 칸을 다시 그릴 때 글자 크기를 무시하고
    크게 표시해서 'COMMUNITY MANAGE…'처럼 칸을 넘쳤음 → Courier를 폼 글꼴 목록에 등록.
    """
    try:
        cat = doc.pdf_catalog()
        af = doc.xref_get_key(cat, "AcroForm")
        if af[0] != "xref":
            return
        af_xref = int(af[1].split()[0])
        if doc.xref_get_key(af_xref, f"DR/Font/{alias}")[0] != "null":
            return
        fx = doc.get_new_xref()
        doc.update_object(fx, f"<< /Type /Font /Subtype /Type1 /BaseFont /{base_font} /Encoding /WinAnsiEncoding >>")
        if doc.xref_get_key(af_xref, "DR")[0] == "null":
            doc.xref_set_key(af_xref, "DR", "<< /Font << >> >>")
        elif doc.xref_get_key(af_xref, "DR/Font")[0] == "null":
            doc.xref_set_key(af_xref, "DR/Font", "<< >>")
        doc.xref_set_key(af_xref, f"DR/Font/{alias}", f"{fx} 0 R")
    except Exception:
        pass

def fill_qa_pdf_interactive(target_bytes, values_by_key, fields):
    """values_by_key: {key: value}. 필드 키 → xref로 정확히 해당 위젯만 채움."""
    doc = fitz.open(stream=target_bytes, filetype="pdf")
    ensure_acroform_font(doc)
    # 같은 이름 텍스트 필드(중복)도 함께 채우기 위해 이름 → 대표 키 매핑
    key_by_name = {f["name"]: f["key"] for f in fields if not f["duplicate"] and f["type"] not in BUTTON_TYPES}
    filled, skipped = 0, []

    for f in fields:
        key = key_by_name.get(f["name"], f["key"]) if f["duplicate"] else f["key"]
        if key not in values_by_key or f.get("always_blank"):
            continue
        val = str(values_by_key[key]).strip()
        if not val:
            continue
        page = doc[f["page"]]
        w = page.load_widget(f["xref"])
        if w is None:
            skipped.append(key)
            continue
        try:
            if f["type"] in BUTTON_TYPES:
                if val.lower() not in TRUE_WORDS:
                    continue  # "No"/"Off"는 아무것도 체크하지 않음
                w.field_value = f["on_state"] or True
                w.update()
            elif f["type"] in CHOICE_TYPES:
                match = next((c for c in f["choices"] if c.lower() == val.lower()), None)
                if match is None and not (w.field_flags & (1 << 18)):  # 편집 불가 콤보박스
                    skipped.append(key)
                    continue
                w.field_value = match or val
                w.update()
            else:
                w.field_flags &= ~1  # ReadOnly 해제
                if w.text_maxlen and len(val) > w.text_maxlen:
                    val = val[:w.text_maxlen]
                w.field_value = val
                w.text_font = "Cour"
                multiline = bool(w.field_flags & (1 << 12))
                if not multiline and not (w.field_flags & (1 << 24)):  # 콤(Comb) 칸 제외
                    base = w.text_fontsize if w.text_fontsize and w.text_fontsize > 0 else 10
                    size = _fit_fontsize(val, w.rect.width, base=min(base, 11))
                    too_long = fitz.get_text_length(val, fontname="cour", fontsize=size) > w.rect.width - 4
                    if too_long or size < 7.5:
                        # 한 줄로는 너무 작거나 안 들어가면 2~3줄로 줄바꿈 (예: 'Intermediate Environmental Designer')
                        ml = _fit_multiline(val, w.rect.width, w.rect.height)
                        if ml and (too_long or ml > size):
                            w.field_flags |= (1 << 12)
                            size = ml
                    w.text_fontsize = size
                w.update()
            filled += 1
        except Exception:
            skipped.append(key)

    doc.need_appearances(False)  # 우리가 만든 모양(글자 크기·줄바꿈) 그대로 표시
    out_buf = io.BytesIO()
    doc.save(out_buf, garbage=1, deflate=True)
    doc.close()
    return out_buf.getvalue(), filled, skipped

# ==========================================
# 4. Word(.docx) 질문지 채우기 엔진
# ==========================================
def _norm_label(s):
    return re.sub(r"[\s:：_\-\.\*]+", " ", (s or "")).strip().lower()

def fill_qa_docx(target_bytes, values_by_key):
    doc = docx.Document(io.BytesIO(target_bytes))
    BLUE_RGB = RGBColor(0x1F, 0x4E, 0x8C)
    # lxml 요소 자체를 set에 보관 (id()는 재사용될 수 있어 오작동)
    # [수정] 긴 라벨부터 매칭 → "Name"이 "Father's Name" 칸을 가로채지 않도록
    items = sorted(
        [(_norm_label(k), str(v).strip()) for k, v in values_by_key.items() if str(v).strip() and _norm_label(k)],
        key=lambda x: -len(x[0]),
    )
    used_labels, used_cells, used_paras = set(), set(), set()

    def insert_run(p, text):
        run = p.add_run(f" {text}")
        run.font.color.rgb = BLUE_RGB
        run.bold = True

    def iter_tables(container):
        for t in container.tables:
            yield t
            for row in t.rows:
                for cell in row.cells:
                    yield from iter_tables(cell)  # 중첩 표 지원

    for table in iter_tables(doc):
        for row in table.rows:
            # 병합 셀 중복 제거
            cells, seen = [], set()
            for c in row.cells:
                if c._tc not in seen:
                    seen.add(c._tc)
                    cells.append(c)
            for c_idx, cell in enumerate(cells):
                if cell._tc in used_cells:
                    continue
                c_norm = _norm_label(cell.text)
                if not c_norm:
                    continue
                for label, ans in items:
                    if label in used_labels:
                        continue
                    if c_norm == label or (label in c_norm and len(c_norm) <= len(label) + 25):
                        nxt = cells[c_idx + 1] if c_idx + 1 < len(cells) else None
                        if nxt is not None and not nxt.text.strip() and nxt._tc not in used_cells:
                            p = nxt.paragraphs[0] if nxt.paragraphs else nxt.add_paragraph()
                            insert_run(p, ans)
                            used_cells.add(nxt._tc)
                        elif "___" in cell.text or cell.text.strip().endswith((":", "：")):
                            p = cell.paragraphs[-1]
                            insert_run(p, ans)
                        else:
                            continue
                        used_cells.add(cell._tc)
                        used_labels.add(label)
                        break

    for p in doc.paragraphs:
        if p._p in used_paras:
            continue
        raw = p.text
        p_norm = _norm_label(raw)
        if not p_norm:
            continue
        for label, ans in items:
            if label in used_labels:
                continue
            if label in p_norm and ("___" in raw or raw.strip().endswith((":", "："))):
                insert_run(p, ans)
                used_paras.add(p._p)
                used_labels.add(label)
                break

    not_placed = [k for k, v in values_by_key.items()
                  if str(v).strip() and _norm_label(k) and _norm_label(k) not in used_labels]
    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue(), len(used_labels), not_placed

def build_output_file(rows, ctx, target_bytes, base_name):
    values = {r["칸 (Key)"]: r["값"] for r in rows if r.get("기입") and str(r.get("값", "")).strip()}
    if ctx["kind"] == "docx":
        out, filled, skipped = fill_qa_docx(target_bytes, values)
        return out, f"{base_name}_Info_Filled.docx", \
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filled, skipped
    out, filled, skipped = fill_qa_pdf_interactive(target_bytes, values, ctx["fields"])
    return out, f"{base_name}_Info_Filled.pdf", "application/pdf", filled, skipped

def files_signature(files):
    h = hashlib.md5()
    for f in files:
        if f is not None:
            h.update(f.name.encode())
            h.update(str(f.size).encode())
    return h.hexdigest()


# ==========================================
# 6. Streamlit 네비게이션 및 UI 구성
# ==========================================
MENU_1 = "🍁 IMM5476 자동 작성"
MENU_2 = "✈️ 한부모 동의서 자동 작성"
MENU_3 = "📋 IMM서류 정보 정리"
MENU_4 = "🏷️ CRM 파일명 생성 및 묶기/분할"
MENU_5 = "📝 질문지 기존 정보 기입"

st.sidebar.title("🦅 CanNest Tool")
app_mode = st.sidebar.radio(
    "원하시는 업무 도구를 선택하세요", [MENU_1, MENU_2, MENU_3, MENU_4, MENU_5]
)

# ------------------------------------------
# 메뉴 1: IMM5476 자동 작성
# ------------------------------------------
if app_mode == MENU_1:
  st.title(MENU_1)
  if "extracted_5476" not in st.session_state:
    st.session_state.extracted_5476 = None

  template_5476_bytes = get_preloaded_file_bytes(
      ["imm5476_template.pdf", "imm5476_template.pdf.pdf"]
  )

  if template_5476_bytes:
    st.success("✅ 사내 표준 'IMM5476' 양식이 자동으로 로드되었습니다.")
  else:
    st.error(
        "⚠️ GitHub에 'imm5476_template.pdf' 파일이 없습니다. 수동으로"
        " 업로드해 주세요."
    )
    template_file = st.file_uploader(
        "IMM5476 템플릿 PDF 선택", type=["pdf"], key="template_5476"
    )
    if template_file:
      template_5476_bytes = template_file.getvalue()

  st.markdown("---")
  client_file = st.file_uploader(
      "1. 손님 여권 또는 퍼밋",
      type=[
          "jpg",
          "jpeg",
          "png",
          "pdf",
          "heic",
          "HEIC",
          "docx",
          "DOCX",
          "doc",
          "DOC",
      ],
      key="client_5476",
  )

  if client_file and st.button("정보 추출하기", use_container_width=True):
    with st.spinner("서류 분석 중입니다. 잠시만 기다려 주세요..."):
      extracted = extract_imm5476_info(
          process_uploaded_file_to_image(client_file)
      )
      if extracted:
        st.session_state.extracted_5476 = extracted
        st.success("정보 추출이 완료되었습니다.")

  if st.session_state.extracted_5476:
    data = st.session_state.extracted_5476
    c1, c2 = st.columns(2)
    with c1:
      surname = st.text_input("성 (Surname)", data.get("surname", ""))
      dob = st.text_input("생년월일", data.get("dob", ""))
      # 💡 UI 렌더링 시 AI 추출 이메일 값 연동
      email = st.text_input("이메일 주소", data.get("email", ""))
    with c2:
      given = st.text_input("이름 (Given Name)", data.get("given_name", ""))
      uci = st.text_input("UCI", data.get("uci", ""))
      sign_date = st.date_input("서명날짜", datetime.date.today())

    # 💡 UI 렌더링 시 AI 추출 주소/번호 값 연동
    address_phone = st.text_input(
        "주소 또는 전화번호 (이메일이 없는 미성년자/신청자용)",
        value=data.get("address_phone", ""),
        placeholder="예: 2301-6658 Dow Ave, Burnaby BC V5H 0C7",
    )

    if st.button("문서 생성 및 다운로드", type="primary"):
      if not template_5476_bytes:
        st.error("템플릿 파일이 없습니다.")
      else:
        final_data = {
            "surname": surname,
            "given_name": given,
            "dob": dob,
            "uci": uci,
            "email": email,
            "address_phone": address_phone,
            "signDate": sign_date.strftime("%Y-%m-%d"),
        }
        pdf_out = fill_imm5476(template_5476_bytes, final_data)
        st.download_button(
            "📥 다운로드",
            pdf_out,
            file_name=f"IMM5476_{surname}_{given}.pdf",
            mime="application/pdf",
        )

    st.markdown("---")
    if st.button(
        "🔄 전체 리셋",
        type="secondary",
        use_container_width=True,
        key="reset_btn_m1",
    ):
      for key in list(st.session_state.keys()):
        if key != "password_correct":
          del st.session_state[key]
      st.rerun()

# ------------------------------------------
# 메뉴 2: 한부모 동의서 자동 작성
# ------------------------------------------
elif app_mode == MENU_2:
  st.title(MENU_2)
  if "consent_non_acc" not in st.session_state:
    st.session_state.consent_non_acc = {}
  if "consent_family" not in st.session_state:
    st.session_state.consent_family = []

  consent_template_bytes = get_preloaded_file_bytes(
      ["consent_template.pdf", "consent_template.pdf.pdf"]
  )

  if consent_template_bytes:
    st.success("✅ 사내 표준 '한부모 동의서' 양식이 자동으로 로드되었습니다.")
  else:
    st.error(
        "⚠️ GitHub에 'consent_template.pdf' 파일이 없습니다. 수동으로"
        " 업로드해 주세요."
    )
    consent_template = st.file_uploader("동의서 양식 선택", type=["pdf"])
    if consent_template:
      consent_template_bytes = consent_template.getvalue()

  st.markdown("---")
  c1, c2 = st.columns(2)
  with c1:
    non_acc_file = st.file_uploader(
        "비동반 부모님 여권 (1장)",
        type=[
            "jpg",
            "jpeg",
            "png",
            "pdf",
            "heic",
            "HEIC",
            "docx",
            "DOCX",
            "doc",
            "DOC",
        ],
    )
  with c2:
    family_files = st.file_uploader(
        "동반 부모/자녀 여권",
        type=[
            "jpg",
            "jpeg",
            "png",
            "pdf",
            "heic",
            "HEIC",
            "docx",
            "DOCX",
            "doc",
            "DOC",
        ],
        accept_multiple_files=True,
    )

  if st.button(
      "여권 정보 추출하기", type="primary", use_container_width=True
  ):
    images = []
    has_non_acc = bool(non_acc_file)
    if non_acc_file:
      images.append(process_uploaded_file_to_image(non_acc_file))
    if family_files:
      images.extend(
          [process_uploaded_file_to_image(f) for f in family_files]
      )

    if images:
      with st.spinner(
          "여권 정보를 분석 중입니다. 잠시만 기다려 주세요..."
      ):
        res = extract_all_passports_batch(has_non_acc, images)
        if res:
          st.session_state.consent_non_acc = (
              res.get("non_accompanying_parent", {}) or {}
          )
          st.session_state.consent_family = (
              res.get("family_members", []) or []
          )
          st.success("여권 정보 추출이 완료되었습니다.")

  non_acc_data = st.session_state.consent_non_acc
  non_acc_name = st.text_input(
      "비동반 부모 성명",
      format_full_name(
          non_acc_data.get("surname", ""),
          non_acc_data.get("given_name", ""),
      ),
  )
  non_acc_address = st.text_input("주소")
  ca, cb = st.columns(2)
  with ca:
    non_acc_phone = st.text_input("전화번호")
  with cb:
    non_acc_email = st.text_input("이메일")

  acc_parents = [
      p
      for p in st.session_state.consent_family
      if not is_minor(p.get("dob", ""))
  ]
  children_list = [
      p for p in st.session_state.consent_family if is_minor(p.get("dob", ""))
  ]

  acc_name, acc_passport, acc_rel = "", "", "Mother"
  if acc_parents:
    p = acc_parents[0]
    acc_name = format_full_name(p.get("surname", ""), p.get("given_name", ""))
    acc_passport = p.get("passport_number", "")
    acc_rel = "Mother" if p.get("gender") == "F" else "Father"

  cp1, cp2, cp3 = st.columns(3)
  with cp1:
    acc_name = st.text_input("동반 부모 성명", acc_name)
  with cp2:
    acc_rel = st.selectbox(
        "관계",
        ["Mother", "Father"],
        index=0 if acc_rel == "Mother" else 1,
    )
  with cp3:
    acc_passport = st.text_input("여권번호", acc_passport)

  final_children = []
  for idx, c in enumerate(children_list):
    cc1, cc2 = st.columns(2)
    with cc1:
      n = st.text_input(
          f"자녀{idx+1} 성명",
          format_full_name(c.get("surname", ""), c.get("given_name", "")),
      )
    with cc2:
      d = st.text_input(f"자녀{idx+1} 생일", c.get("dob", "").replace("-", "/"))
    final_children.append({"name": n, "dob": d})

  if not children_list:
    cc1, cc2 = st.columns(2)
    with cc1:
      n = st.text_input("자녀 성명")
    with cc2:
      d = st.text_input("자녀 생일")
    if n:
      final_children.append({"name": n, "dob": d})

  trip_address = st.text_input("현지 주소")
  ct1, ct2 = st.columns(2)
  with ct1:
    trip_phone = st.text_input("현지 전화")
  with ct2:
    trip_email = st.text_input("현지 이메일")

  trip_date = st.text_input(
      "여행 기간 (Travel Date)",
      placeholder="예: 2026/09/01 ~ 2026/09/30 또는 August 2026",
  )

  sign_date_str = st.date_input("서명일", datetime.date.today()).strftime(
      "%Y/%m/%d"
  )

  if st.button("문서 생성 및 다운로드", type="primary"):
    if not consent_template_bytes:
      st.error("양식 파일이 없습니다.")
    else:
      data_consent = {
          "non_acc_name": non_acc_name,
          "non_acc_address": non_acc_address,
          "non_acc_phone": non_acc_phone,
          "non_acc_email": non_acc_email,
          "children": final_children,
          "acc_name": acc_name,
          "acc_relationship": acc_rel,
          "acc_passport": acc_passport,
          "trip_address": trip_address,
          "trip_phone": trip_phone,
          "trip_email": trip_email,
          "trip_date": trip_date,
          "sign_date": sign_date_str,
      }
      pdf_out = fill_consent_letter(consent_template_bytes, data_consent)

      crm_name = "NAME"
      if non_acc_name:
        if re.search(r"[가-힣]", non_acc_name):
          crm_name = non_acc_name.replace(" ", "")
        else:
          parts = non_acc_name.strip().split()
          if parts:
            crm_name = parts[0].capitalize()

      download_file_name = (
          f"{crm_name}_Consent Letter for Children Travelling Abroad.pdf"
      )

      st.download_button(
          "📥 다운로드", pdf_out, download_file_name, "application/pdf"
      )

  st.markdown("---")
  if st.button(
      "🔄 전체 리셋",
      type="secondary",
      use_container_width=True,
      key="reset_btn_m2",
  ):
    for key in list(st.session_state.keys()):
      if key != "password_correct":
        del st.session_state[key]
    st.rerun()

# ------------------------------------------
# 메뉴 3: 이민서류 정보 정리 (Case File Prep)
# ------------------------------------------
elif app_mode == MENU_3:
  st.title(MENU_3)

  if "prep_result" not in st.session_state:
    st.session_state.prep_result = None

  st.subheader("1. 대상 서식 (IMM PDF)")

  form_map = {
      "직접 파일 업로드 (기타 서식)": None,
      "IMM1294 (SP-OUTSIDE)": "imm1294.pdf",
      "IMM1295 (WP-OUTSIDE)": "imm1295.pdf",
      "IMM5708 (VR-INSIDE)": "imm5708.pdf",
      "IMM5709 (SP-INSIDE)": "imm5709.pdf",
      "IMM5710 (WP-INSIDE)": "imm5710.pdf",
  }

  selected_form = st.selectbox("📌 템플릿 서식 선택", list(form_map.keys()))

  tmpl_bytes = None

  if form_map[selected_form] is not None:
    file_path = form_map[selected_form]
    tmpl_bytes = load_pdf_bytes_cached(file_path)
    if tmpl_bytes:
      st.success(f"✅ '{selected_form}' 양식이 자동으로 로드되었습니다.")
    else:
      st.error(
          f"⚠️ {file_path} 파일이 서버에 없습니다. 파일 업로드 상태를"
          " 확인해 주세요."
      )
  else:
    tmpl_prep_file = st.file_uploader(
        "빈 IMM 서식 (반드시 Print to PDF로 평탄화된 파일)",
        type=["pdf"],
        key="case_tmpl",
    )
    if tmpl_prep_file:
      tmpl_bytes = tmpl_prep_file.getvalue()

  st.markdown("---")
  st.subheader("2. 손님 제출 서류 (복수 선택 가능)")
  client_prep_files = st.file_uploader(
      "질문지, 여권, 퍼밋 등 서류 선택",
      type=[
          "jpg",
          "jpeg",
          "png",
          "pdf",
          "heic",
          "HEIC",
          "docx",
          "DOCX",
          "doc",
          "DOC",
      ],
      accept_multiple_files=True,
      key="case_client_docs",
  )

  if st.button(
      "서류 정보 정리하기", type="primary", use_container_width=True
  ):
    if tmpl_bytes is None:
      st.warning(
          "1번 단계에서 서식이 정상적으로 선택되거나 업로드되지 않았습니다."
      )
    elif not client_prep_files:
      st.warning("2번 단계에서 손님 서류를 1개 이상 올려주세요.")
    else:
      with st.spinner(
          "서류를 대조하여 정보 및 불일치 항목을 확인 중입니다. 잠시만"
          " 기다려 주세요..."
      ):
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
      st.error(
          "서식에서 분석할 항목을 찾지 못했습니다. (파일이 평탄화된 PDF인지"
          " 확인해 주세요)"
      )
    else:
      full_text_list = []
      for sec in sections:
        sec_name = sec.get("section", "기타 항목")
        st.write(f"### 📌 {sec_name}")
        full_text_list.append(f"[{sec_name}]")

        table_data = []
        prev_group = None
        for f in sec.get("fields", []):
          field_lbl = f.get("field", "")
          val = f.get("value", "")
          src = f.get("source", "")

          group_match = re.match(
              r"^(.*?\bEntry\s*\d+)", field_lbl, re.IGNORECASE
          )
          curr_group = (
              group_match.group(1).strip() if group_match else None
          )

          if prev_group and curr_group and prev_group != curr_group:
            table_data.append({
                "항목 (Field)": "──────────",
                "추출값 (Value)": "──────────",
                "출처 (Source)": "──────────",
            })
            full_text_list.append("")

          prev_group = curr_group

          if not val:
            display_val = "⚠️ 확인 필요 (미발견)"
            full_text_list.append(f"{field_lbl}: (확인 필요)")
          else:
            display_val = val
            full_text_list.append(f"{field_lbl}: {val}")

          table_data.append({
              "항목 (Field)": field_lbl,
              "추출값 (Value)": display_val,
              "출처 (Source)": src if src else "-",
          })

        st.table(table_data)
        full_text_list.append("")

      st.markdown("#### 📋 한눈에 복사하기")
      st.text_area(
          "아래 텍스트를 복사하여 서식에 옮겨 적으세요",
          value="\n".join(full_text_list),
          height=250,
      )

      if st.button("＋ 새 케이스 정리하기"):
        st.session_state.prep_result = None
        st.rerun()

# ------------------------------------------
# 메뉴 4: CRM 파일명 자동 생성 및 묶기/분할
# ------------------------------------------
elif app_mode == MENU_4:
  st.title(MENU_4)
  st.caption(
      "개별 낱장 이미지, 여러 장짜리 통짜 PDF, MS Word 서류 등을 섞어서 올려도"
      " AI가 알아서 문서 단위로 묶거나 분할하여 CRM 파일명으로 최적화합니다."
  )

  if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = str(uuid.uuid4())
  if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None

  uploaded_files = st.file_uploader(
      "서류 업로드 (복수 선택 가능)",
      type=[
          "jpg",
          "jpeg",
          "png",
          "pdf",
          "heic",
          "HEIC",
          "docx",
          "DOCX",
          "doc",
          "DOC",
      ],
      accept_multiple_files=True,
      key=st.session_state.uploader_key,
  )

  if uploaded_files:
    if st.button(
        "서류 분석 및 묶기/분할 시작",
        type="primary",
        use_container_width=True,
    ):
      results = []
      status_text = st.empty()
      progress_bar = st.progress(0)

      status_text.text("1. 전체 서류 페이지 스캔 및 미리보기 생성 중...")

      global_pages = []
      page_counter = 1

      # [수정] 장수에 맞춰 미리보기 크기 결정 (40장 이하는 기존과 동일)
      total_pages = count_total_pages(uploaded_files)
      prev_zoom, prev_quality, prev_max_dim = crm_preview_settings(min(total_pages, CRM_MAX_PAGES))

      for file in uploaded_files:
        file_bytes = file.getvalue()
        mime_type = file.type if file.type else "application/pdf"
        fname_lower = file.name.lower()

        if fname_lower.endswith((".doc", ".docx")):
          if page_counter > CRM_MAX_PAGES:
            continue
          text_content = read_word_document_text(file_bytes, file.name)
          img = Image.new("RGB", (800, 1000), color=(255, 255, 255))
          draw = ImageDraw.Draw(img)
          disp_text = f"[Word File: {file.name}]\n\n" + (
              text_content[:800] if text_content else "Word Document"
          )
          draw.text((40, 40), disp_text, fill=(0, 0, 0))

          buf = io.BytesIO()
          img.save(buf, format="JPEG", quality=60)

          global_pages.append({
              "global_idx": page_counter,
              "original_name": file.name,
              "mime_type": mime_type,
              "file_bytes": file_bytes,
              "pdf_page_idx": 0,
              "preview_bytes": buf.getvalue(),
              "is_word": True,
              "word_text": text_content,
          })
          page_counter += 1
        elif "pdf" in mime_type.lower() or fname_lower.endswith(".pdf"):
          doc = fitz.open(stream=file_bytes, filetype="pdf")

          for i in range(len(doc)):
            if page_counter > CRM_MAX_PAGES:
              break

            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(prev_zoom, prev_zoom))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=prev_quality)

            global_pages.append({
                "global_idx": page_counter,
                "original_name": file.name,
                "mime_type": mime_type,
                "file_bytes": file_bytes,
                "pdf_page_idx": i,
                "preview_bytes": buf.getvalue(),
                "is_word": False,
            })
            page_counter += 1
          doc.close()
        else:
          if page_counter > CRM_MAX_PAGES:
            continue
          img = Image.open(io.BytesIO(file_bytes))
          img = ImageOps.exif_transpose(img)
          if img.mode != "RGB":
            img = img.convert("RGB")

          preview = img.copy()
          max_dim = max(preview.width, preview.height)
          if max_dim > prev_max_dim:
            ratio = float(prev_max_dim) / float(max_dim)
            preview = preview.resize(
                (int(preview.width * ratio), int(preview.height * ratio)),
                Image.Resampling.LANCZOS,
            )
          buf = io.BytesIO()
          preview.save(buf, format="JPEG", quality=prev_quality)

          global_pages.append({
              "global_idx": page_counter,
              "original_name": file.name,
              "mime_type": mime_type,
              "file_bytes": file_bytes,
              "pdf_page_idx": 0,
              "preview_bytes": buf.getvalue(),
              "is_word": False,
          })
          page_counter += 1

      if total_pages > CRM_MAX_PAGES:
        st.warning(
            f"⚠️ 업로드된 총 페이지 수({total_pages}장)가 {CRM_MAX_PAGES}장을 초과하여,"
            f" 앞의 {CRM_MAX_PAGES}장까지만 분석 및 병합합니다. 나머지는 따로 올려주세요."
        )
      shrink_previews_to_fit(global_pages)

      status_text.text(
          "2. AI가 페이지별 문맥을 분석하여 연관 서류를 묶거나 나누는"
          " 중입니다..."
      )

      prompt = f"""
            You are an expert AI document classifier for a Canadian immigration firm.
            I am providing {len(global_pages)} pages of documents uploaded by a client. 

            Your task:
            1. Read ALL pages carefully.
            2. GROUP the pages that logically belong to the SAME document type for the SAME client. 
               *CRITICAL MERGE RULE 1 (Passport)*: If you see multiple passports for the same person, SEPARATE them into Old Passport and New Passport by checking the expiry dates. Any stamp or visa pages MUST be merged with the correct passport based on the stamp/visa dates matching the valid period of the passport. DO NOT classify a TRV sticker inside a passport as a Visitor Record.
               *CRITICAL MERGE RULE 2 (ID/License)*: For ID cards, Driver's Licences, and PR Cards, the page containing the primary bio-data (face photo, name, DOB) MUST be ordered as Page 1 (Front), and the backside as Page 2.
               *CRITICAL MERGE RULE 3 (Digital Photo)*: If you see a studio receipt/timestamp page along with a face photo, merge them into ONE single Digital Photo document.
               *CRITICAL MERGE RULE 4*: Merge ALL BANK STATEMENTS, PAYSTUBS, or UTILITY BILLS for the SAME client into a single group.
            3. ROTATION CORRECTION (CRITICAL - TEXT UPRIGHT CHECK): 
               - Look strictly at the MAIN English text/numbers (e.g., 'Canada', 'DRIVER'S LICENCE', 'KIM MIN') and face photo.
               - SPECIAL RULE FOR BC DRIVER'S LICENCES: BC Licences have large text 'DRIVER'S LICENCE BRITISH COLUMBIA' running vertically along the right edge BY DESIGN. IGNORE this vertical side-text! Instead, look at the face photo and the horizontal text 'Canada', 'KIM MIN', 'DOB'. If the face photo is upright and 'KIM MIN' reads left-to-right, IT IS UPRIGHT (0 degrees)!
               - Analyze step-by-step:
                 1) Is the main body text (Name, DOB, Address) reading normally Left-to-Right? -> Output 0
                 2) Is the main body text reading Bottom-to-Top (running vertically pointing left)? -> Output 90
                 3) Is the main body text reading Top-to-Bottom (running vertically pointing right)? -> Output 270
                 4) Is the main body text upside down? -> Output 180
               *Write your physical observation in `rotation_reasoning` first.*
            4. For EACH grouped document, generate an EXACT filename using our strict CRM manual rules provided below.

            [STRICT CRM MANUAL FILENAME RULES]
            Rule 1 (Name): Korean client -> Full Korean Name without spaces (e.g. 홍길동). Foreign client -> VERY FIRST WORD of English Given Name in Title Case (e.g. Richard, Pham).
            Rule 2 (Dates): MUST be YYYY.MM.DD (e.g. 2022.01.02). Use dots '.' as delimiters.
            Rule 3 (Delimiter): Underscore '_' between Name, Document Category, and Date/Company.
            Rule 4 (Date Range): Use hyphen '-' for ranges (e.g. 2022.04.22-2022.05.22).
            Rule 5 (Capitalization): Every English word MUST be Title Case (Capitalize First Letter).

            [CATEGORIES & CMS SUFFIX RULES]
            1. Passport (including stamp/visa pages): {{Name}}_PP_{{ExpiryDate YYYY.MM.DD}}
            2. Work Permit / Study Permit / Visitor Record (IMM 1442) / Coop / PGWP / BOWP: {{Name}}_{{WP/SP/VR/Coop/PGWP/BOWP}}_{{ExpiryDate YYYY.MM.DD}}
            3. Questionnaire (CanNest Q&A / 질문지): {{Name}}_QA_{{Code}}_{{ReceivedDate YYYY.MM.DD}}
               {{Code}} MUST be EXACTLY one of these short codes — never a description:
                 - Work Permit questionnaire (initial, extension, LMIA, PGWP — any work permit) -> WP
                 - Open Work Permit questionnaire (title says 'OWP' / 'Open Work Permit') -> OWP
                 - Working Holiday (IEC) questionnaire -> WH
                 - Study Permit questionnaire (initial or extension) -> SP
                 - Visitor Record questionnaire (visitor extension) -> VR
                 - Temporary Resident Visa (visitor visa) questionnaire -> TRV
                 - Express Entry questionnaire -> EE
                 - EEBC questionnaire -> EEBC
                 - BC PNP questionnaire -> BCPNP
               NEVER write words like 'Work Permit', 'Extension', 'Initial', 'Renewal', 'Application' in the code.
               e.g. 홍길동_QA_WP_2025.01.14, Richard_QA_SP_2024.11.02, 김철수_QA_EEBC_2025.03.01
            4. Police Certificate: {{Name}}_Police Cert_{{CountryInEnglish}}
            5. Letter of Employment / Confirmation of Employment: {{Name}}_LOE_{{CompanyInEnglish}} (CRITICAL: MUST use 'LOE', NEVER 'Employment Letter'. ONLY for documents certifying past/current employment, usually 1 signature.)
            6. Job Offer / Employment Agreement: {{Name}}_Job Offer_{{CompanyInEnglish}} (CRITICAL: Outlines future terms, usually has BOTH employer and employee signatures.)
            7. Paystub: {{Name}}_Paystub_{{CompanyInEnglish}}_{{StartDate-EndDate}}
            8. Education Certificate / WES: {{Name}}_{{Diploma/Bachelor/Master/Highschool/Certificate/WES}}_{{SchoolName}}
            9. Certificate of Income / NOA: {{Name}}_COI_{{Year YYYY}}
            10. Official English Score: {{Name}}_{{IELTS/CELPIP}}_{{Date YYYY.MM.DD}}
            11. Resume: {{Name}}_Resume_{{ReceivedDate YYYY.MM.DD}} (Keep original .doc/.docx format if uploaded)
            12. Medical Exam / Emedical: {{Name}}_Emedical_{{Year YYYY}}
            13. Marriage Certificate: {{Name}}_Marriage Cert_{{IssueDate YYYY.MM.DD}}
            14. Transcript: {{Name}}_Transcript_{{SchoolName}}
            15. Bank Statement: {{Name}}_Bank Statement_{{Year YYYY}}
            16. Family Certificate: {{Name}}_Family Cert_{{IssueDate YYYY.MM.DD}}
            17. Basic Certificate: {{Name}}_Basic Cert
            18. Birth Certificate: {{Name}}_Birth Cert
            19. Travel Consent: {{ChildName}}_Travel Consent
            20. Provincial ID / ECE Licence: {{Name}}_{{ID/ECE Licence}}_{{ExpiryDate YYYY.MM.DD}}
            21. Driver's Licence: {{Name}}_Driver's Licence_{{ExpiryDate YYYY.MM.DD}} (Use 'Licence', not 'License')
            22. Permanent Resident Card: {{Name}}_PR Card_{{ExpiryDate YYYY.MM.DD}}
            23. Digital Photo / Passport Photo: {{Name}}_Digital Photo (Use .pdf if multiple pages, else .jpg)
            24. T4 (Statement of Remuneration Paid): {{Name}}_T4_{{EmployerNameInEnglish}}_{{Year YYYY}}
            25. Letter of Acceptance / Official Admission Letter (입학허가서): {{Name}}_LOA_{{SchoolNameInEnglish}} (e.g. 공원준_LOA_Metropolitan Community College)

            [CRITICAL FALLBACK RULE FOR UNKNOWN DOCUMENTS]
            - Step 1: If a document does NOT match any categories above, extract the official document title printed at the top of the document (in English, Title Case) and format as: {{Name}}_{{DocumentTitleInEnglish}}.
            - Step 2: If the document title/type is completely ambiguous, set suggested_filename as {{Name}}_Unclassified_확인필요.pdf and set "is_unclassified": true.

            Return ONLY a raw JSON object formatted EXACTLY like this:
            {{
              "page_details": [
                {{
                  "page_index": 1,
                  "rotation_reasoning": "Main body text 'KIM MIN' reads horizontally Left-to-Right. Ignoring BC license vertical side-text. It is upright.",
                  "rotation_needed_clockwise": 0
                }}
              ],
              "documents": [
                {{
                  "client_name": "...",
                  "doc_category": "...",
                  "suggested_filename": "...",
                  "page_indices": [1, 2],
                  "is_unclassified": false,
                  "is_resume": false
                }}
              ]
            }}
            """

      contents = [prompt]
      for p in global_pages:
        contents.append(f"--- Page {p['global_idx']} ---")
        contents.append(
            {"mime_type": "image/jpeg", "data": p["preview_bytes"]}
        )

      try:
        response = safe_generate_content(contents)
        clean_text = (
            response.text.strip().replace("```json", "").replace("```", "")
        )
        data = json.loads(clean_text)

        page_details = data.get("page_details", [])
        rotations = {
            str(item.get("page_index")): item.get(
                "rotation_needed_clockwise", 0
            )
            for item in page_details
        }

        raw_docs_info = data.get("documents", [])

        docs_info = []
        for d in raw_docs_info:
          d["suggested_filename"] = normalize_qa_filename(d.get("suggested_filename", ""))
          c_name = d.get("client_name", "").strip()
          c_cat = d.get("doc_category", "").strip()

          existing = None
          for item in docs_info:
            name1 = (
                item.get("suggested_filename", "")
                .replace(".pdf", "")
                .replace(".jpg", "")
                .replace(".docx", "")
                .replace(".doc", "")
                .strip()
            )
            name2 = (
                d.get("suggested_filename", "")
                .replace(".pdf", "")
                .replace(".jpg", "")
                .replace(".docx", "")
                .replace(".doc", "")
                .strip()
            )
            if name1 == name2:
              existing = item
              break
            if (
                c_name
                and c_cat
                and c_name == item.get("client_name", "").strip()
                and c_cat == item.get("doc_category", "").strip()
            ):
              if (
                  "bank statement" in c_cat.lower()
                  or "paystub" in c_cat.lower()
                  or "statement" in c_cat.lower()
              ):
                existing = item
                break

          if existing:
            new_indices = d.get("page_indices", [])
            for ni in new_indices:
              if ni not in existing["page_indices"]:
                existing["page_indices"].append(ni)
          else:
            docs_info.append(d)

      except Exception as e:
        st.error(f"AI 분석 중 오류가 발생했습니다: {e}")
        docs_info = []
        rotations = {}

      status_text.text(
          "3. 분석된 정보를 바탕으로 최종 서류 결합 및 회전/압축 중입니다..."
      )
      progress_step = 1 / max(len(docs_info), 1)

      failed_docs = []

      for idx, doc_info in enumerate(docs_info):
        indices = doc_info.get("page_indices", [])
        if not indices:
          continue

        final_name = doc_info.get("suggested_filename", f"Document_{idx+1}")
        is_unclassified = (
            doc_info.get("is_unclassified", False)
            or "확인필요" in final_name
        )
        is_resume = (
            doc_info.get("is_resume", False)
            or "resume" in final_name.lower()
        )

        source_names = []
        group_pages = []
        for p_idx in indices:
          p_data = next(
              (p for p in global_pages if p["global_idx"] == p_idx), None
          )
          if p_data:
            group_pages.append(p_data)
            if p_data["original_name"] not in source_names:
              source_names.append(p_data["original_name"])

        unique_src_files = list(set([p["original_name"] for p in group_pages]))
        is_all_from_same_pdf = (
            len(unique_src_files) == 1
            and "pdf" in group_pages[0]["mime_type"].lower()
        )
        is_all_from_same_word = (
            len(unique_src_files) == 1 and group_pages[0]["is_word"]
        )
        needs_rotation = any(
            int(rotations.get(str(p["global_idx"]), 0)) != 0
            for p in group_pages
        )

        if is_resume and is_all_from_same_word:
          base_name, _ = os.path.splitext(final_name)
          orig_ext = os.path.splitext(group_pages[0]["original_name"])[1]
          final_name = base_name + orig_ext

          comp_bytes = group_pages[0]["file_bytes"]
          out_mime = group_pages[0]["mime_type"]
          orig_kb = len(comp_bytes) / 1024
          comp_kb = orig_kb

          src_display = ", ".join(source_names)
          if len(src_display) > 30:
            src_display = src_display[:27] + "..."

          results.append({
              "original_name": f"분석결과 ({src_display})",
              "suggested_filename": final_name,
              "category": doc_info.get("doc_category", "기타"),
              "client_name": doc_info.get("client_name", ""),
              "mime": out_mime,
              "orig_kb": orig_kb,
              "comp_kb": comp_kb,
              "bytes": comp_bytes,
              "is_unclassified": is_unclassified,
          })
          progress_bar.progress(min((idx + 1) * progress_step, 1.0))
          continue

        if not (
            final_name.lower().endswith(".pdf")
            or final_name.lower().endswith(".jpg")
            or final_name.lower().endswith(".jpeg")
            or final_name.lower().endswith(".doc")
            or final_name.lower().endswith(".docx")
        ):
          if "photo" in final_name.lower() and len(indices) == 1:
            final_name += ".jpg"
          else:
            final_name += ".pdf"

        try:
          if is_all_from_same_pdf:
            src_doc = fitz.open(
                stream=group_pages[0]["file_bytes"], filetype="pdf"
            )
            pdf_indices = [p["pdf_page_idx"] for p in group_pages]

            if (
                len(pdf_indices) == len(src_doc)
                and not needs_rotation
                and pdf_indices == list(range(len(src_doc)))
            ):
              merged_pdf_bytes = group_pages[0]["file_bytes"]
              src_doc.close()
              final_processed_bytes = sanitize_and_unlock_pdf(merged_pdf_bytes)
            else:
              new_doc = fitz.open()
              for p_data in group_pages:
                p_idx = p_data["pdf_page_idx"]
                rot = int(rotations.get(str(p_data["global_idx"]), 0))
                
                new_doc.insert_pdf(src_doc, from_page=p_idx, to_page=p_idx)
                if rot != 0:
                  new_doc[-1].set_rotation((new_doc[-1].rotation + rot) % 360)

              merged_pdf_bytes_io = io.BytesIO()
              new_doc.save(merged_pdf_bytes_io, deflate=True)
              new_doc.close()
              src_doc.close()
              merged_pdf_bytes = merged_pdf_bytes_io.getvalue()
              final_processed_bytes = sanitize_and_unlock_pdf(merged_pdf_bytes)

            comp_bytes, out_mime = process_and_compress_file(
                final_processed_bytes,
                "application/pdf",
                final_name,
                already_sanitized=True,
            )

          else:
            new_doc = fitz.open()
            for p_data in group_pages:
              rot = int(rotations.get(str(p_data["global_idx"]), 0))

              if p_data.get("is_word"):
                pdf_page = new_doc.new_page(width=595, height=842)
                w_text = p_data.get("word_text", "")
                pdf_page.insert_text(
                    (50, 50),
                    w_text[:3000]
                    if w_text
                    else f"Word Document: {p_data['original_name']}",
                    fontsize=10,
                )
              elif "pdf" in p_data["mime_type"].lower():
                src_doc = fitz.open(
                    stream=p_data["file_bytes"], filetype="pdf"
                )
                p_idx = p_data["pdf_page_idx"]
                new_doc.insert_pdf(src_doc, from_page=p_idx, to_page=p_idx)
                if rot != 0:
                  new_doc[-1].set_rotation((new_doc[-1].rotation + rot) % 360)
                src_doc.close()
              else:
                img = Image.open(io.BytesIO(p_data["file_bytes"]))
                img = ImageOps.exif_transpose(img)
                if img.mode != "RGB":
                  img = img.convert("RGB")

                if rot != 0:
                  img = img.rotate(-rot, expand=True)

                img_buf = io.BytesIO()
                img.save(img_buf, format="JPEG", quality=92)
                pdf_page = new_doc.new_page(
                    width=img.width, height=img.height
                )
                pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())

            merged_pdf_bytes_io = io.BytesIO()
            new_doc.save(merged_pdf_bytes_io, deflate=True)
            new_doc.close()
            merged_pdf_bytes = merged_pdf_bytes_io.getvalue()

            is_jpeg = final_name.lower().endswith((".jpg", ".jpeg"))

            if is_jpeg and len(indices) == 1:
              p_data = group_pages[0]
              if "pdf" in p_data["mime_type"].lower():
                src_doc = fitz.open(
                    stream=p_data["file_bytes"], filetype="pdf"
                )
                page = src_doc.load_page(p_data["pdf_page_idx"])
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                img = Image.frombytes(
                    "RGB", [pix.width, pix.height], pix.samples
                )
                src_doc.close()
              else:
                img = Image.open(io.BytesIO(p_data["file_bytes"]))
                img = ImageOps.exif_transpose(img)

              try:
                rot = int(rotations.get(str(p_data["global_idx"]), 0))
                if rot != 0:
                  img = img.rotate(-rot, expand=True)
              except Exception:
                pass

              if img.mode != "RGB":
                img = img.convert("RGB")

              img_buf = io.BytesIO()
              img.save(img_buf, format="JPEG", quality=95)
              comp_bytes, out_mime = process_and_compress_file(
                  img_buf.getvalue(),
                  "image/jpeg",
                  final_name,
                  already_sanitized=True,
              )
            else:
              final_processed_bytes = sanitize_and_unlock_pdf(merged_pdf_bytes)
              comp_bytes, out_mime = process_and_compress_file(
                  final_processed_bytes,
                  "application/pdf",
                  final_name,
                  already_sanitized=True,
              )

          orig_kb = sum([len(p["file_bytes"]) for p in group_pages]) / 1024
          comp_kb = len(comp_bytes) / 1024

          src_display = ", ".join(source_names)
          if len(src_display) > 30:
            src_display = src_display[:27] + "..."

          results.append({
              "original_name": f"분석결과 ({src_display})",
              "suggested_filename": final_name,
              "category": doc_info.get("doc_category", "기타"),
              "client_name": doc_info.get("client_name", ""),
              "mime": out_mime,
              "orig_kb": orig_kb,
              "comp_kb": comp_kb,
              "bytes": comp_bytes,
              "is_unclassified": is_unclassified,
          })

        except Exception as e:
          failed_docs.append(final_name)

        progress_bar.progress(min((idx + 1) * progress_step, 1.0))

      if failed_docs:
        status_text.warning(
            "일부 서류 처리에 실패했지만, 나머지 서류의 최적화는"
            " 완료되었습니다."
        )
        st.warning(
            "⚠️ **아래 서류는 파일 손상 또는 변환 중 오류가 발생하여"
            " 제외되었습니다. 원본 파일을 직접 확인해 주세요:**\n\n"
            + "\n".join([f"- {f}" for f in failed_docs])
        )
      else:
        status_text.success(
            "모든 서류의 묶기/분할 및 스마트 회전 최적화가 완료되었습니다."
        )

      st.session_state.analysis_results = results

  if st.session_state.analysis_results:
    st.markdown("---")
    st.subheader("변환 완료된 서류 다운로드")

    zip_buffer = io.BytesIO()
    final_downloads = []

    for idx, item in enumerate(st.session_state.analysis_results):
      col1, col2, col3 = st.columns([3, 3, 2])

      with col1:
        st.write(f"**출처**: `{item['original_name']}`")
        st.caption(
            f"{item['orig_kb']:.1f} KB ➡️ **{item['comp_kb']:.1f} KB**"
        )
        if item.get("is_unclassified"):
          st.warning("⚠️ 규칙 미확인 서류 (파일명 수동 확인 필요)")

      with col2:
        user_edited_name = st.text_input(
            "파일명",
            value=item["suggested_filename"],
            key=f"edit_{idx}",
            label_visibility="collapsed",
        )
        final_downloads.append(
            (user_edited_name, item["bytes"], item["mime"])
        )

      with col3:
        st.download_button(
            "⬇️ 개별 다운로드",
            data=item["bytes"],
            file_name=user_edited_name,
            mime=item["mime"],
            key=f"dl_btn_{idx}",
        )

    with zipfile.ZipFile(
        zip_buffer, "w", zipfile.ZIP_DEFLATED
    ) as zip_file:
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
        use_container_width=True,
    )

    st.markdown("---")
    if st.button("🔄 전체 리셋", type="secondary", use_container_width=True):
      for key in list(st.session_state.keys()):
        if key != "password_correct":
          del st.session_state[key]
      st.session_state.uploader_key = str(uuid.uuid4())
      st.rerun()

# ------------------------------------------
# 메뉴 5: 질문지 기존 정보 기입
# ------------------------------------------
elif app_mode == MENU_5:
    st.title(MENU_5)
    st.caption("기존 손님 서류들과 새 빈 질문지를 올리시면, 알고 있는 정보를 서식 칸에 맞춰 기입해 드립니다.")

    if "menu5_uploader_key" not in st.session_state:
        st.session_state.menu5_uploader_key = str(uuid.uuid4())
    for k in ("qa_rows", "qa_ctx", "qa_sig", "qa_warnings", "qa_base"):
        st.session_state.setdefault(k, None)

    UPLOAD_TYPES = ['pdf', 'docx', 'jpg', 'jpeg', 'png', 'heic', 'heif', 'webp']
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 주 신청자 서류")
        main_files = st.file_uploader(
            "주 신청자의 이전 질문지, 여권, 퍼밋, 이력서 등 (복수 선택)",
            type=UPLOAD_TYPES,
            accept_multiple_files=True,
            key=f"qa_main_files_{st.session_state.menu5_uploader_key}"
        )
    with c2:
        st.subheader("2. 가족 서류 (해당 시)")
        family_files = st.file_uploader(
            "배우자·사실혼 파트너·자녀의 이전 질문지, 여권, 퍼밋, 혼인/가족관계증명서 등",
            type=UPLOAD_TYPES,
            accept_multiple_files=True,
            key=f"qa_family_files_{st.session_state.menu5_uploader_key}",
            help="배우자인지, 사실혼 파트너인지, 자녀인지는 서류와 질문지 내용을 보고 자동으로 구분합니다."
        )
    st.caption("※ 구형 .doc 파일은 Word에서 .docx 또는 PDF로 저장 후 올려주세요.")

    st.subheader("3. 새로 채울 빈 질문지")
    new_q_file = st.file_uploader(
        "채워 넣을 대상 질문지 1부 (입력 칸이 있는 .pdf 또는 .docx)",
        type=['pdf', 'docx'],
        key=f"qa_target_{st.session_state.menu5_uploader_key}"
    )

    # [수정] 파일을 바꾸면 이전 손님의 결과가 남아 잘못 다운로드되는 문제 방지
    current_sig = "|".join([files_signature(main_files or []), files_signature(family_files or []), files_signature([new_q_file])])
    if st.session_state.qa_sig and st.session_state.qa_sig != current_sig:
        st.session_state.qa_rows = None
        st.session_state.qa_ctx = None
        st.session_state.qa_warnings = None
        st.session_state.qa_sig = None

    st.markdown("---")

    def reset_menu5():
        for k in ("qa_rows", "qa_ctx", "qa_sig", "qa_warnings", "qa_base"):
            st.session_state[k] = None
        st.session_state.menu5_uploader_key = str(uuid.uuid4())

    # 결과가 나온 뒤에는 기입 버튼을 회색으로 바꾸고, 아래 다운로드 버튼을 강조
    has_result = bool(st.session_state.qa_rows) and new_q_file is not None
    start_label = "🔁 다시 기입하기" if has_result else "✨ 질문지 정보 기입 시작"
    if st.button(start_label, type="secondary" if has_result else "primary", **_WIDE):
        if not main_files:
            st.warning("1단계: 주 신청자의 서류를 최소 1개 이상 업로드해 주세요.")
        elif not new_q_file:
            st.warning("3단계: 새로 채워야 할 질문지 파일을 업로드해 주세요.")
        elif new_q_file.name.lower().endswith(".docx") and not DOCX_AVAILABLE:
            st.error("서버에 python-docx가 설치되어 있지 않습니다. requirements.txt에 python-docx를 추가해 주세요.")
        else:
            with st.spinner("원본 폼 필드를 분석하여 작성 중입니다. 잠시만 기다려 주세요..."):
                data, warns, ctx = analyze_qa_for_fill(main_files, family_files, new_q_file)
            st.session_state.qa_warnings = warns
            if data and ctx:
                st.session_state.qa_rows = build_review_rows(data, ctx)
                st.session_state.qa_ctx = ctx
                st.session_state.qa_sig = current_sig
                st.session_state.qa_base = os.path.splitext(new_q_file.name)[0]
                st.rerun()  # 버튼 색 바로 반영

    for w in (st.session_state.qa_warnings or []):
        st.warning(w)

    # 💡 결과 다운로드 (검토 표 없이 바로 — 다운로드한 파일에서 더블체크)
    if st.session_state.qa_rows is not None and new_q_file is not None:
        rows = st.session_state.qa_rows
        if not rows:
            st.info("자료에서 질문지에 기입할 수 있는 정보를 찾지 못했습니다.")
        else:
            for tw in (st.session_state.qa_ctx or {}).get("table_warnings", []) or []:
                st.warning(tw)
            try:
                out_bytes, out_name, out_mime, filled, skipped = build_output_file(
                    rows, st.session_state.qa_ctx, new_q_file.getvalue(), st.session_state.qa_base)
                st.success(f"✅ {filled}개 칸에 기입했습니다. 확인하지 못한 항목은 비워두었으니 손님께 마저 작성해 달라고 요청하세요.")
                run = st.session_state.get("gemini_last_run")
                if run:
                    st.caption(f"AI 분석: {run['model']} · {run['seconds']}초")
                excluded = [r for r in rows if not r.get("기입") and r.get("확인 필요")]
                if excluded or skipped:
                    with st.expander(f"기입하지 않은 항목 {len(excluded) + len(skipped)}개 (값이 칸과 맞지 않거나 위치를 못 찾음)"):
                        for r in excluded:
                            st.write(f"- {r['주변 문구'] or r['칸 (Key)']}: {r['값']}  — {r['확인 필요']}")
                        for k in skipped:
                            st.write(f"- {k}: 서식에서 위치를 찾지 못함")
                st.download_button(
                    label="📥 Info Filled 질문지 내려받기",
                    data=out_bytes,
                    file_name=out_name,
                    mime=out_mime,
                    type="primary",
                    **_WIDE,
                )
            except Exception as e:
                st.error(f"파일 생성 중 오류가 발생했습니다: {e}")

    # 전체 리셋은 맨 아래
    st.markdown("")
    st.button("🔄 전체 리셋", **_WIDE, on_click=reset_menu5)
