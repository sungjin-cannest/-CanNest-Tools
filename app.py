import streamlit as st

# 📌 가장 먼저 실행되어야 하는 Streamlit 페이지 설정
APP_NAME = "[Demo]CanNest 통합 업무 시스템"
st.set_page_config(page_title=APP_NAME, layout="wide")

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
import gc
from zoneinfo import ZoneInfo

try:
    import pymupdf as fitz  # PyMuPDF 최신 버전
except ImportError:
    import fitz  # 구버전 호환

try:
    from google import genai as gg
    from google.genai import types as gtypes
    NEW_SDK = True
except ImportError:
    import google.generativeai as genai
    NEW_SDK = False

from PIL import Image, ImageDraw, ImageFont, ImageOps
import pillow_heif  # HEIC 지원 라이브러리

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

try:
    _WIDE = {"width": "stretch"} if "width" in _inspect.signature(st.download_button).parameters else {"use_container_width": True}
except Exception:
    _WIDE = {"use_container_width": True}

HANGUL_RE = re.compile(r"[\u3131-\u318E\uAC00-\uD7A3]")

def safe_download_button(label, data=None, file_name=None, mime=None, **kwargs):
    return st.download_button(label, data=data, file_name=file_name, mime=mime, **kwargs)

IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.heic', '.heif', '.webp')

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
        st.title(f"🔒 {APP_NAME}")
        st.text_input("접속 비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.title(f"🔒 {APP_NAME}")
        st.text_input("접속 비밀번호를 입력하세요", type="password", on_change=password_entered, key="password")
        st.error("비밀번호가 틀렸습니다.")
        return False
    return True

if not check_password():
    st.stop()

# ==========================================
# 1. API 키 및 모델 설정 (실제 안정 모델 최우선 배치)
# ==========================================
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]

@st.cache_resource(show_spinner=False)
def get_gemini_client(api_key):
    return gg.Client(api_key=api_key)

if NEW_SDK:
    gclient = get_gemini_client(GEMINI_API_KEY)
else:
    genai.configure(api_key=GEMINI_API_KEY)

THINKING_LEVEL = "low"

# 💡 실존하는 정식 Flash 모델을 최우선으로 배치하여 404 재시도 루프 제거
CANDIDATE_MODELS = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-flash-latest']

def _is_model_gone(msg):
    m = msg.lower()
    return ("404" in m or "not found" in m or "no longer available" in m
            or "not supported" in m or "deprecated" in m)

@st.cache_data(ttl=6 * 3600, show_spinner=False)
def discover_flash_models():
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
    if not isinstance(s, str):
        return s
    try:
        s.encode("utf-8")
        return s
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
            try:
                if "2.5" in model_name:
                    cfg["thinking_config"] = gtypes.ThinkingConfig(thinking_budget=1024)
                else:
                    cfg["thinking_config"] = gtypes.ThinkingConfig(thinking_level=thinking_level)
            except Exception:
                pass
        return gclient.models.generate_content(
            model=model_name, contents=_to_new_sdk_parts(contents),
            config=gtypes.GenerateContentConfig(**cfg))
    gen_config = {}
    if json_mode:
        gen_config = {"temperature": 0, "response_mime_type": "application/json"}
    return genai.GenerativeModel(model_name, generation_config=gen_config).generate_content(contents)

def safe_generate_content(contents, json_mode=False, thinking=None):
    contents = _clean_contents(contents)
    last_error = None
    working = st.session_state.get("gemini_working_model")
    order = ([working] if working else []) + CANDIDATE_MODELS
    tried = set()

    def model_iter():
        for n in order:
            yield n
        for n in discover_flash_models():
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
                                          "api_key_invalid", "permission denied", "quota exceeded")):
                    raise RuntimeError(f"Gemini API 인증/결제 오류: {e}") from e
                if use_thinking and "thinking" in msg:
                    use_thinking = None
                    continue
                if _is_model_gone(msg):
                    break
                if any(k in msg for k in ["429", "500", "503", "unavailable", "overloaded", "deadline", "resource exhausted"]):
                    attempt += 1
                    time.sleep(2 * attempt)
                    continue
                raise
    raise last_error

def parse_json_response(text):
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
                row_txt = " | ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
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
        display_str = f"[Word Document: {file_obj.name}]\n\n" + (text[:800] if text else "Word Document Content")
        draw.text((40, 40), display_str, fill=(0, 0, 0))
        return img
    elif file_obj.type == "application/pdf" or fname.endswith(".pdf"):
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
    else:
        img = Image.open(io.BytesIO(file_bytes))
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")

    max_dim = max(img.width, img.height)
    if max_dim > 1500:
        ratio = 1500.0 / float(max_dim)
        new_size = (int(img.width * ratio), int(img.height * ratio))
        img = img.resize(new_size, Image.Resampling.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=75)
    buf.seek(0)
    return Image.open(buf)

def format_full_name(surname, given_name):
    s = surname.strip()
    g = given_name.strip()
    if not s and not g: return ""
    if not s: return g
    if not g: return s
    return f"{g} {s}"

def set_smart_widget_value(widget, value, default_fontsize=11, min_fontsize=5.5):
    val_str = str(value) if value is not None else ""
    widget.field_value = val_str

    if hasattr(widget, "field_flags") and widget.field_flags:
        widget.field_flags &= ~1

    try:
        widget.text_font = "CourierNew"
    except Exception:
        try:
            widget.text_font = "Cour"
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
        acroform_xref = int(acroform_res[1].split()[0]) if acroform_res[0] == "xr" else (catalog_xref if acroform_res[0] == "dict" else None)

        if acroform_xref and font_xrefs:
            font_dict_entries = " ".join([f"/{alias} {xref} 0 R" for alias, xref in font_xrefs.items()])
            doc.xref_set_key(acroform_xref, "DR/Font", f"<< {font_dict_entries} >>")
    except Exception:
        pass

def sanitize_and_unlock_pdf(pdf_bytes):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page in doc:
            for widget in list(page.widgets()):
                field_type_str = getattr(widget, "field_type_string", "").lower()
                if "sig" in field_type_str or widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
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
    if "word" in mime_type.lower() or "doc" in mime_type.lower() or ext in [".doc", ".docx"]:
        word_text = read_word_document_text(file_bytes, file_name)
        if word_text:
            return [f"\n--- [Word Document: {file_name}] ---\n{word_text[:20000]}\n"]

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
                doc.close()
                return [f"\n--- [Document: {file_name}] ---\n{text[:20000]}\n"]
            else:
                images = []
                for page_num in range(min(len(doc), 10)):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=65)
                    buf.seek(0)
                    images.append({"mime_type": "image/jpeg", "data": buf.getvalue()})
                doc.close()
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
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        return age < 19
    except Exception:
        return True

# ==========================================
# 4. PDF 서식 채우기 로직
# ==========================================
def extract_imm5476_info(image):
    prompt = """
    Analyze this document carefully.
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
        clean_text = response.text.strip().replace("```json", "").replace("```", "")
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"정보 추출 오류: {e}")
        return None

def extract_all_passports_batch(has_non_acc, images):
    prompt = f"""
    You are an expert OCR system specialized in international passports.
    I am providing {len(images)} passport image(s) in exact order.
    For EACH passport image, extract:
    - surname: Title Case
    - given_name: Title Case
    - dob: YYYY-MM-DD
    - passport_number: Uppercase alphanumeric
    - gender: "F" or "M"

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
        clean_text = response.text.strip().replace("```json", "").replace("```", "")
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"여권 일괄 추출 오류: {e}")
        return None

def extract_case_prep_info(tmpl_bytes, client_files):
    prompt = """
    You are an expert Canadian immigration case prep assistant.
    Carefully scan ALL attached client documents to extract all required information matching the IMM form fields.
    Return ONLY a raw valid JSON object:
    {
      "sections": [
        {
          "section": "Section Name",
          "fields": [
            { "field": "Field Label", "value": "Extracted Value or Mismatch Warning", "source": "Source Doc Name(s)" }
          ]
        }
      ]
    }
    """
    contents = [prompt]
    contents.extend(prepare_document_for_gemini(tmpl_bytes, "application/pdf", "Blank_IMM_Form.pdf"))
    contents.extend(batch_process_client_files(client_files))

    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace("```json", "").replace("```", "")
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"서류 정리 오류: {e}")
        return None

def fill_imm5476(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    target_data = {
        "surname": data.get("surname", ""), "given": data.get("given_name", ""),
        "dob": data.get("dob", ""), "email": data.get("email", ""),
        "address_phone": data.get("address_phone", ""),
        "uci": data.get("uci", "").replace("-", ""), "signDate": data.get("signDate", ""),
    }
    flags = {"surname": False, "given": False, "dob": False, "email": False, "address_phone": False, "uci": False}
    page_date_counters = {}

    for page_idx, page in enumerate(doc):
        page_date_counters[page_idx] = 0
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
            elif ("telephone" in fname_lower or "address" in fname_lower) and page_idx == 0 and not flags["address_phone"]:
                val_to_set = target_data["address_phone"]; flags["address_phone"] = True
            elif ("uci" in fname_lower or "unique client identifier" in fname_lower) and not flags["uci"]:
                val_to_set = target_data["uci"]; flags["uci"] = True
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
        child_widgets = [("Information about travelling children", "yyyymmdd"), ("1_2", "2_2"), ("1_3", "2_3")]

        for widget in page.widgets():
            fname = widget.field_name.strip() if widget.field_name else ""
            if not fname: continue
            fname_lower = fname.lower()

            field_type_str = getattr(widget, "field_type_string", "").lower()
            if "check" in field_type_str or "radio" in field_type_str or "check box" in fname_lower or "checkbox" in fname_lower or "alone" in fname_lower:
                try:
                    widget.field_value = "Off"
                    widget.update()
                except Exception:
                    pass
                continue

            val_to_set = None
            if fname == "1": val_to_set = data.get("non_acc_name", "")
            elif fname == "2": val_to_set = data.get("non_acc_address", "")
            elif fname == "3": val_to_set = data.get("non_acc_phone", "")
            elif fname == "email": val_to_set = data.get("non_acc_email", "")
            elif fname == "This child or these children hashave my or our consent to travel with": val_to_set = data.get("acc_name", "")
            elif fname == "Relationship with Children 1": val_to_set = data.get("acc_relationship", "")
            elif fname == "Relationship with Children 2": val_to_set = data.get("acc_passport", "")
            elif fname == "I give my consent for this child to travel to": val_to_set = "Canada"
            elif fname == "2_4" or fname_lower == "to stay with": val_to_set = data.get("acc_name", "")
            elif fname == "At the following addresses 1": val_to_set = data.get("trip_address", "")
            elif fname == "At the following addresses 2": val_to_set = data.get("trip_phone", "")
            elif fname == "email_2": val_to_set = data.get("trip_email", "")
            elif fname == "yyyymmdd_2": val_to_set = data.get("sign_date", "")
            elif fname == "1_4" or "travel date" in fname_lower or fname_lower == "date" or "from" in fname_lower or "departure" in fname_lower or "start date" in fname_lower:
                val_to_set = data.get("trip_date", "")
            elif fname_lower == "to" or "return" in fname_lower or "end date" in fname_lower:
                val_to_set = ""
            else:
                for idx, (name_key, dob_key) in enumerate(child_widgets):
                    if idx < len(page_children):
                        if fname == name_key: val_to_set = page_children[idx].get("name", "")
                        elif fname == dob_key: val_to_set = page_children[idx].get("dob", "")

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
# 5. CRM 스마트 압축 및 PDF 안전 병합 엔진 (CPU 최적화)
# ==========================================
def process_and_compress_file(file_bytes, mime_type, target_filename, already_sanitized: bool = False):
    is_jpeg = target_filename.lower().endswith((".jpg", ".jpeg"))

    if is_jpeg:
        img = Image.open(io.BytesIO(file_bytes))
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")

        max_dim = max(img.width, img.height)
        if max_dim > 1800:
            ratio = 1800.0 / float(max_dim)
            img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)

        buf = io.BytesIO()
        # 💡 CPU 잠김 방지: optimize=True 제거
        img.save(buf, format="JPEG", quality=80)
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
                return (file_bytes, "application/pdf") if already_sanitized else (sanitize_and_unlock_pdf(file_bytes), "application/pdf")

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
                if max(eff_width, eff_height) > 1800:
                    zoom = 1.0

                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                img_buf = io.BytesIO()
                # 💡 CPU 잠김 방지: optimize=True 제거
                img.save(img_buf, format="JPEG", quality=quality)
                img_buf.seek(0)

                new_page_width = pix.width / zoom
                new_page_height = pix.height / zoom
                pdf_page = new_doc.new_page(width=new_page_width, height=new_page_height)
                pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())

            output_pdf = io.BytesIO()
            new_doc.save(output_pdf, deflate=True, garbage=1)
            new_doc.close()
            doc.close()

            compressed_bytes = output_pdf.getvalue()
            final_bytes = compressed_bytes if len(compressed_bytes) < len(file_bytes) else file_bytes
            return (final_bytes, "application/pdf") if already_sanitized else (sanitize_and_unlock_pdf(final_bytes), "application/pdf")

        else:
            target_dpi = 150
            quality = 70
            img = Image.open(io.BytesIO(file_bytes))
            img = ImageOps.exif_transpose(img)
            if img.mode != "RGB":
                img = img.convert("RGB")

            max_dim = max(img.width, img.height)
            if max_dim > 1600:
                ratio = 1600.0 / float(max_dim)
                img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)

            img_buf = io.BytesIO()
            img.save(img_buf, format="JPEG", quality=quality)
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
            return (compressed_bytes, "application/pdf") if already_sanitized else (sanitize_and_unlock_pdf(compressed_bytes), "application/pdf")

# ==========================================
# 5-0. [메뉴 4] CRM 도우미 함수들
# ==========================================
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
    code = next((c for c, pat in QA_CODE_RULES if re.search(pat, type_text, re.I)), "확인필요")
    return f"{name}_QA_{code}" + (f"_{date}" if date else "") + ext

CONTENT_FILENAME_RULES = [
    ("Submission Confirmation Letter",
     re.compile(r"IMM\s*5988|confirms\s+that\s+your\s+application\s+has\s+been\s+received"
                r"|authori[sz]ed\s+to\s+continue\s+(working|studying).{0,120}until\s+a\s+decision", re.I | re.S)),
]

def crm_pages_text(global_pages, page_indices, file_storage, max_chars=6000):
    out, cache = [], {}
    for p in global_pages:
        if p["global_idx"] not in page_indices:
            continue
        try:
            if "pdf" in p["mime_type"].lower():
                fn = p["original_name"]
                if fn not in cache:
                    cache[fn] = fitz.open(stream=file_storage[fn], filetype="pdf")
                out.append(cache[fn].load_page(p["pdf_page_idx"]).get_text("text"))
            elif p.get("word_text"):
                out.append(p["word_text"])
        except Exception:
            pass
        if sum(len(x) for x in out) > max_chars:
            break
    for d in cache.values():
        d.close()
    return "\n".join(out)[:max_chars]

def fix_filename_by_content(filename, text):
    if not filename or not text:
        return filename
    for doc_name, pat in CONTENT_FILENAME_RULES:
        if pat.search(text):
            if doc_name.lower() in filename.lower():
                return filename
            name = filename.split("_")[0].strip()
            ext = os.path.splitext(filename)[1] if re.search(r"\.(pdf|jpe?g|png|docx?)$", filename, re.I) else ""
            return f"{name}_{doc_name}{ext}"
    return filename

UNKNOWN_NAME_RE = re.compile(r"^(unknown|unkown|client|name|n\s*/?\s*a|미상|이름\s*미상|이름없음)(?=[_.]|$)", re.I)

def replace_unknown_name(filename, name):
    if not name: return filename
    return UNKNOWN_NAME_RE.sub(name.strip(), filename, count=1)

def ensure_extension(filename, mime):
    filename = (filename or "").strip() or "Document"
    if re.search(r"\.(pdf|jpe?g|png|docx?)$", filename, re.I):
        return filename
    ext = {"application/pdf": ".pdf", "image/jpeg": ".jpg", "image/png": ".png",
           "application/msword": ".doc"}.get(mime, ".docx" if "word" in (mime or "") else ".pdf")
    return filename + ext

import shutil
import subprocess
import tempfile

def _find_soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")

@st.cache_data(show_spinner=False, max_entries=50)
def word_to_pdf_bytes(file_bytes, file_name):
    soffice = _find_soffice()
    if soffice:
        try:
            with tempfile.TemporaryDirectory() as tmp:
                ext = os.path.splitext(file_name)[1].lower() or ".docx"
                src = os.path.join(tmp, "input" + ext)
                with open(src, "wb") as f: f.write(file_bytes)
                profile = "file://" + os.path.join(tmp, "lo_profile")
                subprocess.run(
                    [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
                     "--convert-to", "pdf", "--outdir", tmp, src],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=40,
                    env={**os.environ, "HOME": tmp})
                out = os.path.join(tmp, "input.pdf")
                if os.path.exists(out) and os.path.getsize(out) > 0:
                    with open(out, "rb") as f: return f.read(), "libreoffice"
        except Exception:
            pass
    try:
        return _docx_to_pdf_basic(file_bytes), "basic"
    except Exception:
        return None, "failed"

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

def _docx_images(element, part):
    out = []
    for drawing in element.iter(_W + "drawing"):
        blip = next(drawing.iter(_A + "blip"), None)
        if blip is None: continue
        rid = blip.get(_R + "embed")
        try: blob = part.related_parts[rid].blob
        except Exception: continue
        ext = next(drawing.iter(_WP + "extent"), None)
        w = int(ext.get("cx")) / 12700 if ext is not None else 100
        h = int(ext.get("cy")) / 12700 if ext is not None else 60
        floating = next(drawing.iter(_WP + "anchor"), None) is not None
        out.append((blob, w, h, floating))
    return out

def _docx_to_pdf_basic(file_bytes):
    import html as _html
    from docx.text.paragraph import Paragraph as _Para
    from docx.table import Table as _Table
    d = docx.Document(io.BytesIO(file_bytes))
    archive = fitz.Archive()
    img_no = [0]

    def add_img(blob, w, h):
        img_no[0] += 1
        name = f"img{img_no[0]}"
        try:
            pil = Image.open(io.BytesIO(blob)).convert("RGB")
            b = io.BytesIO(); pil.save(b, format="PNG"); blob = b.getvalue()
        except Exception: pass
        archive.add((blob, name + ".png"))
        return f"<img src='{name}.png' width='{min(w, 468):.0f}' height='{h * min(1, 468 / max(w, 1)):.0f}'/>"

    def para_html(p, part):
        out = ""
        for r in p.runs:
            for blob, w, h, _fl in _docx_images(r._element, part):
                out += add_img(blob, w, h)
            t = _html.escape(r.text)
            if not t: continue
            if r.bold: t = f"<b>{t}</b>"
            if r.italic: t = f"<i>{t}</i>"
            if r.underline: t = f"<u>{t}</u>"
            out += t
        return out or _html.escape(p.text)

    parts = []
    for child in d.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            p = _Para(child, d)
            txt = para_html(p, d.part)
            style = (p.style.name or "").lower() if p.style is not None else ""
            is_list = "list" in style or child.find(".//" + _W + "numPr") is not None
            align = {1: "center", 2: "right", 3: "justify"}.get(int(p.alignment) if p.alignment is not None else 0, "left")
            if is_list:
                if txt.strip(): parts.append(f"<ul><li>{txt}</li></ul>")
            elif style.startswith("heading"):
                parts.append(f"<h3 style='text-align:{align}'>{txt}</h3>")
            else:
                parts.append(f"<p style='text-align:{align}'>{txt or '&nbsp;'}</p>")
        elif tag == "tbl":
            t = _Table(child, d)
            rows = "".join("<tr>" + "".join(f"<td>{_html.escape(c.text)}</td>" for c in r.cells) + "</tr>" for r in t.rows)
            parts.append(f"<table border='1' cellspacing='0' cellpadding='3'>{rows}</table>")

    mediabox = fitz.paper_rect("letter")
    where = fitz.Rect(54, 40, mediabox.width - 54, mediabox.height - 40)
    css = "* {font-family: sans-serif; font-size: 10.5pt;} p {margin: 0 0 6px 0;} td {font-size: 9.5pt;}"
    story = fitz.Story(html="".join(parts), user_css=css, archive=archive)
    buf = io.BytesIO()
    writer = fitz.DocumentWriter(buf)
    more = True
    while more:
        dev = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()
    return buf.getvalue()

CRM_MAX_PAGES = 150
CRM_THINKING_LEVEL = "low"
CRM_PAYLOAD_LIMIT = 10 * 1024 * 1024

def shrink_previews_to_fit(pages, limit=CRM_PAYLOAD_LIMIT):
    total = sum(len(p["preview_bytes"]) for p in pages)
    tries = 0
    while total > limit and tries < 3:
        scale = (limit / total) ** 0.5 * 0.85
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
# 5-2. [메뉴 5] 질문지 채우기 엔진 도우미
# ==========================================
BUTTON_TYPES = (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON)
CHOICE_TYPES = (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX)
TRUE_WORDS = {'yes', 'true', 'v', 'x', 'on', '1', 'checked', 'y'}

def _safe_on_state(w):
    try:
        s = w.on_state()
        if s and str(s).lower() != "off": return str(s)
    except Exception: pass
    return None

SECTION_HEADER_RE = re.compile(r"(개인\s*정보|가족\s*관계\s*정보|personal\s+(information|details)|family\s+(information|members)|background\s+information)", re.I)
PERSON_HEADER_RE = re.compile(r"^[\[\(\s]*(main\s+applicant|principal\s+applicant|applicant|spouse|common[\s-]*law\s+partner|partner|dependa?e?nt\s+child|child|신청자|주\s*신청자|배우자|자녀)", re.I)

def _person_from_text(t):
    t = (t or "").lower()
    m = re.search(r"자녀\s*\(?(\d+)\)?|child\s*\(?(\d+)\)?", t)
    if m: return f"CHILD {m.group(1) or m.group(2)}"
    if re.match(r"^[\[\(\s]*(자녀|dependa?e?nt|child)", t): return "CHILD"
    if re.match(r"^[\[\(\s]*(배우자|spouse|common[\s-]*law|partner)", t): return "SPOUSE"
    return "MAIN APPLICANT"

def detect_sections(doc):
    headers = []
    for p_idx, page in enumerate(doc):
        for b in page.get_text("blocks"):
            first_line = clean_text((b[4] or "").strip().split("\n")[0].strip())
            if not first_line or len(first_line) > 80: continue
            is_person = bool(PERSON_HEADER_RE.search(first_line)) and (SECTION_HEADER_RE.search(first_line) or first_line.lstrip().startswith("["))
            is_general = bool(SECTION_HEADER_RE.search(first_line)) or bool(re.match(r"^\[\s*[A-Z][A-Z &/-]{3,}\s*\]", first_line))
            if is_person or is_general:
                title = re.split(r"\s+[-–]\s+If applicable|\(\*|\s+-\s+모든", first_line)[0].strip()
                headers.append((p_idx, b[1], title, _person_from_text(first_line) if is_person else None))
    headers.sort(key=lambda h: (h[0], h[1]))
    return headers

def person_for(headers, page_idx, y):
    person = "MAIN APPLICANT"
    for hp, hy, title, who in headers:
        if hp < page_idx or (hp == page_idx and hy <= y + 2):
            if who: person = who
        else: break
    return person

def section_for(headers, page_idx, y):
    current = ""
    for hp, hy, title, who in headers:
        if hp < page_idx or (hp == page_idx and hy <= y + 2): current = title
        else: break
    return current

CURRENT_ADDRESS_RE = re.compile(r"current\s+(residential\s+|home\s+)?address|residential\s+address|present\s+address|home\s+address", re.I)
CURRENT_ADDRESS_KO_RE = re.compile(r"현재(거주지|집주소|거주하는곳|사는곳)")

def _is_current_address_label(text):
    if not text or re.search(r"사망|deceased", text, re.I): return False
    return bool(CURRENT_ADDRESS_RE.search(text) or CURRENT_ADDRESS_KO_RE.search(re.sub(r"\s+", "", text)))

RECENT_ENTRY_PHRASES = ["가장 최근 입국", "최근 입국", "most recent entry", "last entry", "latest entry"]

def mark_always_blank(doc, fields):
    for f in fields:
        if f["type"] != fitz.PDF_WIDGET_TYPE_TEXT: continue
        if re.search(r"family|가족", f.get("section") or "", re.I): continue
        page = doc[f["page"]]
        x0, y0, x1, y1 = f["rect"]
        left = " ".join(page.get_text("text", clip=fitz.Rect(max(0, x0 - 220), y0 - 1, x0, y1 + 1)).split())
        if _is_current_address_label(f.get("printed_label") or "") or _is_current_address_label(left):
            f["always_blank"] = "현재 거주지 주소"

def compute_printed_labels(doc, fields):
    by_page = {}
    for f in fields: by_page.setdefault(f["page"], []).append(f)
    for p_idx, flist in by_page.items():
        page = doc[p_idx]
        words = page.get_text("words")
        rects = [f["rect"] for f in flist]
        for f in flist:
            if f["type"] in BUTTON_TYPES:
                f["printed_label"] = ""; continue
            x0, y0, x1, y1 = f["rect"]
            block_x = max([r[2] for r in rects if r[2] <= x0 + 1 and r[1] < y1 - 2 and r[3] > y0 + 2 and r != f["rect"]] + [x0 - 260])
            left = [w for w in words if w[2] <= x0 + 2 and w[0] >= block_x - 1 and y0 - 2 <= (w[1] + w[3]) / 2 <= y1 + 2]
            label = " ".join(w[4] for w in sorted(left, key=lambda w: (round(w[1] / 3), w[0])))
            f["printed_label"] = clean_text(label.strip())[:120]

def extract_pdf_form_fields(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    headers = detect_sections(doc)
    raw = []
    name_counts = {}
    for p_idx, page in enumerate(doc):
        for w in page.widgets():
            if w.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE: continue
            r = w.rect
            clip = fitz.Rect(max(0, r.x0 - 120), max(0, r.y0 - 20), min(page.rect.width, r.x1 + 60), min(page.rect.height, r.y1 + 20))
            nearby = clean_text(" ".join(page.get_text("text", clip=clip).split())[:160])
            info = {
                "name": w.field_name or "", "label": clean_text((w.field_label or "").strip()),
                "type": w.field_type, "type_str": w.field_type_string,
                "on_state": _safe_on_state(w) if w.field_type in BUTTON_TYPES else None,
                "choices": [clean_text(c if isinstance(c, str) else c[-1]) for c in (w.choice_values or [])] if w.field_type in CHOICE_TYPES else [],
                "nearby": nearby, "section": section_for(headers, p_idx, r.y0),
                "person": person_for(headers, p_idx, r.y0), "page": p_idx, "xref": w.xref, "rect": (r.x0, r.y0, r.x1, r.y1),
            }
            raw.append(info)
            if w.field_type in BUTTON_TYPES:
                name_counts[info["name"]] = name_counts.get(info["name"], 0) + 1

    fields, used = [], set()
    btn_index = {}
    for f in raw:
        key = clean_text(f["name"])
        if f["type"] in BUTTON_TYPES and name_counts.get(f["name"], 0) > 1:
            btn_index[f["name"]] = btn_index.get(f["name"], 0) + 1
            key = f"{key} [{clean_text(f['on_state'] or f'#{btn_index[f['name']]}')}]"
        if f["type"] not in BUTTON_TYPES and key in used:
            f["key"] = key; f["duplicate"] = True; fields.append(f); continue
        base, n = key, 2
        while key in used:
            key = f"{base} ({n})"; n += 1
        used.add(key)
        f["key"] = key; f["duplicate"] = False; fields.append(f)

    compute_printed_labels(doc, fields)
    mark_always_blank(doc, fields)
    doc.close()
    return fields, False

def detect_tables(pdf_bytes, fields):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    cand = [f for f in fields if not f["duplicate"] and f["type"] == fitz.PDF_WIDGET_TYPE_TEXT]
    cand.sort(key=lambda f: (f["page"], f["rect"][1], f["rect"][0]))
    rows = []
    for f in cand:
        if rows and rows[-1]["page"] == f["page"] and abs(rows[-1]["y0"] - f["rect"][1]) <= 4:
            rows[-1]["fields"].append(f)
            rows[-1]["y1"] = max(rows[-1]["y1"], f["rect"][3])
        else:
            rows.append({"page": f["page"], "y0": f["rect"][1], "y1": f["rect"][3], "fields": [f]})
    for r in rows:
        r["fields"].sort(key=lambda f: f["rect"][0])
        r["xs"] = [f["rect"][0] for f in r["fields"]]

    groups, cur = [], []
    for i, r in enumerate(rows):
        ok = len(r["fields"]) >= 2
        if ok and cur:
            prev = cur[-1]
            if not (prev["page"] == r["page"] and r["y0"] - prev["y1"] < 15):
                groups.append(cur); cur = []
        if ok: cur.append(r)
        else:
            if cur: groups.append(cur); cur = []
    if cur: groups.append(cur)

    tables = []
    for g in groups:
        if len(g) < 3: continue
        col_xs = sorted(list({x for r in g for x in r["xs"]}))
        first = g[0]
        col_names = [f"Column {ci + 1}" for ci in range(len(col_xs))]
        t_rows = []
        for r in g:
            cells = {min(range(len(col_xs)), key=lambda c: abs(col_xs[c] - f["rect"][0])): f["key"] for f in r["fields"]}
            t_rows.append(cells)
        tables.append({
            "id": f"T{len(tables) + 1}", "page": first["page"],
            "person": first["fields"][0].get("person", "MAIN APPLICANT"),
            "section": first["fields"][0].get("section", ""),
            "title": f"Table on Page {first['page'] + 1}",
            "columns": col_names, "rows": t_rows, "current_row": None, "to_col": None,
        })
        for cells in t_rows:
            for k in cells.values():
                for f in fields:
                    if f["key"] == k: f["table"] = tables[-1]["id"]
    doc.close()
    return tables

def format_table_catalog(tables):
    out = [f"TABLE {t['id']} | PERSON: {t['person']} | ROWS: {len(t['rows'])}" for t in tables]
    return "\n".join(out)

def format_field_catalog(fields):
    lines = []
    for f in fields:
        if f["duplicate"] or f.get("table") or f.get("always_blank"): continue
        lines.append(f"KEY: {json.dumps(f['key'], ensure_ascii=False)} | PERSON: {f.get('person')} | LABEL: {f.get('printed_label','')}")
    return "\n".join(lines[:120])

def place_table_records(tables, ai_tables):
    values, labels = {}, {}
    by_id = {t["id"]: t for t in tables}
    for tid, records in (ai_tables or {}).items():
        t = by_id.get(tid)
        if not t or not isinstance(records, list): continue
        for ri, rec in enumerate(records):
            if ri >= len(t["rows"]): break
            cells = t["rows"][ri]
            for ci in range(len(t["columns"])):
                val = rec.get(f"c{ci+1}")
                if val and ci in cells:
                    values[cells[ci]] = str(val).strip()
                    labels[cells[ci]] = f"[{t['id']}] Row {ri+1}"
    return values, labels, []

def qa_batch_process_client_files(files, owner=None):
    def worker(f):
        lbl = f"{owner} | {f.name}" if owner else f.name
        ext = os.path.splitext(f.name)[1].lower()
        if ext in ('.doc', '.docx'):
            txt = read_word_document_text(f.getvalue(), f.name)
            return [f"\n--- [Doc: {lbl}] ---\n{txt[:15000]}\n"], None
        elif ext in ('.jpg', '.jpeg', '.png', '.heic', '.webp'):
            img = Image.open(io.BytesIO(f.getvalue()))
            img = ImageOps.exif_transpose(img)
            if img.mode != "RGB": img = img.convert("RGB")
            if max(img.size) > 1200:
                img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=75)
            return [f"\n--- [Doc: {lbl}] ---\n", {"mime_type": "image/jpeg", "data": buf.getvalue()}], None
        else:
            return prepare_document_for_gemini(f.getvalue(), f.type or "application/pdf", lbl), None

    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(executor.map(worker, files))
    flat, warns = [], []
    for parts, w in results:
        flat.extend(parts)
        if w: warns.append(w)
    return flat, warns

QA_RULES = """
[RULES]
1. UNKNOWN INFO = OMIT. NEVER guess.
2. DATES: YYYY-MM-DD.
3. LANGUAGE: All in English (Latin letters).
"""

def analyze_qa_for_fill(main_files, family_files, target_file):
    target_bytes = target_file.getvalue()
    source_parts, warnings = qa_batch_process_client_files(main_files or [], "MAIN APPLICANT")
    if family_files:
        fam_parts, fam_warn = qa_batch_process_client_files(family_files, "FAMILY MEMBER")
        source_parts += fam_parts
        warnings += fam_warn

    fields, _ = extract_pdf_form_fields(target_bytes)
    tables = detect_tables(target_bytes, fields)

    prompt = f"""
    You are an expert Canadian immigration assistant filling out a questionnaire.
    {QA_RULES}
    OUTPUT ONLY JSON: {{"pdf_fields": {{"<KEY>": "<value>"}}, "tables": {{"<TABLE ID>": [{{"c1": "val", "c2": "val"}}]}}}}
    TARGET FORM FIELDS:
    {format_field_catalog(fields)}
    TARGET FORM TABLES:
    {format_table_catalog(tables)}
    """
    try:
        response = safe_generate_content([prompt] + source_parts, json_mode=True, thinking=THINKING_LEVEL)
        data = parse_json_response(response.text)
        return data, warnings, {"kind": "pdf", "fields": fields, "tables": tables}
    except Exception as e:
        st.error(f"분석 중 오류 발생: {e}")
        return None, warnings, None

def build_review_rows(data, context):
    by_key = {f["key"]: f for f in context["fields"] if not f["duplicate"]}
    t_vals, t_labs, _ = place_table_records(context.get("tables", []), data.get("tables"))
    plain = {k: v for k, v in (data.get("pdf_fields") or {}).items() if not (by_key.get(k) or {}).get("table")}
    merged = list(t_vals.items()) + list(plain.items())
    rows = []
    for k, v in merged:
        if v is None or str(v).strip() == "": continue
        f = by_key.get(k)
        rows.append({"기입": f is not None, "칸 (Key)": k, "값": str(v).strip(), "주변 문구": t_labs.get(k, "")})
    return rows

def fill_qa_pdf_interactive(target_bytes, values_by_key, fields):
    doc = fitz.open(stream=target_bytes, filetype="pdf")
    key_by_name = {f["name"]: f["key"] for f in fields if not f["duplicate"] and f["type"] not in BUTTON_TYPES}
    filled = 0
    for f in fields:
        key = key_by_name.get(f["name"], f["key"]) if f["duplicate"] else f["key"]
        if key not in values_by_key: continue
        val = str(values_by_key[key]).strip()
        if not val: continue
        page = doc[f["page"]]
        w = page.load_widget(f["xref"])
        if w is None: continue
        try:
            if f["type"] in BUTTON_TYPES:
                if val.lower() in TRUE_WORDS:
                    w.field_value = f["on_state"] or True
                    w.update()
            else:
                w.field_flags &= ~1
                w.field_value = val
                w.text_font = "CourierNew"
                w.update()
            filled += 1
        except Exception: pass
    out_buf = io.BytesIO()
    doc.save(out_buf, deflate=True)
    doc.close()
    return out_buf.getvalue(), filled, []

def build_output_file(rows, ctx, target_bytes, base_name):
    values = {r["칸 (Key)"]: r["값"] for r in rows if r.get("기입")}
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
# 5-3. [메뉴 6] 서류 정보 일치 확인 엔진
# ==========================================
XC_DOC_TYPES_KO = {
    "questionnaire": "질문지", "resume": "이력서", "employment_letter": "재직/경력증명서",
    "job_offer": "Job Offer", "work_permit": "워크퍼밋", "study_permit": "학생비자",
    "visitor_record": "Visitor Record", "passport": "여권", "lmia": "LMIA", "other": "기타 서류",
}

XC_PROMPT = """
Extract information from attached documents into JSON:
{"documents": [{"file": "name", "doc_type": "type", "person_name": "name", "dob": "YYYY-MM-DD", "passport_no": "num", "jobs": [{"employer": "co", "job_title": "title", "start": "YYYY-MM", "end": "YYYY-MM"}]}]}
"""

def xc_run(files):
    parts, warns = qa_batch_process_client_files(files)
    if not parts: return None, warns
    resp = safe_generate_content([XC_PROMPT] + parts, json_mode=True, thinking=THINKING_LEVEL)
    data = parse_json_response(resp.text)
    return data, warns

# ==========================================
# 6. Streamlit 네비게이션 및 UI 구성
# ==========================================
MENU_1 = "🍁 IMM5476 자동 작성"
MENU_2 = "✈️ 한부모 동의서 자동 작성"
MENU_3 = "📋 IMM서류 정보 정리"
MENU_4 = "🏷️ CRM 파일명 생성 및 묶기/분할"
MENU_5 = "📝 질문지 기존 정보 기입"
MENU_6 = "🔍 서류 정보 일치 확인"

st.sidebar.title("🦅 CanNest Tool")
ALL_MENUS = [MENU_1, MENU_2, MENU_3, MENU_4, MENU_5, MENU_6]
st.session_state.setdefault("app_mode", MENU_1)
st.sidebar.caption("원하시는 업무 도구를 선택하세요")
for _m in ALL_MENUS:
    if st.sidebar.button(_m, key=f"menu_btn_{_m}", **_WIDE,
                         type="primary" if st.session_state.app_mode == _m else "secondary"):
        st.session_state.app_mode = _m
        st.rerun()
app_mode = st.session_state.app_mode

# ------------------------------------------
# 메뉴 1: IMM5476 자동 작성
# ------------------------------------------
if app_mode == MENU_1:
    st.title(MENU_1)
    if "extracted_5476" not in st.session_state: st.session_state.extracted_5476 = None

    template_5476_bytes = get_preloaded_file_bytes(["imm5476_template.pdf", "imm5476_template.pdf.pdf"])
    if template_5476_bytes:
        st.success("✅ 사내 표준 'IMM5476' 양식이 자동으로 로드되었습니다.")
    else:
        st.error("⚠️ GitHub에 'imm5476_template.pdf' 파일이 없습니다. 수동으로 업로드해 주세요.")
        template_file = st.file_uploader("IMM5476 템플릿 PDF 선택", type=["pdf"], key="template_5476")
        if template_file: template_5476_bytes = template_file.getvalue()

    st.markdown("---")
    client_file = st.file_uploader("1. 손님 여권 또는 퍼밋", type=["jpg", "jpeg", "png", "pdf", "heic", "HEIC", "docx", "DOCX", "doc", "DOC"], key="client_5476")

    if client_file and st.button("정보 추출하기", use_container_width=True):
        with st.spinner("서류 분석 중입니다. 잠시만 기다려 주세요..."):
            extracted = extract_imm5476_info(process_uploaded_file_to_image(client_file))
            if extracted:
                st.session_state.extracted_5476 = extracted
                st.success("정보 추출이 완료되었습니다.")

    if st.session_state.extracted_5476:
        data = st.session_state.extracted_5476
        c1, c2 = st.columns(2)
        with c1:
            surname = st.text_input("성 (Surname)", data.get("surname", ""))
            dob = st.text_input("생년월일", data.get("dob", ""))
            email = st.text_input("이메일 주소", data.get("email", ""))
        with c2:
            given = st.text_input("이름 (Given Name)", data.get("given_name", ""))
            uci = st.text_input("UCI", data.get("uci", ""))
            sign_date = st.date_input("서명날짜", datetime.date.today())

        address_phone = st.text_input("주소 또는 전화번호 (이메일이 없는 미성년자/신청자용)", value=data.get("address_phone", ""), placeholder="예: 2301-6658 Dow Ave, Burnaby BC V5H 0C7")

        if st.button("문서 생성 및 다운로드", type="primary"):
            if not template_5476_bytes:
                st.error("템플릿 파일이 없습니다.")
            else:
                final_data = {
                    "surname": surname, "given_name": given, "dob": dob, "uci": uci,
                    "email": email, "address_phone": address_phone, "signDate": sign_date.strftime("%Y-%m-%d"),
                }
                pdf_out = fill_imm5476(template_5476_bytes, final_data)
                safe_download_button("📥 다운로드", pdf_out, file_name=f"IMM5476_{surname}_{given}.pdf", mime="application/pdf")

        st.markdown("---")
        if st.button("🔄 전체 리셋", type="secondary", use_container_width=True, key="reset_btn_m1"):
            for key in list(st.session_state.keys()):
                if key not in ("password_correct", "app_mode"):
                    del st.session_state[key]
            st.rerun()

# ------------------------------------------
# 메뉴 2: 한부모 동의서 자동 작성
# ------------------------------------------
elif app_mode == MENU_2:
    st.title(MENU_2)
    if "consent_non_acc" not in st.session_state: st.session_state.consent_non_acc = {}
    if "consent_family" not in st.session_state: st.session_state.consent_family = []

    consent_template_bytes = get_preloaded_file_bytes(["consent_template.pdf", "consent_template.pdf.pdf"])
    if consent_template_bytes:
        st.success("✅ 사내 표준 '한부모 동의서' 양식이 자동으로 로드되었습니다.")
    else:
        st.error("⚠️ GitHub에 'consent_template.pdf' 파일이 없습니다. 수동으로 업로드해 주세요.")
        consent_template = st.file_uploader("동의서 양식 선택", type=["pdf"])
        if consent_template: consent_template_bytes = consent_template.getvalue()

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1: non_acc_file = st.file_uploader("비동반 부모님 여권 (1장)", type=["jpg", "jpeg", "png", "pdf", "heic", "HEIC", "docx", "DOCX", "doc", "DOC"])
    with c2: family_files = st.file_uploader("동반 부모/자녀 여권", type=["jpg", "jpeg", "png", "pdf", "heic", "HEIC", "docx", "DOCX", "doc", "DOC"], accept_multiple_files=True)

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
    non_acc_name = st.text_input("비동반 부모 성명", format_full_name(non_acc_data.get("surname", ""), non_acc_data.get("given_name", "")))
    non_acc_address = st.text_input("주소")
    ca, cb = st.columns(2)
    with ca: non_acc_phone = st.text_input("전화번호")
    with cb: non_acc_email = st.text_input("이메일")

    acc_parents = [p for p in st.session_state.consent_family if not is_minor(p.get("dob", ""))]
    children_list = [p for p in st.session_state.consent_family if is_minor(p.get("dob", ""))]

    acc_name, acc_passport, acc_rel = "", "", "Mother"
    if acc_parents:
        p = acc_parents[0]
        acc_name = format_full_name(p.get("surname", ""), p.get("given_name", ""))
        acc_passport = p.get("passport_number", "")
        acc_rel = "Mother" if p.get("gender") == "F" else "Father"

    cp1, cp2, cp3 = st.columns(3)
    with cp1: acc_name = st.text_input("동반 부모 성명", acc_name)
    with cp2: acc_rel = st.selectbox("관계", ["Mother", "Father"], index=0 if acc_rel == "Mother" else 1)
    with cp3: acc_passport = st.text_input("여권번호", acc_passport)

    final_children = []
    for idx, c in enumerate(children_list):
        cc1, cc2 = st.columns(2)
        with cc1: n = st.text_input(f"자녀{idx+1} 성명", format_full_name(c.get("surname", ""), c.get("given_name", "")))
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
    trip_date = st.text_input("여행 기간 (Travel Date)", placeholder="예: 2026/09/01 ~ 2026/09/30 또는 August 2026")
    sign_date_str = st.date_input("서명일", datetime.date.today()).strftime("%Y/%m/%d")

    if st.button("문서 생성 및 다운로드", type="primary"):
        if not consent_template_bytes:
            st.error("양식 파일이 없습니다.")
        else:
            data_consent = {
                "non_acc_name": non_acc_name, "non_acc_address": non_acc_address, "non_acc_phone": non_acc_phone,
                "non_acc_email": non_acc_email, "children": final_children, "acc_name": acc_name,
                "acc_relationship": acc_rel, "acc_passport": acc_passport, "trip_address": trip_address,
                "trip_phone": trip_phone, "trip_email": trip_email, "trip_date": trip_date, "sign_date": sign_date_str,
            }
            pdf_out = fill_consent_letter(consent_template_bytes, data_consent)
            crm_name = "NAME"
            if non_acc_name:
                if re.search(r"[가-힣]", non_acc_name): crm_name = non_acc_name.replace(" ", "")
                else:
                    parts = non_acc_name.strip().split()
                    if parts: crm_name = parts[0].capitalize()
            safe_download_button("📥 다운로드", pdf_out, f"{crm_name}_Consent Letter for Children Travelling Abroad.pdf", "application/pdf")

    st.markdown("---")
    if st.button("🔄 전체 리셋", type="secondary", use_container_width=True, key="reset_btn_m2"):
        for key in list(st.session_state.keys()):
            if key not in ("password_correct", "app_mode"):
                del st.session_state[key]
        st.rerun()

# ------------------------------------------
# 메뉴 3: 이민서류 정보 정리 (Case File Prep)
# ------------------------------------------
elif app_mode == MENU_3:
    st.title(MENU_3)
    if "prep_result" not in st.session_state: st.session_state.prep_result = None

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
        if tmpl_bytes: st.success(f"✅ '{selected_form}' 양식이 자동으로 로드되었습니다.")
        else: st.error(f"⚠️ {file_path} 파일이 서버에 없습니다.")
    else:
        tmpl_prep_file = st.file_uploader("빈 IMM 서식 (Print to PDF 평탄화 필수)", type=["pdf"], key="case_tmpl")
        if tmpl_prep_file: tmpl_bytes = tmpl_prep_file.getvalue()

    st.markdown("---")
    st.subheader("2. 손님 제출 서류 (복수 선택 가능)")
    client_prep_files = st.file_uploader("서류 선택", type=["jpg", "jpeg", "png", "pdf", "heic", "HEIC", "docx", "DOCX", "doc", "DOC"], accept_multiple_files=True, key="case_client_docs")

    if st.button("서류 정보 정리하기", type="primary", use_container_width=True):
        if tmpl_bytes is None: st.warning("서식이 정상적으로 선택되지 않았습니다.")
        elif not client_prep_files: st.warning("손님 서류를 1개 이상 올려주세요.")
        else:
            with st.spinner("서류를 대조하여 정리 중입니다..."):
                res = extract_case_prep_info(tmpl_bytes, client_prep_files)
                if res:
                    st.session_state.prep_result = res
                    st.success("서류 정보 정리가 완료되었습니다.")

    if st.session_state.prep_result:
        st.markdown("---")
        st.subheader("3. 정리된 정보 결과")
        parsed = st.session_state.prep_result
        sections = parsed.get("sections", [])
        if not sections: st.error("분석할 항목을 찾지 못했습니다.")
        else:
            full_text_list = []
            for sec in sections:
                sec_name = sec.get("section", "기타 항목")
                st.write(f"### 📌 {sec_name}")
                full_text_list.append(f"[{sec_name}]")
                table_data = []
                for f in sec.get("fields", []):
                    table_data.append({"항목 (Field)": f.get("field", ""), "추출값 (Value)": f.get("value", "") or "⚠️ 확인 필요", "출처 (Source)": f.get("source", "-")})
                    full_text_list.append(f"{f.get('field', '')}: {f.get('value', '')}")
                st.table(table_data)
                full_text_list.append("")
            st.text_area("복사하기", value="\n".join(full_text_list), height=250)

# ------------------------------------------
# 메뉴 4: CRM 파일명 자동 생성 및 묶기/분할 (경량화 & 타임아웃 방지)
# ------------------------------------------
elif app_mode == MENU_4:
    st.title(MENU_4)
    st.caption("개별 낱장 이미지, 여러 장짜리 통짜 PDF, MS Word 서류 등을 묶거나 분할하여 CRM 파일명으로 최적화합니다.")

    if "uploader_key" not in st.session_state: st.session_state.uploader_key = str(uuid.uuid4())
    if "analysis_results" not in st.session_state: st.session_state.analysis_results = None

    uploaded_files = st.file_uploader(
        "서류 업로드 (복수 선택 가능)",
        type=["jpg", "jpeg", "png", "pdf", "heic", "HEIC", "docx", "DOCX", "doc", "DOC"],
        accept_multiple_files=True,
        key=st.session_state.uploader_key,
    )

    if uploaded_files:
        if st.button("서류 분석 및 묶기/분할 시작", type="primary", use_container_width=True):
            results = []
            status_text = st.empty()
            progress_bar = st.progress(0)

            status_text.text("1. 전체 서류 페이지 스캔 및 초경량 썸네일 생성 중...")
            file_storage = {}
            global_pages = []
            page_counter = 1

            for file in uploaded_files:
                file_bytes = file.getvalue()
                file_storage[file.name] = file_bytes
                mime_type = file.type if file.type else "application/pdf"
                fname_lower = file.name.lower()

                if fname_lower.endswith((".doc", ".docx")):
                    text_content = read_word_document_text(file_bytes, file.name)
                    img = Image.new("RGB", (600, 800), color=(255, 255, 255))
                    draw = ImageDraw.Draw(img)
                    draw.text((30, 30), f"[Word: {file.name}]\n\n" + (text_content[:600] if text_content else ""), fill=(0, 0, 0))
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=55)
                    global_pages.append({
                        "global_idx": page_counter, "original_name": file.name, "mime_type": mime_type,
                        "pdf_page_idx": 0, "preview_bytes": buf.getvalue(), "is_word": True, "word_text": text_content,
                    })
                    page_counter += 1
                elif "pdf" in mime_type.lower() or fname_lower.endswith(".pdf"):
                    doc = fitz.open(stream=file_bytes, filetype="pdf")
                    for i in range(len(doc)):
                        page = doc.load_page(i)
                        # 💡 700px 이하로 초경량화하여 1GB 메모리 폭발 완벽 방지
                        scale = min(1.0, 700.0 / max(page.rect.width, page.rect.height))
                        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        buf = io.BytesIO()
                        img.save(buf, format="JPEG", quality=55)
                        global_pages.append({
                            "global_idx": page_counter, "original_name": file.name, "mime_type": mime_type,
                            "pdf_page_idx": i, "preview_bytes": buf.getvalue(), "is_word": False,
                        })
                        page_counter += 1
                    doc.close()
                else:
                    img = Image.open(io.BytesIO(file_bytes))
                    img = ImageOps.exif_transpose(img)
                    if img.mode != "RGB": img = img.convert("RGB")
                    if max(img.width, img.height) > 700:
                        img.thumbnail((700, 700), Image.Resampling.LANCZOS)
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=55)
                    global_pages.append({
                        "global_idx": page_counter, "original_name": file.name, "mime_type": mime_type,
                        "pdf_page_idx": 0, "preview_bytes": buf.getvalue(), "is_word": False,
                    })
                    page_counter += 1

            gc.collect()
            shrink_previews_to_fit(global_pages)

            status_text.text(f"2. AI가 {len(global_pages)}페이지의 문맥을 고속 분석 중입니다...")

            prompt = f"""
            You are an expert AI document classifier for a Canadian immigration firm.
            I am providing {len(global_pages)} pages of documents uploaded by a client. 

            Task:
            1. GROUP pages that logically belong to the SAME document.
            2. MERGE RULES:
               - Passport: Separate Old and New by expiry date. Attach stamp/visa pages to matching passport.
               - ID/Licence/PR Card: Page 1 = Front (photo/name), Page 2 = Back.
               - Digital Photo: Merge photo + studio receipt into 1 document.
               - Bank Statement/Paystub: Merge all pages of same client.
            3. ROTATION CHECK:
               - BC Driver's Licence Trap: Right-side text 'DRIVER'S LICENCE' is vertical by design. IGNORE IT. Look at face and 'Canada' / 'KIM MIN'. If face is upright -> 0.
               - If text points LEFT -> 90. If points RIGHT -> 270. If upside down -> 180. Upright -> 0.
            4. CRM FILENAME RULES:
               - Korean Name -> Korean characters without spaces (e.g. 홍길동). Foreign -> First English Given Name.
               - Date -> YYYY.MM.DD. Underscore '_' delimiter.
               - Suffixes:
                 1. Passport: {{Name}}_PP_{{ExpiryDate}}
                 2. Permit: {{Name}}_{{WP/SP/VR/Coop/PGWP/BOWP}}_{{ExpiryDate}}
                 3. Questionnaire: {{Name}}_QA_{{WP/OWP/WH/SP/VR/TRV/EE/EEBC/BCPNP}}_{{ReceivedDate}}
                 4. LOE: {{Name}}_LOE_{{CompanyInEnglish}}
                 5. Job Offer: {{Name}}_Job Offer_{{CompanyInEnglish}}
                 6. Paystub: {{Name}}_Paystub_{{CompanyInEnglish}}_{{Dates}}
                 7. Driver's Licence: {{Name}}_Driver's Licence_{{ExpiryDate}}
                 8. PR Card: {{Name}}_PR Card_{{ExpiryDate}}
                 9. Digital Photo: {{Name}}_Digital Photo (pdf if >1 pages, else jpg)
                 10. LOA: {{Name}}_LOA_{{SchoolNameInEnglish}}

            Return ONLY raw JSON:
            {{
              "page_details": [{{"page_index": 1, "rotation_needed_clockwise": 0}}],
              "documents": [{{"client_name": "...", "doc_category": "...", "suggested_filename": "...", "page_indices": [1], "is_resume": false}}]
            }}
            """

            contents = [prompt]
            for p in global_pages:
                contents.append(f"--- Page {p['global_idx']} ---")
                contents.append({"mime_type": "image/jpeg", "data": p["preview_bytes"]})

            try:
                response = safe_generate_content(contents, thinking=CRM_THINKING_LEVEL)
                clean_text = response.text.strip().replace("```json", "").replace("```", "")
                data = json.loads(clean_text)

                rotations = {str(item.get("page_index")): item.get("rotation_needed_clockwise", 0) for item in data.get("page_details", [])}
                raw_docs_info = data.get("documents", [])

                docs_info = []
                for d in raw_docs_info:
                    d["suggested_filename"] = normalize_qa_filename(d.get("suggested_filename", ""))
                    docs_info.append(d)

                for d in docs_info:
                    d["suggested_filename"] = fix_filename_by_content(d.get("suggested_filename", ""), crm_pages_text(global_pages, d.get("page_indices", []), file_storage))

            except Exception as e:
                st.error(f"AI 분석 중 오류 발생: {e}")
                docs_info = []
                rotations = {}

            status_text.text("3. 분석 완료. 최종 서류 결합 및 회전 처리 중...")
            progress_step = 1 / max(len(docs_info), 1)

            for idx, doc_info in enumerate(docs_info):
                indices = doc_info.get("page_indices", [])
                if not indices: continue
                final_name = doc_info.get("suggested_filename", f"Document_{idx+1}")
                is_resume = doc_info.get("is_resume", False) or "resume" in final_name.lower()

                group_pages = [p for p in global_pages if p["global_idx"] in indices]
                unique_src_files = list({p["original_name"] for p in group_pages})

                if is_resume and len(unique_src_files) == 1 and group_pages[0]["is_word"]:
                    comp_bytes = file_storage[group_pages[0]["original_name"]]
                    out_mime = group_pages[0]["mime_type"]
                else:
                    new_doc = fitz.open()
                    for p_data in group_pages:
                        rot = int(rotations.get(str(p_data["global_idx"]), 0))
                        src_bytes = file_storage[p_data["original_name"]]
                        if "pdf" in p_data["mime_type"].lower():
                            src_doc = fitz.open(stream=src_bytes, filetype="pdf")
                            new_doc.insert_pdf(src_doc, from_page=p_data["pdf_page_idx"], to_page=p_data["pdf_page_idx"])
                            if rot != 0:
                                new_doc[-1].set_rotation((new_doc[-1].rotation + rot) % 360)
                            src_doc.close()
                        else:
                            img = Image.open(io.BytesIO(src_bytes))
                            img = ImageOps.exif_transpose(img)
                            if img.mode != "RGB": img = img.convert("RGB")
                            if rot != 0: img = img.rotate(-rot, expand=True)
                            img_buf = io.BytesIO()
                            img.save(img_buf, format="JPEG", quality=90)
                            pdf_page = new_doc.new_page(width=img.width, height=img.height)
                            pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())

                    merged_io = io.BytesIO()
                    new_doc.save(merged_io, deflate=True)
                    new_doc.close()

                    is_jpeg = final_name.lower().endswith((".jpg", ".jpeg")) and len(indices) == 1
                    if is_jpeg:
                        comp_bytes, out_mime = process_and_compress_file(merged_io.getvalue(), "image/jpeg", final_name, already_sanitized=True)
                    else:
                        comp_bytes, out_mime = process_and_compress_file(merged_io.getvalue(), "application/pdf", final_name, already_sanitized=True)

                results.append({
                    "original_name": f"묶음 ({len(indices)}p)", "suggested_filename": final_name,
                    "mime": out_mime, "orig_kb": len(comp_bytes) / 1024, "comp_kb": len(comp_bytes) / 1024,
                    "bytes": comp_bytes, "final_name": ensure_extension(final_name, out_mime),
                })
                progress_bar.progress(min((idx + 1) * progress_step, 1.0))

            status_text.success("모든 서류의 묶기/분할 및 최적화가 완료되었습니다.")
            st.session_state.analysis_results = results
            st.session_state.crm_run_id = str(uuid.uuid4())[:8]

    if st.session_state.analysis_results:
        st.markdown("---")
        st.subheader("변환 완료된 서류 다운로드")
        run_id = st.session_state.get("crm_run_id", "0")
        for idx, item in enumerate(st.session_state.analysis_results):
            col1, col2, col3 = st.columns([3, 3, 2])
            edit_key = f"edit_{run_id}_{idx}"
            if edit_key in st.session_state:
                item["final_name"] = ensure_extension(st.session_state[edit_key], item["mime"])
            with col1:
                st.write(f"**{item['original_name']}**")
                st.caption(f"크기: **{item['comp_kb']:.1f} KB**")
            with col2:
                st.text_input("파일명", value=item["final_name"], key=edit_key, label_visibility="collapsed")
            with col3:
                safe_download_button("⬇️ 다운로드", data=item["bytes"], file_name=item["final_name"], mime=item["mime"], key=f"dl_{run_id}_{idx}")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for item in st.session_state.analysis_results:
                zip_file.writestr(item["final_name"], item["bytes"])
        zip_buffer.seek(0)
        st.markdown("---")
        safe_download_button("📦 전체 서류 ZIP 다운로드", data=zip_buffer, file_name=f"CRM_Documents_{datetime.date.today():%Y%m%d}.zip", mime="application/zip", type="primary", **_WIDE)

# ------------------------------------------
# 메뉴 5: 질문지 기존 정보 기입
# ------------------------------------------
elif app_mode == MENU_5:
    st.title(MENU_5)
    st.caption("기존 손님 서류들과 새 빈 질문지를 올리시면, 알고 있는 정보를 서식 칸에 맞춰 기입해 드립니다.")
    if "menu5_uploader_key" not in st.session_state: st.session_state.menu5_uploader_key = str(uuid.uuid4())
    for k in ("qa_rows", "qa_ctx", "qa_sig", "qa_warnings", "qa_base"): st.session_state.setdefault(k, None)

    UPLOAD_TYPES = ['pdf', 'docx', 'jpg', 'jpeg', 'png', 'heic', 'heif', 'webp']
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 주 신청자 서류")
        main_files = st.file_uploader("이전 질문지, 여권, 퍼밋 등", type=UPLOAD_TYPES, accept_multiple_files=True, key=f"qa_main_{st.session_state.menu5_uploader_key}")
    with c2:
        st.subheader("2. 가족 서류 (해당 시)")
        family_files = st.file_uploader("가족 서류", type=UPLOAD_TYPES, accept_multiple_files=True, key=f"qa_fam_{st.session_state.menu5_uploader_key}")

    st.subheader("3. 새로 채울 빈 질문지")
    new_q_file = st.file_uploader("새 질문지 (.pdf 또는 .docx)", type=['pdf', 'docx'], key=f"qa_tgt_{st.session_state.menu5_uploader_key}")

    if st.button("✨ 질문지 정보 기입 시작", type="primary", **_WIDE):
        if not main_files or not new_q_file:
            st.warning("서류를 업로드해 주세요.")
        else:
            with st.spinner("정보를 분석하여 기입 중입니다..."):
                data, warns, ctx = analyze_qa_for_fill(main_files, family_files, new_q_file)
                if data and ctx:
                    st.session_state.qa_rows = build_review_rows(data, ctx)
                    st.session_state.qa_ctx = ctx
                    st.session_state.qa_base = os.path.splitext(new_q_file.name)[0]
                    st.rerun()

    if st.session_state.qa_rows and new_q_file:
        out_bytes, out_name, out_mime, filled, _ = build_output_file(st.session_state.qa_rows, st.session_state.qa_ctx, new_q_file.getvalue(), st.session_state.qa_base)
        st.success(f"✅ {filled}개 칸에 기입 완료되었습니다.")
        safe_download_button("📥 완성된 질문지 내려받기", data=out_bytes, file_name=out_name, mime=out_mime, type="primary", **_WIDE)

# ------------------------------------------
# 메뉴 6: 서류 정보 일치 확인
# ------------------------------------------
elif app_mode == MENU_6:
    st.title(MENU_6)
    st.caption("질문지·이력서·퍼밋·여권 등을 올리면 서류 간 정보 일치 여부를 대조해 드립니다.")
    if "xc_uploader_key" not in st.session_state: st.session_state.xc_uploader_key = str(uuid.uuid4())
    st.session_state.setdefault("xc_result", None)

    xc_files = st.file_uploader("서류 선택 (2개 이상)", type=['pdf', 'docx', 'jpg', 'jpeg', 'png', 'heic', 'webp'], accept_multiple_files=True, key=f"xc_{st.session_state.xc_uploader_key}")

    if st.button("🔍 서류 정보 일치 확인 시작", type="primary", **_WIDE):
        if not xc_files or len(xc_files) < 2:
            st.warning("비교할 서류를 2개 이상 올려주세요.")
        else:
            with st.spinner("서류 정보를 대조 중입니다..."):
                data, warns = xc_run(xc_files)
                if data:
                    st.session_state.xc_result = data
                    st.rerun()

    if st.session_state.xc_result:
        st.success("대조가 완료되었습니다.")
        st.json(st.session_state.xc_result)
