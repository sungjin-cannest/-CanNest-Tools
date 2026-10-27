import streamlit as st

# 📌 페이지 설정
st.set_page_config(page_title="CanNest 통합 업무 시스템", layout="wide")

import google.generativeai as genai
try:
    import pymupdf as fitz  # PyMuPDF 최신 버전
except ImportError:
    import fitz  # 구버전 호환
from PIL import Image, ImageOps, ImageDraw
import pillow_heif
import io
import json
import os
import re
import time
import datetime
import hashlib
import uuid
import zipfile
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor

try:
    import docx
    from docx.shared import RGBColor
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

pillow_heif.register_heif_opener()
Image.MAX_IMAGE_PIXELS = None

LOCAL_TZ = ZoneInfo("America/Vancouver")

# Streamlit 버전에 따라 '가로 꽉 채우기' 옵션 이름이 다름 (신버전: width="stretch")
import inspect as _inspect
try:
    _WIDE = {"width": "stretch"} if "width" in _inspect.signature(st.download_button).parameters else {"use_container_width": True}
except Exception:
    _WIDE = {"use_container_width": True}
HANGUL_RE = re.compile(r"[ㄱ-ㆎ가-힣]")
IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.heic', '.heif', '.webp')

# ==========================================
# 0. API 키 설정 및 보안
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

GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)

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
        for m in genai.list_models():
            name = m.name.split("/")[-1]
            methods = getattr(m, "supported_generation_methods", []) or []
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

def safe_generate_content(contents, json_mode=False):
    """모델 폴백 + 일시적 오류(429/500/503) 재시도."""
    gen_config = {"temperature": 0}
    if json_mode:
        gen_config["response_mime_type"] = "application/json"

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
        mod = genai.GenerativeModel(model_name, generation_config=gen_config)
        for attempt in range(3):
            try:
                resp = mod.generate_content(_clean_contents(contents))
                st.session_state["gemini_working_model"] = model_name
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
                if _is_model_gone(msg):
                    break  # 다음 모델로
                if any(k in msg for k in ["429", "500", "503", "unavailable", "overloaded", "deadline", "resource exhausted"]):
                    time.sleep(2 * (attempt + 1))
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
# 1. 문서 파싱 엔진
# ==========================================
def read_word_document_text(file_bytes, file_name=""):
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

    if fname.endswith(('.doc', '.docx')):
        text = read_word_document_text(file_bytes, file_obj.name)
        img = Image.new('RGB', (800, 1000), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        display_str = f"[Word Document: {file_obj.name}]\n\n" + (text[:800] if text else "Word Document Content")
        draw.text((40, 40), display_str, fill=(0, 0, 0))
        return img
    elif file_obj.type == "application/pdf" or fname.endswith('.pdf'):
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
        img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=75)
    buf.seek(0)
    return Image.open(buf)

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
    try: widget.text_font = "Cour"
    except Exception:
        try: widget.text_font = "Courier"
        except Exception: pass
    widget.update()

def embed_courier_in_doc(doc):
    if not os.path.exists("cour.ttf"): return
    try:
        page = doc[0]
        font_xrefs = {}
        for alias in ["Cour", "Courier", "CoNi", "CourierNew", "CourierNewPSMT"]:
            try:
                xref = page.insert_font(fontfile="cour.ttf", fontname=alias)
                font_xrefs[alias] = xref
            except Exception: pass
        catalog_xref = doc.pdf_catalog()
        acroform_res = doc.xref_get_key(catalog_xref, "AcroForm")
        acroform_xref = int(acroform_res[1].split()[0]) if acroform_res[0] == "xref" else catalog_xref
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

# 질문지 안의 '구역 제목' 패턴 (신청자/배우자/자녀 등). 다른 질문지도 대부분 이 표현을 씀
SECTION_HEADER_RE = re.compile(
    r"(개인\s*정보|가족\s*관계\s*정보|personal\s+(information|details)|family\s+(information|members))",
    re.IGNORECASE)

def detect_sections(doc):
    """각 페이지에서 구역 제목(예: '배우자의 개인 정보')의 위치를 찾아 (page, y, 제목) 목록 반환."""
    headers = []
    for p_idx, page in enumerate(doc):
        for b in page.get_text("blocks"):
            first_line = (b[4] or "").strip().split("\n")[0].strip()
            if first_line and len(first_line) <= 60 and SECTION_HEADER_RE.search(first_line):
                title = re.split(r"\s+-\s+|\(\*", first_line)[0].strip()
                headers.append((p_idx, b[1], title))
    headers.sort(key=lambda h: (h[0], h[1]))
    return headers

def person_for_section(title):
    # 제목의 '주어' 부분만 판단 (예: '신청자 가족관계 정보 (부모님, 배우자, 자녀…)'의 괄호 속 단어는 무시)
    t = (title or "").lower()
    m0 = SECTION_HEADER_RE.search(t)
    if m0:
        t = t[:m0.start()]
    m = re.search(r"자녀\s*\(?(\d+)\)?|child\s*\(?(\d+)\)?", t)
    if m:
        return f"CHILD {m.group(1) or m.group(2)}"
    if "자녀" in t or "child" in t or "dependant" in t or "dependent" in t:
        return "CHILD"
    if "배우자" in t or "spouse" in t or "partner" in t:
        return "SPOUSE"
    return "MAIN APPLICANT"

def section_for(headers, page_idx, y):
    current = ""
    for hp, hy, title in headers:
        if hp < page_idx or (hp == page_idx and hy <= y + 2):
            current = title
        else:
            break
    return current

# ------------------------------------------
# 1-1a. '현재 정보' 칸 (손님이 직접 최신으로 작성) → 항상 비워둠
# ------------------------------------------
CURRENT_ADDRESS_RE = re.compile(
    r"현재\s*거주지|current\s+(residential\s+|home\s+)?address|residential\s+address|present\s+address",
    re.IGNORECASE)
RECENT_ENTRY_PHRASES = ["가장 최근 입국", "최근 입국", "most recent entry", "last entry", "latest entry", "most recent arrival"]
FIRST_ENTRY_PHRASES = ["최초 입국", "처음 입국", "first entry", "original entry", "first arrival"]

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
            for kind, plist in (("recent", RECENT_ENTRY_PHRASES), ("first", FIRST_ENTRY_PHRASES)):
                seen_rects = []
                for ph in plist:
                    for r in page.search_for(ph):
                        if any(r.intersects(x) for x in seen_rects):
                            continue  # '가장 최근 입국' 안의 '최근 입국' 중복 방지
                        seen_rects.append(r)
                        found.append((kind, r))
            phrase_cache[p_idx] = found
        return phrase_cache[p_idx]

    for f in fields:
        if f["type"] != fitz.PDF_WIDGET_TYPE_TEXT:
            continue
        page = doc[f["page"]]
        x0, y0, x1, y1 = f["rect"]
        left = " ".join(page.get_text("text", clip=fitz.Rect(max(0, x0 - 220), y0 - 1, x0, y1 + 1)).split())
        if CURRENT_ADDRESS_RE.search(left) or CURRENT_ADDRESS_RE.search(f.get("label") or ""):
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
                "person": person_for_section(section_for(headers, p_idx, r.y0)),
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
def prepare_document_for_gemini(file_bytes, mime_type, file_name="", max_chars=35000):
    ext = os.path.splitext(file_name)[1].lower() if file_name else ""
    mime_type = (mime_type or "").lower()
    header = f"\n--- [Document: {file_name}] ---\n"

    if ext == ".doc" or mime_type == "application/msword":
        # [수정] 구형 .doc는 python-docx로 못 읽고, Gemini도 받지 못해 전체 분석이 실패했음
        return [], f"{file_name}: 구형 .doc 형식은 읽을 수 없습니다. Word에서 .docx 또는 PDF로 저장 후 올려주세요."

    if "word" in mime_type or ext == ".docx":
        word_text = read_word_document_text(file_bytes, file_name)
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

def batch_process_client_files(files, owner=None):
    """owner: 'MAIN APPLICANT' 또는 'FAMILY MEMBER' — 업로드 칸 기준으로 서류 주인을 표시."""
    def worker(f):
        label = f"{owner} | {f.name}" if owner else f.name
        return prepare_document_for_gemini(f.getvalue(), f.type, label)
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

    source_parts, warnings = batch_process_client_files(main_files or [], "MAIN APPLICANT")
    if family_files:
        fam_parts, fam_warn = batch_process_client_files(family_files, "FAMILY MEMBER")
        source_parts += fam_parts
        warnings += fam_warn
    if not source_parts:
        st.error("읽을 수 있는 손님 자료가 없습니다.")
        return None, warnings, None

    if is_docx:
        form_text = read_word_document_text(target_bytes, target_file.name)
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
        response = safe_generate_content([prompt] + source_parts, json_mode=True)
        data = parse_json_response(response.text)
        if not isinstance(data, dict):
            raise ValueError("AI 응답 형식이 올바르지 않습니다.")
        return data, warnings, context
    except Exception as e:
        st.error(f"서류 분석 중 오류가 발생했습니다: {e}")
        return None, warnings, None

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
            if f is None:
                note = "⚠️ 서식에 없는 칸 (기입 안 됨)"
            elif HANGUL_RE.search(v):
                note = "⚠️ 한글 포함"
            elif f["choices"] and v.lower() not in [c.lower() for c in f["choices"]]:
                note = "⚠️ 선택지에 없는 값"
            rows.append({"기입": f is not None, "구역": ({"MAIN APPLICANT": "신청자", "SPOUSE": "배우자"}.get(f.get("person"), (f.get("person") or "").replace("CHILD", "자녀")) if f else ""),
                         "칸 (Key)": k,
                         "주변 문구": (t_labels.get(k) or (f["nearby"] if f else ""))[:80],
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
MENU_4 = "🏷️ CRM 파일명 생성 및 묶기"
MENU_5 = "📝 질문지 기존 정보 기입"

st.sidebar.title("🦅 CanNest Tool")
app_mode = st.sidebar.radio("원하시는 업무 도구를 선택하세요", [MENU_1, MENU_2, MENU_3, MENU_4, MENU_5])

if app_mode in [MENU_1, MENU_2, MENU_3, MENU_4]:
    st.title(app_mode)
    st.info("좌측 메뉴에서 5번 '질문지 기존 정보 기입' 기능을 테스트해 보세요.")

# ------------------------------------------
# 메뉴 5: 질문지 기존 정보 기입
# ------------------------------------------
elif app_mode == MENU_5:
    st.title(MENU_5)
    st.caption("기존 손님 서류들과 새 빈 질문지를 올리시면, 알고 있는 정보를 서식 칸에 맞춰 기입해 드립니다. "
               "결과는 다운로드 전에 표에서 확인·수정할 수 있습니다.")

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

    # 💡 결과 검토 및 다운로드
    if st.session_state.qa_rows is not None and new_q_file is not None:
        rows = st.session_state.qa_rows
        if not rows:
            st.info("자료에서 질문지에 기입할 수 있는 정보를 찾지 못했습니다.")
        else:
            st.subheader("4. 기입 내용 확인 및 수정")
            st.caption("값을 직접 고치거나, '기입' 체크를 해제하면 그 칸은 비워둡니다. 수정 내용은 다운로드 파일에 바로 반영됩니다.")
            for tw in (st.session_state.qa_ctx or {}).get("table_warnings", []) or []:
                st.warning(tw)
            need_check = sum(1 for r in rows if r["확인 필요"])
            if need_check:
                st.warning(f"⚠️ 확인이 필요한 항목이 {need_check}개 있습니다. '확인 필요' 열을 살펴봐 주세요.")

            edited = st.data_editor(
                rows,
                key=f"qa_editor_{st.session_state.qa_sig}",
                **_WIDE,
                hide_index=True,
                disabled=["구역", "칸 (Key)", "주변 문구", "확인 필요"],
                column_config={
                    "기입": st.column_config.CheckboxColumn(width="small"),
                    "값": st.column_config.TextColumn(width="large"),
                },
            )
            edited_rows = edited.to_dict("records") if hasattr(edited, "to_dict") else edited

            try:
                out_bytes, out_name, out_mime, filled, skipped = build_output_file(
                    edited_rows, st.session_state.qa_ctx, new_q_file.getvalue(), st.session_state.qa_base)
                st.success(f"✅ {filled}개 칸에 기입했습니다. 확인하지 못한 항목은 비워두었으니 손님께 마저 작성해 달라고 요청하세요.")
                if skipped:
                    with st.expander(f"서식에서 위치를 찾지 못해 기입되지 않은 항목 {len(skipped)}개"):
                        st.write(skipped)
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
