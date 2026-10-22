import streamlit as st

# 📌 페이지 설정
st.set_page_config(page_title="CanNest 통합 업무 시스템", layout="wide")

import google.generativeai as genai
import fitz  # PyMuPDF
from PIL import Image, ImageOps, ImageDraw
import pillow_heif
import io
import json
import os
import re
import datetime
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor

try:
    import docx
    from docx.shared import RGBColor
except ImportError:
    pass

pillow_heif.register_heif_opener()
Image.MAX_IMAGE_PIXELS = None

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

def safe_generate_content(contents):
    candidate_models = ['gemini-3.6-flash', 'gemini-1.5-flash']
    last_error = None
    for model_name in candidate_models:
        try:
            mod = genai.GenerativeModel(model_name)
            return mod.generate_content(contents)
        except Exception as e:
            last_error = e
            if "404" in str(e) or "not found" in str(e).lower():
                continue
            raise e
    raise last_error

# ==========================================
# 1. 문서 파싱 및 공통 함수
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
        return "\n".join(text_list)
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
        acroform_xref = int(acroform_res[1].split()[0]) if acroform_res[0] == "xr" else catalog_xref
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
    if "word" in mime_type.lower() or ext in ['.doc', '.docx']:
        word_text = read_word_document_text(file_bytes, file_name)
        if word_text:
            return [f"\n--- [Word Document: {file_name}] ---\n{word_text[:20000]}\n"]

    if "pdf" in mime_type.lower() or ext == '.pdf':
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            field_list = []
            
            for page in doc:
                text += page.get_text("text") + "\n"
                for w in page.widgets():
                    if w.field_value:
                        text += f"[{w.field_name}]: {w.field_value}\n"
                    # 위젯의 정확한 이름과 근처 텍스트를 AI에게 제공
                    r = w.rect
                    clip = fitz.Rect(max(0, r.x0 - 100), max(0, r.y0 - 25), min(page.rect.width, r.x1 + 50), min(page.rect.height, r.y1 + 25))
                    nearby = page.get_text("text", clip=clip).replace('\n', ' ').strip()
                    field_list.append(f"Field Name: '{w.field_name}' | Nearby Text: '{nearby}' | Type: {w.field_type_string}")
            
            if field_list:
                text += "\n\n--- [FILLABLE PDF FIELDS AVAILABLE] ---\n"
                text += "Map client data exactly to the 'Field Name' keys listed below:\n"
                text += "\n".join(field_list) + "\n"

            if len(text.strip()) > 100:
                doc.close()
                return [f"\n--- [Document: {file_name}] ---\n{text[:30000]}\n"]
            else:
                images = []
                for p_idx in range(min(len(doc), 10)):
                    page = doc.load_page(p_idx)
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=70)
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
        return (today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))) < 19
    except Exception:
        return True

# ==========================================
# 메뉴 1~4 기존 로직 생략 (공간 절약을 위해 원본 유지됨)
# ==========================================
def extract_imm5476_info(image):
    prompt = "Analyze this identity document carefully. Extract into exact JSON structure: surname, given_name, dob, uci."
    try:
        response = safe_generate_content([prompt, image])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception: return None

def extract_all_passports_batch(has_non_acc, images):
    prompt = "Extract passport data into JSON. surname, given_name, dob, passport_number, gender."
    try:
        response = safe_generate_content([prompt] + images)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception: return None

def extract_case_prep_info(tmpl_bytes, client_files):
    prompt = "Extract case prep info matching IMM form fields. Return JSON."
    contents = [prompt]
    contents.extend(prepare_document_for_gemini(tmpl_bytes, "application/pdf", "Blank_IMM_Form.pdf"))
    contents.extend(batch_process_client_files(client_files))
    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception: return None

def fill_imm5476(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    flags = {"surname": False, "given": False, "dob": False, "email": False, "address_phone": False, "uci": False}
    for page_idx, page in enumerate(doc):
        date_counter = 0
        for widget in page.widgets():
            fname = widget.field_name.lower() if widget.field_name else ""
            val = None
            if "family name" in fname and not flags["surname"]: val = data.get("surname", ""); flags["surname"] = True
            elif "given name" in fname and not flags["given"]: val = data.get("given_name", ""); flags["given"] = True
            elif "date of birth" in fname and not flags["dob"]: val = data.get("dob", ""); flags["dob"] = True
            elif "email" in fname and not flags["email"]: val = data.get("email", ""); flags["email"] = True
            elif ("telephone" in fname or "address" in fname) and page_idx == 0 and not flags["address_phone"]: val = data.get("address_phone", ""); flags["address_phone"] = True
            elif ("uci" in fname or "unique client identifier" in fname) and not flags["uci"]: val = data.get("uci", "").replace("-", ""); flags["uci"] = True
            elif "date" in fname and "birth" not in fname:
                date_counter += 1
                if page_idx in [2, 3] and date_counter == 1: val = data.get("signDate", "")
            if val is not None:
                set_smart_widget_value(widget, val, default_fontsize=9)
    embed_courier_in_doc(doc)
    doc.need_appearances(True)
    out_buf = io.BytesIO()
    doc.save(out_buf, deflate=True)
    doc.close()
    return out_buf.getvalue()

def fill_consent_letter(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    children = data.get("children", [])
    num_pages = max(1, (len(children) + 2) // 3)
    for _ in range(num_pages - 1): doc.insert_pdf(doc, from_page=0, to_page=0)
    for page_num in range(num_pages):
        page = doc[page_num]
        p_children = children[page_num * 3 : (page_num + 1) * 3]
        child_widgets = [("Information about travelling children", "yyyymmdd"), ("1_2", "2_2"), ("1_3", "2_3")]
        for widget in page.widgets():
            fname = widget.field_name.strip() if widget.field_name else ""
            if not fname: continue
            fl = fname.lower()
            if "check" in getattr(widget, "field_type_string", "").lower() or "radio" in fl or "alone" in fl:
                try: widget.field_value = "Off"; widget.update()
                except Exception: pass
                continue
            v = None
            if fname == "1": v = data.get("non_acc_name", "")
            elif fname == "2": v = data.get("non_acc_address", "")
            elif fname == "3": v = data.get("non_acc_phone", "")
            elif fname == "email": v = data.get("non_acc_email", "")
            elif fname == "This child or these children hashave my or our consent to travel with": v = data.get("acc_name", "")
            elif fname == "Relationship with Children 1": v = data.get("acc_relationship", "")
            elif fname == "Relationship with Children 2": v = data.get("acc_passport", "")
            elif fname == "I give my consent for this child to travel to": v = "Canada"
            elif fname == "2_4" or fl == "to stay with": v = data.get("acc_name", "")
            elif fname == "At the following addresses 1": v = data.get("trip_address", "")
            elif fname == "At the following addresses 2": v = data.get("trip_phone", "")
            elif fname == "email_2": v = data.get("trip_email", "")
            elif fname == "yyyymmdd_2": v = data.get("sign_date", "")
            elif fname == "1_4" or "travel date" in fl or fl == "date" or "from" in fl or "start date" in fl: v = data.get("trip_date", "")
            elif fl == "to" or "return" in fl or "end date" in fl: v = ""
            else:
                for idx, (n_k, d_k) in enumerate(child_widgets):
                    if idx < len(p_children):
                        if fname == n_k: v = p_children[idx].get("name", "")
                        elif fname == d_k: v = p_children[idx].get("dob", "")
            if v is not None: set_smart_widget_value(widget, v, default_fontsize=11)
    embed_courier_in_doc(doc)
    doc.need_appearances(True)
    out_buf = io.BytesIO()
    doc.save(out_buf, deflate=True)
    doc.close()
    return out_buf.getvalue()

def process_and_compress_file(file_bytes, mime_type, target_filename, already_sanitized=False):
    if target_filename.lower().endswith(('.jpg', '.jpeg')):
        img = Image.open(io.BytesIO(file_bytes))
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB": img = img.convert("RGB")
        if max(img.width, img.height) > 2000:
            r = 2000.0 / float(max(img.width, img.height))
            img = img.resize((int(img.width * r), int(img.height * r)), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=80, optimize=True)
        return buf.getvalue(), "image/jpeg"
    else:
        if "pdf" in mime_type.lower():
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            total_text_len = sum([len(doc.load_page(i).get_text("text").strip()) for i in range(min(len(doc), 5))])
            has_widgets = any(list(page.widgets()) for page in doc)
            if total_text_len > 50 or has_widgets:
                doc.close()
                return (file_bytes, "application/pdf") if already_sanitized else (sanitize_and_unlock_pdf(file_bytes), "application/pdf")
            
            new_doc = fitz.open()
            for page in doc:
                rot = page.rotation % 360
                zoom = 150 / 72.0 if max(page.rect.width, page.rect.height) <= 2000 else 1.0
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img_buf = io.BytesIO()
                img.save(img_buf, format="JPEG", quality=65, optimize=True)
                new_page = new_doc.new_page(width=pix.width/zoom, height=pix.height/zoom)
                new_page.insert_image(new_page.rect, stream=img_buf.getvalue())
            out_buf = io.BytesIO()
            new_doc.save(out_buf, deflate=True, garbage=4)
            new_doc.close(); doc.close()
            final_bytes = out_buf.getvalue() if len(out_buf.getvalue()) < len(file_bytes) else file_bytes
            return (final_bytes, "application/pdf") if already_sanitized else (sanitize_and_unlock_pdf(final_bytes), "application/pdf")
        else:
            img = Image.open(io.BytesIO(file_bytes))
            img = ImageOps.exif_transpose(img)
            if img.mode != "RGB": img = img.convert("RGB")
            if max(img.width, img.height) > 1800:
                r = 1800.0 / float(max(img.width, img.height))
                img = img.resize((int(img.width * r), int(img.height * r)), Image.Resampling.LANCZOS)
            img_buf = io.BytesIO()
            img.save(img_buf, format="JPEG", quality=70, optimize=True)
            new_doc = fitz.open()
            pdf_page = new_doc.new_page(width=img.width * 72.0/150, height=img.height * 72.0/150)
            pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())
            out_buf = io.BytesIO()
            new_doc.save(out_buf)
            new_doc.close()
            return (out_buf.getvalue(), "application/pdf") if already_sanitized else (sanitize_and_unlock_pdf(out_buf.getvalue()), "application/pdf")

# ==========================================
# 5. 질문지 자동 이어채우기 (메뉴 5 전용)
# ==========================================
def analyze_qa_for_fill(client_files, target_file):
    prompt = """
    You are an expert Canadian immigration assistant filling out a blank questionnaire.
    The FIRST document is the NEW BLANK QUESTIONNAIRE. The rest are CLIENT SOURCE MATERIALS.

    [CRITICAL RULES - MUST OBEY]
    1. UNKNOWN INFO = DO NOT TOUCH: If the client documents do NOT explicitly provide the answer (e.g., no mention of Tuberculosis, Medical issues, Crime, Visa Refusals, or Biometrics), you MUST leave the field completely UNMAPPED. Do NOT output the field in the JSON. Do NOT guess "No".
    2. ENGLISH TESTS: If the client uploaded an IELTS, CELPIP, or Duolingo score, you MUST find the "YES" radio button/checkbox for the English test question and map it to "Yes".
    3. NO ROW SKIPPING IN TABLES: The PDF has sequentially named fields for tables. ALWAYS map the most recent history to the FIRST row, the next to the SECOND row, etc. DO NOT skip the first row even if there is no "Current" job. There must be NO empty gaps between rows.
    4. TRAVEL HISTORY: Ensure you extract and map ALL travel history if available in the source materials to the Travel Table fields.
    5. LANGUAGE: All output values MUST be in ENGLISH (Latin alphabet, numbers). NO KOREAN inside values.

    Return ONLY a raw JSON object formatted strictly as:
    {
      "pdf_fields": {
         "Exact_Field_Name_1": "Extracted Value",
         "Exact_Checkbox_Name": "Yes"
      },
      "docx_fields": {
         "출생 도시": "Value if the target was DOCX"
      }
    }
    """
    contents = [prompt]
    mime_target = target_file.type if target_file.type else "application/pdf"
    contents.extend(prepare_document_for_gemini(target_file.getvalue(), mime_target, f"[NEW_FORM] {target_file.name}"))
    contents.extend(batch_process_client_files(client_files))

    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"서류 분석 중 오류가 발생했습니다: {e}")
        return None

def fill_qa_pdf_interactive(target_bytes, qa_data):
    """
    원본 PDF의 폼 필드(Widget)를 그대로 활용하여 
    위치 밀림(도장 현상) 없이 깨끗하게 텍스트를 기입합니다.
    """
    doc = fitz.open(stream=target_bytes, filetype="pdf")
    pdf_fields = qa_data.get("pdf_fields", {})
    fontfile = "cour.ttf" if os.path.exists("cour.ttf") else None
    
    has_widgets = any(list(page.widgets()) for page in doc)
    
    if has_widgets:
        for page in doc:
            for w in page.widgets():
                fn = w.field_name
                if fn in pdf_fields:
                    val = pdf_fields[fn]
                    if val:
                        ftype = getattr(w, "field_type_string", "").lower()
                        
                        # 라디오/체크박스인 경우
                        if "check" in ftype or "radio" in ftype:
                            if str(val).lower() in ['yes', 'true', 'v', 'on', '1', 'checked']:
                                try:
                                    # PDF 내부에 정의된 '선택' 상태값 확인 (보통 'Yes', '1', 'On' 등)
                                    on_states = [s for s in w.on_state() if s.lower() != 'off']
                                    if on_states:
                                        w.field_value = on_states[0]
                                    else:
                                        w.field_value = True
                                    w.update()
                                except Exception:
                                    pass
                        # 텍스트 필드인 경우
                        else:
                            # 읽기 전용 해제하여 수정 가능하도록 보장
                            if hasattr(w, "field_flags"):
                                w.field_flags &= ~1 
                            w.field_value = str(val)
                            try:
                                w.text_font = "Cour" if fontfile else "Courier"
                            except Exception:
                                pass
                            w.update()
                            
    doc.need_appearances(True)
    out_buf = io.BytesIO()
    doc.save(out_buf, deflate=True)
    doc.close()
    return out_buf.getvalue()

def fill_qa_docx(target_bytes, qa_data):
    doc = docx.Document(io.BytesIO(target_bytes))
    BLUE_RGB = RGBColor(0x1F, 0x4E, 0x8C)
    fill_map = {k.strip().lower(): str(v).strip() for k, v in qa_data.get("docx_fields", {}).items() if v}

    def insert_run(p, text):
        run = p.add_run(f" {text}")
        run.font.color.rgb = BLUE_RGB
        run.bold = True

    for table in doc.tables:
        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                c_text = cell.text.strip().lower()
                for q_label, ans in fill_map.items():
                    if q_label in c_text and len(c_text) < len(q_label) + 20:
                        if c_idx + 1 < len(row.cells) and not row.cells[c_idx + 1].text.strip():
                            p = row.cells[c_idx + 1].paragraphs[0] if row.cells[c_idx + 1].paragraphs else row.cells[c_idx + 1].add_paragraph()
                            insert_run(p, ans)
                        elif "___" in cell.text or cell.text.endswith(":"):
                            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                            insert_run(p, ans)

    for p in doc.paragraphs:
        p_text = p.text.strip().lower()
        for q_label, ans in fill_map.items():
            if q_label in p_text and ("___" in p.text or p.text.strip().endswith(":")):
                insert_run(p, ans)

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()

# ==========================================
# 6. Streamlit 네비게이션 및 UI 구성
# ==========================================
MENU_1 = "🍁 IMM5476 자동 작성"
MENU_2 = "✈️ 한부모 동의서 자동 작성"
MENU_3 = "📋 IMM서류 정보 정리"
MENU_4 = "🏷️ CRM 파일명 생성 및 묶기"
MENU_5 = "📝 질문지 자동 완성"

st.sidebar.title("🦅 CanNest Tool")
app_mode = st.sidebar.radio("원하시는 업무 도구를 선택하세요", [MENU_1, MENU_2, MENU_3, MENU_4, MENU_5])

# 메뉴 1~4 생략 (기존과 동일하게 작동하므로 변경사항 없음)
if app_mode == MENU_1:
    st.title(MENU_1)
    if "extracted_5476" not in st.session_state: st.session_state.extracted_5476 = None
    template_5476_bytes = get_preloaded_file_bytes(["imm5476_template.pdf", "imm5476_template.pdf.pdf"])
    if template_5476_bytes: st.success("✅ 사내 표준 'IMM5476' 양식이 자동으로 로드되었습니다.")
    else:
        st.error("⚠️ GitHub에 'imm5476_template.pdf' 파일이 없습니다. 수동으로 업로드해 주세요.")
        template_file = st.file_uploader("IMM5476 템플릿 PDF 선택", type=['pdf'], key="template_5476")
        if template_file: template_5476_bytes = template_file.getvalue()
    st.markdown("---")
    client_file = st.file_uploader("1. 손님 여권 또는 퍼밋", type=['jpg', 'jpeg', 'png', 'pdf', 'heic', 'HEIC', 'docx', 'doc'], key="client_5476")
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
            email = st.text_input("이메일 주소", data.get("email", ""))
        with c2:
            given = st.text_input("이름 (Given Name)", data.get("given_name", ""))
            uci = st.text_input("UCI", data.get("uci", ""))
            sign_date = st.date_input("서명날짜", datetime.date.today())
        address_phone = st.text_input("주소 또는 전화번호", value=data.get("address_phone", ""), placeholder="예: 2301-6658 Dow Ave, Burnaby BC V5H 0C7")
        if st.button("문서 생성 및 다운로드", type="primary"):
            if not template_5476_bytes: st.error("템플릿 파일이 없습니다.")
            else:
                final_data = {"surname": surname, "given_name": given, "dob": dob, "uci": uci, "email": email, "address_phone": address_phone, "signDate": sign_date.strftime("%Y-%m-%d")}
                pdf_out = fill_imm5476(template_5476_bytes, final_data)
                st.download_button("📥 다운로드", pdf_out, file_name=f"IMM5476_{surname}_{given}.pdf", mime="application/pdf")
        st.markdown("---")
        if st.button("🔄 전체 리셋", type="secondary", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key != "password_correct": del st.session_state[key]
            st.rerun()

elif app_mode == MENU_2:
    st.title(MENU_2)
    if "consent_non_acc" not in st.session_state: st.session_state.consent_non_acc = {}
    if "consent_family" not in st.session_state: st.session_state.consent_family = []
    consent_template_bytes = get_preloaded_file_bytes(["consent_template.pdf", "consent_template.pdf.pdf"])
    if consent_template_bytes: st.success("✅ 사내 표준 '한부모 동의서' 양식이 자동으로 로드되었습니다.")
    else:
        st.error("⚠️ GitHub에 'consent_template.pdf' 파일이 없습니다. 수동으로 업로드해 주세요.")
        consent_template = st.file_uploader("동의서 양식 선택", type=['pdf'])
        if consent_template: consent_template_bytes = consent_template.getvalue()
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1: non_acc_file = st.file_uploader("비동반 부모님 여권 (1장)", type=['jpg', 'jpeg', 'png', 'pdf', 'heic', 'HEIC', 'docx', 'doc'])
    with c2: family_files = st.file_uploader("동반 부모/자녀 여권", type=['jpg', 'jpeg', 'png', 'pdf', 'heic', 'HEIC', 'docx', 'doc'], accept_multiple_files=True)
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
    trip_date = st.text_input("여행 기간 (Travel Date)", placeholder="예: 2026/09/01 ~ 2026/09/30 또는 August 2026")
    sign_date_str = st.date_input("서명일", datetime.date.today()).strftime("%Y/%m/%d")
    if st.button("문서 생성 및 다운로드", type="primary"):
        if not consent_template_bytes: st.error("양식 파일이 없습니다.")
        else:
            data_consent = {
                "non_acc_name": non_acc_name, "non_acc_address": non_acc_address, "non_acc_phone": non_acc_phone, "non_acc_email": non_acc_email,
                "children": final_children, "acc_name": acc_name, "acc_relationship": acc_rel, "acc_passport": acc_passport,
                "trip_address": trip_address, "trip_phone": trip_phone, "trip_email": trip_email, 
                "trip_date": trip_date, "sign_date": sign_date_str
            }
            pdf_out = fill_consent_letter(consent_template_bytes, data_consent)
            crm_name = "NAME"
            if non_acc_name:
                if re.search(r'[가-힣]', non_acc_name): crm_name = non_acc_name.replace(" ", "")
                else:
                    parts = non_acc_name.strip().split()
                    if parts: crm_name = parts[0].capitalize()
            download_file_name = f"{crm_name}_Consent Letter for Children Travelling Abroad.pdf"
            st.download_button("📥 다운로드", pdf_out, download_file_name, "application/pdf")
    st.markdown("---")
    if st.button("🔄 전체 리셋", type="secondary", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key != "password_correct": del st.session_state[key]
        st.rerun()

elif app_mode == MENU_3:
    st.title(MENU_3)
    if "prep_result" not in st.session_state: st.session_state.prep_result = None
    st.subheader("1. 대상 서식 (IMM PDF)")
    form_map = {
        "직접 파일 업로드 (기타 서식)": None, "IMM1294 (SP-OUTSIDE)": "imm1294.pdf", "IMM1295 (WP-OUTSIDE)": "imm1295.pdf",
        "IMM5708 (VR-INSIDE)": "imm5708.pdf", "IMM5709 (SP-INSIDE)": "imm5709.pdf", "IMM5710 (WP-INSIDE)": "imm5710.pdf"
    }
    selected_form = st.selectbox("📌 템플릿 서식 선택", list(form_map.keys()))
    tmpl_bytes = None
    if form_map[selected_form] is not None:
        file_path = form_map[selected_form]
        tmpl_bytes = load_pdf_bytes_cached(file_path)
        if tmpl_bytes: st.success(f"✅ '{selected_form}' 양식이 자동으로 로드되었습니다.")
        else: st.error(f"⚠️ {file_path} 파일이 서버에 없습니다. 파일 업로드 상태를 확인해 주세요.")
    else:
        tmpl_prep_file = st.file_uploader("빈 IMM 서식 (반드시 Print to PDF로 평탄화된 파일)", type=['pdf'], key="case_tmpl")
        if tmpl_prep_file: tmpl_bytes = tmpl_prep_file.getvalue()
    st.markdown("---")
    st.subheader("2. 손님 제출 서류 (복수 선택 가능)")
    client_prep_files = st.file_uploader("질문지, 여권, 퍼밋 등 서류 선택", type=['jpg', 'jpeg', 'png', 'pdf', 'heic', 'HEIC', 'docx', 'DOCX', 'doc', 'DOC'], accept_multiple_files=True, key="case_client_docs")
    if st.button("서류 정보 정리하기", type="primary", use_container_width=True):
        if tmpl_bytes is None: st.warning("1번 단계에서 서식이 정상적으로 선택되거나 업로드되지 않았습니다.")
        elif not client_prep_files: st.warning("2번 단계에서 손님 서류를 1개 이상 올려주세요.")
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
        if not sections: st.error("서식에서 분석할 항목을 찾지 못했습니다. (파일이 평탄화된 PDF인지 확인해 주세요)")
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
                    group_match = re.match(r'^(.*?\bEntry\s*\d+)', field_lbl, re.IGNORECASE)
                    curr_group = group_match.group(1).strip() if group_match else None
                    if prev_group and curr_group and prev_group != curr_group:
                        table_data.append({"항목 (Field)": "──────────", "추출값 (Value)": "──────────", "출처 (Source)": "──────────"})
                        full_text_list.append("")
                    prev_group = curr_group
                    if not val:
                        display_val = "⚠️ 확인 필요 (미발견)"
                        full_text_list.append(f"{field_lbl}: (확인 필요)")
                    else:
                        display_val = val
                        full_text_list.append(f"{field_lbl}: {val}")
                    table_data.append({"항목 (Field)": field_lbl, "추출값 (Value)": display_val, "출처 (Source)": src if src else "-"})
                st.table(table_data)
                full_text_list.append("")
            st.markdown("#### 📋 한눈에 복사하기")
            st.text_area("아래 텍스트를 복사하여 서식에 옮겨 적으세요", value="\n".join(full_text_list), height=250)
            if st.button("＋ 새 케이스 정리하기"):
                st.session_state.prep_result = None
                st.rerun()

elif app_mode == MENU_4:
    st.title(MENU_4)
    st.caption("개별 낱장 이미지, 여러 장짜리 통짜 PDF, MS Word 서류 등을 섞어서 올려도 AI가 알아서 문서 단위로 묶거나 분할하여 CRM 파일명으로 최적화합니다.")
    if "uploader_key" not in st.session_state: st.session_state.uploader_key = str(uuid.uuid4())
    if "analysis_results" not in st.session_state: st.session_state.analysis_results = None
    uploaded_files = st.file_uploader("서류 업로드 (복수 선택 가능)", type=['jpg', 'jpeg', 'png', 'pdf', 'heic', 'HEIC', 'docx', 'DOCX', 'doc', 'DOC'], accept_multiple_files=True, key=st.session_state.uploader_key)
    if uploaded_files:
        if st.button("서류 분석 및 묶기/분할 시작", type="primary", use_container_width=True):
            results = []
            status_text = st.empty()
            progress_bar = st.progress(0)
            status_text.text("1. 전체 서류 페이지 스캔 및 미리보기 생성 중...")
            global_pages = []
            page_counter = 1
            for file in uploaded_files:
                file_bytes = file.getvalue()
                mime_type = file.type if file.type else "application/pdf"
                fname_lower = file.name.lower()
                if fname_lower.endswith(('.doc', '.docx')):
                    text_content = read_word_document_text(file_bytes, file.name)
                    img = Image.new('RGB', (800, 1000), color=(255, 255, 255))
                    draw = ImageDraw.Draw(img)
                    disp_text = f"[Word File: {file.name}]\n\n" + (text_content[:800] if text_content else "Word Document")
                    draw.text((40, 40), disp_text, fill=(0, 0, 0))
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=60)
                    global_pages.append({"global_idx": page_counter, "original_name": file.name, "mime_type": mime_type, "file_bytes": file_bytes, "pdf_page_idx": 0, "preview_bytes": buf.getvalue(), "is_word": True, "word_text": text_content})
                    page_counter += 1
                elif "pdf" in mime_type.lower() or fname_lower.endswith('.pdf'):
                    doc = fitz.open(stream=file_bytes, filetype="pdf")
                    for i in range(len(doc)):
                        if page_counter > 40: break
                        page = doc.load_page(i)
                        pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2))
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        buf = io.BytesIO()
                        img.save(buf, format="JPEG", quality=60)
                        global_pages.append({"global_idx": page_counter, "original_name": file.name, "mime_type": mime_type, "file_bytes": file_bytes, "pdf_page_idx": i, "preview_bytes": buf.getvalue(), "is_word": False})
                        page_counter += 1
                    doc.close()
                else:
                    if page_counter > 40: continue
                    img = Image.open(io.BytesIO(file_bytes))
                    img = ImageOps.exif_transpose(img)
                    if img.mode != "RGB": img = img.convert("RGB")
                    preview = img.copy()
                    max_dim = max(preview.width, preview.height)
                    if max_dim > 1200:
                        ratio = 1200.0 / float(max_dim)
                        preview = preview.resize((int(preview.width * ratio), int(preview.height * ratio)), Image.Resampling.LANCZOS)
                    buf = io.BytesIO()
                    preview.save(buf, format="JPEG", quality=60)
                    global_pages.append({"global_idx": page_counter, "original_name": file.name, "mime_type": mime_type, "file_bytes": file_bytes, "pdf_page_idx": 0, "preview_bytes": buf.getvalue(), "is_word": False})
                    page_counter += 1
            if page_counter > 40: st.warning("⚠️ 업로드된 총 페이지 수가 40장을 초과하여, 앞의 40장까지만 분석 및 병합합니다.")
            status_text.text("2. AI가 페이지별 문맥을 분석하여 연관 서류를 묶거나 나누는 중입니다...")
            prompt = f"""
            You are an expert AI document classifier for a Canadian immigration firm.
            I am providing {len(global_pages)} pages of documents uploaded by a client. 

            Your task:
            1. Read ALL pages carefully.
            2. GROUP the pages that logically belong to the SAME document type for the SAME client. 
            3. ROTATION CORRECTION
            4. For EACH grouped document, generate an EXACT filename using our strict CRM manual rules.

            Return ONLY a raw JSON object formatted EXACTLY like this:
            {{
              "page_details": [
                {{
                  "page_index": 1,
                  "rotation_reasoning": "...",
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
                contents.append({"mime_type": "image/jpeg", "data": p['preview_bytes']})
            try:
                response = safe_generate_content(contents)
                clean_text = response.text.strip().replace('```json', '').replace('```', '')
                data = json.loads(clean_text)
                page_details = data.get("page_details", [])
                rotations = {str(item.get("page_index")): item.get("rotation_needed_clockwise", 0) for item in page_details}
                raw_docs_info = data.get("documents", [])
                docs_info = []
                for d in raw_docs_info:
                    c_name = d.get("client_name", "").strip()
                    c_cat = d.get("doc_category", "").strip()
                    existing = None
                    for item in docs_info:
                        name1 = item.get("suggested_filename", "").replace(".pdf", "").replace(".jpg", "").replace(".docx", "").replace(".doc", "").strip()
                        name2 = d.get("suggested_filename", "").replace(".pdf", "").replace(".jpg", "").replace(".docx", "").replace(".doc", "").strip()
                        if name1 == name2:
                            existing = item; break
                        if c_name and c_cat and c_name == item.get("client_name", "").strip() and c_cat == item.get("doc_category", "").strip():
                            if "bank statement" in c_cat.lower() or "paystub" in c_cat.lower() or "statement" in c_cat.lower():
                                existing = item; break
                    if existing:
                        for ni in d.get("page_indices", []):
                            if ni not in existing["page_indices"]: existing["page_indices"].append(ni)
                    else: docs_info.append(d)
            except Exception as e:
                st.error(f"AI 분석 중 오류가 발생했습니다: {e}")
                docs_info = []; rotations = {}

            status_text.text("3. 분석된 정보를 바탕으로 최종 서류 결합 및 회전/압축 중입니다...")
            progress_step = 1 / max(len(docs_info), 1)
            failed_docs = []

            for idx, doc_info in enumerate(docs_info):
                indices = doc_info.get("page_indices", [])
                if not indices: continue
                final_name = doc_info.get("suggested_filename", f"Document_{idx+1}")
                is_unclassified = doc_info.get("is_unclassified", False) or "확인필요" in final_name
                is_resume = doc_info.get("is_resume", False) or "resume" in final_name.lower()
                source_names = []
                group_pages = []
                for p_idx in indices:
                    p_data = next((p for p in global_pages if p['global_idx'] == p_idx), None)
                    if p_data:
                        group_pages.append(p_data)
                        if p_data["original_name"] not in source_names: source_names.append(p_data["original_name"])

                unique_src_files = list(set([p["original_name"] for p in group_pages]))
                is_all_from_same_pdf = (len(unique_src_files) == 1 and "pdf" in group_pages[0]["mime_type"].lower())
                is_all_from_same_word = (len(unique_src_files) == 1 and group_pages[0]["is_word"])
                needs_rotation = any(int(rotations.get(str(p["global_idx"]), 0)) != 0 for p in group_pages)
                
                if is_resume and is_all_from_same_word:
                    base_name, _ = os.path.splitext(final_name)
                    orig_ext = os.path.splitext(group_pages[0]["original_name"])[1]
                    final_name = base_name + orig_ext
                    comp_bytes = group_pages[0]["file_bytes"]
                    out_mime = group_pages[0]["mime_type"]
                    orig_kb = len(comp_bytes) / 1024
                    src_display = ", ".join(source_names)
                    if len(src_display) > 30: src_display = src_display[:27] + "..."
                    results.append({"original_name": f"분석결과 ({src_display})", "suggested_filename": final_name, "category": doc_info.get("doc_category", "기타"), "client_name": doc_info.get("client_name", ""), "mime": out_mime, "orig_kb": orig_kb, "comp_kb": orig_kb, "bytes": comp_bytes, "is_unclassified": is_unclassified})
                    progress_bar.progress(min((idx + 1) * progress_step, 1.0))
                    continue

                if not (final_name.lower().endswith(".pdf") or final_name.lower().endswith(".jpg") or final_name.lower().endswith(".jpeg") or final_name.lower().endswith(".doc") or final_name.lower().endswith(".docx")):
                    if "photo" in final_name.lower() and len(indices) == 1: final_name += ".jpg"
                    else: final_name += ".pdf"
                        
                try:
                    if is_all_from_same_pdf:
                        src_doc = fitz.open(stream=group_pages[0]["file_bytes"], filetype="pdf")
                        pdf_indices = [p["pdf_page_idx"] for p in group_pages]
                        if len(pdf_indices) == len(src_doc) and not needs_rotation and pdf_indices == list(range(len(src_doc))):
                            merged_pdf_bytes = group_pages[0]["file_bytes"]
                            src_doc.close()
                            final_processed_bytes = sanitize_and_unlock_pdf(merged_pdf_bytes)
                        else:
                            new_doc = fitz.open()
                            for p_data in group_pages:
                                p_idx = p_data["pdf_page_idx"]
                                rot = int(rotations.get(str(p_data["global_idx"]), 0))
                                new_doc.insert_pdf(src_doc, from_page=p_idx, to_page=p_idx)
                                if rot != 0: new_doc[-1].set_rotation((new_doc[-1].rotation + rot) % 360)
                            merged_pdf_bytes_io = io.BytesIO()
                            new_doc.save(merged_pdf_bytes_io, deflate=True)
                            new_doc.close(); src_doc.close()
                            merged_pdf_bytes = merged_pdf_bytes_io.getvalue()
                            final_processed_bytes = sanitize_and_unlock_pdf(merged_pdf_bytes)
                        comp_bytes, out_mime = process_and_compress_file(final_processed_bytes, "application/pdf", final_name, already_sanitized=True)
                    else:
                        new_doc = fitz.open()
                        for p_data in group_pages:
                            rot = int(rotations.get(str(p_data["global_idx"]), 0))
                            if p_data.get("is_word"):
                                pdf_page = new_doc.new_page(width=595, height=842)
                                w_text = p_data.get("word_text", "")
                                pdf_page.insert_text((50, 50), w_text[:3000] if w_text else f"Word Document: {p_data['original_name']}", fontsize=10)
                            elif "pdf" in p_data["mime_type"].lower():
                                src_doc = fitz.open(stream=p_data["file_bytes"], filetype="pdf")
                                p_idx = p_data["pdf_page_idx"]
                                new_doc.insert_pdf(src_doc, from_page=p_idx, to_page=p_idx)
                                if rot != 0: new_doc[-1].set_rotation((new_doc[-1].rotation + rot) % 360)
                                src_doc.close()
                            else:
                                img = Image.open(io.BytesIO(p_data["file_bytes"]))
                                img = ImageOps.exif_transpose(img)
                                if img.mode != "RGB": img = img.convert("RGB")
                                if rot != 0: img = img.rotate(-rot, expand=True)
                                img_buf = io.BytesIO()
                                img.save(img_buf, format="JPEG", quality=92)
                                pdf_page = new_doc.new_page(width=img.width, height=img.height)
                                pdf_page.insert_image(pdf_page.rect, stream=img_buf.getvalue())
                        merged_pdf_bytes_io = io.BytesIO()
                        new_doc.save(merged_pdf_bytes_io, deflate=True)
                        new_doc.close()
                        final_processed_bytes = merged_pdf_bytes_io.getvalue()
                        is_jpeg = final_name.lower().endswith(('.jpg', '.jpeg'))
                        if is_jpeg and len(indices) == 1:
                            p_data = group_pages[0]
                            if "pdf" in p_data["mime_type"].lower():
                                src_doc = fitz.open(stream=p_data["file_bytes"], filetype="pdf")
                                page = src_doc.load_page(p_data["pdf_page_idx"])
                                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                                src_doc.close()
                            else:
                                img = Image.open(io.BytesIO(p_data["file_bytes"]))
                                img = ImageOps.exif_transpose(img)
                            try:
                                rot = int(rotations.get(str(p_data["global_idx"]), 0))
                                if rot != 0: img = img.rotate(-rot, expand=True)
                            except Exception: pass
                            if img.mode != "RGB": img = img.convert("RGB")
                            img_buf = io.BytesIO()
                            img.save(img_buf, format="JPEG", quality=95)
                            comp_bytes, out_mime = process_and_compress_file(img_buf.getvalue(), "image/jpeg", final_name, already_sanitized=True)
                        else:
                            final_processed_bytes = sanitize_and_unlock_pdf(merged_pdf_bytes)
                            comp_bytes, out_mime = process_and_compress_file(final_processed_bytes, "application/pdf", final_name, already_sanitized=True)
                        
                    orig_kb = sum([len(p["file_bytes"]) for p in group_pages]) / 1024
                    comp_kb = len(comp_bytes) / 1024
                    src_display = ", ".join(source_names)
                    if len(src_display) > 30: src_display = src_display[:27] + "..."
                    results.append({"original_name": f"분석결과 ({src_display})", "suggested_filename": final_name, "category": doc_info.get("doc_category", "기타"), "client_name": doc_info.get("client_name", ""), "mime": out_mime, "orig_kb": orig_kb, "comp_kb": comp_kb, "bytes": comp_bytes, "is_unclassified": is_unclassified})
                except Exception as e:
                    failed_docs.append(final_name)
                progress_bar.progress(min((idx + 1) * progress_step, 1.0))
                
            if failed_docs:
                status_text.warning("일부 서류 처리에 실패했지만, 나머지 서류의 최적화는 완료되었습니다.")
                st.warning("⚠️ **아래 서류는 파일 손상 또는 변환 중 오류가 발생하여 제외되었습니다. 원본 파일을 직접 확인해 주세요:**\n\n" + "\n".join([f"- {f}" for f in failed_docs]))
            else:
                status_text.success("모든 서류의 묶기/분할 및 스마트 회전 최적화가 완료되었습니다.")
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
                st.caption(f"{item['orig_kb']:.1f} KB ➡️ **{item['comp_kb']:.1f} KB**")
                if item.get("is_unclassified"): st.warning("⚠️ 규칙 미확인 서류 (파일명 수동 확인 필요)")
            with col2:
                user_edited_name = st.text_input("파일명", value=item['suggested_filename'], key=f"edit_{idx}", label_visibility="collapsed")
                final_downloads.append((user_edited_name, item['bytes'], item['mime']))
            with col3:
                st.download_button("⬇️ 개별 다운로드", data=item['bytes'], file_name=user_edited_name, mime=item['mime'], key=f"dl_btn_{idx}")
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for fname, fbytes, _ in final_downloads:
                zip_file.writestr(fname, fbytes)
        zip_buffer.seek(0)
        today_str = datetime.date.today().strftime("%Y%m%d")
        st.markdown("---")
        st.download_button("📦 전체 서류 ZIP 다운로드", data=zip_buffer, file_name=f"CRM_Documents_{today_str}.zip", mime="application/zip", type="primary", use_container_width=True)
        st.markdown("---")
        if st.button("🔄 전체 리셋", type="secondary", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key != "password_correct": del st.session_state[key]
            st.session_state.uploader_key = str(uuid.uuid4())
            st.rerun()

# ------------------------------------------
# 메뉴 5: 질문지 자동 이어채우기 (QA Re-fill)
# ------------------------------------------
elif app_mode == MENU_5:
    st.title(MENU_5)
    st.caption("기존 손님 서류들과 새 빈 질문지를 올리시면, 알고 있는 정보를 폼 필드 칸에 정렬하여 완벽하게 채워 드립니다.")

    if "qa_result_files" not in st.session_state:
        st.session_state.qa_result_files = None

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 손님의 기존 자료들")
        source_files = st.file_uploader(
            "이전 질문지, 여권, 퍼밋, 이력서 등 (복수 선택)",
            type=['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'heic'],
            accept_multiple_files=True,
            key="qa_sources"
        )
    with c2:
        st.subheader("2. 새로 채울 빈 질문지")
        new_q_file = st.file_uploader(
            "채워 넣을 대상 질문지 1부 (.pdf 또는 .docx)",
            type=['pdf', 'docx'],
            key="qa_target"
        )

    st.markdown("---")

    col_action, col_reset = st.columns([8, 2])

    with col_action:
        if st.button("✨ 질문지 자동 완성하기", type="primary", use_container_width=True):
            if not source_files:
                st.warning("1단계: 손님의 기존 서류를 최소 1개 이상 업로드해 주세요.")
            elif not new_q_file:
                st.warning("2단계: 새로 채워야 할 질문지 파일을 업로드해 주세요.")
            else:
                with st.spinner("원본 폼 필드를 분석하여 빈틈없이 작성 중입니다. 잠시만 기다려 주세요..."):
                    qa_data = analyze_qa_for_fill(source_files, new_q_file)
                    
                    if qa_data:
                        base_name = os.path.splitext(new_q_file.name)[0]
                        is_docx = new_q_file.name.lower().endswith(".docx")
                        
                        if is_docx:
                            filled_bytes = fill_docx_questionnaire(new_q_file.getvalue(), qa_data)
                            q_filename = f"{base_name}_완성본.docx"
                            q_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        else:
                            filled_bytes = fill_qa_pdf_interactive(new_q_file.getvalue(), qa_data)
                            q_filename = f"{base_name}_완성본.pdf"
                            q_mime = "application/pdf"
                        
                        # 결과를 Session State에 안전하게 저장하여 버튼 증발 방지
                        st.session_state.qa_result_files = {
                            "q_filename": q_filename,
                            "q_bytes": filled_bytes,
                            "q_mime": q_mime
                        }

    with col_reset:
        if st.button("🔄 전체 리셋", use_container_width=True, key="reset_menu5"):
            st.session_state.qa_result_files = None
            if "qa_sources" in st.session_state: del st.session_state["qa_sources"]
            if "qa_target" in st.session_state: del st.session_state["qa_target"]
            st.rerun()

    # 💡 다운로드 버튼 렌더링 (State 유지)
    if st.session_state.qa_result_files:
        res = st.session_state.qa_result_files
        st.success("✅ 질문지 작성이 완벽하게 완료되었습니다! 아래 버튼을 클릭하세요.")
        
        st.download_button(
            label=f"📥 완성된 질문지 내려받기",
            data=res["q_bytes"],
            file_name=res["q_filename"],
            mime=res["q_mime"],
            type="primary",
            use_container_width=True
        )
