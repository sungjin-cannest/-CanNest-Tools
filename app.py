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
# 1. 문서 파싱 엔진
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
            return [f"\n--- [Word Document: {file_name}] ---\n{word_text[:35000]}\n"]
            
    if "pdf" in mime_type.lower() or ext == '.pdf':
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            field_list = []
            
            for page in doc:
                text += page.get_text("text") + "\n"
                
                # 대화형 서식(Widget) 필드 리스트를 AI에게 제공
                for w in page.widgets():
                    if w.field_value:
                        text += f"[{w.field_name}]: {w.field_value}\n"
                    
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
                return [f"\n--- [Document: {file_name}] ---\n{text[:35000]}\n"]
            else:
                images = []
                for p_idx in range(min(len(doc), 10)):
                    page = doc.load_page(p_idx)
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=75)
                    images.append({"mime_type": "image/jpeg", "data": buf.getvalue()})
                doc.close()
                return images
        except Exception:
            pass
    return [{"mime_type": mime_type, "data": file_bytes}]

def batch_process_client_files(files):
    def worker(f):
        mime = f.type if f.type else "application/pdf"
        return prepare_document_for_gemini(f.getvalue(), mime, f.name)
    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(executor.map(worker, files))
    flat = []
    for r in results:
        flat.extend(r)
    return flat

# ==========================================
# 2. AI 질문지 분석 및 완벽 매핑
# ==========================================
def analyze_qa_for_fill(client_files, target_file):
    prompt = """
    You are an expert Canadian immigration assistant filling out a blank questionnaire based on client documents.
    The FIRST document is the NEW BLANK QUESTIONNAIRE. The rest are CLIENT SOURCE MATERIALS.

    [CRITICAL RULES - MUST OBEY OR FAIL]
    1. UNKNOWN INFO = OMIT: If the client documents do NOT explicitly mention an answer for Tuberculosis, Medical issues, Crime, Visa Refusals, or Biometrics, YOU MUST OMIT those fields from the JSON entirely. ABSOLUTELY DO NOT GUESS "NO", "Off", or "False". Leave them unmapped and blank.
    2. ENGLISH TESTS: If AND ONLY IF the client explicitly has an IELTS, CELPIP, or Duolingo score in the materials, map "Yes" to the English test YES checkbox. Otherwise, LEAVE IT COMPLETELY BLANK.
    3. CANADA VISA HISTORY: If the client currently holds a Canadian permit (e.g., Visitor Record, Work Permit, Study Permit), they obviously applied for it. You MUST map this fact to the "Canada Visa Application History" text field (e.g., "Visitor Record (Valid until YYYY-MM-DD)").
    4. CURRENT JOB LOGIC: Today is October 2026. A job ending in "2026-09" or earlier is NOT current. In the work history table, the FIRST ROW is reserved for the "Current" job. IF THE CLIENT DOES NOT HAVE A CURRENT JOB, YOU MUST LEAVE THE ENTIRE FIRST ROW BLANK. Start mapping their past jobs from the SECOND ROW downwards.
    5. TABLE SEQUENTIAL MAPPING: Map Education and Travel history sequentially starting from Row 1. NO empty gaps between rows. Map Work history from Row 2 if there is no current job, otherwise Row 1.
    6. OVERSEAS RESIDENCE: If you infer overseas residence from overseas work experience, DO NOT fill in the "From/To" dates for that residence (leave dates blank), but DO fill in the City/Country.
    7. FAMILY INFO: Extract Spouse and Children information if available and map them.
    8. SIGNATURE DATE = BLANK: DO NOT fill in the signature Date. Leave it blank.
    9. LANGUAGE: All output values MUST be in ENGLISH (Latin alphabet, numbers). NO KOREAN inside values.

    Return ONLY a raw JSON object formatted strictly as:
    {
      "pdf_fields": {
         "Exact_Field_Name_1": "Extracted Value",
         "Exact_Checkbox_Name_For_Yes": "Yes"
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

# ==========================================
# 3. PDF 대화형 폼(Widget) 직접 채우기 엔진
# ==========================================
def fill_qa_pdf_interactive(target_bytes, qa_data):
    doc = fitz.open(stream=target_bytes, filetype="pdf")
    pdf_fields = qa_data.get("pdf_fields", {})
    fontfile = "cour.ttf" if os.path.exists("cour.ttf") else None
    
    for page in doc:
        for w in page.widgets():
            fn = w.field_name
            if fn in pdf_fields:
                val = pdf_fields[fn]
                if val:
                    ftype = getattr(w, "field_type_string", "").lower()
                    
                    if "check" in ftype or "radio" in ftype:
                        if str(val).lower() in ['yes', 'true', 'v', 'on', '1', 'checked']:
                            try:
                                on_states = [s for s in w.on_state() if s.lower() != 'off']
                                if on_states:
                                    w.field_value = on_states[0]
                                else:
                                    w.field_value = True
                                w.update()
                            except Exception:
                                pass
                    else:
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
    st.caption("기존 손님 서류들과 새 빈 질문지를 올리시면, 알고 있는 정보를 서식 칸에 맞춰 기입해 드립니다.")

    # 파일 업로더 초기화를 위한 동적 Key 생성
    if "menu5_uploader_key" not in st.session_state:
        st.session_state.menu5_uploader_key = str(uuid.uuid4())

    if "qa_result_files" not in st.session_state:
        st.session_state.qa_result_files = None

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 손님의 기존 자료들")
        source_files = st.file_uploader(
            "이전 질문지, 여권, 퍼밋, 이력서 등 (복수 선택)",
            type=['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'heic'],
            accept_multiple_files=True,
            key=f"qa_sources_{st.session_state.menu5_uploader_key}"
        )
    with c2:
        st.subheader("2. 새로 채울 빈 질문지")
        new_q_file = st.file_uploader(
            "채워 넣을 대상 질문지 1부 (.pdf 또는 .docx)",
            type=['pdf', 'docx'],
            key=f"qa_target_{st.session_state.menu5_uploader_key}"
        )

    st.markdown("---")

    col_action, col_reset = st.columns([8, 2])

    with col_action:
        if st.button("✨ 질문지 정보 기입 시작", type="primary", use_container_width=True):
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
                            filled_bytes = fill_qa_docx(new_q_file.getvalue(), qa_data)
                            q_filename = f"{base_name}_Info_Filled.docx"
                            q_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        else:
                            filled_bytes = fill_qa_pdf_interactive(new_q_file.getvalue(), qa_data)
                            q_filename = f"{base_name}_Info_Filled.pdf"
                            q_mime = "application/pdf"

                        # 결과를 Session State에 저장하여 버튼 증발 방지
                        st.session_state.qa_result_files = {
                            "q_filename": q_filename,
                            "q_bytes": filled_bytes,
                            "q_mime": q_mime
                        }

    with col_reset:
        if st.button("🔄 전체 리셋", use_container_width=True):
            st.session_state.qa_result_files = None
            st.session_state.menu5_uploader_key = str(uuid.uuid4()) # 새로운 키를 생성하여 업로더 완전 초기화
            st.rerun()

    # 💡 결과 렌더링
    if st.session_state.qa_result_files:
        res = st.session_state.qa_result_files
        st.success("✅ 기존 정보 기입이 완료되었습니다! 확인하지 못한 항목은 비워두었으니 손님께 마저 작성해 달라고 요청하세요.")
        
        st.download_button(
            label=f"📥 Info Filled 질문지 내려받기",
            data=res["q_bytes"],
            file_name=res["q_filename"],
            mime=res["q_mime"],
            type="primary"
        )
