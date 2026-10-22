import streamlit as st

# 📌 페이지 설정
st.set_page_config(page_title="CanNest 질문지 자동 완성", layout="wide")

import google.generativeai as genai
import fitz  # PyMuPDF
from PIL import Image, ImageOps
import pillow_heif
import io
import json
import os
from concurrent.futures import ThreadPoolExecutor

try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    pass

pillow_heif.register_heif_opener()
Image.MAX_IMAGE_PIXELS = None

# ==========================================
# 0. API 키 설정
# ==========================================
if "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Cloud의 Secrets에 GEMINI_API_KEY 설정이 필요합니다.")
    st.stop()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

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
# 1. 파일 데이터 파싱 (PDF 폼 필드 이름 추출 기능 포함)
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

def prepare_document_for_gemini(file_bytes, mime_type, file_name=""):
    ext = os.path.splitext(file_name)[1].lower() if file_name else ""
    if "word" in mime_type.lower() or ext in ['.doc', '.docx']:
        text = read_word_document_text(file_bytes, file_name)
        if text:
            return [f"\n--- [Word Document: {file_name}] ---\n{text[:35000]}\n"]
            
    if "pdf" in mime_type.lower() or ext == '.pdf':
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            field_list = []
            
            for page in doc:
                text += page.get_text("text") + "\n"
                
                # 💡 [핵심 기술] PDF 내부에 존재하는 모든 클릭 가능한 서식(Widget)의 고유 이름(Field Name)을 AI에게 전달
                for w in page.widgets():
                    if w.field_value:
                        text += f"[{w.field_name}]: {w.field_value}\n"
                    
                    r = w.rect
                    clip = fitz.Rect(max(0, r.x0 - 100), max(0, r.y0 - 25), min(page.rect.width, r.x1 + 50), min(page.rect.height, r.y1 + 25))
                    nearby = page.get_text("text", clip=clip).replace('\n', ' ').strip()
                    field_list.append(f"Field Name: '{w.field_name}' | Context Nearby: '{nearby}' | Type: {w.field_type_string}")
            
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
# 2. AI 질문지 분석 및 매핑 (오류 수정판)
# ==========================================
def analyze_qa_for_fill(client_files, target_file):
    prompt = """
    You are an expert Canadian immigration assistant filling out a blank questionnaire based on client documents.
    The FIRST document is the NEW BLANK QUESTIONNAIRE. The rest are CLIENT SOURCE MATERIALS.

    [CRITICAL RULES]
    1. CURRENT DATE: Today is October 2026.
    2. CURRENT JOB LOGIC: If a job's end date is BEFORE October 2026 (e.g., '2026-09-26' or '2025-06-30'), it is NOT 'Current'. ONLY jobs marked as 'Present' or with end dates in the future (e.g., '2028-05') belong in the "Current Occupation" section.
    3. LANGUAGE: All output values MUST be in ENGLISH (Latin alphabet, numbers). NO KOREAN inside values.
    4. PDF MAPPING: Look at the [FILLABLE PDF FIELDS AVAILABLE] section I provided for the blank form. Put the EXACT "Field Name" as the key in the `pdf_fields` dictionary, and the extracted English answer as the value.
    5. CHECKBOXES: For checkboxes/radio buttons in PDF, set the value to "Yes" if it applies.
    6. FULL HISTORY ARRAYS: Always provide the COMPLETE chronological history in the `full_education_history`, `full_work_history`, and `full_travel_history` arrays. We will use these arrays to generate an Additional Information Word document if the history exceeds the PDF boxes. Fill the PDF boxes with the most recent entries up to the available slots.

    Return ONLY a raw JSON object formatted strictly as:
    {
      "applicant_name": "Full Name in Title Case",
      "pdf_fields": {
         "Exact_Field_Name_1": "Danyang",
         "Exact_Field_Name_2": "Republic of Korea",
         "Exact_Checkbox_Name": "Yes"
      },
      "docx_fields": {
         "출생 도시": "Value if the target was DOCX",
         "기혼": "V"
      },
      "full_education_history": [
        {"from": "YYYY-MM", "to": "YYYY-MM", "field_degree": "Degree", "school_name": "School", "location": "City, Country"}
      ],
      "full_work_history": [
        {"from": "YYYY-MM", "to": "YYYY-MM", "job_title": "Title", "company_name": "Company", "location": "City, Country"}
      ],
      "full_travel_history": [
        {"from": "YYYY-MM-DD", "to": "YYYY-MM-DD", "city": "City", "country": "Country", "purpose": "Tourism"}
      ]
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
    """
    도장 찍듯 덮어씌우는 것이 아니라, 원본 PDF의 폼 필드(Widget) 속성에 
    직접 값을 기입하여 정렬과 폰트 크기를 PDF 고유 서식에 맡기는 완벽한 방식입니다.
    """
    doc = fitz.open(stream=target_bytes, filetype="pdf")
    pdf_fields = qa_data.get("pdf_fields", {})
    fontfile = "cour.ttf" if os.path.exists("cour.ttf") else None
    
    # PDF 문서 내의 모든 폼 필드를 순회하며 값 주입
    for page in doc:
        for w in page.widgets():
            fn = w.field_name
            if fn in pdf_fields:
                val = pdf_fields[fn]
                if val:
                    ftype = getattr(w, "field_type_string", "").lower()
                    
                    # 체크박스 처리
                    if "check" in ftype or "radio" in ftype:
                        if str(val).lower() in ['yes', 'true', 'v', 'on', '1', 'checked']:
                            try:
                                on_states = w.on_state()
                                if on_states:
                                    w.field_value = on_states[0]  # PDF에 정의된 '선택' 값 (예: 'Yes')
                                else:
                                    w.field_value = "Yes"
                            except Exception:
                                w.field_value = True
                    
                    # 일반 텍스트 필드 처리
                    else:
                        w.field_value = str(val)
                        if hasattr(w, "field_flags"):
                            w.field_flags &= ~1  # 수정 가능 상태로 유지
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

# ==========================================
# 4. Word 질문지 작성 엔진 (.docx)
# ==========================================
def fill_docx_questionnaire(target_bytes, qa_data):
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
# 5. 별첨 Word 부속 서류(Additional History) 생성 엔진
# ==========================================
def create_qa_appendix_docx(qa_data):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.6)
        s.bottom_margin = Inches(0.6)
        s.left_margin = Inches(0.6)
        s.right_margin = Inches(0.6)

    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(9.5)

    client_name = qa_data.get("applicant_name", "APPLICANT")
    h1 = doc.add_paragraph()
    r1 = h1.add_run("ADDITIONAL INFORMATION SHEET")
    r1.bold = True; r1.font.size = Pt(14)
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph()
    sub.add_run(f"Applicant Name: {client_name}").bold = True
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    def set_cell_bg(cell, fill_hex):
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_hex)
        cell._element.get_or_add_tcPr().append(shd)

    def build_table(title, items, headers, keys):
        if not items: return
        p = doc.add_paragraph()
        p.add_run(title).bold = True
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = 'Table Grid'
        for i, text in enumerate(headers):
            t.cell(0, i).text = text
            t.cell(0, i).paragraphs[0].runs[0].bold = True
            set_cell_bg(t.cell(0, i), "EAECEE")
        for item in items:
            row = t.add_row().cells
            for i, key in enumerate(keys):
                row[i].text = item.get(key, "")
        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    build_table("1. Full Education History", qa_data.get("full_education_history", []), 
                ["From", "To", "Field of Study / Degree", "School Name", "City / Country"], 
                ["from", "to", "field_degree", "school_name", "location"])
                
    build_table("2. Full Work / Activity History", qa_data.get("full_work_history", []), 
                ["From", "To", "Occupation / Title", "Company Name", "City / Country"], 
                ["from", "to", "job_title", "company_name", "location"])
                
    build_table("3. Travel History", qa_data.get("full_travel_history", []), 
                ["From", "To", "City", "Country", "Purpose"], 
                ["from", "to", "city", "country", "purpose"])

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()

# ==========================================
# 6. UI 화면 구성 (단독 5번 도구)
# ==========================================
st.title("📝 질문지 자동 이어채우기")
st.caption("기존 손님 서류들과 새 빈 질문지를 올리시면, 알고 있는 정보를 빈칸에 맞게 정확히 기입해 드립니다.")

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
    if st.button("질문지 자동 완성하기", type="primary", use_container_width=True):
        if not source_files:
            st.warning("1단계: 손님의 기존 서류를 최소 1개 이상 업로드해 주세요.")
        elif not new_q_file:
            st.warning("2단계: 새로 채워야 할 질문지 파일을 업로드해 주세요.")
        else:
            with st.spinner("원본 서식 필드(Widget)를 스캔하고 정확한 위치에 답변을 채우는 중입니다..."):
                qa_data = analyze_qa_for_fill(source_files, new_q_file)
                
                if qa_data:
                    base_name = os.path.splitext(new_q_file.name)[0]
                    is_docx = new_q_file.name.lower().endswith(".docx")
                    app_info = qa_data.get("applicant_name", "Client").replace(" ", "_")
                    
                    if is_docx:
                        filled_bytes = fill_docx_questionnaire(new_q_file.getvalue(), qa_data)
                        q_filename = f"{base_name}_완성본.docx"
                        q_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    else:
                        filled_bytes = fill_qa_pdf_interactive(new_q_file.getvalue(), qa_data)
                        q_filename = f"{base_name}_완성본.pdf"
                        q_mime = "application/pdf"

                    # 💡 PDF 폼 필드 칸 수(보통 2~3개)를 초과하는 히스토리가 있을 때만 별첨 생성 판단
                    # 대부분의 질문지 양식에는 학력 2개, 경력 2개가 들어갑니다.
                    edu_count = len(qa_data.get("full_education_history", []))
                    work_count = len(qa_data.get("full_work_history", []))
                    travel_count = len(qa_data.get("full_travel_history", []))
                    
                    has_extra = edu_count > 2 or work_count > 2 or travel_count > 3
                    
                    appendix_bytes = None
                    appendix_filename = ""
                    if has_extra:
                        appendix_bytes = create_qa_appendix_docx(qa_data)
                        appendix_filename = f"{app_info}_Additional_History.docx"
                    
                    # 💡 다운로드 버튼을 눌러도 사라지지 않도록 Session State에 데이터 보존
                    st.session_state.qa_result_files = {
                        "q_filename": q_filename,
                        "q_bytes": filled_bytes,
                        "q_mime": q_mime,
                        "has_extra": has_extra,
                        "app_filename": appendix_filename,
                        "app_bytes": appendix_bytes
                    }

with col_reset:
    if st.button("🔄 전체 리셋", use_container_width=True):
        st.session_state.qa_result_files = None
        st.rerun()

# 💡 Session State에 저장된 파일이 있으면 다운로드 버튼을 렌더링
if st.session_state.qa_result_files:
    res = st.session_state.qa_result_files
    st.success("✅ 질문지 작성이 완벽하게 완료되었습니다! 아래 버튼을 클릭하세요.")
    
    dl_c1, dl_c2 = st.columns(2)
    with dl_c1:
        st.download_button(
            label=f"📥 1. 완성된 질문지 내려받기 (.pdf/.docx)",
            data=res["q_bytes"],
            file_name=res["q_filename"],
            mime=res["q_mime"],
            type="primary",
            use_container_width=True
        )
        
    with dl_c2:
        if res["has_extra"]:
            st.download_button(
                label=f"📄 2. [별첨] 추가 학력·경력 내역서 내려받기 (.docx)",
                data=res["app_bytes"],
                file_name=res["app_filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="secondary",
                use_container_width=True
            )
        else:
            st.info("💡 학력/경력이 질문지 칸 내에 모두 들어가서 별첨 서류는 생성되지 않았습니다.")
