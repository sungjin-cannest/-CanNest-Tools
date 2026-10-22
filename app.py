import streamlit as st

# 📌 페이지 설정
st.set_page_config(page_title="CanNest 질문지 자동 완성", layout="wide")

import google.generativeai as genai
import fitz  # PyMuPDF
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

# ==========================================
# 0. API 키 설정
# ==========================================
if "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Cloud의 Secrets에 GEMINI_API_KEY 설정이 필요합니다.")
    st.stop()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

def safe_generate_content(contents):
    candidate_models = ['gemini-3.6-flash']
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
# 1. 문서 파싱 및 AI 전송 준비
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
            return [f"\n--- [Document: {file_name}] ---\n{text[:35000]}\n"]
            
    if "pdf" in mime_type.lower() or ext == '.pdf':
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            for page in doc:
                text += page.get_text("text") + "\n"
                for widget in page.widgets():
                    if widget.field_value:
                        text += f"[{widget.field_name}]: {widget.field_value}\n"
            
            if len(text.strip()) > 100:
                doc.close()
                return [f"\n--- [Document: {file_name}] ---\n{text[:35000]}\n"]
            else:
                from PIL import Image
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
# 2. AI 질문지 분석 및 정형 추출
# ==========================================
def analyze_and_extract_qa(client_files, target_file):
    prompt = """
    You are an expert Canadian immigration case assistant.
    The FIRST document is a NEW BLANK IMMIGRATION QUESTIONNAIRE.
    The REMAINING documents are CLIENT MATERIALS (past questionnaires, passports, permits, resumes, certificates).

    [TASK]
    Extract all verified facts from client materials to fill into the new questionnaire.

    [STRICT RULES]
    1. LANGUAGE: All extracted values MUST be in ENGLISH (numbers, Latin alphabet, standard punctuation). NO Korean characters in values.
    2. NAMES: Strictly match Passport Title Case (e.g., 'Jihyun Cheung').
    3. ACCURACY: If information is not confirmed in the materials, return an empty string "". Do not guess.
    4. EDUCATION & WORK: Extract ALL post-secondary and work history entries chronologically (most recent first).

    Return ONLY a raw JSON object with this exact structure:
    {
      "applicant": {
        "full_name": "Full Name",
        "dob": "YYYY-MM-DD",
        "birth_city": "City",
        "birth_country": "Country",
        "email": "Email",
        "phone": "Phone Number",
        "address": "Full Address with Postal Code",
        "marital_status": "Married / Single / Common-Law / Divorced",
        "spouse_surname": "Spouse Surname or empty",
        "spouse_given_name": "Spouse Given Name or empty",
        "marriage_date": "YYYY-MM-DD or empty",
        "spouse_dob": "YYYY-MM-DD or empty",
        "prev_spouse_surname": "",
        "prev_spouse_given_name": "",
        "prev_marriage_from": "",
        "prev_marriage_to": "",
        "prev_spouse_dob": ""
      },
      "checkbox_checks": [
        "기혼", "NO", "아니오", "Single"
      ],
      "education_history": [
        {
          "from": "YYYY-MM",
          "to": "YYYY-MM",
          "field_degree": "Degree / Field of Study",
          "school_name": "School Name",
          "location": "City, Country"
        }
      ],
      "work_history": [
        {
          "from": "YYYY-MM",
          "to": "YYYY-MM",
          "job_title": "Occupation / Job Title",
          "company_name": "Company / Organization Name",
          "location": "City, Country"
        }
      ],
      "travel_history": [
        {
          "from": "YYYY-MM-DD",
          "to": "YYYY-MM-DD",
          "city": "City",
          "country": "Country",
          "purpose": "Tourism / Study / Work"
        }
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
# 3. PDF 질문지 작성 엔진 (PyMuPDF 표 칸 정확한 매핑)
# ==========================================
def fill_pdf_questionnaire(target_bytes, qa_data):
    doc = fitz.open(stream=target_bytes, filetype="pdf")
    BLUE_COLOR = (0.12, 0.31, 0.55)  # 타이핑 색상 (#1F4E8C)
    app = qa_data.get("applicant", {})

    # 1. 인적사항 단일 필드 매핑
    label_mapping = [
        ("출생 도시", app.get("birth_city", ""), False),
        ("출생 국가", app.get("birth_country", ""), False),
        ("이메일 주소", app.get("email", ""), False),
        ("연락처", app.get("phone", ""), False),
        ("현재 거주지 주소", app.get("address", ""), True),
        ("사실혼이 시작된 날짜", app.get("marriage_date", ""), False),
        ("3) 배우자 생년월일", app.get("spouse_dob", ""), False)
    ]

    for label, val, is_long in label_mapping:
        if not val:
            continue
        for page in doc:
            rects = page.search_for(label)
            if rects:
                r = rects[0]
                cell_box = None
                try:
                    for tab in page.find_tables():
                        for row in tab.rows:
                            for c_idx, cell in enumerate(row.cells):
                                if cell and fitz.Rect(cell).intersects(r):
                                    if c_idx + 1 < len(row.cells) and row.cells[c_idx + 1]:
                                        cell_box = fitz.Rect(row.cells[c_idx + 1])
                                        break
                except Exception:
                    pass

                if cell_box:
                    target_rect = fitz.Rect(cell_box.x0 + 4, cell_box.y0 + 2, cell_box.x1 - 4, cell_box.y1 - 2)
                else:
                    if is_long:
                        target_rect = fitz.Rect(r.x1 + 10, r.y0 - 2, page.rect.width - 35, r.y1 + 18)
                    else:
                        target_rect = fitz.Rect(r.x1 + 8, r.y0 - 2, min(r.x1 + 220, page.rect.width - 35), r.y1 + 4)

                fs = 8.5 if is_long else 9.5
                try:
                    page.insert_textbox(target_rect, val, fontsize=fs, fontname="helv", color=BLUE_COLOR, align=0)
                except Exception:
                    pass
                break

    # 2. 배우자 성/이름 처리
    for page in doc:
        rects_surname = page.search_for("성:")
        rects_given = page.search_for("이름:")
        if rects_surname and app.get("spouse_surname"):
            r_s = rects_surname[0]
            page.insert_textbox(fitz.Rect(r_s.x1 + 5, r_s.y0 - 2, r_s.x1 + 100, r_s.y1 + 3), app.get("spouse_surname"), fontsize=9.5, fontname="helv", color=BLUE_COLOR)
        if rects_given and app.get("spouse_given_name"):
            r_g = rects_given[0]
            page.insert_textbox(fitz.Rect(r_g.x1 + 5, r_g.y0 - 2, r_g.x1 + 120, r_g.y1 + 3), app.get("spouse_given_name"), fontsize=9.5, fontname="helv", color=BLUE_COLOR)

    # 3. 체크박스 마킹
    for chk in qa_data.get("checkbox_checks", []):
        if not chk:
            continue
        for page in doc:
            rects = page.search_for(chk)
            if rects:
                r = rects[0]
                box = fitz.Rect(r.x0 - 16, r.y0 - 1, r.x0 - 2, r.y1 + 1)
                try:
                    page.insert_textbox(box, "V", fontsize=10, fontname="helv", color=BLUE_COLOR, align=1)
                except Exception:
                    pass
                break

    # 4. 학력 및 직장 표 정확한 칸 채우기
    def fill_table_safely(page, keyword, items_list, col_keys):
        if not items_list:
            return
            
        try:
            tables = page.find_tables()
            matched_tab = None
            for tab in tables:
                tab_text = page.get_text("text", clip=tab.bbox)
                if keyword in tab_text:
                    matched_tab = tab
                    break
            
            if matched_tab and len(matched_tab.rows) > 1:
                available_rows = len(matched_tab.rows) - 1
                for idx in range(min(available_rows, len(items_list))):
                    data = items_list[idx]
                    row_cells = matched_tab.rows[1 + idx].cells
                    for c_idx, key in enumerate(col_keys):
                        if c_idx < len(row_cells) and row_cells[c_idx]:
                            c_rect = fitz.Rect(row_cells[c_idx])
                            # 셀 내부 패딩(Padding)을 주어 도장처럼 겹치지 않게 함
                            pad_rect = fitz.Rect(c_rect.x0 + 4, c_rect.y0 + 3, c_rect.x1 - 4, c_rect.y1 - 3)
                            val = data.get(key, "")
                            if val:
                                page.insert_textbox(pad_rect, str(val), fontsize=8.0, fontname="helv", color=BLUE_COLOR, align=0)
                
                # 칸이 부족한 경우 마지막 셀 하단에 별첨 참조 표시
                if len(items_list) > available_rows:
                    last_cell = matched_tab.rows[-1].cells[-1]
                    if last_cell:
                        r_annot = fitz.Rect(last_cell)
                        page.insert_textbox(
                            fitz.Rect(r_annot.x0, r_annot.y1 - 10, r_annot.x1, r_annot.y1),
                            "*See attached sheet",
                            fontsize=7.0,
                            fontname="helv",
                            color=BLUE_COLOR
                        )
        except Exception:
            pass

    for page in doc:
        fill_table_safely(page, "학교 정보", qa_data.get("education_history", []), ["from", "to", "field_degree", "school_name", "location"])
        fill_table_safely(page, "직장 정보", qa_data.get("work_history", []), ["from", "to", "job_title", "company_name", "location"])
        fill_table_safely(page, "여행 정보", qa_data.get("travel_history", []), ["from", "to", "city", "country", "purpose"])

    doc.need_appearances(True)
    out_buf = io.BytesIO()
    doc.save(out_buf, deflate=True)
    doc.close()
    return out_buf.getvalue()

# ==========================================
# 4. Word 질문지 채우기 엔진 (.docx)
# ==========================================
def fill_docx_questionnaire(target_bytes, qa_data):
    doc = docx.Document(io.BytesIO(target_bytes))
    BLUE_RGB = RGBColor(0x1F, 0x4E, 0x8C)
    app = qa_data.get("applicant", {})

    fill_map = {
        "출생 도시": app.get("birth_city", ""),
        "출생 국가": app.get("birth_country", ""),
        "이메일 주소": app.get("email", ""),
        "연락처": app.get("phone", ""),
        "현재 거주지 주소": app.get("address", ""),
        "성:": app.get("spouse_surname", ""),
        "이름:": app.get("spouse_given_name", ""),
        "사실혼이 시작된 날짜": app.get("marriage_date", ""),
        "배우자 생년월일": app.get("spouse_dob", "")
    }

    def insert_run(p, text):
        run = p.add_run(f" {text}")
        run.font.color.rgb = BLUE_RGB
        run.bold = True

    for table in doc.tables:
        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                c_text = cell.text.strip().lower()
                for q_label, ans in fill_map.items():
                    if ans and q_label.lower() in c_text and len(c_text) < len(q_label) + 20:
                        if c_idx + 1 < len(row.cells) and not row.cells[c_idx + 1].text.strip():
                            p = row.cells[c_idx + 1].paragraphs[0] if row.cells[c_idx + 1].paragraphs else row.cells[c_idx + 1].add_paragraph()
                            insert_run(p, ans)
                        elif "___" in cell.text or cell.text.endswith(":"):
                            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                            insert_run(p, ans)

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()

# ==========================================
# 5. 별첨 Word 부속 서류 자동 생성 엔진
# ==========================================
def create_appendix_docx(qa_data):
    """질문지 표 칸을 초과한 전체 학력 및 경력 내역을 담은 공식 Word 부속 서류 생성"""
    doc = docx.Document()
    
    # 여백 설정 (Narrow 0.5인치)
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(0.6)
        s.bottom_margin = Inches(0.6)
        s.left_margin = Inches(0.6)
        s.right_margin = Inches(0.6)

    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(9.5)

    app = qa_data.get("applicant", {})
    client_name = app.get("full_name", "APPLICANT")
    client_dob = app.get("dob", "")

    # 타이틀
    h1 = doc.add_paragraph()
    r1 = h1.add_run("ADDITIONAL INFORMATION SHEET")
    r1.bold = True
    r1.font.size = Pt(14)
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph()
    sub.add_run(f"Applicant Name: {client_name}   |   Date of Birth: {client_dob}").bold = True
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    def set_cell_background(cell, fill_hex):
        tcPr = cell._element.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_hex)
        tcPr.append(shd)

    # 1. 학력 정보 표
    edu_list = qa_data.get("education_history", [])
    if edu_list:
        p_edu = doc.add_paragraph()
        p_edu.add_run("1. Full Education History").bold = True
        p_edu.paragraph_format.space_after = Pt(2)

        t_edu = doc.add_table(rows=1, cols=5)
        t_edu.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_edu.style = 'Table Grid'
        headers = ["From (YYYY-MM)", "To (YYYY-MM)", "Field of Study / Degree", "School Name", "City / Country"]
        
        hdr_cells = t_edu.rows[0].cells
        for idx, text in enumerate(headers):
            hdr_cells[idx].text = text
            hdr_cells[idx].paragraphs[0].runs[0].bold = True
            set_cell_background(hdr_cells[idx], "EAECEE")

        for item in edu_list:
            row_cells = t_edu.add_row().cells
            row_cells[0].text = item.get("from", "")
            row_cells[1].text = item.get("to", "")
            row_cells[2].text = item.get("field_degree", "")
            row_cells[3].text = item.get("school_name", "")
            row_cells[4].text = item.get("location", "")

        doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 2. 직장 정보 표
    work_list = qa_data.get("work_history", [])
    if work_list:
        p_work = doc.add_paragraph()
        p_work.add_run("2. Full Work / Activity History").bold = True
        p_work.paragraph_format.space_after = Pt(2)

        t_work = doc.add_table(rows=1, cols=5)
        t_work.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_work.style = 'Table Grid'
        headers_work = ["From (YYYY-MM)", "To (YYYY-MM)", "Occupation / Title", "Company / School Name", "City / Country"]
        
        hdr_w = t_work.rows[0].cells
        for idx, text in enumerate(headers_work):
            hdr_w[idx].text = text
            hdr_w[idx].paragraphs[0].runs[0].bold = True
            set_cell_background(hdr_w[idx], "EAECEE")

        for item in work_list:
            row_cells = t_work.add_row().cells
            row_cells[0].text = item.get("from", "")
            row_cells[1].text = item.get("to", "")
            row_cells[2].text = item.get("job_title", "")
            row_cells[3].text = item.get("company_name", "")
            row_cells[4].text = item.get("location", "")

        doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 3. 여행 정보 표
    travel_list = qa_data.get("travel_history", [])
    if travel_list:
        p_trv = doc.add_paragraph()
        p_trv.add_run("3. Travel History").bold = True
        p_trv.paragraph_format.space_after = Pt(2)

        t_trv = doc.add_table(rows=1, cols=5)
        t_trv.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_trv.style = 'Table Grid'
        headers_trv = ["From (YYYY-MM-DD)", "To (YYYY-MM-DD)", "City", "Country", "Purpose of Travel"]
        
        hdr_t = t_trv.rows[0].cells
        for idx, text in enumerate(headers_trv):
            hdr_t[idx].text = text
            hdr_t[idx].paragraphs[0].runs[0].bold = True
            set_cell_background(hdr_t[idx], "EAECEE")

        for item in travel_list:
            row_cells = t_trv.add_row().cells
            row_cells[0].text = item.get("from", "")
            row_cells[1].text = item.get("to", "")
            row_cells[2].text = item.get("city", "")
            row_cells[3].text = item.get("country", "")
            row_cells[4].text = item.get("purpose", "")

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()

# ==========================================
# 6. UI 화면 구성 (단독 도구)
# ==========================================
st.title("📝 질문지 자동 이어채우기")
st.caption("기존 손님 서류들과 새 빈 질문지를 올리시면, 알고 있는 정보를 표 칸에 맞게 자동 기입하고 초과된 내역은 별도 Word 양식으로 생성해 드립니다.")

c1, c2 = st.columns(2)
with c1:
    st.subheader("1. 손님의 기존 자료들")
    source_files = st.file_uploader(
        "이전 질문지, 여권, 퍼밋, 이력서 등 (복수 선택 가능)",
        type=['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'heic'],
        accept_multiple_files=True
    )

with c2:
    st.subheader("2. 새로 채울 빈 질문지")
    new_q_file = st.file_uploader(
        "대상 질문지 1부 (.pdf 또는 .docx)",
        type=['pdf', 'docx']
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
            with st.spinner("서류를 대조하여 질문지 칸에 맞게 작성 중입니다. 잠시만 기다려 주세요..."):
                qa_data = analyze_and_extract_qa(source_files, new_q_file)
                
                if qa_data:
                    base_name = os.path.splitext(new_q_file.name)[0]
                    is_docx = new_q_file.name.lower().endswith(".docx")
                    
                    app_info = qa_data.get("applicant", {})
                    client_name = app_info.get("full_name", "Client").replace(" ", "_")

                    # 1. 질문지 본문 작성
                    if is_docx:
                        filled_bytes = fill_docx_questionnaire(new_q_file.getvalue(), qa_data)
                        q_filename = f"{base_name}_완성본.docx"
                        q_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    else:
                        filled_bytes = fill_pdf_questionnaire(new_q_file.getvalue(), qa_data)
                        q_filename = f"{base_name}_완성본.pdf"
                        q_mime = "application/pdf"

                    st.success("✅ 작성이 완료되었습니다! 아래 버튼을 눌러 파일을 내려받으세요.")
                    
                    # 2. 다운로드 버튼 제공
                    dl_col1, dl_col2 = st.columns(2)
                    with dl_col1:
                        st.download_button(
                            label=f"📥 완성된 질문지 내려받기 ({q_filename})",
                            data=filled_bytes,
                            file_name=q_filename,
                            mime=q_mime,
                            type="primary",
                            use_container_width=True
                        )

                    # 3. 학력/직장 이력이 2건 초과인 경우에만 별첨 생성
                    edu_count = len(qa_data.get("education_history", []))
                    work_count = len(qa_data.get("work_history", []))
                    
                    # 💡 PDF 질문지 칸을 초과한 정보가 있을 때만 별첨 생성 버튼 표시
                    if edu_count > 2 or work_count > 2:
                        appendix_bytes = create_appendix_docx(qa_data)
                        appendix_filename = f"{client_name}_Additional_History.docx"
                        with dl_col2:
                            st.download_button(
                                label=f"📄 [별첨] 전체 학력·경력 내역서 내려받기",
                                data=appendix_bytes,
                                file_name=appendix_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="secondary",
                                use_container_width=True
                            )

with col_reset:
    if st.button("🔄 전체 리셋", use_container_width=True):
        st.rerun()
