import streamlit as st

# 📌 Streamlit 페이지 설정
st.set_page_config(page_title="CanNest 질문지 자동 완성 도구", layout="wide")

import google.generativeai as genai
import fitz  # PyMuPDF
from PIL import Image, ImageOps, ImageDraw
import pillow_heif
import io
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor

try:
    import docx
    from docx.shared import RGBColor
except ImportError:
    pass

pillow_heif.register_heif_opener()
Image.MAX_IMAGE_PIXELS = None

# ==========================================
# 0. 보안 및 API 설정
# ==========================================
if "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Cloud의 Secrets에 GEMINI_API_KEY 설정이 필요합니다.")
    st.stop()

GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)

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
# 1. 문서 파싱 및 Gemini 전달 준비
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
            return [f"\n--- [Document: {file_name}] ---\n{text[:25000]}\n"]
            
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
                return [f"\n--- [Document: {file_name}] ---\n{text[:25000]}\n"]
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
# 2. AI 질문지 분석 및 매핑 엔진
# ==========================================
def analyze_and_extract_qa(client_files, target_file, staff_notes=""):
    prompt = """
    You are an AI assistant specialized in auto-filling Canadian immigration questionnaires.
    
    [INPUT INSTRUCTIONS]
    1. The FIRST attached document is the NEW BLANK QUESTIONNAIRE.
    2. The REMAINING attached documents are CLIENT SOURCE MATERIALS (past questionnaires, passports, permits, certificates, resumes, etc.).

    [STRICT RULES]
    1. Match all facts from the client materials and map them into the questions asked in the new questionnaire.
    2. Only use proven facts from the materials. If an item is NOT mentioned or cannot be confirmed, simply EXCLUDE it.
    3. For personal details, use standard Canadian immigration formatting (Names in Title Case, Dates in YYYY-MM-DD or YYYY-MM).
    4. For single-choice checkboxes (e.g. Yes/No, Marital Status), identify the exact choice label that should be marked.

    Return ONLY a raw JSON object with this exact schema:
    {
      "key_values": [
        {
          "page": 4,
          "label": "출생 도시",
          "value": "Danyang",
          "type": "text"
        },
        {
          "page": 4,
          "label": "기혼",
          "value": "V",
          "type": "check"
        }
      ],
      "table_rows": [
        {
          "page": 5,
          "table_keyword": "학교 정보",
          "rows": [
            ["2016-06", "2018-03", "고등학교", "Jacinto Catholic School", "San Jacinto, Philippines"]
          ]
        }
      ]
    }
    """
    contents = [prompt]
    mime_target = target_file.type if target_file.type else "application/pdf"
    contents.extend(prepare_document_for_gemini(target_file.getvalue(), mime_target, f"[NEW_BLANK_FORM] {target_file.name}"))
    contents.extend(batch_process_client_files(client_files))
    if staff_notes.strip():
        contents.append(f"\n--- [Staff Notes] ---\n{staff_notes.strip()}\n")

    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"질문지 분석 중 오류가 발생했습니다: {e}")
        return None

# ==========================================
# 3. PDF 질문지 채우기 엔진 (Flat PDF + AcroForm 지원)
# ==========================================
def fill_pdf_questionnaire(target_bytes, qa_data):
    doc = fitz.open(stream=target_bytes, filetype="pdf")
    fontfile = "cour.ttf" if os.path.exists("cour.ttf") else None
    
    key_values = qa_data.get("key_values", [])
    table_rows = qa_data.get("table_rows", [])

    # 1. 단일 항목 및 체크박스 기입
    for item in key_values:
        p_idx = int(item.get("page", 1)) - 1
        if p_idx < 0 or p_idx >= len(doc):
            continue
        page = doc[p_idx]
        label = item.get("label", "").strip()
        val = str(item.get("value", "")).strip()
        itype = item.get("type", "text")

        if not label or not val:
            continue

        # 대화형 서식(Widget) 필드가 존재하는 경우 우선 매핑
        filled_widget = False
        for widget in page.widgets():
            w_name = widget.field_name.lower() if widget.field_name else ""
            if label.lower() in w_name:
                widget.field_value = val
                widget.update()
                filled_widget = True
                break
        if filled_widget:
            continue

        # 일반 평탄화(Flat) PDF인 경우 텍스트 위치 탐색
        rects = page.search_for(label)
        if not rects:
            # 특수문자나 공백 제외하고 재시도
            clean_lbl = re.sub(r'[\s\:\(\)\_\-\?]+', '', label)
            if len(clean_lbl) >= 2:
                rects = page.search_for(clean_lbl[:4])
        
        if rects:
            r = rects[0]
            if itype == "check":
                # 체크박스 기호(☐) 자리 또는 텍스트 바로 앞쪽에 V 기입
                check_rect = fitz.Rect(r.x0 - 15, r.y0 - 1, r.x0 - 1, r.y1 + 1)
                page.insert_textbox(check_rect, "V", fontsize=10, fontname="helv", color=(0.12, 0.31, 0.55), align=1)
            else:
                # 표 안의 셀인지 감지
                cell_box = None
                try:
                    tables = page.find_tables()
                    for tab in tables:
                        for row in tab.rows:
                            for c_idx, c_rect in enumerate(row.cells):
                                if c_rect and fitz.Rect(c_rect).intersects(r):
                                    if c_idx + 1 < len(row.cells) and row.cells[c_idx + 1]:
                                        cell_box = fitz.Rect(row.cells[c_idx + 1])
                                        break
                except Exception:
                    pass

                if cell_box:
                    # 표의 인접 셀 안에 입력
                    target_rect = fitz.Rect(cell_box.x0 + 4, cell_box.y0 + 2, cell_box.x1 - 4, cell_box.y1 - 2)
                else:
                    # 일반 문장형 레이블 우측 공란에 입력
                    target_rect = fitz.Rect(r.x1 + 6, r.y0 - 1, min(r.x1 + 350, page.rect.width - 20), r.y1 + 3)

                fs = 9.0
                if fontfile and target_rect.width > 0:
                    try:
                        f = fitz.Font(fontfile=fontfile)
                        t_len = f.text_length(val, fontsize=fs)
                        if t_len > target_rect.width:
                            fs = max(5.5, min(9.0, target_rect.width / (f.text_length(val, fontsize=1) or 1)))
                    except Exception:
                        pass

                page.insert_textbox(
                    target_rect,
                    val,
                    fontsize=fs,
                    fontfile=fontfile,
                    fontname="Cour" if fontfile else "courier",
                    color=(0.12, 0.31, 0.55),
                    align=0
                )

    # 2. 다중 행 표(학력, 경력, 여행 등) 기입
    for t_item in table_rows:
        p_idx = int(t_item.get("page", 1)) - 1
        if p_idx < 0 or p_idx >= len(doc):
            continue
        page = doc[p_idx]
        keyword = t_item.get("table_keyword", "")
        rows_data = t_item.get("rows", [])
        
        if not rows_data:
            continue

        try:
            tables = page.find_tables()
            matched_table = None
            for tab in tables:
                tab_text = page.get_text("text", clip=tab.bbox)
                if keyword in tab_text or any(keyword[:3] in tab_text for _ in [0]):
                    matched_table = tab
                    break
            
            if matched_table and len(matched_table.rows) > 1:
                # 헤더 다음 빈 행부터 순서대로 기입
                start_row_idx = 1
                for d_idx, data_row in enumerate(rows_data):
                    target_r_idx = start_row_idx + d_idx
                    if target_r_idx >= len(matched_table.rows):
                        break
                    row_cells = matched_table.rows[target_r_idx].cells
                    for col_idx, text_val in enumerate(data_row):
                        if col_idx < len(row_cells) and row_cells[col_idx]:
                            c_rect = fitz.Rect(row_cells[col_idx])
                            pad_rect = fitz.Rect(c_rect.x0 + 3, c_rect.y0 + 2, c_rect.x1 - 3, c_rect.y1 - 2)
                            page.insert_textbox(
                                pad_rect,
                                str(text_val),
                                fontsize=8.0,
                                fontfile=fontfile,
                                fontname="Cour" if fontfile else "courier",
                                color=(0.12, 0.31, 0.55),
                                align=0
                            )
        except Exception:
            pass

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
    key_values = qa_data.get("key_values", [])
    
    # 빠른 조회를 위한 매핑 테이블
    fill_map = {item['label'].strip().lower(): str(item['value']).strip() for item in key_values if item.get('label') and item.get('value')}

    def insert_run(p, text):
        run = p.add_run(f" {text}")
        run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x8C)
        run.bold = True

    # 1. 표(Table) 내부 질문 검색 후 빈 셀에 기입
    for table in doc.tables:
        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                c_text = cell.text.strip().lower()
                for q_label, ans in fill_map.items():
                    if q_label in c_text and len(c_text) < len(q_label) + 25:
                        if c_idx + 1 < len(row.cells) and not row.cells[c_idx + 1].text.strip():
                            p = row.cells[c_idx + 1].paragraphs[0] if row.cells[c_idx + 1].paragraphs else row.cells[c_idx + 1].add_paragraph()
                            insert_run(p, ans)
                        elif "___" in cell.text or cell.text.endswith(":"):
                            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                            insert_run(p, ans)

    # 2. 본문 문단(Paragraph) 빈칸 채우기
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
# 5. UI 화면 구성 (5번 기능 단독)
# ==========================================
st.title("📝 질문지 자동 이어채우기")
st.caption("손님의 기존 서류와 새 질문지를 업로드하면, 알고 있는 정보를 새 질문지에 자동 기입하여 완성본을 만들어 드립니다.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 손님 기존 보유 자료")
    client_source_files = st.file_uploader(
        "이전 질문지, 여권, 퍼밋, 이력서 등 (복수 선택 가능)",
        type=['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'heic'],
        accept_multiple_files=True
    )
    staff_notes = st.text_area("담당자 메모 (선택)", placeholder="예: 2026년 3월 이사 완료 등 특이사항", height=70)

with col2:
    st.subheader("2. 새로 채워야 할 빈 질문지")
    target_q_file = st.file_uploader(
        "대상 질문지 1부 (.docx 또는 .pdf)",
        type=['docx', 'pdf']
    )

if st.button("질문지 자동 채우기 시작", type="primary", use_container_width=True):
    if not client_source_files:
        st.warning("1단계: 손님의 기존 자료를 1개 이상 올려주세요.")
    elif not target_q_file:
        st.warning("2단계: 새로 채울 질문지 파일을 올려주세요.")
    else:
        with st.spinner("서류를 대조하여 질문지를 채우는 중입니다. 잠시만 기다려 주세요..."):
            qa_plan = analyze_and_extract_qa(client_source_files, target_q_file, staff_notes)
            
            if qa_plan:
                fname_lower = target_q_file.name.lower()
                base_name = os.path.splitext(target_q_file.name)[0]
                
                # Word 서식 처리
                if fname_lower.endswith(".docx"):
                    filled_bytes = fill_docx_questionnaire(target_q_file.getvalue(), qa_plan)
                    download_name = f"{base_name}_자동완성.docx"
                    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                # PDF 서식 처리 (Flat PDF 및 AcroForm 완벽 지원)
                else:
                    filled_bytes = fill_pdf_questionnaire(target_q_file.getvalue(), qa_plan)
                    download_name = f"{base_name}_자동완성.pdf"
                    mime = "application/pdf"
                
                st.markdown("---")
                st.success("✅ 질문지 자동 작성이 완료되었습니다! 아래 버튼을 눌러 내려받으세요.")
                st.download_button(
                    label=f"📥 {download_name} 다운로드",
                    data=filled_bytes,
                    file_name=download_name,
                    mime=mime,
                    type="primary",
                    use_container_width=True
                )
