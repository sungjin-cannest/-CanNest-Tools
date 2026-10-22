import streamlit as st

# 📌 페이지 설정
st.set_page_config(page_title="CanNest 질문지 자동 완성", layout="wide")

import google.generativeai as genai
import fitz  # PyMuPDF
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
# 1. 파일 읽기 및 텍스트/이미지 변환
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
            return [f"\n--- [Document: {file_name}] ---\n{text[:30000]}\n"]
            
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
                return [f"\n--- [Document: {file_name}] ---\n{text[:30000]}\n"]
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
# 2. AI 대조 및 답변 매핑
# ==========================================
def analyze_and_extract_qa(client_files, target_file):
    prompt = """
    You are an expert AI assistant that auto-fills Canadian immigration questionnaires.
    
    [INPUT]
    1. The FIRST attached document is the NEW BLANK QUESTIONNAIRE.
    2. The REMAINING attached documents are CLIENT MATERIALS (past questionnaires, passports, permits, certificates, resumes, etc.).

    [STRICT RULES]
    1. Extract all verified facts from the client materials and map them into the questions asked in the new questionnaire.
    2. STRICT LANGUAGE RULE: All answers MUST be in ENGLISH (Latin letters, numbers, standard punctuation only). NEVER use Korean characters in the values.
       - Names: strictly as in Passport (e.g., 'Jihyun Cheung').
       - Marital Status: 'Married', 'Single', 'Divorced', 'Common-Law'.
       - Dates: YYYY-MM-DD or YYYY-MM.
    3. If information for a question does NOT exist in the materials, simply DO NOT include it. Leave it blank.
    4. For Checkboxes (Yes/No, Marital status, etc.), provide the exact text label next to the box that should be checked (e.g. '기혼', 'YES', 'NO', '아니오').

    Return ONLY a raw valid JSON object:
    {
      "text_fields": [
        {"anchor": "출생 도시", "value": "Danyang"},
        {"anchor": "출생 국가", "value": "Republic of Korea"},
        {"anchor": "이메일 주소", "value": "ccamjik2@gmail.com"},
        {"anchor": "연락처", "value": "514-567-5791"},
        {"anchor": "현재 거주지 주소", "value": "#401, 2232 Douglas Rd. Burnaby BC V5C 5A7", "is_long": true},
        {"anchor": "1) 배우자 이름", "sub_fields": [{"anchor": "성:", "value": "Jeong"}, {"anchor": "이름:", "value": "Taekhun"}]},
        {"anchor": "사실혼이 시작된 날짜", "value": "2026-03-28"}
      ],
      "checkboxes": [
        {"anchor": "기혼"},
        {"anchor": "NO"}
      ],
      "table_records": [
        {
          "table_anchor": "학교 정보",
          "rows": [
            [
              {"col_anchor": "From", "value": "2016-06"},
              {"col_anchor": "To", "value": "2018-03"},
              {"col_anchor": "전공", "value": "Senior High School"},
              {"col_anchor": "학교명", "value": "Jacinto Catholic School"},
              {"col_anchor": "도시", "value": "San Jacinto, Philippines"}
            ]
          ]
        },
        {
          "table_anchor": "직장 정보",
          "rows": [
            [
              {"col_anchor": "From", "value": "2019-08"},
              {"col_anchor": "To", "value": "2020-07"},
              {"col_anchor": "직업", "value": "Customer Representative"},
              {"col_anchor": "회사명", "value": "Tech Mahindra Limited"},
              {"col_anchor": "도시", "value": "Quezon City, Philippines"}
            ]
          ]
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
# 3. PDF 질문지 작성 엔진 (위치 탐색 & 삽입)
# ==========================================
def find_anchor_in_doc(doc, text_query):
    """전체 PDF 페이지를 스캔하여 라벨의 위치 사각형과 해당 페이지를 찾습니다."""
    queries = [text_query, text_query.replace(" ", ""), text_query[:4]]
    for q in queries:
        if not q or len(q.strip()) < 2:
            continue
        for page in doc:
            rects = page.search_for(q)
            if rects:
                return page, rects[0]
    return None, None

def fill_pdf_questionnaire(target_bytes, qa_data):
    doc = fitz.open(stream=target_bytes, filetype="pdf")
    BLUE_COLOR = (0.12, 0.31, 0.55)  # #1F4E8C
    
    # 1. 체크박스 마킹
    for cb in qa_data.get("checkboxes", []):
        anchor = cb.get("anchor", "").strip()
        if not anchor:
            continue
        page, rect = find_anchor_in_doc(doc, anchor)
        if page and rect:
            # 체크박스(☐) 사각형 위치 계산 (라벨 왼쪽 약 16pt 지점)
            chk_rect = fitz.Rect(rect.x0 - 16, rect.y0 - 1, rect.x0 - 2, rect.y1 + 1)
            try:
                page.insert_textbox(chk_rect, "V", fontsize=10, fontname="helv", color=BLUE_COLOR, align=1)
            except Exception:
                pass

    # 2. 단일 텍스트 필드 채우기
    for item in qa_data.get("text_fields", []):
        # 복합 필드(성/이름 등) 처리
        sub_fields = item.get("sub_fields")
        if sub_fields:
            for sub in sub_fields:
                s_anchor = sub.get("anchor", "").strip()
                s_val = str(sub.get("value", "")).strip()
                if not s_val:
                    continue
                page, rect = find_anchor_in_doc(doc, s_anchor)
                if page and rect:
                    box = fitz.Rect(rect.x1 + 6, rect.y0 - 2, rect.x1 + 140, rect.y1 + 3)
                    try:
                        page.insert_textbox(box, s_val, fontsize=9.0, fontname="helv", color=BLUE_COLOR, align=0)
                    except Exception:
                        pass
            continue

        anchor = item.get("anchor", "").strip()
        val = str(item.get("value", "")).strip()
        is_long = item.get("is_long", False)

        if not anchor or not val:
            continue

        page, rect = find_anchor_in_doc(doc, anchor)
        if page and rect:
            if is_long:
                # 주소 등 긴 필드는 라벨 우측부터 페이지 여백 끝까지 넉넉하게 지정
                box = fitz.Rect(rect.x1 + 10, rect.y0 - 2, page.rect.width - 40, rect.y1 + 22)
                fs = 8.5
            else:
                # 일반 질문은 라벨 오른쪽 인접 칸에 배치
                box = fitz.Rect(rect.x1 + 10, rect.y0 - 2, min(rect.x1 + 180, page.rect.width - 30), rect.y1 + 3)
                fs = 9.0

            try:
                page.insert_textbox(box, val, fontsize=fs, fontname="helv", color=BLUE_COLOR, align=0)
            except Exception:
                pass

    # 3. 다중 행 표(학력, 경력, 여행 등) 채우기
    for t_record in qa_data.get("table_records", []):
        t_anchor = t_record.get("table_anchor", "")
        rows = t_record.get("rows", [])
        if not rows:
            continue
            
        page, rect = find_anchor_in_doc(doc, t_anchor)
        if page and rect:
            # 표 헤더 기준 Y좌표 시작점
            start_y = rect.y1 + 8
            row_height = 20.0
            
            for r_idx, row_data in enumerate(rows):
                cur_y0 = start_y + (r_idx * row_height)
                cur_y1 = cur_y0 + row_height - 2
                
                for cell in row_data:
                    c_anchor = cell.get("col_anchor", "")
                    c_val = str(cell.get("value", "")).strip()
                    if not c_val:
                        continue
                    
                    # 해당 페이지 내에서 열 헤더 검색
                    c_rects = page.search_for(c_anchor)
                    if c_rects:
                        col_r = c_rects[0]
                        cell_box = fitz.Rect(col_r.x0 - 2, cur_y0, col_r.x1 + 80, cur_y1)
                        try:
                            page.insert_textbox(cell_box, c_val, fontsize=8.0, fontname="helv", color=BLUE_COLOR, align=0)
                        except Exception:
                            pass

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
    
    # 단일 항목 매핑 사전 구성
    fill_map = {}
    for item in qa_data.get("text_fields", []):
        if item.get("sub_fields"):
            for sub in item["sub_fields"]:
                if sub.get("anchor") and sub.get("value"):
                    fill_map[sub["anchor"].strip().lower()] = str(sub["value"]).strip()
        elif item.get("anchor") and item.get("value"):
            fill_map[item["anchor"].strip().lower()] = str(item["value"]).strip()

    def insert_blue_run(p, text):
        run = p.add_run(f" {text}")
        run.font.color.rgb = BLUE_RGB
        run.bold = True

    # 1. 표(Table) 내부 탐색 후 입력
    for table in doc.tables:
        for row in table.rows:
            for c_idx, cell in enumerate(row.cells):
                c_text = cell.text.strip().lower()
                for q_label, ans in fill_map.items():
                    if q_label in c_text and len(c_text) < len(q_label) + 20:
                        # 다음 열 셀이 비어있으면 채움
                        if c_idx + 1 < len(row.cells) and not row.cells[c_idx + 1].text.strip():
                            p = row.cells[c_idx + 1].paragraphs[0] if row.cells[c_idx + 1].paragraphs else row.cells[c_idx + 1].add_paragraph()
                            insert_blue_run(p, ans)
                        elif "___" in cell.text or cell.text.endswith(":"):
                            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                            insert_blue_run(p, ans)

    # 2. 본문 문단(Paragraph) 빈칸 채우기
    for p in doc.paragraphs:
        p_text = p.text.strip().lower()
        for q_label, ans in fill_map.items():
            if q_label in p_text and ("___" in p.text or p.text.strip().endswith(":")):
                insert_blue_run(p, ans)

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()

# ==========================================
# 5. 간결한 UI (단독 도구)
# ==========================================
st.title("📝 질문지 자동 이어채우기")
st.caption("기존 손님 서류들과 새 빈 질문지를 올리시면, 이미 있는 정보를 자동으로 채워 완성된 서류를 바로 내려받을 수 있습니다.")

c1, c2 = st.columns(2)
with c1:
    st.subheader("1. 손님의 기존 자료들")
    source_files = st.file_uploader(
        "이전 질문지, 여권, 퍼밋, 이력서 등 (복수 선택)",
        type=['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png'],
        accept_multiple_files=True
    )

with c2:
    st.subheader("2. 새로 채울 빈 질문지")
    new_q_file = st.file_uploader(
        "채워 넣을 새 질문지 1부 (.pdf 또는 .docx)",
        type=['pdf', 'docx']
    )

st.markdown("---")

if st.button("질문지 자동 완성하기", type="primary", use_container_width=True):
    if not source_files:
        st.warning("1단계: 손님의 기존 서류를 최소 1개 이상 업로드해 주세요.")
    elif not new_q_file:
        st.warning("2단계: 새로 채워야 할 질문지 파일을 업로드해 주세요.")
    else:
        with st.spinner("기존 서류들을 대조하여 새 질문지를 작성 중입니다. 잠시만 기다려 주세요..."):
            qa_data = analyze_and_extract_qa(source_files, new_q_file)
            
            if qa_data:
                base_name = os.path.splitext(new_q_file.name)[0]
                is_docx = new_q_file.name.lower().endswith(".docx")
                
                if is_docx:
                    completed_bytes = fill_docx_questionnaire(new_q_file.getvalue(), qa_data)
                    out_filename = f"{base_name}_완성본.docx"
                    mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    completed_bytes = fill_pdf_questionnaire(new_q_file.getvalue(), qa_data)
                    out_filename = f"{base_name}_완성본.pdf"
                    mime_type = "application/pdf"
                
                st.success("✅ 질문지 작성이 완료되었습니다! 아래 버튼을 눌러 확인해 보세요.")
                st.download_button(
                    label=f"📥 {out_filename} 내려받기",
                    data=completed_bytes,
                    file_name=out_filename,
                    mime=mime_type,
                    type="primary",
                    use_container_width=True
                )
