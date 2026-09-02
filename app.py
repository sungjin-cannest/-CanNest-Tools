import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
from PIL import Image
import io
import zipfile
import json
import os
import datetime

# ==========================================
# 0. 기본 설정 (반드시 코드 맨 윗줄에 1번만 있어야 함!)
# ==========================================
st.set_page_config(page_title="[DEV] CRM 서류 판별 및 압축", layout="wide")

if "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Secrets에 GEMINI_API_KEY가 필요합니다.")
    st.stop()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

# ==========================================
# 1. CRM 서류 판별 로직 (AI 모델 자동 탐색 적용)
# ==========================================
def analyze_document_with_crm_rules(file_bytes, mime_type, original_filename):
    prompt = """
    You are an expert AI document classifier for a Canadian immigration firm.
    Read the provided document carefully and generate an EXACT filename according to our STRICT internal CRM rules.

    [CRITICAL NAMING RULES]
    1. Client Name:
       - Korean client: Full Name in Korean with NO SPACES (e.g., 고주하, 홍길동, 김영미).
       - Non-Korean client: STRICTLY the VERY FIRST WORD of their First Name in Title Case. (e.g., If the name is "Jose Tover Fernando", use "Jose". If "Maria Luisa", use "Maria").
    2. Dates:
       - Format MUST be YYYY.MM.DD (e.g., 2022.01.02).
       - Date ranges: Use a hyphen (e.g., 2022.04.22-2022.05.22).
       - Year ONLY where specified in the manual (e.g., COI, Emedical, Bank Statement).
    3. Delimiter: Always use underscore '_' between Name, Category, Details, and Dates.
    4. Unlisted Documents (CRITICAL): If the document does NOT match any category in the manual list below, look at the TOP of the document to find its exact title in English. 
       - Format for unlisted: {Name}_{ExactEnglishTitleAtTheTop}
       - Example: If the document is a tax assessment, use "{Name}_Notice of Assessment".

    [MANUAL CATEGORY & FORMAT SPECIFICATIONS]
    - Digital Photo / Passport Photo: {Name}_Digital Photo.jpg (MUST use .jpg extension)
    - Passport: {Name}_PP_{ExpiryDate YYYY.MM.DD}
    - Work Permit: {Name}_WP_{ExpiryDate YYYY.MM.DD}
    - Study Permit: {Name}_SP_{ExpiryDate YYYY.MM.DD}
    - Visitor Record: {Name}_VR_{ExpiryDate YYYY.MM.DD}
    - Post-Grad Work Permit: {Name}_PGWP_{ExpiryDate YYYY.MM.DD}
    - Bridging Open Work Permit: {Name}_BOWP_{ExpiryDate YYYY.MM.DD}
    - Coop Permit: {Name}_Coop_{ExpiryDate YYYY.MM.DD}
    - Questionnaire: {Name}_QA_{Type e.g. WP/EE/PR Card/PNP/Spouse WP/VR}_{ReceivedDate YYYY.MM.DD}
    - Police Certificate: {Name}_Police Cert_{CountryNameInEnglish}
    - LOE / Employment Letter: {Name}_LOE_{CompanyInEnglish}_{한글/영문/공증}
    - Paystub: {Name}_Paystub_{CompanyInEnglish}_{StartDate YYYY.MM.DD-EndDate YYYY.MM.DD}
    - Degree/Diploma: {Name}_{Diploma/Bachelor/Highschool/Master/Certificate}_{SchoolName}
    - WES: {Name}_WES
    - Certificate of Income: {Name}_COI_{Year YYYY}
    - Language Test: {Name}_{IELTS/CELPIP}_{TestDate YYYY.MM.DD}
    - Resume: {Name}_Resume_{ReceivedDate YYYY.MM.DD}
    - Medical Exam: {Name}_Emedical_{Year YYYY}
    - Marriage Cert: {Name}_Marriage Cert_{IssueDate YYYY.MM.DD}
    - Transcript: {Name}_Transcript_{SchoolName}
    - Bank Statement: {Name}_Bank Statement_{Year YYYY}
    - Family Cert: {Name}_Family Cert_{IssueDate YYYY.MM.DD}
    - Basic Cert: {Name}_Basic Cert
    - Birth Cert: {Name}_Birth Cert
    - Travel Consent: {ChildName}_Travel Consent
    - ECE License: {Name}_ECE License_{Province e.g. BC/SK}
    - LOA (Letter of Acceptance): {Name}_LOA_{SchoolName}
    - Tuition Receipt: {Name}_Tuition Receipt_{SchoolName}
    - Confirmation of Enrollment: {Name}_Confirmation of Enrollment_{SchoolName}

    Return ONLY a raw JSON object with this format (no markdown blocks, no formatting):
    {
        "client_name": "Extracted formatted name (e.g., 고주하 or Jose)",
        "doc_category": "Category from manual OR exact English title",
        "suggested_filename": "Full_Generated_Filename_With_Correct_Extension"
    }
    """
    
    img_data = None
    if "pdf" in mime_type.lower():
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=70)
            img_data = {"mime_type": "image/jpeg", "data": buf.getvalue()}
            doc.close()
        except Exception as e:
            st.error(f"PDF 이미지 변환 실패 ({original_filename}): {e}")
            return {"client_name": "고객명", "doc_category": "기타", "suggested_filename": f"미분류_{original_filename}"}
    else:
        img_data = {"mime_type": mime_type, "data": file_bytes}

    # 404 에러 방지를 위한 5중 모델 안전망
    candidate_models = [
        'gemini-1.5-flash-latest', 
        'gemini-1.5-flash', 
        'gemini-1.5-pro-latest', 
        'gemini-1.5-pro',
        'gemini-pro-vision'
    ]
    
    response = None
    last_error = None
    
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content([prompt, img_data])
            break # 성공하면 즉시 루프 탈출
        except Exception as e:
            last_error = e
            if "404" in str(e).lower() or "not found" in str(e).lower():
                continue # 404 에러면 다음 모델로 넘어감
            else:
                break # 다른 치명적 에러면 루프 중단

    try:
        if response is None:
            raise Exception(f"사용 가능한 모델이 없습니다. 마지막 에러: {last_error}")
            
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        data = json.loads(clean_text)
        
        filename = data.get("suggested_filename", "미분류_서류")
        
        if not (filename.lower().endswith(".pdf") or filename.lower().endswith(".jpg") or filename.lower().endswith(".jpeg")):
            filename += ".pdf"
            
        data["suggested_filename"] = filename
        return data
        
    except Exception as e:
        st.error(f"AI 분석 중 에러 발생 ({original_filename}): {e}")
        base_name = os.path.splitext(original_filename)[0]
        return {"client_name": "고객명", "doc_category": "기타", "suggested_filename": f"{base_name}.pdf"}

# ==========================================
# 2. 파일 압축 변환 엔진 (PDF / JPG 분기)
# ==========================================
def process_and_compress_file(file_bytes, mime_type, target_filename):
    is_jpeg = target_filename.lower().endswith(('.jpg', '.jpeg'))
    
    if is_jpeg:
        img = Image.open(io.BytesIO(file_bytes))
        if img.mode != "RGB":
            img = img.convert("RGB")
            
        if img.width > 2400:
            ratio = 2400 / img.width
            img = img.resize((2400, int(img.height * ratio)), Image.Resampling.LANCZOS)
            
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85, optimize=True)
        buf.seek(0)
        return buf.getvalue(), "image/jpeg"
        
    else:
        output_pdf = io.BytesIO()
        target_dpi = 150
        quality = 65
        
        if "pdf" in mime_type.lower():
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            new_doc = fitz.open()
            
            for page in doc:
                zoom = target_dpi / 72.0
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                img_buf = io.BytesIO()
                img.save(img_buf, format="JPEG", quality=quality, optimize=True)
                img_buf.seek(0)
                
                img_doc = fitz.open(stream=img_buf.getvalue(), filetype="jpeg")
                pdf_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
                pdf_page.show_pdf_page(pdf_page.rect, img_doc, 0)
                img_doc.close()
                
            new_doc.save(output_pdf, deflate=True, garbage=4)
            new_doc.close()
            doc.close()
        else:
            img = Image.open(io.BytesIO(file_bytes))
            if img.mode != "RGB":
                img = img.convert("RGB")
                
            if img.width > 1800:
                ratio = 1800 / img.width
                img = img.resize((1800, int(img.height * ratio)), Image.Resampling.LANCZOS)
                
            img_buf = io.BytesIO()
            img.save(img_buf, format="JPEG", quality=quality, optimize=True)
            img_buf.seek(0)
            
            img_doc = fitz.open(stream=img_buf.getvalue(), filetype="jpeg")
            new_doc = fitz.open()
            page = new_doc.new_page(width=img.width * 72 / target_dpi, height=img.height * 72 / target_dpi)
            page.show_pdf_page(page.rect, img_doc, 0)
            new_doc.save(output_pdf)
            new_doc.close()
            img_doc.close()
            
        output_pdf.seek(0)
        return output_pdf.getvalue(), "application/pdf"

# ==========================================
# 3. 화면 UI 구성
# ==========================================
st.title("🧪 [DEV] CRM 규격 서류 자동 판별 & 압축 시스템")
st.caption("손님 서류 및 증명사진을 올려주시면 CRM 규칙 기반 파일명 생성 및 압축 변환을 진행합니다.")

st.info("🔒 **보안 안내**: 업로드된 모든 서류는 메모리에서만 처리되며, 작업 후 하단의 '파기' 버튼을 누르면 즉시 영구 삭제됩니다.")

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None

uploaded_files = st.file_uploader(
    "손님 서류 및 사진 업로드 (복수 선택 가능)", 
    type=['jpg', 'jpeg', 'png', 'pdf'], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🔍 서류 분석 및 파일명 생성", type="primary", use_container_width=True):
        results = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, file in enumerate(uploaded_files):
            status_text.text(f"분석 중 ({idx+1}/{len(uploaded_files)}): {file.name}")
            file_bytes = file.getvalue()
            mime_type = file.type if file.type else "application/pdf"
            
            analysis = analyze_document_with_crm_rules(file_bytes, mime_type, file.name)
            
            results.append({
                "file_obj": file,
                "file_bytes": file_bytes,
                "mime_type": mime_type,
                "original_name": file.name,
                "suggested_filename": analysis.get("suggested_filename", file.name),
                "category": analysis.get("doc_category", "기타"),
                "client_name": analysis.get("client_name", "")
            })
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        status_text.success("✅ 파일 분석 완료! 아래에서 생성된 파일명을 확인 및 수정해 주세요.")
        st.session_state.analysis_results = results

if st.session_state.analysis_results:
    st.markdown("---")
    st.subheader("📋 생성된 파일명 검토 및 수정")
    
    updated_filenames = []
    
    for idx, item in enumerate(st.session_state.analysis_results):
        col1, col2 = st.columns([2, 3])
        with col1:
            st.write(f"**원본**: `{item['original_name']}`")
            st.caption(f"분류: {item['category']} | 명의자: {item['client_name']}")
        with col2:
            new_name = st.text_input(
                f"최종 파일명 #{idx+1}", 
                value=item['suggested_filename'], 
                key=f"fn_{idx}"
            )
            updated_filenames.append(new_name)

    st.markdown("---")
    if st.button("🚀 최종 파일 압축 변환 및 패키징", type="primary", use_container_width=True):
        zip_buffer = io.BytesIO()
        final_outputs = []
        
        with st.spinner("파일 최적화 압축 및 패키징 중입니다..."):
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, item in enumerate(st.session_state.analysis_results):
                    final_name = updated_filenames[idx]
                    
                    compressed_bytes, out_mime = process_and_compress_file(
                        item['file_bytes'], 
                        item['mime_type'], 
                        final_name
                    )
                    
                    orig_kb = len(item['file_bytes']) / 1024
                    comp_kb = len(compressed_bytes) / 1024
                    
                    zip_file.writestr(final_name, compressed_bytes)
                    
                    final_outputs.append({
                        "filename": final_name,
                        "mime": out_mime,
                        "orig_kb": orig_kb,
                        "comp_kb": comp_kb,
                        "bytes": compressed_bytes
                    })
                    
        st.success("🎉 파일 변환 및 ZIP 패키징 완료!")
        
        st.markdown("### 📥 변환 결과 및 개별 다운로드")
        for out in final_outputs:
            c1, c2, c3 = st.columns([4, 2, 2])
            with c1:
                st.write(f"📄 **{out['filename']}**")
            with c2:
                st.caption(f"{out['orig_kb']:.1f} KB ➡️ **{out['comp_kb']:.1f} KB**")
            with c3:
                st.download_button(
                    "다운로드", 
                    data=out['bytes'], 
                    file_name=out['filename'], 
                    mime=out['mime'],
                    key=f"dl_final_{out['filename']}"
                )
                
        zip_buffer.seek(0)
        today_str = datetime.date.today().strftime("%Y%m%d")
        st.download_button(
            "📦 전체 서류 ZIP 다운로드",
            data=zip_buffer,
            file_name=f"CRM_Documents_{today_str}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

    # ==========================================
    # 4. 보안 데이터 파기 버튼
    # ==========================================
    st.markdown("---")
    st.subheader("🗑️ 보안 및 데이터 파기")
    if st.button("작업 완료 및 모든 데이터 즉시 파기", type="secondary", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.success("✅ 모든 서류와 데이터가 메모리에서 안전하게 삭제되었습니다.")
        st.rerun()
