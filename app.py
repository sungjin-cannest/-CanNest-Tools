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
# 0. 설정 및 API 키
# ==========================================
st.set_page_config(page_title="[DEV] CRM 규격 서류 판별 & 압축 툴", layout="wide")

if "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Secrets에 GEMINI_API_KEY가 필요합니다.")
    st.stop()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

# ==========================================
# 1. CRM 매뉴얼 기반 서류 판별 AI 로직
# ==========================================
def analyze_document_with_crm_rules(file_bytes, mime_type, original_filename):
    """
    [매뉴얼] CRM 파일저장 규칙을 준수하여 서류 종류 판별 및 파일명 자동 생성
    """
    prompt = """
    You are an expert AI document classifier for a Canadian immigration firm.
    Analyze the provided document image and extract metadata to generate an EXACT filename according to our internal CRM manual rules.

    [CRM NAMING RULES]
    1. Client Name:
       - Korean client: Full Name in Korean (e.g. 홍길동, 김영미)
       - Non-Korean client: English First Name only in Title Case (e.g. Richard, Pham, Jennifer)
    2. Dates:
       - Format: YYYY.MM.DD (using dots, e.g. 2022.01.02, 2021.11.11).
       - Date ranges: Use hyphen (e.g. 2022.04.22-2022.05.22).
       - Year only where specified (e.g. COI, Emedical, Bank Statement).
    3. Delimiter: Always use underscore '_' between Name, Category, Details, and Dates.
    4. English Capitalization: Capitalize the first letter of each English word (Title Case).

    [CATEGORY & FORMAT SPECIFICATIONS]
    - Passport: {Name}_PP_{ExpiryDate YYYY.MM.DD}
    - Work Permit: {Name}_WP_{ExpiryDate YYYY.MM.DD}
    - Study Permit: {Name}_SP_{ExpiryDate YYYY.MM.DD}
    - Visitor Record: {Name}_VR_{ExpiryDate YYYY.MM.DD}
    - Post-Grad Work Permit: {Name}_PGWP_{ExpiryDate YYYY.MM.DD}
    - Bridging Open Work Permit: {Name}_BOWP_{ExpiryDate YYYY.MM.DD}
    - Coop Permit: {Name}_Coop_{ExpiryDate YYYY.MM.DD}
    - Questionnaire: {Name}_QA_{Type e.g. WP/EE/PR Card/PNP/Spouse WP/VR}_{ReceivedDate YYYY.MM.DD}
    - Police Certificate: {Name}_Police Cert_{CountryNameInEnglish}
    - LOE / Employment Letter: {Name}_LOE_{CompanyInEnglish}_{한글/영문/공증}[_{ReceivedDate YYYY.MM.DD if multiple}]
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
    - Family Cert (가족관계증명서): {Name}_Family Cert_{IssueDate YYYY.MM.DD} (Name = Applicant on cert)
    - Basic Cert (기본증명서): {Name}_Basic Cert
    - Birth Cert (출생증명서): {Name}_Birth Cert
    - Travel Consent: {ChildName}_Travel Consent
    - ECE License: {Name}_ECE License_{Province e.g. BC/SK}
    - LOA: {Name}_LOA_{SchoolName}
    - Tuition Receipt: {Name}_Tuition Receipt_{SchoolName}
    - Confirmation of Enrollment: {Name}_Confirmation of Enrollment_{SchoolName}
    - Other Document: {Name}_{DocEnglishTitle}

    Return ONLY a raw JSON object with this format (without markdown block):
    {
        "client_name": "Extracted Client Name",
        "doc_category": "Detected Category",
        "suggested_filename": "Full_Generated_Filename_Without_Extension"
    }
    """
    
    # PDF일 경우 1페이지를 이미지로 변환하여 AI에 전달
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
        except Exception:
            return {"client_name": "고객명", "doc_category": "기타", "suggested_filename": f"미분류_{original_filename}"}
    else:
        img_data = {"mime_type": mime_type, "data": file_bytes}

    try:
        model = genai.GenerativeModel('gemini-3.6-flash')
        response = model.generate_content([prompt, img_data])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        data = json.loads(clean_text)
        
        # .pdf 확장자 제거 처리
        filename = data.get("suggested_filename", "미분류_서류").replace(".pdf", "").replace(".PDF", "")
        data["suggested_filename"] = f"{filename}.pdf"
        return data
    except Exception:
        base_name = os.path.splitext(original_filename)[0]
        return {"client_name": "고객명", "doc_category": "기타", "suggested_filename": f"{base_name}.pdf"}

# ==========================================
# 2. PDF 압축 및 최적화 엔진 (IRCC 제출용)
# ==========================================
def compress_to_pdf(file_bytes, mime_type, target_dpi=150, quality=65):
    """이미지 및 PDF 파일을 IRCC 규격 고압축 PDF로 변환"""
    output_pdf = io.BytesIO()
    
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
    return output_pdf.getvalue()

# ==========================================
# 3. Streamlit UI
# ==========================================
st.title("🧪 [DEV] CRM 규격 서류 자동 판별 & 압축 시스템")
st.caption("손님에게 받은 파일들을 올려주시면 CRM 저장 규칙에 맞춘 파일명 생성 및 압축 PDF 변환을 진행합니다.")

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None

uploaded_files = st.file_uploader(
    "손님 서류 업로드 (복수 선택 가능)", 
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
            
            # AI 분석 및 파일명 제안
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
    st.info("💡 AI가 생성한 파일명이 CRM 매뉴얼 규격과 맞는지 확인하시고, 필요시 직접 수정할 수 있습니다.")
    
    updated_filenames = []
    
    for idx, item in enumerate(st.session_state.analysis_results):
        col1, col2 = st.columns([2, 3])
        with col1:
            st.write(f"**원본**: `{item['original_name']}`")
            st.caption(f"분석 분류: {item['category']} | 명의자: {item['client_name']}")
        with col2:
            new_name = st.text_input(
                f"최종 파일명 #{idx+1}", 
                value=item['suggested_filename'], 
                key=f"fn_{idx}"
            )
            updated_filenames.append(new_name)

    st.markdown("---")
    if st.button("🚀 최종 PDF 압축 변환 및 패키징", type="primary", use_container_width=True):
        zip_buffer = io.BytesIO()
        final_outputs = []
        
        with st.spinner("PDF 압축 및 파일 패키징 중입니다..."):
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, item in enumerate(st.session_state.analysis_results):
                    final_name = updated_filenames[idx]
                    if not final_name.lower().endswith(".pdf"):
                        final_name += ".pdf"
                        
                    compressed_bytes = compress_to_pdf(item['file_bytes'], item['mime_type'])
                    
                    orig_kb = len(item['file_bytes']) / 1024
                    comp_kb = len(compressed_bytes) / 1024
                    
                    zip_file.writestr(final_name, compressed_bytes)
                    
                    final_outputs.append({
                        "filename": final_name,
                        "orig_kb": orig_kb,
                        "comp_kb": comp_kb,
                        "bytes": compressed_bytes
                    })
                    
        st.success("🎉 PDF 변환 및 ZIP 패키징이 completed!")
        
        # 변환 결과 다운로드 영역
        st.markdown("### 📥 변환 결과 및 다운로드")
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
                    mime="application/pdf",
                    key=f"dl_final_{out['filename']}"
                )
                
        zip_buffer.seek(0)
        today_str = datetime.date.today().strftime("%Y%m%d")
        st.download_button(
            "📦 전체 compressed 서류 ZIP 다운로드",
            data=zip_buffer,
            file_name=f"CRM_Documents_{today_str}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )
