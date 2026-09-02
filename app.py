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
    - Digital Photo / Passport Photo: {Name}_Digital Photo.jpg (MUST use .jpg extension, do NOT use .pdf)
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
    - Family Cert (가족관계증명서): {Name}_Family Cert_{IssueDate YYYY.MM.DD}
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
        except Exception:
            return {"client_name": "고객명", "doc_category": "기타", "suggested_filename": f"미분류_{original_filename}"}
    else:
        img_data = {"mime_type": mime_type, "data": file_bytes}

    try:
        model = genai.GenerativeModel('gemini-3.6-flash')
        response = model.generate_content([prompt, img_data])
        clean_text = response.text.strip().replace('```json', '').replace('
