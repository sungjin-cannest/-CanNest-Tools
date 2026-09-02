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

    Return ONLY a raw JSON object with this format (without markdown block or backticks):
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

    try:
        # ✅ 모델명을 올바른 최신 모델(gemini-1.5-flash)로 변경했습니다!
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        # JSON 응답을 강제하기 위한 설정 (안정성 향상)
        response = model.generate_content(
            [prompt, img_data],
            generation_config=genai.types.GenerationConfig(
                response_mime_type="application/json"
            )
        )
        
        data = json.loads(response.text)
        
        filename = data.get("suggested_filename", "미분류_서류")
        
        if not (filename.lower().endswith(".pdf") or filename.lower().endswith(".jpg") or filename.lower().endswith(".jpeg")):
            filename += ".pdf"
            
        data["suggested_filename"] = filename
        return data
        
    except Exception as e:
        # ✅ 에러를 숨기지 않고 화면에 띄워줍니다.
        st.error(f"AI 분석 중 에러 발생 ({original_filename}): {e}")
        base_name = os.path.splitext(original_filename)[0]
        return {"client_name": "고객명", "doc_category": "기타", "suggested_filename": f"{base_name}.pdf"}
