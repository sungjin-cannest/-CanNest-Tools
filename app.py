import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import json
import io
import os
from PIL import Image
import datetime

# 고해상도 스캔 이미지 용량 제한 해제
Image.MAX_IMAGE_PIXELS = None

# ==========================================
# 0. Secrets 안전 검사 및 보안 비밀번호 설정
# ==========================================
if "APP_PASSWORD" not in st.secrets or "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Streamlit Cloud의 Secrets 설정이 필요합니다.")
    st.info("우측 하단 [Manage app] -> [Settings] -> [Secrets]에 GEMINI_API_KEY와 APP_PASSWORD를 입력해 주세요.")
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

# ==========================================
# 1. API 키 및 모델 호환 자동 검색 설정
# ==========================================
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)

def safe_generate_content(contents):
    """404 모델 에러 발생 시 지원 가능한 모델(2.5-flash, 1.5-flash-latest 등)을 순차 탐색하여 호출"""
    candidate_models = [
        'gemini-2.5-flash', 
        'gemini-1.5-flash-latest', 
        'gemini-1.5-flash', 
        'gemini-2.0-flash'
    ]
    last_error = None
    for model_name in candidate_models:
        try:
            mod = genai.GenerativeModel(model_name)
            response = mod.generate_content(contents)
            return response
        except Exception as e:
            last_error = e
            if "404" in str(e) or "not found" in str(e).lower():
                continue
            else:
                raise e
    raise last_error

# ==========================================
# 2. 공통 및 AI 데이터 추출 함수
# ==========================================
def process_uploaded_file_to_image(file_obj):
    """여권 등 이미지 분석 시 선명도 유지 및 경량화"""
    if file_obj.type == "application/pdf":
        doc = fitz.open(stream=file_obj.read(), filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    else:
        img = Image.open(file_obj)
        if img.mode != "RGB":
            img = img.convert("RGB")
    
    if img.width > 1400:
        ratio = 1400 / img.width
        new_size = (1400, int(img.height * ratio))
        img = img.resize(new_size, Image.Resampling.LANCZOS)
    
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    buf.seek(0)
    return Image.open(buf)

def format_full_name(surname, given_name):
    s = surname.strip()
    g = given_name.strip()
    if not s and not g: return ""
    if not s: return g
    if not g: return s
    return f"{g} {s}"

def prepare_document_for_gemini(file_bytes, mime_type, file_name=""):
    """PDF에서 텍스트만 0.1초 만에 뽑아 속도를 극대화합니다."""
    if "pdf" in mime_type.lower():
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = ""
            for page in doc:
                text += page.get_text("text") + "\n"
            
            if len(text.strip()) > 100:
                return f"\n--- [Document: {file_name}] ---\n{text}\n"
        except:
            pass
    
    return {"mime_type": mime_type, "data": file_bytes}

def extract_imm5476_info(image):
    prompt = """
    Analyze this identity document (passport/permit/visa) carefully.
    Extract the following details into exact JSON structure:
    - surname: Family name in English uppercase
    - given_name: Given names in English uppercase
    - dob: Date of birth in YYYY-MM-DD format
    - uci: UCI numbers only if present (10 digits or 8 digits), else empty string
    Return ONLY raw valid JSON object without markdown or code formatting.
    """
    try:
        response = safe_generate_content([prompt, image])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"정보 추출 오류: {e}")
        return None

def extract_all_passports_batch(has_non_acc, images):
    prompt = f"""
    You are an expert OCR system specialized in international passports.
    I am providing {len(images)} passport image(s) in exact order.
    Order structure:
    - {'Image 1 is the non-accompanying parent passport.' if has_non_acc else 'There is no non-accompanying parent passport provided.'}
    - {'Remaining images (Image ' + ('2' if has_non_acc else '1') + f' to {len(images)}) are accompanying family members (parents or children).' if (len(images) > (1 if has_non_acc else 0)) else 'No family passports provided.'}

    For EACH passport image, extract:
    - surname: Surname / Family name in English uppercase
    - given_name: Given name(s) in English uppercase
    - dob: Date of birth in YYYY-MM-DD format
    - passport_number: Passport number in uppercase alphanumeric
    - gender: Sex of the person, strictly "F" or "M"

    Return ONLY a raw valid JSON object:
    {{
      "non_accompanying_parent": {{
        "surname": "...", "given_name": "...", "dob": "YYYY-MM-DD", "passport_number": "...", "gender": "M"
      }} or null,
      "family_members": [
        {{ "surname": "...", "given_name": "...", "dob": "YYYY-MM-DD", "passport_number": "...", "gender": "F" }}
      ]
    }}
    """
    contents = [prompt] + images
    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"여권 일괄 추출 오류: {e}")
        return None

def extract_case_prep_info(tmpl_bytes, client_files):
    prompt = """
    You are helping immigration case-prep staff. 
    The FIRST document attached is a BLANK reference immigration form (Canadian IRCC "IMM" series) — read it to find every field that requires client-specific data entry.
    The REMAINING attached documents are the client's own materials (intake questionnaire, passport, permit, etc.) — use them as your source of truth for values.

    Return ONLY a raw JSON object in this exact shape:
    {
      "sections": [
        {
          "section": "short section name from the form",
          "fields": [
            { "field": "field label", "value": "value found or empty", "source": "source document or empty" }
          ]
        }
      ]
    }
    Rules: Keep values exact. If missing, set value/source to empty string. Maintain strict order.
    """
    
    contents = [prompt]
    contents.append(prepare_document_for_gemini(tmpl_bytes, "application/pdf", "Blank_IMM_Form.pdf"))
    
    for f in client_files:
        mime = f.type if f.type else "application/pdf"
        contents.append(prepare_document_for_gemini(f.getvalue(), mime, f.name))

    try:
        response = safe_generate_content(contents)
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"서류 정리 오류: {e}")
        return None

def is_minor(dob_str):
    try:
        birth_date = datetime.datetime.strptime(dob_str, "%Y-%m-%d").date()
        today = datetime.date.today()
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        return age < 19
    except:
        return True

def get_preloaded_file(file_names):
    for fn in file_names:
        if os.path.exists(fn):
            return fn
    return None

# ==========================================
# 3. PDF 서식 채우기 로직
# ==========================================
def fill_imm5476(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    target_data = {
        "surname": data.get("surname", ""), "given": data.get("given_name", ""),
        "dob": data.get("dob", ""), "email": data.get("email", ""),
        "uci": data.get("uci", "").replace("-", ""), "signDate": data.get("signDate", "")  
    }
    flags = {"surname": False, "given": False, "dob": False, "email": False, "uci": False}
    date_counter = 0
    
    for page in doc:
        for widget in page.widgets():
            field_name = widget.field_name
            if not field_name: continue
            fname_lower = field_name.lower()
            
            if "family name" in fname_lower and not flags["surname"]:
                widget.field_value = target_data["surname"]; widget.update(); flags["surname"] = True
            elif "given name" in fname_lower and not flags["given"]:
                widget.field_value = target_data["given"]; widget.update(); flags["given"] = True
            elif "date of birth" in fname_lower and not flags["dob"]:
                widget.field_value = target_data["dob"]; widget.update(); flags["dob"] = True
            elif "email" in fname_lower and not flags["email"]:
                widget.field_value = target_data["email"]; widget.update(); flags["email"] = True
            elif ("uci" in fname_lower or "unique client identifier" in fname_lower) and not flags["uci"]:
                widget.field_value = target_data["uci"]; widget.update(); flags["uci"] = True
            elif "date" in fname_lower and "birth" not in fname_lower:
                date_counter += 1
                if date_counter == 1: widget.field_value = target_data["signDate"]; widget.update()

    output_pdf = io.BytesIO()
    doc.save(output_pdf); doc.close(); output_pdf.seek(0)
    return output_pdf

def fill_consent_letter(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    children = data.get("children", [])
    num_pages_needed = max(1, (len(children) + 2) // 3)
    for _ in range(num_pages_needed - 1): doc.insert_pdf(doc, from_page=0, to_page=0)
        
    for page_num in range(num_pages_needed):
        page = doc[page_num]
        page_children = children[page_num * 3 : (page_num + 1) * 3]
        child_widgets = [("Information about travelling children", "yyyymmdd"), ("1_2", "2_2"), ("1_3", "2_3")]
        
        for widget in page.widgets():
            fname = widget.field_name.strip() if widget.field_name else ""
            if not fname: continue
            if fname == "1": widget.field_value = data.get("non_acc_name", ""); widget.update()
            elif fname == "2": widget.field_value = data.get("non_acc_address", ""); widget.update()
            elif fname == "3": widget.field_value = data.get("non_acc_phone", ""); widget.update()
            elif fname == "email": widget.field_value = data.get("non_acc_email", ""); widget.update()
            elif fname == "Check Box1": widget.field_value = "0"; widget.update()
            elif fname == "This child or these children hashave my or our consent to travel with": widget.field_value = data.get("acc_name", ""); widget.update()
            elif fname == "Relationship with Children 1": widget.field_value = data.get("acc_relationship", ""); widget.update()
            elif fname == "Relationship with Children 2": widget.field_value = data.get("acc_passport", ""); widget.update()
            elif fname == "I give my consent for this child to travel to": widget.field_value = "Canada"; widget.update()
            elif fname == "2_4": widget.field_value = data.get("acc_name", ""); widget.update()
            elif fname == "At the following addresses 1": widget.field_value = data.get("trip_address", ""); widget.update()
            elif fname == "At the following addresses 2": widget.field_value = data.get("trip_phone", ""); widget.update()
            elif fname == "email_2": widget.field_value = data.get("trip_email", ""); widget.update()
            elif fname == "yyyymmdd_2": widget.field_value = data.get("sign_date", ""); widget.update()
                
            for idx, (name_key, dob_key) in enumerate(child_widgets):
                if idx < len(page_children):
                    if fname == name_key: widget.field_value = page_children[idx].get("name", ""); widget.update()
                    elif fname == dob_key: widget.field_value = page_children[idx].get("dob", ""); widget.update()

    output_pdf = io.BytesIO()
    doc.save(output_pdf); doc.close(); output_pdf.seek(0)
    return output_pdf

# ==========================================
# 4. Streamlit 네비게이션 및 UI 구성
# ==========================================
st.set_page_config(page_title="CanNest 업무 자동화 툴", layout="centered")

st.sidebar.title("🦅 CanNest Tool")
app_mode = st.sidebar.radio("원하시는 업무 도구를 선택하세요", [
    "🍁 IMM5476 자동 작성", 
    "✈️ 한부모 동의서 자동 작성",
    "📋 이민서류 정보 정리 (Case File Prep)"
])

# ------------------------------------------
# 메뉴 1: IMM5476 자동 작성
# ------------------------------------------
if app_mode == "🍁 IMM5476 자동 작성":
    st.title("🍁 IMM5476 자동 작성 도구")
    if "extracted_5476" not in st.session_state: st.session_state.extracted_5476 = None
    template_5476_bytes = None

    found_5476 = get_preloaded_file(["imm5476_template.pdf", "imm5476_template.pdf.pdf"])
    
    if found_5476:
        st.success("✅ 사내 표준 'IMM5476' 양식이 시스템에서 자동으로 로드되었습니다.")
        with open(found_5476, "rb") as f: template_5476_bytes = f.read()
    else:
        st.error("⚠️ GitHub에 'imm5476_template.pdf' 파일이 없습니다. 우선 수동으로 업로드해주세요.")
        template_file = st.file_uploader("IMM5476 템플릿 PDF 수동 업로드", type=['pdf'], key="template_5476")
        if template_file: template_5476_bytes = template_file.getvalue()

    st.markdown("---")
    client_file = st.file_uploader("1. 손님 여권 또는 퍼밋", type=['jpg', 'jpeg', 'png', 'pdf'], key="client_5476")

    if client_file and st.button("🚀 AI 정보 추출하기", use_container_width=True):
        with st.spinner("서류 분석 중..."):
            extracted = extract_imm5476_info(process_uploaded_file_to_image(client_file))
            if extracted: st.session_state.extracted_5476 = extracted; st.success("성공!")

    if st.session_state.extracted_5476:
        data = st.session_state.extracted_5476
        c1, c2 = st.columns(2)
        with c1:
            surname = st.text_input("성 (Surname)", data.get("surname", ""))
            dob = st.text_input("생년월일", data.get("dob", ""))
            email = st.text_input("이메일", "")
        with c2:
            given = st.text_input("이름 (Given Name)", data.get("given_name", ""))
            uci = st.text_input("UCI", data.get("uci", ""))
            sign_date = st.date_input("서명날짜", datetime.date.today())

        if st.button("문서 생성 및 다운로드", type="primary"):
            if not template_5476_bytes:
                st.error("템플릿 파일이 없습니다.")
            else:
                final_data = {"surname": surname, "given_name": given, "dob": dob, "uci": uci, "email": email, "signDate": sign_date.strftime("%Y-%m-%d")}
                pdf_out = fill_imm5476(template_5476_bytes, final_data)
                st.download_button("📥 다운로드", pdf_out, file_name=f"IMM5476_{surname}_{given}.pdf", mime="application/pdf")

# ------------------------------------------
# 메뉴 2: 한부모 동의서 자동 작성
# ------------------------------------------
elif app_mode == "✈️ 한부모 동의서 자동 작성":
    st.title("✈️ 한부모 동의서 자동 작성 도구")
    if "consent_non_acc" not in st.session_state: st.session_state.consent_non_acc = {}
    if "consent_family" not in st.session_state: st.session_state.consent_family = []

    consent_template_bytes = None
    found_consent = get_preloaded_file(["consent_template.pdf", "consent_template.pdf.pdf"])

    if found_consent:
        st.success("✅ 사내 표준 '한부모 동의서' 양식이 시스템에서 자동으로 로드되었습니다.")
        with open(found_consent, "rb") as f: consent_template_bytes = f.read()
    else:
        st.error("⚠️ GitHub에 'consent_template.pdf' 파일이 없습니다. 우선 수동으로 업로드해주세요.")
        consent_template = st.file_uploader("동의서 양식 수동 업로드", type=['pdf'])
        if consent_template: consent_template_bytes = consent_template.getvalue()

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1: non_acc_file = st.file_uploader("비동반 부모님 여권 (1장)", type=['jpg', 'jpeg', 'png', 'pdf'])
    with c2: family_files = st.file_uploader("동반 부모/자녀 여권", type=['jpg', 'jpeg', 'png', 'pdf'], accept_multiple_files=True)

    if st.button("🚀 일괄 추출하기", type="primary", use_container_width=True):
        images = []
        has_non_acc = bool(non_acc_file)
        if non_acc_file: images.append(process_uploaded_file_to_image(non_acc_file))
        if family_files: images.extend([process_uploaded_file_to_image(f) for f in family_files])
        
        if images:
            with st.spinner("분석 중..."):
                res = extract_all_passports_batch(has_non_acc, images)
                if res:
                    st.session_state.consent_non_acc = res.get("non_accompanying_parent", {}) or {}
                    st.session_state.consent_family = res.get("family_members", []) or []
                    st.success("완료!")

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
    sign_date_str = st.date_input("서명일", datetime.date.today()).strftime("%Y/%m/%d")

    if st.button("문서 생성 및 다운로드", type="primary"):
        if not consent_template_bytes:
            st.error("양식 파일이 없습니다.")
        else:
            data_consent = {
                "non_acc_name": non_acc_name, "non_acc_address": non_acc_address, "non_acc_phone": non_acc_phone, "non_acc_email": non_acc_email,
                "children": final_children, "acc_name": acc_name, "acc_relationship": acc_rel, "acc_passport": acc_passport,
                "trip_address": trip_address, "trip_phone": trip_phone, "trip_email": trip_email, "sign_date": sign_date_str
            }
            pdf_out = fill_consent_letter(consent_template_bytes, data_consent)
            st.download_button("📥 다운로드", pdf_out, f"Consent_{non_acc_name}.pdf", "application/pdf")

# ------------------------------------------
# 메뉴 3: 이민서류 정보 정리 (Case File Prep)
# ------------------------------------------
elif app_mode == "📋 이민서류 정보 정리 (Case File Prep)":
    st.title("📋 이민서류 정보 정리 도구")
    
    if "prep_result" not in st.session_state:
        st.session_state.prep_result = None

    st.subheader("1. 대상 서식 (IMM PDF)")
    
    form_map = {
        "직접 파일 업로드 (기타 서식)": None,
        "IMM1294 (SP-OUTSIDE)": "imm1294.pdf",
        "IMM1295 (WP-OUTSIDE)": "imm1295.pdf",
        "IMM5708 (VR-INSIDE)": "imm5708.pdf",
        "IMM5709 (SP-INSIDE)": "imm5709.pdf",
        "IMM5710 (WP-INSIDE)": "imm5710.pdf"
    }
    
    selected_form = st.selectbox("📌 템플릿 서식 선택 (GitHub 박제본)", list(form_map.keys()))
    
    tmpl_bytes = None
    
    if form_map[selected_form] is not None:
        file_path = form_map[selected_form]
        if os.path.exists(file_path):
            with open(file_path, "rb") as f:
                tmpl_bytes = f.read()
            st.success(f"✅ 서버에 저장된 '{selected_form}' 양식이 자동으로 로드되었습니다.")
        else:
            st.error(f"⚠️ {file_path} 파일이 아직 GitHub에 없습니다. 파일을 업로드해 주세요!")
    else:
        tmpl_prep_file = st.file_uploader("빈 IMM 서식 (반드시 Print to PDF로 평탄화된 파일)", type=['pdf'], key="case_tmpl")
        if tmpl_prep_file:
            tmpl_bytes = tmpl_prep_file.getvalue()

    st.markdown("---")
    st.subheader("2. 손님 제출 서류 (복수 선택 가능)")
    client_prep_files = st.file_uploader("질문지, 여권, 퍼밋 등 서류 올리기", type=['jpg', 'jpeg', 'png', 'pdf'], accept_multiple_files=True, key="case_client_docs")

    if st.button("🚀 서류 읽고 항목별 정보 정리하기", type="primary", use_container_width=True):
        if tmpl_bytes is None:
            st.warning("1번 단계에서 서식이 정상적으로 선택되거나 업로드되지 않았습니다.")
        elif not client_prep_files:
            st.warning("2번 단계에서 손님 서류를 1개 이상 올려주세요.")
        else:
            with st.spinner("⚡ AI가 초고속으로 문서를 읽고 정리 중입니다... (약 10~20초 소요)"):
                res = extract_case_prep_info(tmpl_bytes, client_prep_files)
                if res:
                    st.session_state.prep_result = res
                    st.success("🎉 서류 정보 정리가 완료되었습니다!")

    if st.session_state.prep_result:
        st.markdown("---")
        st.subheader("3. 정리된 정보 결과")

        parsed = st.session_state.prep_result
        sections = parsed.get("sections", [])

        if not sections:
            st.error("서식에서 분석할 항목을 찾지 못했습니다. (파일이 평탄화된 PDF인지 확인하세요)")
        else:
            full_text_list = []
            for sec in sections:
                sec_name = sec.get("section", "기타 항목")
                st.write(f"### 📌 {sec_name}")
                full_text_list.append(f"[{sec_name}]")

                table_data = []
                for f in sec.get("fields", []):
                    field_lbl = f.get("field", "")
                    val = f.get("value", "")
                    src = f.get("source", "")

                    if not val:
                        display_val = "⚠️ 확인 필요 (미발견)"
                        full_text_list.append(f"{field_lbl}: (확인 필요)")
                    else:
                        display_val = val
                        full_text_list.append(f"{field_lbl}: {val}")

                    table_data.append({
                        "항목 (Field)": field_lbl,
                        "추출값 (Value)": display_val,
                        "출처 (Source)": src if src else "-"
                    })

                st.table(table_data)
                full_text_list.append("")

            st.markdown("#### 📋 한눈에 복사하기")
            st.text_area("아래 텍스트를 복사하여 서식에 옮겨 적으세요", value="\n".join(full_text_list), height=250)
            
            if st.button("＋ 새 케이스 정리하기"):
                st.session_state.prep_result = None
                st.rerun()
