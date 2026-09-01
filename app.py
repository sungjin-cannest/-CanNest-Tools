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
# 1. API 키 및 모델 설정
# ==========================================
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel('gemini-3.6-flash')

# ==========================================
# 2. 공통 및 AI 데이터 추출 함수
# ==========================================
def process_uploaded_file_to_image(file_obj):
    if file_obj.type == "application/pdf":
        doc = fitz.open(stream=file_obj.read(), filetype="pdf")
        page = doc.load_page(0)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    else:
        img = Image.open(file_obj)
    
    if img.width > 1800:
        ratio = 1800 / img.width
        new_size = (1800, int(img.height * ratio))
        img = img.resize(new_size)
    return img

def format_full_name(surname, given_name, style="FIRST_LAST"):
    """영문 성명 표기법 조합 함수"""
    s = surname.strip()
    g = given_name.strip()
    if not s and not g:
        return ""
    if not s:
        return g
    if not g:
        return s
    
    if style == "FIRST_LAST":
        return f"{g} {s}"        # 예: Huja Ko
    elif style == "LAST_COMMA_FIRST":
        return f"{s}, {g}"       # 예: Ko, Huja
    return f"{g} {s}"

def extract_imm5476_info(image):
    prompt = """
    Analyze this document (passport/visa/permit) and extract the following information.
    Return ONLY a valid JSON object without markdown tags.
    Format required:
    {
      "surname": "Family name in English uppercase",
      "given_name": "Given names in English uppercase",
      "dob": "Date of birth in YYYY-MM-DD format",
      "uci": "UCI digits only if present, else empty string"
    }
    """
    try:
        response = model.generate_content([prompt, image])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"정보 추출 오류: {e}")
        return None

def extract_passport_details(image):
    prompt = """
    Analyze this passport document and extract key details.
    Return ONLY a valid JSON object without markdown tags.
    Format required:
    {
      "surname": "Surname in English uppercase",
      "given_name": "Given name in English uppercase",
      "dob": "Date of birth in YYYY-MM-DD format",
      "passport_number": "Passport number uppercase",
      "gender": "F or M"
    }
    """
    try:
        response = model.generate_content([prompt, image])
        clean_text = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(clean_text)
    except Exception as e:
        st.error(f"여권 정보 추출 중 오류: {e}")
        return None

def is_minor(dob_str):
    try:
        birth_date = datetime.datetime.strptime(dob_str, "%Y-%m-%d").date()
        today = datetime.date.today()
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        return age < 19
    except:
        return True

# ==========================================
# 3. PDF 서식 채우기 로직
# ==========================================
def fill_imm5476(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    target_data = {
        "surname": data.get("surname", ""),
        "given": data.get("given_name", ""),
        "dob": data.get("dob", ""),  
        "email": data.get("email", ""),
        "uci": data.get("uci", "").replace("-", ""),  
        "signDate": data.get("signDate", "")  
    }
    flags = {"surname": False, "given": False, "dob": False, "email": False, "uci": False}
    date_counter = 0
    fields_found = False
    
    for page in doc:
        for widget in page.widgets():
            field_name = widget.field_name
            if not field_name: continue
            fields_found = True
            fname_lower = field_name.lower()
            
            if "family name" in fname_lower and not flags["surname"]:
                widget.field_value = target_data["surname"]
                widget.update()
                flags["surname"] = True
            elif "given name" in fname_lower and not flags["given"]:
                widget.field_value = target_data["given"]
                widget.update()
                flags["given"] = True
            elif "date of birth" in fname_lower and not flags["dob"]:
                widget.field_value = target_data["dob"]
                widget.update()
                flags["dob"] = True
            elif "email" in fname_lower and not flags["email"]:
                widget.field_value = target_data["email"]
                widget.update()
                flags["email"] = True
            elif ("uci" in fname_lower or "unique client identifier" in fname_lower) and not flags["uci"]:
                widget.field_value = target_data["uci"]
                widget.update()
                flags["uci"] = True
            elif "date" in fname_lower and "birth" not in fname_lower:
                date_counter += 1
                if date_counter == 1:
                    widget.field_value = target_data["signDate"]
                    widget.update()

    if not fields_found:
        raise Exception("PDF 입력 칸을 찾을 수 없습니다.")

    output_pdf = io.BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf

def fill_consent_letter(template_bytes, data):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    children = data.get("children", [])
    num_children = len(children)
    
    num_pages_needed = max(1, (num_children + 2) // 3)
    for _ in range(num_pages_needed - 1):
        doc.insert_pdf(doc, from_page=0, to_page=0)
        
    for page_num in range(num_pages_needed):
        page = doc[page_num]
        page_children = children[page_num * 3 : (page_num + 1) * 3]
        child_widgets = [
            ("Information about travelling children", "yyyymmdd"),
            ("1_2", "2_2"),
            ("1_3", "2_3")
        ]
        
        for widget in page.widgets():
            fname = widget.field_name.strip() if widget.field_name else ""
            if not fname: continue
            
            if fname == "1":
                widget.field_value = data.get("non_acc_name", "")
                widget.update()
            elif fname == "2":
                widget.field_value = data.get("non_acc_address", "")
                widget.update()
            elif fname == "3":
                widget.field_value = data.get("non_acc_phone", "")
                widget.update()
            elif fname == "email":
                widget.field_value = data.get("non_acc_email", "")
                widget.update()
                
            elif fname == "Check Box1":
                widget.field_value = "0"
                widget.update()
            elif fname == "This child or these children hashave my or our consent to travel with":
                widget.field_value = data.get("acc_name", "")
                widget.update()
            elif fname == "Relationship with Children 1":
                widget.field_value = data.get("acc_relationship", "")
                widget.update()
            elif fname == "Relationship with Children 2":
                widget.field_value = data.get("acc_passport", "")
                widget.update()
                
            elif fname == "I give my consent for this child to travel to":
                widget.field_value = "Canada"
                widget.update()
            elif fname == "1_4":
                widget.field_value = ""
                widget.update()
            elif fname == "2_4":
                widget.field_value = data.get("acc_name", "")
                widget.update()
            elif fname == "At the following addresses 1":
                widget.field_value = data.get("trip_address", "")
                widget.update()
            elif fname == "At the following addresses 2":
                widget.field_value = data.get("trip_phone", "")
                widget.update()
            elif fname == "email_2":
                widget.field_value = data.get("trip_email", "")
                widget.update()
            elif fname == "yyyymmdd_2":
                widget.field_value = data.get("sign_date", "")
                widget.update()
                
            for idx, (name_key, dob_key) in enumerate(child_widgets):
                if idx < len(page_children):
                    if fname == name_key:
                        widget.field_value = page_children[idx].get("name", "")
                        widget.update()
                    elif fname == dob_key:
                        widget.field_value = page_children[idx].get("dob", "")
                        widget.update()

    output_pdf = io.BytesIO()
    doc.save(output_pdf)
    doc.close()
    output_pdf.seek(0)
    return output_pdf

# ==========================================
# 4. Streamlit 네비게이션 및 UI 구성
# ==========================================
st.set_page_config(page_title="CanNest 업무 자동화 툴", layout="centered")

st.sidebar.title("🦅 CanNest Tool")
app_mode = st.sidebar.radio("원하시는 업무 도구를 선택하세요", ["🍁 IMM5476 자동 작성", "✈️ 한부모 동의서 자동 작성"])

# ------------------------------------------
# 메뉴 1: IMM5476 자동 작성
# ------------------------------------------
if app_mode == "🍁 IMM5476 자동 작성":
    st.title("🍁 IMM5476 자동 작성 도구")
    st.write("여권이나 퍼밋 서류를 올리면 AI가 데이터를 추출하여 서식을 채워줍니다.")

    if "extracted_5476" not in st.session_state:
        st.session_state.extracted_5476 = None

    default_5476_path = "imm5476_template.pdf"
    template_5476_bytes = None

    if os.path.exists(default_5476_path):
        st.success("✅ 사내 표준 'IMM5476' 양식이 자동으로 로드되었습니다.")
        with open(default_5476_path, "rb") as f:
            template_5476_bytes = f.read()
    else:
        st.subheader("1. IMM5476 템플릿 업로드")
        template_file = st.file_uploader("IMM5476 템플릿 PDF 선택", type=['pdf'], key="template_5476")
        if template_file:
            template_5476_bytes = template_file.getvalue()

    st.markdown("---")
    st.subheader("1. 손님 여권 또는 퍼밋 (이미지 또는 PDF)")
    client_file = st.file_uploader("여권/퍼밋 선택", type=['jpg', 'jpeg', 'png', 'pdf'], key="client_5476")

    if client_file is not None:
        if st.button("🚀 AI 정보 추출하기", key="btn_5476"):
            with st.spinner("AI가 분석 중입니다..."):
                img = process_uploaded_file_to_image(client_file)
                extracted = extract_imm5476_info(img)
                if extracted:
                    st.session_state.extracted_5476 = extracted
                    st.success("정보 추출 성공!")

    if st.session_state.extracted_5476 is not None:
        st.subheader("2. 정보 확인 및 입력")
        data = st.session_state.extracted_5476
        col1, col2 = st.columns(2)
        with col1:
            surname = st.text_input("성 (Surname)", value=data.get("surname", ""))
            dob = st.text_input("생년월일 (YYYY-MM-DD)", value=data.get("dob", ""))
            email = st.text_input("이메일", value="")
        with col2:
            given = st.text_input("이름 (Given Name)", value=data.get("given_name", ""))
            uci = st.text_input("UCI (있는 경우)", value=data.get("uci", ""))
            sign_date = st.date_input("서명란 날짜", value=datetime.date.today())

        if st.button("문서 생성 및 다운로드", type="primary"):
            if not template_5476_bytes:
                st.error("1번 단계에서 IMM5476 템플릿 PDF를 올리거나, GitHub 저장소에 imm5476_template.pdf 파일을 추가해 주세요!")
            elif not email:
                st.warning("이메일을 입력해 주세요.")
            else:
                final_data = {
                    "surname": surname, "given_name": given, "dob": dob,
                    "uci": uci, "email": email, "signDate": sign_date.strftime("%Y-%m-%d")
                }
                try:
                    pdf_out = fill_imm5476(template_5476_bytes, final_data)
                    st.download_button("📥 완성된 IMM5476 다운로드", pdf_out, file_name=f"IMM5476_{surname}_{given}.pdf", mime="application/pdf")
                    st.balloons()
                except Exception as e:
                    st.error(f"오류 발생: {e}")

# ------------------------------------------
# 메뉴 2: 한부모 동의서 자동 작성
# ------------------------------------------
elif app_mode == "✈️ 한부모 동의서 자동 작성":
    st.title("✈️ 한부모 동의서(Consent Letter) 자동 작성 도구")
    
    if "consent_non_acc" not in st.session_state: st.session_state.consent_non_acc = {}
    if "consent_family" not in st.session_state: st.session_state.consent_family = []

    default_template_path = "consent_template.pdf"
    consent_template_bytes = None

    if os.path.exists(default_template_path):
        st.success("✅ 사내 표준 '한부모 동의서' 양식이 자동으로 로드되었습니다.")
        with open(default_template_path, "rb") as f:
            consent_template_bytes = f.read()
    else:
        st.subheader("1. 동의서 템플릿 PDF 업로드")
        consent_template = st.file_uploader("한부모 동의서 양식 PDF 선택", type=['pdf'], key="consent_tmpl")
        if consent_template:
            consent_template_bytes = consent_template.getvalue()

    st.markdown("---")
    st.subheader("1. 여권 파일 업로드 (한번에 처리)")
    
    # 영문 이름 표기법 선택 옵션
    name_style_choice = st.radio("영문 성명 표기 방식 선택", ["Huja Ko (이름 성)", "Ko, Huja (성, 이름)"], horizontal=True)
    name_style = "FIRST_LAST" if "Huja Ko" in name_style_choice else "LAST_COMMA_FIRST"

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        non_acc_file = st.file_uploader("비동반 부모님 여권 (1장)", type=['jpg', 'jpeg', 'png', 'pdf'], key="non_acc")
    with col_up2:
        family_files = st.file_uploader("동반 부모 및 자녀 여권 (복수 선택)", type=['jpg', 'jpeg', 'png', 'pdf'], accept_multiple_files=True, key="family_files")

    # 모든 여권 일괄 추출 버튼
    if st.button("🚀 모든 여권 정보 한 번에 AI 추출하기", type="primary", use_container_width=True):
        if not non_acc_file and not family_files:
            st.warning("분석할 여권 파일을 1장 이상 올려주세요.")
        else:
            with st.spinner("AI가 업로드된 모든 여권을 한 번에 분석 중입니다..."):
                # 1. 비동반 부모 여권 처리
                if non_acc_file:
                    img_non = process_uploaded_file_to_image(non_acc_file)
                    st.session_state.consent_non_acc = extract_passport_details(img_non) or {}

                # 2. 동반 가족 여권들 처리
                if family_files:
                    st.session_state.consent_family = []
                    for f in family_files:
                        img_fam = process_uploaded_file_to_image(f)
                        p_info = extract_passport_details(img_fam)
                        if p_info:
                            st.session_state.consent_family.append(p_info)

            st.success("🎉 모든 여권 정보가 성공적으로 추출되었습니다!")

    # ------------------------------------------
    # 정보 입력 폼
    # ------------------------------------------
    st.markdown("---")
    st.subheader("2. 비동반 부모님 정보 (동의서 작성인)")
    non_acc_data = st.session_state.consent_non_acc
    default_non_acc_name = format_full_name(non_acc_data.get('surname', ''), non_acc_data.get('given_name', ''), name_style)
    
    non_acc_name = st.text_input("비동반 부모 성명 (I, ...)", value=default_non_acc_name)
    non_acc_address = st.text_input("비동반 부모 주소 (Address)", value="")
    col_a, col_b = st.columns(2)
    with col_a:
        non_acc_phone = st.text_input("비동반 부모 전화번호", value="")
    with col_b:
        non_acc_email = st.text_input("비동반 부모 이메일", value="")

    st.markdown("---")
    st.subheader("3. 동반 부모 및 자녀 정보")
    
    acc_parents = []
    children_list = []

    for person in st.session_state.consent_family:
        dob = person.get("dob", "")
        if is_minor(dob):
            children_list.append(person)
        else:
            acc_parents.append(person)

    acc_name, acc_passport, acc_rel = "", "", "Mother"
    if acc_parents:
        selected_parent_str = st.selectbox("동반 부모님 선택", [f"{format_full_name(p.get('surname',''), p.get('given_name',''), name_style)} ({p.get('passport_number')})" for p in acc_parents])
        selected_idx = 0
        for idx, p in enumerate(acc_parents):
            formatted_p_name = format_full_name(p.get('surname',''), p.get('given_name',''), name_style)
            if formatted_p_name in selected_parent_str:
                selected_idx = idx
                break
        
        parent_info = acc_parents[selected_idx]
        acc_name = format_full_name(parent_info.get('surname',''), parent_info.get('given_name',''), name_style)
        acc_passport = parent_info.get("passport_number", "")
        acc_rel = "Mother" if parent_info.get("gender") == "F" else "Father"
    
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        acc_name = st.text_input("동반 부모 성명", value=acc_name)
    with col_p2:
        acc_rel = st.selectbox("관계 (Relationship)", ["Mother", "Father"], index=0 if acc_rel == "Mother" else 1)
    with col_p3:
        acc_passport = st.text_input("여권번호", value=acc_passport)

    st.write("##### 👶 자녀 목록 (3명 초과 시 추가 페이지 자동 생성)")
    final_children = []
    if children_list:
        for idx, child in enumerate(children_list):
            c_col1, c_col2 = st.columns(2)
            c_full_name = format_full_name(child.get('surname',''), child.get('given_name',''), name_style)
            with c_col1:
                c_name = st.text_input(f"자녀 #{idx+1} 성명", value=c_full_name, key=f"cname_{idx}")
            with c_col2:
                c_dob = st.text_input(f"자녀 #{idx+1} 생년월일 (YYYY/MM/DD)", value=child.get("dob", "").replace("-", "/"), key=f"cdob_{idx}")
            final_children.append({"name": c_name, "dob": c_dob})
    else:
        st.info("업로드된 자녀 여권이 없습니다. 필요 시 아래에 직접 입력하세요.")
        c1, c2 = st.columns(2)
        with c1: c_name_manual = st.text_input("자녀 #1 성명", value="")
        with c2: c_dob_manual = st.text_input("자녀 #1 생년월일", value="")
        if c_name_manual:
            final_children.append({"name": c_name_manual, "dob": c_dob_manual})

    st.markdown("---")
    st.subheader("4. 캐나다 현지 체류 정보 (Contact Information during Trip)")
    st.info(f"💡 'To stay with'는 동반 부모님 이름({acc_name})으로 자동 기입되며, Travel Date는 빈칸으로 남겨집니다.")
    
    trip_address = st.text_input("캐나다 현지 주소 (At the following address)", value="")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        trip_phone = st.text_input("현지 연락처 전화번호", value="")
    with col_t2:
        trip_email = st.text_input("현지 연락처 이메일", value="")

    st.markdown("---")
    st.subheader("5. 동의서 문서 생성")
    sign_date_str = st.date_input("서명란 자동 기입 날짜", value=datetime.date.today()).strftime("%Y/%m/%d")

    if st.button("✈️ 한부모 동의서 생성 및 다운로드", type="primary"):
        if not consent_template_bytes:
            st.error("양식 PDF 파일(consent_template.pdf)을 찾을 수 없습니다. GitHub 저장소에 consent_template.pdf 파일이 있는지 확인해 주세요.")
        elif not non_acc_name:
            st.warning("비동반 부모님 성명을 입력해 주세요.")
        else:
            data_consent = {
                "non_acc_name": non_acc_name,
                "non_acc_address": non_acc_address,
                "non_acc_phone": non_acc_phone,
                "non_acc_email": non_acc_email,
                "children": final_children,
                "acc_name": acc_name,
                "acc_relationship": acc_rel,
                "acc_passport": acc_passport,
                "trip_address": trip_address,
                "trip_phone": trip_phone,
                "trip_email": trip_email,
                "sign_date": sign_date_str
            }
            try:
                filled_consent = fill_consent_letter(consent_template_bytes, data_consent)
                st.download_button("📥 완성된 한부모 동의서 다운로드", filled_consent, file_name=f"Consent_Letter_{non_acc_name}.pdf", mime="application/pdf")
                st.balloons()
            except Exception as e:
                st.error(f"동의서 생성 오류: {e}")
