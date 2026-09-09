# =====================================================================
# ui_integration.py
# app_v3_patch.py의 함수들을 실제로 Streamlit UI에 연결하는 코드입니다.
# 기존 app.py의 "6. Streamlit UI" 섹션 중,
#   "1. 손님 정보 (여권 업로드)" ~ "AI 채용공고 & 여권 실시간 분석 시작" 버튼
# 부분을 아래 내용으로 통째로 교체하시면 됩니다.
# =====================================================================

import streamlit as st

# app_v3_patch.py 에서 가져옴
# from app_v3_patch import generate_job_offer_from_existing

st.subheader("1. 서류 업로드")

passport_file = st.file_uploader(
    "손님 여권 이미지 또는 PDF",
    type=["jpg", "jpeg", "png", "pdf", "heic", "HEIC"],
    key="jo_passport",
)

old_jo_file = st.file_uploader(
    "기준으로 사용할 기존 잡오퍼 (있으면 이 문서를 그대로 편집합니다)",
    type=["docx"],
    key="jo_old_file",
)

job_posting_text = st.text_area(
    "채용 공고 링크(URL) 또는 공고 텍스트",
    height=100,
)

logo_file = st.file_uploader(
    "회사 로고 이미지 (기존 잡오퍼 편집 모드에서는 무시됩니다 — 원본 로고 그대로 유지)",
    type=["jpg", "jpeg", "png"],
    key="jo_logo",
)

if st.button("AI 분석 시작", type="primary", use_container_width=True):
    with st.spinner("정보 추출 중..."):
        extracted_info = {}

        # 신규 케이스에서만 로고 사용 (기존 문서 편집 모드는 원본 로고 유지)
        if logo_file and not old_jo_file:
            extracted_info["logo_bytes"] = logo_file.getvalue()

        if passport_file:
            pass_img = process_uploaded_file_to_image(passport_file)
            pass_data = extract_imm5476_info(pass_img)
            if pass_data:
                extracted_info["client_name"] = format_full_name(
                    pass_data.get("surname", ""), pass_data.get("given_name", "")
                )
                extracted_info["client_dob"] = pass_data.get("dob", "")

        if old_jo_file:
            existing_parsed = parse_existing_job_offer(
                old_jo_file.getvalue(), old_jo_file.type or "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            if existing_parsed:
                extracted_info.update(existing_parsed)
            # 최종 생성 시 "그대로 편집"할 원본 바이트를 세션에 보관
            st.session_state["_old_jo_bytes"] = old_jo_file.getvalue()
        else:
            st.session_state.pop("_old_jo_bytes", None)

        if job_posting_text.strip():
            raw_input = job_posting_text.strip()
            is_url = bool(re.search(r"https?://[^\s]+|www\.[^\s]+", raw_input))
            if is_url:
                url_match = re.search(r"(https?://[^\s]+|www\.[^\s]+)", raw_input).group(0)
                fetched_text, err = fetch_url_content_safe(url_match)
                if err:
                    st.warning(f"URL 조회 실패: {err} (입력하신 텍스트 자체를 분석합니다)")
                text_to_analyze = fetched_text if fetched_text else raw_input
            else:
                text_to_analyze = raw_input

            job_extracted = call_gemini_json(
                [f"Analyze this job posting content and extract the requested fields:\n\n{text_to_analyze}"],
                JOB_POSTING_SCHEMA,
            )
            if job_extracted:
                extracted_info.update(job_extracted)

        st.session_state.job_offer_data.update(extracted_info)
        st.success("분석이 완료되었습니다.")


# =====================================================================
# 최종 생성 버튼 (기존 "4. 최종 잡오퍼 정보 확인 및 수정" 섹션 맨 아래,
# "📄 MS Word (.docx) 생성 및 다운로드" 버튼의 onClick 로직 교체)
# =====================================================================

if st.button("📄 MS Word (.docx) 생성 및 다운로드", type="primary", use_container_width=True):
    if not c_name or not emp_name or not j_title:
        st.error("손님 성명, 회사명, 직책은 필수 입력 항목입니다.")
    else:
        final_jo_dict = {
            "client_name": c_name,
            "client_dob": c_dob,
            "offer_date": offer_dt,
            "employment_term": term_str,
            "start_date": start_dt_str,
            "employer_name": emp_name,
            "signer_name": signer_n,
            "signer_title": signer_t,
            "employer_address": emp_addr,
            "employer_phone": emp_phone,
            "employer_email": emp_email,
            "job_title": j_title,
            "noc_code": j_noc,
            "wage": j_wage,
            "hours": j_hours,
            "job_location": j_loc,
            "benefits": j_benefits,
            "overtime_clause": j_ot,
            "job_duties": j_duties_text,
            "logo_bytes": jo_data.get("logo_bytes"),
        }

        old_bytes = st.session_state.get("_old_jo_bytes")
        if old_bytes:
            # 기존 문서를 그대로 편집 (로고/서명/회사헤더 유지)
            docx_bytes = generate_job_offer_from_existing(
                old_bytes, final_jo_dict, PROVINCES, detect_province
            )
        else:
            # 신규 케이스: 빈 템플릿 기반 생성
            if not os.path.exists(TEMPLATE_PATH):
                st.error(f"템플릿 파일을 찾을 수 없습니다: {TEMPLATE_PATH}")
                st.stop()
            docx_bytes = generate_job_offer_docx(final_jo_dict)

        crm_client = "NAME"
        if c_name:
            parts = c_name.strip().split()
            if parts:
                crm_client = parts[0].capitalize()
        out_filename = f"[Job Offer]_{crm_client}.docx"

        st.success("잡오퍼 DOCX 문서 생성이 완료되었습니다!")
        st.download_button(
            label="📥 Job Offer 다운로드",
            data=docx_bytes,
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
