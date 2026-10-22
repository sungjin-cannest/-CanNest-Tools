MENU_1 = "🍁 IMM5476 자동 작성"
MENU_2 = "✈️ 한부모 동의서 자동 작성"
MENU_3 = "📋 IMM서류 정보 정리"
MENU_4 = "🏷️ CRM 파일명 생성 및 묶기/분할"
MENU_5 = "📝 질문지 자동 이어채우기 (QA Re-fill)"

st.sidebar.title("🦅 CanNest Tool")
app_mode = st.sidebar.radio("원하시는 업무 도구를 선택하세요", [MENU_1, MENU_2, MENU_3, MENU_4, MENU_5])

# ------------------------------------------
# 메뉴 5: 질문지 자동 이어채우기 (QA Re-fill)
# ------------------------------------------
elif app_mode == MENU_5:
    st.title("📝 질문지 자동 이어채우기 (QA Re-fill)")
    st.caption("손님이 과거에 작성했던 질문지나 여권/기본서류를 넣고 새 질문지를 올리면, 기존 정보를 자동으로 채우고 손님께 추가로 받아야 할 질문만 정리해 드립니다.")

    if "qa_plan_result" not in st.session_state:
        st.session_state.qa_plan_result = None

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 손님 기존 보유 자료")
        source_client_files = st.file_uploader(
            "이전 질문지, 여권, 퍼밋, 이력서 등 (복수 선택)",
            type=['pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png'],
            accept_multiple_files=True,
            key="qa_src_files"
        )
        staff_notes = st.text_area("담당자 메모 (선택)", placeholder="예: 2026년 2월에 이사함, 배우자 동반 여부 미정 등 특이사항 기재", height=85)

    with c2:
        st.subheader("2. 새로 채워야 할 질문지")
        target_q_file = st.file_uploader(
            "작성 대상 새 질문지 1부 (.docx 또는 .pdf)",
            type=['docx', 'pdf'],
            key="qa_target_file"
        )
        mark_blue_opt = st.checkbox("자동 입력된 답변 파란색으로 강조", value=True)

    if st.button("질문지 대조 및 자동 채우기 시작", type="primary", use_container_width=True):
        if not source_client_files:
            st.warning("1단계: 손님의 기존 자료를 1개 이상 업로드해 주세요.")
        elif not target_q_file:
            st.warning("2단계: 새로 채워 넣을 질문지 파일을 업로드해 주세요.")
        else:
            with st.spinner("기존 서류들을 대조하여 새 질문지를 자동 작성 중입니다..."):
                plan = extract_qa_fill_plan(source_client_files, target_q_file, staff_notes)
                if plan:
                    st.session_state.qa_plan_result = plan
                    st.success("질문지 채우기 및 미비 항목 분석이 완료되었습니다!")

    if st.session_state.qa_plan_result:
        res = st.session_state.qa_plan_result
        summary = res.get("summary", {})
        fill_items = res.get("fill_items", [])
        missing_items = res.get("missing_questions", [])
        msg_draft = res.get("client_message_draft", "")

        st.markdown("---")
        st.subheader("3. 분석 및 채우기 결과")

        # 상단 요약 지표
        m1, m2, m3 = st.columns(3)
        m1.metric("자동 채움 완료", f"{len(fill_items)}개")
        m2.metric("확인 필요 (변동 가능성)", f"{len([x for x in fill_items if x.get('needs_check')])}개")
        m3.metric("손님 추가 요청 필요", f"{len(missing_items)}개")

        # 1) 다운로드 영역
        is_docx_target = target_q_file.name.lower().endswith(".docx")
        if is_docx_target:
            filled_bytes = fill_docx_questionnaire(target_q_file.getvalue(), fill_items, mark_blue=mark_blue_opt)
            out_filename = f"{target_q_file.name.rsplit('.', 1)[0]}_자동완성.docx"
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            filled_bytes = target_q_file.getvalue()
            out_filename = f"{target_q_file.name.rsplit('.', 1)[0]}_확인용.pdf"
            mime_type = "application/pdf"

        st.download_button(
            label=f"📥 완성된 질문지 내려받기 ({out_filename})",
            data=filled_bytes,
            file_name=out_filename,
            mime=mime_type,
            type="primary"
        )

        # 2) 고객 전송용 카카오톡/이메일 안내문 복사 영역
        st.write("### 📲 손님 안내 메시지 (카카오톡/이메일용)")
        st.caption("아래 내용을 복사하여 손님께 그대로 전송하시면 편리합니다.")
        st.text_area("고객 발송용 텍스트", value=msg_draft, height=200)

        # 3) 상세 대조 테이블
        t1, t2 = st.tabs(["✅ 자동 채워진 항목", "⚠️ 손님께 새로 받아야 할 항목"])
        with t1:
            table_rows = []
            for item in fill_items:
                status = "⚠️ 확인 요망" if item.get("needs_check") else "정상 채움"
                table_rows.append({
                    "질문 항목": item.get("question_label"),
                    "채워진 답변": item.get("matched_answer"),
                    "상태": status,
                    "정보 출처": item.get("source_doc")
                })
            st.table(table_rows)

        with t2:
            missing_rows = []
            for item in missing_items:
                missing_rows.append({
                    "미비 질문 항목": item.get("question_label"),
                    "사유": item.get("reason")
                })
            st.table(missing_rows)

        if st.button("🔄 새 질문지 케이스 진행하기"):
            st.session_state.qa_plan_result = None
            st.rerun()
