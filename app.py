# =====================================================================
# app_v3_patch.py
# 기존 app.py(v2)에 아래 내용을 "교체/추가"하는 패치입니다.
# 섹션별로 "어디를 바꾸는지" 주석에 명시했으니, 해당 위치의 기존 함수를
# 통째로 이 버전으로 교체하시면 됩니다. (새 함수는 그냥 추가)
# =====================================================================

import io
import re
import copy
import datetime

import docx
from docx.oxml.ns import qn


# =====================================================================
# [수정 1] Cloudflare 이메일 난독화 디코딩
# → fetch_url_content_safe() 안에서 html 읽은 직후에 호출합니다.
#    (섹션 3. URL 보안 강화 fetch 부분)
# 원인: 채용공고 사이트가 스팸봇 방지로 이메일을 [email protected] 로
#       표시하고, 실제 주소는 data-cfemail="hex..." 안에 XOR 인코딩되어
#       숨어있습니다. JS를 실행하지 않는 한 그대로 노출되지 않던 것입니다.
# =====================================================================
def _decode_cf_email(cfemail_hex: str) -> str:
    try:
        r = int(cfemail_hex[:2], 16)
        return "".join(
            chr(int(cfemail_hex[i:i + 2], 16) ^ r)
            for i in range(2, len(cfemail_hex), 2)
        )
    except Exception:
        return ""


_CF_EMAIL_RE = re.compile(
    r'data-cfemail="([0-9a-fA-F]+)"[^>]*>\s*\[email\s*protected\]\s*<',
    re.IGNORECASE,
)


def _decloak_cloudflare_emails(html: str) -> str:
    """[email protected] 로 가려진 부분을 실제 이메일 주소로 되돌립니다."""
    return _CF_EMAIL_RE.sub(lambda m: _decode_cf_email(m.group(1)) + "<", html)


# fetch_url_content_safe() 안에서 이 줄만 추가하면 됩니다:
#
#     html = raw.decode("utf-8", errors="ignore")
#     html = _decloak_cloudflare_emails(html)      # <-- 이 줄 추가
#
# (그 아래 <script>/<style> 제거, 태그 제거 로직은 그대로 둡니다)


# =====================================================================
# [수정 2] 프로빈스별 overtime 조항 — intro 중복 버그 수정
# 기존 PROVINCES[code]["overtime"] 값에 "Overtime will be paid..." 인트로
# 문장이 매번 포함되어 있는데, 실제 문서에는 "Overtime:" 섹션에 이미
# 공통 인트로 문장이 한 줄 있어서 2번 나오는 버그였습니다.
# → overtime 값에서 "불릿(요율) 줄만" 남기도록 분리합니다.
# =====================================================================
def get_provincial_overtime_bullets(address_text, PROVINCES, detect_province):
    """intro 문장 없이, 요율 불릿 줄만 리스트로 반환."""
    full_clause = PROVINCES[detect_province(address_text)]["overtime"]
    lines = [l.strip() for l in full_clause.split("\n") if l.strip()]
    # 각 province 딕셔너리의 첫 줄이 "Overtime will be paid..." 로 시작하면 제외
    if lines and lines[0].lower().startswith("overtime will be paid"):
        lines = lines[1:]
    return lines


# =====================================================================
# [수정 3] wage/hours 문장 정규화 — "$$22.00/hour per hour" 같은
# 단위 중복 버그 방지. clean_wage/clean_hours로 숫자만 뽑은 뒤,
# 문장을 매번 "새로 조립"해서 넣습니다 (기존 문장에 끼워넣지 않음).
# =====================================================================
def build_wage_sentence(wage: str, hours: str) -> str:
    return f"The employee will be paid ${wage}/hour, based on a minimum of {hours} hours per week."


def build_term_sentence(term_years: str) -> str:
    # term_years 예: "1-year", "3-year"
    return f"This is a full-time, {term_years} employment term starting from the date agreed upon by the employer and employee."


# =====================================================================
# [핵심 추가] 기존 업로드 문서를 "그 자체로" 편집하는 엔진
# → 빈 TEMPLATE_PATH를 열지 않고, 사용자가 올린 기존 잡오퍼 docx를
#   그대로 열어서 바뀌는 항목만 라벨 기준으로 치환합니다.
#   로고 / 서명란 / 회사 헤더 / Confidentiality 문구 등은 건드리지 않습니다.
# =====================================================================

def _iter_all_paragraphs(doc):
    """app.py의 iter_all_paragraphs와 동일 (표 안까지 재귀)"""
    def _iter_cell(cell):
        for p in cell.paragraphs:
            yield p
        for table in cell.tables:
            for row in table.rows:
                for c in row.cells:
                    yield from _iter_cell(c)

    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_cell(cell)


def _set_paragraph_text(paragraph, new_text):
    """첫 run의 서식(굵기/폰트/크기)을 유지하며 전체 텍스트 교체.
    app.py의 set_paragraph_text와 동일 로직."""
    if not paragraph.runs:
        paragraph.add_run(new_text)
        return
    base_rpr = paragraph.runs[0]._r.find(qn("w:rPr"))
    for r in list(paragraph.runs):
        r._r.getparent().remove(r._r)
    run = paragraph.add_run()
    if base_rpr is not None:
        run._r.insert(0, copy.deepcopy(base_rpr))
    run.text = new_text


def _find_para_startswith(doc, prefix):
    for p in doc.paragraphs:
        if p.text.strip().lower().startswith(prefix.lower()):
            return p
    return None


def _replace_label_line(doc, label, new_value):
    """'Job Title: Chef' 처럼 라벨과 내용이 한 줄에 있는 경우,
    라벨은 유지하고 내용만 교체. 못 찾으면 아무 것도 안 함."""
    p = _find_para_startswith(doc, label)
    if p is None:
        return False
    _set_paragraph_text(p, f"{label} {new_value}")
    return True


def _replace_next_paragraph(doc, label, new_value):
    """'Job Location:' 처럼 라벨 단독 줄 다음에 내용이 오는 경우,
    라벨 줄은 그대로 두고 바로 다음 문단 텍스트만 통째로 교체."""
    p = _find_para_startswith(doc, label)
    if p is None:
        return False
    nxt = p._p.getnext()
    if nxt is None or nxt.tag != qn("w:p"):
        return False
    wrapped = docx.text.paragraph.Paragraph(nxt, p._parent)
    _set_paragraph_text(wrapped, new_value)
    return True


def _replace_bullets_after(doc, label, new_lines):
    """label 문단 뒤에 이어지는 'List Paragraph' 스타일 형제 문단들을
    모조리 제거하고 new_lines로 재삽입. (Job Duties / Overtime 공용)
    label 문단과 첫 불릿 사이에 인트로 문장이 껴 있으면 그건 건드리지 않고
    그 다음부터 나오는 List Paragraph만 대상으로 합니다."""
    anchor = _find_para_startswith(doc, label)
    if anchor is None or not new_lines:
        return False

    node = anchor._p.getnext()
    # 인트로 문단(들)은 건너뛰고, 첫 List Paragraph를 찾음
    while node is not None and node.tag == qn("w:p"):
        wrapped = docx.text.paragraph.Paragraph(node, anchor._parent)
        if wrapped.style.name == "List Paragraph":
            break
        node = node.getnext()

    bullets = []
    cur = node
    while cur is not None and cur.tag == qn("w:p"):
        wrapped = docx.text.paragraph.Paragraph(cur, anchor._parent)
        if wrapped.style.name != "List Paragraph":
            break
        bullets.append(cur)
        cur = cur.getnext()

    if not bullets:
        return False

    template_xml = copy.deepcopy(bullets[0])
    ref = bullets[0].getprevious()
    for b in bullets:
        b.getparent().remove(b)
    for line in new_lines:
        new_p = copy.deepcopy(template_xml)
        ref.addnext(new_p)
        ref = new_p
        wrapped = docx.text.paragraph.Paragraph(new_p, anchor._parent)
        _set_paragraph_text(wrapped, line)
    return True


def _update_salutation(doc, new_name):
    p = _find_para_startswith(doc, "Dear ")
    if p is None:
        return False
    text = p.text
    new_text = re.sub(r"^Dear .+?,", f"Dear {new_name},", text)
    _set_paragraph_text(p, new_text)
    return True


def _update_job_title_line(doc, job_title, noc_code):
    """'Job Title: Chef' 또는 'Job Title: Chef (NOC 63200)' 형태 모두 처리."""
    p = _find_para_startswith(doc, "Job Title:")
    if p is None:
        return False
    if noc_code and str(noc_code).strip():
        new_text = f"Job Title: {job_title} (NOC {noc_code})"
    else:
        new_text = f"Job Title: {job_title}"
    _set_paragraph_text(p, new_text)
    return True


def generate_job_offer_from_existing(
    existing_file_bytes: bytes,
    data: dict,
    PROVINCES: dict,
    detect_province,
) -> bytes:
    """업로드된 '기존 잡오퍼' docx를 그대로 열어서, 바뀌는 항목만 교체합니다.
    로고 / 서명란 / 회사 헤더 / Confidentiality 문구는 손대지 않습니다.

    data 예시 키: client_name, offer_date, job_title, noc_code,
                  employment_term (예: '1-year'), wage, hours,
                  job_location, job_duties(list[str]), start_date
    """
    doc = docx.Document(io.BytesIO(existing_file_bytes))

    # 1) 상단 날짜 (문서 맨 위 'April 4, 2025' 같은 단독 줄)
    if data.get("offer_date"):
        for p in doc.paragraphs[:5]:
            if re.match(r"^[A-Z][a-z]+ \d{1,2},? \d{4}$", p.text.strip()):
                _set_paragraph_text(p, data["offer_date"])
                break

    # 2) 받는 사람 (Dear X,)
    if data.get("client_name"):
        _update_salutation(doc, data["client_name"])

    # 3) Job Title (+ NOC)
    if data.get("job_title"):
        _update_job_title_line(doc, data["job_title"], data.get("noc_code", ""))

    # 4) Job Duties 불릿 교체
    duties = data.get("job_duties", [])
    if isinstance(duties, str):
        duties = [re.sub(r"^[•\-\*]\s*", "", d.strip()) for d in duties.split("\n") if d.strip()]
    if duties:
        _replace_bullets_after(doc, "Job Duties:", duties)

    # 5) 급여/시간 문장 — 라벨 다음 줄을 새로 조립해서 교체 (단위 중복 방지)
    if data.get("wage") and data.get("hours"):
        sentence = build_wage_sentence(data["wage"], data["hours"])
        if not _replace_next_paragraph(doc, "Hourly wage and hours", sentence):
            _replace_next_paragraph(doc, "Hourly Wage and Hours", sentence)

    # 6) 고용 기간 문장
    if data.get("employment_term"):
        term_sentence = build_term_sentence(data["employment_term"])
        _replace_next_paragraph(doc, "Terms of Employment:", term_sentence)

    # 7) 근무지 주소 (기존 회사 주소와 다를 때만 의미있음)
    if data.get("job_location"):
        _replace_next_paragraph(doc, "Job Location:", data["job_location"])

    # 8) Overtime — 인트로는 그대로 두고 요율 불릿만 주(province)에 맞게 교체
    bullets = get_provincial_overtime_bullets(
        data.get("job_location", ""), PROVINCES, detect_province
    )
    if bullets:
        _replace_bullets_after(doc, "Overtime:", bullets)

    # 9) Start Date
    if data.get("start_date"):
        _replace_next_paragraph(doc, "Start Date:", data["start_date"])

    # → 로고, 서명 이미지, 회사명/주소/전화, Confidentiality 문구는
    #    위 어떤 함수도 건드리지 않으므로 원본 그대로 유지됩니다.

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =====================================================================
# [UI 흐름 통합] 2단계 추출 → 1단계로
# 기존: "기존 잡오퍼 분석하여 정보 가져오기" 버튼 + 별도로
#       "AI 채용공고 & 여권 실시간 분석 시작" 버튼 → 총 2번 클릭, 2번 분석
# 변경: 파일들(기존 오퍼 / 여권 / 채용공고)을 다 올려놓고
#       버튼 "AI 분석 시작" 한 번으로 전부 처리.
# =====================================================================
"""
Streamlit UI 쪽 교체 예시 (기존 섹션 1~3을 아래처럼 하나의 흐름으로):

st.subheader("1. 서류 업로드")
passport_file = st.file_uploader("손님 여권", type=[...], key="jo_passport")
old_jo_file = st.file_uploader(
    "기준으로 사용할 기존 잡오퍼 (있으면 이 문서를 그대로 편집합니다)",
    type=["docx"], key="jo_old_file"
)
job_posting_text = st.text_area("채용 공고 링크(URL) 또는 텍스트", height=100)
logo_file = st.file_uploader("회사 로고 (기존 잡오퍼 편집 모드에서는 무시됨)", ...)

if st.button("AI 분석 시작", type="primary", use_container_width=True):
    with st.spinner("정보 추출 중..."):
        extracted_info = {}

        if passport_file:
            ... # 기존과 동일

        if old_jo_file:
            existing_parsed = parse_existing_job_offer(
                old_jo_file.getvalue(), old_jo_file.type or "..."
            )
            extracted_info.update(existing_parsed or {})
            st.session_state["_old_jo_bytes"] = old_jo_file.getvalue()  # 편집 대상 원본 보관

        if job_posting_text.strip():
            ... # 기존과 동일하게 URL fetch + call_gemini_json

        st.session_state.job_offer_data.update(extracted_info)
        st.success("분석 완료")

# 최종 생성 버튼에서:
if st.session_state.get("_old_jo_bytes"):
    docx_bytes = generate_job_offer_from_existing(
        st.session_state["_old_jo_bytes"], final_jo_dict, PROVINCES, detect_province
    )
else:
    docx_bytes = generate_job_offer_docx(final_jo_dict)  # 기존 템플릿 기반 (신규 케이스용)
"""
