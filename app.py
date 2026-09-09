import os
import io
import datetime
import docx
from docx import Document
from docx.shared import Pt, Inches
from docxtpl import DocxTemplate
import streamlit as st

# ==========================================
# 템플릿 자동 생성 엔진 (GitHub에 파일 올릴 필요 없음)
# ==========================================
def ensure_templates_exist():
    os.makedirs("templates", exist_ok=True)
    path_a = "templates/Style_A_Template.docx"
    path_b = "templates/Style_B_Template.docx"

    # Style A 템플릿이 없으면 자동 작성 및 저장
    if not os.path.exists(path_a):
        doc = Document()
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)

        doc.add_paragraph("{{ offer_date }}")
        
        p_dear = doc.add_paragraph()
        p_dear.add_run("Dear ")
        p_dear.add_run("{{ client_name }}").bold = True
        p_dear.add_run(",")

        p_intro = doc.add_paragraph()
        p_intro.add_run("We are pleased to offer you a full-time position as ")
        p_intro.add_run("{{ job_title }}").bold = True
        p_intro.add_run(" for a ")
        p_intro.add_run("{{ employment_term }}").bold = True
        p_intro.add_run(" term with ")
        p_intro.add_run("{{ employer_name }}").bold = True
        p_intro.add_run(" based on the following terms and conditions:")

        p = doc.add_paragraph(); p.add_run("Job Title: ").bold = True; p.add_run("{{ job_title }}")

        p = doc.add_paragraph(); p.add_run("Job Duties:").bold = True
        doc.add_paragraph("{% for duty in job_duties %}")
        doc.add_paragraph("• {{ duty }}")
        doc.add_paragraph("{% endfor %}")

        p = doc.add_paragraph(); p.add_run("Compensation:\n").bold = True; p.add_run("Hourly wage and hours\nThe employee will be paid ${{ wage }} per hour, based on a minimum of {{ hours }} hours per week.")
        p = doc.add_paragraph(); p.add_run("Terms of Employment: \n").bold = True; p.add_run("This is a full-time, {{ employment_term }} employment term starting from the date agreed upon by the employer and employee.")
        p = doc.add_paragraph(); p.add_run("Job Location: \n").bold = True; p.add_run("{{ job_location }}")
        p = doc.add_paragraph(); p.add_run("Benefits:\n").bold = True; p.add_run("{{ benefits }}")
        p = doc.add_paragraph(); p.add_run("Confidentiality:\n").bold = True; p.add_run("By accepting the terms of this offer, the employee agrees to keep all confidential information obtained during their employment with {{ employer_name }} strictly confidential. The employee further agrees that, upon termination of employment for any reason, they will return all physical and digital property belonging to or originating from {{ employer_name }} within five days of receiving notice of termination.")
        p = doc.add_paragraph(); p.add_run("Start Date:\n").bold = True; p.add_run("{{ start_date }}")
        p = doc.add_paragraph(); p.add_run("Overtime:\n").bold = True; p.add_run("{{ overtime_clause }}")

        doc.add_paragraph("We are pleased to extend this offer of employment to you on behalf of {{ employer_name }}. We are confident that you will make a valuable contribution to our company, and we look forward to working with you.")

        doc.add_paragraph("Sincerely,\t\t\t\t\tI accept the terms of this offer:")
        doc.add_paragraph("X\t\t\t\t\t\t______")
        
        p_sig = doc.add_paragraph()
        p_sig.add_run("{{ signer_name }}\n").bold = True
        p_sig.add_run("{{ signer_title }}\n{{ employer_name }}\nT. {{ employer_phone }}")

        p_client_sig = doc.add_paragraph()
        p_client_sig.add_run("{{ client_name }}\n").bold = True
        p_client_sig.add_run("(DOB: {{ client_dob }})")

        doc.save(path_a)

    # Style B 템플릿이 없으면 자동 작성 및 저장
    if not os.path.exists(path_b):
        doc = Document()
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)

        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)

        doc.add_paragraph("{{ offer_date }}")
        
        p_dear = doc.add_paragraph()
        p_dear.add_run("Dear ")
        p_dear.add_run("{{ client_name }}").bold = True
        p_dear.add_run(",")

        p_intro = doc.add_paragraph()
        p_intro.add_run("{{ employer_name }}").bold = True
        p_intro.add_run(" is pleased to offer you the position of ")
        p_intro.add_run("{{ job_title }}").bold = True
        p_intro.add_run(". We are confident that your skills will be a great asset to our business. The details of your employment are outlined below:")

        def add_kv(label, val):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            p.add_run(f"{label}: ").bold = True
            p.add_run(val)

        add_kv("Job Title", "{{ job_title }}")
        add_kv("Employment Type", "Full-time")
        add_kv("Terms of Employment", "{{ employment_term }}")
        add_kv("Start Date", "{{ start_date }}")
        add_kv("Hourly wage", "{{ wage }} CAD per hour")

        p_ot = doc.add_paragraph(); p_ot.paragraph_format.space_after = Pt(2)
        p_ot.add_run("Overtime:\n").bold = True; p_ot.add_run("{{ overtime_clause }}")

        add_kv("Hours of Work", "{{ hours }} hours per week")
        add_kv("Work Location", "{{ job_location }}")
        add_kv("Vacation Pay", "{{ benefits }}")

        p = doc.add_paragraph(); p.add_run("Job Duties:").bold = True
        doc.add_paragraph("{% for duty in job_duties %}")
        doc.add_paragraph("• {{ duty }}")
        doc.add_paragraph("{% endfor %}")

        p_conf = doc.add_paragraph()
        p_conf.add_run("Confidentiality Agreement:\n").bold = True
        p_conf.add_run("By agreeing to the terms of this offer, the employee further agrees that they will hold all confidential information with which they are entrusted while employed by {{ employer_name }} in strict confidence. The employee also agrees that upon termination of employment for any reason, they will return all physical and digital property which belongs to or originates at {{ employer_name }} within 5 days of notice of termination of employment.")

        doc.add_paragraph("We look forward to your acceptance of this offer.\nPlease sign below to confirm your agreement.")
        doc.add_paragraph("Yours truly,")
        doc.add_paragraph("X_____________________________________________")

        p_sig = doc.add_paragraph(); p_sig.paragraph_format.space_after = Pt(2)
        p_sig.add_run("{{ signer_name }}\n").bold = True
        p_sig.add_run("{{ signer_title }}\n{{ employer_name }}\n{{ employer_address }}\nT. {{ employer_phone }}\nE. {{ employer_email }}")

        doc.add_paragraph("\nI accept the terms of this offer:")
        doc.add_paragraph("___________________________________________")

        p_acc = doc.add_paragraph()
        p_acc.add_run("{{ client_name }}\n").bold = True
        p_acc.add_run("(DOB: {{ client_dob }})")

        doc.save(path_b)

# 앱 실행 시 템플릿 파일이 없으면 서버에서 자동으로 즉시 생성
ensure_templates_exist()


# ==========================================
# 잡오퍼 DOCX 생성 함수
# ==========================================
def generate_job_offer_from_template(data, selected_style="Style A"):
    template_path = "templates/Style_A_Template.docx" if "Style A" in selected_style else "templates/Style_B_Template.docx"
    doc = DocxTemplate(template_path)
    
    duties = data.get('job_duties', [])
    if isinstance(duties, str):
        duties = [d.strip() for d in duties.split('\n') if d.strip()]

    context = {
        'offer_date': data.get('offer_date', datetime.date.today().strftime("%B %d, %Y")),
        'client_name': data.get('client_name', ''),
        'client_dob': data.get('client_dob', ''),
        'employer_name': data.get('employer_name', ''),
        'signer_name': data.get('signer_name', ''),
        'signer_title': data.get('signer_title', ''),
        'employer_address': data.get('employer_address', ''),
        'employer_phone': data.get('employer_phone', ''),
        'employer_email': data.get('employer_email', ''),
        'job_title': data.get('job_title', ''),
        'employment_term': data.get('employment_term', '3 years'),
        'start_date': data.get('start_date', 'As soon as possible upon authorization to work in Canada'),
        'wage': str(data.get('wage', '')),
        'hours': str(data.get('hours', '30')),
        'job_location': data.get('job_location', ''),
        'benefits': data.get('benefits', '4% vacation pay'),
        'overtime_clause': data.get('overtime_clause', ''),
        'job_duties': duties
    }
    
    doc.render(context)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()
