import os
import subprocess
import fitz
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

FONT = "Garamond"
BLACK = RGBColor(0, 0, 0)
NAVY = RGBColor(0x11, 0x33, 0x59)

def _run(p, text, size=10.5, bold=False, italic=False, color=BLACK):
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.name = FONT
    r.font.color.rgb = color
    rpr = r._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)
    return r

def build_cover_letter():
    d = Document()
    sec = d.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11.0)
    sec.left_margin = Inches(0.75)
    sec.right_margin = Inches(0.75)
    sec.top_margin = Inches(0.65)
    sec.bottom_margin = Inches(0.65)

    # Header
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, "ASMI BASNET", 15, bold=True, color=NAVY)
    
    p2 = d.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(14)
    _run(p2, "Tallahassee, FL · (314) 732-5088 · asmibas18@gmail.com · linkedin.com/in/asmibasnet · asmibasnet.github.io", 9.5)

    # Date & Addressee
    p3 = d.add_paragraph()
    p3.paragraph_format.space_after = Pt(8)
    _run(p3, "October 22, 2026\n\nSearch Committee for Financial Data & Reporting Analyst (Requisition JR102156)\nRaymond A. Mason School of Business | Miller Hall\nWilliam & Mary\nWilliamsburg, VA 23187", 10.0)

    # Salutation
    p4 = d.add_paragraph()
    p4.paragraph_format.space_after = Pt(8)
    _run(p4, "Dear Senior Director Financial Planning & Data Analysis, Director Data Analytics, and Members of the Search Committee,", 10.0)

    # Body Paragraph 1
    p5 = d.add_paragraph()
    p5.paragraph_format.space_after = Pt(7)
    p5.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p5, "I am writing to express my enthusiastic interest in the Financial Data & Reporting Analyst position (JR102156) with the Raymond A. Mason School of Business at William & Mary. Currently completing my Master of Science in Analytics at Saint Louis University, I bring three years of hands-on experience in higher education institutional analytics, recurring financial and operational reporting, budget-to-actual variance analysis, and automated data reconciliation. Having worked directly within university administration supporting executive decision-making, I am excited by the opportunity to bring my analytical rigor, technical expertise in Power BI and SQL, and commitment to institutional data stewardship to the Mason School's Finance and Administration team.", 10.0)

    # Body Paragraph 2
    p6 = d.add_paragraph()
    p6.paragraph_format.space_after = Pt(7)
    p6.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p6, "In my role as Graduate Researcher with the SPS Corps Analytics Initiative at Saint Louis University, I collaborate with academic leadership and departmental administrators to deliver recurring financial, enrollment, and operational reports. I perform routine budget-to-actual evaluations and variance analyses across complex fiscal cycles, isolating contributing factors and presenting executive summaries to non-technical stakeholders. To support faster, cleaner reporting cycles, I architected automated SQL data pipelines across 8+ operational databases, refactoring legacy queries into modular CTEs and analytical window functions to reduce extraction latency by 40%. Additionally, I design and maintain interactive Power BI dashboards backed by star-schema models and robust Excel reporting workbooks that monitor tuition trends, departmental budgets, and operational performance.", 10.0)

    # Body Paragraph 3
    p7 = d.add_paragraph()
    p7.paragraph_format.space_after = Pt(7)
    p7.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p7, "A central pillar of the Mason School's mission is data accuracy and integrity. During my tenure at SLU, I established automated Independent Verification and Validation (IV&V) frameworks in Python and SQL across 10,000+ institutional transaction records, proactively catching schema drift, referential integrity violations, and ledger discrepancies before reports reached administrators. I am deeply familiar with maintaining institutional data dictionaries, documenting report logic, and strictly adhering to university privacy and FERPA governance standards. Whether validating data extracted from Workday and student information systems or developing time-series budget projections to flag anomalies, I take immense pride in ensuring every metric presented is thoroughly reconciled and defensible.", 10.0)

    # Body Paragraph 4
    p8 = d.add_paragraph()
    p8.paragraph_format.space_after = Pt(10)
    p8.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p8, "William & Mary's storied heritage of academic distinction and the Raymond A. Mason School of Business's dedication to developing principled leaders inspire me. I am fully prepared to meet the on-site and hybrid demands at Miller Hall and contribute immediately to your annual reporting, accreditation cycles, and process automation initiatives. Thank you for your time, consideration, and review of my application materials.", 10.0)

    # Sign-off
    p9 = d.add_paragraph()
    p9.paragraph_format.space_after = Pt(4)
    _run(p9, "Sincerely,", 10.0)
    
    p10 = d.add_paragraph()
    _run(p10, "Asmi Basnet\n(314) 732-5088 | asmibas18@gmail.com", 10.0, bold=True)

    return d

def convert_to_pdf(docx_path, pdf_path):
    ps_script = f"""
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {{
        $doc = $word.Documents.Open('{docx_path}', $false, $true, $false)
        $doc.SaveAs([ref]'{pdf_path}', [ref]17)
        $doc.Close([ref]0)
    }} catch {{
        Write-Output ("Error: " + $_.Exception.Message)
    }} finally {{
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }}
    """
    ps_file = "convert_cl_temp.ps1"
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps_script)
    res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file], capture_output=True, text=True)
    if os.path.exists(ps_file):
        os.remove(ps_file)
    print("Conversion output:", res.stdout, res.stderr)

if __name__ == "__main__":
    out_dir = r"c:\Users\Shayara Basnet\Downloads\portfolio"
    docx_path = os.path.join(out_dir, "Asmi_Basnet_Cover_Letter_William_Mary.docx")
    pdf_path = os.path.join(out_dir, "Asmi_Basnet_Cover_Letter_William_Mary.pdf")
    
    doc = build_cover_letter()
    doc.save(docx_path)
    convert_to_pdf(docx_path, pdf_path)
    if os.path.exists(pdf_path):
        p1 = len(fitz.open(pdf_path))
        print(f"Cover Letter PDF generated: {pdf_path} (Page count: {p1})")
