import os
import subprocess
import fitz
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

FONT = "Garamond"
BLACK = RGBColor(0, 0, 0)
NAVY = RGBColor(0x00, 0x33, 0x66)

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
    _run(p2, "Tallahassee, FL (Local Resident) · (314) 732-5088 · asmibas18@gmail.com · linkedin.com/in/asmibasnet · asmibasnet.github.io", 9.5)

    # Date & Addressee
    p3 = d.add_paragraph()
    p3.paragraph_format.space_after = Pt(8)
    _run(p3, "October 22, 2026\n\nHiring Team — Provider Network Analytics (Job ID 2039725)\nMolina Healthcare of Florida\nTallahassee, FL", 10.0)

    # Salutation
    p4 = d.add_paragraph()
    p4.paragraph_format.space_after = Pt(8)
    _run(p4, "Dear Hiring Manager and Provider Network Analytics Team,", 10.0)

    # Body Paragraph 1
    p5 = d.add_paragraph()
    p5.paragraph_format.space_after = Pt(7)
    p5.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p5, "I am writing to express my enthusiastic interest in the Senior Analyst, Data & Analytics (Provider Network - FL) position (Job ID: 2039725) at Molina Healthcare. As a Tallahassee local resident completing my Master of Science in Analytics at Saint Louis University, I offer over three years of applied experience in healthcare data analytics, provider performance tracking, utilization analysis, and automated data engineering. Having built clinical analytics engines and optimized enterprise data pipelines, I am eager to contribute to Molina Healthcare's mission of delivering cost-effective, high-quality care to vulnerable populations across Florida.", 10.0)

    # Body Paragraph 2
    p6 = d.add_paragraph()
    p6.paragraph_format.space_after = Pt(7)
    p6.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p6, "In my role as Graduate Researcher with the SPS Corps Analytics Initiative at Saint Louis University, I collaborate with operational, clinical, and financial stakeholders to deliver actionable analytical solutions and executive dashboards. To support faster reporting and cost-containment analyses, I refactored complex SQL extraction queries across 8+ operational databases using modular CTEs and analytical window functions, slashing query execution latency by 40%. Additionally, I design and maintain interactive Power BI dashboards backed by star-schema models and advanced Excel workbooks (Power Pivot, dynamic modeling formulas) that track operational throughput, resource utilization, and cost benchmarks across departments.", 10.0)

    # Body Paragraph 3
    p7 = d.add_paragraph()
    p7.paragraph_format.space_after = Pt(7)
    p7.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p7, "My technical healthcare experience directly complements Molina's provider network priorities. In my CareGap analytics platform, I analyzed 34,000+ longitudinal patient records using columnar DuckDB and PostgreSQL to evaluate quality performance metrics, clinical care-gap closure rates, and cost drivers. For the WashU Clinical Analytics Datathon, I engineered an end-to-end predictive pipeline across 62,000+ clinical records, developing 59 predictive features (visit frequency, medication count, volatility) to isolate high-risk utilization cohorts (AUC 0.8520). Furthermore, I established automated Independent Verification & Validation (IV&V) frameworks that detect schema drift and reconcile transaction discrepancies, ensuring total data fidelity and adherence to strict HIPAA standards.", 10.0)

    # Body Paragraph 4
    p8 = d.add_paragraph()
    p8.paragraph_format.space_after = Pt(10)
    p8.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p8, "Being physically based in Tallahassee, Florida, I bring immediate availability, local regional alignment, and a deep dedication to supporting Molina's Florida provider networks and state-mandated regulatory reporting. I would welcome the opportunity to discuss how my technical expertise in SQL, Power BI, and healthcare predictive modeling can support Molina Healthcare's strategic objectives. Thank you for your time, consideration, and review.", 10.0)

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
    ps_file = "convert_molina_cl_temp.ps1"
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps_script)
    res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file], capture_output=True, text=True)
    if os.path.exists(ps_file):
        os.remove(ps_file)
    print("Conversion output:", res.stdout, res.stderr)

if __name__ == "__main__":
    out_dir = r"c:\Users\Shayara Basnet\Downloads\portfolio"
    docx_path = os.path.join(out_dir, "Asmi_Basnet_Cover_Letter_Molina_Healthcare.docx")
    pdf_path = os.path.join(out_dir, "Asmi_Basnet_Cover_Letter_Molina_Healthcare.pdf")
    
    doc = build_cover_letter()
    doc.save(docx_path)
    convert_to_pdf(docx_path, pdf_path)
    if os.path.exists(pdf_path):
        p1 = len(fitz.open(pdf_path))
        print(f"Cover Letter PDF generated: {pdf_path} (Page count: {p1})")
