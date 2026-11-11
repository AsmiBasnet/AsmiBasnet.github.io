import os
import sys
import subprocess
import fitz
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn

FONT = "Garamond"
BLACK = RGBColor(0, 0, 0)
NAVY = RGBColor(0x11, 0x33, 0x59)  # Deep William & Mary / Executive Navy

def _run(p, text, size, bold=False, italic=False, color=BLACK):
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

class ResumeDoc:
    def __init__(self, fs=8.2, top_margin=0.28, bottom_margin=0.28, left_margin=0.38, right_margin=0.38, line_spacing=1.02):
        self.fs = fs
        self.d = Document()
        sec = self.d.sections[0]
        sec.page_width = Inches(8.5)
        sec.page_height = Inches(11.0)
        sec.left_margin = Inches(left_margin)
        sec.right_margin = Inches(right_margin)
        sec.top_margin = Inches(top_margin)
        sec.bottom_margin = Inches(bottom_margin)
        self.tab = sec.page_width - Inches(left_margin + right_margin)
        
        for name in ("Normal", "Heading 1", "List Bullet"):
            st = self.d.styles[name]
            st.font.name = FONT
            st.font.size = Pt(fs)
            st.font.color.rgb = BLACK
            rpr = st.element.get_or_add_rPr()
            rf = rpr.find(qn("w:rFonts"))
            if rf is None:
                rf = rpr.makeelement(qn("w:rFonts"), {})
                rpr.append(rf)
            for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                rf.set(qn(a), FONT)
            for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
                if rf.get(qn(a)) is not None:
                    del rf.attrib[qn(a)]
            pf = st.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = line_spacing
            
        h = self.d.styles["Heading 1"]
        h.font.bold = True
        h.font.italic = False
        self.d.core_properties.author = "Asmi Basnet"
        self.d.core_properties.title = "Asmi Basnet - William & Mary Financial Data & Reporting Analyst Resume"

    def p(self, style=None, align=None):
        p = self.d.add_paragraph(style=style) if style else self.d.add_paragraph()
        if align is not None:
            p.alignment = align
        return p

    def heading(self, text, space_before=1.6, space_after=0.1):
        p = self.p("Heading 1")
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        _run(p, text.upper(), self.fs + 0.8, bold=True, color=NAVY)

    def lr(self, left, right, bold=False, italic=False, space_before=0.5):
        p = self.p()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.tab_stops.add_tab_stop(self.tab, WD_TAB_ALIGNMENT.RIGHT)
        _run(p, left, self.fs, bold=bold, italic=italic)
        _run(p, "\t" + right, self.fs, bold=bold, italic=italic)

    def bullet(self, text, space_after=0.1):
        p = self.p("List Bullet", WD_ALIGN_PARAGRAPH.JUSTIFY)
        pf = p.paragraph_format
        pf.left_indent = Inches(0.18)
        pf.first_line_indent = Inches(-0.12)
        pf.space_after = Pt(space_after)
        _run(p, text, self.fs)

    def plain(self, text, justify=True, italic_lead=None, space_after=0.1):
        p = self.p(None, WD_ALIGN_PARAGRAPH.JUSTIFY if justify else None)
        p.paragraph_format.space_after = Pt(space_after)
        if italic_lead:
            _run(p, italic_lead, self.fs, bold=True)
        _run(p, text, self.fs)


def build_1page_resume(fs=8.2):
    doc = ResumeDoc(fs=fs, top_margin=0.28, bottom_margin=0.28, left_margin=0.38, right_margin=0.38, line_spacing=1.02)
    
    # Header
    p = doc.p(None, WD_ALIGN_PARAGRAPH.CENTER)
    _run(p, "ASMI BASNET", fs + 3.2, bold=True, color=NAVY)
    
    p = doc.p(None, WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_after = Pt(1.2)
    _run(p, "Tallahassee, FL (Hybrid / Relocation Ready) | asmibas18@gmail.com | (314) 732-5088 | linkedin.com/in/asmibasnet | github.com/asmibasnet | asmibasnet.github.io", fs - 0.4)
    
    # Summary
    doc.heading("Professional Summary", space_before=1.5, space_after=0.1)
    doc.plain(
        "Financial Data & Reporting Analyst with extensive higher education institutional analytics experience (Master of Science in Analytics candidate, Saint Louis University; B.E. in Electronics & Communication). Demonstrated expertise supporting university administration and finance teams through recurring financial, operational, and enrollment reporting, budget-to-actual variance analysis, and automated ETL reconciliation. Skilled in Power BI (DAX, Power Query, star-schema data modeling), advanced Microsoft Excel (Power Pivot, complex modeling formulas, VBA), SQL (complex CTEs, window functions, query tuning), and institutional ERP/CRM data pipelines (Banner/Workday-aligned architectures). Proven record reducing reporting query latency by 40% across 8+ databases, validating data integrity across 10,000+ institutional records, establishing automated IV&V quality frameworks, and maintaining comprehensive data dictionaries under FERPA standards. Ready to support the Raymond A. Mason School of Business at William & Mary."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=1.8, space_after=0.1)
    doc.plain(
        "Microsoft Power BI (Advanced DAX, Power Query, Star/Snowflake Schema Modeling, RLS, Service), Tableau, Microsoft Excel (Power Pivot, XLOOKUP, INDEX/MATCH, Dynamic Arrays, Financial Modeling, VBA/Macros), Looker Studio, Executive Dashboards.",
        justify=False, italic_lead="BI, Reporting & Visualization: "
    )
    doc.plain(
        "Budget-to-Actual Analysis, Variance Analysis, Trend Analysis, Fiscal Allocations, Enrollment Metrics, Revenue & Expenditure Tracking, Institutional Research, Accreditation & Ranking Support, KPI Design, Financial Presentations.",
        justify=False, italic_lead="Financial & Operational Analytics: "
    )
    doc.plain(
        "Advanced SQL (Complex CTEs, Analytical Window Functions, Multi-Table Joins, Performance Tuning), PostgreSQL, SQL Server, Snowflake, DuckDB, Python (Pandas, NumPy, SQLAlchemy), R, Automated Data Reconciliation, Data Cleansing, ETL Pipelines.",
        justify=False, italic_lead="SQL & Data Management: "
    )
    doc.plain(
        "Enterprise Systems (Workday/Banner Architecture Familiarity), Independent Verification & Validation (IV&V), Data Quality Assertion, Metadata Repositories, Data Dictionaries, FERPA Compliance, JIRA, Git / GitHub, Agile Sprints.",
        justify=False, italic_lead="Governance, Quality & Systems: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=1.8, space_after=0.1)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Graduate Researcher, SPS Corps – Analytics Initiative", "August 2025 – Present", italic=True, space_before=0.1)
    doc.bullet("Deliver recurring financial, operational, and enrollment analytical reports to university academic leadership, translating complex multi-departmental datasets into actionable insights for strategic decision-making.")
    doc.bullet("Perform routine budget-to-actual variance analyses, revenue/expenditure tracking, and comparative trend evaluations, identifying contributing drivers for material changes across fiscal cycles.")
    doc.bullet("Architect automated data extraction and reconciliation pipelines across 8+ operational databases, refactoring nested queries into modular SQL CTEs and analytical window functions to reduce query latency by 40%.")
    doc.bullet("Develop and maintain interactive Power BI dashboards, advanced DAX measures, and Excel financial reporting models to monitor fiscal allocations, tuition/enrollment benchmarks, and departmental KPIs.")
    doc.bullet("Built automated data quality verification suites (IV&V) that enforce referential integrity and detect schema drift, establishing centralized metadata repositories and comprehensive data dictionaries.")
    doc.bullet("Developed a predictive forecasting engine pairing Prophet with PyTorch LSTM neural networks and automated anomaly detection (flagging deviations beyond ±2.5 standard deviations) to provide leadership with early budget variance alerts.")
    doc.bullet("Collaborate with academic administrators and cross-functional stakeholders in Agile cycles to translate reporting needs into technical specifications; conduct user guidance sessions for 25+ staff members.")
    
    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Data Analyst Intern / BI Developer", "June 2024 – August 2024", italic=True, space_before=0.1)
    doc.bullet("Built and refreshed 6 interactive Power BI and Tableau dashboards backed by star-schema models, consolidating 15+ operational and budgetary KPIs to enable leadership to monitor departmental productivity and milestones.")
    doc.bullet("Authored and tuned SQL extraction queries across institutional student databases (Banner-aligned), executing multi-table joins and data mappings to eliminate recurring reporting bottlenecks.")
    doc.bullet("Engineered automated data profiling and validation scripts in Python (Pandas) and SQL across 10,000+ institutional records, establishing webhook alerts that instantly catch data anomalies and transaction discrepancies.")
    doc.bullet("Enforced strict institutional data governance policies, establishing row-level security (RLS), FERPA compliance routines, and audit trails across sensitive student and financial records.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd.", "Kathmandu, Nepal (Remote)", bold=True, space_before=0.5)
    doc.lr("Data Analyst / Researcher & Systems Analyst", "December 2022 – July 2023", italic=True, space_before=0.5)
    doc.bullet("Extracted, cleansed, mapped, and reconciled multi-source operational and financial transaction datasets (5,000+ records) into standardized 3NF relational schemas using Python (Pandas, SQLAlchemy) and SQL.")
    doc.bullet("Engineered automated validation routines and exception handlers in Python and C/C++, implementing boundary tests and reconciliation scripts to maintain 99.8% pipeline reliability across recurring reporting runs.")
    doc.bullet("Authored comprehensive technical documentation, data definitions, and reporting procedures to ensure reporting continuity and governance across teams.")

    # Featured Projects
    doc.heading("Featured Financial & Data Architecture Projects", space_before=1.8, space_after=0.1)
    
    # SCD Type-2 Warehouse
    doc.lr("SCD Type-2 Dimensional Data Warehouse & Analytics Engine", "Python, SQL, DuckDB, Star Schema, Kimball Modeling", bold=True, space_before=0.4)
    doc.bullet("Designed a Kimball dimensional data warehouse implementing Slowly Changing Dimensions (SCD Type-2) with surrogate keys and date tracking across 10,000+ records, preserving 100% point-in-time audit lineage for operational and financial analysis.")
    doc.bullet("Authored complex SQL window functions (LAG, LEAD, DENSE_RANK) to calculate rolling cohort retention, promotion velocity, and budget allocation trends, feeding an interactive analytics reporting dashboard.")
    
    # CareGap
    doc.lr("CareGap — Decision Support & Operational Analytics Engine", "Python, DuckDB, PostgreSQL, Power BI, Streamlit", bold=True, space_before=0.4)
    doc.bullet("Engineered a high-performance analytics data engine over 34,000+ longitudinal records, integrating columnar DuckDB for sub-second analytical aggregations with PostgreSQL for transactional and relational integrity.")
    doc.bullet("Delivered 6 interactive Power BI / Streamlit views tracking quality performance measures and operational gap closure rates, implementing automated variance tracking and comparative benchmark analyses.")

    # Education & Certifications
    doc.heading("Education & Professional Development", space_before=1.8, space_after=0.1)
    doc.lr("Saint Louis University | Master of Science in Analytics (STEM)", "St. Louis, MO | Expected May 2026", bold=True, space_before=0.4)
    doc.lr("Tribhuvan University | Bachelor of Engineering in Electronics & Communication (STEM)", "Kathmandu, Nepal | Conferred 2023", bold=True, space_before=0.1)
    doc.plain("Certifications: AI Engineer for Data Scientists Associate (DataCamp) · Building with Claude & Anthropic API (Prompt Caching & Tools) · Project Management & Agile Certification (GlobalShala & SLU).", space_after=0.1)

    return doc.d


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
    ps_file = "convert_wm_temp.ps1"
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps_script)
    res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file], capture_output=True, text=True)
    if os.path.exists(ps_file):
        os.remove(ps_file)
    print("Conversion output:", res.stdout, res.stderr)


if __name__ == "__main__":
    out_dir = r"c:\Users\Shayara Basnet\Downloads\portfolio"
    
    docx_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_William_Mary_Financial_Data_Analyst.docx")
    pdf_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_William_Mary_Financial_Data_Analyst.pdf")
    
    print("Building 1-page docx...")
    d1 = build_1page_resume(fs=8.2)
    d1.save(docx_1p)
    print("Converting 1-page to PDF...")
    convert_to_pdf(docx_1p, pdf_1p)
    if os.path.exists(pdf_1p):
        p1 = len(fitz.open(pdf_1p))
        print(f"1-Page PDF generated: {pdf_1p} (Page count: {p1})")
