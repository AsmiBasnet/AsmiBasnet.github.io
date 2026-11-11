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
NAVY = RGBColor(0x1F, 0x3A, 0x5F)

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
        self.d.core_properties.title = "Asmi Basnet - HNTB BI Developer II Resume"

    def p(self, style=None, align=None):
        p = self.d.add_paragraph(style=style) if style else self.d.add_paragraph()
        if align is not None:
            p.alignment = align
        return p

    def heading(self, text, space_before=1.5, space_after=0.1):
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
    _run(p, "Tallahassee, FL (Local Resident · On-Site Ready) | asmibas18@gmail.com | (314) 732-5088 | linkedin.com/in/asmibasnet | github.com/asmibasnet | asmibasnet.github.io", fs - 0.4)
    
    # Summary
    doc.heading("Professional Summary", space_before=1.5, space_after=0.1)
    doc.plain(
        "BI Developer and Data Analytics Specialist based locally in Tallahassee, FL (Master of Science in Analytics candidate, Saint Louis University; B.E. in Electronics & Communication) with 3+ years of experience engineering enterprise BI dashboards, Kimball dimensional models, automated ETL workflows, and GIS geospatial intelligence solutions. Proficient in Power BI (DAX, Power Query, star-schema modeling), Tableau, advanced SQL (complex CTEs, window functions, query tuning), and GIS spatial integration (GeoJSON, Leaflet.js). Proven track record optimizing queries across 8+ databases to slash latency by 40%, architecting SCD Type-2 dimensional warehouses across 10,000+ records, and establishing automated data quality assertion frameworks (IV&V). Ready to support HNTB's digital infrastructure and transportation engineering projects across Florida."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=1.8, space_after=0.1)
    doc.plain(
        "Microsoft Power BI (Advanced DAX, Power Query, Star/Snowflake Schemas, RLS, Incremental Refresh, Service), Tableau (LOD Expressions, Calculated Fields), Looker Studio, Excel (Power Pivot, VBA/Macros), Executive Reporting, KPI Design.",
        justify=False, italic_lead="BI & Visualization: "
    )
    doc.plain(
        "Advanced SQL (Complex CTEs, Analytical Window Functions, Multi-Table Joins, Stored Procedures, Performance Tuning), PostgreSQL, SQL Server, Snowflake, DuckDB, BigQuery, SQLite, Python (Pandas, NumPy, SQLAlchemy), R.",
        justify=False, italic_lead="SQL & Data Engineering: "
    )
    doc.plain(
        "GIS & Geospatial Data Integration, GeoJSON, Spatial Feeds & Highway Closure Telemetry, Leaflet.js, Kimball Dimensional Modeling, Star/Snowflake Schema Design, Slowly Changing Dimensions (SCD Type-2), 3NF Relational Schemas, Data Lineage.",
        justify=False, italic_lead="GIS & Dimensional Modeling: "
    )
    doc.plain(
        "Automated Data Quality Assurance (IV&V), Referential Integrity Checking, Schema Drift Detection, Metadata Repositories, Data Dictionaries, Technical Documentation, Unit Testing, Agile/Scrum Sprints, JIRA, Git / GitHub.",
        justify=False, italic_lead="Governance & Best Practices: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=1.8, space_after=0.1)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Graduate Researcher, SPS Corps – Analytics Initiative", "August 2025 – Present", italic=True, space_before=0.1)
    doc.bullet("Architected enterprise data pipelines across 8+ operational databases, refactoring nested subqueries into modular SQL CTEs and analytical window functions (ROW_NUMBER, DENSE_RANK, LAG/LEAD) to reduce query latency by 40%.")
    doc.bullet("Built dynamic DAX measures, dimensional star-schema models, and interactive Power BI reports to present fiscal allocations, enrollment trends, and operational KPIs directly to university executive leadership.")
    doc.bullet("Implemented automated data quality verification suites (IV&V) that enforce referential integrity and detect schema drift, establishing centralized metadata repositories and comprehensive data dictionaries.")
    doc.bullet("Engineered automated ETL workflows and reconciliation checks, eliminating manual data entry cycles and accelerating weekly reporting turnarounds.")
    doc.bullet("Developed a hybrid forecasting engine pairing Prophet with PyTorch LSTM neural networks and automated anomaly detection (flagging variances beyond ±2.5 standard deviations) to provide leadership with predictive quarterly budget insights.")
    doc.bullet("Led requirements workshops with departmental stakeholders in Agile sprints to define reporting goals, authoring technical specifications, user stories, and conducting training sessions for 25+ staff members.")
    
    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Data Analyst Intern / BI Developer", "June 2024 – August 2024", italic=True, space_before=0.1)
    doc.bullet("Built and deployed 6 interactive Power BI and Tableau dashboards backed by star-schema data models, consolidating 15+ operational KPIs to enable leadership to monitor departmental productivity and project milestones.")
    doc.bullet("Authored and tuned SQL extraction queries across enterprise student databases using multi-table joins and window functions, reducing recurring reporting lag.")
    doc.bullet("Engineered automated data profiling and validation scripts in Python (Pandas) and SQL across 10,000+ institutional records, establishing webhook alerts that instantly flag data drift and transaction discrepancies.")
    doc.bullet("Enforced strict institutional data governance rules, establishing row-level security (RLS) and audit routines across sensitive reporting models.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd.", "Kathmandu, Nepal (Remote)", bold=True, space_before=0.5)
    doc.lr("Data Analyst / Researcher & Systems Analyst", "December 2022 – July 2023", italic=True, space_before=0.5)
    doc.bullet("Extracted, cleansed, and mapped multi-source operational and telemetry datasets (5,000+ records) into standardized 3NF relational schemas using Python (Pandas, SQLAlchemy) and SQL, guaranteeing complete data fidelity.")
    doc.bullet("Validated algorithmic logic and automated backend processing scripts in Python and C/C++, implementing robust exception handling, boundary tests, and data reconciliation routines to maintain 99.8% pipeline reliability.")
    doc.bullet("Authored formal technical documentation detailing data dictionaries, validation rules, and sampling methodologies to establish organizational data governance standards.")

    # Featured Projects
    doc.heading("Featured BI, GIS & Dimensional Modeling Projects", space_before=1.8, space_after=0.1)
    
    # Nepal Disaster Watch
    doc.lr("Nepal Disaster Watch — Geospatial Intelligence & Real-Time Telemetry Platform", "TypeScript, GeoJSON, Leaflet.js, REST APIs", bold=True, space_before=0.4)
    doc.bullet("Architected a real-time GIS command platform integrating GeoJSON spatial feeds, live highway closures, river sensors, and meteorological data across all 77 districts of Nepal for OpenAI's WebMCP Challenge.")
    doc.bullet("Engineered resilient status adapters (live, snapshot, unreachable) and dynamic district cross-referencing tools, eliminating silent telemetry failures and enabling interactive spatial risk triage.")
    
    # SCD Type-2 Warehouse
    doc.lr("SCD Type-2 Dimensional Data Warehouse & Analytics Engine", "Python, SQL, DuckDB, Star Schema, Kimball Modeling", bold=True, space_before=0.4)
    doc.bullet("Designed a Kimball dimensional data warehouse with SCD Type-2 tracking (surrogate keys, effective/end dates, is_current flags) across 10,000+ records, preserving 100% point-in-time audit lineage.")
    doc.bullet("Authored complex SQL window functions (LAG, LEAD, DENSE_RANK) to calculate rolling retention, promotion velocity, and attrition metrics, feeding an interactive analytics dashboard.")
    
    # CareGap
    doc.lr("CareGap — Clinical Decision Support & Population Health Analytics Engine", "Python, DuckDB, PostgreSQL, Power BI, Streamlit", bold=True, space_before=0.4)
    doc.bullet("Built a healthcare analytics data engine over 34,000+ longitudinal patient records, using columnar DuckDB for sub-second aggregations and PostgreSQL for transactional integrity.")
    doc.bullet("Delivered 6 interactive Power BI / Streamlit views tracking quality measures and care-gap closure rates, incorporating CDC growth-chart Z-score logic for pediatric cohorts.")

    # Education & Certifications
    doc.heading("Education & Certifications", space_before=1.8, space_after=0.1)
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
    ps_file = "convert_temp.ps1"
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps_script)
    res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file], capture_output=True, text=True)
    if os.path.exists(ps_file):
        os.remove(ps_file)
    print("Conversion output:", res.stdout, res.stderr)


if __name__ == "__main__":
    out_dir = r"c:\Users\Shayara Basnet\Downloads\portfolio"
    
    docx_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_HNTB_BI_Developer.docx")
    pdf_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_HNTB_BI_Developer.pdf")
    
    print("Building 1-page docx...")
    d1 = build_1page_resume(fs=8.2)
    d1.save(docx_1p)
    print("Converting 1-page to PDF...")
    convert_to_pdf(docx_1p, pdf_1p)
    if os.path.exists(pdf_1p):
        p1 = len(fitz.open(pdf_1p))
        print(f"1-Page PDF generated: {pdf_1p} (Page count: {p1})")
