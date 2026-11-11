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
NAVY = RGBColor(0x00, 0x33, 0x66)  # Molina Deep Blue / Professional Navy

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
        self.d.core_properties.title = "Asmi Basnet - Molina Healthcare Senior Data Analyst Resume"

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
    _run(p, "Tallahassee, FL (Local Resident · Florida Network Ready) | asmibas18@gmail.com | (314) 732-5088 | linkedin.com/in/asmibasnet | github.com/asmibasnet | asmibasnet.github.io", fs - 0.4)
    
    # Summary
    doc.heading("Professional Summary", space_before=1.5, space_after=0.1)
    doc.plain(
        "Based locally in Tallahassee, Florida, I am a data and analytics specialist and graduate researcher (Saint Louis University M.S. in Analytics candidate; B.E. in Electronics & Communication) with over three years of applied experience working with healthcare datasets, operational reporting, and data reconciliation. My work centers on turning complex operational, claims, and administrative records into dependable, easy-to-digest reporting that helps leadership understand utilization patterns, track provider network performance, and identify cost-containment opportunities. I write advanced SQL to clean and query data across relational databases, build interactive dashboards in Power BI and Tableau, and develop Excel financial and utilization models. Having developed predictive models across 34,000+ patient records and instituted automated validation checks across multi-table environments, I focus on delivering accurate, auditable insights that clinical and business teams can trust."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=1.8, space_after=0.1)
    doc.plain(
        "Healthcare Utilization Metrics, Cost-Containment Analytics, Medical Economics, Provider Network Performance, Claims, Pharmacy & Lab Data Analysis, Quality Measures & Performance Indicators, Risk Stratification, State Regulatory Submissions (AHCA Alignment).",
        justify=False, italic_lead="Healthcare & Operational Analytics: "
    )
    doc.plain(
        "Microsoft Power BI (Advanced DAX, Power Query, Star/Snowflake Schema Modeling, Row-Level Security), Tableau, Microsoft Excel (Advanced Formulas, Power Pivot, Dynamic Arrays, Financial & Utilization Modeling, VBA), SSRS Architecture Familiarity, Executive Presentations.",
        justify=False, italic_lead="Reporting & Business Intelligence: "
    )
    doc.plain(
        "Advanced SQL (Multi-Table Joins, Complex CTEs, Analytical Window Functions, Query Performance Optimization), PostgreSQL, SQL Server, Snowflake, DuckDB, Python (Pandas, NumPy, Scikit-Learn, SQLAlchemy), R, Automated ETL Reconciliation & Data Cleansing.",
        justify=False, italic_lead="Data Engineering & SQL: "
    )
    doc.plain(
        "Automated Quality Assurance (IV&V Frameworks), Referential Integrity Checking, Schema Drift Detection, Data Dictionaries & Metadata Repositories, HIPAA & Privacy Compliance, Project Management & Agile Sprints (Certified), JIRA, Git / GitHub.",
        justify=False, italic_lead="Data Governance & Methodologies: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=1.8, space_after=0.1)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Graduate Researcher, SPS Corps – Analytics Initiative", "August 2025 – Present", italic=True, space_before=0.1)
    doc.bullet("Partnered with university administration, department heads, and cross-functional teams to deliver recurring utilization, operational, and budgetary reports that inform strategic planning and resource allocation.")
    doc.bullet("Re-engineered data extraction queries across 8+ operational databases by replacing slow, nested subqueries with modular SQL common table expressions and analytical window functions (ROW_NUMBER, DENSE_RANK, LAG/LEAD), reducing query runtimes by 40%.")
    doc.bullet("Designed and maintained centralized Power BI dashboards and advanced Excel models to track key operational indicators, fiscal allocations, and workload throughput for non-technical stakeholders and executive reviewers.")
    doc.bullet("Built an automated data-quality assertion framework (IV&V) that routinely validates referential integrity and flags unexpected schema changes, preventing corrupted or mismatched records from entering recurring reporting layers.")
    doc.bullet("Developed a time-series forecasting model combining Prophet with LSTM neural networks, incorporating automated anomaly checks to alert management whenever spending or utilization drifted beyond ±2.5 standard deviations.")
    doc.bullet("Authored clear technical documentation, report definitions, and standardized data dictionaries, and conducted hands-on walkthroughs for over 25 administrative staff members to ensure continuity across reporting cycles.")
    
    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Data Analyst Intern / BI Developer", "June 2024 – August 2024", italic=True, space_before=0.1)
    doc.bullet("Built and published a suite of 6 interactive Power BI and Tableau dashboards backed by clean dimensional star schemas, consolidating 15+ operational and performance metrics into single-pane views for program directors.")
    doc.bullet("Extracted and reconciled high-volume transactional data from enterprise databases using multi-table joins and SQL views, cutting down recurring monthly reporting delays and eliminating manual spreadsheet consolidation.")
    doc.bullet("Created automated data profiling and validation scripts in Python (Pandas) and SQL across 10,000+ institutional records, wiring up webhook notifications that immediately surface data discrepancies or missing attributes.")
    doc.bullet("Established strict row-level security (RLS) rules and audit logs across reporting datasets to safeguard sensitive records in full compliance with institutional privacy and data governance policies.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd.", "Kathmandu, Nepal (Remote)", bold=True, space_before=0.5)
    doc.lr("Data Analyst / Researcher & Systems Analyst", "December 2022 – July 2023", italic=True, space_before=0.5)
    doc.bullet("Extracted, cleaned, and mapped messy multi-source survey and operational transaction datasets (5,000+ records) into standardized relational schemas using Python (Pandas, SQLAlchemy) and SQL to ensure consistent reporting.")
    doc.bullet("Wrote automated data validation scripts with rigorous boundary testing, exception handling, and reconciliation checks in Python and C/C++, maintaining a 99.8% pipeline reliability rate across regular reporting schedules.")
    doc.bullet("Documented data dictionaries, transformation logic, and sampling protocols to establish institutional governance standards and make reporting repeatable across distributed teams.")

    # Featured Projects
    doc.heading("Featured Healthcare Data & Analytics Projects", space_before=1.8, space_after=0.1)
    
    # CareGap
    doc.lr("CareGap — Clinical Decision Support & Population Health Analytics Engine", "Python, DuckDB, Scikit-Learn, XGBoost, SHAP, Power BI, Streamlit", bold=True, space_before=0.4)
    doc.bullet("Developed a healthcare analytics platform that queries 34,000+ longitudinal synthetic patient records with sub-second response times using columnar DuckDB, training ensemble models (XGBoost, Random Forest) to identify clinical care gaps and cost drivers.")
    doc.bullet("Integrated SHAP TreeExplainer to give clinicians per-patient transparent risk explanations, paired with 6 interactive Power BI and Streamlit views that track quality measures, patient follow-up rates, and preventive care utilization.")

    # WashU Datathon
    doc.lr("WashU Clinical Diabetes Risk Prediction & Healthcare Utilization Datathon", "Python, Scikit-Learn, LightGBM, XGBoost, Hyperopt", bold=True, space_before=0.4)
    doc.bullet("Built an end-to-end predictive pipeline across 62,000+ clinical records, engineering 59 domain features (such as visit frequency, prescription counts, and glycemic fluctuations) to identify patient cohorts at elevated risk of high utilization.")
    doc.bullet("Evaluated models using stratified cross-validation and hyperparameter optimization, achieving an AUC of 0.8520 to help clinical teams prioritize interventions and reduce avoidable hospital readmissions.")

    # SCD Type-2 Warehouse
    doc.lr("SCD Type-2 Dimensional Data Warehouse & Analytics Engine", "Python, SQL, DuckDB, Star Schema, Kimball Modeling", bold=True, space_before=0.4)
    doc.bullet("Designed a Kimball-style dimensional data warehouse with Slowly Changing Dimensions (SCD Type-2) across 10,000+ records, using surrogate keys and valid-from/valid-to date tracking to guarantee a complete, point-in-time audit trail for operational and financial analysis.")

    # Education & Certifications
    doc.heading("Education & Professional Credentials", space_before=1.8, space_after=0.1)
    doc.lr("Saint Louis University | Master of Science in Analytics (STEM)", "St. Louis, MO | Expected May 2026", bold=True, space_before=0.4)
    doc.lr("Tribhuvan University | Bachelor of Engineering in Electronics & Communication (STEM)", "Kathmandu, Nepal | Conferred 2023", bold=True, space_before=0.1)
    doc.plain("Certifications: Project Management & Agile Certification (GlobalShala & SLU) · AI Engineer for Data Scientists Associate (DataCamp) · Building with Claude & Anthropic API (Prompt Caching & Tools).", space_after=0.1)

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
    ps_file = "convert_molina_temp.ps1"
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps_script)
    res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file], capture_output=True, text=True)
    if os.path.exists(ps_file):
        os.remove(ps_file)
    print("Conversion output:", res.stdout, res.stderr)


if __name__ == "__main__":
    out_dir = r"c:\Users\Shayara Basnet\Downloads\portfolio"
    
    docx_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Molina_Healthcare_Senior_Data_Analyst.docx")
    pdf_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Molina_Healthcare_Senior_Data_Analyst.pdf")
    
    print("Building 1-page docx...")
    d1 = build_1page_resume(fs=8.2)
    d1.save(docx_1p)
    print("Converting 1-page to PDF...")
    convert_to_pdf(docx_1p, pdf_1p)
    if os.path.exists(pdf_1p):
        p1 = len(fitz.open(pdf_1p))
        print(f"1-Page PDF generated: {pdf_1p} (Page count: {p1})")
