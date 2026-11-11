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
NAVY = RGBColor(0x1B, 0x36, 0x5D)  # Federal / Civic Tech Deep Navy

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
        self.d.core_properties.title = "Asmi Basnet - Peregrine Advisors Data Scientist Resume"

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
    _run(p, "Washington, DC Metro Ready · Tallahassee, FL | asmibas18@gmail.com | (314) 732-5088 | linkedin.com/in/asmibasnet | github.com/asmibasnet | asmibasnet.github.io", fs - 0.4)
    
    # Summary
    doc.heading("Professional Summary", space_before=1.5, space_after=0.1)
    doc.plain(
        "Applied Data Scientist and Graduate Researcher (Saint Louis University M.S. in Analytics candidate; B.E. in Electronics & Communication) with over three years of experience building statistical models, automated analytical pipelines, and document-intelligence systems. I specialize in carrying complex analytical questions from raw, messy data into defensible, reproducible answers that withstand scrutiny. Experienced working across Python, SQL, and cloud environments (AWS S3, Aurora/PostgreSQL architectures), with a strong background hardening exploratory notebooks into reviewed, tested pipelines. Proven track record benchmarking supervised models (XGBoost, LightGBM, Random Forest) against rigorous baselines, developing time-series anomaly detection engines, architecting auditable NLP/RAG document verification suites (PolicyGuard), and communicating findings through clear visual narratives for institutional decision-makers. Eager to support mission-critical federal analytics at Peregrine Advisors."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=1.8, space_after=0.1)
    doc.plain(
        "Supervised & Unsupervised Machine Learning, Statistical Hypothesis Testing, Time-Series Forecasting (Prophet, LSTM, ARIMA), Classification, Anomaly Detection, Model Validation (K-Fold Cross-Validation, ROC-AUC, PR-AUC, F1), Error Analysis, Baseline Benchmarking, Feature Engineering.",
        justify=False, italic_lead="Statistical Modeling & ML: "
    )
    doc.plain(
        "Python (Pandas, NumPy, Scikit-Learn, PyTorch, SQLAlchemy, FastAPI), SQL (Complex CTEs, Analytical Window Functions, Multi-Table Joins, Query Optimization), PostgreSQL, DuckDB, Snowflake, SQLite, C/C++, R, Statistical Analysis (Regression, ANOVA, Chi-Square).",
        justify=False, italic_lead="Programming & Query Languages: "
    )
    doc.plain(
        "Amazon Web Services (AWS) Data Platforms Familiarity (Amazon S3, Amazon Aurora / Relational Databases, Cloud Pipelines), Containerization (Docker, Docker Compose), Linux/Bash Scripting, Git / GitHub Version Control, Automated Testing & Verification Loops.",
        justify=False, italic_lead="Cloud & Data Engineering (AWS): "
    )
    doc.plain(
        "Foundation Models & LLM Integration (Amazon Bedrock / Claude & OpenAI APIs), Document Parsing, Offline RAG Suites, Citation Verification, Independent Verification & Validation (IV&V), Data Governance, Power BI, Tableau, Excel (Modeling, VBA), Technical Writing.",
        justify=False, italic_lead="AI, Document Intelligence & Governance: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=1.8, space_after=0.1)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Graduate Researcher, SPS Corps – Analytics Initiative", "August 2025 – Present", italic=True, space_before=0.1)
    doc.bullet("Frame core analytical questions with departmental leaders, turning high-level operational inquiries into testable statistical designs and automated data pipelines across 8+ operational databases.")
    doc.bullet("Refactored complex SQL queries by replacing nested subqueries with modular common table expressions and window functions (ROW_NUMBER, DENSE_RANK, LAG/LEAD), reducing query runtimes by 40% and standardizing data definitions.")
    doc.bullet("Developed a hybrid forecasting engine combining Facebook Prophet with PyTorch LSTM neural networks, integrating automated anomaly detection (flagging deviations beyond ±2.5σ) to deliver defensible quarterly budget and workload projections.")
    doc.bullet("Hardened exploratory analysis scripts into reviewed, reproducible pipelines and implemented automated data quality checks (IV&V) that enforce referential integrity and catch schema drift before records reach reporting layers.")
    doc.bullet("Designed interactive Power BI dashboards and executive briefings, presenting complex statistical trends and model limitations in accessible plain language to university leadership.")
    doc.bullet("Authored rigorous technical runbooks, data dictionaries, and model evaluation reports, conducting hands-on training sessions for over 25 administrative staff members to ensure organizational continuity.")
    
    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Data Analyst Intern / BI Developer", "June 2024 – August 2024", italic=True, space_before=0.1)
    doc.bullet("Built and deployed 6 interactive Power BI and Tableau dashboards backed by clean dimensional star schemas, consolidating 15+ institutional KPIs to give administrators clear visibility into operational performance.")
    doc.bullet("Engineered automated data profiling, cleaning, and reconciliation scripts in Python (Pandas) and SQL across 10,000+ transaction records, setting up automated webhook alerts to immediately flag data discrepancies.")
    doc.bullet("Enforced strict institutional data governance standards, establishing row-level security (RLS) and audit logging routines to ensure sensitive academic records remained fully compliant with privacy regulations.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd.", "Kathmandu, Nepal (Remote)", bold=True, space_before=0.5)
    doc.lr("Data Analyst / Researcher & Systems Analyst", "December 2022 – July 2023", italic=True, space_before=0.5)
    doc.bullet("Extracted, cleansed, and normalized multi-source operational and survey datasets (5,000+ records) into standardized relational schemas using Python (Pandas, SQLAlchemy) and SQL, guaranteeing complete data consistency.")
    doc.bullet("Engineered automated validation scripts and exception-handling routines in Python and C/C++, implementing boundary tests and audit checks to maintain a 99.8% pipeline reliability rate across recurring analysis runs.")
    doc.bullet("Authored formal methodology documentation, data dictionaries, and statistical audit reports (SPSS/R regression, ANOVA, chi-square) to ensure transparent, peer-reviewable research standards.")

    # Featured Projects
    doc.heading("Featured Applied Machine Learning & Document Intelligence Projects", space_before=1.8, space_after=0.1)
    
    # PolicyGuard
    doc.lr("PolicyGuard — Document Intelligence & Automated Citation Verification Suite", "Python, Ollama, ChromaDB, FastAPI, Docker, Streamlit", bold=True, space_before=0.4)
    doc.bullet("Engineered a document-intelligence compliance application designed to analyze complex regulatory and policy corpora locally with zero external API exposure, keeping sensitive agency documents strictly protected on premises.")
    doc.bullet("Implemented automated chain-of-thought citation verification that evaluates retrieved context sentence-by-sentence, flagging contradictory requirements (FLAG_CONFLICT) and missing evidence (CLARIFY) to produce an auditable governance health score.")

    # WashU Datathon
    doc.lr("WashU Clinical Analytics & Predictive Risk Modeling Datathon", "Python, Scikit-Learn, LightGBM, XGBoost, Hyperopt", bold=True, space_before=0.4)
    doc.bullet("Built an end-to-end predictive modeling pipeline across 62,000+ records, engineering 59 domain features (visit frequencies, variance metrics, interaction terms) to identify elevated-risk patient cohorts.")
    doc.bullet("Benchmarked Logistic Regression, Random Forest, LightGBM, and XGBoost against baseline models using stratified K-fold cross-validation and hyperparameter optimization, achieving an AUC of 0.8520 with documented error distributions.")

    # SCD Type-2 Warehouse
    doc.lr("SCD Type-2 Dimensional Data Warehouse & Reproducible Analytics Engine", "Python, SQL, DuckDB, Star Schema, Kimball Modeling", bold=True, space_before=0.4)
    doc.bullet("Designed a Kimball dimensional data warehouse implementing Slowly Changing Dimensions (SCD Type-2) across 10,000+ records with surrogate keys and date tracking, guaranteeing a complete point-in-time audit trail for defensible retrospective analysis.")

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
    ps_file = "convert_peregrine_temp.ps1"
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(ps_script)
    res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file], capture_output=True, text=True)
    if os.path.exists(ps_file):
        os.remove(ps_file)
    print("Conversion output:", res.stdout, res.stderr)


if __name__ == "__main__":
    out_dir = r"c:\Users\Shayara Basnet\Downloads\portfolio"
    
    docx_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Peregrine_Advisors_Data_Scientist.docx")
    pdf_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Peregrine_Advisors_Data_Scientist.pdf")
    
    print("Building 1-page docx...")
    d1 = build_1page_resume(fs=8.2)
    d1.save(docx_1p)
    print("Converting 1-page to PDF...")
    convert_to_pdf(docx_1p, pdf_1p)
    if os.path.exists(pdf_1p):
        p1 = len(fitz.open(pdf_1p))
        print(f"1-Page PDF generated: {pdf_1p} (Page count: {p1})")
