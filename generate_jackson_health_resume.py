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
    def __init__(self, fs=8.2, top_margin=0.30, bottom_margin=0.30, left_margin=0.38, right_margin=0.38, line_spacing=1.02):
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
        self.d.core_properties.title = "Asmi Basnet - Jackson Health System Machine Learning Engineer Resume"

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
    _run(p, "Miami, FL (In-State Florida Candidate / Relocating to Miami) · Tallahassee, FL | asmibas18@gmail.com | (314) 732-5088 | linkedin.com/in/asmibasnet | github.com/asmibasnet | asmibasnet.github.io", fs - 0.4)
    
    # Summary
    doc.heading("Professional Summary", space_before=1.5, space_after=0.1)
    doc.plain(
        "Machine Learning Engineer and Applied AI Specialist (Master of Science in Analytics candidate, Saint Louis University; B.E. in Electronics & Communication) with 3+ years of experience engineering production predictive ML models, generative AI architectures, and intelligent clinical applications. Proven expertise analyzing complex structured and unstructured healthcare datasets (34,000+ EHR records, clinical notes, imaging), training ensemble classifiers (XGBoost, Random Forest, PyTorch), and deploying Explainable AI (SHAP TreeExplainer) to ensure clinical transparency. Skilled in architecting privacy-first, offline retrieval-augmented generation (RAG) suites with automated citation verification (PolicyGuard), time-series forecasting (Prophet + LSTM), and modular SQL across enterprise databases. Florida resident eager to deliver high-impact, compliant, and reliable AI/ML solutions for Jackson Health System."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=1.8, space_after=0.1)
    doc.plain(
        "Supervised & Unsupervised ML (Scikit-Learn, XGBoost, LightGBM, Random Forest, Lasso), Deep Learning (PyTorch, Torchvision, CNNs, LSTM), Explainable AI (SHAP TreeExplainer), Time-Series Forecasting (Prophet, ARIMA), K-Fold Cross-Validation, Hyperparameter Optimization (Hyperopt).",
        justify=False, italic_lead="Machine Learning & Modeling: "
    )
    doc.plain(
        "Retrieval-Augmented Generation (RAG), Local LLM Inference (Ollama: Qwen 2.5, Llama 3, Phi-3), OpenAI GPT-4o API, Anthropic Claude API (Prompt Caching & Tools), Vector Databases (ChromaDB, FAISS), Citation Verification, Clinical Guardrails, Agentic Workflows (Agno, WebMCP).",
        justify=False, italic_lead="Generative AI & LLMs: "
    )
    doc.plain(
        "Python (FastAPI, Flask, Pandas, NumPy, Scikit-Learn, PyTorch, SQLAlchemy), SQL (Snowflake, PostgreSQL, DuckDB, SQL Server, BigQuery: CTEs, Window Functions), C/C++, TypeScript, Linux/Bash CLI, Git / GitHub, Docker, Docker Compose, RESTful APIs, Webhooks.",
        justify=False, italic_lead="Programming & Infrastructure: "
    )
    doc.plain(
        "EHR Data Modeling, Clinical Informatics (Synthea EHR), DICOM / Medical Imaging, Automated Data Quality Assurance (IV&V), Schema Drift Detection, HIPAA / FERPA Privacy Compliance, Audit Logging, Business Intelligence (Power BI, DAX, Tableau).",
        justify=False, italic_lead="Healthcare Data & Governance: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=1.8, space_after=0.1)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Graduate Machine Learning & AI Researcher (SPS-Corps)", "August 2025 – Present", italic=True, space_before=0.1)
    doc.bullet("Architected and deployed an air-gapped, privacy-first offline RAG compliance suite in Python (FastAPI, ChromaDB, Ollama) that runs local quantized LLMs to process sensitive policy and clinical guidelines with zero external data exposure.")
    doc.bullet("Engineered a chain-of-thought verification engine evaluating retrieved context sentence-by-sentence, flagging contradictory requirements (FLAG_CONFLICT), and requesting missing context (CLARIFY) to automate governance audits.")
    doc.bullet("Built a hybrid time-series forecasting engine combining Facebook Prophet with PyTorch LSTM neural networks and automated anomaly detection (flagging variances beyond ±2.5σ) to deliver operational demand and resource predictions.")
    doc.bullet("Optimized modular SQL queries across 8+ operational databases using complex CTEs and analytical window functions (ROW_NUMBER, DENSE_RANK), standardizing data dictionaries and reducing query execution latency by 40%.")
    doc.bullet("Implemented automated data quality verification suites (IV&V) enforcing referential integrity, schema drift detection, and rigorous data governance standards aligned with institutional and privacy compliance frameworks.")
    doc.bullet("Collaborated directly with cross-functional leadership in Agile sprints to translate operational requirements into technical specifications, authoring runbooks and conducting user training workshops for 25+ staff members.")
    
    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Data & AI Solutions Analyst Intern", "June 2024 – August 2024", italic=True, space_before=0.1)
    doc.bullet("Engineered automated data profiling, cleaning, and validation pipelines in Python (Pandas) and SQL across 10,000+ transactional records, establishing automated webhook alerts that instantly flag schema drift and data discrepancies.")
    doc.bullet("Built and deployed 6 interactive Power BI and Tableau executive dashboards backed by star-schema data models, tracking 15+ operational KPIs to enable leadership to monitor resource allocation, productivity, and project milestones.")
    doc.bullet("Enforced strict institutional data governance and privacy rules, building row-level data quality validation checks across reporting layers.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd. (Mayahold Group)", "Remote / Kathmandu, Nepal", bold=True, space_before=0.5)
    doc.lr("Systems & Data Analyst", "December 2022 – July 2023", italic=True, space_before=0.1)
    doc.bullet("Extracted, cleansed, and normalized multi-source operational and telemetry datasets (5,000+ records) into standardized 3NF relational schemas using Python (Pandas, SQLAlchemy) and SQL, guaranteeing complete data fidelity.")
    doc.bullet("Validated algorithmic logic and automated processing scripts in Python and C/C++, implementing robust exception handling, boundary tests, and audit logging to maintain 99.8% pipeline reliability.")

    # Featured Projects
    doc.heading("Featured Clinical AI & Machine Learning Projects", space_before=1.8, space_after=0.1)
    
    # CareGap
    doc.lr("CareGap — Clinical Decision Support & Population Health Analytics", "Python, DuckDB, Scikit-Learn, XGBoost, SHAP, FAISS, Streamlit", bold=True, space_before=0.4)
    doc.bullet("Developed a full-stack clinical decision platform querying 34,000+ longitudinal synthetic EHR patient records with sub-second latency using DuckDB columnar storage, training an ensemble classifier (XGBoost, Random Forest, Lasso) for patient risk stratification.")
    doc.bullet("Integrated SHAP TreeExplainer for per-patient transparent risk attribution alongside a dual-tier RAG clinical assistant (FAISS + local Ollama with failover to Gemini) featuring strict guardrails that block unauthorized dosage recommendations.")
    
    # PolicyGuard
    doc.lr("PolicyGuard — Enterprise Offline RAG & Compliance Intelligence Suite", "Python, Ollama (Qwen 2.5 / Llama 3), ChromaDB, FastAPI, Docker", bold=True, space_before=0.4)
    doc.bullet("Engineered an air-gapped, privacy-first RAG application running entirely on local quantized LLMs via Ollama to analyze confidential enterprise documentation and clinical guidelines with zero third-party API exposure and zero data leakage.")
    doc.bullet("Implemented automated chain-of-thought citation verification and dynamic cosine similarity guardrails, evaluating retrieved context sentence-by-sentence to detect contradictory clauses (FLAG_CONFLICT) and eliminate manual review overhead.")
    
    # WashU Datathon
    doc.lr("WashU Clinical Diabetes Risk Prediction Datathon", "Python, Scikit-Learn, XGBoost, LightGBM, Hyperopt", bold=True, space_before=0.4)
    doc.bullet("Engineered an end-to-end ML pipeline on 62,000+ clinical EHR records to predict uncontrolled HbA1c, engineering visit-frequency, medication-count, and glycemic-volatility features.")
    doc.bullet("Benchmarked Logistic Regression, Random Forest, LightGBM, and XGBoost using stratified K-fold cross-validation and hyperparameter tuning, optimizing ROC-AUC, PR-AUC, and F1 to minimize clinical false negatives.")

    # Education & Certifications
    doc.heading("Education & Certifications", space_before=1.8, space_after=0.1)
    doc.lr("Saint Louis University | Master of Science in Analytics (STEM)", "St. Louis, MO | Expected May 2026", bold=True, space_before=0.4)
    doc.lr("Tribhuvan University | Bachelor of Engineering in Electronics & Communication (STEM)", "Kathmandu, Nepal | Conferred 2023", bold=True, space_before=0.1)
    doc.plain("Certifications: AI Engineer for Data Scientists Associate (DataCamp: PyTorch Deep Learning & Scikit-Learn ML Pipelines) · Building with Claude & Anthropic API (Prompt Caching & Tools) · Microsoft Certified: Power BI Data Analyst Associate (PL-300) · Project Management & Agile Certification.", space_after=0.1)

    return doc.d


def build_2page_resume(fs=8.6):
    doc = ResumeDoc(fs=fs, top_margin=0.35, bottom_margin=0.35, left_margin=0.40, right_margin=0.40, line_spacing=1.04)
    
    # Header
    p = doc.p(None, WD_ALIGN_PARAGRAPH.CENTER)
    _run(p, "ASMI BASNET", fs + 3.5, bold=True, color=NAVY)
    
    p = doc.p(None, WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_after = Pt(2.0)
    _run(p, "Miami, FL (In-State Florida Candidate / Relocating to Miami) · Tallahassee, FL\nasmibas18@gmail.com | (314) 732-5088 | linkedin.com/in/asmibasnet | github.com/asmibasnet | asmibasnet.github.io", fs - 0.3)
    
    # Summary
    doc.heading("Professional Summary", space_before=2.2, space_after=0.15)
    doc.plain(
        "Machine Learning Engineer and Applied AI Specialist (Master of Science in Analytics candidate, Saint Louis University; B.E. in Electronics & Communication) with 3+ years of experience engineering production predictive ML models, generative AI architectures, and intelligent clinical applications. Proven expertise analyzing complex structured and unstructured healthcare datasets (34,000+ EHR records, clinical notes, imaging), training ensemble classifiers (XGBoost, Random Forest, PyTorch), and deploying Explainable AI (SHAP TreeExplainer) to ensure clinical transparency. Skilled in architecting privacy-first, offline retrieval-augmented generation (RAG) suites with automated citation verification (PolicyGuard), time-series forecasting (Prophet + LSTM), and modular SQL across enterprise databases. Florida resident eager to deliver high-impact, compliant, and reliable AI/ML solutions for Jackson Health System."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=2.2, space_after=0.15)
    doc.plain(
        "Supervised & Unsupervised ML (Scikit-Learn, XGBoost, LightGBM, Random Forest, Lasso/Ridge), Deep Learning (PyTorch, Torchvision, CNNs, LSTM), Explainable AI (SHAP TreeExplainer), Time-Series Forecasting (Prophet, ARIMA), K-Fold Cross-Validation, Hyperparameter Optimization (Hyperopt), ROC-AUC / PR-AUC Optimization.",
        justify=False, italic_lead="Machine Learning & Modeling: "
    )
    doc.plain(
        "Retrieval-Augmented Generation (RAG), Local LLM Inference (Ollama: Qwen 2.5, Llama 3, Phi-3), OpenAI GPT-4o API, Anthropic Claude API (Prompt Caching & Tools), Vector Databases (ChromaDB, FAISS), Citation Verification, Clinical Guardrails, Agentic Workflows (Agno, WebMCP), Context Window Optimization.",
        justify=False, italic_lead="Generative AI & LLMs: "
    )
    doc.plain(
        "Python (FastAPI, Flask, Requests, Pandas, NumPy, Scikit-Learn, PyTorch, SQLAlchemy), SQL (Snowflake, PostgreSQL, DuckDB, SQL Server, BigQuery: CTEs, Window Functions), C/C++, TypeScript, Linux/Bash CLI, Git / GitHub, Docker, Docker Compose, RESTful APIs, CI/CD Actions.",
        justify=False, italic_lead="Programming & Infrastructure: "
    )
    doc.plain(
        "EHR Data Modeling (Synthea EHR), DICOM / Medical Imaging, Automated Data Quality Assurance (IV&V), Schema Drift Detection, HIPAA / FERPA Privacy Compliance, Audit Logging, Star-Schema Modeling, Business Intelligence (Microsoft Power BI, DAX, Tableau).",
        justify=False, italic_lead="Healthcare Data & Governance: "
    )
    doc.plain(
        "Agile / Scrum Sprint Planning, JIRA / Project Tracking, Technical Runbooks & API Documentation, Standard Operating Procedures (SOPs), End-User Training & Knowledge-Sharing Sessions, Peer Code Reviews & Design Reviews.",
        justify=False, italic_lead="Project Delivery & Collaboration: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=2.2, space_after=0.15)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.8)
    doc.lr("Graduate Machine Learning & AI Researcher (SPS-Corps)", "August 2025 – Present", italic=True, space_before=0.2)
    doc.bullet("Architected and deployed an air-gapped, privacy-first offline RAG compliance suite in Python (FastAPI, ChromaDB, Ollama) that runs local quantized LLMs to process sensitive policy and clinical guidelines with zero external data exposure and zero data leakage.")
    doc.bullet("Engineered a chain-of-thought verification engine evaluating retrieved context sentence-by-sentence, flagging contradictory requirements (FLAG_CONFLICT), and requesting missing context (CLARIFY) to automate governance audits.")
    doc.bullet("Built a hybrid time-series forecasting engine combining Facebook Prophet with PyTorch LSTM neural networks and automated anomaly detection (flagging variances beyond ±2.5σ) to deliver operational demand and resource predictions.")
    doc.bullet("Optimized modular SQL queries across 8+ operational databases using complex CTEs and analytical window functions (ROW_NUMBER, DENSE_RANK), standardizing data dictionaries and reducing query execution latency by 40%.")
    doc.bullet("Designed cooperating multi-agent automation workflows leveraging the Agno framework and Claude / OpenAI APIs, integrating custom tool-calling routines and REST API connectors that automated cross-departmental research and data synthesis, cutting manual compilation cycles by 90%.")
    doc.bullet("Implemented automated data quality verification suites (IV&V) enforcing referential integrity, schema drift detection, and rigorous data governance standards aligned with institutional and privacy compliance frameworks.")
    doc.bullet("Collaborated directly with cross-functional leadership in Agile sprints to translate operational requirements into technical specifications, authoring runbooks and conducting user training workshops for 25+ staff members.")

    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.8)
    doc.lr("Data & AI Solutions Analyst Intern", "June 2024 – August 2024", italic=True, space_before=0.2)
    doc.bullet("Engineered automated data profiling, cleaning, and validation pipelines in Python (Pandas) and SQL across 10,000+ transactional records, establishing automated webhook alerts that instantly flag schema drift and data discrepancies.")
    doc.bullet("Built and deployed 6 interactive Power BI and Tableau executive dashboards backed by star-schema data models, tracking 15+ operational KPIs to enable leadership to monitor resource allocation, productivity, and project milestones.")
    doc.bullet("Enforced strict institutional data governance and privacy rules, building row-level data quality validation checks across reporting layers.")
    doc.bullet("Partnered closely with program directors and clinical/administrative coordinators to refine metric definitions, delivering ad-hoc analyses on tight deadlines.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd. (Mayahold Group)", "Remote / Kathmandu, Nepal", bold=True, space_before=0.8)
    doc.lr("Systems & Data Analyst", "December 2022 – July 2023", italic=True, space_before=0.2)
    doc.bullet("Extracted, cleansed, and normalized multi-source operational and telemetry datasets (5,000+ records) into standardized 3NF relational schemas using Python (Pandas, SQLAlchemy) and SQL, guaranteeing complete data fidelity.")
    doc.bullet("Executed rigorous statistical hypothesis testing, multivariate linear regressions, and ANOVA modeling in Python, R, and SPSS to evaluate behavioral trends and output audit-ready reporting procedures.")
    doc.bullet("Validated algorithmic logic and automated processing scripts in Python and C/C++, implementing robust exception handling, boundary tests, and audit logging to maintain 99.8% pipeline reliability.")

    # Featured Projects (Clean Page 2 Break)
    doc.d.add_page_break()
    doc.heading("Featured Clinical AI & Machine Learning Projects", space_before=1.5, space_after=0.2)
    
    # CareGap
    doc.lr("CareGap — Clinical Decision Support & Population Health Analytics", "Python, DuckDB, Scikit-Learn, XGBoost, SHAP, FAISS, Streamlit", bold=True, space_before=0.6)
    doc.bullet("Developed a full-stack clinical decision platform querying 34,000+ longitudinal synthetic EHR patient records with sub-second latency using DuckDB columnar storage, training an ensemble classifier (XGBoost, Random Forest, Lasso) for patient risk stratification.")
    doc.bullet("Integrated SHAP TreeExplainer for per-patient transparent risk attribution alongside a dual-tier RAG clinical assistant (FAISS + local Ollama with failover to Gemini) featuring strict guardrails that block unauthorized dosage recommendations.")
    doc.bullet("Delivered interactive clinician-facing dashboards tracking quality measures, care-gap closure rates, and CDC growth-chart Z-score logic for pediatric cohorts.")

    # PolicyGuard
    doc.lr("PolicyGuard — Enterprise Offline RAG & Compliance Intelligence Suite", "Python, Ollama (Qwen 2.5 / Llama 3), ChromaDB, FastAPI, Docker", bold=True, space_before=0.6)
    doc.bullet("Engineered an air-gapped, privacy-first RAG application running entirely on local quantized LLMs via Ollama to analyze confidential enterprise documentation and clinical guidelines with zero third-party API exposure and zero data leakage.")
    doc.bullet("Implemented automated chain-of-thought citation verification and dynamic cosine similarity guardrails, evaluating retrieved context sentence-by-sentence to detect contradictory clauses (FLAG_CONFLICT) and eliminate manual review overhead.")
    doc.bullet("Containerized the multi-service architecture using Docker Compose, creating an intuitive Streamlit interface for business reviewers to inspect sourced conflicts alongside an aggregated governance health score.")
    
    # WashU Datathon
    doc.lr("WashU Clinical Diabetes Risk Prediction Datathon", "Python, Scikit-Learn, XGBoost, LightGBM, Hyperopt", bold=True, space_before=0.6)
    doc.bullet("Engineered an end-to-end ML pipeline on 62,000+ clinical EHR records to predict uncontrolled HbA1c, engineering visit-frequency, medication-count, and glycemic-volatility features.")
    doc.bullet("Benchmarked Logistic Regression, Random Forest, LightGBM, and XGBoost using stratified K-fold cross-validation and hyperparameter tuning, optimizing ROC-AUC, PR-AUC, and F1 to minimize clinical false negatives.")

    # Pulmonary Condition Classifier
    doc.lr("Pulmonary Condition Classification with Deep CNNs", "PyTorch, Torchvision, ResNet, DICOM Processing", bold=True, space_before=0.6)
    doc.bullet("Trained multi-class convolutional neural networks on 33,000+ chest radiographs (COVID-QU-Ex), benchmarking custom CNN architectures against pre-trained ResNet with data augmentation, dropout, batch normalization, and learning-rate scheduling.")
    doc.bullet("Evaluated models with confusion matrices, sensitivity, and specificity across COVID-19, viral pneumonia, and normal classes to support diagnostic workflows.")

    # Education & Academic Background
    doc.heading("Education & Academic Background", space_before=2.2, space_after=0.2)
    doc.lr("Saint Louis University | Master of Science in Analytics (STEM)", "St. Louis, MO | Expected May 2026", bold=True, space_before=0.6)
    doc.plain("Relevant Coursework: Machine Learning, Artificial Intelligence, Deep Learning (CNN, LSTM, Transformers), Cloud Data Architecture, Predictive Big Data Analytics, Database Management & Warehousing.", space_after=0.2)
    doc.lr("Tribhuvan University | Bachelor of Engineering in Electronics & Communication (STEM)", "Kathmandu, Nepal | Conferred 2023", bold=True, space_before=0.4)
    doc.plain("Relevant Coursework: Microprocessors & Microcontrollers (C/C++), Control Systems & PID Theory, Computer Networks, Digital Signal Processing, Probability & Statistics.", space_after=0.2)
    
    doc.heading("Certifications & Accreditations", space_before=2.0, space_after=0.2)
    doc.bullet("AI Engineer for Data Scientists Associate — DataCamp: PyTorch Deep Learning Architectures, Scikit-Learn ML Pipelines, Model Validation.")
    doc.bullet("Building with Claude & Anthropic API — Anthropic: Prompt Caching, Model Context Protocol (MCP) Tools, Production Evaluation Loops.")
    doc.bullet("Microsoft Certified: Power BI Data Analyst Associate (PL-300) — Enterprise Star-Schema Data Modeling, Advanced DAX, KPI Dashboards.")
    doc.bullet("Project Management & Agile Certification — GlobalShala & Saint Louis University: Agile/Scrum Sprint Planning, Milestone Tracking, Stakeholder Communications.")

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
    
    # 1. Generate 1-Page Resume
    docx_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Jackson_Health_System_Machine_Learning_Engineer.docx")
    pdf_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Jackson_Health_System_Machine_Learning_Engineer.pdf")
    
    print("Building 1-page docx...")
    d1 = build_1page_resume(fs=8.2)
    d1.save(docx_1p)
    print("Converting 1-page to PDF...")
    convert_to_pdf(docx_1p, pdf_1p)
    if os.path.exists(pdf_1p):
        p1 = len(fitz.open(pdf_1p))
        print(f"1-Page PDF generated: {pdf_1p} (Page count: {p1})")
    
    # 2. Generate 2-Page Resume
    docx_2p = os.path.join(out_dir, "Asmi_Basnet_Resume_Jackson_Health_System_Machine_Learning_Engineer_2Page.docx")
    pdf_2p = os.path.join(out_dir, "Asmi_Basnet_Resume_Jackson_Health_System_Machine_Learning_Engineer_2Page.pdf")
    
    print("Building 2-page docx...")
    d2 = build_2page_resume(fs=8.6)
    d2.save(docx_2p)
    print("Converting 2-page to PDF...")
    convert_to_pdf(docx_2p, pdf_2p)
    if os.path.exists(pdf_2p):
        p2 = len(fitz.open(pdf_2p))
        print(f"2-Page PDF generated: {pdf_2p} (Page count: {p2})")
