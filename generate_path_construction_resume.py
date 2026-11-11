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
    def __init__(self, fs=8.3, top_margin=0.32, bottom_margin=0.32, left_margin=0.38, right_margin=0.38, line_spacing=1.0):
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
        self.d.core_properties.title = "Asmi Basnet - Path Construction AI Engineer Resume"

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
    doc = ResumeDoc(fs=fs, top_margin=0.30, bottom_margin=0.30, left_margin=0.38, right_margin=0.38, line_spacing=1.02)
    
    # Header
    p = doc.p(None, WD_ALIGN_PARAGRAPH.CENTER)
    _run(p, "ASMI BASNET", fs + 3.2, bold=True, color=NAVY)
    
    p = doc.p(None, WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_after = Pt(1.2)
    _run(p, "Arlington Heights, IL (Open to On-Site Relocation / Greater Chicago Area) · Tallahassee, FL | asmibas18@gmail.com | (314) 732-5088 | linkedin.com/in/asmibasnet | github.com/asmibasnet | asmibasnet.github.io", fs - 0.4)
    
    # Summary
    doc.heading("Professional Summary", space_before=1.5, space_after=0.1)
    doc.plain(
        "AI Engineer and Applied Systems Specialist (M.S. in Analytics candidate, Saint Louis University; B.E. in Electronics & Communication) with 3+ years of experience engineering production LLM applications, autonomous agent workflows, and enterprise automation systems. Proficient in Python, C/C++, TypeScript, and SQL, with proven experience building retrieval-augmented generation (RAG) suites, tool-calling agents (Agno, WebMCP), and predictive forecasting engines. Background architecting air-gapped document intelligence with automated citation verification (PolicyGuard), time-series budget and demand forecasting (Prophet + LSTM), and enterprise data pipelines across Microsoft 365, SharePoint, and REST APIs. Adept at collaborating directly with cross-functional business leaders to translate operational workflows into secure, high-impact software solutions with strict data privacy."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=1.8, space_after=0.1)
    doc.plain(
        "Retrieval-Augmented Generation (RAG), OpenAI GPT-4o API, Anthropic Claude API (Prompt Caching & Tools), Local LLM Inference (Ollama: Qwen 2.5, Llama 3, Phi-3), Agno Multi-Agent Framework, Model Context Protocol (WebMCP), Vector Databases (ChromaDB, FAISS), Citation Verification, Hallucination Guardrails, Context Window Optimization.",
        justify=False, italic_lead="AI Engineering & LLMs: "
    )
    doc.plain(
        "Python (FastAPI, Flask, Requests, Pandas, NumPy, Scikit-Learn, PyTorch, SQLAlchemy), JavaScript / TypeScript, C/C++, Go (Golang), SQL (PostgreSQL, DuckDB, Snowflake, SQL Server: CTEs, Window Functions), RESTful APIs, Webhooks, Linux CLI, Git / GitHub, Docker Compose.",
        justify=False, italic_lead="Programming & Software: "
    )
    doc.plain(
        "Workflow Automation, Enterprise Systems Integration (REST APIs, Webhooks, SharePoint / Microsoft 365 Graph), Automated Multi-Tab Excel Reporting (OpenPyXL, Power Query), Time-Series Forecasting (Prophet, LSTM), Anomaly Detection, Operational Dashboarding.",
        justify=False, italic_lead="Operations & Systems: "
    )
    doc.plain(
        "Automated Data Quality Assurance (IV&V), Schema Drift Monitoring, Audit Logging, Responsible AI & Data Privacy (Air-Gapped / Zero-Leakage Architectures), Star-Schema Dimensional Modeling, Business Intelligence (Microsoft Power BI, DAX, Tableau).",
        justify=False, italic_lead="Data Systems & Governance: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=1.8, space_after=0.1)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Graduate AI Engineer & Research Assistant (SPS-Corps)", "August 2025 – Present", italic=True, space_before=0.1)
    doc.bullet("Architected and deployed an air-gapped, offline RAG policy compliance suite in Python (FastAPI, ChromaDB, Ollama) that runs local quantized LLMs to analyze confidential enterprise documentation with zero third-party API exposure and zero data leakage.")
    doc.bullet("Engineered a chain-of-thought verification engine that evaluates citations sentence-by-sentence, flags contradictory requirements (FLAG_CONFLICT), and requests missing context (CLARIFY), aggregating findings into a governance health score to automate contract and policy audits.")
    doc.bullet("Designed cooperating multi-agent automation workflows leveraging the Agno framework and Anthropic Claude / OpenAI APIs, integrating custom tool-calling routines and REST API connectors that automated cross-departmental research and data synthesis, cutting manual compilation cycles by 90%.")
    doc.bullet("Engineered automated data ingestion, validation, and ETL pipelines across 8+ operational databases using Python and modular SQL CTEs with window functions, standardizing data dictionaries and cutting query execution times by 40%.")
    doc.bullet("Built a hybrid forecasting and anomaly detection engine pairing Facebook Prophet with PyTorch LSTM neural networks, automatically flagging demand surges and budget variances beyond ±2.5σ to deliver predictive resource allocation insights to leadership.")
    doc.bullet("Partnered with departmental leadership in Agile sprints to translate operational needs into technical specifications, authoring runbooks and conducting hands-on user training workshops for 25+ administrative staff to drive seamless tool adoption.")
    
    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Data & AI Solutions Analyst Intern", "June 2024 – August 2024", italic=True, space_before=0.1)
    doc.bullet("Built and deployed 6 interactive executive dashboards across Power BI and Tableau backed by star-schema data models, consolidating 15+ operational KPIs to enable leadership to track resource utilization, productivity, and project milestones.")
    doc.bullet("Engineered automated data profiling, cleaning, and validation scripts in Python (Pandas) and SQL across 10,000+ institutional records, establishing automated webhook alerts that instantly flag schema drift and transactional anomalies.")
    doc.bullet("Partnered closely with cross-functional administrative directors and technical teams to map business requirements into technical deliverables, managing sprint milestones and delivering weekly progress presentations to executive stakeholders.")
    doc.bullet("Implemented automated validation checks and audit routines on sensitive records, enforcing strict governance and institutional privacy compliance.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd. (Mayahold Group)", "Remote / Kathmandu, Nepal", bold=True, space_before=0.5)
    doc.lr("Systems & Data Analyst", "December 2022 – July 2023", italic=True, space_before=0.1)
    doc.bullet("Extracted, cleaned, and normalized multi-source operational and telemetry datasets (5,000+ records) into standardized 3NF relational schemas using Python (Pandas, SQLAlchemy) and SQL, ensuring complete data consistency across enterprise systems.")
    doc.bullet("Validated algorithmic logic and automated backend processing scripts in Python and C/C++, implementing robust exception handling, boundary checks, and audit logging to maintain 99.8% pipeline reliability.")
    doc.bullet("Authored formal technical documentation detailing data dictionaries, data validation rules, and sampling methodologies to establish organizational data governance standards.")

    # Featured Projects
    doc.heading("Featured AI Engineering & Automation Projects", space_before=1.8, space_after=0.1)
    
    # PolicyGuard
    doc.lr("PolicyGuard — Enterprise Offline RAG & Document Intelligence Suite", "Python, Ollama, ChromaDB, FastAPI, Docker", bold=True, space_before=0.4)
    doc.bullet("Engineered an air-gapped, privacy-first RAG application running entirely on local quantized LLMs via Ollama to analyze confidential enterprise policy documents and contracts with zero data leakage.")
    doc.bullet("Implemented automated chain-of-thought citation verification and dynamic cosine similarity guardrails, evaluating retrieved context sentence-by-sentence to detect contradictory clauses (FLAG_CONFLICT) and eliminate manual review overhead.")
    
    # Nepal Disaster Watch
    doc.lr("Nepal Disaster Watch — WebMCP Real-Time Autonomous Command Center", "TypeScript, WebMCP, REST APIs, GeoJSON", bold=True, space_before=0.4)
    doc.bullet("Architected an AI-driven crisis intelligence command center featuring 11 native WebMCP tools (6 read, 5 write) for OpenAI's WebMCP Challenge, synchronizing 6 live institutional feeds across all 77 districts of Nepal.")
    doc.bullet("Engineered resilient feed status adapters (live, snapshot, unreachable), allowing autonomous AI agents to dynamically query spatial risks, pan map layers, and prevent silent system failures during critical operations.")
    
    # Autonomous Multi-Agent System
    doc.lr("Autonomous Multi-Agent Workflow & Research System", "Python, Agno Framework, OpenAI GPT-4o / Claude API, Zapier", bold=True, space_before=0.4)
    doc.bullet("Built a cooperating multi-agent workflow pairing a Market Scout Agent with an Analytical Synthesis Agent to autonomously monitor data feeds, call external APIs, evaluate valuation anomalies, and synthesize structured executive briefs.")
    doc.bullet("Integrated event-driven Zapier webhooks and automated messaging endpoints to ingest real-time operational triggers and distribute structured summaries, reducing manual information gathering and reporting cycles by 90%.")

    # Education & Certifications
    doc.heading("Education & Certifications", space_before=1.8, space_after=0.1)
    doc.lr("Saint Louis University | Master of Science in Analytics (STEM)", "St. Louis, MO | Expected May 2026", bold=True, space_before=0.4)
    doc.lr("Tribhuvan University | Bachelor of Engineering in Electronics & Communication (STEM)", "Kathmandu, Nepal | Conferred 2023", bold=True, space_before=0.1)
    doc.plain("Certifications: Building with Claude & Anthropic API (Prompt Caching & Tools) · AI Engineer for Data Scientists Associate (DataCamp: PyTorch & Scikit-Learn) · Microsoft Certified: Power BI Data Analyst Associate (PL-300) · Project Management & Agile Certification (GlobalShala & SLU).", space_after=0.1)

    return doc.d


def build_2page_resume(fs=8.6):
    doc = ResumeDoc(fs=fs, top_margin=0.35, bottom_margin=0.35, left_margin=0.40, right_margin=0.40, line_spacing=1.04)
    
    # Header
    p = doc.p(None, WD_ALIGN_PARAGRAPH.CENTER)
    _run(p, "ASMI BASNET", fs + 3.5, bold=True, color=NAVY)
    
    p = doc.p(None, WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_after = Pt(2.0)
    _run(p, "Arlington Heights, IL (Open to On-Site Relocation / Greater Chicago Area) · Tallahassee, FL\nasmibas18@gmail.com | (314) 732-5088 | linkedin.com/in/asmibasnet | github.com/asmibasnet | asmibasnet.github.io", fs - 0.3)
    
    # Summary
    doc.heading("Professional Summary", space_before=2.2, space_after=0.15)
    doc.plain(
        "AI Engineer and Applied Systems Specialist (M.S. in Analytics candidate, Saint Louis University; B.E. in Electronics & Communication) with 3+ years of experience engineering production LLM applications, autonomous agent workflows, and enterprise automation systems. Proficient in Python, C/C++, TypeScript, and SQL, with proven experience building retrieval-augmented generation (RAG) suites, tool-calling agents (Agno, WebMCP), and predictive forecasting engines. Background architecting air-gapped document intelligence with automated citation verification (PolicyGuard), time-series budget and demand forecasting (Prophet + LSTM), and enterprise data pipelines across Microsoft 365, SharePoint, and REST APIs. Adept at collaborating directly with cross-functional business leaders to translate operational workflows into secure, high-impact software solutions with strict data privacy."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=2.2, space_after=0.15)
    doc.plain(
        "Retrieval-Augmented Generation (RAG), OpenAI GPT-4o API, Anthropic Claude API (Prompt Caching & Tools), Local LLM Inference (Ollama: Qwen 2.5, Llama 3, Phi-3), Agno Multi-Agent Framework, Model Context Protocol (WebMCP), Vector Databases (ChromaDB, FAISS), Citation Verification, Hallucination Guardrails, Context Window Optimization.",
        justify=False, italic_lead="AI Engineering & LLMs: "
    )
    doc.plain(
        "Python (FastAPI, Flask, Requests, Pandas, NumPy, Scikit-Learn, PyTorch, SQLAlchemy), JavaScript / TypeScript, C/C++, Go (Golang), SQL (PostgreSQL, DuckDB, Snowflake, SQL Server: CTEs, Window Functions), RESTful APIs, Webhooks, Linux CLI, Git / GitHub, Docker, Docker Compose, CI/CD Actions.",
        justify=False, italic_lead="Programming & Custom Software: "
    )
    doc.plain(
        "Workflow Automation, Enterprise Systems Integration (REST APIs, Webhooks, SharePoint / Microsoft 365 Graph), Automated Multi-Tab Excel Reporting (OpenPyXL, Power Query), Time-Series Forecasting (Prophet, LSTM), Anomaly Detection, Operational Dashboarding.",
        justify=False, italic_lead="Operations & Systems: "
    )
    doc.plain(
        "Automated Data Quality Assurance (IV&V), Schema Drift Monitoring, Audit Logging, Responsible AI & Data Privacy (Air-Gapped / Zero-Leakage Architectures), Star-Schema Dimensional Modeling, Business Intelligence (Microsoft Power BI, DAX, Tableau).",
        justify=False, italic_lead="Data Systems, QA & Governance: "
    )
    doc.plain(
        "Agile / Scrum Sprint Planning, JIRA / Project Tracking, Technical Runbooks & API Documentation, Standard Operating Procedures (SOPs), End-User Training & Knowledge-Sharing Sessions, Tool Adoption Telemetry & Stakeholder Engagement.",
        justify=False, italic_lead="Project Delivery & Collaboration: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=2.2, space_after=0.15)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.8)
    doc.lr("Graduate AI Engineer & Research Assistant (SPS-Corps)", "August 2025 – Present", italic=True, space_before=0.2)
    doc.bullet("Architected and deployed an air-gapped, offline RAG policy compliance suite in Python (FastAPI, ChromaDB, Ollama) that runs local quantized LLMs to analyze confidential enterprise documentation with zero third-party API exposure and zero data leakage.")
    doc.bullet("Engineered a chain-of-thought verification engine that evaluates citations sentence-by-sentence, flags contradictory requirements (FLAG_CONFLICT), and requests missing context (CLARIFY), aggregating findings into a governance health score to automate contract and policy audits.")
    doc.bullet("Designed cooperating multi-agent automation workflows leveraging the Agno framework and Anthropic Claude / OpenAI APIs, integrating custom tool-calling routines and REST API connectors that automated cross-departmental research and data synthesis, cutting manual compilation cycles by 90%.")
    doc.bullet("Engineered automated data ingestion, validation, and ETL pipelines across 8+ operational databases using Python and modular SQL CTEs with window functions, standardizing data dictionaries and cutting query execution times by 40%.")
    doc.bullet("Built a hybrid forecasting and anomaly detection engine pairing Facebook Prophet with PyTorch LSTM neural networks, automatically flagging demand surges and budget variances beyond ±2.5σ to deliver predictive resource allocation insights to leadership.")
    doc.bullet("Authored comprehensive technical documentation, architecture blueprints, and API specifications for deployed AI tools, while leading interactive user training workshops and knowledge-sharing sessions for 25+ administrative staff to facilitate frictionless tool adoption.")
    doc.bullet("Established a structured enhancement request tracking system to collect end-user feedback, prioritize technical debt and feature requests, and iteratively push containerized updates via Docker Compose and GitHub Actions.")

    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.8)
    doc.lr("Data & AI Solutions Analyst Intern", "June 2024 – August 2024", italic=True, space_before=0.2)
    doc.bullet("Built and deployed 6 interactive executive dashboards across Power BI and Tableau backed by star-schema data models, consolidating 15+ operational KPIs to enable leadership to track resource utilization, productivity, and project milestones.")
    doc.bullet("Engineered automated data profiling, cleaning, and validation scripts in Python (Pandas) and SQL across 10,000+ institutional records, establishing automated webhook alerts that instantly flag schema drift and transactional anomalies.")
    doc.bullet("Partnered closely with cross-functional administrative directors and technical teams to map business requirements into technical deliverables, managing sprint milestones and delivering weekly progress presentations to executive stakeholders.")
    doc.bullet("Implemented automated validation checks and audit routines on sensitive records, enforcing strict governance and institutional privacy compliance.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd. (Mayahold Group)", "Remote / Kathmandu, Nepal", bold=True, space_before=0.8)
    doc.lr("Systems & Data Analyst", "December 2022 – July 2023", italic=True, space_before=0.2)
    doc.bullet("Extracted, cleaned, and normalized multi-source operational and telemetry datasets (5,000+ records) into standardized 3NF relational schemas using Python (Pandas, SQLAlchemy) and SQL, ensuring complete data consistency across enterprise systems.")
    doc.bullet("Validated algorithmic logic and automated backend processing scripts in Python and C/C++, implementing robust exception handling, boundary checks, and audit logging to maintain 99.8% pipeline reliability.")
    doc.bullet("Authored formal technical documentation detailing data dictionaries, data validation rules, and sampling methodologies to establish organizational data governance standards.")

    # Featured Projects (Clean Page 2 Break)
    doc.d.add_page_break()
    doc.heading("Featured AI Engineering & Automation Projects", space_before=1.5, space_after=0.2)
    
    # PolicyGuard
    doc.lr("PolicyGuard — Enterprise Offline RAG & Document Intelligence Suite", "Python, Ollama, ChromaDB, FastAPI, Docker", bold=True, space_before=0.6)
    doc.bullet("Engineered an air-gapped, privacy-first RAG application running entirely on local quantized LLMs via Ollama to analyze confidential enterprise policy documents and contracts with zero data leakage.")
    doc.bullet("Implemented automated chain-of-thought citation verification and dynamic cosine similarity guardrails, evaluating retrieved context sentence-by-sentence to detect contradictory clauses (FLAG_CONFLICT) and eliminate manual review overhead.")
    doc.bullet("Containerized the multi-service architecture using Docker Compose, creating an intuitive Streamlit interface for business reviewers to inspect sourced conflicts alongside an aggregated governance health score.")
    
    # Nepal Disaster Watch
    doc.lr("Nepal Disaster Watch — WebMCP Real-Time Autonomous Command Center", "TypeScript, WebMCP, REST APIs, Leaflet.js, GeoJSON", bold=True, space_before=0.6)
    doc.bullet("Architected an AI-driven crisis intelligence command center featuring 11 native WebMCP tools (6 read, 5 write) for OpenAI's WebMCP Challenge, synchronizing 6 live institutional feeds (NDRRMA, river sensors, Copernicus GloFAS) across all 77 districts of Nepal.")
    doc.bullet("Engineered resilient feed status adapters (live, snapshot, unreachable), allowing autonomous AI agents to dynamically query spatial risks, pan map layers, and prevent silent system failures during critical operations.")
    doc.bullet("Implemented dynamic sensor data parsing, event-driven status alerting, and diagnostic verification logic, establishing standard protocols for telemetry verification and remote operations monitoring.")

    # Autonomous Multi-Agent System
    doc.lr("Autonomous Multi-Agent Workflow & Research System", "Python, Agno Framework, OpenAI GPT-4o / Claude API, Zapier, REST APIs", bold=True, space_before=0.6)
    doc.bullet("Built a cooperating multi-agent workflow pairing a Market Scout Agent with an Analytical Synthesis Agent to autonomously monitor data feeds, call external APIs, evaluate valuation anomalies, and synthesize structured executive briefs.")
    doc.bullet("Integrated event-driven Zapier webhooks and automated messaging endpoints to ingest real-time operational triggers and distribute structured summaries, reducing manual information gathering and reporting cycles by 90%.")

    # CareGap
    doc.lr("CareGap — Clinical Decision Support & Explainable AI Risk Platform", "Python, DuckDB, Scikit-Learn, XGBoost, SHAP, FAISS, Streamlit", bold=True, space_before=0.6)
    doc.bullet("Developed a full-stack decision engine querying 34,000+ synthetic EHR records with sub-second latency using DuckDB columnar Parquet storage, training an ensemble classifier (XGBoost, Random Forest, Lasso) for risk stratification.")
    doc.bullet("Integrated SHAP TreeExplainer for transparent per-record feature attribution alongside a dual-tier RAG assistant featuring strict prompt guardrails to prevent unsafe recommendations.")

    # Education & Academic Background
    doc.heading("Education & Academic Background", space_before=2.2, space_after=0.2)
    doc.lr("Saint Louis University | Master of Science in Analytics (STEM)", "St. Louis, MO | Expected May 2026", bold=True, space_before=0.6)
    doc.plain("Relevant Coursework: Machine Learning, Artificial Intelligence, LLM Prompt Engineering & Agent Architectures, Cloud Data Architecture, Predictive Big Data Analytics, Database Management & Warehousing.", space_after=0.2)
    doc.lr("Tribhuvan University | Bachelor of Engineering in Electronics & Communication (STEM)", "Kathmandu, Nepal | Conferred 2023", bold=True, space_before=0.4)
    doc.plain("Relevant Coursework: Microprocessors & Microcontrollers (C/C++), Control Systems & PID Theory, Computer Networks, Digital Signal Processing, Probability & Statistics.", space_after=0.2)
    
    doc.heading("Certifications & Accreditations", space_before=2.0, space_after=0.2)
    doc.bullet("Building with Claude & Anthropic API — Anthropic: Prompt Caching, Model Context Protocol (MCP) Tools, Production Evaluation Loops.")
    doc.bullet("AI Engineer for Data Scientists Associate — DataCamp: PyTorch Deep Learning Architectures, Scikit-Learn ML Pipelines, Model Validation.")
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
    docx_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Path_Construction_AI_Engineer.docx")
    pdf_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Path_Construction_AI_Engineer.pdf")
    
    print("Building 1-page docx...")
    d1 = build_1page_resume(fs=8.2)
    d1.save(docx_1p)
    print("Converting 1-page to PDF...")
    convert_to_pdf(docx_1p, pdf_1p)
    if os.path.exists(pdf_1p):
        p1 = len(fitz.open(pdf_1p))
        print(f"1-Page PDF generated: {pdf_1p} (Page count: {p1})")
    
    # 2. Generate 2-Page Resume
    docx_2p = os.path.join(out_dir, "Asmi_Basnet_Resume_Path_Construction_AI_Engineer_2Page.docx")
    pdf_2p = os.path.join(out_dir, "Asmi_Basnet_Resume_Path_Construction_AI_Engineer_2Page.pdf")
    
    print("Building 2-page docx...")
    d2 = build_2page_resume(fs=8.6)
    d2.save(docx_2p)
    print("Converting 2-page to PDF...")
    convert_to_pdf(docx_2p, pdf_2p)
    if os.path.exists(pdf_2p):
        p2 = len(fitz.open(pdf_2p))
        print(f"2-Page PDF generated: {pdf_2p} (Page count: {p2})")
