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
        self.d.core_properties.title = "Asmi Basnet - Danfoss Turbocor AI Software Engineer Resume"

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
        "AI Software Engineer and Systems Developer based locally in Tallahassee, FL (M.S. in Analytics candidate, Saint Louis University; B.E. in Electronics & Communication) with 3+ years of experience engineering production software, autonomous agentic workflows, and AI-enabled developer tooling. Strong programming background spanning Python, TypeScript, Go, C/C++, and SQL, with deep experience building retrieval-augmented generation (RAG) suites, Model Context Protocol (MCP) tool execution pipelines, and automated software validation frameworks. Uniquely bridges software engineering with electrical and physical systems engineering (control theory, sensor telemetry, microcontrollers), enabling effective cross-functional collaboration with mechanical, electrical, and systems engineers at Danfoss Turbocor to accelerate product modernization and ship reliable AI-driven solutions."
    )
    
    # Skills
    doc.heading("Technical Skills & Core Competencies", space_before=1.8, space_after=0.1)
    doc.plain(
        "Autonomous AI Agents (Agno, ReAct / Supervisor Patterns), Model Context Protocol (WebMCP Tools), Retrieval-Augmented Generation (RAG), Anthropic Claude API (Prompt Caching & Tools), OpenAI GPT-4o API, Local LLM Inference (Ollama: Qwen 2.5, Llama 3), Vector Databases (ChromaDB, FAISS), Citation Verification, Output Evaluation.",
        justify=False, italic_lead="Agentic Workflows & GenAI: "
    )
    doc.plain(
        "Python (FastAPI, Flask, Requests, Pandas, NumPy, Scikit-Learn, PyTorch, SQLAlchemy), TypeScript, JavaScript (ES6+), Go (Golang), C/C++, SQL (PostgreSQL, DuckDB, Snowflake, SQL Server: CTEs, Window Functions), RESTful API Engineering, Webhooks, Linux/Bash CLI, Git / GitHub, Docker, Docker Compose, CI/CD Actions.",
        justify=False, italic_lead="Software & Programming: "
    )
    doc.plain(
        "Internal Developer Tooling, Automated Testing & Verification, Firmware & Algorithmic Debugging, Telemetry Stream Processing, Star-Schema Modeling, Automated Data Quality Assurance (IV&V), Schema Drift Monitoring, Business Intelligence (Microsoft Power BI, DAX, Tableau).",
        justify=False, italic_lead="Systems & Automation: "
    )
    doc.plain(
        "Agile / Scrum Sprint Planning, JIRA, Technical Runbooks & Architecture Blueprints, Standard Operating Procedures (SOPs), Peer Code Reviews, Cross-Functional Engineering Collaboration (Software, Electrical, Mechanical).",
        justify=False, italic_lead="Engineering Delivery: "
    )
    
    # Experience
    doc.heading("Professional Experience", space_before=1.8, space_after=0.1)
    
    # SLU Graduate Role
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Graduate AI Software Engineer & Research Assistant (SPS-Corps)", "August 2025 – Present", italic=True, space_before=0.1)
    doc.bullet("Architected and deployed an air-gapped, offline RAG software application in Python (FastAPI, ChromaDB, Ollama) that runs local quantized LLMs to analyze complex technical specifications with zero external API exposure and zero data leakage.")
    doc.bullet("Engineered a chain-of-thought verification engine evaluating retrieved context sentence-by-sentence, flagging contradictory requirements (FLAG_CONFLICT), and requesting missing context (CLARIFY) to eliminate hallucinations and automate compliance reviews.")
    doc.bullet("Designed cooperating multi-agent automation workflows leveraging the Agno framework and Anthropic Claude / OpenAI APIs, integrating custom tool-calling routines and REST API connectors that automated cross-departmental research and data synthesis, cutting manual compilation cycles by 90%.")
    doc.bullet("Engineered automated data ingestion, validation, and ETL pipelines across 8+ operational databases using Python and modular SQL CTEs with window functions, standardizing data dictionaries and cutting query execution times by 40%.")
    doc.bullet("Built a hybrid forecasting and anomaly detection engine pairing Facebook Prophet with PyTorch LSTM neural networks, automatically flagging demand surges and telemetry variances beyond ±2.5 standard deviations to deliver predictive operational insights.")
    doc.bullet("Partnered with engineering and project leadership in Agile sprints to translate operational requirements into technical specifications, authoring runbooks and conducting hands-on user training workshops for 25+ staff members to drive seamless tool adoption.")
    
    # SLU Intern
    doc.lr("Saint Louis University", "St. Louis, MO", bold=True, space_before=0.5)
    doc.lr("Data & AI Solutions Analyst Intern", "June 2024 – August 2024", italic=True, space_before=0.1)
    doc.bullet("Built and deployed 6 interactive executive dashboards across Power BI and Tableau backed by star-schema data models, consolidating 15+ operational KPIs to enable leadership to track resource utilization, productivity, and project milestones.")
    doc.bullet("Engineered automated data profiling, cleaning, and validation scripts in Python (Pandas) and SQL across 10,000+ institutional records, establishing automated webhook alerts that instantly flag schema drift and transactional anomalies.")
    doc.bullet("Enforced strict institutional data governance and privacy rules, building row-level data quality validation checks across reporting layers.")

    # Entegra Sources
    doc.lr("Entegra Sources Pvt. Ltd. (Mayahold Group)", "Remote / Kathmandu, Nepal", bold=True, space_before=0.5)
    doc.lr("Systems & Software Analyst", "December 2022 – July 2023", italic=True, space_before=0.1)
    doc.bullet("Extracted, cleaned, and normalized multi-source operational and telemetry datasets (5,000+ records) into standardized 3NF relational schemas using Python (Pandas, SQLAlchemy) and SQL, ensuring complete data consistency across enterprise systems.")
    doc.bullet("Validated algorithmic logic and automated backend processing scripts in Python and C/C++, implementing robust exception handling, boundary checks, and audit logging to maintain 99.8% pipeline reliability.")
    doc.bullet("Authored formal technical documentation detailing data dictionaries, data validation rules, and sampling methodologies to establish organizational data governance standards.")

    # Featured Projects
    doc.heading("Featured AI Engineering & Automation Projects", space_before=1.8, space_after=0.1)
    
    # Nepal Disaster Watch
    doc.lr("Nepal Disaster Watch — WebMCP Real-Time Autonomous Command Center", "TypeScript, WebMCP, REST APIs, GeoJSON", bold=True, space_before=0.4)
    doc.bullet("Architected an AI-driven crisis intelligence command center featuring 11 native WebMCP tools (6 read, 5 write) for OpenAI's WebMCP Challenge, synchronizing 6 live institutional feeds across all 77 districts of Nepal.")
    doc.bullet("Engineered resilient feed status adapters (live, snapshot, unreachable), allowing autonomous AI agents to dynamically query spatial risks, pan map layers, and prevent silent system failures during critical operations.")
    
    # PolicyGuard
    doc.lr("PolicyGuard — Enterprise Offline RAG & Technical Document Suite", "Python, Ollama, ChromaDB, FastAPI, Docker", bold=True, space_before=0.4)
    doc.bullet("Engineered an air-gapped, privacy-first RAG application running entirely on local quantized LLMs via Ollama to analyze confidential enterprise engineering documents and specifications with zero data leakage.")
    doc.bullet("Implemented automated chain-of-thought citation verification and dynamic cosine similarity guardrails, evaluating retrieved context sentence-by-sentence to detect contradictory clauses (FLAG_CONFLICT) and eliminate manual review overhead.")
    
    # Autonomous Multi-Agent System
    doc.lr("Autonomous Multi-Agent Workflow & Tool Execution System", "Python, Agno Framework, Claude API / GPT-4o, REST APIs", bold=True, space_before=0.4)
    doc.bullet("Built a cooperating multi-agent workflow pairing a Market Scout Agent with an Analytical Synthesis Agent to autonomously monitor data feeds, call external APIs, evaluate valuation anomalies, and synthesize structured executive briefs.")
    doc.bullet("Integrated event-driven Zapier webhooks and automated messaging endpoints to ingest real-time operational triggers and distribute structured summaries, reducing manual information gathering and reporting cycles by 90%.")

    # Education & Certifications
    doc.heading("Education & Certifications", space_before=1.8, space_after=0.1)
    doc.lr("Saint Louis University | Master of Science in Analytics (STEM)", "St. Louis, MO | Expected May 2026", bold=True, space_before=0.4)
    doc.lr("Tribhuvan University | Bachelor of Engineering in Electronics & Communication (STEM)", "Kathmandu, Nepal | Conferred 2023", bold=True, space_before=0.1)
    doc.plain("Certifications: Building with Claude & Anthropic API (Prompt Caching & Tools) · AI Engineer for Data Scientists Associate (DataCamp: PyTorch & Scikit-Learn) · Microsoft Certified: Power BI Data Analyst Associate (PL-300) · Project Management & Agile Certification (GlobalShala & SLU).", space_after=0.1)

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
    
    docx_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Danfoss_AI_Software_Engineer.docx")
    pdf_1p = os.path.join(out_dir, "Asmi_Basnet_Resume_Danfoss_AI_Software_Engineer.pdf")
    
    print("Building 1-page docx...")
    d1 = build_1page_resume(fs=8.2)
    d1.save(docx_1p)
    print("Converting 1-page to PDF...")
    convert_to_pdf(docx_1p, pdf_1p)
    if os.path.exists(pdf_1p):
        p1 = len(fitz.open(pdf_1p))
        print(f"1-Page PDF generated: {pdf_1p} (Page count: {p1})")
