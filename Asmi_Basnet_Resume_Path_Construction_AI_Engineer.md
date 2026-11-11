# ASMI BASNET
**Arlington Heights, IL (Open to On-Site Relocation / Greater Chicago Area) · Tallahassee, FL**  
**asmibas18@gmail.com | (314) 732-5088 | [linkedin.com/in/asmibasnet](https://linkedin.com/in/asmibasnet) | [github.com/asmibasnet](https://github.com/asmibasnet) | [asmibasnet.github.io](https://asmibasnet.github.io)**

---

## PROFESSIONAL SUMMARY
AI Engineer and Applied Systems Specialist (M.S. in Analytics candidate, Saint Louis University; B.E. in Electronics & Communication) with 3+ years of experience engineering production LLM applications, autonomous agent workflows, and enterprise automation systems. Proficient in Python, C/C++, TypeScript, and SQL, with proven experience building retrieval-augmented generation (RAG) suites, tool-calling agents (Agno, WebMCP), and predictive forecasting engines. Background architecting air-gapped document intelligence with automated citation verification (PolicyGuard), time-series budget and demand forecasting (Prophet + LSTM), and enterprise data pipelines across Microsoft 365, SharePoint, and REST APIs. Adept at collaborating directly with cross-functional business leaders to translate operational workflows into secure, high-impact software solutions with strict data privacy.

---

## TECHNICAL SKILLS & CORE COMPETENCIES
* **AI Engineering & LLM Systems:** Retrieval-Augmented Generation (RAG), OpenAI GPT-4o API, Anthropic Claude API (Prompt Caching & Tools), Local LLM Inference (Ollama: Qwen 2.5, Llama 3, Phi-3), Agno Multi-Agent Framework, Model Context Protocol (WebMCP), Vector Databases (ChromaDB, FAISS), Citation Verification, Hallucination Guardrails, Context Window Optimization.
* **Programming & Custom Software:** Python (FastAPI, Flask, Requests, Pandas, NumPy, Scikit-Learn, PyTorch, SQLAlchemy), JavaScript / TypeScript, C/C++, Go (Golang), SQL (PostgreSQL, DuckDB, Snowflake, SQL Server: CTEs, Window Functions), RESTful APIs, Webhooks, Linux CLI, Git / GitHub, Docker, Docker Compose, CI/CD Actions.
* **Operations & Systems:** Workflow Automation, Enterprise Systems Integration (REST APIs, Webhooks, SharePoint / Microsoft 365 Graph), Automated Multi-Tab Excel Reporting (OpenPyXL, Power Query), Time-Series Forecasting (Prophet, LSTM), Anomaly Detection, Operational Dashboarding.
* **Data Systems, QA & Governance:** Automated Data Quality Assurance (IV&V), Schema Drift Monitoring, Audit Logging, Responsible AI & Data Privacy (Air-Gapped / Zero-Leakage Architectures), Star-Schema Dimensional Modeling, Business Intelligence (Microsoft Power BI, DAX, Tableau).
* **Project Delivery & Collaboration:** Agile / Scrum Sprint Planning, JIRA / Project Tracking, Technical Runbooks & API Documentation, Standard Operating Procedures (SOPs), End-User Training & Knowledge-Sharing Sessions, Tool Adoption Telemetry & Stakeholder Engagement.

---

## PROFESSIONAL EXPERIENCE

### **Saint Louis University** | St. Louis, MO
**Graduate AI Engineer & Research Assistant (SPS-Corps)** | *August 2025 – Present*
* Architected and deployed an air-gapped, offline RAG policy compliance suite in Python (FastAPI, ChromaDB, Ollama) that runs local quantized LLMs to analyze confidential enterprise documentation with zero third-party API exposure and zero data leakage.
* Engineered a chain-of-thought verification engine that evaluates citations sentence-by-sentence, flags contradictory requirements (`FLAG_CONFLICT`), and requests missing context (`CLARIFY`), aggregating findings into a governance health score to automate contract and policy audits.
* Designed cooperating multi-agent automation workflows leveraging the Agno framework and Anthropic Claude / OpenAI APIs, integrating custom tool-calling routines and REST API connectors that automated cross-departmental research and data synthesis, cutting manual compilation cycles by 90%.
* Engineered automated data ingestion, validation, and ETL pipelines across 8+ operational databases using Python and modular SQL CTEs with window functions, standardizing data dictionaries and cutting query execution times by 40%.
* Built a hybrid forecasting and anomaly detection engine pairing Facebook Prophet with PyTorch LSTM neural networks, automatically flagging demand surges and budget variances beyond ±2.5σ to deliver predictive resource allocation insights to leadership.
* Authored comprehensive technical documentation, architecture blueprints, and API specifications for deployed AI tools, while leading interactive user training workshops and knowledge-sharing sessions for 25+ administrative staff to facilitate frictionless tool adoption.
* Established a structured enhancement request tracking system to collect end-user feedback, prioritize technical debt and feature requests, and iteratively push containerized updates via Docker Compose and GitHub Actions.

### **Saint Louis University** | St. Louis, MO
**Data & AI Solutions Analyst Intern** | *June 2024 – August 2024*
* Built and deployed 6 interactive executive dashboards across Power BI and Tableau backed by star-schema data models, consolidating 15+ operational KPIs to enable leadership to track resource utilization, productivity, and project milestones.
* Engineered automated data profiling, cleaning, and validation scripts in Python (Pandas) and SQL across 10,000+ institutional records, establishing automated webhook alerts that instantly flag schema drift and transactional anomalies.
* Partnered closely with cross-functional administrative directors and technical teams to map business requirements into technical deliverables, managing sprint milestones and delivering weekly progress presentations to executive stakeholders.
* Implemented automated validation checks and audit routines on sensitive records, enforcing strict governance and institutional privacy compliance.

### **Entegra Sources Pvt. Ltd. (Mayahold Group)** | Remote / Kathmandu, Nepal
**Systems & Data Analyst** | *December 2022 – July 2023*
* Extracted, cleaned, and normalized multi-source operational and telemetry datasets (5,000+ records) into standardized 3NF relational schemas using Python (Pandas, SQLAlchemy) and SQL, ensuring complete data consistency across enterprise systems.
* Validated algorithmic logic and automated backend processing scripts in Python and C/C++, implementing robust exception handling, boundary checks, and audit logging to maintain 99.8% pipeline reliability.
* Authored formal technical documentation detailing data dictionaries, data validation rules, and sampling methodologies to establish organizational data governance standards.

---

## FEATURED AI ENGINEERING & AUTOMATION PROJECTS

### **PolicyGuard — Enterprise Offline RAG & Document Intelligence Suite**
*(Python, Ollama, ChromaDB, FastAPI, Streamlit, Docker Compose)*
* Engineered an air-gapped, privacy-first RAG application running entirely on local quantized LLMs via Ollama to analyze confidential enterprise policy documents and contracts with zero data leakage.
* Implemented automated chain-of-thought citation verification and dynamic cosine similarity guardrails, evaluating retrieved context sentence-by-sentence to detect contradictory clauses (`FLAG_CONFLICT`) and eliminate manual review overhead.
* Containerized the multi-service architecture using Docker Compose, creating an intuitive Streamlit interface for business reviewers to inspect sourced conflicts alongside an aggregated governance health score.

### **Nepal Disaster Watch — WebMCP Real-Time Autonomous Command Center**
*(TypeScript, Model Context Protocol, REST APIs, Leaflet.js, GeoJSON)*
* Architected an AI-driven crisis intelligence command center featuring 11 native WebMCP tools (6 read, 5 write) for OpenAI's WebMCP Challenge, synchronizing 6 live institutional feeds (NDRRMA, river sensors, Copernicus GloFAS) across all 77 districts of Nepal.
* Engineered resilient feed status adapters (`live`, `snapshot`, `unreachable`), allowing autonomous AI agents to dynamically query spatial risks, pan map layers, and prevent silent system failures during critical operations.
* Implemented dynamic sensor data parsing, event-driven status alerting, and diagnostic verification logic, establishing standard protocols for telemetry verification and remote operations monitoring.

### **Autonomous Multi-Agent Workflow & Research System**
*(Python, Agno Framework, OpenAI GPT-4o / Claude API, Zapier, REST APIs)*
* Built a cooperating multi-agent workflow pairing a Market Scout Agent with an Analytical Synthesis Agent to autonomously monitor data feeds, call external APIs, evaluate valuation anomalies, and synthesize structured executive briefs.
* Integrated event-driven Zapier webhooks and automated messaging endpoints to ingest real-time operational triggers and distribute structured summaries, reducing manual information gathering and reporting cycles by 90%.

### **CareGap — Clinical Decision Support & Explainable AI Risk Platform**
*(Python, DuckDB, Scikit-Learn, XGBoost, SHAP, FAISS, Streamlit)*
* Developed a full-stack decision engine querying 34,000+ synthetic EHR records with sub-second latency using DuckDB columnar Parquet storage, training an ensemble classifier (XGBoost, Random Forest, Lasso) for risk stratification.
* Integrated SHAP TreeExplainer for transparent per-record feature attribution alongside a dual-tier RAG assistant featuring strict prompt guardrails to prevent unsafe recommendations.

---

## EDUCATION & ACADEMIC BACKGROUND

### **Saint Louis University** | St. Louis, MO
**Master of Science in Analytics (STEM)** | *Expected May 2026*
* **Relevant Coursework:** Machine Learning, Artificial Intelligence, LLM Prompt Engineering & Agent Architectures, Cloud Data Architecture, Predictive Big Data Analytics, Database Management & Warehousing.

### **Tribhuvan University** | Kathmandu, Nepal
**Bachelor of Engineering in Electronics & Communication (STEM)** | *Conferred 2023 (2017 – 2022)*
* **Relevant Coursework:** Microprocessors & Microcontrollers (C/C++), Control Systems & PID Theory, Computer Networks, Digital Signal Processing, Numerical Methods, Probability & Statistics.

---

## CERTIFICATIONS & ACCREDITATIONS
* **Building with Claude & Anthropic API** — Anthropic: Prompt Caching, Model Context Protocol (MCP) Tools, Production Evaluation Loops.
* **AI Engineer for Data Scientists Associate** — DataCamp: PyTorch Deep Learning Architectures, Scikit-Learn ML Pipelines, Model Validation.
* **Microsoft Certified: Power BI Data Analyst Associate (PL-300)** — Enterprise Star-Schema Data Modeling, Advanced DAX, KPI Dashboards.
* **Project Management & Agile Certification** — GlobalShala & Saint Louis University: Agile/Scrum Sprint Planning, Milestone Tracking, Stakeholder Communications.
