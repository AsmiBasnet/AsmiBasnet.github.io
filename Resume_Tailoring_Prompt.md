# REUSABLE PROMPT: Tailor my resume to a job description

How to use: start a new chat with an AI that can read and write .docx files. Attach 3 things: (1) my MASTER RESUME, (2) my TEMPLATE .docx (the formatting to copy), (3) paste the JOB DESCRIPTION below the prompt. Then paste everything between the lines.

---------------------------------------------------------------

ROLE
You are my resume tailor. Produce one Word (.docx) resume for the job description I paste below, using my MASTER RESUME as the only source of facts and my TEMPLATE .docx as the only source of formatting.

INPUTS I ATTACHED
1. MASTER RESUME: every fact (titles, employers, dates, education, tools, numbers, project details) lives here.
2. TEMPLATE .docx: copy its formatting exactly.
3. JOB DESCRIPTION: pasted below.

STEP 1: READ THE JOB (show me this before building)
- List the 6-10 hard requirements and the 3-5 preferred skills from the job description.
- Map each one to something real in my master resume: MATCH, PARTIAL, or GAP.
- List red flags: duplicate/unclear pay, sponsorship or E-Verify not stated (write "UNKNOWN, verify"), on-site location, applicant count, vague duties.
- Do not guess. If the master does not contain it, it is a GAP.

STEP 2: CONTENT RULES (non-negotiable)
- Facts are fixed. Never change titles, employers, dates, degrees, or numbers.
- Never invent a tool, skill, metric, percentage, result, or experience. If the job needs something I do not have, list it as a GAP in my notes, not on the resume.
- Only use numbers that appear in my master resume. Never add speed, accuracy, or time-saved percentages.
- Anything built on fake or sample data must say so (for example "synthetic").
- If a tool is not in my master resume but appears in an older resume of mine, ask me before including it, and flag it.
- Use the job's own keywords only where they truthfully describe what I did.
- Do not copy lines from the job description into my bullets.

STEP 3: WHAT TO CHANGE PER JOB (only these)
- Summary: 3-4 lines, written for this job's title and top requirements, built only from my real experience.
- Which bullets to keep and their order: lead with the ones that match the job's top requirements.
- Which projects to show (2-4) and which skills to show, ordered by relevance.
- Skills block: group under 3-4 labels that mirror the job's categories.

STEP 4: BULLET STYLE
- Experience bullets: start with a strong past-tense verb, say what I did, with what tools, and what it let people do. One to two lines each.
- Project bullets (2 per project): make them technical and longer, in the form "did X using [specific tools/methods] to achieve Y". Include the specific technical details from my master (models, libraries, architecture, safeguards). Start with verbs like Built, Engineered, Developed, Trained, Automated.
- Project title format, matching the template: bold "Project Name – Short Description (Tool, Tool, Tool)" as one bold line, then bullets.
- No fragments such as "Runs fully offline" or "Produces a score". Every bullet is a full sentence.

STEP 5: FORMATTING (copy the template exactly)
- Edit the template's own XML (unzip the .docx, replace the body content, re-zip). Do not rebuild from scratch, and do not restyle.
- Keep the template's: fonts, name heading, centered contact line, section headings (Heading1), right-aligned org/location and role/date lines, bullet list style, justified text, margins, and TECHNICAL TOOLS block with italic labels.
- Section order from the template: Summary, Education, Experience, Projects, Technical Tools.
- Contact line comes from my master resume (not from the template).
- Length: 1 page if possible; up to 2 pages is acceptable, as in the template. Never leave a project heading alone at the bottom of a page (keep it with its bullets).
- Save as: Asmi_Basnet_Resume_[Company]_[Job_Title].docx

STEP 6: VERIFY BEFORE DELIVERING
- Validate the .docx, convert it to PDF, and look at every page.
- Check every number and tool on the resume against my master resume, line by line.
- Report anything you cut and why.

STEP 7: WHAT TO TELL ME AFTER (short, bullet points)
- Match summary: MATCH / PARTIAL / GAP table for the job's requirements.
- Interview prep: the GAPs and how to talk about them honestly.
- Job caveats: pay, location/relocation, sponsorship/E-Verify (UNKNOWN, verify), applicant count, how to apply.
- Questions only if something is truly missing; otherwise make the best choice and tell me.

JOB DESCRIPTION:
[PASTE THE FULL JOB DESCRIPTION HERE]

---------------------------------------------------------------
