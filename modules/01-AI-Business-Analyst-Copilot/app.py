import os
import re
from io import BytesIO
from xml.sax.saxutils import escape

import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from docx import Document

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

st.set_page_config(
    page_title="AI Business Analyst Copilot",
    page_icon="🤖",
    layout="wide",
)

load_dotenv(override=True)
api_key = os.getenv("OPENAI_API_KEY")

if not api_key:
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        api_key = None


def clean_inline_markdown(text):
    return text.replace("**", "").replace("__", "").replace("`", "")


def safe_filename(value, fallback="Business_Requirements"):
    cleaned = value.strip().replace(" ", "_")
    cleaned = "".join(
        c for c in cleaned if c.isalnum() or c in ("_", "-")
    )
    return cleaned or fallback


def generate_ai_text(prompt, spinner_message):
    client = OpenAI(api_key=api_key)
    with st.spinner(spinner_message):
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=prompt,
        )
    return response.output_text



# =========================================================
# DASHBOARD METRIC HELPERS
# =========================================================

def count_unique_ids(text, prefix):
    if not text:
        return 0

    pattern = rf"\b{re.escape(prefix)}-\d{{3}}\b"

    return len(
        set(
            re.findall(
                pattern,
                text,
            )
        )
    )


def get_requirement_ids(text):
    if not text:
        return []

    requirement_ids = re.findall(
        r"\b(?:FR|NFR)-\d{3}\b",
        text,
    )

    return sorted(
        set(requirement_ids),
        key=lambda item: (
            item.startswith("NFR"),
            item,
        ),
    )


def parse_rtm_coverage(rtm_text):
    coverage_counts = {
        "Covered": 0,
        "Partial": 0,
        "Gap": 0,
    }

    if not rtm_text:
        return coverage_counts

    for line in rtm_text.splitlines():
        stripped = line.strip()

        if not stripped.startswith("|"):
            continue

        if "---" in stripped:
            continue

        cells = [
            cell.strip()
            for cell in stripped.strip("|").split("|")
        ]

        if len(cells) < 6:
            continue

        if cells[0] == "Requirement ID":
            continue

        status = cells[5]

        if status in coverage_counts:
            coverage_counts[status] += 1

    return coverage_counts


def calculate_traceability_score(coverage_counts):
    total = sum(
        coverage_counts.values()
    )

    if total == 0:
        return 0

    weighted_coverage = (
        coverage_counts["Covered"]
        + (
            0.5
            * coverage_counts["Partial"]
        )
    )

    return round(
        (
            weighted_coverage
            / total
        )
        * 100
    )


def create_word_document(document_title, content_text, company, project_name):
    buffer = BytesIO()
    document = Document()
    document.add_heading(document_title, level=0)
    document.add_paragraph(f"Company: {company}")
    document.add_paragraph(f"Project: {project_name}")
    document.add_paragraph("")

    for line in content_text.splitlines():
        cleaned_line = line.strip()

        if not cleaned_line:
            document.add_paragraph("")
            continue

        if cleaned_line.startswith("### "):
            document.add_heading(
                clean_inline_markdown(cleaned_line[4:]), level=3
            )
        elif cleaned_line.startswith("## "):
            document.add_heading(
                clean_inline_markdown(cleaned_line[3:]), level=2
            )
        elif cleaned_line.startswith("# "):
            document.add_heading(
                clean_inline_markdown(cleaned_line[2:]), level=1
            )
        elif cleaned_line.startswith("- "):
            document.add_paragraph(
                clean_inline_markdown(cleaned_line[2:]),
                style="List Bullet",
            )
        else:
            document.add_paragraph(clean_inline_markdown(cleaned_line))

    document.save(buffer)
    buffer.seek(0)
    return buffer


def create_pdf_document(document_title, content_text, company, project_name):
    buffer = BytesIO()
    pdf = SimpleDocTemplate(
        buffer,
        pagesize=LETTER,
        rightMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
    )

    styles = getSampleStyleSheet()
    story = [
        Paragraph(escape(document_title), styles["Title"]),
        Spacer(1, 12),
        Paragraph(f"<b>Company:</b> {escape(company)}", styles["Normal"]),
        Paragraph(f"<b>Project:</b> {escape(project_name)}", styles["Normal"]),
        Spacer(1, 18),
    ]

    for line in content_text.splitlines():
        cleaned_line = line.strip()

        if not cleaned_line:
            story.append(Spacer(1, 8))
            continue

        if cleaned_line.startswith("### "):
            story.append(
                Paragraph(
                    escape(clean_inline_markdown(cleaned_line[4:])),
                    styles["Heading3"],
                )
            )
        elif cleaned_line.startswith("## "):
            story.append(
                Paragraph(
                    escape(clean_inline_markdown(cleaned_line[3:])),
                    styles["Heading2"],
                )
            )
        elif cleaned_line.startswith("# "):
            story.append(
                Paragraph(
                    escape(clean_inline_markdown(cleaned_line[2:])),
                    styles["Heading1"],
                )
            )
        elif cleaned_line.startswith("- "):
            story.append(
                Paragraph(
                    "&bull; " + escape(clean_inline_markdown(cleaned_line[2:])),
                    styles["Normal"],
                )
            )
        else:
            story.append(
                Paragraph(
                    escape(clean_inline_markdown(cleaned_line)),
                    styles["Normal"],
                )
            )

        story.append(Spacer(1, 6))

    pdf.build(story)
    buffer.seek(0)
    return buffer


def show_download_buttons(
    document_title,
    content_text,
    company,
    project_name,
    filename_suffix,
):
    base_name = safe_filename(
        project_name,
        fallback="Business_Analysis_Project",
    )

    word_file = create_word_document(
        document_title,
        content_text,
        company,
        project_name,
    )

    pdf_file = create_pdf_document(
        document_title,
        content_text,
        company,
        project_name,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            label="📘 Download Word",
            data=word_file,
            file_name=f"{base_name}_{filename_suffix}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with col2:
        st.download_button(
            label="📕 Download PDF",
            data=pdf_file,
            file_name=f"{base_name}_{filename_suffix}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with col3:
        st.download_button(
            label="📝 Download Markdown",
            data=content_text,
            file_name=f"{base_name}_{filename_suffix}.md",
            mime="text/markdown",
            use_container_width=True,
        )


default_state = {
    "generated_requirements": None,
    "generated_company": "",
    "generated_project_name": "",
    "generated_user_stories": None,
    "generated_test_cases": None,
    "generated_rtm": None,
    "generated_executive_analysis": None,
}

for key, value in default_state.items():
    if key not in st.session_state:
        st.session_state[key] = value


st.title("🤖 AI Business Analyst Copilot")
st.caption(
    "Generate professional Business Analysis documentation using AI."
)
st.divider()

st.subheader("📋 Business Requirements Generator")

company = st.text_input("Company")
project_name = st.text_input("Project Name")
problem = st.text_area("Business Problem", height=120)
goal = st.text_area("Business Goal", height=120)
stakeholders = st.text_area("Stakeholders (one per line)", height=100)
constraints = st.text_area("Business Constraints", height=100)


if st.button(
    "✨ Generate Business Requirements",
    use_container_width=True,
):
    required_fields_missing = (
        not company.strip()
        or not project_name.strip()
        or not problem.strip()
        or not goal.strip()
    )

    if required_fields_missing:
        st.warning(
            "Please complete Company, Project Name, "
            "Business Problem, and Business Goal."
        )

    elif not api_key:
        st.error(
            "The OpenAI API key was not found. "
            "Check your .env file locally or Streamlit Secrets in the live app."
        )

    else:
        brd_prompt = f'''
You are a senior Business Systems Analyst.

Create a professional Business Requirements Document using the information below.

Company:
{company}

Project Name:
{project_name}

Business Problem:
{problem}

Business Goal:
{goal}

Stakeholders:
{stakeholders or "Not provided"}

Business Constraints:
{constraints or "Not provided"}

Create the following sections:

1. Executive Summary
2. Business Problem Statement
3. Business Objectives
4. Project Scope
5. Stakeholder Analysis
6. Functional Requirements
7. Non-Functional Requirements
8. User Stories
9. Acceptance Criteria
10. Assumptions
11. Dependencies
12. Risks and Mitigation Strategies
13. Success Metrics

Requirements:

- Use clear, professional enterprise business-analysis language.
- Make the requirements specific to the information provided.
- Number functional requirements using FR-001, FR-002, FR-003, etc.
- Number non-functional requirements using NFR-001, NFR-002, NFR-003, etc.
- Write user stories using:
  "As a [user], I want [capability], so that [business benefit]."
- Make acceptance criteria measurable wherever possible.
- Clearly distinguish assumptions from confirmed requirements.
- Do not invent highly specific company facts that were not provided.
'''

        try:
            requirements = generate_ai_text(
                brd_prompt,
                "Generating business requirements with AI...",
            )

            if not requirements:
                st.error(
                    "OpenAI returned an empty response. Please try again."
                )
            else:
                st.session_state["generated_requirements"] = requirements
                st.session_state["generated_company"] = company
                st.session_state["generated_project_name"] = project_name
                st.session_state["generated_user_stories"] = None
                st.session_state["generated_test_cases"] = None
                st.session_state["generated_rtm"] = None
                st.session_state["generated_executive_analysis"] = None

        except Exception as error:
            st.error(f"OpenAI connection error: {error}")


if st.session_state["generated_requirements"]:
    requirements = st.session_state["generated_requirements"]
    generated_company = st.session_state["generated_company"]
    generated_project_name = st.session_state["generated_project_name"]

    st.success("Business Requirements Generated")
    st.divider()
    st.header("📄 Generated Business Requirements Document")
    st.markdown(requirements)

    st.subheader("📥 Export Business Requirements")

    show_download_buttons(
        "Business Requirements Document",
        requirements,
        generated_company,
        generated_project_name,
        "BRD",
    )

    st.divider()
    st.header("📋 Jira-Ready User Stories")
    st.caption(
        "Generate implementation-ready user stories from the approved "
        "Business Requirements Document."
    )

    if st.button(
        "🧩 Generate Jira User Stories",
        use_container_width=True,
    ):
        if not api_key:
            st.error("The OpenAI API key was not found.")
        else:
            stories_prompt = f'''
You are a senior Business Analyst and Product Owner.

Using ONLY the Business Requirements Document below, create a set of
Jira-ready user stories.

Company:
{generated_company}

Project:
{generated_project_name}

BUSINESS REQUIREMENTS DOCUMENT:

{requirements}

Create enough user stories to cover the major functional requirements
without creating unnecessary duplicates.

Use this exact Markdown structure for every story:

## US-001 — [Concise Jira Summary]

- Epic: [Logical epic name]
- Priority: [High, Medium, or Low]
- Story Points: [1, 2, 3, 5, or 8]
- Linked Requirements: [FR-### and/or NFR-###]

### User Story

As a [specific user/persona], I want [capability], so that [business benefit].

### Acceptance Criteria

1. Given [context], when [action], then [measurable outcome].
2. Given [context], when [action], then [measurable outcome].
3. Given [context], when [action], then [measurable outcome].

### Business Value

[One concise sentence explaining why the story matters.]

Rules:

- Number stories sequentially: US-001, US-002, etc.
- Use requirements traceability wherever possible.
- Do not invent requirements that are not supported by the BRD.
- Write acceptance criteria in Given/When/Then format.
- Keep Jira summaries concise and action-oriented.
- Assign realistic story points based on relative complexity.
'''

            try:
                user_stories = generate_ai_text(
                    stories_prompt,
                    "Generating Jira-ready user stories...",
                )

                if not user_stories:
                    st.error("OpenAI returned an empty response.")
                else:
                    st.session_state["generated_user_stories"] = user_stories
                    st.session_state["generated_test_cases"] = None
                    st.session_state["generated_rtm"] = None
                    st.session_state["generated_executive_analysis"] = None

            except Exception as error:
                st.error(f"User story generation error: {error}")


    if st.session_state["generated_user_stories"]:
        user_stories = st.session_state["generated_user_stories"]

        st.success("Jira User Stories Generated")
        st.markdown(user_stories)

        st.subheader("📥 Export Jira User Stories")

        show_download_buttons(
            "Jira-Ready User Stories",
            user_stories,
            generated_company,
            generated_project_name,
            "Jira_User_Stories",
        )

        st.divider()
        st.header("🧪 AI-Generated QA Test Cases")
        st.caption(
            "Generate traceable functional and negative test cases "
            "from the BRD and Jira stories."
        )

        if st.button(
            "🔬 Generate QA Test Cases",
            use_container_width=True,
        ):
            if not api_key:
                st.error("The OpenAI API key was not found.")
            else:
                test_prompt = f'''
You are a senior QA Analyst working with a Business Analyst and Product Owner.

Create professional QA test cases using ONLY the Business Requirements
Document and Jira user stories provided below.

Company:
{generated_company}

Project:
{generated_project_name}

BUSINESS REQUIREMENTS DOCUMENT:

{requirements}

JIRA USER STORIES:

{user_stories}

Create test coverage for the important business flows, validation rules,
acceptance criteria, and major negative scenarios.

Use this exact Markdown structure for every test case:

## TC-001 — [Concise Test Case Title]

- Linked User Story: [US-###]
- Linked Requirement: [FR-### and/or NFR-###]
- Priority: [High, Medium, or Low]
- Test Type: [Functional, Negative, Integration, Security, or Performance]

### Objective

[What this test validates.]

### Preconditions

- [Required starting condition]

### Test Steps

1. [Action]
2. [Action]
3. [Action]

### Expected Result

[Specific measurable expected outcome.]

Rules:

- Number test cases sequentially: TC-001, TC-002, etc.
- Include positive and negative coverage where appropriate.
- Preserve traceability to user stories and requirements.
- Do not invent functionality not supported by the source documents.
- Make steps executable by a QA tester.
- Keep expected results objective and testable.
'''

                try:
                    test_cases = generate_ai_text(
                        test_prompt,
                        "Generating QA test cases...",
                    )

                    if not test_cases:
                        st.error("OpenAI returned an empty response.")
                    else:
                        st.session_state["generated_test_cases"] = test_cases
                        st.session_state["generated_rtm"] = None
                        st.session_state["generated_executive_analysis"] = None

                except Exception as error:
                    st.error(f"Test case generation error: {error}")


        if st.session_state["generated_test_cases"]:
            test_cases = st.session_state["generated_test_cases"]

            st.success("QA Test Cases Generated")
            st.markdown(test_cases)

            st.subheader("📥 Export QA Test Cases")

            show_download_buttons(
                "QA Test Cases",
                test_cases,
                generated_company,
                generated_project_name,
                "QA_Test_Cases",
            )

            st.divider()
            st.header("🔗 Requirements Traceability Matrix")
            st.caption(
                "Map each functional and non-functional requirement "
                "to its Jira user stories and QA test cases."
            )

            if st.button(
                "🧭 Generate Traceability Matrix",
                use_container_width=True,
            ):
                if not api_key:
                    st.error("The OpenAI API key was not found.")
                else:
                    rtm_prompt = f'''
You are a senior Business Systems Analyst responsible for requirements
governance, quality assurance, and traceability.

Using ONLY the source artifacts below, create a professional Requirements
Traceability Matrix (RTM).

Company:
{generated_company}

Project:
{generated_project_name}

BUSINESS REQUIREMENTS DOCUMENT:

{requirements}

JIRA USER STORIES:

{user_stories}

QA TEST CASES:

{test_cases}

Your job is to map every functional requirement (FR-###) and every
non-functional requirement (NFR-###) to the user stories and test cases
that validate it.

Return the RTM as a Markdown table using EXACTLY these columns:

| Requirement ID | Requirement Summary | Priority | User Story ID(s) | Test Case ID(s) | Coverage Status |
|---|---|---|---|---|---|

Rules:

- Include one row for every FR-### and NFR-### found in the BRD.
- Do not invent requirement IDs, user story IDs, or test case IDs.
- Requirement Summary must be concise.
- Priority must be High, Medium, or Low based on the business impact
  supported by the source artifacts.
- If multiple user stories or test cases apply, separate the IDs with commas.
- If no user story is mapped, write: Not mapped
- If no test case is mapped, write: Not mapped
- Coverage Status must be exactly one of: Covered, Partial, Gap
- Covered means the requirement has at least one mapped user story AND at
  least one mapped test case.
- Partial means it has either a mapped user story OR a mapped test case,
  but not both.
- Gap means neither is mapped.

After the table, add:

## Traceability Summary

- Total Requirements: [number]
- Covered: [number]
- Partial: [number]
- Gaps: [number]

## Recommended Actions

Provide a short prioritized list of actions for any Partial or Gap items.

Do not create new requirements merely to improve coverage. Report actual
traceability gaps honestly.
'''

                    try:
                        rtm = generate_ai_text(
                            rtm_prompt,
                            "Generating requirements traceability matrix...",
                        )

                        if not rtm:
                            st.error("OpenAI returned an empty response.")
                        else:
                            st.session_state["generated_rtm"] = rtm
                            st.session_state["generated_executive_analysis"] = None

                    except Exception as error:
                        st.error(
                            f"Traceability matrix generation error: {error}"
                        )

            if st.session_state["generated_rtm"]:
                rtm = st.session_state["generated_rtm"]

                st.success("Requirements Traceability Matrix Generated")
                st.markdown(rtm)

                st.subheader("📥 Export Traceability Matrix")

                show_download_buttons(
                    "Requirements Traceability Matrix",
                    rtm,
                    generated_company,
                    generated_project_name,
                    "Requirements_Traceability_Matrix",
                )

                # =============================================
                # EXECUTIVE PROJECT DASHBOARD
                # =============================================

                st.divider()

                st.header(
                    "📊 Executive Project Dashboard"
                )

                st.caption(
                    "A deterministic view of delivery artifacts, "
                    "requirements coverage, and traceability health."
                )

                requirement_ids = get_requirement_ids(
                    requirements
                )

                requirement_count = len(
                    requirement_ids
                )

                user_story_count = count_unique_ids(
                    user_stories,
                    "US",
                )

                test_case_count = count_unique_ids(
                    test_cases,
                    "TC",
                )

                coverage_counts = parse_rtm_coverage(
                    rtm
                )

                traceability_score = (
                    calculate_traceability_score(
                        coverage_counts
                    )
                )

                metric1, metric2, metric3, metric4 = (
                    st.columns(4)
                )

                with metric1:
                    st.metric(
                        "Requirements",
                        requirement_count,
                    )

                with metric2:
                    st.metric(
                        "User Stories",
                        user_story_count,
                    )

                with metric3:
                    st.metric(
                        "QA Test Cases",
                        test_case_count,
                    )

                with metric4:
                    st.metric(
                        "Traceability Score",
                        f"{traceability_score}%",
                    )

                coverage1, coverage2, coverage3 = (
                    st.columns(3)
                )

                with coverage1:
                    st.metric(
                        "Covered",
                        coverage_counts["Covered"],
                    )

                with coverage2:
                    st.metric(
                        "Partial",
                        coverage_counts["Partial"],
                    )

                with coverage3:
                    st.metric(
                        "Gaps",
                        coverage_counts["Gap"],
                    )

                st.progress(
                    traceability_score / 100
                )

                st.caption(
                    "Traceability Score formula: "
                    "Covered requirements receive full credit, "
                    "Partial requirements receive half credit, "
                    "and Gap requirements receive no credit."
                )

                if coverage_counts["Gap"] > 0:
                    st.warning(
                        "Traceability gaps remain. Review the RTM "
                        "before treating the project as implementation-ready."
                    )
                elif coverage_counts["Partial"] > 0:
                    st.info(
                        "No complete traceability gaps were detected, "
                        "but some requirements have partial coverage."
                    )
                else:
                    st.success(
                        "All requirements represented in the RTM "
                        "have full story and test-case coverage."
                    )


                # =============================================
                # MOSCOW PRIORITIZATION + RISK ANALYSIS
                # =============================================

                st.divider()

                st.header(
                    "🎯 MoSCoW Prioritization & Risk Analysis"
                )

                st.caption(
                    "Create an executive-ready prioritization, "
                    "risk register, readiness assessment, "
                    "and recommended next actions."
                )

                if st.button(
                    "📈 Generate Executive Analysis",
                    use_container_width=True,
                ):

                    if not api_key:

                        st.error(
                            "The OpenAI API key was not found."
                        )

                    else:

                        executive_prompt = f"""
You are a senior Business Systems Analyst,
Product Owner, and delivery governance advisor.

Create an executive project analysis using
ONLY the source artifacts below.

Company:
{generated_company}

Project:
{generated_project_name}

CURRENT TRACEABILITY METRICS:

- Requirements: {requirement_count}
- User Stories: {user_story_count}
- QA Test Cases: {test_case_count}
- Covered Requirements: {coverage_counts["Covered"]}
- Partial Requirements: {coverage_counts["Partial"]}
- Gap Requirements: {coverage_counts["Gap"]}
- Traceability Score: {traceability_score}%

BUSINESS REQUIREMENTS DOCUMENT:

{requirements}

JIRA USER STORIES:

{user_stories}

QA TEST CASES:

{test_cases}

REQUIREMENTS TRACEABILITY MATRIX:

{rtm}

Create the following sections.

## Executive Summary

Provide a concise executive-level summary of
the project, business objective, delivery
position, and most important concerns.

## MoSCoW Prioritization

Return a Markdown table using EXACTLY these columns:

| Requirement ID | MoSCoW Category | Business Rationale |
|---|---|---|

Rules:

- Include every FR-### and NFR-### from the BRD.
- Category must be exactly one of:
  Must Have
  Should Have
  Could Have
  Won't Have This Release
- Do not invent requirement IDs.
- Base the category on business criticality,
  dependencies, risk, and stated project goals.
- "Won't Have This Release" means deferred,
  not rejected.

## Risk Register

Return a Markdown table using EXACTLY these columns:

| Risk ID | Risk Description | Probability | Impact | Score | Rating | Mitigation | Owner Role |
|---|---|---|---|---|---|---|---|

Rules:

- Number risks RISK-001, RISK-002, etc.
- Probability must be an integer from 1 to 5.
- Impact must be an integer from 1 to 5.
- Score must equal Probability multiplied by Impact.
- Rating must use:
  Low = 1-5
  Medium = 6-10
  High = 11-15
  Critical = 16-25
- Identify only risks supported by the source
  artifacts or reasonable delivery risks directly
  implied by them.
- Owner Role must be a role, not a person's name.
- Mitigation must be actionable.

## Delivery Readiness Assessment

Use this exact structure:

- Readiness: [Green, Amber, or Red]
- Traceability Score: {traceability_score}%
- Key Strengths: [concise statement]
- Key Constraints: [concise statement]
- Decision Basis: [concise explanation]

Guidance:

- Green means the artifacts show strong coverage
  with no material unresolved delivery blockers.
- Amber means implementation may proceed only
  after identified gaps, risks, or dependencies
  are addressed.
- Red means major unresolved gaps or critical
  risks make implementation premature.

## Recommended Next Actions

Provide 3 to 7 prioritized actions.
Start each action with:
1.
2.
3.
and continue sequentially.

Important:

- Do not invent company facts.
- Do not hide traceability gaps.
- Do not claim that AI-generated analysis is
  formal business approval.
- Clearly distinguish delivery recommendations
  from confirmed project decisions.
"""

                        try:

                            executive_analysis = (
                                generate_ai_text(
                                    executive_prompt,
                                    (
                                        "Generating MoSCoW priorities, "
                                        "risk analysis, and executive "
                                        "readiness assessment..."
                                    ),
                                )
                            )

                            if not executive_analysis:

                                st.error(
                                    "OpenAI returned "
                                    "an empty response."
                                )

                            else:

                                st.session_state[
                                    "generated_executive_analysis"
                                ] = executive_analysis

                        except Exception as error:

                            st.error(
                                "Executive analysis "
                                f"generation error: {error}"
                            )


                # =============================================
                # DISPLAY EXECUTIVE ANALYSIS
                # =============================================

                if (
                    st.session_state[
                        "generated_executive_analysis"
                    ]
                ):

                    executive_analysis = (
                        st.session_state[
                            "generated_executive_analysis"
                        ]
                    )

                    st.success(
                        "Executive Analysis Generated"
                    )

                    st.markdown(
                        executive_analysis
                    )

                    st.subheader(
                        "📥 Export Executive Analysis"
                    )

                    show_download_buttons(
                        (
                            "Executive Project Analysis - "
                            "MoSCoW Prioritization and Risk Register"
                        ),
                        executive_analysis,
                        generated_company,
                        generated_project_name,
                        "Executive_Project_Analysis",
                    )