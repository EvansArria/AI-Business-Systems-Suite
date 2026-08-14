import os
from io import BytesIO
from xml.sax.saxutils import escape

import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from docx import Document

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
)


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="AI Business Analyst Copilot",
    page_icon="🤖",
    layout="wide",
)


# =========================================================
# API KEY / ENVIRONMENT SETUP
# =========================================================

load_dotenv(override=True)

api_key = os.getenv("OPENAI_API_KEY")

if not api_key:
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        api_key = None


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_inline_markdown(text):
    return (
        text
        .replace("**", "")
        .replace("__", "")
        .replace("`", "")
    )


def safe_filename(
    value,
    fallback="Business_Requirements",
):
    cleaned = (
        value
        .strip()
        .replace(" ", "_")
    )

    cleaned = "".join(
        char
        for char in cleaned
        if char.isalnum()
        or char in ("_", "-")
    )

    return cleaned or fallback


def generate_ai_text(
    prompt,
    spinner_message,
):
    client = OpenAI(
        api_key=api_key
    )

    with st.spinner(
        spinner_message
    ):
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=prompt,
        )

    return response.output_text


# =========================================================
# WORD EXPORT
# =========================================================

def create_word_document(
    document_title,
    content_text,
    company,
    project_name,
):
    buffer = BytesIO()

    document = Document()

    document.add_heading(
        document_title,
        level=0,
    )

    document.add_paragraph(
        f"Company: {company}"
    )

    document.add_paragraph(
        f"Project: {project_name}"
    )

    document.add_paragraph("")

    for line in content_text.splitlines():

        cleaned_line = line.strip()

        if not cleaned_line:
            document.add_paragraph("")
            continue

        if cleaned_line.startswith("### "):

            document.add_heading(
                clean_inline_markdown(
                    cleaned_line[4:]
                ),
                level=3,
            )

        elif cleaned_line.startswith("## "):

            document.add_heading(
                clean_inline_markdown(
                    cleaned_line[3:]
                ),
                level=2,
            )

        elif cleaned_line.startswith("# "):

            document.add_heading(
                clean_inline_markdown(
                    cleaned_line[2:]
                ),
                level=1,
            )

        elif cleaned_line.startswith("- "):

            document.add_paragraph(
                clean_inline_markdown(
                    cleaned_line[2:]
                ),
                style="List Bullet",
            )

        else:

            document.add_paragraph(
                clean_inline_markdown(
                    cleaned_line
                )
            )

    document.save(
        buffer
    )

    buffer.seek(0)

    return buffer


# =========================================================
# PDF EXPORT
# =========================================================

def create_pdf_document(
    document_title,
    content_text,
    company,
    project_name,
):
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

    story = []

    story.append(
        Paragraph(
            escape(
                document_title
            ),
            styles["Title"],
        )
    )

    story.append(
        Spacer(
            1,
            12,
        )
    )

    story.append(
        Paragraph(
            f"<b>Company:</b> "
            f"{escape(company)}",
            styles["Normal"],
        )
    )

    story.append(
        Paragraph(
            f"<b>Project:</b> "
            f"{escape(project_name)}",
            styles["Normal"],
        )
    )

    story.append(
        Spacer(
            1,
            18,
        )
    )

    for line in content_text.splitlines():

        cleaned_line = line.strip()

        if not cleaned_line:

            story.append(
                Spacer(
                    1,
                    8,
                )
            )

            continue

        if cleaned_line.startswith(
            "### "
        ):

            story.append(
                Paragraph(
                    escape(
                        clean_inline_markdown(
                            cleaned_line[4:]
                        )
                    ),
                    styles["Heading3"],
                )
            )

        elif cleaned_line.startswith(
            "## "
        ):

            story.append(
                Paragraph(
                    escape(
                        clean_inline_markdown(
                            cleaned_line[3:]
                        )
                    ),
                    styles["Heading2"],
                )
            )

        elif cleaned_line.startswith(
            "# "
        ):

            story.append(
                Paragraph(
                    escape(
                        clean_inline_markdown(
                            cleaned_line[2:]
                        )
                    ),
                    styles["Heading1"],
                )
            )

        elif cleaned_line.startswith(
            "- "
        ):

            story.append(
                Paragraph(
                    "&bull; "
                    + escape(
                        clean_inline_markdown(
                            cleaned_line[2:]
                        )
                    ),
                    styles["Normal"],
                )
            )

        else:

            story.append(
                Paragraph(
                    escape(
                        clean_inline_markdown(
                            cleaned_line
                        )
                    ),
                    styles["Normal"],
                )
            )

        story.append(
            Spacer(
                1,
                6,
            )
        )

    pdf.build(
        story
    )

    buffer.seek(0)

    return buffer


# =========================================================
# REUSABLE DOWNLOAD BUTTONS
# =========================================================

def show_download_buttons(
    document_title,
    content_text,
    company,
    project_name,
    filename_suffix,
):
    base_name = safe_filename(
        project_name,
        fallback=(
            "Business_Analysis_Project"
        ),
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

    col1, col2, col3 = (
        st.columns(3)
    )

    with col1:

        st.download_button(
            label="📘 Download Word",
            data=word_file,
            file_name=(
                f"{base_name}_"
                f"{filename_suffix}.docx"
            ),
            mime=(
                "application/vnd."
                "openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with col2:

        st.download_button(
            label="📕 Download PDF",
            data=pdf_file,
            file_name=(
                f"{base_name}_"
                f"{filename_suffix}.pdf"
            ),
            mime="application/pdf",
            use_container_width=True,
        )

    with col3:

        st.download_button(
            label="📝 Download Markdown",
            data=content_text,
            file_name=(
                f"{base_name}_"
                f"{filename_suffix}.md"
            ),
            mime="text/markdown",
            use_container_width=True,
        )


# =========================================================
# SESSION STATE
# =========================================================

default_state = {
    "generated_requirements": None,
    "generated_company": "",
    "generated_project_name": "",
    "generated_user_stories": None,
    "generated_test_cases": None,
}

for key, value in (
    default_state.items()
):
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# APP HEADER
# =========================================================

st.title(
    "🤖 AI Business Analyst Copilot"
)

st.caption(
    "Generate professional "
    "Business Analysis documentation "
    "using AI."
)

st.divider()


# =========================================================
# BUSINESS REQUIREMENTS INPUT
# =========================================================

st.subheader(
    "📋 Business Requirements Generator"
)

company = st.text_input(
    "Company"
)

project_name = st.text_input(
    "Project Name"
)

problem = st.text_area(
    "Business Problem",
    height=120,
)

goal = st.text_area(
    "Business Goal",
    height=120,
)

stakeholders = st.text_area(
    "Stakeholders (one per line)",
    height=100,
)

constraints = st.text_area(
    "Business Constraints",
    height=100,
)


# =========================================================
# GENERATE BRD
# =========================================================

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
            "Please complete Company, "
            "Project Name, Business Problem, "
            "and Business Goal."
        )

    elif not api_key:

        st.error(
            "The OpenAI API key was not found. "
            "Check your .env file locally "
            "or Streamlit Secrets in the "
            "live app."
        )

    else:

        brd_prompt = f"""
You are a senior Business Systems Analyst.

Create a professional Business Requirements
Document using the information below.

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

- Use clear, professional enterprise
  business-analysis language.

- Make the requirements specific to the
  information provided.

- Number functional requirements using
  FR-001, FR-002, FR-003, etc.

- Number non-functional requirements using
  NFR-001, NFR-002, NFR-003, etc.

- Write user stories using:
  "As a [user], I want [capability],
  so that [business benefit]."

- Make acceptance criteria measurable
  wherever possible.

- Clearly distinguish assumptions from
  confirmed requirements.

- Do not invent highly specific company
  facts that were not provided.
"""

        try:

            requirements = (
                generate_ai_text(
                    brd_prompt,
                    (
                        "Generating business "
                        "requirements with AI..."
                    ),
                )
            )

            if not requirements:

                st.error(
                    "OpenAI returned an empty "
                    "response. Please try again."
                )

            else:

                st.session_state[
                    "generated_requirements"
                ] = requirements

                st.session_state[
                    "generated_company"
                ] = company

                st.session_state[
                    "generated_project_name"
                ] = project_name

                # Reset downstream artifacts
                # when a new BRD is created.
                st.session_state[
                    "generated_user_stories"
                ] = None

                st.session_state[
                    "generated_test_cases"
                ] = None

        except Exception as error:

            st.error(
                f"OpenAI connection error: "
                f"{error}"
            )


# =========================================================
# DISPLAY BRD
# =========================================================

if (
    st.session_state[
        "generated_requirements"
    ]
):

    requirements = (
        st.session_state[
            "generated_requirements"
        ]
    )

    generated_company = (
        st.session_state[
            "generated_company"
        ]
    )

    generated_project_name = (
        st.session_state[
            "generated_project_name"
        ]
    )

    st.success(
        "Business Requirements Generated"
    )

    st.divider()

    st.header(
        "📄 Generated Business "
        "Requirements Document"
    )

    st.markdown(
        requirements
    )

    st.subheader(
        "📥 Export Business Requirements"
    )

    show_download_buttons(
        "Business Requirements Document",
        requirements,
        generated_company,
        generated_project_name,
        "BRD",
    )


    # =====================================================
    # JIRA-READY USER STORIES
    # =====================================================

    st.divider()

    st.header(
        "📋 Jira-Ready User Stories"
    )

    st.caption(
        "Generate implementation-ready "
        "user stories from the approved "
        "Business Requirements Document."
    )

    if st.button(
        "🧩 Generate Jira User Stories",
        use_container_width=True,
    ):

        if not api_key:

            st.error(
                "The OpenAI API key "
                "was not found."
            )

        else:

            stories_prompt = f"""
You are a senior Business Analyst
and Product Owner.

Using ONLY the Business Requirements
Document below, create a set of
Jira-ready user stories.

Company:
{generated_company}

Project:
{generated_project_name}

BUSINESS REQUIREMENTS DOCUMENT:

{requirements}

Create enough user stories to cover
the major functional requirements
without creating unnecessary duplicates.

Use this exact Markdown structure
for every story:

## US-001 — [Concise Jira Summary]

- Epic: [Logical epic name]
- Priority: [High, Medium, or Low]
- Story Points: [1, 2, 3, 5, or 8]
- Linked Requirements:
  [FR-### and/or NFR-###]

### User Story

As a [specific user/persona],
I want [capability],
so that [business benefit].

### Acceptance Criteria

1. Given [context], when [action],
   then [measurable outcome].

2. Given [context], when [action],
   then [measurable outcome].

3. Given [context], when [action],
   then [measurable outcome].

### Business Value

[One concise sentence explaining
why the story matters.]

Rules:

- Number stories sequentially:
  US-001, US-002, etc.

- Use requirements traceability
  wherever possible.

- Do not invent requirements that
  are not supported by the BRD.

- Write acceptance criteria in
  Given/When/Then format.

- Keep Jira summaries concise
  and action-oriented.

- Assign realistic story points
  based on relative complexity.
"""

            try:

                user_stories = (
                    generate_ai_text(
                        stories_prompt,
                        (
                            "Generating Jira-ready "
                            "user stories..."
                        ),
                    )
                )

                if not user_stories:

                    st.error(
                        "OpenAI returned "
                        "an empty response."
                    )

                else:

                    st.session_state[
                        "generated_user_stories"
                    ] = user_stories

                    st.session_state[
                        "generated_test_cases"
                    ] = None

            except Exception as error:

                st.error(
                    "User story generation "
                    f"error: {error}"
                )


    # =====================================================
    # DISPLAY JIRA STORIES
    # =====================================================

    if (
        st.session_state[
            "generated_user_stories"
        ]
    ):

        user_stories = (
            st.session_state[
                "generated_user_stories"
            ]
        )

        st.success(
            "Jira User Stories Generated"
        )

        st.markdown(
            user_stories
        )

        st.subheader(
            "📥 Export Jira User Stories"
        )

        show_download_buttons(
            "Jira-Ready User Stories",
            user_stories,
            generated_company,
            generated_project_name,
            "Jira_User_Stories",
        )


        # =================================================
        # QA TEST CASES
        # =================================================

        st.divider()

        st.header(
            "🧪 AI-Generated QA Test Cases"
        )

        st.caption(
            "Generate traceable functional "
            "and negative test cases from "
            "the BRD and Jira stories."
        )

        if st.button(
            "🔬 Generate QA Test Cases",
            use_container_width=True,
        ):

            if not api_key:

                st.error(
                    "The OpenAI API key "
                    "was not found."
                )

            else:

                test_prompt = f"""
You are a senior QA Analyst working
with a Business Analyst and Product Owner.

Create professional QA test cases using
ONLY the Business Requirements Document
and Jira user stories provided below.

Company:
{generated_company}

Project:
{generated_project_name}

BUSINESS REQUIREMENTS DOCUMENT:

{requirements}

JIRA USER STORIES:

{user_stories}

Create test coverage for the important
business flows, validation rules,
acceptance criteria, and major
negative scenarios.

Use this exact Markdown structure
for every test case:

## TC-001 — [Concise Test Case Title]

- Linked User Story: [US-###]
- Linked Requirement:
  [FR-### and/or NFR-###]
- Priority: [High, Medium, or Low]
- Test Type:
  [Functional, Negative, Integration,
  Security, or Performance]

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

- Number test cases sequentially:
  TC-001, TC-002, etc.

- Include positive and negative
  coverage where appropriate.

- Preserve traceability to user
  stories and requirements.

- Do not invent functionality not
  supported by the source documents.

- Make steps executable by
  a QA tester.

- Keep expected results objective
  and testable.
"""

                try:

                    test_cases = (
                        generate_ai_text(
                            test_prompt,
                            (
                                "Generating QA "
                                "test cases..."
                            ),
                        )
                    )

                    if not test_cases:

                        st.error(
                            "OpenAI returned "
                            "an empty response."
                        )

                    else:

                        st.session_state[
                            "generated_test_cases"
                        ] = test_cases

                except Exception as error:

                    st.error(
                        "Test case generation "
                        f"error: {error}"
                    )


        # =================================================
        # DISPLAY QA TEST CASES
        # =================================================

        if (
            st.session_state[
                "generated_test_cases"
            ]
        ):

            test_cases = (
                st.session_state[
                    "generated_test_cases"
                ]
            )

            st.success(
                "QA Test Cases Generated"
            )

            st.markdown(
                test_cases
            )

            st.subheader(
                "📥 Export QA Test Cases"
            )

            show_download_buttons(
                "QA Test Cases",
                test_cases,
                generated_company,
                generated_project_name,
                "QA_Test_Cases",
            )