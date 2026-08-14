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

# Local development:
# Reads OPENAI_API_KEY from the .env file.
load_dotenv(override=True)

api_key = os.getenv("OPENAI_API_KEY")

# Streamlit Cloud fallback:
# Reads OPENAI_API_KEY from Streamlit Secrets if needed.
if not api_key:
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        api_key = None


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_inline_markdown(text):
    """
    Removes simple Markdown symbols so exported
    Word/PDF files look cleaner.
    """
    return (
        text
        .replace("**", "")
        .replace("__", "")
        .replace("`", "")
    )


# =========================================================
# WORD EXPORT FUNCTION
# =========================================================

def create_word_document(
    requirements_text,
    company,
    project_name,
):
    buffer = BytesIO()

    document = Document()

    document.add_heading(
        "AI Business Analyst Copilot",
        level=0,
    )

    document.add_paragraph(
        f"Company: {company}"
    )

    document.add_paragraph(
        f"Project: {project_name}"
    )

    document.add_paragraph("")

    for line in requirements_text.splitlines():

        cleaned_line = line.strip()

        if not cleaned_line:
            document.add_paragraph("")
            continue

        if cleaned_line.startswith("### "):

            heading_text = clean_inline_markdown(
                cleaned_line[4:]
            )

            document.add_heading(
                heading_text,
                level=3,
            )

        elif cleaned_line.startswith("## "):

            heading_text = clean_inline_markdown(
                cleaned_line[3:]
            )

            document.add_heading(
                heading_text,
                level=2,
            )

        elif cleaned_line.startswith("# "):

            heading_text = clean_inline_markdown(
                cleaned_line[2:]
            )

            document.add_heading(
                heading_text,
                level=1,
            )

        elif cleaned_line.startswith("- "):

            bullet_text = clean_inline_markdown(
                cleaned_line[2:]
            )

            document.add_paragraph(
                bullet_text,
                style="List Bullet",
            )

        else:

            document.add_paragraph(
                clean_inline_markdown(cleaned_line)
            )

    document.save(buffer)

    buffer.seek(0)

    return buffer


# =========================================================
# PDF EXPORT FUNCTION
# =========================================================

def create_pdf_document(
    requirements_text,
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
            "AI Business Analyst Copilot",
            styles["Title"],
        )
    )

    story.append(
        Spacer(1, 12)
    )

    story.append(
        Paragraph(
            f"<b>Company:</b> {escape(company)}",
            styles["Normal"],
        )
    )

    story.append(
        Paragraph(
            f"<b>Project:</b> {escape(project_name)}",
            styles["Normal"],
        )
    )

    story.append(
        Spacer(1, 18)
    )

    for line in requirements_text.splitlines():

        cleaned_line = line.strip()

        if not cleaned_line:
            story.append(
                Spacer(1, 8)
            )
            continue

        if cleaned_line.startswith("### "):

            heading_text = clean_inline_markdown(
                cleaned_line[4:]
            )

            story.append(
                Paragraph(
                    escape(heading_text),
                    styles["Heading3"],
                )
            )

        elif cleaned_line.startswith("## "):

            heading_text = clean_inline_markdown(
                cleaned_line[3:]
            )

            story.append(
                Paragraph(
                    escape(heading_text),
                    styles["Heading2"],
                )
            )

        elif cleaned_line.startswith("# "):

            heading_text = clean_inline_markdown(
                cleaned_line[2:]
            )

            story.append(
                Paragraph(
                    escape(heading_text),
                    styles["Heading1"],
                )
            )

        elif cleaned_line.startswith("- "):

            bullet_text = clean_inline_markdown(
                cleaned_line[2:]
            )

            story.append(
                Paragraph(
                    "&bull; " + escape(bullet_text),
                    styles["Normal"],
                )
            )

        else:

            paragraph_text = clean_inline_markdown(
                cleaned_line
            )

            story.append(
                Paragraph(
                    escape(paragraph_text),
                    styles["Normal"],
                )
            )

        story.append(
            Spacer(1, 6)
        )

    pdf.build(story)

    buffer.seek(0)

    return buffer


# =========================================================
# SESSION STATE
# Keeps generated content available after button actions.
# =========================================================

if "generated_requirements" not in st.session_state:
    st.session_state.generated_requirements = None

if "generated_company" not in st.session_state:
    st.session_state.generated_company = ""

if "generated_project_name" not in st.session_state:
    st.session_state.generated_project_name = ""


# =========================================================
# APP HEADER
# =========================================================

st.title("🤖 AI Business Analyst Copilot")

st.caption(
    "Generate professional Business Analysis "
    "documentation using AI."
)

st.divider()


# =========================================================
# BUSINESS REQUIREMENTS INPUT FORM
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
# GENERATE BUTTON
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
            "Please complete Company, Project Name, "
            "Business Problem, and Business Goal."
        )

    elif not api_key:

        st.error(
            "The OpenAI API key was not found. "
            "Check your .env file locally or "
            "Streamlit Secrets in the live app."
        )

    else:

        prompt = f"""
You are a senior Business Systems Analyst.

Create a professional Business Requirements Document
using the information below.

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

- Number functional requirements using:
  FR-001, FR-002, FR-003, etc.

- Number non-functional requirements using:
  NFR-001, NFR-002, NFR-003, etc.

- Write user stories using this structure:
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

            client = OpenAI(
                api_key=api_key
            )

            with st.spinner(
                "Generating business requirements "
                "with AI..."
            ):

                response = client.responses.create(
                    model="gpt-4.1-mini",
                    input=prompt,
                )

            requirements = response.output_text

            if not requirements:

                st.error(
                    "OpenAI returned an empty response. "
                    "Please try again."
                )

            else:

                st.session_state.generated_requirements = (
                    requirements
                )

                st.session_state.generated_company = (
                    company
                )

                st.session_state.generated_project_name = (
                    project_name
                )

        except Exception as error:

            st.error(
                f"OpenAI connection error: {error}"
            )


# =========================================================
# DISPLAY GENERATED BRD
# =========================================================

if st.session_state.generated_requirements:

    requirements = (
        st.session_state.generated_requirements
    )

    generated_company = (
        st.session_state.generated_company
    )

    generated_project_name = (
        st.session_state.generated_project_name
    )

    st.success(
        "Business Requirements Generated"
    )

    st.divider()

    st.header(
        "📄 Generated Business Requirements Document"
    )

    st.markdown(
        requirements
    )


    # =====================================================
    # CREATE EXPORT FILES
    # =====================================================

    safe_project_name = (
        generated_project_name
        .strip()
        .replace(" ", "_")
    )

    if not safe_project_name:
        safe_project_name = "Business_Requirements"


    word_file = create_word_document(
        requirements,
        generated_company,
        generated_project_name,
    )


    pdf_file = create_pdf_document(
        requirements,
        generated_company,
        generated_project_name,
    )


    # =====================================================
    # DOWNLOAD BUTTONS
    # =====================================================

    st.divider()

    st.subheader(
        "📥 Export Business Requirements"
    )

    col1, col2, col3 = st.columns(3)


    with col1:

        st.download_button(
            label="📘 Download Word",
            data=word_file,
            file_name=(
                f"{safe_project_name}_BRD.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.wordprocessingml.document"
            ),
            use_container_width=True,
        )


    with col2:

        st.download_button(
            label="📕 Download PDF",
            data=pdf_file,
            file_name=(
                f"{safe_project_name}_BRD.pdf"
            ),
            mime="application/pdf",
            use_container_width=True,
        )


    with col3:

        st.download_button(
            label="📝 Download Markdown",
            data=requirements,
            file_name=(
                f"{safe_project_name}_BRD.md"
            ),
            mime="text/markdown",
            use_container_width=True,
        )