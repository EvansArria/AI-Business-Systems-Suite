import os

import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI


# Load the API key from the .env file.
# override=True forces the app to use the key currently saved in .env.
load_dotenv(override=True)

api_key = os.getenv("OPENAI_API_KEY")

st.set_page_config(
    page_title="AI Business Analyst Copilot",
    page_icon="🤖",
    layout="wide",
)

st.title("🤖 AI Business Analyst Copilot")
st.caption("Generate professional Business Analysis documentation using AI.")

st.divider()

st.subheader("📋 Business Requirements Generator")

company = st.text_input("Company")

project_name = st.text_input("Project Name")

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
            "Check that your .env file contains "
            "OPENAI_API_KEY=your-key."
        )

    else:
        prompt = f"""
You are a senior Business Systems Analyst.

Create a professional Business Requirements Document using the
information below.

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
- Number the functional and non-functional requirements.
- Write user stories in this format:
  "As a [user], I want [capability], so that [business benefit]."
- Make acceptance criteria measurable where possible.
- Do not invent highly specific company facts that were not provided.
"""

        try:
            client = OpenAI(api_key=api_key)

            with st.spinner(
                "Generating business requirements with AI..."
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
                st.success("Business Requirements Generated")

                st.divider()

                st.markdown(requirements)

                st.download_button(
                    label="⬇️ Download Requirements",
                    data=requirements,
                    file_name=(
                        f"{project_name.strip().replace(' ', '_')}"
                        "_business_requirements.md"
                    ),
                    mime="text/markdown",
                    use_container_width=True,
                )

        except Exception as error:
            st.error(f"OpenAI connection error: {error}")