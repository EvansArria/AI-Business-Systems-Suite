import streamlit as st

from generator import generate_brd


st.set_page_config(
    page_title="AI Business Analyst Copilot",
    page_icon="🤖",
    layout="wide",
)

st.title("🤖 AI Business Analyst Copilot")
st.caption("Generate professional business analysis documentation.")

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
    "Stakeholders",
    placeholder="Business Sponsor, Product Owner, Business Analyst...",
    height=100,
)

constraints = st.text_area(
    "Business Constraints",
    placeholder="Budget, timeline, staffing, compliance, technology...",
    height=100,
)

if st.button(
    "✨ Generate Business Requirements",
    use_container_width=True,
):
    required_fields = {
        "Company": company,
        "Project Name": project_name,
        "Business Problem": problem,
        "Business Goal": goal,
    }

    missing_fields = [
        field_name
        for field_name, value in required_fields.items()
        if not value.strip()
    ]

    if missing_fields:
        st.error(
            "Please complete: "
            + ", ".join(missing_fields)
        )
    else:
        brd = generate_brd(
            company=company,
            project=project_name,
            problem=problem,
            goal=goal,
            stakeholders=stakeholders,
            constraints=constraints,
        )

        st.success("Business Requirements Draft Generated")
        st.markdown(brd)