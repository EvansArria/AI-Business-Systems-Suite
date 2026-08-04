def generate_brd(company, project, problem, goal, stakeholders, constraints):
    return f"""
# Executive Summary

Company:
{company}

Project:
{project}

Business Problem:
{problem}

Business Goal:
{goal}

Stakeholders:
{stakeholders}

Constraints:
{constraints}

---

## Business Objectives

- Improve operational efficiency
- Reduce manual work
- Increase productivity
- Improve customer satisfaction

---

## Functional Requirements

- Capture business information
- Store project details
- Generate documentation
- Export documentation

---

## Non-Functional Requirements

- Secure
- Reliable
- Fast
- Easy to use
"""