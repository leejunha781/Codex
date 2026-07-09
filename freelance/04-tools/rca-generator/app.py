"""RCA Report Generator — Streamlit UI."""

from __future__ import annotations

import streamlit as st

from report_builder import build_docx_report, build_markdown_report, load_sample_data

st.set_page_config(
    page_title="RCA Report Generator",
    page_icon="📋",
    layout="wide",
)

st.title("RCA Report Generator")
st.caption(
    "Engineering Document Automation — portfolio demo by Joonha Lee. "
    "Turn field issue inputs into a structured RCA report draft."
)

with st.sidebar:
    st.header("Actions")
    if st.button("Load sample data", use_container_width=True):
        st.session_state["form"] = load_sample_data()
        st.rerun()
    st.markdown("---")
    st.markdown(
        "**Services:** Technical documentation · FAT/SAT · Proposals  \n"
        "**Contact:** leejunha781@gmail.com"
    )

defaults = st.session_state.get("form", {})

col_left, col_right = st.columns(2)

with col_left:
    st.subheader("Issue intake")
    doc_id = st.text_input("Document ID", value=defaults.get("doc_id", "RCA-2026-001"))
    report_date = st.text_input("Report date", value=defaults.get("report_date", ""))
    prepared_by = st.text_input("Prepared by", value=defaults.get("prepared_by", "Joonha Lee"))
    status = st.selectbox(
        "Status",
        ["Open", "In progress", "Closed — corrective action verified"],
        index=["Open", "In progress", "Closed — corrective action verified"].index(
            defaults.get("status", "Open")
        )
        if defaults.get("status") in ["Open", "In progress", "Closed — corrective action verified"]
        else 0,
    )
    equipment_name = st.text_input("Equipment name", value=defaults.get("equipment_name", ""))
    customer_site = st.text_input("Customer / site", value=defaults.get("customer_site", ""))
    severity_options = ["Low", "Medium", "High", "Critical"]
    default_severity = defaults.get("severity", "Medium")
    if default_severity not in severity_options:
        default_severity = default_severity.split(" ")[0] if default_severity else "Medium"
    if default_severity not in severity_options:
        default_severity = "Medium"
    severity = st.selectbox(
        "Severity",
        severity_options,
        index=severity_options.index(default_severity),
    )
    symptom = st.text_area("Symptom summary", value=defaults.get("symptom", ""), height=80)

with col_right:
    st.subheader("Analysis & closure")
    problem_description = st.text_area(
        "Problem description",
        value=defaults.get("problem_description", ""),
        height=100,
    )
    environment = st.text_area("Environment", value=defaults.get("environment", ""), height=80)
    test_results = st.text_area("Test results & evidence", value=defaults.get("test_results", ""), height=100)
    root_cause = st.text_area("Root cause", value=defaults.get("root_cause", ""), height=80)
    contributing_factors = st.text_area(
        "Contributing factors",
        value=defaults.get("contributing_factors", ""),
        height=60,
    )
    corrective_actions = st.text_area(
        "Corrective actions",
        value=defaults.get("corrective_actions", ""),
        height=80,
    )
    verification = st.text_area("Verification", value=defaults.get("verification", ""), height=80)
    preventive_recommendations = st.text_area(
        "Preventive recommendations",
        value=defaults.get("preventive_recommendations", ""),
        height=60,
    )
    executive_summary = st.text_area(
        "Executive summary (auto-suggest: fill last)",
        value=defaults.get("executive_summary", ""),
        height=80,
    )

form_data = {
    "doc_id": doc_id,
    "report_date": report_date,
    "prepared_by": prepared_by,
    "status": status,
    "equipment_name": equipment_name,
    "customer_site": customer_site,
    "severity": severity,
    "symptom": symptom,
    "problem_description": problem_description,
    "environment": environment,
    "test_results": test_results,
    "root_cause": root_cause,
    "contributing_factors": contributing_factors,
    "corrective_actions": corrective_actions,
    "verification": verification,
    "preventive_recommendations": preventive_recommendations,
    "executive_summary": executive_summary,
}

st.session_state["form"] = form_data

st.divider()
st.subheader("Preview & export")

markdown_report = build_markdown_report(form_data)
st.markdown(markdown_report)

col_a, col_b = st.columns(2)
with col_a:
    st.download_button(
        label="Download Markdown (.md)",
        data=markdown_report,
        file_name=f"{doc_id.replace(' ', '_')}.md",
        mime="text/markdown",
        use_container_width=True,
    )
with col_b:
    docx_bytes = build_docx_report(form_data)
    st.download_button(
        label="Download Word (.docx)",
        data=docx_bytes,
        file_name=f"{doc_id.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
