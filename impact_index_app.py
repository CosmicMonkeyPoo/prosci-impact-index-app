import io
import textwrap

import streamlit as st
import pandas as pd

import json
from openai import OpenAI

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


# ------------- CONFIG AND QUESTIONS -------------

CC_QUESTIONS = [
    "Scope of change",
    "Number of impacted employees",
    "Variation in groups that are impacted",
    "Type of change",
    "Degree of process change",
    "Degree of technology and system change",
    "Degree of job role changes",
    "Degree of organization restructuring",
    "Amount of change overall",
    "Impact on employee compensation",
    "Reduction in total staffing levels",
    "Timeframe for change",
]

OA_QUESTIONS = [
    "Perceived need for change among employees and managers",
    "Impact of past changes on employees",
    "Change capacity (how much else is changing)",
    "Past changes (success and management quality)",
    "Shared vision and direction for the organization",
    "Resources and funding availability",
    "Culture and responsiveness to change",
    "Organizational reinforcement (how change is rewarded)",
    "Leadership style and power distribution",
    "Senior management change competency",
    "Middle management change competency",
    "Employee change competency",
]

GROUP_ASPECTS = [
    "Processes",
    "Systems",
    "Tools",
    "Job role",
    "Critical behaviors",
    "Mindset / Attitude / Beliefs",
    "Reporting structure",
    "Performance reviews",
    "Compensation",
    "Location",
]


# ------------- HELPER FUNCTIONS -------------

def generate_change_plan_with_gpt(project_info, group_impacts, oa_impacts=None):
    """
    Call the OpenAI API to generate a high-level change plan
    based on project info and impact variations.
    """
    api_key = st.secrets.get("OPENAI_API_KEY")
    if not api_key:
        st.error("OpenAI API key is not configured. Please set OPENAI_API_KEY in Streamlit secrets.")
        return None

    client = OpenAI(api_key=api_key)

    payload = {
        "project_info": project_info,
        "group_impacts": group_impacts,
        "oa_impacts": oa_impacts,
    }

    system_msg = (
        "You are an expert change management consultant using the Prosci methodology. "
        "You specialize in translating change impact assessments into practical, "
        "role-based change plans for complex organizations."
    )

    user_msg = (
        "Using the following change impact assessment data, create a concise, high-level change plan. "
        "The plan should:\n"
        "- Summarize the overall change and key drivers.\n"
        "- Highlight which groups are most impacted and how.\n"
        "- Recommend tailored change tactics for each group based on their impact level.\n"
        "- Organize tactics into phases (for example: Awareness, Desire, Knowledge, Ability, Reinforcement).\n"
        "- Be written so that a project sponsor or change manager could use it to guide planning.\n\n"
        f"Here is the structured data (JSON):\n{json.dumps(payload, indent=2)}"
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4.1-mini",  # you can change the model later if you want
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ],
            temperature=0.4,
        )
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Error calling OpenAI API: {e}")
        return None

def compute_cc_score(cc_answers):
    """Sum scores for Change Characteristics and compute percentage."""
    total = sum(cc_answers.values())
    max_score = len(cc_answers) * 5
    percent = (total / max_score * 100) if max_score > 0 else 0
    return total, max_score, percent


def compute_oa_score(oa_answers):
    """Sum scores for Organizational Attributes and compute percentage."""
    total = sum(oa_answers.values())
    max_score = len(oa_answers) * 5
    percent = (total / max_score * 100) if max_score > 0 else 0
    return total, max_score, percent


def compute_group_impact(groups_data):
    """
    For each group:
      - count of aspects with score > 0
      - degree of impact on a 0–5 scale, matching the Excel formula:
        IF(SUM(G:P)>0, (SUM(G:P)/50)*5, 0)
    """
    results = []
    for g in groups_data:
        scores = [g["aspects"].get(a, 0) for a in GROUP_ASPECTS]
        total_score = sum(scores)
        aspects_impacted = sum(1 for s in scores if s > 0)
        if total_score > 0:
            degree_impact = (total_score / 50.0) * 5.0
        else:
            degree_impact = 0.0

        results.append({
            "Group name": g["name"],
            "Employees": g["employees"],
            "Aspects impacted (out of 10)": aspects_impacted,
            "Degree of impact (0-5)": round(degree_impact, 2),
        })

    if not results:
        return pd.DataFrame()

    df = pd.DataFrame(results)
    # Make the row index start at 1 instead of 0 for display
    df.index = range(1, len(df) + 1)
    df.index.name = "#"
    return df


def build_pdf_summary(
    project_name,
    sponsor_name,
    org_name,
    assessment_owner,
    project_desc,
    cc_total,
    cc_max,
    cc_pct,
    oa_total,
    oa_max,
    oa_pct,
    cc_answers,
    oa_answers,
    group_df,
):
    """
    Build a summary-only PDF including:
      - Project information
      - CC & OA summary
      - CC & OA items with score >= 3
      - High-impact groups (Degree of impact >= 3)
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def draw_wrapped(text, x, y, max_width, leading=12, font_name="Helvetica", font_size=10):
        """Draw wrapped text and return updated y position."""
        c.setFont(font_name, font_size)
        for line in textwrap.wrap(text, width=max_width):
            c.drawString(x, y, line)
            y -= leading
        return y

    y = height - 50

    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Change Impact Assessment – Summary")
    y -= 30

    # Project information
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Project information")
    y -= 18

    c.setFont("Helvetica", 10)
    if project_name:
        y = draw_wrapped(f"Project: {project_name}", 60, y, 90)
    if org_name:
        y = draw_wrapped(f"Organization / Dept: {org_name}", 60, y, 90)
    if sponsor_name:
        y = draw_wrapped(f"Sponsor: {sponsor_name}", 60, y, 90)
    if assessment_owner:
        y = draw_wrapped(f"Assessment completed by: {assessment_owner}", 60, y, 90)

    if project_desc:
        y -= 6
        c.setFont("Helvetica-Bold", 11)
        c.drawString(50, y, "Change description")
        y -= 14
        y = draw_wrapped(project_desc, 60, y, 95)

    y -= 10

    # CC summary
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Change Characteristics (CC)")
    y -= 16
    c.setFont("Helvetica", 10)
    c.drawString(
        60, y,
        f"Total CC score: {cc_total} / {cc_max} ({cc_pct:.1f}%)"
    )
    y -= 18

    # CC items with score >= 3
    cc_high = []
    for i, q in enumerate(CC_QUESTIONS, start=1):
        score = cc_answers.get(f"CC_{i}", 0)
        if score >= 3:
            cc_high.append((i, q, score))

    if cc_high:
        c.setFont("Helvetica-Bold", 11)
        c.drawString(60, y, "Areas of higher change impact (CC items scored 3 or above):")
        y -= 14
        c.setFont("Helvetica", 10)
        for i, q, score in cc_high:
            line = f"- [{score}] {i}) {q}"
            y = draw_wrapped(line, 70, y, 90)
            if y < 70:
                c.showPage()
                y = height - 50
    else:
        c.setFont("Helvetica", 10)
        c.drawString(60, y, "No CC items scored 3 or above.")
        y -= 16

    y -= 6

    # OA summary
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Organizational Attributes (OA)")
    y -= 16
    c.setFont("Helvetica", 10)
    c.drawString(
        60, y,
        f"Total OA score: {oa_total} / {oa_max} ({oa_pct:.1f}%)"
    )
    y -= 18

    # OA items with score >= 3 (higher risk)
    oa_high = []
    for i, q in enumerate(OA_QUESTIONS, start=1):
        score = oa_answers.get(f"OA_{i}", 0)
        if score >= 3:
            oa_high.append((i, q, score))

    if oa_high:
        c.setFont("Helvetica-Bold", 11)
        c.drawString(60, y, "Areas of higher organizational risk (OA items scored 3 or above):")
        y -= 14
        c.setFont("Helvetica", 10)
        for i, q, score in oa_high:
            line = f"- [{score}] {i}) {q}"
            y = draw_wrapped(line, 70, y, 90)
            if y < 70:
                c.showPage()
                y = height - 50
    else:
        c.setFont("Helvetica", 10)
        c.drawString(60, y, "No OA items scored 3 or above.")
        y -= 16

    y -= 6

        # Group impact summary – all groups
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Group Impact Summary")
    y -= 16
    c.setFont("Helvetica", 10)

    if group_df is not None and not group_df.empty:
        c.drawString(
            60, y,
            "Impacted groups and their degree of impact:"
        )
        y -= 14

        # Iterate over all groups (already indexed 1,2,3... in the app)
        for idx, row in group_df.iterrows():
            line = (
                f"- {row['Group name']} "
                f"(Employees: {row['Employees']}, "
                f"Aspects impacted: {row['Aspects impacted (out of 10)']}, "
                f"Degree of impact: {row['Degree of impact (0-5)']})"
            )
            y = draw_wrapped(line, 70, y, 90)
            if y < 70:
                c.showPage()
                y = height - 50
    else:
        c.drawString(60, y, "No group impact data entered.")
        y -= 16


    # Close out
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

def build_change_plan_pdf(project_info, plan_text):
    """
    Build a PDF containing:
      - Project information (if available)
      - The AI-generated change plan text
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def draw_wrapped(text, x, y, max_width, leading=12, font_name="Helvetica", font_size=10):
        """Draw wrapped text and return updated y position."""
        c.setFont(font_name, font_size)
        for line in textwrap.wrap(text, width=max_width):
            c.drawString(x, y, line)
            y -= leading
        return y

    y = height - 50

    # Title
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "AI-Generated Change Plan")
    y -= 30

    # Project info (if available)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Project information")
    y -= 18
    c.setFont("Helvetica", 10)

    if project_info:
        if project_info.get("project_name"):
            y = draw_wrapped(f"Project: {project_info['project_name']}", 60, y, 90)
        if project_info.get("organization_name"):
            y = draw_wrapped(f"Organization / Dept: {project_info['organization_name']}", 60, y, 90)
        if project_info.get("sponsor_name"):
            y = draw_wrapped(f"Sponsor: {project_info['sponsor_name']}", 60, y, 90)
        if project_info.get("assessment_owner"):
            y = draw_wrapped(f"Assessment completed by: {project_info['assessment_owner']}", 60, y, 90)

        if project_info.get("description"):
            y -= 6
            c.setFont("Helvetica-Bold", 11)
            c.drawString(50, y, "Change description")
            y -= 14
            y = draw_wrapped(project_info["description"], 60, y, 95)

    y -= 10

    # Plan body
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Recommended Change Plan")
    y -= 18

    c.setFont("Helvetica", 10)
    if not plan_text:
        y = draw_wrapped("No plan text generated.", 60, y, 95)
    else:
        # Allow for long text with page breaks
        for paragraph in plan_text.split("\n"):
            paragraph = paragraph.strip()
            if not paragraph:
                y -= 6
                continue
            y = draw_wrapped(paragraph, 60, y, 95)
            if y < 70:
                c.showPage()
                y = height - 50

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer


# ------------- STREAMLIT APP -------------

st.set_page_config(
    page_title="Prosci Impact Index – Impact Assessment",
    layout="wide"
)

# --- Custom colors: pink bars + blue headings ---
st.markdown(
    """
    <style>
    /* Headings */
    h1, h2, h3, h4, h5, h6 {
        color: #06AFE6 !important;
    }

    /* Buttons */
    .stButton>button {
        background-color: #DA10AB !important;
        border-color: #DA10AB !important;
        color: white !important;
    }

    /* Slider thumb */
    .stSlider [role="slider"] {
        background-color: #DA10AB !important;
        border-color: #DA10AB !important;
    }

    /* Filled track (blue → pink gradient) */
    .stSlider [data-baseweb="slider"] > div > div > div:nth-child(2) {
        background: linear-gradient(90deg, #06AFE6 0%, #DA10AB 100%) !important;
    }

    /* Unfilled track (light grey) */
    .stSlider [data-baseweb="slider"] > div > div > div:nth-child(3) {
        background-color: #E0E0E0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


st.title("Prosci Impact Index – Impact Assessment App")

st.markdown(
    """
    This app is a simplified, Excel-free interface for the Prosci Impact Index.
    Fill in the sections below to assess change risk and group-level impact.
    """
)

# ---------- SECTION 1: PROJECT BASICS ----------

st.header("1. Project information")

col1, col2 = st.columns(2)

with col1:
    project_name = st.text_input("Project name", value="")
    sponsor_name = st.text_input("Primary sponsor name", value="")

with col2:
    org_name = st.text_input("Organization or department", value="")
    assessment_owner = st.text_input("Assessment completed by", value="")

project_desc = st.text_area(
    "Short description of the change",
    placeholder="Describe the change at a high level (what is changing and why)."
)

st.markdown("---")

# ---------- SECTION 2: CHANGE CHARACTERISTICS ----------

st.header("2. Change Characteristics (1–5 scale)")

st.markdown(
    "Rate each item from **1 (low impact)** to **5 (high impact)** based on the characteristics of this change."
)

cc_answers = {}

for i, q in enumerate(CC_QUESTIONS, start=1):
    cc_answers[f"CC_{i}"] = st.slider(
        f"{i}) {q}",
        min_value=1,
        max_value=5,
        value=3,
        step=1,
    )

cc_total, cc_max, cc_pct = compute_cc_score(cc_answers)

st.subheader("Change Characteristics summary")
st.write(f"Total CC score: **{cc_total}** out of {cc_max}")
st.write(f"Percent of maximum: **{cc_pct:.1f}%**")

st.markdown("---")

# ---------- SECTION 3: ORGANIZATIONAL ATTRIBUTES ----------

st.header("3. Organizational Attributes (1–5 scale)")

st.markdown(
    "Rate each item from **1 (low risk, more favorable)** to **5 (high risk, less favorable)** "
    "based on how the organization currently operates."
)

oa_answers = {}

for i, q in enumerate(OA_QUESTIONS, start=1):
    oa_answers[f"OA_{i}"] = st.slider(
        f"{i}) {q}",
        min_value=1,
        max_value=5,
        value=3,
        step=1,
    )

oa_total, oa_max, oa_pct = compute_oa_score(oa_answers)

st.subheader("Organizational Attributes summary")
st.write(f"Total OA score: **{oa_total}** out of {oa_max}")
st.write(f"Percent of maximum: **{oa_pct:.1f}%**")

st.markdown("---")

# ---------- SECTION 4: GROUP IMPACT INVENTORY ----------

st.header("4. Group Impact Inventory")

st.markdown(
    """
For each impacted group, capture how strongly the change affects different aspects of their work.
Use a **0–5** scale where:

- 0 = no impact  
- 1 = very low impact  
- 5 = extremely high impact
"""
)

num_groups = st.number_input(
    "How many groups would you like to assess?",
    min_value=1,
    max_value=24,
    value=3,
    step=1,
)

groups_data = []

for i in range(int(num_groups)):
    st.subheader(f"Group {i + 1}")
    with st.expander(f"Details for group {i + 1}", expanded=True if i == 0 else False):
        g_name = st.text_input(
            "Group name",
            key=f"group_name_{i}",
            placeholder="Example: Customer Service, Finance, IT Operations"
        )
        g_employees = st.number_input(
            "Number of employees in this group",
            min_value=0,
            value=0,
            step=1,
            key=f"group_employees_{i}",
        )

        st.markdown("### Impact on aspects (0–5)")
        aspect_scores = {}
        for aspect in GROUP_ASPECTS:
            aspect_scores[aspect] = st.slider(
                aspect,
                min_value=0,
                max_value=5,
                value=0,
                step=1,
                key=f"group_{i}_{aspect}",
            )

        groups_data.append({
            "name": g_name,
            "employees": g_employees,
            "aspects": aspect_scores,
        })

group_df = compute_group_impact(groups_data)

st.subheader("Group impact summary")
if group_df.empty:
    st.info("Fill in at least one group with some non-zero impact scores to see results.")
else:
    st.dataframe(group_df, use_container_width=True)

    # ---------- EXCEL EXPORT (Group Impact + OA) ----------
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        group_df.to_excel(writer, sheet_name="Group Impact")
        oa_summary_df = pd.DataFrame({
            "Metric": ["Total OA score", "Max OA score", "Percent of max"],
            "Value": [oa_total, oa_max, oa_pct],
        })
        oa_details_df = pd.DataFrame({
            "Question": OA_QUESTIONS,
            "Score": list(oa_answers.values()),
        })
        oa_summary_df.to_excel(writer, sheet_name="OA Summary", index=False)
        oa_details_df.to_excel(writer, sheet_name="OA Details", index=False)
    excel_buffer.seek(0)

    st.download_button(
        label="Download impact results as Excel",
        data=excel_buffer,
        file_name="impact_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")

# ---------- SECTION 5: OVERALL SUMMARY & PDF EXPORT ----------

st.header("5. Overall impact summary")

col_a, col_b = st.columns(2)

with col_a:
    st.subheader("Change risk scores")
    st.write(f"- Change Characteristics: **{cc_total} / {cc_max}** ({cc_pct:.1f}%)")
    st.write(f"- Organizational Attributes: **{oa_total} / {oa_max}** ({oa_pct:.1f}%)")

with col_b:
    st.subheader("Top impacted groups")
    if not group_df.empty:
        top_groups = group_df.sort_values(
            by="Degree of impact (0-5)",
            ascending=False
        ).head(5)
        st.write(top_groups[["Group name", "Employees", "Degree of impact (0-5)"]])
    else:
        st.write("No group impact data yet.")

# PDF summary (OA + Group Impact) – SUMMARY-ONLY PDF
pdf_buffer = build_pdf_summary(
    project_name=project_name,
    sponsor_name=sponsor_name,
    org_name=org_name,
    assessment_owner=assessment_owner,
    project_desc=project_desc,
    cc_total=cc_total,
    cc_max=cc_max,
    cc_pct=cc_pct,
    oa_total=oa_total,
    oa_max=oa_max,
    oa_pct=oa_pct,
    cc_answers=cc_answers,
    oa_answers=oa_answers,
    group_df=group_df,
)

# Convert buffer to raw bytes for Streamlit download
pdf_bytes = pdf_buffer.getvalue()

st.download_button(
    label="Download PDF Summary",
    data=pdf_bytes,
    file_name="impact_summary.pdf",
    mime="application/pdf"
)

st.markdown(
    """
You can use these scores to determine where targeted change management activity is most needed:
- Higher **CC** and **OA** scores indicate higher overall change risk.  
- Groups with higher **Degree of impact** and many **Aspects impacted** will likely need more support and tailored change plans.
"""
)

# ------------------------------
# AI-Generated Change Plan section
# ------------------------------
st.markdown("---")
st.subheader("AI-Generated Change Plan (Optional)")

st.markdown(
    "Click the button below to send this impact assessment to an AI assistant and "
    "generate a high-level change plan that accounts for the different groups and their impact levels."
)

# Ensure a place in session_state for the plan
if "change_plan" not in st.session_state:
    st.session_state["change_plan"] = None

# Build group impacts structure from the actual group_df
group_impacts = None
if group_df is not None and not group_df.empty:
    # Include the row index (#) as well for clarity
    group_impacts = group_df.reset_index().to_dict(orient="records")

# Build OA impacts structure from current OA scores
oa_impacts = {
    "summary": {
        "total_oa_score": oa_total,
        "max_oa_score": oa_max,
        "percent_of_max": oa_pct,
    },
    "details": [
        {
            "id": i,
            "question": q,
            "score": oa_answers.get(f"OA_{i}", None),
        }
        for i, q in enumerate(OA_QUESTIONS, start=1)
    ],
}

# Project info passed to the AI
project_info = {
    "project_name": project_name,
    "sponsor_name": sponsor_name,
    "organization_name": org_name,
    "assessment_owner": assessment_owner,
    "description": project_desc,
}

# Button to generate the plan
if st.button("Generate AI Change Plan"):
    with st.spinner("Generating change plan..."):
        plan = generate_change_plan_with_gpt(
            project_info=project_info,
            group_impacts=group_impacts,
            oa_impacts=oa_impacts,
        )

    if plan:
        st.session_state["change_plan"] = plan
        st.success("Change plan generated.")

# If we have a saved plan, display it and offer PDF download
if st.session_state.get("change_plan"):
    st.markdown("### Recommended Change Plan")
    st.write(st.session_state["change_plan"])

    # Build PDF of the AI change plan
    plan_pdf_buffer = build_change_plan_pdf(project_info, st.session_state["change_plan"])
    plan_pdf_bytes = plan_pdf_buffer.getvalue()

    st.download_button(
        label="Download Change Plan as PDF",
        data=plan_pdf_bytes,
        file_name="change_plan.pdf",
        mime="application/pdf",
    )
