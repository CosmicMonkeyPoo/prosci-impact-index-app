import io
import textwrap
import re  

import streamlit as st
import pandas as pd

import json
from openai import OpenAI

# Updated imports for the PDF generation
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


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

    # UPDATED PROMPT: Explicitly forbids ASCII tables to prevent PDF formatting issues
    system_msg = (
        "You are an expert change management consultant using the Prosci methodology. "
        "You specialize in translating change impact assessments into practical, "
        "role-based change plans for complex organizations.\n\n"
        "IMPORTANT FORMATTING RULES:\n"
        "1. Do NOT use Markdown tables (ASCII tables with | and -). They break the PDF rendering.\n"
        "2. Instead of tables, use clear bulleted lists or grouped text sections.\n"
        "3. Use '### ' (triple hash) for your Section Headers so we can style them.\n"
        "4. Do not use HTML tags like <br>."
    )

    user_msg = (
        "Using the following change impact assessment data, create a concise, high-level change plan. "
        "The plan should:\n"
        "- Summarize the overall change and key drivers.\n"
        "- Highlight which groups are most impacted and how (Use a list, not a table).\n"
        "- Recommend tailored change tactics for each group based on their impact level.\n"
        "- Organize tactics into phases (for example: Awareness, Desire, Knowledge, Ability, Reinforcement).\n"
        "- Be written so that a project sponsor or change manager could use it to guide planning.\n\n"
        f"Here is the structured data (JSON):\n{json.dumps(payload, indent=2)}"
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4.1-mini", 
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
            "Degree of impact (0-5)": round(degree_impact, 1), # Rounded to nearest 10th
        })

    if not results:
        return pd.DataFrame()

    df = pd.DataFrame(results)
    # Make the row index start at 1 instead of 0 for display
    df.index = range(1, len(df) + 1)
    df.index.name = "#"
    return df

def style_impact_table(df: pd.DataFrame):
    """
    Style a DataFrame with:
      - black background
      - blue borders
      - pink text
      - Number formatting (1 decimal place)
    """
    # Define the blue border string for reuse
    blue_border = "1px solid #06AFE6"
    
    return (
        df.style
        .format({"Degree of impact (0-5)": "{:.1f}"})  # Force 1 decimal place (e.g., 3.0)
        .set_properties(**{
            "background-color": "#000000",
            "color": "#DA10AB",           # Pink text
            "border": blue_border,        # Blue border
        })
        .set_table_styles([
            {
                "selector": "th", 
                "props": [
                    ("background-color", "#000000"), 
                    ("color", "#DA10AB"), 
                    ("border", blue_border)
                ]
            },
            {
                "selector": "tr", 
                "props": [("background-color", "#000000")]
            },
            {
                "selector": "td", 
                "props": [("border", blue_border)]
            }
        ])
    )


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
    Build a PDF summary using ReportLab Platypus for proper tables and styling.
    Theme: White background, Blue Headings, Pink Table Text, Blue Borders.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story = []
    
    # --- Styles ---
    styles = getSampleStyleSheet()
    
    # Custom Heading (Blue #06AFE6)
    title_style = ParagraphStyle(
        'TitleCustom', 
        parent=styles['Heading1'], 
        textColor=colors.HexColor("#06AFE6"),
        spaceAfter=12
    )
    h2_style = ParagraphStyle(
        'Heading2Custom', 
        parent=styles['Heading2'], 
        textColor=colors.HexColor("#06AFE6"),
        spaceBefore=12, 
        spaceAfter=6
    )
    
    # Normal Text (Black for readability on white paper)
    normal_style = styles['Normal']
    
    # Pink Text for Table Content
    pink_text_style = ParagraphStyle(
        'PinkText',
        parent=styles['Normal'],
        textColor=colors.HexColor("#DA10AB")
    )
    
    # --- Content ---
    
    # Title
    story.append(Paragraph("Change Impact Assessment – Summary", title_style))
    
    # 1. Project Info
    story.append(Paragraph("1. Project information", h2_style))
    
    # Helper to add bold label + text
    def add_info_line(label, value):
        if value:
            text = f"<b>{label}</b> {value}"
            story.append(Paragraph(text, normal_style))
            story.append(Spacer(1, 4))

    add_info_line("Project:", project_name)
    add_info_line("Organization / Dept:", org_name)
    add_info_line("Sponsor:", sponsor_name)
    add_info_line("Assessment completed by:", assessment_owner)
    
    if project_desc:
        story.append(Spacer(1, 6))
        story.append(Paragraph("<b>Change description:</b>", normal_style))
        story.append(Paragraph(project_desc, normal_style))

    story.append(Spacer(1, 12))

    # 2. Change Characteristics
    story.append(Paragraph("2. Change Characteristics (CC)", h2_style))
    story.append(Paragraph(f"<b>Total Score:</b> {cc_total} / {cc_max} ({cc_pct:.1f}%)", normal_style))
    
    # CC High Impact
    cc_high = [q for i, q in enumerate(CC_QUESTIONS, 1) if cc_answers.get(f"CC_{i}", 0) >= 3]
    if cc_high:
        story.append(Spacer(1, 6))
        story.append(Paragraph("<b>High impact areas (Score 3+):</b>", normal_style))
        for item in cc_high:
            story.append(Paragraph(f"• {item}", normal_style))
    else:
        story.append(Paragraph("No items scored 3 or above.", normal_style))

    # 3. Organizational Attributes
    story.append(Paragraph("3. Organizational Attributes (OA)", h2_style))
    story.append(Paragraph(f"<b>Total Score:</b> {oa_total} / {oa_max} ({oa_pct:.1f}%)", normal_style))
    
    # OA High Impact
    oa_high = [q for i, q in enumerate(OA_QUESTIONS, 1) if oa_answers.get(f"OA_{i}", 0) >= 3]
    if oa_high:
        story.append(Spacer(1, 6))
        story.append(Paragraph("<b>High risk areas (Score 3+):</b>", normal_style))
        for item in oa_high:
            story.append(Paragraph(f"• {item}", normal_style))
    else:
        story.append(Paragraph("No items scored 3 or above.", normal_style))
        
    story.append(Spacer(1, 12))

    # 4. Group Impact Summary (THE TABLE)
    story.append(Paragraph("4. Group Impact Summary", h2_style))
    
    if group_df is not None and not group_df.empty:
        # Construct Table Data
        # Headers
        table_data = [[
            Paragraph("<b>Group Name</b>", pink_text_style),
            Paragraph("<b>Employees</b>", pink_text_style),
            Paragraph("<b>Aspects (10)</b>", pink_text_style),
            Paragraph("<b>Impact (0-5)</b>", pink_text_style)
        ]]
        
        # Rows
        for _, row in group_df.iterrows():
            # Round impact to 1 decimal
            impact_val = f"{row['Degree of impact (0-5)']:.1f}"
            
            row_data = [
                Paragraph(str(row['Group name']), pink_text_style),
                Paragraph(str(row['Employees']), pink_text_style),
                Paragraph(str(row['Aspects impacted (out of 10)']), pink_text_style),
                Paragraph(impact_val, pink_text_style)
            ]
            table_data.append(row_data)
        
        # Create Table
        # Adjust col widths as needed
        t = Table(table_data, colWidths=[3.0*inch, 1.0*inch, 1.2*inch, 1.2*inch])
        
        # Apply the "Webpage Look" (White BG, Pink Text, Blue Grid)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.white),       # White Background
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#06AFE6")), # Blue Borders
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),                # Center numbers
        ]))
        
        story.append(t)
    else:
        story.append(Paragraph("No group data entered.", normal_style))
        
    # Build
    doc.build(story)
    buffer.seek(0)
    return buffer

def build_change_plan_pdf(project_info, plan_text):
    """
    Build a PDF for the AI Change Plan.
    Theme: White background, Blue Headings.
    Parses Markdown headers (###) into real PDF styles.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story = []
    
    # --- Styles ---
    styles = getSampleStyleSheet()
    
    # Blue Headings
    title_style = ParagraphStyle('TitleCustom', parent=styles['Heading1'], textColor=colors.HexColor("#06AFE6"), spaceAfter=12)
    h2_style = ParagraphStyle('Heading2Custom', parent=styles['Heading2'], textColor=colors.HexColor("#06AFE6"), spaceBefore=12, spaceAfter=6)
    
    # Normal Text
    normal_style = styles['Normal']
    
    # --- Content ---
    story.append(Paragraph("AI-Generated Change Plan", title_style))
    
    # Project Info
    story.append(Paragraph("Project Information", h2_style))
    if project_info:
        if project_info.get("project_name"):
            story.append(Paragraph(f"<b>Project:</b> {project_info['project_name']}", normal_style))
        if project_info.get("sponsor_name"):
            story.append(Paragraph(f"<b>Sponsor:</b> {project_info['sponsor_name']}", normal_style))
            
    story.append(Spacer(1, 12))
    
    # The Plan Content Parsing
    # We don't add a hardcoded header here because the AI usually provides its own headers.
    
    if plan_text:
        # Split by newlines
        lines = plan_text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 1. Clean up <br> tags and XML characters
            clean_line = line.replace('<br>', '').replace('<br/>', '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            # 2. Check for Markdown Headers (### Header)
            if clean_line.startswith('###'):
                # Strip the hashtags and whitespace
                header_text = clean_line.replace('#', '').strip()
                # Add as a Styled Heading
                story.append(Paragraph(header_text, h2_style))
            
            # 3. Check for Markdown Headers (## Header) - just in case
            elif clean_line.startswith('##'):
                header_text = clean_line.replace('#', '').strip()
                story.append(Paragraph(header_text, h2_style))

            # 4. Standard Text
            else:
                # Apply Bold Formatting safely (Regex)
                formatted_text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', clean_line)
                
                try:
                    story.append(Paragraph(formatted_text, normal_style))
                except Exception:
                    story.append(Paragraph(clean_line, normal_style))
                
                # Add a small spacer after paragraphs for readability
                story.append(Spacer(1, 6))
    else:
        story.append(Paragraph("No plan generated.", normal_style))

    doc.build(story)
    buffer.seek(0)
    return buffer

# ------------- STREAMLIT APP -------------

st.set_page_config(
    page_title="Prosci Impact Index – Impact Assessment",
    layout="wide"
)

# --- Custom colors: pink bars + blue headings ---
# --- Custom colors: dark theme + pink/blue accents ---
# --- Custom dark theme + colors ---
st.markdown(
    """
    <style>
    /* ---------- Top header bar (very top of page) ---------- */
    [data-testid="stHeader"] {
        background-color: #DA10AB !important;  /* pink bar */
    }
    [data-testid="stHeader"] * {
        color: #FFFFFF !important;  /* white icons/text in the top bar */
    }

    /* ---------- Base layout & background ---------- */
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #000000 !important;  /* pure black */
        color: #FFFFFF !important;             /* default text white */
    }

    [data-testid="stAppViewContainer"] > .main {
        background-color: #000000 !important;
    }

    [data-testid="stSidebar"] {
        background-color: #050505 !important;
    }

    /* ---------- Headings in main content ---------- */
    h1, h2, h3, h4, h5, h6 {
        color: #06AFE6 !important;   /* your blue for headings */
    }

    /* ---------- General text / labels ---------- */
    label,
    .stMarkdown,
    .stTextInput label,
    .stNumberInput label,
    .stSlider label {
        color: #FFFFFF !important;
    }

    /* ---------- Inputs & text areas ---------- */
    input, textarea {
        background-color: #111111 !important;
        color: #FFFFFF !important;
        border: 1px solid #444444 !important;
    }

    /* Placeholder / helper text */
    input::placeholder,
    textarea::placeholder {
        color: #BBBBBB !important;  /* light grey on dark */
    }

    /* ---------- Buttons (all types) ---------- */
    .stButton > button,
    .stDownloadButton > button,
    button[kind="primary"],
    button[kind="secondary"] {
        background-color: #DA10AB !important;   /* pink */
        color: #FFFFFF !important;              /* white text */
        border: 1px solid #DA10AB !important;
        border-radius: 4px !important;
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover,
    button[kind="primary"]:hover,
    button[kind="secondary"]:hover {
        background-color: #b10b8a !important;   /* darker pink on hover */
        border-color: #b10b8a !important;
        color: #FFFFFF !important;
    }

    /* ---------- Sliders (thumb + track) ---------- */
    [data-baseweb="slider"] [role="slider"] {
        background-color: #DA10AB !important;  /* pink thumb */
        border-color: #DA10AB !important;
    }

    /* Filled track: blue → pink gradient */
    [data-baseweb="slider"] > div > div > div:nth-child(2) {
        background: linear-gradient(90deg, #06AFE6 0%, #DA10AB 100%) !important;
    }

    /* Unfilled track */
    [data-baseweb="slider"] > div > div > div:nth-child(3) {
        background-color: #333333 !important;
    }

    /* ---------- Expanders (“Details for group #”) ---------- */
    div[data-testid="stExpander"] {
        background-color: #111111 !important;
        color: #FFFFFF !important;
        border: 1px solid #333333 !important;
    }

    div[data-testid="stExpander"] summary {
        background-color: #111111 !important;
        color: #FFFFFF !important;
    }

    div[data-testid="stExpander"] summary svg {
        fill: #FFFFFF !important;
    }

    /* ---------- Tables (replaces DataFrame styling) ---------- */
    /* Target the st.table specifically */
    [data-testid="stTable"] {
        background-color: #000000 !important;
        width: 100%;
    }
    
    /* Force the table structure to be black */
    [data-testid="stTable"] table {
        background-color: #000000 !important;
        border-collapse: collapse !important;
    }
    
    /* Target headers (th) - Black bg, Pink Text, Blue Border */
    [data-testid="stTable"] th {
        background-color: #000000 !important;
        color: #DA10AB !important;            /* Pink text */
        border: 1px solid #06AFE6 !important; /* Blue border */
        font-weight: bold;
    }
    
    /* Target cells (td) - Black bg, Pink Text, Blue Border */
    [data-testid="stTable"] td {
        background-color: #000000 !important;
        color: #DA10AB !important;            /* Pink text */
        border: 1px solid #06AFE6 !important; /* Blue border */
    }

    </style>
    """,
    unsafe_allow_html=True,
)


st.title("Prosci Impact Index – Impact Assessment App")

# --- SIDEBAR: RESET CONTROL ---
with st.sidebar:
    st.header("App Controls")
    st.write("Need to start over?")
    if st.button("Reset / Clear All Data"):
        # Clear all session state keys to reset inputs
        for key in st.session_state.keys():
            del st.session_state[key]
        st.rerun()

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
    styled_group_df = style_impact_table(group_df)
    # Using st.table instead of st.dataframe for perfect color control
    st.table(styled_group_df)

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

    excel_filename = (
    f"{project_name.strip().replace(' ', '_')}_impact_results.xlsx"
    if project_name else
    "impact_results.xlsx"
)

st.download_button(
    label="Download impact results as Excel",
    data=excel_buffer,
    file_name=excel_filename,
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
        top_display = top_groups[["Group name", "Employees", "Degree of impact (0-5)"]]
        styled_top = style_impact_table(top_display)
        # Using st.table instead of st.dataframe for perfect color control
        st.table(styled_top)
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

summary_filename = (
    f"{project_name.strip().replace(' ', '_')}_impact_summary.pdf"
    if project_name else
    "impact_summary.pdf"
)

st.download_button(
    label="Download PDF Summary",
    data=pdf_bytes,
    file_name=summary_filename,
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

    # Dynamic filename using project name
    plan_filename = (
        f"{project_name.strip().replace(' ', '_')}_change_plan.pdf"
        if project_name else
        "change_plan.pdf"
    )

    st.download_button(
        label="Download Change Plan as PDF",
        data=plan_pdf_bytes,
        file_name=plan_filename,
        mime="application/pdf",
    )