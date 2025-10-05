# app.py
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ReportLab for PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import utils
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# python-docx for editable docx
from docx import Document
from docx.shared import Inches

# ---------- CONFIG ----------
ACADEMY_NAME_DEFAULT = "Avenir Academy"
STRONG_TH = 70.0
AVERAGE_LO = 40.0
TOP_COUNT = 3
BOTTOM_COUNT = 3

IGNORE_COLS_DEFAULT = [
    "Student Id Fk", "Roll Number", "Section Roll Number",
    "Student Name", "Father Name", "Class Name", "Section Name",
    "Evaluation", "Total Marks", "Obtained Marks",
    "Current Percentage", "Current Position", "Current Grade"
]
# ----------------------------

st.set_page_config(page_title="Student Performance Analyzer", page_icon="📊", layout="wide")
st.title("📊 Student Performance Analyzer")

# Sidebar options
st.sidebar.header("Report Settings")
academy_name = st.sidebar.text_input("Academy/School name", value=ACADEMY_NAME_DEFAULT)
show_logo = st.sidebar.checkbox("Upload logo to include in reports", value=False)
logo_file = None
if show_logo:
    logo_file = st.sidebar.file_uploader("Upload logo image (png/jpg)", type=["png", "jpg", "jpeg"])

report_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
st.sidebar.markdown(f"**Report date:** {report_date}")

# Upload Excel
uploaded = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Upload an Excel file with student records. Example columns: Student Name, Class Name, Chemistry, Biology, Math, English, Urdu, etc.")
    st.stop()

# Load dataframe
try:
    df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"Could not read Excel file: {e}")
    st.stop()

df.columns = df.columns.str.strip()
st.success("File uploaded and read successfully!")
st.write("Detected columns:", df.columns.tolist())

# Detect subjects (exclude ignore list and non-subject columns)
ignore_cols = IGNORE_COLS_DEFAULT.copy()
subjects = [c for c in df.columns if c not in ignore_cols and c.lower() not in [ic.lower() for ic in ignore_cols]]

if len(subjects) == 0:
    st.error("No subject columns detected. Please check your Excel file column names or update the ignore list in the code.")
    st.stop()

st.write("Detected subjects:", subjects)

# ---------- Parsing functions ----------
def find_fraction(s: str):
    if not isinstance(s, str):
        return None
    m = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", s)
    if m:
        try:
            return float(m.group(1)), float(m.group(2))
        except:
            return None
    return None

def find_number(s: str):
    if not isinstance(s, str):
        return None
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if m:
        try:
            return float(m.group(1))
        except:
            return None
    return None

def parse_series_to_pct(series: pd.Series):
    s = series.copy()
    parsed = []
    orig = s.fillna("").astype(str)
    has_slash = orig.str.contains("/").any()

    # If no slashes, attempt numeric coercion
    if not has_slash:
        numeric = pd.to_numeric(s, errors="coerce")
        nonnull = numeric.dropna()
        if len(nonnull) == 0:
            # fallback extract first number
            for val in s:
                if pd.isna(val):
                    parsed.append(np.nan)
                    continue
                num = find_number(str(val))
                parsed.append(num if num is not None else np.nan)
            parsed = pd.Series(parsed, index=s.index)
            # decide normalization
            if parsed.dropna().empty:
                return parsed  # all NaN
            maxv = parsed.max()
            if maxv > 100:
                # treat as raw marks, infer total = max
                return (parsed / maxv * 100).clip(0,100)
            return parsed.clip(0,100)
        else:
            maxv = nonnull.max()
            if maxv <= 100:
                return numeric.clip(0,100)
            else:
                # infer total
                return (numeric / maxv * 100).clip(0,100)

    # If has slashes, parse row-wise
    for val in s:
        if pd.isna(val):
            parsed.append(np.nan)
            continue
        if isinstance(val, (int, float, np.integer, np.floating)):
            parsed.append(float(val))
            continue
        fs = find_fraction(str(val))
        if fs:
            obt, tot = fs
            if tot == 0:
                parsed.append(np.nan)
            else:
                parsed.append((obt / tot) * 100)
            continue
        # fallback: first number
        num = find_number(str(val))
        parsed.append(num if num is not None else np.nan)

    parsed = pd.Series(parsed, index=s.index).astype(float)
    # For rows that came from fractions -> already percent; for others if values >100 infer total from those non-fraction values
    # We can't perfectly detect which rows came from fraction, but if most values are <=100 it's fine.
    nonnull = parsed.dropna()
    if nonnull.empty:
        return parsed
    maxv = nonnull.max()
    if maxv > 100:
        # infer total among numeric values >100
        # compute max among original numeric-looking non-fraction rows
        # simpler: if max >100, scale all values by max to get %
        return (parsed / maxv * 100).clip(0,100)
    return parsed.clip(0,100)

# Apply parsing to subject columns and create pct cols
pct_cols = []
for subj in subjects:
    pct_col = subj + "_pct"
    df[pct_col] = parse_series_to_pct(df[subj])
    pct_cols.append(pct_col)

# If "Current Percentage" exists and looks valid, prefer it for overall ranking; otherwise compute average across detected subject pct columns
if "Current Percentage" in df.columns and pd.to_numeric(df["Current Percentage"], errors="coerce").notna().sum() > 0:
    df["Overall_pct"] = pd.to_numeric(df["Current Percentage"], errors="coerce")
else:
    df["Overall_pct"] = df[pct_cols].mean(axis=1)

# Fill NaN overall with 0 to avoid sorting issues (but keep original NaNs if needed)
# df["Overall_pct"] = df["Overall_pct"].fillna(0)

# ---------- Class statistics ----------
class_name = df.get("Class Name", pd.Series([academy_name])).iloc[0] if "Class Name" in df.columns else academy_name
total_students = len(df)
highest_pct = df["Overall_pct"].max(skipna=True)
lowest_pct = df["Overall_pct"].min(skipna=True)
class_avg = df["Overall_pct"].mean(skipna=True)

# Sort by overall for top/low
df_sorted = df.sort_values("Overall_pct", ascending=False).reset_index(drop=True)

top_students = df_sorted.head(TOP_COUNT)
bottom_students = df_sorted.tail(BOTTOM_COUNT).sort_values("Overall_pct", ascending=True)

# Define average students = those not in top or bottom and whose overall pct within [AVERAGE_LO, STRONG_TH) maybe
# User asked: "in average all student that are average" -> we will treat average students as those with Overall_pct in [AVERAGE_LO, STRONG_TH)
avg_students = df[(df["Overall_pct"] >= AVERAGE_LO) & (df["Overall_pct"] < STRONG_TH)].sort_values("Overall_pct", ascending=False)

# For distribution graph, create histogram bins
overall_valid = df["Overall_pct"].dropna()

# ---------- Helper: matplotlib plotting to BytesIO ----------
def fig_to_bytes(fig, dpi=150):
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=dpi)
    plt.close(fig)
    buf.seek(0)
    return buf

def create_subject_avg_bar_chart(subjects_list, df_local):
    avgs = [df_local[s + "_pct"].mean(skipna=True) for s in subjects_list]
    fig, ax = plt.subplots(figsize=(6,3))
    ax.bar(subjects_list, avgs)  # do not specify colors
    ax.set_title("Subject-wise Average (%)")
    ax.set_ylabel("Average %")
    ax.set_ylim(0,100)
    ax.set_xticklabels(subjects_list, rotation=45, ha='right')
    fig.tight_layout()
    return fig_to_bytes(fig)

def create_overall_distribution(hist_series):
    fig, ax = plt.subplots(figsize=(6,3))
    ax.hist(hist_series.dropna(), bins=10)
    ax.set_title("Overall Percentage Distribution")
    ax.set_xlabel("Percentage")
    ax.set_ylabel("Number of Students")
    fig.tight_layout()
    return fig_to_bytes(fig)

def create_student_bar(student_row, subjects_list):
    vals = []
    labels = []
    for s in subjects_list:
        pct = student_row.get(s + "_pct", np.nan)
        if pd.isna(pct):
            vals.append(0.0)
        else:
            vals.append(pct)
        labels.append(s)
    fig, ax = plt.subplots(figsize=(4,1.5))
    ax.bar(labels, vals)
    ax.set_ylim(0,100)
    ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
    ax.set_yticks([0,50,100])
    fig.tight_layout()
    return fig_to_bytes(fig)

# ---------- Generate PDF ----------
def generate_pdf_buffer(df_local, subjects_list, logo_bytes=None):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = []

    # Header (logo + academy name + date)
    header_data = []
    if logo_bytes:
        # prepare ReportLab Image with fixed height
        img = utils.ImageReader(logo_bytes)
        iw, ih = img.getSize()
        aspect = ih / float(iw)
        rl_img = RLImage(logo_bytes, width=80, height=(80 * aspect))
        header_data.append([rl_img, Paragraph(f"<b>{academy_name}</b><br/><i>Class: {class_name}</i><br/>{report_date}", styles["Normal"])])
        t = Table(header_data, colWidths=[90, 420])
    else:
        story.append(Paragraph(f"<b>{academy_name}</b>", styles["Title"]))
        story.append(Paragraph(f"Class: {class_name}", styles["Normal"]))
        story.append(Paragraph(f"Report generated: {report_date}", styles["Normal"]))
        story.append(Spacer(1, 12))
        # skip table creation
        t = None

    if t:
        t.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
        story.append(t)
        story.append(Spacer(1,12))

    # Class summary
    story.append(Paragraph("<b>Class Summary</b>", styles["Heading2"]))
    summary_data = [
        ["Total Students", str(total_students)],
        ["Highest Percentage", f"{highest_pct:.2f}%"],
        ["Class Average (mean)", f"{class_avg:.2f}%"],
        ["Lowest Percentage", f"{lowest_pct:.2f}%"]
    ]
    sum_table = Table(summary_data, colWidths=[200, 300])
    sum_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.3,colors.black),("BACKGROUND",(0,0),(1,0),colors.lightgrey)]))
    story.append(sum_table)
    story.append(Spacer(1,12))

    # Top/Average/Lowest students
    story.append(Paragraph("<b>Top Students</b>", styles["Heading3"]))
    for _, r in top_students.iterrows():
        story.append(Paragraph(f"{r.get('Student Name','-')} — {r['Overall_pct']:.2f}%", styles["Normal"]))
    story.append(Spacer(1,6))

    story.append(Paragraph("<b>Average Students</b>", styles["Heading3"]))
    if avg_students.empty:
        story.append(Paragraph("No average students found (based on thresholds).", styles["Normal"]))
    else:
        # list all average students
        for _, r in avg_students.iterrows():
            story.append(Paragraph(f"{r.get('Student Name','-')} — {r['Overall_pct']:.2f}%", styles["Normal"]))
    story.append(Spacer(1,6))

    story.append(Paragraph("<b>Lowest Students</b>", styles["Heading3"]))
    for _, r in bottom_students.iterrows():
        story.append(Paragraph(f"{r.get('Student Name','-')} — {r['Overall_pct']:.2f}%", styles["Normal"]))
    story.append(Spacer(1,12))

    # Graphs
    story.append(Paragraph("<b>Graphs</b>", styles["Heading2"]))
    subj_chart = create_subject_avg_bar_chart(subjects_list, df_local)
    dist_chart = create_overall_distribution(df_local["Overall_pct"])
    # Insert both side by side table
    img1 = RLImage(subj_chart, width=260, height=120)
    img2 = RLImage(dist_chart, width=260, height=120)
    gtable = Table([[img1, img2]], colWidths=[260,260])
    story.append(gtable)
    story.append(PageBreak())

    # Individual student pages
    story.append(Paragraph("<b>Individual Student Performance</b>", styles["Heading2"]))
    story.append(Spacer(1,6))
    for idx, row in df_local.iterrows():
        story.append(Paragraph(f"<b>{row.get('Student Name','Student')}</b>", styles["Heading3"]))
        # subject-wise percentages with labels
        sub_rows = [["Subject", "Percentage", "Category"]]
        for s in subjects_list:
            pct = row.get(s + "_pct", np.nan)
            pct_str = "N/A" if pd.isna(pct) else f"{pct:.2f}%"
            if pd.isna(pct):
                cat = "N/A"
            elif pct >= STRONG_TH:
                cat = "Strongest"
            elif pct >= AVERAGE_LO:
                cat = "Average"
            else:
                cat = "Weakest"
            sub_rows.append([s, pct_str, cat])
        sub_table = Table(sub_rows, colWidths=[180,80,120])
        sub_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.3,colors.black),("BACKGROUND",(0,0),(-1,0),colors.lightgrey)]))
        story.append(sub_table)
        story.append(Spacer(1,6))
        # small per-student chart
        student_fig = create_student_bar(row, subjects_list)
        rlimg = RLImage(student_fig, width=300, height=80)
        story.append(rlimg)
        story.append(Spacer(1,12))
        # page break every 4 students
        if (idx + 1) % 4 == 0:
            story.append(PageBreak())

    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------- Generate DOCX ----------
def generate_docx_buffer(df_local, subjects_list, logo_bytes=None):
    buffer = BytesIO()
    docx = Document()
    # Header
    if logo_bytes:
        # Save logo bytes to temp and insert
        logo_stream = BytesIO(logo_bytes.read() if hasattr(logo_bytes, 'read') else logo_bytes)
        docx.add_picture(logo_stream, width=Inches(1))
    docx.add_heading(academy_name, level=1)
    docx.add_paragraph(f"Class: {class_name}")
    docx.add_paragraph(f"Report generated: {report_date}")
    docx.add_paragraph("")
    # Summary
    docx.add_heading("Class Summary", level=2)
    tbl = docx.add_table(rows=5, cols=2)
    tbl.style = 'Light List Accent 1' if 'Light List Accent 1' in docx.styles else tbl.style
    tbl.cell(0,0).text = "Total Students"
    tbl.cell(0,1).text = str(total_students)
    tbl.cell(1,0).text = "Highest Percentage"
    tbl.cell(1,1).text = f"{highest_pct:.2f}%"
    tbl.cell(2,0).text = "Class Average (mean)"
    tbl.cell(2,1).text = f"{class_avg:.2f}%"
    tbl.cell(3,0).text = "Lowest Percentage"
    tbl.cell(3,1).text = f"{lowest_pct:.2f}%"
    tbl.cell(4,0).text = "Top Students"
    tbl.cell(4,1).text = ", ".join([f"{r.get('Student Name','-')} ({r['Overall_pct']:.2f}%)" for _, r in top_students.iterrows()])

    docx.add_paragraph("")
    docx.add_heading("Graphs", level=2)
    # Insert graphs as images
    subj_chart = create_subject_avg_bar_chart(subjects_list, df_local)
    dist_chart = create_overall_distribution(df_local["Overall_pct"])
    docx.add_picture(subj_chart, width=Inches(4))
    docx.add_picture(dist_chart, width=Inches(4))
    docx.add_page_break()

    # Individual students
    docx.add_heading("Individual Student Performance", level=2)
    for _, r in df_local.iterrows():
        docx.add_heading(str(r.get("Student Name","Student")), level=3)
        for s in subjects_list:
            pct = r.get(s + "_pct", np.nan)
            pct_str = "N/A" if pd.isna(pct) else f"{pct:.2f}%"
            if pd.isna(pct):
                cat = "N/A"
            elif pct >= STRONG_TH:
                cat = "Strongest"
            elif pct >= AVERAGE_LO:
                cat = "Average"
            else:
                cat = "Weakest"
            docx.add_paragraph(f"{s}: {pct_str} — {cat}")
        # Add small chart image
        student_fig = create_student_bar(r, subjects_list)
        docx.add_picture(student_fig, width=Inches(4))
        docx.add_paragraph("")
    docx.save(buffer)
    buffer.seek(0)
    return buffer

# ---------- UI actions ----------
st.markdown("---")
st.subheader("Preview & Options")
st.write(f"Class: **{class_name}**  |  Total students: **{total_students}**  |  Generated: **{report_date}**")
st.write(f"Top {TOP_COUNT} students, Average students count: {len(avg_students)}, Lowest {BOTTOM_COUNT} students")

# Show small graphs in UI
col1, col2 = st.columns(2)
with col1:
    st.write("📘 Subject-wise Averages")
    subj_img = create_subject_avg_bar_chart(subjects, df)
    st.image(subj_img)
with col2:
    st.write("📊 Overall Performance Distribution")
    dist_img = create_overall_distribution(df["Overall_pct"])
    st.image(dist_img)

st.write("🏅 Top students preview:")
st.table(
    top_students[["Student Name", "Overall_pct"]]
    .rename(columns={"Overall_pct": "Overall %"})
    .head(TOP_COUNT)
)

# Generate buttons & output selection
output_type = st.radio("Choose output format:", ("PDF (with graphs)", "Editable Word (.docx)"))

if st.button("Generate & Download Report"):
    logo_bytes = None
    if logo_file is not None:
        logo_bytes = logo_file
    subjects_list = subjects

    # ✅ Detect Class and Group for filename
    class_name_safe = str(df.get("Class Name", pd.Series(["Unknown"])).iloc[0]).replace(" ", "_")
    group_name_safe = ""

    # Check if there's a Group column or detect from subject names
    if "Group" in df.columns:
        group_name_safe = str(df["Group"].iloc[0]).replace(" ", "_")
    elif any("Biology" in str(col) for col in df.columns):
        group_name_safe = "Bio"
    elif any("Computer" in str(col) for col in df.columns):
        group_name_safe = "Comp"

    # Build file base name
    if group_name_safe and group_name_safe.lower() != "nan":
        file_base = f"Student_Performance_{class_name_safe}_{group_name_safe}"
    else:
        file_base = f"Student_Performance_{class_name_safe}"

    with st.spinner("⏳ Generating report, please wait..."):
        if output_type.startswith("PDF"):
            pdf_buf = generate_pdf_buffer(df, subjects_list, logo_bytes=logo_bytes)
            st.success("✅ PDF generated successfully!")
            st.download_button(
                "📥 Download PDF Report",
                data=pdf_buf,
                file_name=f"{file_base}.pdf",
                mime="application/pdf"
            )
        else:
            docx_buf = generate_docx_buffer(df, subjects_list, logo_bytes=logo_bytes)
            st.success("✅ DOCX generated successfully!")
            st.download_button(
                "📥 Download DOCX Report",
                data=docx_buf,
                file_name=f"{file_base}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

st.markdown("---")
st.markdown("> 📎 Note: You can also use graphs and charts to make the document more helpful for analysis.")
