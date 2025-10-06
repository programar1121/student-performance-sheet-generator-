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

# Initialize session state for comparison mode
if 'comparison_mode' not in st.session_state:
    st.session_state.comparison_mode = False
if 'df_week1' not in st.session_state:
    st.session_state.df_week1 = None
if 'df_week2' not in st.session_state:
    st.session_state.df_week2 = None
if 'comparison_data' not in st.session_state:
    st.session_state.comparison_data = None

# Sidebar options
st.sidebar.header("Report Settings")
academy_name = st.sidebar.text_input("Academy/School name", value=ACADEMY_NAME_DEFAULT)

# Comparison mode toggle
comparison_mode = st.sidebar.checkbox("📈 Enable Two-Week Comparison", value=st.session_state.comparison_mode)
if comparison_mode != st.session_state.comparison_mode:
    st.session_state.comparison_mode = comparison_mode
    st.rerun()

# Configurable thresholds
st.sidebar.subheader("Performance Thresholds")
strong_threshold = st.sidebar.number_input("Strong Performance Threshold (%)", 
                                         min_value=0.0, max_value=100.0, 
                                         value=STRONG_TH, step=1.0)
average_threshold = st.sidebar.number_input("Average Performance Threshold (%)", 
                                          min_value=0.0, max_value=100.0, 
                                          value=AVERAGE_LO, step=1.0)
top_count = st.sidebar.number_input("Number of Top Students to Show", 
                                  min_value=1, max_value=20, value=TOP_COUNT, step=1)
bottom_count = st.sidebar.number_input("Number of Bottom Students to Show", 
                                     min_value=1, max_value=20, value=BOTTOM_COUNT, step=1)

# Logo upload
show_logo = st.sidebar.checkbox("Upload logo to include in reports", value=False)
logo_file = None
if show_logo:
    logo_file = st.sidebar.file_uploader("Upload logo image (png/jpg)", type=["png", "jpg", "jpeg"])

report_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
st.sidebar.markdown(f"**Report date:** {report_date}")

# ---------- Enhanced Parsing functions ----------
def find_fraction(s: str):
    """Extract numerator and denominator from fraction strings"""
    if not isinstance(s, str):
        return None
    m = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", s)
    if m:
        try:
            numerator = float(m.group(1))
            denominator = float(m.group(2))
            if denominator == 0:
                return None
            return numerator, denominator
        except (ValueError, TypeError):
            return None
    return None

def find_number(s: str):
    """Extract numeric values from strings"""
    if not isinstance(s, str):
        return None
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if m:
        try:
            return float(m.group(1))
        except (ValueError, TypeError):
            return None
    return None

def parse_series_to_pct(series: pd.Series):
    """Convert various score formats to percentages with enhanced error handling"""
    s = series.copy()
    parsed = []
    orig = s.fillna("").astype(str)
    has_slash = orig.str.contains("/").any()

    if not has_slash:
        numeric = pd.to_numeric(s, errors="coerce")
        nonnull = numeric.dropna()
        
        if len(nonnull) == 0:
            for val in s:
                if pd.isna(val):
                    parsed.append(np.nan)
                    continue
                num = find_number(str(val))
                parsed.append(num if num is not None else np.nan)
            
            parsed = pd.Series(parsed, index=s.index)
            if parsed.dropna().empty:
                return parsed
            maxv = parsed.max()
            if maxv > 100:
                return (parsed / maxv * 100).clip(0, 100)
            return parsed.clip(0, 100)
        else:
            maxv = nonnull.max()
            if maxv <= 100:
                return numeric.clip(0, 100)
            else:
                return (numeric / maxv * 100).clip(0, 100)

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
            parsed.append((obt / tot) * 100)
            continue
            
        num = find_number(str(val))
        parsed.append(num if num is not None else np.nan)

    parsed = pd.Series(parsed, index=s.index).astype(float)
    nonnull = parsed.dropna()
    
    if nonnull.empty:
        return parsed
    
    maxv = nonnull.max()
    if maxv > 100:
        return (parsed / maxv * 100).clip(0, 100)
    return parsed.clip(0, 100)

def process_dataframe(df, subjects):
    """Process a dataframe and return processed data with statistics"""
    df_processed = df.copy()
    df_processed.columns = df_processed.columns.str.strip()
    
    # Apply parsing to subjects
    pct_cols = []
    for subj in subjects:
        pct_col = subj + "_pct"
        df_processed[pct_col] = parse_series_to_pct(df_processed[subj])
        pct_cols.append(pct_col)

    # Calculate overall percentage
    if "Current Percentage" in df_processed.columns:
        current_pct_numeric = pd.to_numeric(df_processed["Current Percentage"], errors="coerce")
        if current_pct_numeric.notna().sum() > len(df_processed) * 0.5:
            df_processed["Overall_pct"] = current_pct_numeric
        else:
            df_processed["Overall_pct"] = df_processed[pct_cols].mean(axis=1)
    else:
        df_processed["Overall_pct"] = df_processed[pct_cols].mean(axis=1)

    # Remove students with no valid percentage data
    initial_count = len(df_processed)
    df_processed = df_processed[df_processed["Overall_pct"].notna()]
    
    # Calculate statistics
    total_students = len(df_processed)
    highest_pct = df_processed["Overall_pct"].max(skipna=True)
    lowest_pct = df_processed["Overall_pct"].min(skipna=True)
    class_avg = df_processed["Overall_pct"].mean(skipna=True)
    class_median = df_processed["Overall_pct"].median(skipna=True)

    df_sorted = df_processed.sort_values("Overall_pct", ascending=False).reset_index(drop=True)
    top_students = df_sorted.head(top_count)
    
    # Enhanced lowest students logic
    nonzero_df = df_sorted[df_sorted["Overall_pct"] > 0].copy()
    if len(nonzero_df) >= bottom_count:
        bottom_students = nonzero_df.tail(bottom_count).sort_values("Overall_pct", ascending=True)
    else:
        bottom_students = nonzero_df.sort_values("Overall_pct", ascending=True)

    # Performance categories
    strong_students = df_processed[df_processed["Overall_pct"] >= strong_threshold]
    avg_students = df_processed[(df_processed["Overall_pct"] >= average_threshold) & 
                               (df_processed["Overall_pct"] < strong_threshold)]
    weak_students = df_processed[df_processed["Overall_pct"] < average_threshold]

    # Subject averages
    subject_stats = {}
    for subject in subjects:
        pct_col = subject + "_pct"
        if pct_col in df_processed.columns:
            subject_data = df_processed[pct_col].dropna()
            if len(subject_data) > 0:
                subject_stats[subject] = {
                    'average': subject_data.mean(),
                    'median': subject_data.median(),
                    'highest': subject_data.max(),
                    'lowest': subject_data.min(),
                    'std_dev': subject_data.std()
                }

    return {
        'df': df_processed,
        'total_students': total_students,
        'highest_pct': highest_pct,
        'lowest_pct': lowest_pct,
        'class_avg': class_avg,
        'class_median': class_median,
        'top_students': top_students,
        'bottom_students': bottom_students,
        'strong_students': strong_students,
        'avg_students': avg_students,
        'weak_students': weak_students,
        'subject_stats': subject_stats,
        'pct_cols': pct_cols,
        'subjects': subjects
    }

# ---------- Comparison Functions ----------
def compare_weeks(week1_data, week2_data, week1_name="Week 1", week2_name="Week 2"):
    """Compare two weeks of data and return comparison metrics"""
    
    comparison = {
        'class_avg_change': week2_data['class_avg'] - week1_data['class_avg'],
        'class_avg_change_pct': ((week2_data['class_avg'] - week1_data['class_avg']) / week1_data['class_avg'] * 100) if week1_data['class_avg'] > 0 else 0,
        'highest_change': week2_data['highest_pct'] - week1_data['highest_pct'],
        'lowest_change': week2_data['lowest_pct'] - week1_data['lowest_pct'],
        'strong_students_change': len(week2_data['strong_students']) - len(week1_data['strong_students']),
        'weak_students_change': len(week2_data['weak_students']) - len(week1_data['weak_students']),
        'subject_comparison': {}
    }
    
    # Compare subject averages
    for subject in week1_data['subjects']:
        if subject in week1_data['subject_stats'] and subject in week2_data['subject_stats']:
            week1_avg = week1_data['subject_stats'][subject]['average']
            week2_avg = week2_data['subject_stats'][subject]['average']
            change = week2_avg - week1_avg
            change_pct = (change / week1_avg * 100) if week1_avg > 0 else 0
            
            comparison['subject_comparison'][subject] = {
                'week1_avg': week1_avg,
                'week2_avg': week2_avg,
                'change': change,
                'change_pct': change_pct
            }
    
    # Student-level comparison (for students present in both weeks)
    common_students = []
    student_subject_comparison = {}
    
    if 'Student Name' in week1_data['df'].columns and 'Student Name' in week2_data['df'].columns:
        week1_students = set(week1_data['df']['Student Name'].dropna())
        week2_students = set(week2_data['df']['Student Name'].dropna())
        common_names = week1_students.intersection(week2_students)
        
        for name in common_names:
            week1_row = week1_data['df'][week1_data['df']['Student Name'] == name].iloc[0]
            week2_row = week2_data['df'][week2_data['df']['Student Name'] == name].iloc[0]
            
            week1_score = week1_row['Overall_pct']
            week2_score = week2_row['Overall_pct']
            change = week2_score - week1_score
            
            # Subject-wise comparison for this student
            subject_changes = {}
            for subject in week1_data['subjects']:
                week1_subj = week1_row.get(subject + '_pct', np.nan)
                week2_subj = week2_row.get(subject + '_pct', np.nan)
                if not pd.isna(week1_subj) and not pd.isna(week2_subj):
                    subject_changes[subject] = {
                        'week1_score': week1_subj,
                        'week2_score': week2_subj,
                        'change': week2_subj - week1_subj
                    }
            
            common_students.append({
                'name': name,
                'week1_score': week1_score,
                'week2_score': week2_score,
                'change': change,
                'change_pct': (change / week1_score * 100) if week1_score > 0 else 0,
                'subject_changes': subject_changes
            })
            
            student_subject_comparison[name] = subject_changes
    
    comparison['common_students'] = pd.DataFrame(common_students)
    comparison['common_students_count'] = len(common_students)
    comparison['student_subject_comparison'] = student_subject_comparison
    
    return comparison

# ---------- Visualization Functions ----------
def fig_to_bytes(fig, dpi=150):
    """Convert matplotlib figure to bytes"""
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=dpi)
    plt.close(fig)
    buf.seek(0)
    return buf

def create_subject_avg_bar_chart(subjects_list, df_local):
    """Create bar chart of subject averages"""
    avgs = [df_local[s + "_pct"].mean(skipna=True) for s in subjects_list]
    fig, ax = plt.subplots(figsize=(8, 4))
    bars = ax.bar(subjects_list, avgs, color='skyblue', alpha=0.7)
    ax.set_title("Subject-wise Average Performance (%)", fontsize=14, fontweight='bold')
    ax.set_ylabel("Average %", fontweight='bold')
    ax.set_ylim(0, 100)
    ax.grid(axis='y', alpha=0.3)
    
    for bar, avg in zip(bars, avgs):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 1,
                f'{avg:.1f}%', ha='center', va='bottom', fontweight='bold')
    
    ax.set_xticklabels(subjects_list, rotation=45, ha='right')
    fig.tight_layout()
    return fig_to_bytes(fig)

def create_overall_distribution(hist_series):
    """Create histogram of overall percentage distribution"""
    fig, ax = plt.subplots(figsize=(8, 4))
    n, bins, patches = ax.hist(hist_series.dropna(), bins=12, color='lightgreen', 
                              alpha=0.7, edgecolor='black')
    ax.set_title("Overall Percentage Distribution", fontsize=14, fontweight='bold')
    ax.set_xlabel("Percentage", fontweight='bold')
    ax.set_ylabel("Number of Students", fontweight='bold')
    ax.grid(axis='y', alpha=0.3)
    
    for i, (count, patch) in enumerate(zip(n, patches)):
        if count > 0:
            ax.text(patch.get_x() + patch.get_width()/2, count + 0.1,
                   f'{int(count)}', ha='center', va='bottom', fontweight='bold')
    
    fig.tight_layout()
    return fig_to_bytes(fig)

def create_performance_pie_chart(strong_count, avg_count, weak_count):
    """Create pie chart of performance categories"""
    fig, ax = plt.subplots(figsize=(6, 6))
    sizes = [strong_count, avg_count, weak_count]
    labels = [f'Strong\n(≥{strong_threshold}%)', 
             f'Average\n({average_threshold}-{strong_threshold}%)', 
             f'Weak\n(<{average_threshold}%)']
    colors = ['#2ecc71', '#f39c12', '#e74c3c']
    
    wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%',
                                     startangle=90)
    
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
    
    ax.set_title('Performance Category Distribution', fontsize=14, fontweight='bold')
    return fig_to_bytes(fig)

def create_comparison_chart(week1_data, week2_data, week1_name="Week 1", week2_name="Week 2"):
    """Create comparison chart between two weeks"""
    metrics = ['Class Average', 'Highest Score', 'Lowest Score']
    week1_values = [week1_data['class_avg'], week1_data['highest_pct'], week1_data['lowest_pct']]
    week2_values = [week2_data['class_avg'], week2_data['highest_pct'], week2_data['lowest_pct']]
    
    x = np.arange(len(metrics))
    width = 0.35
    
    fig, ax = plt.subplots(figsize=(10, 6))
    bars1 = ax.bar(x - width/2, week1_values, width, label=week1_name, color='#3498db', alpha=0.7)
    bars2 = ax.bar(x + width/2, week2_values, width, label=week2_name, color='#2ecc71', alpha=0.7)
    
    ax.set_ylabel('Percentage', fontweight='bold')
    ax.set_title('Week-over-Week Comparison', fontsize=14, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(metrics)
    ax.legend()
    ax.grid(axis='y', alpha=0.3)
    
    # Add value labels on bars
    for bars in [bars1, bars2]:
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 1,
                   f'{height:.1f}%', ha='center', va='bottom', fontweight='bold')
    
    fig.tight_layout()
    return fig_to_bytes(fig)

def create_subject_comparison_chart(comparison_data, week1_name="Week 1", week2_name="Week 2"):
    """Create subject-wise comparison chart"""
    subjects = list(comparison_data.keys())
    week1_avgs = [comparison_data[subj]['week1_avg'] for subj in subjects]
    week2_avgs = [comparison_data[subj]['week2_avg'] for subj in subjects]
    changes = [comparison_data[subj]['change'] for subj in subjects]
    
    x = np.arange(len(subjects))
    width = 0.35
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
    
    # Bar chart for averages
    bars1 = ax1.bar(x - width/2, week1_avgs, width, label=week1_name, color='#3498db', alpha=0.7)
    bars2 = ax1.bar(x + width/2, week2_avgs, width, label=week2_name, color='#2ecc71', alpha=0.7)
    
    ax1.set_ylabel('Average %', fontweight='bold')
    ax1.set_title('Subject-wise Average Comparison', fontsize=14, fontweight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(subjects, rotation=45, ha='right')
    ax1.legend()
    ax1.grid(axis='y', alpha=0.3)
    
    # Add value labels
    for bars in [bars1, bars2]:
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 1,
                    f'{height:.1f}%', ha='center', va='bottom', fontsize=8, fontweight='bold')
    
    # Bar chart for changes
    colors = ['#2ecc71' if x >= 0 else '#e74c3c' for x in changes]
    bars3 = ax2.bar(x, changes, color=colors, alpha=0.7)
    ax2.set_ylabel('Change (%)', fontweight='bold')
    ax2.set_title('Subject-wise Performance Change', fontsize=14, fontweight='bold')
    ax2.set_xticks(x)
    ax2.set_xticklabels(subjects, rotation=45, ha='right')
    ax2.grid(axis='y', alpha=0.3)
    
    # Add change labels
    for bar, change in zip(bars3, changes):
        height = bar.get_height()
        va = 'bottom' if height >= 0 else 'top'
        color = 'green' if height >= 0 else 'red'
        ax2.text(bar.get_x() + bar.get_width()/2., height + (0.5 if height >= 0 else -0.5),
                f'{change:+.1f}%', ha='center', va=va, color=color, fontweight='bold')
    
    ax2.axhline(y=0, color='black', linestyle='-', alpha=0.3)
    
    fig.tight_layout()
    return fig_to_bytes(fig)

def create_student_comparison_chart(student_data, week1_name="Week 1", week2_name="Week 2"):
    """Create individual student subject comparison chart"""
    subjects = list(student_data.keys())
    week1_scores = [student_data[subj]['week1_score'] for subj in subjects]
    week2_scores = [student_data[subj]['week2_score'] for subj in subjects]
    changes = [student_data[subj]['change'] for subj in subjects]
    
    x = np.arange(len(subjects))
    width = 0.35
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8))
    
    # Bar chart for scores
    bars1 = ax1.bar(x - width/2, week1_scores, width, label=week1_name, color='#3498db', alpha=0.7)
    bars2 = ax1.bar(x + width/2, week2_scores, width, label=week2_name, color='#2ecc71', alpha=0.7)
    
    ax1.set_ylabel('Score (%)', fontweight='bold')
    ax1.set_title('Student Subject Performance Comparison', fontsize=14, fontweight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(subjects, rotation=45, ha='right')
    ax1.legend()
    ax1.grid(axis='y', alpha=0.3)
    ax1.set_ylim(0, 100)
    
    # Add value labels
    for bars in [bars1, bars2]:
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 1,
                    f'{height:.1f}%', ha='center', va='bottom', fontsize=8, fontweight='bold')
    
    # Bar chart for changes
    colors = ['#2ecc71' if x >= 0 else '#e74c3c' for x in changes]
    bars3 = ax2.bar(x, changes, color=colors, alpha=0.7)
    ax2.set_ylabel('Change (%)', fontweight='bold')
    ax2.set_title('Subject-wise Performance Change', fontsize=14, fontweight='bold')
    ax2.set_xticks(x)
    ax2.set_xticklabels(subjects, rotation=45, ha='right')
    ax2.grid(axis='y', alpha=0.3)
    
    # Add change labels
    for bar, change in zip(bars3, changes):
        height = bar.get_height()
        va = 'bottom' if height >= 0 else 'top'
        color = 'green' if height >= 0 else 'red'
        ax2.text(bar.get_x() + bar.get_width()/2., height + (0.5 if height >= 0 else -0.5),
                f'{change:+.1f}%', ha='center', va=va, color=color, fontweight='bold')
    
    ax2.axhline(y=0, color='black', linestyle='-', alpha=0.3)
    
    fig.tight_layout()
    return fig_to_bytes(fig)

# ---------- PDF Generation for Comparison ----------
def generate_comparison_pdf_buffer(week1_data, week2_data, comparison_data, week1_name, week2_name, logo_bytes=None):
    """Generate PDF report buffer for comparison"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = []

    # Header with logo
    header_data = []
    if logo_bytes:
        try:
            img = utils.ImageReader(logo_bytes)
            iw, ih = img.getSize()
            aspect = ih / float(iw)
            rl_img = RLImage(logo_bytes, width=80, height=(80 * aspect))
            header_data.append([rl_img, Paragraph(f"<b>{academy_name}</b><br/><i>Two-Week Performance Comparison</i><br/>{report_date}", styles["Normal"])])
            t = Table(header_data, colWidths=[90, 420])
            t.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
            story.append(t)
        except Exception as e:
            story.append(Paragraph(f"<b>{academy_name}</b>", styles["Title"]))
            story.append(Paragraph("Two-Week Performance Comparison", styles["Heading2"]))
    else:
        story.append(Paragraph(f"<b>{academy_name}</b>", styles["Title"]))
        story.append(Paragraph("Two-Week Performance Comparison", styles["Heading2"]))
    
    story.append(Paragraph(f"Report generated: {report_date}", styles["Normal"]))
    story.append(Spacer(1, 12))

    # Comparison Periods
    story.append(Paragraph("<b>Comparison Periods</b>", styles["Heading3"]))
    period_data = [
        ["Period", "Name", "Total Students", "Class Average"],
        [week1_name, week1_name, str(week1_data['total_students']), f"{week1_data['class_avg']:.2f}%"],
        [week2_name, week2_name, str(week2_data['total_students']), f"{week2_data['class_avg']:.2f}%"]
    ]
    period_table = Table(period_data, colWidths=[100, 150, 100, 100])
    period_table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold")
    ]))
    story.append(period_table)
    story.append(Spacer(1,12))

    # Key Metrics Comparison
    story.append(Paragraph("<b>Key Metrics Comparison</b>", styles["Heading2"]))
    metrics_data = [
        ["Metric", f"{week1_name}", f"{week2_name}", "Change", "Change %"],
        ["Class Average", f"{week1_data['class_avg']:.2f}%", f"{week2_data['class_avg']:.2f}%", 
         f"{comparison_data['class_avg_change']:+.2f}%", f"{comparison_data['class_avg_change_pct']:+.1f}%"],
        ["Highest Score", f"{week1_data['highest_pct']:.2f}%", f"{week2_data['highest_pct']:.2f}%", 
         f"{comparison_data['highest_change']:+.2f}%", "-"],
        ["Lowest Score", f"{week1_data['lowest_pct']:.2f}%", f"{week2_data['lowest_pct']:.2f}%", 
         f"{comparison_data['lowest_change']:+.2f}%", "-"],
        ["Strong Students", str(len(week1_data['strong_students'])), str(len(week2_data['strong_students'])), 
         f"{comparison_data['strong_students_change']:+d}", "-"],
        ["Weak Students", str(len(week1_data['weak_students'])), str(len(week2_data['weak_students'])), 
         f"{comparison_data['weak_students_change']:+d}", "-"]
    ]
    metrics_table = Table(metrics_data, colWidths=[120, 80, 80, 80, 80])
    metrics_table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.3, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BACKGROUND", (3,1), (3,-1), colors.lightgreen if comparison_data['class_avg_change'] >= 0 else colors.pink),
        ("BACKGROUND", (4,1), (4,1), colors.lightgreen if comparison_data['class_avg_change'] >= 0 else colors.pink),
    ]))
    story.append(metrics_table)
    story.append(Spacer(1,12))

    # Subject-wise Comparison
    story.append(Paragraph("<b>Subject-wise Performance Comparison</b>", styles["Heading2"]))
    subject_data = [["Subject", f"{week1_name} Avg", f"{week2_name} Avg", "Change", "Change %"]]
    
    for subject, stats in comparison_data['subject_comparison'].items():
        change_color = colors.lightgreen if stats['change'] >= 0 else colors.pink
        subject_data.append([
            subject,
            f"{stats['week1_avg']:.2f}%",
            f"{stats['week2_avg']:.2f}%",
            f"{stats['change']:+.2f}%",
            f"{stats['change_pct']:+.1f}%"
        ])
    
    subject_table = Table(subject_data, colWidths=[120, 80, 80, 80, 80])
    subject_table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.3, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BACKGROUND", (3,1), (4,-1), colors.lightgreen),
    ]))
    
    # Apply color to negative changes
    for i in range(1, len(subject_data)):
        if comparison_data['subject_comparison'][subject_data[i][0]]['change'] < 0:
            subject_table.setStyle(TableStyle([
                ("BACKGROUND", (3,i), (4,i), colors.pink)
            ]))
    
    story.append(subject_table)
    story.append(PageBreak())

    # Individual Student Comparison
    if comparison_data['common_students_count'] > 0:
        story.append(Paragraph("<b>Individual Student Performance Comparison</b>", styles["Heading2"]))
        story.append(Paragraph(f"Tracking {comparison_data['common_students_count']} common students", styles["Normal"]))
        story.append(Spacer(1,12))

        # Top Improvers
        story.append(Paragraph("<b>Top 5 Improvers</b>", styles["Heading3"]))
        top_improvers = comparison_data['common_students'].nlargest(5, 'change')
        improver_data = [["Student", f"{week1_name}", f"{week2_name}", "Change", "Change %"]]
        
        for _, student in top_improvers.iterrows():
            improver_data.append([
                student['name'],
                f"{student['week1_score']:.2f}%",
                f"{student['week2_score']:.2f}%",
                f"{student['change']:+.2f}%",
                f"{student['change_pct']:+.1f}%"
            ])
        
        improver_table = Table(improver_data, colWidths=[200, 80, 80, 80, 80])
        improver_table.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.3, colors.black),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("BACKGROUND", (3,1), (4,-1), colors.lightgreen),
        ]))
        story.append(improver_table)
        story.append(Spacer(1,12))

        # Top Decliners
        story.append(Paragraph("<b>Top 5 Decliners</b>", styles["Heading3"]))
        top_decliners = comparison_data['common_students'].nsmallest(5, 'change')
        decliner_data = [["Student", f"{week1_name}", f"{week2_name}", "Change", "Change %"]]
        
        for _, student in top_decliners.iterrows():
            decliner_data.append([
                student['name'],
                f"{student['week1_score']:.2f}%",
                f"{student['week2_score']:.2f}%",
                f"{student['change']:+.2f}%",
                f"{student['change_pct']:+.1f}%"
            ])
        
        decliner_table = Table(decliner_data, colWidths=[200, 80, 80, 80, 80])
        decliner_table.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.3, colors.black),
            ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("BACKGROUND", (3,1), (4,-1), colors.pink),
        ]))
        story.append(decliner_table)
        story.append(PageBreak())

        # Detailed Student Subject Analysis
        story.append(Paragraph("<b>Detailed Student Subject Analysis</b>", styles["Heading2"]))
        
        for idx, (student_name, subject_data) in enumerate(comparison_data['student_subject_comparison'].items()):
            if idx > 0 and idx % 2 == 0:
                story.append(PageBreak())
                
            story.append(Paragraph(f"<b>{student_name}</b>", styles["Heading3"]))
            
            student_overall = comparison_data['common_students'][comparison_data['common_students']['name'] == student_name].iloc[0]
            story.append(Paragraph(f"Overall: {student_overall['week1_score']:.2f}% → {student_overall['week2_score']:.2f}% "
                                 f"(Change: {student_overall['change']:+.2f}%)", styles["Normal"]))
            
            # Subject performance table
            subject_rows = [["Subject", f"{week1_name}", f"{week2_name}", "Change"]]
            for subject, scores in subject_data.items():
                subject_rows.append([
                    subject,
                    f"{scores['week1_score']:.2f}%",
                    f"{scores['week2_score']:.2f}%",
                    f"{scores['change']:+.2f}%"
                ])
            
            student_table = Table(subject_rows, colWidths=[150, 80, 80, 80])
            student_table.setStyle(TableStyle([
                ("GRID", (0,0), (-1,-1), 0.3, colors.black),
                ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ]))
            
            # Color code changes
            for i in range(1, len(subject_rows)):
                change = subject_data[subject_rows[i][0]]['change']
                if change >= 0:
                    student_table.setStyle(TableStyle([
                        ("BACKGROUND", (3,i), (3,i), colors.lightgreen)
                    ]))
                else:
                    student_table.setStyle(TableStyle([
                        ("BACKGROUND", (3,i), (3,i), colors.pink)
                    ]))
            
            story.append(student_table)
            story.append(Spacer(1,12))

    # Summary Insights
    story.append(PageBreak())
    story.append(Paragraph("<b>Summary Insights</b>", styles["Heading2"]))
    
    insights = []
    if comparison_data['class_avg_change'] > 0:
        insights.append(f"✅ <b>Overall Improvement</b>: Class average increased by {comparison_data['class_avg_change']:.2f}% ({comparison_data['class_avg_change_pct']:.1f}%)")
    else:
        insights.append(f"⚠️ <b>Overall Decline</b>: Class average decreased by {abs(comparison_data['class_avg_change']):.2f}% ({abs(comparison_data['class_avg_change_pct']):.1f}%)")
    
    if comparison_data['strong_students_change'] > 0:
        insights.append(f"✅ <b>More Strong Performers</b>: Number of strong students increased by {comparison_data['strong_students_change']}")
    else:
        insights.append(f"⚠️ <b>Fewer Strong Performers</b>: Number of strong students decreased by {abs(comparison_data['strong_students_change'])}")
    
    # Find best and worst performing subjects
    if comparison_data['subject_comparison']:
        best_subject = max(comparison_data['subject_comparison'].items(), key=lambda x: x[1]['change'])
        worst_subject = min(comparison_data['subject_comparison'].items(), key=lambda x: x[1]['change'])
        
        insights.append(f"🎯 <b>Best Improvement</b>: {best_subject[0]} improved by {best_subject[1]['change']:+.2f}%")
        insights.append(f"📉 <b>Largest Decline</b>: {worst_subject[0]} declined by {abs(worst_subject[1]['change']):.2f}%")
    
    for insight in insights:
        story.append(Paragraph(insight, styles["Normal"]))
        story.append(Spacer(1,6))

    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------- Main UI ----------
if not st.session_state.comparison_mode:
    # Single file analysis mode (existing code)
    st.header("📊 Single File Analysis")
    
    uploaded = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"], key="single_file")
    if not uploaded:
        st.info("Upload an Excel file with student records. Example columns: Student Name, Class Name, Chemistry, Biology, Math, English, Urdu, etc.")
        st.stop()

    # ... [Rest of single file analysis code] ...

else:
    # Two-file comparison mode
    st.header("📈 Two-Week Comparison Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📅 Week 1 Data")
        week1_file = st.file_uploader("Upload Week 1 Excel file", type=["xlsx"], key="week1")
        week1_name = st.text_input("Week 1 Name", value="Week 1")
        
    with col2:
        st.subheader("📅 Week 2 Data")
        week2_file = st.file_uploader("Upload Week 2 Excel file", type=["xlsx"], key="week2")
        week2_name = st.text_input("Week 2 Name", value="Week 2")
    
    if week1_file and week2_file:
        try:
            # Load and process both files
            with st.spinner("Loading and processing both files..."):
                df_week1 = pd.read_excel(week1_file)
                df_week2 = pd.read_excel(week2_file)
                
                # Detect subjects for both files
                ignore_cols = IGNORE_COLS_DEFAULT.copy()
                subjects_week1 = [c for c in df_week1.columns if c not in ignore_cols and c.lower() not in [ic.lower() for ic in ignore_cols]]
                subjects_week2 = [c for c in df_week2.columns if c not in ignore_cols and c.lower() not in [ic.lower() for ic in ignore_cols]]
                
                # Use common subjects
                common_subjects = list(set(subjects_week1) & set(subjects_week2))
                if not common_subjects:
                    st.error("❌ No common subjects found between the two files.")
                    st.stop()
                
                st.success(f"✅ Found {len(common_subjects)} common subjects: {common_subjects}")
                
                # Process both datasets
                week1_data = process_dataframe(df_week1, common_subjects)
                week2_data = process_dataframe(df_week2, common_subjects)
                
                # Perform comparison
                comparison = compare_weeks(week1_data, week2_data, week1_name, week2_name)
                st.session_state.comparison_data = comparison
                
            # Display comparison results
            st.markdown("---")
            st.header("📊 Comparison Results")
            
            # Key metrics comparison
            st.subheader("📈 Key Metrics Comparison")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                delta_class_avg = comparison['class_avg_change']
                st.metric(
                    "Class Average",
                    f"{week2_data['class_avg']:.2f}%",
                    f"{delta_class_avg:+.2f}%",
                    delta_color="normal" if delta_class_avg >= 0 else "inverse"
                )
            
            with col2:
                delta_highest = comparison['highest_change']
                st.metric(
                    "Highest Score",
                    f"{week2_data['highest_pct']:.2f}%",
                    f"{delta_highest:+.2f}%",
                    delta_color="normal" if delta_highest >= 0 else "inverse"
                )
            
            with col3:
                delta_strong = comparison['strong_students_change']
                st.metric(
                    "Strong Students",
                    len(week2_data['strong_students']),
                    f"{delta_strong:+d}",
                    delta_color="normal" if delta_strong >= 0 else "inverse"
                )
            
            with col4:
                delta_weak = comparison['weak_students_change']
                st.metric(
                    "Weak Students",
                    len(week2_data['weak_students']),
                    f"{delta_weak:+d}",
                    delta_color="normal" if delta_weak <= 0 else "inverse"
                )
            
            # Comparison charts
            st.subheader("📊 Comparison Charts")
            
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Overall Metrics Comparison**")
                comp_chart = create_comparison_chart(week1_data, week2_data, week1_name, week2_name)
                st.image(comp_chart)
            
            with col2:
                st.write("**Performance Category Changes**")
                cat_data = {
                    'Strong': [len(week1_data['strong_students']), len(week2_data['strong_students'])],
                    'Average': [len(week1_data['avg_students']), len(week2_data['avg_students'])],
                    'Weak': [len(week1_data['weak_students']), len(week2_data['weak_students'])]
                }
                
                fig, ax = plt.subplots(figsize=(8, 6))
                x = np.arange(3)
                width = 0.35
                
                bars1 = ax.bar(x - width/2, [cat_data[cat][0] for cat in ['Strong', 'Average', 'Weak']], 
                              width, label=week1_name, color='#3498db', alpha=0.7)
                bars2 = ax.bar(x + width/2, [cat_data[cat][1] for cat in ['Strong', 'Average', 'Weak']], 
                              width, label=week2_name, color='#2ecc71', alpha=0.7)
                
                ax.set_ylabel('Number of Students')
                ax.set_title('Performance Category Changes')
                ax.set_xticks(x)
                ax.set_xticklabels(['Strong', 'Average', 'Weak'])
                ax.legend()
                ax.grid(axis='y', alpha=0.3)
                
                for bars in [bars1, bars2]:
                    for bar in bars:
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                               f'{int(height)}', ha='center', va='bottom', fontweight='bold')
                
                st.pyplot(fig)
            
            # Subject-wise comparison
            st.subheader("📚 Subject-wise Performance Changes")
            subject_comp_chart = create_subject_comparison_chart(comparison['subject_comparison'], week1_name, week2_name)
            st.image(subject_comp_chart)
            
            # Detailed subject comparison table
            st.subheader("📋 Subject-wise Detailed Comparison")
            subject_comparison_data = []
            for subject, stats in comparison['subject_comparison'].items():
                subject_comparison_data.append({
                    'Subject': subject,
                    f'{week1_name} Avg': f"{stats['week1_avg']:.2f}%",
                    f'{week2_name} Avg': f"{stats['week2_avg']:.2f}%",
                    'Change': f"{stats['change']:+.2f}%",
                    'Change %': f"{stats['change_pct']:+.1f}%"
                })
            
            subject_df = pd.DataFrame(subject_comparison_data)
            st.dataframe(subject_df, use_container_width=True)
            
            # Individual Student Comparison Section
            if comparison['common_students_count'] > 0:
                st.markdown("---")
                st.header("👥 Individual Student Comparison")
                
                # Student selector
                student_names = comparison['common_students']['name'].tolist()
                selected_student = st.selectbox("Select Student for Detailed Analysis", student_names)
                
                if selected_student:
                    student_data = comparison['student_subject_comparison'][selected_student]
                    student_overall = comparison['common_students'][comparison['common_students']['name'] == selected_student].iloc[0]
                    
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        st.subheader(f"📊 {selected_student} - Subject Performance")
                        student_chart = create_student_comparison_chart(student_data, week1_name, week2_name)
                        st.image(student_chart)
                    
                    with col2:
                        st.subheader("Overall Performance")
                        st.metric(
                            "Overall Score",
                            f"{student_overall['week2_score']:.2f}%",
                            f"{student_overall['change']:+.2f}%",
                            delta_color="normal" if student_overall['change'] >= 0 else "inverse"
                        )
                        
                        st.write("**Subject Changes:**")
                        for subject, scores in student_data.items():
                            change_icon = "📈" if scores['change'] >= 0 else "📉"
                            st.write(f"{change_icon} {subject}: {scores['change']:+.1f}%")
                
                # Top and Bottom Improvers
                st.subheader("🏆 Student Progress Leaders")
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Top 5 Improvers**")
                    top_improvers = comparison['common_students'].nlargest(5, 'change')
                    display_top = top_improvers[['name', 'week1_score', 'week2_score', 'change']].copy()
                    display_top.columns = ['Student', f'{week1_name}', f'{week2_name}', 'Change']
                    display_top[f'{week1_name}'] = display_top[f'{week1_name}'].round(2)
                    display_top[f'{week2_name}'] = display_top[f'{week2_name}'].round(2)
                    display_top['Change'] = display_top['Change'].round(2)
                    st.dataframe(display_top, use_container_width=True)
                
                with col2:
                    st.write("**Top 5 Decliners**")
                    bottom_improvers = comparison['common_students'].nsmallest(5, 'change')
                    display_bottom = bottom_improvers[['name', 'week1_score', 'week2_score', 'change']].copy()
                    display_bottom.columns = ['Student', f'{week1_name}', f'{week2_name}', 'Change']
                    display_bottom[f'{week1_name}'] = display_bottom[f'{week1_name}'].round(2)
                    display_bottom[f'{week2_name}'] = display_bottom[f'{week2_name}'].round(2)
                    display_bottom['Change'] = display_bottom['Change'].round(2)
                    st.dataframe(display_bottom, use_container_width=True)
                
                # Student progress distribution
                st.write("**Student Progress Distribution**")
                fig, ax = plt.subplots(figsize=(10, 6))
                changes = comparison['common_students']['change']
                n, bins, patches = ax.hist(changes, bins=15, color='lightblue', alpha=0.7, edgecolor='black')
                ax.set_xlabel('Score Change (%)')
                ax.set_ylabel('Number of Students')
                ax.set_title('Distribution of Student Score Changes')
                ax.grid(axis='y', alpha=0.3)
                ax.axvline(x=0, color='red', linestyle='--', alpha=0.7, label='No Change')
                ax.legend()
                
                for i, (count, patch) in enumerate(zip(n, patches)):
                    if count > 0:
                        ax.text(patch.get_x() + patch.get_width()/2, count + 0.1,
                               f'{int(count)}', ha='center', va='bottom', fontweight='bold')
                
                st.pyplot(fig)
            
            # Summary insights
            st.markdown("---")
            st.subheader("💡 Summary Insights")
            
            insights = []
            if comparison['class_avg_change'] > 0:
                insights.append(f"✅ **Overall Improvement**: Class average increased by {comparison['class_avg_change']:.2f}% ({comparison['class_avg_change_pct']:.1f}%)")
            else:
                insights.append(f"⚠️ **Overall Decline**: Class average decreased by {abs(comparison['class_avg_change']):.2f}% ({abs(comparison['class_avg_change_pct']):.1f}%)")
            
            if comparison['strong_students_change'] > 0:
                insights.append(f"✅ **More Strong Performers**: Number of strong students increased by {comparison['strong_students_change']}")
            else:
                insights.append(f"⚠️ **Fewer Strong Performers**: Number of strong students decreased by {abs(comparison['strong_students_change'])}")
            
            # Find best and worst performing subjects
            if comparison['subject_comparison']:
                best_subject = max(comparison['subject_comparison'].items(), key=lambda x: x[1]['change'])
                worst_subject = min(comparison['subject_comparison'].items(), key=lambda x: x[1]['change'])
                
                insights.append(f"🎯 **Best Improvement**: {best_subject[0]} improved by {best_subject[1]['change']:+.2f}%")
                insights.append(f"📉 **Largest Decline**: {worst_subject[0]} declined by {abs(worst_subject[1]['change']):.2f}%")
            
            for insight in insights:
                st.write(insight)
            
            # PDF Download Button
            st.markdown("---")
            st.subheader("📄 Download Comparison Report")
            
            if st.button("📥 Generate & Download Comparison PDF Report", type="primary"):
                logo_bytes = None
                if logo_file is not None:
                    logo_bytes = logo_file.getvalue() if hasattr(logo_file, 'getvalue') else logo_file
                
                with st.spinner("Generating comprehensive PDF report..."):
                    try:
                        pdf_buffer = generate_comparison_pdf_buffer(
                            week1_data, week2_data, comparison, week1_name, week2_name, logo_bytes
                        )
                        
                        # Create safe filename
                        class_name = week1_data['df'].get('Class Name', 'Unknown Class').iloc[0] if 'Class Name' in week1_data['df'].columns else 'Unknown_Class'
                        class_name_safe = re.sub(r'[^\w\-_]', '', str(class_name).replace(" ", "_"))
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"comparison_report_{class_name_safe}_{timestamp}.pdf"
                        
                        st.download_button(
                            label="📥 Download PDF Report",
                            data=pdf_buffer,
                            file_name=filename,
                            mime="application/pdf",
                            type="primary"
                        )
                        
                        st.success("✅ PDF report generated successfully! Click the download button above.")
                        
                    except Exception as e:
                        st.error(f"❌ Error generating PDF report: {e}")
                
        except Exception as e:
            st.error(f"❌ Error processing comparison: {e}")
            st.stop()
    else:
        st.info("👆 Please upload both Week 1 and Week 2 Excel files to begin comparison analysis.")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>"
    "Student Performance Analyzer • Built with Streamlit • "
    f"Report generated on {report_date}"
    "</div>",
    unsafe_allow_html=True
)