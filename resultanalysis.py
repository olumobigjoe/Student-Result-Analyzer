"""
Student Results Analyzer (Streamlit only, no AI/Ollama)
--------------------------------------------------------
Upload student results (CSV/Excel) with CA, Practical, Exam and Total scores
(practical optional), view a detailed statistical performance analysis, and
download a full PDF report.

Run with:
    pip install streamlit pandas numpy openpyxl matplotlib reportlab
    streamlit run student_results_app.py
"""

import io
import datetime

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)

st.set_page_config(page_title="Student Results Analyzer", layout="wide", page_icon="📊")

# =============================================================================
# Column auto-detection helpers
# =============================================================================

NAME_HINTS = ["name", "student", "matric", "reg no", "regno", "id number", "student id"]
CA_HINTS = ["ca", "continuous assessment", "c.a", "test score", "coursework"]
PRACTICAL_HINTS = ["practical", "prac", "lab score", "lab"]
EXAM_HINTS = ["exam", "examination", "final exam"]
TOTAL_HINTS = ["total", "grand total", "overall"]
COURSE_HINTS = ["course", "subject", "class"]

DEFAULT_GRADE_BOUNDARIES = [
    ("A", 70, 100),
    ("B", 60, 69.999),
    ("C", 50, 59.999),
    ("D", 45, 49.999),
    ("F", 0, 44.999),
]


def load_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file)
    else:
        raise ValueError("Unsupported file type. Please upload a .csv, .xlsx, or .xls file.")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _find_column(columns, hints):
    lower_map = {c: c.lower() for c in columns}
    for c, lc in lower_map.items():
        for h in hints:
            if h in lc:
                return c
    return None


def detect_columns(df: pd.DataFrame) -> dict:
    columns = list(df.columns)
    guess = {
        "name": _find_column(columns, NAME_HINTS),
        "ca": _find_column(columns, CA_HINTS),
        "practical": _find_column(columns, PRACTICAL_HINTS),
        "exam": _find_column(columns, EXAM_HINTS),
        "total": _find_column(columns, TOTAL_HINTS),
        "course": _find_column(columns, COURSE_HINTS),
    }
    if guess["name"] is None:
        used = {v for v in guess.values() if v}
        for c in columns:
            if c in used:
                continue
            if not pd.api.types.is_numeric_dtype(df[c]):
                guess["name"] = c
                break
    return guess


def coerce_numeric(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def compute_total_if_missing(df, ca_col, practical_col, exam_col, total_col):
    df = df.copy()
    if total_col and total_col in df.columns:
        return df, total_col
    component_cols = [c for c in [ca_col, practical_col, exam_col] if c]
    if not component_cols:
        return df, None
    df["Computed Total"] = df[component_cols].sum(axis=1, skipna=True)
    return df, "Computed Total"


# =============================================================================
# Analysis helpers
# =============================================================================

def assign_grade(score, boundaries=DEFAULT_GRADE_BOUNDARIES):
    if pd.isna(score):
        return "N/A"
    for label, low, high in boundaries:
        if low <= score <= high:
            return label
    return "N/A"


def component_stats(df: pd.DataFrame, score_cols: dict) -> pd.DataFrame:
    rows = []
    for label, col in score_cols.items():
        if col is None or col not in df.columns:
            continue
        series = df[col].dropna()
        if series.empty:
            continue
        rows.append({
            "Component": label,
            "Mean": round(series.mean(), 2),
            "Median": round(series.median(), 2),
            "Std Dev": round(series.std(), 2),
            "Min": round(series.min(), 2),
            "Max": round(series.max(), 2),
            "Count": int(series.count()),
        })
    return pd.DataFrame(rows)


def pass_fail_summary(df: pd.DataFrame, total_col: str, pass_mark: float) -> dict:
    valid = df[total_col].dropna()
    total_students = len(valid)
    passed = int((valid >= pass_mark).sum())
    failed = total_students - passed
    pass_rate = round((passed / total_students) * 100, 1) if total_students else 0.0
    return {"total_students": total_students, "passed": passed, "failed": failed, "pass_rate": pass_rate}


def grade_distribution(df: pd.DataFrame, total_col: str, boundaries=DEFAULT_GRADE_BOUNDARIES) -> pd.DataFrame:
    grades = df[total_col].apply(lambda x: assign_grade(x, boundaries))
    counts = grades.value_counts().reindex([b[0] for b in boundaries], fill_value=0)
    dist = counts.reset_index()
    dist.columns = ["Grade", "Number of Students"]
    return dist


def top_bottom_performers(df: pd.DataFrame, name_col: str, total_col: str, n: int = 5):
    ranked = df[[name_col, total_col]].dropna().sort_values(total_col, ascending=False)
    top = ranked.head(n).reset_index(drop=True)
    bottom = ranked.tail(n).sort_values(total_col).reset_index(drop=True)
    return top, bottom


def correlation_matrix(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    valid_cols = [c for c in cols if c and c in df.columns]
    if len(valid_cols) < 2:
        return pd.DataFrame()
    return df[valid_cols].corr().round(2)


def component_contribution(df: pd.DataFrame, score_cols: dict, total_col: str) -> pd.DataFrame:
    rows = []
    total_mean = df[total_col].mean() if total_col in df.columns else np.nan
    for label, col in score_cols.items():
        if col is None or col not in df.columns:
            continue
        comp_mean = df[col].mean()
        pct = round((comp_mean / total_mean) * 100, 1) if total_mean else np.nan
        rows.append({"Component": label, "Avg Score": round(comp_mean, 2), "% of Total": pct})
    return pd.DataFrame(rows)


def build_summary_text(course_label, stats_df, pass_summary, grade_df, component_df, top_df, bottom_df) -> str:
    """Plain-text, rule-based narrative (no AI) built from the computed stats."""
    lines = []
    lines.append(f"Class Performance Summary for {course_label}")
    lines.append(
        f"A total of {pass_summary['total_students']} students were assessed. "
        f"{pass_summary['passed']} passed and {pass_summary['failed']} failed, "
        f"giving a pass rate of {pass_summary['pass_rate']}%."
    )

    if not component_df.empty:
        best = component_df.loc[component_df["Avg Score"].idxmax()]
        worst = component_df.loc[component_df["Avg Score"].idxmin()]
        lines.append(
            f"Among the assessed components, '{best['Component']}' had the highest average "
            f"score ({best['Avg Score']}), while '{worst['Component']}' had the lowest "
            f"average score ({worst['Avg Score']})."
        )

    if not grade_df.empty:
        top_grade_row = grade_df.loc[grade_df["Number of Students"].idxmax()]
        lines.append(
            f"The most common grade was '{top_grade_row['Grade']}', achieved by "
            f"{top_grade_row['Number of Students']} student(s)."
        )
        fail_row = grade_df[grade_df["Grade"] == "F"]
        if not fail_row.empty and fail_row.iloc[0]["Number of Students"] > 0:
            lines.append(
                f"{fail_row.iloc[0]['Number of Students']} student(s) scored an F grade "
                f"and may need remedial support."
            )

    if not stats_df.empty:
        for _, row in stats_df.iterrows():
            lines.append(
                f"For {row['Component']}: average = {row['Mean']}, "
                f"median = {row['Median']}, std dev = {row['Std Dev']} "
                f"(range {row['Min']}-{row['Max']})."
            )

    if not top_df.empty:
        lines.append("Top performers: " + ", ".join(top_df.iloc[:, 0].astype(str).tolist()) + ".")
    if not bottom_df.empty:
        lines.append("Students who may need support: " + ", ".join(bottom_df.iloc[:, 0].astype(str).tolist()) + ".")

    return "\n".join(lines)


# =============================================================================
# PDF report generation
# =============================================================================

def _df_to_table(df, col_widths=None):
    data = [list(df.columns)] + df.astype(str).values.tolist()
    table = Table(data, colWidths=col_widths, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return table


def _fig_to_image(fig, width=15 * cm):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return Image(buf, width=width, height=width * 0.55)


def _grade_pie_chart(grade_df):
    fig, ax = plt.subplots(figsize=(5, 4))
    non_zero = grade_df[grade_df["Number of Students"] > 0]
    if non_zero.empty:
        ax.text(0.5, 0.5, "No data", ha="center")
    else:
        ax.pie(non_zero["Number of Students"], labels=non_zero["Grade"], autopct="%1.1f%%", startangle=90)
        ax.set_title("Grade Distribution")
    return fig


def _score_hist_chart(df, total_col):
    fig, ax = plt.subplots(figsize=(6, 3.5))
    df[total_col].dropna().hist(bins=15, ax=ax, color="#2980b9", edgecolor="white")
    ax.set_xlabel("Total Score")
    ax.set_ylabel("Number of Students")
    ax.set_title("Distribution of Total Scores")
    ax.grid(axis="y", alpha=0.3)
    return fig


def _component_bar_chart(component_df):
    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.bar(component_df["Component"], component_df["Avg Score"], color="#27ae60")
    ax.set_ylabel("Average Score")
    ax.set_title("Average Score by Component")
    ax.grid(axis="y", alpha=0.3)
    return fig


def generate_pdf_report(course_name, stats_df, pass_summary, grade_df, top_df, bottom_df,
                         component_df, narrative_text, full_results_df=None, total_col=None) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm, leftMargin=1.5 * cm, rightMargin=1.5 * cm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SectionHeader", fontSize=14, spaceAfter=8, spaceBefore=14,
                               textColor=colors.HexColor("#2c3e50"), fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="Body", fontSize=10, leading=14))

    story = []
    story.append(Paragraph("Student Performance Report", styles["Title"]))
    story.append(Paragraph(f"Course: {course_name}", styles["Heading2"]))
    story.append(Paragraph(f"Generated: {datetime.datetime.now().strftime('%d %B %Y, %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 0.6 * cm))

    story.append(Paragraph("Class Summary", styles["SectionHeader"]))
    summary_line = (
        f"Total students: {pass_summary.get('total_students')} &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Passed: {pass_summary.get('passed')} &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Failed: {pass_summary.get('failed')} &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Pass rate: {pass_summary.get('pass_rate')}%"
    )
    story.append(Paragraph(summary_line, styles["Body"]))
    story.append(Spacer(1, 0.4 * cm))

    if stats_df is not None and not stats_df.empty:
        story.append(Paragraph("Component Statistics", styles["SectionHeader"]))
        story.append(_df_to_table(stats_df))
        story.append(Spacer(1, 0.4 * cm))

    if component_df is not None and not component_df.empty:
        story.append(_fig_to_image(_component_bar_chart(component_df)))
        story.append(Spacer(1, 0.4 * cm))

    if grade_df is not None and not grade_df.empty:
        story.append(Paragraph("Grade Distribution", styles["SectionHeader"]))
        story.append(_df_to_table(grade_df))
        story.append(Spacer(1, 0.3 * cm))
        story.append(_fig_to_image(_grade_pie_chart(grade_df), width=10 * cm))
        story.append(Spacer(1, 0.4 * cm))

    if total_col and full_results_df is not None and total_col in full_results_df.columns:
        story.append(_fig_to_image(_score_hist_chart(full_results_df, total_col)))
        story.append(Spacer(1, 0.4 * cm))

    if top_df is not None and not top_df.empty:
        story.append(Paragraph("Top Performers", styles["SectionHeader"]))
        story.append(_df_to_table(top_df))
        story.append(Spacer(1, 0.4 * cm))

    if bottom_df is not None and not bottom_df.empty:
        story.append(Paragraph("Students Who May Need Support", styles["SectionHeader"]))
        story.append(_df_to_table(bottom_df))
        story.append(Spacer(1, 0.4 * cm))

    story.append(PageBreak())
    story.append(Paragraph("Detailed Analysis", styles["SectionHeader"]))
    for para in (narrative_text or "No analysis available.").split("\n"):
        if para.strip():
            story.append(Paragraph(para.strip(), styles["Body"]))
            story.append(Spacer(1, 0.15 * cm))

    if full_results_df is not None:
        story.append(PageBreak())
        story.append(Paragraph("Full Student Results", styles["SectionHeader"]))
        story.append(_df_to_table(full_results_df.fillna("")))

    doc.build(story)
    buffer.seek(0)
    return buffer.read()


# =============================================================================
# Streamlit UI
# =============================================================================

st.sidebar.title("⚙️ Settings")
pass_mark = st.sidebar.number_input("Pass mark (out of 100)", min_value=0, max_value=100, value=40)

st.sidebar.markdown("**Grade boundaries**")
grade_a = st.sidebar.number_input("A ≥", min_value=0, max_value=100, value=70)
grade_b = st.sidebar.number_input("B ≥", min_value=0, max_value=100, value=60)
grade_c = st.sidebar.number_input("C ≥", min_value=0, max_value=100, value=50)
grade_d = st.sidebar.number_input("D ≥", min_value=0, max_value=100, value=45)

boundaries = [
    ("A", grade_a, 100),
    ("B", grade_b, grade_a - 0.001),
    ("C", grade_c, grade_b - 0.001),
    ("D", grade_d, grade_c - 0.001),
    ("F", 0, grade_d - 0.001),
]

st.title("📊 Student Results Analyzer")
st.write(
    "Upload a CSV or Excel file with student results (CA, Practical, Exam, Total — "
    "or CA, Exam, Total if there's no practical component) to get a detailed "
    "performance analysis and a downloadable PDF report."
)

uploaded_file = st.file_uploader("Upload results file", type=["csv", "xlsx", "xls"])
if uploaded_file is None:
    st.info("👆 Upload a file to get started.")
    st.stop()

try:
    df_raw = load_file(uploaded_file)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

st.subheader("Preview of uploaded data")
st.dataframe(df_raw.head(10), use_container_width=True)

guess = detect_columns(df_raw)
columns = list(df_raw.columns)
none_option = "— None —"
options = [none_option] + columns


def col_index(guessed_col):
    return options.index(guessed_col) if guessed_col in options else 0


st.subheader("Confirm column mapping")
st.caption("Auto-detected where possible — adjust any that look wrong.")

c1, c2, c3 = st.columns(3)
with c1:
    name_col = st.selectbox("Student name / ID column", options, index=col_index(guess["name"]))
    course_col = st.selectbox("Course column (optional)", options, index=col_index(guess["course"]))
with c2:
    ca_col = st.selectbox("CA score column", options, index=col_index(guess["ca"]))
    practical_col = st.selectbox("Practical score column (optional)", options, index=col_index(guess["practical"]))
with c3:
    exam_col = st.selectbox("Exam score column", options, index=col_index(guess["exam"]))
    total_col_choice = st.selectbox("Total score column (optional — computed if none)", options, index=col_index(guess["total"]))

name_col = None if name_col == none_option else name_col
course_col = None if course_col == none_option else course_col
ca_col = None if ca_col == none_option else ca_col
practical_col = None if practical_col == none_option else practical_col
exam_col = None if exam_col == none_option else exam_col
total_col_choice = None if total_col_choice == none_option else total_col_choice

if name_col is None or (ca_col is None and exam_col is None):
    st.warning("Please make sure at least a student name column and a CA or Exam column are selected.")
    st.stop()

numeric_cols = [c for c in [ca_col, practical_col, exam_col, total_col_choice] if c]
df = coerce_numeric(df_raw, numeric_cols)
df, total_col = compute_total_if_missing(df, ca_col, practical_col, exam_col, total_col_choice)

if total_col is None:
    st.error("Could not determine or compute a Total score. Please check your column selections.")
    st.stop()

selected_course = "All Students"
if course_col:
    courses = ["All Students"] + sorted(df[course_col].dropna().unique().tolist())
    selected_course = st.selectbox("Filter by course", courses)
    if selected_course != "All Students":
        df = df[df[course_col] == selected_course]

course_label = selected_course if selected_course != "All Students" else (course_col or "Uploaded Dataset")

score_cols = {}
if ca_col:
    score_cols["CA"] = ca_col
if practical_col:
    score_cols["Practical"] = practical_col
if exam_col:
    score_cols["Exam"] = exam_col

st.markdown("---")
st.header(f"📈 Analysis: {course_label}")

stats_df = component_stats(df, score_cols)
pass_summary = pass_fail_summary(df, total_col, pass_mark)
grade_df = grade_distribution(df, total_col, boundaries)
top_df, bottom_df = top_bottom_performers(df, name_col, total_col, n=5)
component_df = component_contribution(df, score_cols, total_col)
corr_df = correlation_matrix(df, list(score_cols.values()) + [total_col])

m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Students", pass_summary["total_students"])
m2.metric("Passed", pass_summary["passed"])
m3.metric("Failed", pass_summary["failed"])
m4.metric("Pass Rate", f"{pass_summary['pass_rate']}%")

st.subheader("Component Statistics")
st.dataframe(stats_df, use_container_width=True)

col_a, col_b = st.columns(2)
with col_a:
    st.subheader("Grade Distribution")
    st.dataframe(grade_df, use_container_width=True)
    fig, ax = plt.subplots()
    non_zero = grade_df[grade_df["Number of Students"] > 0]
    if not non_zero.empty:
        ax.pie(non_zero["Number of Students"], labels=non_zero["Grade"], autopct="%1.1f%%", startangle=90)
        ax.set_title("Grade Distribution")
    st.pyplot(fig)

with col_b:
    st.subheader("Total Score Distribution")
    fig2, ax2 = plt.subplots()
    df[total_col].dropna().hist(bins=15, ax=ax2, color="#2980b9", edgecolor="white")
    ax2.set_xlabel("Total Score")
    ax2.set_ylabel("Number of Students")
    st.pyplot(fig2)

if not component_df.empty:
    st.subheader("Average Score by Component")
    st.dataframe(component_df, use_container_width=True)
    fig3, ax3 = plt.subplots()
    ax3.bar(component_df["Component"], component_df["Avg Score"], color="#27ae60")
    ax3.set_ylabel("Average Score")
    st.pyplot(fig3)

if not corr_df.empty:
    st.subheader("Correlation Between Components")
    st.dataframe(corr_df, use_container_width=True)

col_c, col_d = st.columns(2)
with col_c:
    st.subheader("🏆 Top Performers")
    st.dataframe(top_df, use_container_width=True)
with col_d:
    st.subheader("⚠️ May Need Support")
    st.dataframe(bottom_df, use_container_width=True)

st.subheader("Full Results Table")
st.dataframe(df, use_container_width=True)

st.markdown("---")
st.header("📝 Detailed Analysis")
narrative_text = build_summary_text(course_label, stats_df, pass_summary, grade_df, component_df, top_df, bottom_df)
st.write(narrative_text)

st.markdown("---")
st.header("📄 Export PDF Report")

if st.button("Generate PDF Report"):
    with st.spinner("Building PDF..."):
        pdf_bytes = generate_pdf_report(
            course_name=course_label,
            stats_df=stats_df,
            pass_summary=pass_summary,
            grade_df=grade_df,
            top_df=top_df,
            bottom_df=bottom_df,
            component_df=component_df,
            narrative_text=narrative_text,
            full_results_df=df[[c for c in [name_col, course_col] + list(score_cols.values()) + [total_col] if c]],
            total_col=total_col,
        )
    st.success("PDF ready!")
    st.download_button(
        label="⬇️ Download PDF Report",
        data=pdf_bytes,
        file_name=f"{course_label.replace(' ', '_')}_performance_report.pdf",
        mime="application/pdf",
    )