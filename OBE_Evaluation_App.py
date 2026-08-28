
import io
import re
from datetime import date
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="OBE Evaluation Report Generator",
    page_icon="🎓",
    layout="wide",
)

# -----------------------------
# Helpers
# -----------------------------
def clean_name(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def parse_clo_list(text: str) -> List[str]:
    items = []
    for line in (text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        m = re.match(r"^(CLO\s*\d+)\s*[:\-–]\s*(.*)$", line, re.I)
        if m:
            items.append(f"{m.group(1).upper()}: {m.group(2).strip()}")
        else:
            items.append(line)
    return items


def clo_ids(clo_text: str) -> List[str]:
    ids = []
    for i, item in enumerate(parse_clo_list(clo_text), 1):
        m = re.match(r"^(CLO\s*\d+)", item, re.I)
        ids.append(m.group(1).upper().replace(" ", "") if m else f"CLO{i}")
    return ids


def safe_float(x, default=0.0):
    try:
        return float(x)
    except Exception:
        return default


def read_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None, "No file uploaded."
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(uploaded_file), None
        if name.endswith(".xlsx"):
            return pd.read_excel(uploaded_file), None
        return None, "Please upload an .xlsx or .csv file."
    except Exception as e:
        return None, f"Could not read the file: {e}"


def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    cols = [str(c) for c in df.columns]
    low = {c: c.lower().strip() for c in cols}

    def find(patterns):
        for c in cols:
            if any(re.search(p, low[c], re.I) for p in patterns):
                return c
        return None

    return {
        "student_id": find([r"\broll\b", r"\breg\b", r"student.?id", r"registration"]),
        "name": find([r"^name$", r"student.?name", r"student"]),
        "total": find([r"grand.?total", r"g\.?\s*tot", r"\btotal\b"]),
    }


def infer_clo_columns(df: pd.DataFrame, clo_list: List[str]) -> Dict[str, str]:
    """Find likely CLO attainment/performance columns."""
    result = {}
    cols = [str(c) for c in df.columns]
    for clo in clo_list:
        n = re.sub(r"\s+", "", clo).upper()
        candidates = []
        for c in cols:
            compact = re.sub(r"\s+", "", c).upper()
            if n in compact and ("ATTAIN" in compact or "PERFORM" in compact or "%" in compact):
                candidates.append(c)
        if candidates:
            result[clo] = candidates[0]
    return result


def direct_attainment(series: pd.Series, threshold: float) -> Tuple[float, int, int]:
    s = pd.to_numeric(series, errors="coerce").dropna()
    n = len(s)
    if n == 0:
        return np.nan, 0, 0
    achieved = int((s >= threshold).sum())
    return achieved / n * 100, achieved, n


def weighted_attainment(df: pd.DataFrame, contribution_cols: List[Tuple[str, float]], threshold: float):
    if not contribution_cols:
        return np.nan, 0, 0
    work = pd.DataFrame(index=df.index)
    total_w = sum(w for _, w in contribution_cols)
    if total_w <= 0:
        return np.nan, 0, 0
    weighted = pd.Series(0.0, index=df.index)
    available = pd.Series(False, index=df.index)
    for col, weight in contribution_cols:
        if col in df.columns:
            vals = pd.to_numeric(df[col], errors="coerce")
            weighted = weighted + vals.fillna(0) * (weight / total_w)
            available = available | vals.notna()
    vals = weighted[available]
    if len(vals) == 0:
        return np.nan, 0, 0
    achieved = int((vals >= threshold).sum())
    return achieved / len(vals) * 100, achieved, len(vals)


def classify(attainment: float, bands: List[Dict]) -> str:
    if pd.isna(attainment):
        return "Not available"
    for band in sorted(bands, key=lambda x: x["min"], reverse=True):
        if attainment >= band["min"]:
            return band["label"]
    return "Not Achieved"


def make_excel_report(summary_df, assessment_df, mapping_df, student_df, metadata) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([metadata]).to_excel(writer, sheet_name="Course Information", index=False)
        summary_df.to_excel(writer, sheet_name="CLO Attainment", index=False)
        assessment_df.to_excel(writer, sheet_name="Assessments", index=False)
        mapping_df.to_excel(writer, sheet_name="CLO-PLO Mapping", index=False)
        student_df.to_excel(writer, sheet_name="Student Data", index=False)
    return output.getvalue()


# -----------------------------
# Session state defaults
# -----------------------------
if "assessments" not in st.session_state:
    st.session_state.assessments = [
        {"name": "Quiz 1", "total": 10.0, "weight": 5.0, "clos": ["CLO1"]},
        {"name": "Midterm", "total": 30.0, "weight": 20.0, "clos": ["CLO1", "CLO2"]},
        {"name": "Final Examination", "total": 50.0, "weight": 40.0, "clos": ["CLO2", "CLO3"]},
    ]

st.title("🎓 OBE Evaluation Report Generator")
st.caption("Configurable Course Learning Outcome (CLO) attainment and evidence-based course evaluation")

with st.sidebar:
    st.header("Report Settings")
    method = st.selectbox(
        "CLO attainment method",
        [
            "Direct threshold attainment",
            "Weighted assessment-based attainment",
            "Use uploaded/pre-calculated CLO attainment columns",
        ],
        help="Select the methodology required by your institution. The application does not hard-code one institutional formula.",
    )
    threshold = st.number_input(
        "CLO achievement threshold (%)",
        min_value=0.0, max_value=100.0, value=50.0, step=1.0,
        help="For direct attainment, a student is counted as achieving a CLO when their CLO performance is at or above this percentage."
    )
    target = st.number_input(
        "Target CLO attainment (%)",
        min_value=0.0, max_value=100.0, value=70.0, step=1.0,
    )

    st.subheader("Achievement Bands")
    default_bands = [
        ("Level 3", 90.0, "Excellent Achievement"),
        ("Level 2", 70.0, "Satisfactory Achievement"),
        ("Level 1", 50.0, "Partial Achievement"),
        ("Level 0", 0.0, "Not Achieved"),
    ]
    bands = []
    for level, minimum, label in default_bands:
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            lv = st.text_input("Level", value=level, key=f"lv_{level}")
        with c2:
            mn = st.number_input("Min %", 0.0, 100.0, minimum, 1.0, key=f"mn_{level}")
        with c3:
            lb = st.text_input("Description", value=label, key=f"lb_{level}")
        bands.append({"level": lv, "min": mn, "label": lb})

# -----------------------------
# Course information
# -----------------------------
st.header("1. Course Information")
c1, c2, c3 = st.columns(3)
with c1:
    institution = st.text_input("University / Institution", "FAST-NUCES")
    department = st.text_input("Department", "Sciences & Humanities")
    program = st.text_input("Program", "BSBA")
    course_title = st.text_input("Course Title", "English – II")
with c2:
    course_code = st.text_input("Course Code", "SS1006")
    semester = st.text_input("Semester", "Spring")
    academic_year = st.text_input("Academic Year", "2026")
    instructor = st.text_input("Instructor / Course Teacher", "Iram Amjad")
with c3:
    enrolled = st.number_input("Number Enrolled", 0, 10000, 25)
    assessed = st.number_input("Number Assessed", 0, 10000, 25)
    credit_hours = st.number_input("Credit Hours", 0.0, 20.0, 3.0, 0.5)
    report_date = st.date_input("Report Date", date.today())

# -----------------------------
# CLO / PLO
# -----------------------------
st.header("2. CLOs, PLOs & CLO–PLO Mapping")
default_clos = """CLO1: Demonstrate effective principles of communication.
CLO2: Apply oral and presentation skills appropriately.
CLO3: Demonstrate active listening and audience awareness.
CLO4: Adapt verbal and visual communication to context.
CLO5: Apply persuasive communication strategies."""
clo_text = st.text_area("Course Learning Outcomes (one per line)", default_clos, height=150)
clos = clo_ids(clo_text)
clo_descriptions = parse_clo_list(clo_text)

default_plos = "PLO1\nPLO2\nPLO3\nPLO4\nPLO5\nPLO6\nPLO7\nPLO8\nPLO9\nPLO10\nPLO11\nPLO12"
plo_text = st.text_area("Program Learning Outcomes (one per line)", default_plos, height=120)
plos = [x.strip() for x in plo_text.splitlines() if x.strip()]

mapping_rows = []
st.subheader("Interactive CLO–PLO Mapping")
if clos and plos:
    header = st.columns([1] + [1] * len(plos))
    header[0].markdown("**CLO**")
    for j, p in enumerate(plos):
        header[j+1].markdown(f"**{p}**")
    for i, clo in enumerate(clos):
        cells = st.columns([1] + [1] * len(plos))
        cells[0].markdown(f"**{clo}**")
        row = {"CLO": clo}
        for j, p in enumerate(plos):
            val = cells[j+1].selectbox(
                f"{clo}-{p}", [0, 1, 2, 3], index=0,
                key=f"map_{clo}_{p}",
                label_visibility="collapsed",
            )
            row[p] = val
        mapping_rows.append(row)

mapping_df = pd.DataFrame(mapping_rows)

# -----------------------------
# Assessment components
# -----------------------------
st.header("3. Assessment Components")
st.write("Add assessments and specify total marks, weightage, and CLO(s) assessed.")

if st.button("➕ Add Assessment"):
    st.session_state.assessments.append(
        {"name": f"Assessment {len(st.session_state.assessments)+1}", "total": 10.0, "weight": 5.0, "clos": [clos[0]] if clos else []}
    )

for idx, a in enumerate(st.session_state.assessments):
    with st.expander(f"{idx+1}. {a['name']}", expanded=True):
        q1, q2, q3, q4 = st.columns([3, 1.2, 1.2, 4])
        a["name"] = q1.text_input("Assessment Name", a["name"], key=f"an_{idx}")
        a["total"] = q2.number_input("Total Marks", 0.0, 10000.0, float(a["total"]), 0.5, key=f"at_{idx}")
        a["weight"] = q3.number_input("Weightage (%)", 0.0, 100.0, float(a["weight"]), 0.5, key=f"aw_{idx}")
        a["clos"] = q4.multiselect("CLO(s) Assessed", clos, default=[x for x in a["clos"] if x in clos], key=f"ac_{idx}")

if st.session_state.assessments:
    assessment_config_df = pd.DataFrame([
        {
            "Assessment Name": a["name"],
            "Total Marks": a["total"],
            "Weightage (%)": a["weight"],
            "CLO(s) Assessed": ", ".join(a["clos"]),
        }
        for a in st.session_state.assessments
    ])
    st.dataframe(assessment_config_df, use_container_width=True, hide_index=True)

# -----------------------------
# Data upload
# -----------------------------
st.header("4. Student Assessment Data")
uploaded = st.file_uploader("Upload Excel (.xlsx) or CSV (.csv)", type=["xlsx", "csv"])

df = None
if uploaded:
    df, err = read_uploaded_file(uploaded)
    if err:
        st.error(err)
    else:
        st.success(f"Loaded {len(df):,} rows and {len(df.columns):,} columns from **{uploaded.name}**.")
        st.dataframe(df.head(10), use_container_width=True)

if df is not None and len(df.columns):
    st.subheader("Column Mapping")
    detected = detect_columns(df)
    cols = [str(c) for c in df.columns]

    mc1, mc2, mc3 = st.columns(3)
    with mc1:
        student_id_col = st.selectbox("Student ID / Roll No.", ["— None —"] + cols,
                                      index=(cols.index(detected["student_id"]) + 1 if detected["student_id"] in cols else 0))
    with mc2:
        name_col = st.selectbox("Student Name", ["— None —"] + cols,
                                index=(cols.index(detected["name"]) + 1 if detected["name"] in cols else 0))
    with mc3:
        total_col = st.selectbox("Overall / Grand Total", ["— None —"] + cols,
                                 index=(cols.index(detected["total"]) + 1 if detected["total"] in cols else 0))

    selected_student_id = None if student_id_col == "— None —" else student_id_col
    selected_name = None if name_col == "— None —" else name_col

    st.subheader("CLO Performance Column Mapping")
    st.caption("For the uploaded/pre-calculated method, map each CLO to its percentage/performance column. You can override automatic suggestions.")
    inferred = infer_clo_columns(df, clos)
    clo_upload_map = {}
    map_cols = st.columns(min(3, max(1, len(clos))))
    for i, clo in enumerate(clos):
        with map_cols[i % len(map_cols)]:
            opts = ["— Not mapped —"] + cols
            suggested = inferred.get(clo)
            default_index = opts.index(suggested) if suggested in opts else 0
            chosen = st.selectbox(f"{clo} performance column", opts, index=default_index, key=f"uploadmap_{clo}")
            if chosen != "— Not mapped —":
                clo_upload_map[clo] = chosen

    if st.button("🔎 Preview Mapped Student Data"):
        preview = df.copy()
        rename_map = {}
        if selected_student_id:
            rename_map[selected_student_id] = "Student ID"
        if selected_name:
            rename_map[selected_name] = "Student Name"
        preview = preview.rename(columns=rename_map)
        st.dataframe(preview.head(20), use_container_width=True)

# -----------------------------
# CLO calculation
# -----------------------------
st.header("5. CLO Attainment Analysis")

summary_rows = []
calculation_notes = []

if df is not None and len(df) > 0:
    working = df.copy()

    # Uploaded/pre-calculated CLO columns
    if method == "Use uploaded/pre-calculated CLO attainment columns":
        for clo in clos:
            col = clo_upload_map.get(clo)
            if col and col in working.columns:
                vals = pd.to_numeric(working[col], errors="coerce")
                attainment = vals.dropna().mean()
                n = vals.notna().sum()
                achieved = int((vals.dropna() >= threshold).sum())
                summary_rows.append({
                    "CLO": clo,
                    "Attainment (%)": attainment,
                    "Students Achieving Threshold": achieved,
                    "Students Assessed": n,
                    "Target (%)": target,
                    "Gap to Target (pp)": attainment - target if not pd.isna(attainment) else np.nan,
                    "Achievement Level": classify(attainment, bands),
                    "Status": "Target Met" if attainment >= target else "Below Target",
                })
                calculation_notes.append(f"{clo}: mean of mapped uploaded percentage column '{col}'.")

    else:
        st.info("For generic Excel/CSV files, map assessment percentage columns below. A percentage column should contain each student's performance for that assessment (0–100).")
        percent_cols = [str(c) for c in working.columns]
        assessment_map = []
        for i, a in enumerate(st.session_state.assessments):
            with st.expander(f"Map: {a['name']}", expanded=False):
                col = st.selectbox(
                    "Student percentage/performance column",
                    ["— Not mapped —"] + percent_cols,
                    key=f"assessmap_{i}",
                )
                assessment_map.append((a, None if col == "— Not mapped —" else col))

        for clo in clos:
            relevant = [(a, col) for a, col in assessment_map if clo in a["clos"] and col]
            if method == "Direct threshold attainment":
                # Direct method: combine assessment evidence for this CLO using the mean
                # of the mapped assessment percentages per student.
                if relevant:
                    tmp = pd.DataFrame(index=working.index)
                    for a, col in relevant:
                        tmp[a["name"]] = pd.to_numeric(working[col], errors="coerce")
                    student_clo = tmp.mean(axis=1, skipna=True).dropna()
                    attainment, achieved, n = direct_attainment(student_clo, threshold)
                    calculation_notes.append(
                        f"{clo}: student CLO performance = mean of mapped assessment percentages; "
                        f"attainment = students at/above {threshold:.1f}% ÷ students with CLO evidence × 100."
                    )
                else:
                    attainment, achieved, n = np.nan, 0, 0
            else:
                contributions = []
                for a, col in relevant:
                    contributions.append((col, safe_float(a["weight"])))
                attainment, achieved, n = weighted_attainment(working, contributions, threshold)
                calculation_notes.append(
                    f"{clo}: weighted assessment evidence using configured assessment weightages."
                )

            summary_rows.append({
                "CLO": clo,
                "Attainment (%)": attainment,
                "Students Achieving Threshold": achieved,
                "Students Assessed": n,
                "Target (%)": target,
                "Gap to Target (pp)": attainment - target if not pd.isna(attainment) else np.nan,
                "Achievement Level": classify(attainment, bands),
                "Status": "Target Met" if not pd.isna(attainment) and attainment >= target else "Below Target",
            })

summary_df = pd.DataFrame(summary_rows)

if not summary_df.empty:
    summary_df["Attainment (%)"] = summary_df["Attainment (%)"].round(2)
    summary_df["Gap to Target (pp)"] = summary_df["Gap to Target (pp)"].round(2)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    valid = summary_df["Attainment (%)"].dropna()
    if len(valid):
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Average CLO Attainment", f"{valid.mean():.2f}%")
        m2.metric("Highest CLO", summary_df.loc[summary_df["Attainment (%)"].idxmax(), "CLO"])
        m3.metric("Lowest CLO", summary_df.loc[summary_df["Attainment (%)"].idxmin(), "CLO"])
        m4.metric("CLOs Meeting Target", f"{int((valid >= target).sum())}/{len(valid)}")

    st.subheader("CLO Attainment Visualization")
    chart_df = summary_df.dropna(subset=["Attainment (%)"]).set_index("CLO")[["Attainment (%)", "Target (%)"]]
    st.bar_chart(chart_df)

    st.subheader("Auditable Calculation Method")
    st.markdown(f"**Selected method:** {method}")
    if method == "Direct threshold attainment":
        st.latex(
            r"\text{CLO Attainment (\%)} = "
            r"\frac{\text{Number of students achieving the CLO threshold}}"
            r"{\text{Number of students with valid CLO evidence}}\times100"
        )
    elif method == "Weighted assessment-based attainment":
        st.latex(
            r"\text{CLO Attainment} = \sum_i "
            r"\left(\text{Assessment Weight}_i \times \text{CLO Performance}_i\right)"
        )
        st.caption("Weights are normalized over the assessments mapped to each CLO when necessary.")
    else:
        st.write("The application uses the uploaded/pre-calculated CLO performance columns selected in the mapping section.")

    for note in calculation_notes:
        st.write("•", note)

    strong = summary_df[summary_df["Attainment (%)"] >= target]["CLO"].tolist()
    weak = summary_df[summary_df["Attainment (%)"] < target]["CLO"].tolist()

    st.subheader("Strengths, Weaknesses & Recommendations")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Strong CLOs**")
        st.write(", ".join(strong) if strong else "No CLO met the target.")
    with col2:
        st.markdown("**CLOs Requiring Improvement**")
        st.write(", ".join(weak) if weak else "All CLOs met the target.")

    if weak:
        recs = []
        for clo in weak:
            val = summary_df.loc[summary_df["CLO"] == clo, "Attainment (%)"].iloc[0]
            gap = target - val if not pd.isna(val) else None
            if gap is not None:
                recs.append(
                    f"**{clo}**: attainment is {val:.2f}%, which is {gap:.2f} percentage points below the target. "
                    "Review the assessment evidence mapped to this CLO, provide targeted formative practice, "
                    "and consider adjusting teaching/assessment alignment in the next offering."
                )
        for r in recs:
            st.markdown("• " + r)
    else:
        st.success("All evaluated CLOs met or exceeded the configured target.")

# -----------------------------
# Report narrative
# -----------------------------
st.header("6. OBE Evaluation Report")
if summary_df.empty:
    st.info("Upload student assessment data and complete the CLO mappings to generate the report.")
else:
    valid = summary_df["Attainment (%)"].dropna()
    avg = valid.mean() if len(valid) else np.nan
    met = int((valid >= target).sum()) if len(valid) else 0
    total = len(valid)

    narrative = f"""
### Course Evaluation Summary

The OBE evaluation for **{course_title} ({course_code})** was conducted for **{program}**, {semester} {academic_year}. 
The course was taught by **{instructor}** and carries **{credit_hours:g} credit hours**. 
The analysis included **{assessed} assessed students** out of **{enrolled} enrolled students**.

Using the selected **{method}** methodology and a configured achievement threshold of **{threshold:.1f}%**, 
the average attainment across evaluated CLOs was **{avg:.2f}%**. 
**{met} of {total} evaluated CLOs** met or exceeded the configured target of **{target:.1f}%**.

The attainment results should be interpreted together with the assessment-to-CLO mapping and the institution's approved OBE policy. 
CLOs below the target should be considered priorities for continuous quality improvement (CQI), including review of teaching strategies, 
assessment design, learner practice opportunities, and alignment between CLOs and assessments.
"""
    st.markdown(narrative)

# -----------------------------
# Export
# -----------------------------
st.header("7. Export")
if df is not None:
    metadata = {
        "Institution": institution,
        "Department": department,
        "Program": program,
        "Course Title": course_title,
        "Course Code": course_code,
        "Semester": semester,
        "Academic Year": academic_year,
        "Instructor": instructor,
        "Enrolled": enrolled,
        "Assessed": assessed,
        "Credit Hours": credit_hours,
        "Method": method,
        "CLO Threshold (%)": threshold,
        "Target (%)": target,
    }
    excel_bytes = make_excel_report(
        summary_df,
        assessment_config_df,
        mapping_df,
        df,
        metadata,
    )
    st.download_button(
        "⬇️ Download OBE Evaluation Report (Excel)",
        data=excel_bytes,
        file_name=f"OBE_Evaluation_{course_code}_{semester}_{academic_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    report_text = f"""OBE EVALUATION REPORT

Institution: {institution}
Department: {department}
Program: {program}
Course: {course_title} ({course_code})
Semester: {semester} {academic_year}
Instructor: {instructor}
Students Enrolled: {enrolled}
Students Assessed: {assessed}
Credit Hours: {credit_hours}

CLO ATTAINMENT
{summary_df.to_string(index=False)}

METHODOLOGY
{method}
Achievement threshold: {threshold:.1f}%
Target attainment: {target:.1f}%

The application reports the configured calculation method and preserves the underlying student data and mappings in the Excel export.
"""
    st.download_button(
        "⬇️ Download Report Summary (TXT)",
        data=report_text,
        file_name=f"OBE_Report_Summary_{course_code}.txt",
        mime="text/plain",
    )

st.divider()
st.caption("OBE Evaluation Report Generator • Configurable methodology • Auditable calculations • Excel/CSV input")
