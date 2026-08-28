import io
import re
import streamlit as st
import pandas as pd
import numpy as np

try:
    from docx import Document
except Exception:
    Document = None

try:
    import matplotlib.pyplot as plt
except Exception:
    plt = None

st.set_page_config(
    page_title="OBE Evaluation Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

NOT_AVAILABLE = "Not available in the provided files."

st.markdown("""
<style>
.main-title{font-size:3rem;font-weight:800;color:#172554;margin-bottom:.2rem}
.subtitle{font-size:1.05rem;color:#64748b;margin-bottom:2rem}
.section-title{font-size:2rem;font-weight:800;color:#173b68;margin-top:1.8rem}
.card{background:#f4f7fb;border:1px solid #dbe4ef;border-radius:14px;padding:18px;min-height:105px}
.card-label{color:#64748b;font-size:.9rem}
.card-value{color:#172033;font-size:1.05rem;font-weight:650;margin-top:8px}
.info-box{background:#eaf4ff;border-radius:10px;padding:14px 18px;color:#075985;border:1px solid #cfe7ff}
</style>
""", unsafe_allow_html=True)

# ---------------- Sidebar ----------------
with st.sidebar:
    st.markdown("## Report Settings")
    method = st.selectbox(
        "CLO attainment method",
        ["Direct student-threshold attainment",
         "Assessment-score attainment",
         "Use evidence available in Excel"]
    )
    benchmark = st.number_input("Achievement threshold (%)", 0.0, 100.0, 70.0, 1.0)
    target = st.number_input("Target CLO attainment (%)", 0.0, 100.0, 70.0, 1.0)
    st.divider()
    st.caption("Enter information manually, upload source files, or use both.")

# ---------------- Header ----------------
st.markdown(
    '<div class="main-title">📊 OBE Evaluation Report Generator</div>',
    unsafe_allow_html=True
)
st.markdown(
    '<div class="subtitle">Generate an auditable Course Learning Outcome (CLO) attainment report from assessment data.</div>',
    unsafe_allow_html=True
)

# ---------------- 1 Course Information ----------------
st.markdown('<div class="section-title">1. Course Information</div>', unsafe_allow_html=True)
st.caption("You can type these details yourself. They do not have to come from the uploaded files.")

c1, c2, c3 = st.columns(3)

with c1:
    institution = st.text_input("Institution", placeholder="e.g., FAST-NUCES")
    department = st.text_input("Department", placeholder="e.g., English")
    program = st.text_input("Program", placeholder="e.g., BS English")

with c2:
    course_title = st.text_input("Course Title", placeholder="e.g., English II")
    course_code = st.text_input("Course Code", placeholder="e.g., SS-1006")
    semester = st.text_input("Semester", placeholder="e.g., Spring 2026")

with c3:
    academic_year = st.text_input("Academic Year", placeholder="e.g., 2025-26")
    instructor = st.text_input("Course Teacher / Instructor", placeholder="Instructor name")
    credit_hours = st.number_input(
        "Credit Hours", min_value=0.0, max_value=20.0, value=3.0, step=0.5
    )

# ---------------- 2 CLO ----------------
st.markdown('<div class="section-title">2. Course Learning Outcomes</div>', unsafe_allow_html=True)
st.caption("Type the official CLO wording manually OR upload the Course Outline and review the detected CLOs.")

outline_file = st.file_uploader(
    "Upload Course Outline (.docx)",
    type=["docx"],
    key="outline"
)

def docx_text(data):
    if Document is None:
        return ""
    try:
        doc = Document(io.BytesIO(data))
        parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))
        return "\n".join(parts)
    except Exception:
        return ""

def detect_clos(text):
    rows = []
    for line in text.splitlines():
        line = line.strip()
        match = re.match(
            r"^(CLO\s*\d+)\s*[:.\-–—]?\s*(.+)$",
            line,
            flags=re.I
        )
        if match:
            rows.append({
                "CLO": match.group(1).upper(),
                "Official CLO Description": match.group(2).strip()
            })
    return rows

default_clos = pd.DataFrame(
    [{"CLO": "CLO 1", "Official CLO Description": ""}],
    columns=["CLO", "Official CLO Description"]
)

if outline_file:
    detected = detect_clos(docx_text(outline_file.getvalue()))
    if detected:
        default_clos = pd.DataFrame(detected)
        st.success(f"{len(detected)} CLO(s) detected. Review/edit them below.")
    else:
        st.warning(
            "No clearly labelled CLO lines were detected. "
            "Enter the official CLO wording manually."
        )

clo_df = st.data_editor(
    default_clos,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "CLO": st.column_config.TextColumn("CLO", required=True),
        "Official CLO Description": st.column_config.TextColumn(
            "Official CLO Description", width="large"
        ),
    },
    key="clo_editor"
)

# ---------------- 3 Student Assessment Data ----------------
st.markdown('<div class="section-title">3. Student Assessment Data</div>', unsafe_allow_html=True)
st.caption(
    "Upload the official Excel/CSV here. Student names/IDs do not need to be stored in GitHub."
)

assessment_file = st.file_uploader(
    "Upload Student Assessment Data",
    type=["xlsx", "xls", "csv"],
    key="assessment"
)

df = None

if assessment_file:
    try:
        if assessment_file.name.lower().endswith(".csv"):
            df = pd.read_csv(assessment_file)
        else:
            xls = pd.ExcelFile(assessment_file)
            sheet = st.selectbox("Select worksheet", xls.sheet_names)
            df = pd.read_excel(assessment_file, sheet_name=sheet)

        st.success(f"Loaded: {assessment_file.name}")
        st.dataframe(df.head(20), use_container_width=True)

    except Exception as e:
        st.error(f"Could not load the file: {e}")

# ---------------- 4 Manual Assessment Entry ----------------
st.markdown('<div class="section-title">4. Manual Assessment Entry</div>', unsafe_allow_html=True)
st.caption(
    "If you do not have an Excel file, enable this option and enter a demonstration dataset."
)

manual = st.checkbox("Enable manual student/assessment entry")

if manual:
    demo = pd.DataFrame({
        "Student ID": ["S001", "S002", "S003"],
        "CLO 1": [78, 65, 84],
        "CLO 2": [72, 61, 90],
        "CLO 3": [81, 69, 76],
    })

    manual_df = st.data_editor(
        demo,
        num_rows="dynamic",
        use_container_width=True,
        key="manual_data"
    )

    if st.button("Use manual data for analysis"):
        df = manual_df.copy()
        st.session_state["manual_selected"] = True
        st.success("Manual data selected.")

if st.session_state.get("manual_selected", False) and manual:
    df = manual_df.copy()

# ---------------- 5 Generate Dashboard ----------------
st.markdown('<div class="section-title">5. Generate OBE Dashboard</div>', unsafe_allow_html=True)

generate = st.button(
    "📊 Generate / Refresh OBE Analysis",
    type="primary",
    use_container_width=True
)

if generate:

    # Course information cards
    st.markdown("### Course Information")

    info = [
        ("Institution", institution or NOT_AVAILABLE),
        ("Department", department or NOT_AVAILABLE),
        ("Program", program or NOT_AVAILABLE),
        ("Course Title", course_title or NOT_AVAILABLE),
        ("Course Code", course_code or NOT_AVAILABLE),
        ("Semester", semester or NOT_AVAILABLE),
        ("Academic Year", academic_year or NOT_AVAILABLE),
        ("Course Teacher / Instructor", instructor or NOT_AVAILABLE),
        ("Credit Hours", f"{credit_hours:.2f}" if credit_hours else NOT_AVAILABLE),
    ]

    cols = st.columns(3)

    for i, (label, value) in enumerate(info):
        with cols[i % 3]:
            st.markdown(
                f'<div class="card"><div class="card-label">{label}</div>'
                f'<div class="card-value">{value}</div></div>',
                unsafe_allow_html=True
            )

    # CLO table
    st.markdown("### Course Learning Outcomes")
    st.dataframe(clo_df, use_container_width=True, hide_index=True)

    if df is None or df.empty:
        st.markdown(
            '<div class="info-box">Please upload the OBE Excel/CSV file or enable manual assessment entry.</div>',
            unsafe_allow_html=True
        )

    else:
        numeric = df.select_dtypes(include=np.number).columns.tolist()

        if not numeric:
            st.warning("No numeric assessment columns were detected.")

        else:
            values = df[numeric].apply(pd.to_numeric, errors="coerce")
            total = values.sum(axis=1, min_count=1).dropna()

            if len(total):
                st.markdown("### Student Performance")

                a, b, c, d, e, f = st.columns(6)
                a.metric("N", len(total))
                b.metric("Highest", f"{total.max():.2f}")
                c.metric("Lowest", f"{total.min():.2f}")
                d.metric("Mean", f"{total.mean():.2f}")
                e.metric("Median", f"{total.median():.2f}")
                f.metric(
                    "Benchmark",
                    f"{(total >= benchmark).mean() * 100:.1f}%"
                )

            # Detect CLO columns
            clo_cols = [
                c for c in df.columns
                if re.match(r"^CLO\s*\d+", str(c), re.I)
            ]

            if clo_cols:

                rows = []

                for col in clo_cols:
                    s = pd.to_numeric(df[col], errors="coerce").dropna()

                    if len(s):
                        mean = float(s.mean())

                        if mean >= 80:
                            status = "Strong"
                        elif mean >= 70:
                            status = "Satisfactory"
                        else:
                            status = "Needs Improvement"

                        rows.append({
                            "CLO": str(col),
                            "Mean Attainment (%)": round(mean, 2),
                            "Students ≥ Benchmark": int((s >= benchmark).sum()),
                            "Benchmark Achievement (%)": round(
                                (s >= benchmark).mean() * 100, 2
                            ),
                            "Status": status
                        })

                result = pd.DataFrame(rows)

                st.markdown("### CLO Attainment")
                st.dataframe(result, use_container_width=True, hide_index=True)

                if plt is not None and not result.empty:

                    fig, ax = plt.subplots(figsize=(9, 4.8))
                    ax.bar(result["CLO"], result["Mean Attainment (%)"])
                    ax.axhline(
                        benchmark,
                        linestyle="--",
                        linewidth=2,
                        label=f"{benchmark:.0f}% benchmark"
                    )
                    ax.set_ylim(0, 100)
                    ax.set_ylabel("Attainment (%)")
                    ax.set_xlabel("Course Learning Outcome")
                    ax.set_title("CLO Attainment")
                    ax.legend()
                    ax.grid(axis="y", alpha=0.2)
                    st.pyplot(fig)
                    st.caption(
                        "Figure 1. CLO mean attainment compared with the selected benchmark."
                    )

                    fig2, ax2 = plt.subplots(figsize=(9, 4.8))
                    ax2.bar(
                        result["CLO"],
                        result["Benchmark Achievement (%)"]
                    )
                    ax2.axhline(
                        70,
                        linestyle="--",
                        linewidth=2,
                        label="70% benchmark achievement"
                    )
                    ax2.set_ylim(0, 100)
                    ax2.set_ylabel("Students achieving benchmark (%)")
                    ax2.set_xlabel("Course Learning Outcome")
                    ax2.set_title("Benchmark Achievement by CLO")
                    ax2.legend()
                    ax2.grid(axis="y", alpha=0.2)
                    st.pyplot(fig2)
                    st.caption(
                        "Figure 2. Percentage of students meeting or exceeding the benchmark for each CLO."
                    )

            else:
                st.info(
                    "No columns labelled CLO 1, CLO 2, etc. were detected. "
                    "The workbook may store its CLO mapping on another worksheet."
                )

# ---------------- Audit rules ----------------
st.markdown('<div class="section-title">6. Audit Rules</div>', unsafe_allow_html=True)
st.markdown("""
- Course Outline is authoritative for exact official CLO wording.
- Excel is authoritative for numerical calculations, assessment evidence and mappings.
- Missing information is reported as **“Not available in the provided files.”**
- Maximum marks are never inferred when absent.
- Status: **≥80 Strong; 70–79.99 Satisfactory; <70 Needs Improvement.**
- Real student-identifiable files should be uploaded inside the app, not committed to a public GitHub repository.
""")
