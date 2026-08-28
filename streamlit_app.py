import io
import re
from pathlib import Path

import pandas as pd
import numpy as np
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import matplotlib.pyplot as plt


# ---------------------------------------------------------
# PAGE / UI
# ---------------------------------------------------------
st.set_page_config(
    page_title="OBE Evaluation Report Generator",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
.main-title {
    font-size: 42px;
    font-weight: 800;
    color: #17365D;
    margin-bottom: 0;
}
.subtitle {
    color: #6B7280;
    font-size: 17px;
    margin-bottom: 25px;
}
.section-title {
    color: #17365D;
    font-size: 30px;
    font-weight: 750;
    margin-top: 25px;
}
.info-card {
    background: #F4F7FB;
    border: 1px solid #D9E2F3;
    border-radius: 12px;
    padding: 14px 18px;
    min-height: 90px;
}
.info-label {
    color: #64748B;
    font-size: 14px;
}
.info-value {
    color: #172033;
    font-size: 17px;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 OBE Evaluation Report Generator</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Generate an auditable CLO attainment analysis, Word report, Excel workbook and separate chart files from your Course Outline and OBE assessment data.</div>',
    unsafe_allow_html=True
)

# ---------------------------------------------------------
# SIDEBAR SETTINGS
# ---------------------------------------------------------
with st.sidebar:
    st.header("Report Settings")
    benchmark = st.number_input(
        "OBE benchmark (%)", min_value=0.0, max_value=100.0, value=70.0, step=1.0
    )
    st.caption("Status: ≥80 Strong | 70–79.99 Satisfactory | <70 Needs Improvement")
    st.divider()
    st.subheader("Master Prompt")
    st.caption("The supplied OBE master-prompt rules are applied to the generated report package.")

BENCHMARK = benchmark


# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------
def clean(v):
    if pd.isna(v):
        return ""
    return str(v).replace("\ufeff", "").strip()


def num(v):
    try:
        return float(v)
    except Exception:
        return np.nan


def status(x):
    if pd.isna(x):
        return "Not available"
    if x >= 80:
        return "Strong"
    if x >= 70:
        return "Satisfactory"
    return "Needs Improvement"


def safe_value(v):
    return v if v else "Not available in the provided files."


# ---------------------------------------------------------
# COURSE OUTLINE PARSER
# ---------------------------------------------------------
def parse_outline(data):
    d = Document(io.BytesIO(data))
    lines = [clean(p.text) for p in d.paragraphs if clean(p.text)]
    text = "\n".join(lines)

    info = {
        k: ""
        for k in [
            "Institution",
            "Department",
            "Program",
            "Course Title",
            "Course Code",
            "Semester",
            "Academic Year",
            "Campus",
            "Instructor/Faculty",
            "Credit Hours",
            "Section",
            "Course Description",
        ]
    }

    patterns = {
        "Course Title": r"(?:Course|Course Title):\s*([^\n]+)",
        "Course Code": r"(?:Course code|Course Code):\s*([^\n]+)",
        "Semester": r"(?:Year/Semester|Semester):\s*([^\n]+)",
        "Program": r"Program:\s*([^\n]+)",
        "Credit Hours": r"(?:Units/Cr Hrs\.|Credit Hours|Cr Hrs\.):\s*([^\n]+)",
        "Instructor/Faculty": r"(?:Instructor|Teacher|Faculty):\s*([^\n]*)",
        "Academic Year": r"Academic Year:\s*([^\n]+)",
        "Department": r"Department:\s*([^\n]+)",
        "Institution": r"Institution:\s*([^\n]+)",
        "Campus": r"Campus:\s*([^\n]+)",
        "Section": r"Section:\s*([^\n]+)",
    }

    for key, pattern in patterns.items():
        m = re.search(pattern, text, re.I)
        if m:
            info[key] = clean(m.group(1))

    if not info["Campus"] and "Lahore Campus" in text:
        info["Campus"] = "Lahore Campus"

    m = re.search(
        r"COURSE DESCRIPTION\s*(.*?)(?:Program Educational Objectives|Course Objectives|Course Learning Outcomes)",
        text,
        re.I | re.S,
    )
    if m:
        info["Course Description"] = " ".join(m.group(1).split())

    # CLOs from tables
    clos = {}
    for table in d.tables:
        for row in table.rows:
            vals = [clean(c.text) for c in row.cells]
            if vals and re.fullmatch(r"CLO\s*\d+", vals[0], re.I) and len(vals) > 1:
                clos[re.sub(r"\s+", "", vals[0]).upper()] = vals[1]

    # More robust paragraph fallback
    if not clos:
        for i, line in enumerate(lines):
            m = re.match(r"^(CLO\s*\d+)\s*[:\-]?\s*(.*)$", line, re.I)
            if m:
                key = re.sub(r"\s+", "", m.group(1)).upper()
                desc = m.group(2).strip()
                if desc:
                    clos[key] = desc
                elif i + 1 < len(lines):
                    clos[key] = lines[i + 1]

    objectives = []
    m = re.search(
        r"Course Objectives\s*(.*?)(?:Program Learning Outcome|Course Learning Outcomes|CLO\s*1)",
        text,
        re.I | re.S,
    )
    if m:
        block = m.group(1)
        for n in range(1, 15):
            mm = re.search(rf"(?:^|\n){n}\s+(.+?)(?=\n\d\s+|$)", block, re.S)
            if mm:
                objectives.append(" ".join(mm.group(1).split()))

    return info, objectives, clos


# ---------------------------------------------------------
# CURRENT OBE EXCEL ANALYSIS
# Keeps the established SS1006-style OBE workbook mapping.
# ---------------------------------------------------------
def analyze(raw, clos):
    pct_cols = {
        "CLO1": 9,
        "CLO2": 14,
        "CLO3": 22,
        "CLO4": 29,
        "CLO5": 34,
    }
    total_col = 35

    groups = {
        "CLO1": [(5, "Qz :1"), (6, "S-I :3"), (7, "Final :5")],
        "CLO2": [(10, "S-II :2"), (11, "Qz :2"), (12, "Final :4")],
        "CLO3": [
            (14, "PRS :1"),
            (15, "PRS :2"),
            (16, "S-I :2"),
            (17, "Qz :3"),
            (18, "PRS :5"),
            (19, "Final :6"),
        ],
        "CLO4": [
            (23, "S-I :1"),
            (24, "PRS :3"),
            (25, "PRS :4"),
            (26, "Final :2"),
            (27, "Final :3"),
        ],
        "CLO5": [(30, "S-II :1"), (31, "Final :1"), (32, "Final :7")],
    }

    assessments = []
    for clo, items in groups.items():
        for col, label in items:
            if col < raw.shape[1]:
                assessments.append({
                    "clo": clo,
                    "label": label,
                    "weightage": num(raw.iloc[4, col]),
                    "average": num(raw.iloc[5, col]),
                    "date": clean(raw.iloc[2, col]),
                })

    rows = list(range(10, min(35, raw.shape[0])))
    stats = {}

    for clo in clos:
        if clo not in pct_cols or pct_cols[clo] >= raw.shape[1]:
            stats[clo] = {
                "n": 0, "mean": np.nan, "sd": np.nan,
                "n70": 0, "pct70": np.nan
            }
            continue

        s = pd.to_numeric(raw.loc[rows, pct_cols[clo]], errors="coerce").dropna()
        stats[clo] = {
            "n": len(s),
            "mean": s.mean() if len(s) else np.nan,
            "sd": s.std(ddof=1) if len(s) > 1 else np.nan,
            "n70": int((s >= BENCHMARK).sum()) if len(s) else 0,
            "pct70": (s >= BENCHMARK).mean() * 100 if len(s) else np.nan,
        }

    if total_col < raw.shape[1]:
        total = pd.to_numeric(raw.loc[rows, total_col], errors="coerce").dropna()
    else:
        total = pd.Series(dtype=float)

    overall = {
        "n": len(total),
        "highest": total.max() if len(total) else np.nan,
        "lowest": total.min() if len(total) else np.nan,
        "mean": total.mean() if len(total) else np.nan,
        "median": total.median() if len(total) else np.nan,
        "sd": total.std(ddof=1) if len(total) > 1 else np.nan,
        "benchmark_pct": (total >= BENCHMARK).mean() * 100 if len(total) else np.nan,
    }

    gtot = num(raw.iloc[5, total_col]) if total_col < raw.shape[1] else np.nan
    return assessments, stats, overall, gtot, rows, pct_cols, total_col


# ---------------------------------------------------------
# CHARTS
# ---------------------------------------------------------
def charts(stats, assessments, out):
    out = Path(out)
    out.mkdir(exist_ok=True)
    paths = []
    cs = list(stats)

    # Figure 1
    p = out / "CLO_Attainment_Chart.png"
    vals = [stats[c]["mean"] for c in cs]
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(cs, vals)
    ax.axhline(BENCHMARK, linestyle="--", linewidth=1.5, label=f"{BENCHMARK:.0f}% benchmark")
    ax.set_title("Figure 1. CLO-wise OBE Attainment")
    ax.set_xlabel("Course Learning Outcome")
    ax.set_ylabel("Mean Attainment (%)")
    ax.set_ylim(0, 100)
    ax.legend()
    for i, v in enumerate(vals):
        if not pd.isna(v):
            ax.text(i, min(v + 2, 97), f"{v:.2f}%", ha="center")
    fig.tight_layout()
    fig.savefig(p, dpi=200, bbox_inches="tight")
    plt.close(fig)
    paths.append(p)

    # Figure 2
    p = out / "Benchmark_Achievement_Chart.png"
    vals = [stats[c]["pct70"] for c in cs]
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(cs, vals)
    ax.axhline(BENCHMARK, linestyle="--", linewidth=1.5, label=f"{BENCHMARK:.0f}% benchmark")
    ax.set_title(f"Figure 2. Students Achieving ≥{BENCHMARK:.0f}% by CLO")
    ax.set_xlabel("Course Learning Outcome")
    ax.set_ylabel(f"Students achieving ≥{BENCHMARK:.0f}% (%)")
    ax.set_ylim(0, 100)
    ax.legend()
    for i, v in enumerate(vals):
        if not pd.isna(v):
            ax.text(i, min(v + 2, 97), f"{v:.0f}%", ha="center")
    fig.tight_layout()
    fig.savefig(p, dpi=200, bbox_inches="tight")
    plt.close(fig)
    paths.append(p)

    # Figure 3
    p = out / "Assessment_Performance_Chart.png"
    vals = [a["average"] for a in assessments]
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.bar(range(len(vals)), vals)
    ax.set_title("Figure 3. Assessment Mean Scores as Reported in Excel")
    ax.set_xlabel("Assessment / CLO")
    ax.set_ylabel("Mean score (raw scale as provided)")
    ax.set_xticks(range(len(vals)))
    ax.set_xticklabels(
        [f"{a['clo']}\n{a['label']}" for a in assessments],
        rotation=55, ha="right", fontsize=8
    )
    fig.tight_layout()
    fig.savefig(p, dpi=200, bbox_inches="tight")
    plt.close(fig)
    paths.append(p)

    return paths


# ---------------------------------------------------------
# EXCEL OUTPUT
# ---------------------------------------------------------
def workbook(clos, assessments, stats, overall, student_df, out):
    wb = Workbook()
    wb.remove(wb.active)

    title = PatternFill("solid", fgColor="1F4E78")
    head = PatternFill("solid", fgColor="D9EAF7")

    def setup(ws, title_text, headers):
        ws.merge_cells(
            start_row=1, start_column=1,
            end_row=1, end_column=len(headers)
        )
        ws.cell(1, 1, title_text)
        ws.cell(1, 1).fill = title
        ws.cell(1, 1).font = Font(color="FFFFFF", bold=True, size=12)

        for j, h in enumerate(headers, 1):
            ws.cell(3, j, h).fill = head
            ws.cell(3, j).font = Font(bold=True)
            ws.cell(3, j).alignment = Alignment(wrap_text=True)

        ws.freeze_panes = "A4"
        ws.sheet_view.showGridLines = False

    def put(ws, rows, start=4):
        for r, row in enumerate(rows, start):
            for j, value in enumerate(row, 1):
                ws.cell(r, j, value)
                ws.cell(r, j).alignment = Alignment(
                    vertical="top", wrap_text=True
                )

    ws = wb.create_sheet("OBE Summary")
    setup(ws, "OBE Summary", [
        "CLO", "Official CLO Description", "Mean Attainment (%)",
        f"Students ≥{BENCHMARK:.0f}%", f"Students ≥{BENCHMARK:.0f}% (%)",
        "SD", "Status"
    ])
    put(ws, [
        [
            c, clos[c], stats[c]["mean"], stats[c]["n70"],
            stats[c]["pct70"], stats[c]["sd"], status(stats[c]["mean"])
        ]
        for c in clos
    ])

    ws = wb.create_sheet("CLO–Assessment Mapping")
    setup(ws, "CLO–Assessment Mapping", [
        "CLO", "Official CLO Description", "Assessment/Question",
        "Weightage", "Average/Mean Score", "Maximum Marks",
        "Attainment %", "Source"
    ])
    put(ws, [
        [
            a["clo"],
            clos.get(a["clo"], "Unmatched CLO in Excel"),
            a["label"],
            a["weightage"],
            a["average"],
            "Not available in the provided files.",
            "Not available in the provided files.",
            "Excel OBE sheet",
        ]
        for a in assessments
    ])

    ws = wb.create_sheet("Assessment Analysis")
    setup(ws, "Assessment Analysis", [
        "CLO", "Assessment/Question", "Weightage",
        "Average/Mean Score", "Assessment Attainment %",
        "Interpretation"
    ])
    put(ws, [
        [
            a["clo"], a["label"], a["weightage"], a["average"],
            "Not available in the provided files.",
            "Raw maximum marks are not separately supplied; normalized attainment is not inferred."
        ]
        for a in assessments
    ])

    ws = wb.create_sheet("Student/CLO Data")
    setup(ws, "Student/CLO Data", list(student_df.columns))
    put(ws, student_df.values.tolist())

    ws = wb.create_sheet("CQI Action Plan")
    setup(ws, "CQI Action Plan", [
        "CLO/Area", "Identified Issue", "Recommended Action",
        "Teaching/Learning Intervention", "Follow-up Evidence", "Target"
    ])
    put(ws, [
        [
            c,
            f"Mean attainment {stats[c]['mean']:.2f}% "
            f"{'is below' if stats[c]['mean'] < BENCHMARK else 'meets/exceeds'} "
            f"the {BENCHMARK:.0f}% benchmark." if not pd.isna(stats[c]["mean"])
            else "Not available in the provided files.",
            f"Align intervention directly with the official CLO: {clos[c]}",
            "Use targeted practice, formative assessment, guided application and feedback tied to this CLO.",
            "Repeat CLO-aligned assessment and compare attainment and benchmark achievement.",
            f"Mean attainment ≥{BENCHMARK:.0f}%.",
        ]
        for c in clos
    ])

    ws = wb.create_sheet("Chart Data")
    setup(ws, "Chart Data", [
        "CLO", "Mean Attainment (%)",
        f"Students ≥{BENCHMARK:.0f}% (%)", "Status"
    ])
    put(ws, [
        [c, stats[c]["mean"], stats[c]["pct70"], status(stats[c]["mean"])]
        for c in clos
    ])

    wb.save(out)


# ---------------------------------------------------------
# WORD REPORT
# ---------------------------------------------------------
def report(info, objectives, clos, assessments, stats, overall, gtot,
           student_df, chart_paths, out):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = __import__("docx").shared.Inches(.65)
    sec.left_margin = sec.right_margin = __import__("docx").shared.Inches(.75)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(
        "OUTCOME-BASED EDUCATION (OBE)\n"
        "EVALUATION REPORT\n"
    ).bold = True
    p.add_run(
        f"\n{safe_value(info.get('Course Title'))} "
        f"({safe_value(info.get('Course Code'))})\n"
        f"Program: {safe_value(info.get('Program'))}\n"
        f"Section: {safe_value(info.get('Section'))}\n"
        f"Semester: {safe_value(info.get('Semester'))}\n"
    )

    doc.add_page_break()

    def table(headers, rows):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        for i, h in enumerate(headers):
            t.rows[0].cells[i].text = str(h)
        for row in rows:
            cells = t.add_row().cells
            for i, value in enumerate(row):
                cells[i].text = str(value)

    valid_means = {c: stats[c]["mean"] for c in clos if not pd.isna(stats[c]["mean"])}
    strong = max(valid_means, key=valid_means.get) if valid_means else "Not available"
    weak = min(valid_means, key=valid_means.get) if valid_means else "Not available"

    doc.add_heading("1. Executive Summary", 1)
    table(["Item", "Result"], [
        ["Course", f"{safe_value(info.get('Course Title'))} ({safe_value(info.get('Course Code'))})"],
        ["Program / Section", f"{safe_value(info.get('Program'))} / {safe_value(info.get('Section'))}"],
        ["Semester", safe_value(info.get("Semester"))],
        ["Instructor", safe_value(info.get("Instructor/Faculty"))],
        ["Number of students", overall["n"]],
        ["Overall CLO/course attainment", f"{gtot:.2f}%" if not pd.isna(gtot) else "Not available in the provided files."],
        ["Strongest CLO", f"{strong} – {stats[strong]['mean']:.2f}%" if strong != "Not available" else "Not available"],
        ["Weakest CLO", f"{weak} – {stats[weak]['mean']:.2f}%" if weak != "Not available" else "Not available"],
        [f"CLOs ≥{BENCHMARK:.0f}%", sum(
            stats[c]["mean"] >= BENCHMARK for c in clos
            if not pd.isna(stats[c]["mean"])
        )],
        [f"CLOs <{BENCHMARK:.0f}%", sum(
            stats[c]["mean"] < BENCHMARK for c in clos
            if not pd.isna(stats[c]["mean"])
        )],
        ["Key CQI recommendation",
         f"Prioritize {weak} through CLO-specific intervention."
         if weak != "Not available" else "Not available in the provided files."],
    ])

    doc.add_heading("2. Course Information", 1)
    table(["Field", "Information"], [
        [k, safe_value(v)] for k, v in info.items()
    ])

    doc.add_heading("2.1 Course Description", 2)
    doc.add_paragraph(safe_value(info.get("Course Description")))

    doc.add_heading("2.2 Course Objectives", 2)
    if objectives:
        for i, obj in enumerate(objectives, 1):
            doc.add_paragraph(f"{i}. {obj}")
    else:
        doc.add_paragraph("Not available in the provided files.")

    doc.add_heading("2.3 Official CLOs", 2)
    table(["CLO", "Official CLO"], [[c, clos[c]] for c in clos])

    doc.add_heading("3. CLO–Assessment Alignment", 1)
    table(
        ["CLO", "Official CLO Description", "Assessment/Question",
         "Weightage", "Maximum Marks"],
        [
            [
                a["clo"], clos.get(a["clo"], "Unmatched CLO in Excel"),
                a["label"], a["weightage"],
                "Not available in the provided files."
            ]
            for a in assessments
        ]
    )

    doc.add_heading("4. Methodology", 1)
    doc.add_paragraph(
        "The Course Outline is authoritative for course information and exact "
        "official CLO wording. The Excel workbook is authoritative for "
        "numerical OBE analysis. Missing maximum marks are not inferred. "
        f"Benchmark = {BENCHMARK:.0f}%; ≥80% Strong, 70–79.99% Satisfactory, "
        "<70% Needs Improvement."
    )

    doc.add_heading("5. CLO-wise OBE Attainment", 1)
    table(
        ["CLO", "Exact Official CLO", "Mean Attainment (%)",
         f"Students ≥{BENCHMARK:.0f}%", "Status"],
        [
            [
                c, clos[c],
                f"{stats[c]['mean']:.2f}%" if not pd.isna(stats[c]["mean"]) else "Not available in the provided files.",
                f"{stats[c]['n70']} ({stats[c]['pct70']:.0f}%)"
                if not pd.isna(stats[c]["pct70"]) else "Not available in the provided files.",
                status(stats[c]["mean"])
            ]
            for c in clos
        ]
    )

    doc.add_heading("6. Assessment-wise Analysis", 1)
    table(
        ["CLO", "Assessment/Question", "Weightage",
         "Average/Mean Score", "Attainment %", "Interpretation"],
        [
            [
                a["clo"], a["label"], a["weightage"], a["average"],
                "Not available in the provided files.",
                "Raw maximum marks are not separately available; attainment is not inferred."
            ]
            for a in assessments
        ]
    )

    doc.add_heading("7. Student Performance Analysis", 1)
    table(["Metric", "Result"], [
        ["Number assessed", overall["n"]],
        ["Highest overall score", f"{overall['highest']:.2f}" if not pd.isna(overall["highest"]) else "Not available in the provided files."],
        ["Lowest overall score", f"{overall['lowest']:.2f}" if not pd.isna(overall["lowest"]) else "Not available in the provided files."],
        ["Mean overall score", f"{overall['mean']:.2f}" if not pd.isna(overall["mean"]) else "Not available in the provided files."],
        ["Median", f"{overall['median']:.2f}" if not pd.isna(overall["median"]) else "Not available in the provided files."],
        ["Standard deviation", f"{overall['sd']:.2f}" if not pd.isna(overall["sd"]) else "Not available in the provided files."],
        [f"Meeting {BENCHMARK:.0f}% benchmark",
         f"{overall['benchmark_pct']:.0f}%" if not pd.isna(overall["benchmark_pct"]) else "Not available in the provided files."],
    ])

    doc.add_heading("8. Charts and Visual Evidence", 1)
    captions = [
        f"Figure 1. CLO-wise OBE attainment with the {BENCHMARK:.0f}% benchmark.",
        f"Figure 2. Percentage of students achieving ≥{BENCHMARK:.0f}% for each CLO.",
        "Figure 3. Assessment mean scores as reported in Excel.",
    ]
    for pth, cap in zip(chart_paths, captions):
        doc.add_picture(str(pth), width=__import__("docx").shared.Inches(6.7))
        q = doc.add_paragraph(cap)
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("9. CLO Alignment with Course Outline", 1)
    table(
        ["CLO", "Official CLO", "Assessment Evidence",
         "Attainment %", f"Benchmark Achievement %",
         "Status", "CQI Priority"],
        [
            [
                c,
                clos[c],
                "; ".join(a["label"] for a in assessments if a["clo"] == c)
                or "No assessment evidence for this CLO was identified in the provided Excel file.",
                f"{stats[c]['mean']:.2f}%"
                if not pd.isna(stats[c]["mean"]) else "Not available in the provided files.",
                f"{stats[c]['pct70']:.0f}%"
                if not pd.isna(stats[c]["pct70"]) else "Not available in the provided files.",
                status(stats[c]["mean"]),
                "High" if (not pd.isna(stats[c]["mean"]) and stats[c]["mean"] < BENCHMARK) else "Maintain"
            ]
            for c in clos
        ]
    )

    doc.add_heading("10. Evidence-based OBE Findings", 1)
    if valid_means:
        doc.add_paragraph(
            f"Overall attainment reported by the Excel workbook is "
            f"{gtot:.2f}%." if not pd.isna(gtot) else
            "Overall attainment is not available in the provided files."
        )
        doc.add_paragraph(
            f"{strong} has the highest CLO mean attainment ({stats[strong]['mean']:.2f}%), "
            f"while {weak} has the lowest ({stats[weak]['mean']:.2f}%)."
        )
        doc.add_paragraph(
            "These findings describe the available assessment evidence and do not "
            "constitute unsupported claims about student ability, instructor effectiveness, "
            "or teaching quality."
        )
    else:
        doc.add_paragraph("Not available in the provided files.")

    doc.add_heading("11. CQI Action Plan", 1)
    table(
        ["CLO/Area", "Identified Issue", "Recommended Action",
         "Teaching/Learning Intervention", "Follow-up Evidence", "Target"],
        [
            [
                c,
                f"Mean attainment {stats[c]['mean']:.2f}% is below the benchmark."
                if not pd.isna(stats[c]["mean"]) and stats[c]["mean"] < BENCHMARK
                else "CLO meets or exceeds the benchmark."
                if not pd.isna(stats[c]["mean"])
                else "Not available in the provided files.",
                f"Use the exact official CLO as the intervention focus: {clos[c]}",
                "Targeted practice, formative assessment, guided application and feedback tied to the CLO.",
                "Repeat CLO-aligned evidence and compare attainment and benchmark achievement.",
                f"Mean attainment ≥{BENCHMARK:.0f}%.",
            ]
            for c in clos
        ]
    )

    doc.add_heading("12. Formal Conclusion", 1)
    doc.add_paragraph(
        "The OBE evaluation summarizes the available course-outline and assessment "
        "evidence. CQI priorities should focus on CLOs below the stated benchmark, "
        "while CLOs meeting the benchmark should be maintained and monitored."
    )

    doc.add_heading("13. OBE Quality Check", 1)
    doc.add_paragraph(
        "The generated package uses the same student, CLO, assessment and chart data "
        "across the Word report and Excel workbook. Missing values are explicitly "
        "reported rather than inferred."
    )
    table(["Check", "Result"], [
        ["Students analyzed", overall["n"]],
        ["Overall attainment", f"{gtot:.2f}%" if not pd.isna(gtot) else "Not available in the provided files."],
        ["Strongest CLO", strong],
        ["Weakest CLO", weak],
        [f"CLOs ≥{BENCHMARK:.0f}%", sum(stats[c]["mean"] >= BENCHMARK for c in clos if not pd.isna(stats[c]["mean"]))],
        [f"CLOs <{BENCHMARK:.0f}%", sum(stats[c]["mean"] < BENCHMARK for c in clos if not pd.isna(stats[c]["mean"]))],
        ["Assessments analyzed", len(assessments)],
        ["CLOs matched", sum(1 for c in clos if any(a["clo"] == c for a in assessments))],
        ["Calculations internally consistent", "Yes; tables and charts use the same analysis."],
    ])

    doc.save(out)


# ---------------------------------------------------------
# MASTER PROMPT TEXT
# ---------------------------------------------------------
MASTER_PROMPT = """OBE EVALUATION REPORT – MASTER PROMPT

Analyze the uploaded Course Outline and OBE Excel workbook together.

SOURCE HIERARCHY
Course Outline = authoritative for course information and exact official CLO wording.
Excel = authoritative for all numerical calculations, student performance, marks, assessment results, assessment-to-CLO mapping and attainment.
Never invent, estimate, modify or reinterpret values. Never change official CLO wording. Never infer missing maximum marks. Missing information must be reported as “Not available in the provided files.”

REQUIRED OUTPUT
1. Course Information
2. CLO–Assessment Alignment
3. OBE Analysis
4. Status: ≥80 Strong; 70–79.99 Satisfactory; <70 Needs Improvement.
5. CLO Attainment Table
6. Assessment-to-CLO Mapping
7. Student Performance
8. Charts
9. Evidence-based OBE interpretation
10. CQI Action Plan
11. Dedicated CLO Alignment with Course Outline
12. Formal conclusion
13. Executive Summary
14. Word report
15. Separate Excel workbook
16. DOCX, XLSX and three PNG charts
17. Validate all numerical and CLO consistency
18. Final quality check

If a Course Outline CLO has no Excel evidence, write exactly:
“No assessment evidence for this CLO was identified in the provided Excel file.”

If an Excel assessment CLO cannot be matched to the Course Outline, flag the discrepancy clearly.
"""


# ---------------------------------------------------------
# INTERFACE
# ---------------------------------------------------------
st.markdown('<div class="section-title">1. Course Information</div>', unsafe_allow_html=True)
st.caption(
    "Enter information here when it is not available in the uploaded Course Outline. "
    "Uploaded source information remains authoritative and is not silently overwritten."
)

c1, c2, c3 = st.columns(3)
with c1:
    manual_institution = st.text_input("Institution", value="")
    manual_department = st.text_input("Department", value="")
    manual_program = st.text_input("Program", value="")
with c2:
    manual_title = st.text_input("Course Title", value="")
    manual_code = st.text_input("Course Code", value="")
    manual_semester = st.text_input("Semester", value="")
with c3:
    manual_year = st.text_input("Academic Year", value="")
    manual_instructor = st.text_input("Course Teacher / Instructor", value="")
    manual_credit = st.number_input("Credit Hours", min_value=0.0, value=3.0, step=0.5)

manual_section = st.text_input("Section (optional)", value="")
manual_campus = st.text_input("Campus (optional)", value="")

st.markdown('<div class="section-title">2. Course Learning Outcomes</div>', unsafe_allow_html=True)
st.caption("Upload the Course Outline to extract and display the exact official CLO wording.")

outline = st.file_uploader(
    "Upload Course Outline (.docx)",
    type=["docx"],
    key="outline_upload",
)

st.markdown('<div class="section-title">3. Student Assessment Data</div>', unsafe_allow_html=True)
st.caption(
    "Upload the OBE Excel workbook containing assessment marks, CLO mappings and student-level attainment evidence."
)

excel = st.file_uploader(
    "Upload OBE Assessment Excel (.xlsx)",
    type=["xlsx"],
    key="excel_upload",
)

# Download master prompt
with st.expander("View / download the OBE Master Prompt"):
    st.text_area("Master Prompt", MASTER_PROMPT, height=350)
    st.download_button(
        "Download OBE Master Prompt",
        MASTER_PROMPT,
        file_name="OBE_Evaluation_Master_Prompt.txt",
        mime="text/plain",
    )


# ---------------------------------------------------------
# PROCESS INPUTS
# ---------------------------------------------------------
info = {}
objectives = []
clos = {}

if outline:
    try:
        info, objectives, clos = parse_outline(outline.getvalue())

        # Manual entries fill only blanks.
        manual_map = {
            "Institution": manual_institution,
            "Department": manual_department,
            "Program": manual_program,
            "Course Title": manual_title,
            "Course Code": manual_code,
            "Semester": manual_semester,
            "Academic Year": manual_year,
            "Instructor/Faculty": manual_instructor,
            "Credit Hours": str(manual_credit) if manual_credit else "",
            "Section": manual_section,
            "Campus": manual_campus,
        }
        for key, value in manual_map.items():
            if not info.get(key) and value:
                info[key] = value

        st.success("Course Outline loaded. Official CLO wording extracted from the document.")

        if clos:
            clo_df = pd.DataFrame(
                [{"CLO": c, "CLO Description": clos[c]} for c in clos]
            )
            st.dataframe(clo_df, use_container_width=True, hide_index=True)
        else:
            st.warning(
                "No CLOs were detected in the Course Outline. Please check the document format."
            )

    except Exception as e:
        st.error(f"Could not read the Course Outline: {e}")

elif any([manual_institution, manual_department, manual_program, manual_title, manual_code]):
    info = {
        "Institution": manual_institution,
        "Department": manual_department,
        "Program": manual_program,
        "Course Title": manual_title,
        "Course Code": manual_code,
        "Semester": manual_semester,
        "Academic Year": manual_year,
        "Campus": manual_campus,
        "Instructor/Faculty": manual_instructor,
        "Credit Hours": str(manual_credit),
        "Section": manual_section,
        "Course Description": "",
    }

if excel:
    try:
        raw = pd.read_excel(io.BytesIO(excel.getvalue()), sheet_name="OBE", header=None)
        st.success("OBE Excel workbook loaded successfully.")

        if not clos:
            st.warning(
                "Please upload the Course Outline as well so the app can use the exact official CLO wording."
            )
        else:
            assessments, stats, overall, gtot, rows, pct_cols, total_col = analyze(raw, clos)

            st.markdown('<div class="section-title">4. OBE Analysis Dashboard</div>', unsafe_allow_html=True)

            # Course cards
            cards = [
                ("Institution", safe_value(info.get("Institution"))),
                ("Department", safe_value(info.get("Department"))),
                ("Program", safe_value(info.get("Program"))),
                ("Course Title", safe_value(info.get("Course Title"))),
                ("Course Code", safe_value(info.get("Course Code"))),
                ("Semester", safe_value(info.get("Semester"))),
                ("Academic Year", safe_value(info.get("Academic Year"))),
                ("Course Teacher / Instructor", safe_value(info.get("Instructor/Faculty"))),
                ("Credit Hours", safe_value(info.get("Credit Hours"))),
            ]

            for start in range(0, len(cards), 3):
                cols = st.columns(3)
                for col, (label, value) in zip(cols, cards[start:start + 3]):
                    with col:
                        st.markdown(
                            f'<div class="info-card"><div class="info-label">{label}</div>'
                            f'<div class="info-value">{value}</div></div>',
                            unsafe_allow_html=True,
                        )

            st.markdown("### CLO Attainment")
            result_df = pd.DataFrame([
                {
                    "CLO": c,
                    "Exact Official CLO": clos[c],
                    "Mean Attainment (%)": stats[c]["mean"],
                    f"Students ≥{BENCHMARK:.0f}%": stats[c]["n70"],
                    f"Students ≥{BENCHMARK:.0f}% (%)": stats[c]["pct70"],
                    "SD": stats[c]["sd"],
                    "Status": status(stats[c]["mean"]),
                }
                for c in clos
            ])
            st.dataframe(result_df, use_container_width=True, hide_index=True)

            # Summary metrics
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Students Assessed", overall["n"])
            m2.metric(
                "Overall Attainment",
                f"{gtot:.2f}%" if not pd.isna(gtot) else "N/A"
            )
            valid = {c: stats[c]["mean"] for c in clos if not pd.isna(stats[c]["mean"])}
            if valid:
                strongest = max(valid, key=valid.get)
                weakest = min(valid, key=valid.get)
                m3.metric("Strongest CLO", f"{strongest} ({valid[strongest]:.2f}%)")
                m4.metric("Weakest CLO", f"{weakest} ({valid[weakest]:.2f}%)")

            st.markdown("### Generate Complete OBE Report Package")

            if st.button("Generate Complete OBE Report Package", type="primary", use_container_width=True):
                out = Path("obe_output")
                out.mkdir(exist_ok=True)

                chart_paths = charts(stats, assessments, out)

                student_data = {
                    "Sr.": range(1, overall["n"] + 1),
                }

                # Student identifiers are included only when available in the workbook.
                if raw.shape[1] > 2:
                    student_data["Roll No."] = [
                        clean(raw.iloc[r, 2]) for r in rows[:overall["n"]]
                    ]
                if raw.shape[1] > 3:
                    student_data["Section"] = [
                        clean(raw.iloc[r, 3]) for r in rows[:overall["n"]]
                    ]

                for c in clos:
                    if c in pct_cols and pct_cols[c] < raw.shape[1]:
                        vals = [
                            num(raw.iloc[r, pct_cols[c]])
                            for r in rows[:overall["n"]]
                        ]
                        student_data[c + " Attainment %"] = vals

                if total_col < raw.shape[1]:
                    student_data["Overall Score"] = [
                        num(raw.iloc[r, total_col])
                        for r in rows[:overall["n"]]
                    ]

                sdf = pd.DataFrame(student_data)

                xlsx_path = out / "OBE_Analysis.xlsx"
                docx_path = out / "OBE_Evaluation_Report.docx"

                workbook(
                    clos, assessments, stats, overall, sdf, xlsx_path
                )
                report(
                    info, objectives, clos, assessments, stats,
                    overall, gtot, sdf, chart_paths, docx_path
                )

                st.success(
                    "Complete package generated: DOCX + XLSX + three separate PNG charts."
                )

                st.markdown("### Download Outputs")

                st.download_button(
                    "📄 Download Word OBE Evaluation Report",
                    docx_path.read_bytes(),
                    file_name=docx_path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

                st.download_button(
                    "📊 Download Excel OBE Analysis Workbook",
                    xlsx_path.read_bytes(),
                    file_name=xlsx_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                for pth in chart_paths:
                    st.download_button(
                        f"📈 Download {pth.name}",
                        pth.read_bytes(),
                        file_name=pth.name,
                        mime="image/png",
                        use_container_width=True,
                    )

                st.markdown("### Chart Preview")
                for pth in chart_paths:
                    st.image(str(pth), caption=pth.stem, use_container_width=True)

    except Exception as e:
        st.error(
            f"Could not analyze the Excel workbook. Please make sure the workbook "
            f"contains the expected 'OBE' sheet and assessment/CLO data. Error: {e}"
        )

else:
    st.info(
        "Start by entering course information and uploading the Course Outline "
        "and OBE Excel workbook. The complete OBE analysis becomes available after upload."
    )
