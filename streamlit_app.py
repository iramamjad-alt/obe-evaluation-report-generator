import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import matplotlib.pyplot as plt

st.set_page_config(page_title="OBE Evaluation Report Generator", page_icon="📊", layout="wide")

st.markdown("""
<style>
.main-title{font-size:42px;font-weight:800;color:#17365D;margin-bottom:0}
.subtitle{color:#6B7280;font-size:17px;margin-bottom:25px}
.section-title{color:#17365D;font-size:30px;font-weight:750;margin-top:25px}
.info-card{background:#F4F7FB;border:1px solid #D9E2F3;border-radius:12px;padding:14px 18px;min-height:90px}
.info-label{color:#64748B;font-size:14px}.info-value{color:#172033;font-size:17px;font-weight:600}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 OBE Evaluation Report Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Generate an auditable CLO attainment analysis, Word report, Excel workbook and separate chart files from your Course Outline and OBE assessment data.</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Report Settings")
    benchmark = st.number_input("OBE benchmark (%)", 0.0, 100.0, 70.0, 1.0)
    st.caption("Benchmark used for CLO/student achievement calculations.")
    st.divider()
    st.subheader("Input options")
    st.caption("Course details may be entered manually when they are not available in the Course Outline.")

BENCHMARK = benchmark
MISSING = "Not available in the provided files."
NO_EVIDENCE = "No assessment evidence for this CLO was identified in the provided Excel file."


def clean(v):
    if pd.isna(v): return ""
    return str(v).replace("\ufeff", "").strip()


def num(v):
    try: return float(v)
    except Exception: return np.nan


def status(x):
    if pd.isna(x): return MISSING
    if x >= 80: return "Strong"
    if x >= 70: return "Satisfactory"
    return "Needs Improvement"


def safe_value(v): return v if clean(v) else MISSING


def normalize_label(x):
    return re.sub(r"[^a-z0-9]+", "", clean(x).lower())


def detect_clo_ids(values):
    found = []
    for v in values:
        s = clean(v)
        for m in re.finditer(r"\bCLO\s*[-_]?\s*(\d+)\b", s, re.I):
            c = f"CLO{int(m.group(1))}"
            if c not in found: found.append(c)
    return found


def parse_outline(data):
    d = Document(io.BytesIO(data))
    lines = [clean(p.text) for p in d.paragraphs if clean(p.text)]
    text = "\n".join(lines)
    info = {k:"" for k in ["Institution","Department","Program","Course Title","Course Code","Semester","Academic Year","Campus","Instructor/Faculty","Credit Hours","Section","Course Description"]}
    patterns = {
        "Course Title": r"(?:Course|Course Title):\s*([^\n]+)", "Course Code": r"(?:Course code|Course Code):\s*([^\n]+)",
        "Semester": r"(?:Year/Semester|Semester):\s*([^\n]+)", "Program": r"Program:\s*([^\n]+)",
        "Credit Hours": r"(?:Units/Cr Hrs\.|Credit Hours|Cr Hrs\.):\s*([^\n]+)", "Instructor/Faculty": r"(?:Instructor|Teacher|Faculty):\s*([^\n]*)",
        "Academic Year": r"Academic Year:\s*([^\n]+)", "Department": r"Department:\s*([^\n]+)", "Institution": r"Institution:\s*([^\n]+)",
        "Campus": r"Campus:\s*([^\n]+)", "Section": r"Section:\s*([^\n]+)"
    }
    for key, pattern in patterns.items():
        m = re.search(pattern, text, re.I)
        if m: info[key] = clean(m.group(1))
    if not info["Campus"] and "Lahore Campus" in text: info["Campus"] = "Lahore Campus"
    m = re.search(r"COURSE DESCRIPTION\s*(.*?)(?:Program Educational Objectives|Course Objectives|Course Learning Outcomes|CLO\s*1)", text, re.I|re.S)
    if m: info["Course Description"] = " ".join(m.group(1).split())

    clos = {}
    for table in d.tables:
        for row in table.rows:
            vals = [clean(c.text) for c in row.cells]
            if vals:
                ids = detect_clo_ids(vals[:1])
                if ids and len(vals)>1 and clean(vals[1]): clos[ids[0]] = vals[1]
    for i,line in enumerate(lines):
        m = re.match(r"^(CLO\s*\d+)\s*[:\-]?\s*(.*)$", line, re.I)
        if m:
            c = f"CLO{int(re.search(r'\d+',m.group(1)).group())}"
            desc = m.group(2).strip() or (lines[i+1] if i+1<len(lines) else "")
            if desc: clos[c] = desc
    objectives=[]
    m = re.search(r"Course Objectives\s*(.*?)(?:Program Learning Outcome|Course Learning Outcomes|CLO\s*1)", text, re.I|re.S)
    if m:
        block=m.group(1)
        for line in block.splitlines():
            line=re.sub(r"^\s*\d+[.)]\s*", "", line).strip()
            if line: objectives.append(line)
    return info, objectives, dict(sorted(clos.items(), key=lambda x:int(re.search(r'\d+',x[0]).group())))


def read_excel_sheets(data):
    return pd.read_excel(io.BytesIO(data), sheet_name=None, header=None)


def find_header_row(df):
    best=(-1,0)
    for r in range(min(15,len(df))):
        vals=[clean(x) for x in df.iloc[r].tolist()]
        score=sum(bool(re.search(r"CLO\s*\d+|student|roll|id|assessment|marks|score|attainment|question", v, re.I)) for v in vals)
        if score>best[1]: best=(r,score)
    return best[0] if best[1]>=1 else 0


def to_table(df):
    if df.empty: return pd.DataFrame()
    hr=find_header_row(df)
    headers=[]
    seen={}
    for i,v in enumerate(df.iloc[hr].tolist()):
        h=clean(v) or f"Column_{i+1}"
        seen[h]=seen.get(h,0)+1
        headers.append(h if seen[h]==1 else f"{h}_{seen[h]}")
    out=df.iloc[hr+1:].copy()
    out.columns=headers
    out=out.dropna(how="all").reset_index(drop=True)
    return out


def looks_like_percent(header):
    s=clean(header).lower()
    return "%" in s or "percent" in s or "percentage" in s or "attainment" in s or "achievement" in s


def find_student_table(sheets):
    candidates=[]
    for name, raw in sheets.items():
        t=to_table(raw)
        if t.empty: continue
        cols=list(t.columns)
        clo_cols=[c for c in cols if detect_clo_ids([c])]
        numeric=sum(pd.to_numeric(t[c],errors="coerce").notna().sum() for c in cols)
        id_cols=[c for c in cols if re.search(r"student|roll|registration|reg\.?\s*no|id|name", clean(c), re.I)]
        score=len(clo_cols)*10 + len(id_cols)*4 + min(numeric,100)/100
        if clo_cols: candidates.append((score,name,t,clo_cols,id_cols))
    if not candidates: return None,None,None,None
    _,name,t,clo_cols,id_cols=max(candidates,key=lambda x:x[0])
    return name,t,clo_cols,id_cols


def analyze_dynamic(sheets, clos):
    name, table, clo_cols, id_cols = find_student_table(sheets)
    if table is None:
        raise ValueError("No student-level Excel table with CLO-labelled columns was detected. CLO columns should contain labels such as CLO1, CLO 1, CLO-1, etc.")

    # Match Excel CLO columns to official Course Outline CLOs.
    excel_clos=detect_clo_ids(clo_cols)
    for c in excel_clos:
        if c not in clos:
            clos[c] = "Excel assessment CLO has no matching Course Outline CLO."
    ordered=list(clos.keys())
    assessments=[]
    stats={}
    student_data={}
    for c in ordered:
        matching=[col for col in clo_cols if c in detect_clo_ids([col])]
        vals=[]
        for col in matching:
            s=pd.to_numeric(table[col],errors="coerce")
            if s.notna().any():
                vals.append(s)
                assessments.append({"clo":c,"label":clean(col),"average":s.mean(),"maximum":MISSING,"attainment":s.mean() if looks_like_percent(col) else np.nan,"source":name})
        if vals:
            combined=pd.concat(vals,axis=1).mean(axis=1,skipna=True).dropna()
            is_pct=all(looks_like_percent(col) for col in matching)
            stats[c]={"n":len(combined),"mean":combined.mean() if is_pct else np.nan,"raw_mean":combined.mean(),"sd":combined.std(ddof=1) if len(combined)>1 else np.nan,"n70":int((combined>=BENCHMARK).sum()) if is_pct else 0,"pct70":(combined>=BENCHMARK).mean()*100 if is_pct else np.nan,"evidence":matching,"is_percent":is_pct}
            student_data[c+" Attainment %" if is_pct else c+" Score"] = combined.reindex(table.index).tolist()
        else:
            stats[c]={"n":0,"mean":np.nan,"raw_mean":np.nan,"sd":np.nan,"n70":0,"pct70":np.nan,"evidence":[] ,"is_percent":False}

    # Identify an overall percentage column only when its label clearly indicates percentage/attainment.
    overall_col=next((c for c in table.columns if re.search(r"overall|total|final|aggregate|course",clean(c),re.I) and looks_like_percent(c)),None)
    overall_series=pd.to_numeric(table[overall_col],errors="coerce").dropna() if overall_col else pd.Series(dtype=float)
    overall={"n":len(overall_series),"highest":overall_series.max() if len(overall_series) else np.nan,"lowest":overall_series.min() if len(overall_series) else np.nan,"mean":overall_series.mean() if len(overall_series) else np.nan,"median":overall_series.median() if len(overall_series) else np.nan,"sd":overall_series.std(ddof=1) if len(overall_series)>1 else np.nan,"benchmark_pct":(overall_series>=BENCHMARK).mean()*100 if len(overall_series) else np.nan}
    if id_cols:
        for c in id_cols[:2]: student_data[clean(c)]=table[c].tolist()
    student_df=pd.DataFrame(student_data)
    return assessments,stats,overall,student_df,name,table


def charts(stats, assessments, out):
    out=Path(out); out.mkdir(exist_ok=True); paths=[]; cs=list(stats)
    vals=[stats[c]["mean"] for c in cs]
    p=out/"CLO_Attainment_Chart.png"; fig,ax=plt.subplots(figsize=(9,5)); ax.bar(cs,[0 if pd.isna(v) else v for v in vals]); ax.axhline(BENCHMARK,ls="--",lw=1.5,label=f"{BENCHMARK:.0f}% benchmark"); ax.set_title("Figure 1. CLO-wise OBE Attainment"); ax.set_xlabel("Course Learning Outcome"); ax.set_ylabel("Mean Attainment (%)"); ax.set_ylim(0,100); ax.legend()
    for i,v in enumerate(vals):
        if not pd.isna(v): ax.text(i,min(v+2,97),f"{v:.2f}%",ha="center")
    fig.tight_layout();fig.savefig(p,dpi=200,bbox_inches="tight");plt.close(fig);paths.append(p)
    p=out/"Benchmark_Achievement_Chart.png"; vals=[stats[c]["pct70"] for c in cs]; fig,ax=plt.subplots(figsize=(9,5));ax.bar(cs,[0 if pd.isna(v) else v for v in vals]);ax.set_title(f"Figure 2. Students Achieving ≥{BENCHMARK:.0f}% by CLO");ax.set_xlabel("Course Learning Outcome");ax.set_ylabel("Students achieving benchmark (%)");ax.set_ylim(0,100)
    for i,v in enumerate(vals):
        if not pd.isna(v): ax.text(i,min(v+2,97),f"{v:.0f}%",ha="center")
    fig.tight_layout();fig.savefig(p,dpi=200,bbox_inches="tight");plt.close(fig);paths.append(p)
    p=out/"Assessment_Performance_Chart.png"; valid=[a for a in assessments if not pd.isna(a["average"])]; vals=[a["average"] for a in valid]; labels=[f"{a['clo']}\n{a['label']}" for a in valid]; fig,ax=plt.subplots(figsize=(12,6));ax.bar(range(len(vals)),vals);ax.set_title("Figure 3. Assessment Mean Scores as Reported in Excel");ax.set_xlabel("Assessment / CLO");ax.set_ylabel("Mean score / attainment as reported");ax.set_xticks(range(len(vals)));ax.set_xticklabels(labels,rotation=55,ha="right",fontsize=8);fig.tight_layout();fig.savefig(p,dpi=200,bbox_inches="tight");plt.close(fig);paths.append(p)
    return paths


def make_xlsx(clos,assessments,stats,overall,student_df,out):
    wb=Workbook();wb.remove(wb.active); title=PatternFill("solid",fgColor="1F4E78");head=PatternFill("solid",fgColor="D9EAF7")
    def setup(ws,title_text,headers):
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(headers));ws.cell(1,1,title_text).fill=title;ws.cell(1,1).font=Font(color="FFFFFF",bold=True,size=12)
        for j,h in enumerate(headers,1): ws.cell(3,j,h).fill=head;ws.cell(3,j).font=Font(bold=True);ws.cell(3,j).alignment=Alignment(wrap_text=True)
        ws.freeze_panes="A4"
    def put(ws,rows):
        for r,row in enumerate(rows,4):
            for j,v in enumerate(row,1): ws.cell(r,j,v)
    ws=wb.create_sheet("OBE Summary");setup(ws,"OBE Summary",["CLO","Official CLO Description","Mean Attainment (%)",f"Students ≥{BENCHMARK:.0f}%",f"Students ≥{BENCHMARK:.0f}% (%)","SD","Status"]);put(ws,[[c,clos[c],stats[c]["mean"],stats[c]["n70"],stats[c]["pct70"],stats[c]["sd"],status(stats[c]["mean"])] for c in clos])
    ws=wb.create_sheet("CLO–Assessment Mapping");setup(ws,"CLO–Assessment Mapping",["CLO","Official CLO Description","Assessment/Question","Maximum Marks","Average/Mean","Attainment %","Source"]);put(ws,[[a["clo"],clos.get(a["clo"],"Excel assessment CLO has no matching Course Outline CLO."),a["label"],a["maximum"],a["average"],a["attainment"] if not pd.isna(a["attainment"]) else MISSING,a["source"]] for a in assessments])
    ws=wb.create_sheet("Assessment Analysis");setup(ws,"Assessment Analysis",["CLO","Assessment/Question","Average/Mean","Attainment %","Interpretation"]);put(ws,[[a["clo"],a["label"],a["average"],a["attainment"] if not pd.isna(a["attainment"]) else MISSING,"Attainment is reported only where the Excel column is explicitly percentage/attainment based."] for a in assessments])
    ws=wb.create_sheet("Student/CLO Data");setup(ws,"Student/CLO Data",list(student_df.columns));put(ws,student_df.values.tolist())
    ws=wb.create_sheet("CQI Action Plan");setup(ws,"CQI Action Plan",["CLO/Area","Issue","Action","Intervention","Follow-up Evidence","Target"]);put(ws,[[c,(f"Mean attainment {stats[c]['mean']:.2f}% is below benchmark." if not pd.isna(stats[c]['mean']) and stats[c]['mean']<BENCHMARK else "CLO meets/exceeds benchmark." if not pd.isna(stats[c]['mean']) else MISSING),f"Address the exact official CLO: {clos[c]}","Use CLO-aligned practice, formative assessment and feedback.","Repeat CLO-aligned assessment and compare results.",f"Mean attainment ≥{BENCHMARK:.0f}%."] for c in clos])
    ws=wb.create_sheet("Chart Data");setup(ws,"Chart Data",["CLO","Mean Attainment (%)",f"Benchmark Achievement (%)","Status"]);put(ws,[[c,stats[c]["mean"],stats[c]["pct70"],status(stats[c]["mean"])] for c in clos]);wb.save(out)


def make_docx(info,objectives,clos,assessments,stats,overall,student_df,chart_paths,out):
    doc=Document(); sec=doc.sections[0];sec.top_margin=sec.bottom_margin=Inches(.65);sec.left_margin=sec.right_margin=Inches(.75)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.add_run("OUTCOME-BASED EDUCATION (OBE)\nEVALUATION REPORT\n").bold=True;p.add_run(f"\n{safe_value(info.get('Course Title'))} ({safe_value(info.get('Course Code'))})\nProgram: {safe_value(info.get('Program'))}\nSection: {safe_value(info.get('Section'))}\nSemester: {safe_value(info.get('Semester'))}\n")
    doc.add_page_break()
    def table(headers,rows):
        t=doc.add_table(rows=1,cols=len(headers));t.style="Table Grid"
        for i,h in enumerate(headers):t.rows[0].cells[i].text=str(h)
        for row in rows:
            cells=t.add_row().cells
            for i,v in enumerate(row):cells[i].text=str(v)
    valid={c:stats[c]["mean"] for c in clos if not pd.isna(stats[c]["mean"])}; strongest=max(valid,key=valid.get) if valid else MISSING;weakest=min(valid,key=valid.get) if valid else MISSING
    doc.add_heading("1. Executive Summary",1);table(["Item","Result"],[["Course",f"{safe_value(info.get('Course Title'))} ({safe_value(info.get('Course Code'))})"],["Number of students",overall["n"]],["Strongest CLO",f"{strongest} – {valid[strongest]:.2f}%" if strongest!=MISSING else MISSING],["Weakest CLO",f"{weakest} – {valid[weakest]:.2f}%" if weakest!=MISSING else MISSING]])
    doc.add_heading("2. Course Information",1);table(["Field","Information"],[[k,safe_value(v)] for k,v in info.items()]);doc.add_heading("2.1 Course Description",2);doc.add_paragraph(safe_value(info.get("Course Description")));doc.add_heading("2.2 Course Objectives",2)
    for i,x in enumerate(objectives,1):doc.add_paragraph(f"{i}. {x}")
    if not objectives:doc.add_paragraph(MISSING)
    doc.add_heading("2.3 Official CLOs",2);table(["CLO","Official CLO"],[[c,clos[c]] for c in clos])
    doc.add_heading("3. CLO–Assessment Alignment",1);table(["CLO","Official CLO","Assessment/Question","Maximum Marks","Attainment %"],[[a["clo"],clos.get(a["clo"],"Excel assessment CLO has no matching Course Outline CLO."),a["label"],MISSING,a["attainment"] if not pd.isna(a["attainment"]) else MISSING] for a in assessments])
    doc.add_heading("4. OBE Attainment",1);table(["CLO","Exact Official CLO","Mean Attainment (%)",f"Students ≥{BENCHMARK:.0f}%","Status"],[[c,clos[c],f"{stats[c]['mean']:.2f}%" if not pd.isna(stats[c]['mean']) else MISSING,f"{stats[c]['n70']} ({stats[c]['pct70']:.0f}%)" if not pd.isna(stats[c]['pct70']) else MISSING,status(stats[c]['mean'])] for c in clos])
    doc.add_heading("5. Student Performance",1);table(["Metric","Result"],[["Number assessed",overall["n"]],["Highest",overall["highest"] if not pd.isna(overall["highest"]) else MISSING],["Lowest",overall["lowest"] if not pd.isna(overall["lowest"]) else MISSING],["Mean",overall["mean"] if not pd.isna(overall["mean"]) else MISSING],["Median",overall["median"] if not pd.isna(overall["median"]) else MISSING],["SD",overall["sd"] if not pd.isna(overall["sd"]) else MISSING],[f"Meeting {BENCHMARK:.0f}%",f"{overall['benchmark_pct']:.0f}%" if not pd.isna(overall["benchmark_pct"]) else MISSING]])
    doc.add_heading("6. Charts",1)
    for i,pth in enumerate(chart_paths,1):doc.add_picture(str(pth),width=Inches(6.7));q=doc.add_paragraph(f"Figure {i}. {pth.stem.replace('_',' ')}.");q.alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading("7. CLO Alignment with Course Outline",1);table(["CLO","Official CLO","Assessment Evidence","Attainment %","Benchmark Achievement %","Status","CQI Priority"],[[c,clos[c],"; ".join(stats[c]["evidence"]) if stats[c]["evidence"] else NO_EVIDENCE,f"{stats[c]['mean']:.2f}%" if not pd.isna(stats[c]['mean']) else MISSING,f"{stats[c]['pct70']:.0f}%" if not pd.isna(stats[c]['pct70']) else MISSING,status(stats[c]['mean']),"High" if not pd.isna(stats[c]['mean']) and stats[c]['mean']<BENCHMARK else "Maintain"] for c in clos])
    doc.add_heading("8. Evidence-based Findings",1);doc.add_paragraph("The findings summarize the assessment evidence detected in the uploaded files. They do not by themselves establish claims about student ability, instructor effectiveness or teaching quality.")
    doc.add_heading("9. CQI Action Plan",1);table(["CLO/Area","Issue","Action","Intervention","Follow-up Evidence","Target"],[[c,(f"Mean attainment {stats[c]['mean']:.2f}% is below benchmark." if not pd.isna(stats[c]['mean']) and stats[c]['mean']<BENCHMARK else "CLO meets/exceeds benchmark." if not pd.isna(stats[c]['mean']) else MISSING),f"Address the exact official CLO: {clos[c]}","Use CLO-aligned practice, formative assessment and feedback.","Repeat CLO-aligned assessment and compare results.",f"Mean attainment ≥{BENCHMARK:.0f}%."] for c in clos])
    doc.add_heading("10. Formal Conclusion",1);doc.add_paragraph("This OBE evaluation presents the course-outline and assessment evidence detected from the uploaded files and identifies CLO-level areas for continued monitoring or CQI action.")
    doc.add_heading("11. Quality Check",1);table(["Check","Result"],[["Students analyzed",overall["n"]],["CLOs in Course Outline",sum(1 for c in clos if not c.startswith("CLO") or c in clos)],["Assessments detected",len(assessments)],["CLOs with assessment evidence",sum(bool(stats[c]["evidence"]) for c in clos)],["Excel CLOs without Course Outline match",", ".join(c for c in detect_clo_ids([a['clo'] for a in assessments]) if c not in clos) or "None detected"]])
    doc.save(out)

MASTER_PROMPT="""OBE EVALUATION REPORT – MASTER PROMPT\n\nAnalyze the uploaded Course Outline and OBE Excel workbook together.\n\nSOURCE HIERARCHY\nCourse Outline = authoritative for course information and exact official CLO wording.\nExcel = authoritative for numerical calculations, student performance, marks, assessment results, assessment-to-CLO mapping and attainment.\n\nREQUIRED OUTPUT\n1. Course Information\n2. CLO–Assessment Alignment\n3. OBE Analysis\n4. Status classification\n5. CLO Attainment Table\n6. Assessment-to-CLO Mapping\n7. Student Performance\n8. Charts\n9. Evidence-based OBE interpretation\n10. CQI Action Plan\n11. CLO Alignment with Course Outline\n12. Formal conclusion\n13. Executive Summary\n14. Word report\n15. Separate Excel workbook\n16. DOCX, XLSX and three PNG charts\n17. Validate student count, mappings, marks, means, attainment, benchmark percentages, exact CLO wording and consistency\n18. Final quality check\n\nIf a Course Outline CLO has no Excel evidence, write exactly:\n“No assessment evidence for this CLO was identified in the provided Excel file.”\n\nIf an Excel assessment CLO cannot be matched to the Course Outline, flag the discrepancy clearly.\n"""

st.markdown('<div class="section-title">1. Course Information</div>',unsafe_allow_html=True)
st.caption("Enter course information directly, or leave fields blank so the uploaded Course Outline can supply it.")
c1,c2,c3=st.columns(3)
with c1:
    manual_institution=st.text_input("Institution");manual_department=st.text_input("Department");manual_program=st.text_input("Program")
with c2:
    manual_title=st.text_input("Course Title");manual_code=st.text_input("Course Code");manual_semester=st.text_input("Semester")
with c3:
    manual_year=st.text_input("Academic Year");manual_instructor=st.text_input("Course Teacher / Instructor");manual_credit=st.number_input("Credit Hours",0.0,20.0,3.0,.5)
manual_section=st.text_input("Section (optional)");manual_campus=st.text_input("Campus (optional)")

st.markdown('<div class="section-title">2. Course Learning Outcomes</div>',unsafe_allow_html=True)
outline=st.file_uploader("Upload Course Outline (.docx)",type=["docx"],key="outline_upload")

st.markdown('<div class="section-title">3. Student Assessment Data</div>',unsafe_allow_html=True)
excel=st.file_uploader("Upload OBE Assessment Excel (.xlsx)",type=["xlsx"],key="excel_upload")

with st.expander("View / download the OBE Master Prompt"):
    st.text_area("Master Prompt",MASTER_PROMPT,height=300)
    st.download_button("Download OBE Master Prompt",MASTER_PROMPT,file_name="OBE_Evaluation_Master_Prompt.txt",mime="text/plain")

info={};objectives=[];clos={}
if outline:
    try:
        info,objectives,clos=parse_outline(outline.getvalue())
        manual_map={"Institution":manual_institution,"Department":manual_department,"Program":manual_program,"Course Title":manual_title,"Course Code":manual_code,"Semester":manual_semester,"Academic Year":manual_year,"Instructor/Faculty":manual_instructor,"Credit Hours":str(manual_credit) if manual_credit else "","Section":manual_section,"Campus":manual_campus}
        for k,v in manual_map.items():
            if not info.get(k) and v: info[k]=v
        st.success("Course Outline loaded. Exact official CLO wording extracted where detected.")
        if clos: st.dataframe(pd.DataFrame([{"CLO":c,"CLO Description":d} for c,d in clos.items()]),use_container_width=True,hide_index=True)
        else: st.warning("No CLOs were detected in the Course Outline. You may still enter CLOs manually below.")
    except Exception as e: st.error(f"Could not read the Course Outline: {e}")
else:
    info={"Institution":manual_institution,"Department":manual_department,"Program":manual_program,"Course Title":manual_title,"Course Code":manual_code,"Semester":manual_semester,"Academic Year":manual_year,"Campus":manual_campus,"Instructor/Faculty":manual_instructor,"Credit Hours":str(manual_credit),"Section":manual_section,"Course Description":""}

if not clos:
    st.markdown("#### Optional manual CLO entry")
    st.caption("Use this only when a Course Outline is unavailable or its CLO table cannot be read. Enter the official wording exactly as approved.")
    n_clo=st.number_input("Number of CLOs",1,20,3,1)
    for i in range(1,n_clo+1): clos[f"CLO{i}"]=st.text_area(f"CLO {i} official wording",key=f"manual_clo_{i}")
    clos={k:v for k,v in clos.items() if clean(v)}

if excel and clos:
    try:
        sheets=read_excel_sheets(excel.getvalue())
        assessments,stats,overall,student_df,sheet_name,student_table=analyze_dynamic(sheets,clos)
        st.success(f"Excel loaded. Dynamic student/CLO table detected on sheet: {sheet_name}")
        st.markdown('<div class="section-title">4. OBE Analysis Dashboard</div>',unsafe_allow_html=True)
        cards=[("Institution",safe_value(info.get("Institution")),),("Department",safe_value(info.get("Department"))), ("Program",safe_value(info.get("Program"))), ("Course Title",safe_value(info.get("Course Title"))), ("Course Code",safe_value(info.get("Course Code"))), ("Semester",safe_value(info.get("Semester"))), ("Academic Year",safe_value(info.get("Academic Year"))), ("Course Teacher / Instructor",safe_value(info.get("Instructor/Faculty"))), ("Credit Hours",safe_value(info.get("Credit Hours")))]
        for start in range(0,len(cards),3):
            cols=st.columns(3)
            for col,(label,value) in zip(cols,cards[start:start+3]):
                with col: st.markdown(f'<div class="info-card"><div class="info-label">{label}</div><div class="info-value">{value}</div></div>',unsafe_allow_html=True)
        st.markdown("### CLO Attainment")
        result_df=pd.DataFrame([{"CLO":c,"Exact Official CLO":clos[c],"Mean Attainment (%)":stats[c]["mean"],f"Students ≥{BENCHMARK:.0f}%":stats[c]["n70"],f"Students ≥{BENCHMARK:.0f}% (%)":stats[c]["pct70"],"SD":stats[c]["sd"],"Status":status(stats[c]["mean"])} for c in clos])
        st.dataframe(result_df,use_container_width=True,hide_index=True)
        m1,m2,m3,m4=st.columns(4);m1.metric("Students Assessed",max(stats[c]["n"] for c in stats) if stats else 0);m2.metric("Overall %",f"{overall['mean']:.2f}%" if not pd.isna(overall['mean']) else "N/A")
        valid={c:stats[c]["mean"] for c in clos if not pd.isna(stats[c]["mean"])}
        if valid:
            s=max(valid,key=valid.get);w=min(valid,key=valid.get);m3.metric("Strongest CLO",f"{s} ({valid[s]:.2f}%)");m4.metric("Weakest CLO",f"{w} ({valid[w]:.2f}%)")
        st.markdown("### Detected Assessment Evidence")
        st.dataframe(pd.DataFrame(assessments),use_container_width=True,hide_index=True)
        if st.button("Generate Complete OBE Report Package",type="primary",use_container_width=True):
            out=Path("obe_output");out.mkdir(exist_ok=True)
            chart_paths=charts(stats,assessments,out)
            xlsx_path=out/"OBE_Analysis.xlsx"
            docx_path=out/"OBE_Evaluation_Report.docx"
            make_xlsx(clos,assessments,stats,overall,student_df,xlsx_path)
            make_docx(info,objectives,clos,assessments,stats,overall,student_df,chart_paths,docx_path)
            st.success("Complete package generated: DOCX + XLSX + three separate PNG charts.")
            st.download_button("📄 Download Word OBE Evaluation Report",docx_path.read_bytes(),docx_path.name,"application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
            st.download_button("📊 Download Excel OBE Analysis Workbook",xlsx_path.read_bytes(),xlsx_path.name,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            for p in chart_paths: st.download_button(f"📈 Download {p.name}",p.read_bytes(),p.name,"image/png",use_container_width=True)
            st.markdown("### Chart Preview")
            for p in chart_paths: st.image(str(p),caption=p.stem,use_container_width=True)
    except Exception as e:
        st.error(f"Could not analyze the uploaded Excel workbook dynamically: {e}")
elif excel and not clos:
    st.warning("Please upload the Course Outline or enter the official CLOs manually before analysis.")
else:
    st.info("Enter course information and upload the Course Outline and OBE Excel workbook. The app will detect CLO and assessment columns dynamically.")
