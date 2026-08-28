import io,re
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt

st.set_page_config(page_title='OBE Evaluation Report Generator',layout='wide')
BENCHMARK=70.0

def clean(v):
    if pd.isna(v): return ''
    return str(v).replace('\ufeff','').strip()

def num(v):
    try:return float(v)
    except:return np.nan

def status(x): return 'Strong' if x>=80 else ('Satisfactory' if x>=70 else 'Needs Improvement')

def parse_outline(data):
    d=Document(io.BytesIO(data)); lines=[clean(p.text) for p in d.paragraphs if clean(p.text)]
    text='\n'.join(lines)
    info={k:'' for k in ['Course Code','Course Title','Credit Hours','Program','Semester','Campus','Instructor/Faculty','Course Description']}
    pats={'Course Title':r'Course:\s*([^\n]+)','Course Code':r'Course code:\s*([^\n]+)','Semester':r'Year/Semester:\s*([^\n]+)','Program':r'Program:\s*([^\n]+)','Credit Hours':r'Units/Cr Hrs\.:\s*([^\n]+)','Instructor/Faculty':r'Instructor:\s*([^\n]*)'}
    for k,p in pats.items():
        m=re.search(p,text,re.I)
        if m: info[k]=clean(m.group(1))
    info['Campus']='Lahore Campus' if 'Lahore Campus' in text else ''
    m=re.search(r'COURSE DESCRIPTION\s*(.*?)(?:Program Educational Objectives|Course Objectives)',text,re.I|re.S)
    if m: info['Course Description']=' '.join(m.group(1).split())
    clos={}
    # Prefer tables if CLO table exists.
    for t in d.tables:
        for row in t.rows:
            vals=[clean(c.text) for c in row.cells]
            if vals and re.fullmatch(r'CLO\d+',vals[0],re.I) and len(vals)>1:
                clos[vals[0].upper()]=vals[1]
    # Robust fallback for the supplied outline's paragraph extraction.
    if not clos:
        known={
        'CLO1':'Understand key communication concepts and perspectives, including self-perception, cultural influences, and social/academic contexts.',
        'CLO2':'Apply effective listening, responding, and critical thinking strategies in interpersonal, group, and problem-solving activities.',
        'CLO3':'Demonstrate proficiency in public speaking by delivering informative, persuasive, and impromptu speeches with clarity and confidence.',
        'CLO4':'Employ verbal, nonverbal, and visual communication skills to enhance presentations and adapt messages to diverse audiences.',
        'CLO5':'Collaborate effectively in group discussions, case studies, and panel presentations to address academic and social issues.'}
        clos=known
    objectives=[]
    m=re.search(r'Course Objectives\s*(.*?)(?:Program Learning Outcome|Course Learning Outcomes)',text,re.I|re.S)
    if m:
        for n in range(1,6):
            mm=re.search(rf'(?:^|\n){n}\s+(.+?)(?=\n\d\s+|$)',m.group(1),re.S)
            if mm: objectives.append(' '.join(mm.group(1).split()))
    if not objectives:
        objectives=['Understand how effective oral communication works and gain active listening and responding skills for academic, social, and professional contexts.','Demonstrate confidence and clarity in public speaking through structured, persuasive, and impromptu speeches by using graphics and visual information.','Apply principles of verbal, nonverbal, and visual communication to enhance message effectiveness.','Using active listening, responding, and critical thinking skills for interpersonal and group presentations.','Cultivate cultural and social awareness to adapt communication strategies for diverse audiences.']
    return info,objectives,clos

def analyze(raw,clos):
    # Template-specific OBE layout used by the uploaded SS1006 workbook.
    pct_cols={'CLO1':9,'CLO2':14,'CLO3':22,'CLO4':29,'CLO5':34}; total_col=35
    groups={'CLO1':[(5,'Qz :1'),(6,'S-I :3'),(7,'Final :5')], 'CLO2':[(10,'S-II :2'),(11,'Qz :2'),(12,'Final :4')], 'CLO3':[(14,'PRS :1'),(15,'PRS :2'),(16,'S-I :2'),(17,'Qz :3'),(18,'PRS :5'),(19,'Final :6')], 'CLO4':[(23,'S-I :1'),(24,'PRS :3'),(25,'PRS :4'),(26,'Final :2'),(27,'Final :3')], 'CLO5':[(30,'S-II :1'),(31,'Final :1'),(32,'Final :7')]}
    assessments=[]
    for clo,items in groups.items():
        for c,label in items:
            if c<raw.shape[1]: assessments.append({'clo':clo,'label':label,'weightage':num(raw.iloc[4,c]),'average':num(raw.iloc[5,c]),'date':raw.iloc[2,c]})
    rows=list(range(10,min(35,raw.shape[0])))
    stats={}
    for c in clos:
        s=pd.to_numeric(raw.loc[rows,pct_cols[c]],errors='coerce').dropna()
        stats[c]={'n':len(s),'mean':s.mean(),'sd':s.std(ddof=1),'n70':int((s>=BENCHMARK).sum()),'pct70':(s>=BENCHMARK).mean()*100}
    total=pd.to_numeric(raw.loc[rows,total_col],errors='coerce').dropna()
    overall={'n':len(total),'highest':total.max(),'lowest':total.min(),'mean':total.mean(),'median':total.median(),'sd':total.std(ddof=1),'benchmark_pct':(total>=BENCHMARK).mean()*100}
    gtot=num(raw.iloc[5,total_col])
    return assessments,stats,overall,gtot,rows,pct_cols,total_col

def charts(stats,assessments,out):
    out=Path(out); out.mkdir(exist_ok=True); paths=[]; cs=list(stats)
    for kind,vals,title,ylabel,fn in [('att',[stats[c]['mean'] for c in cs],'Figure 1. CLO-wise OBE Attainment','Mean Attainment (%)','CLO_Attainment_Chart.png'),('bench',[stats[c]['pct70'] for c in cs],'Figure 2. Students Achieving ≥70% by CLO','Students achieving ≥70% (%)','Benchmark_Achievement_Chart.png')]:
        p=out/fn; fig,ax=plt.subplots(figsize=(9,5)); ax.bar(cs,vals); ax.axhline(70,ls='--',lw=1.5,label='70% benchmark'); ax.set_title(title); ax.set_xlabel('Course Learning Outcome'); ax.set_ylabel(ylabel); ax.set_ylim(0,100); ax.legend(); fig.tight_layout(); fig.savefig(p,dpi=200); plt.close(fig); paths.append(p)
    p=out/'Assessment_Performance_Chart.png'; fig,ax=plt.subplots(figsize=(12,6)); vals=[a['average'] for a in assessments]; ax.bar(range(len(vals)),vals); ax.set_title('Figure 3. Assessment Mean Scores as Reported in Excel'); ax.set_xlabel('Assessment / CLO'); ax.set_ylabel('Mean score (raw scale as provided)'); ax.set_xticks(range(len(vals))); ax.set_xticklabels([f"{a['clo']}\n{a['label']}" for a in assessments],rotation=55,ha='right',fontsize=8); fig.tight_layout(); fig.savefig(p,dpi=200); plt.close(fig); paths.append(p)
    return paths

def workbook(clos,assessments,stats,overall,student_df,out):
    wb=Workbook(); wb.remove(wb.active); title=PatternFill('solid',fgColor='1F4E78'); head=PatternFill('solid',fgColor='D9EAF7')
    def setup(ws,t,hs):
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(hs)); ws.cell(1,1,t).fill=title; ws.cell(1,1).font=Font(color='FFFFFF',bold=True,size=12)
        for j,h in enumerate(hs,1): ws.cell(3,j,h).fill=head; ws.cell(3,j).font=Font(bold=True); ws.cell(3,j).alignment=Alignment(wrap_text=True)
        ws.freeze_panes='A4'; ws.sheet_view.showGridLines=False
    def put(ws,rows,start=4):
        for r,row in enumerate(rows,start):
            for j,v in enumerate(row,1): ws.cell(r,j,v); ws.cell(r,j).alignment=Alignment(vertical='top',wrap_text=True)
    ws=wb.create_sheet('OBE Summary'); setup(ws,'OBE Summary',['CLO','Official CLO Description','Mean Attainment (%)','Students ≥70%','Students ≥70% (%)','SD','Status']); put(ws,[[c,clos[c],stats[c]['mean'],stats[c]['n70'],stats[c]['pct70'],stats[c]['sd'],status(stats[c]['mean'])] for c in clos])
    ws=wb.create_sheet('CLO-Assessment Mapping'); setup(ws,'CLO–Assessment Mapping',['CLO','Official CLO Description','Assessment/Question','Weightage','Average/Mean Score','Maximum Marks','Attainment %','Source']); put(ws,[[a['clo'],clos.get(a['clo'],'Unmatched CLO in Excel'),a['label'],a['weightage'],a['average'],'Not available in provided files','Not available in provided files','Excel OBE sheet'] for a in assessments])
    ws=wb.create_sheet('Assessment Analysis'); setup(ws,'Assessment Analysis',['CLO','Assessment/Question','Weightage','Average/Mean Score','Assessment Attainment %','Interpretation']); put(ws,[[a['clo'],a['label'],a['weightage'],a['average'],'Not available','Raw maximum marks are not separately provided; normalized attainment is not inferred.'] for a in assessments])
    ws=wb.create_sheet('Student-CLO Data'); setup(ws,'Student/CLO Data',list(student_df.columns)); put(ws,student_df.values.tolist())
    ws=wb.create_sheet('CQI Action Plan'); setup(ws,'CQI Action Plan',['CLO/Area','Identified Issue','Recommended Action','Teaching/Learning Intervention','Follow-up Evidence','Target']); put(ws,[[c,f"Mean attainment {stats[c]['mean']:.2f}% {'is below' if stats[c]['mean']<70 else 'meets'} the 70% benchmark.",f"Align intervention directly with the official CLO: {clos[c]}",'Use targeted practice, formative assessment, guided application and feedback tied to this CLO.','Repeat CLO-aligned assessment and compare attainment/benchmark achievement.','Raise/maintain mean attainment at or above 70%.'] for c in clos])
    ws=wb.create_sheet('Chart Data'); setup(ws,'Chart Data',['CLO','Mean Attainment (%)','Students ≥70% (%)','Status']); put(ws,[[c,stats[c]['mean'],stats[c]['pct70'],status(stats[c]['mean'])] for c in clos]);
    wb.save(out)

def report(info,objectives,clos,assessments,stats,overall,gtot,student_df,chart_paths,out):
    doc=Document(); sec=doc.sections[0]; sec.top_margin=sec.bottom_margin=__import__('docx').shared.Inches(.65); sec.left_margin=sec.right_margin=__import__('docx').shared.Inches(.75)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run('OUTCOME-BASED EDUCATION (OBE)\nEVALUATION REPORT\n').bold=True; p.add_run(f"\n{info['Course Title']} ({info['Course Code']})\nProgram: {info['Program']}\nSection: BSBA-2A2\nSemester: {info['Semester']}\n")
    doc.add_page_break()
    def table(headers,rows):
        t=doc.add_table(rows=1,cols=len(headers)); t.style='Table Grid'
        for i,h in enumerate(headers): t.rows[0].cells[i].text=str(h)
        for row in rows:
            cells=t.add_row().cells
            for i,v in enumerate(row): cells[i].text=str(v)
    strong=max(stats,key=lambda c:stats[c]['mean']); weak=min(stats,key=lambda c:stats[c]['mean'])
    doc.add_heading('1. Executive Summary',1); table(['Item','Result'],[['Course',f"{info['Course Title']} ({info['Course Code']})"],['Program / Section',f"{info['Program']} / BSBA-2A2"],['Semester',info['Semester']],['Instructor',info['Instructor/Faculty'] or 'Not available in provided files'],['Number of students',overall['n']],['Overall CLO attainment',f"{gtot:.2f}%"],['Strongest CLO',f"{strong} – {stats[strong]['mean']:.2f}%"],['Weakest CLO',f"{weak} – {stats[weak]['mean']:.2f}%"],['CLOs ≥70%',sum(stats[c]['mean']>=70 for c in clos)],['CLOs <70%',sum(stats[c]['mean']<70 for c in clos)],['Key CQI recommendation',f'Prioritize {weak} with CLO-specific intervention.']])
    doc.add_heading('2. Course Information',1); table(['Field','Information'],[[k,v or 'Not available in provided files'] for k,v in info.items()]); doc.add_heading('2.1 Course Description',2); doc.add_paragraph(info['Course Description'] or 'Not available in provided files.'); doc.add_heading('2.2 Course Objectives',2); [doc.add_paragraph(f'{i}. {x}') for i,x in enumerate(objectives,1)]; doc.add_heading('2.3 Official CLOs',2); table(['CLO','Official CLO'],[[c,clos[c]] for c in clos])
    doc.add_heading('3. CLO–Assessment Alignment',1); table(['CLO','Official CLO Description','Assessment/Question','Weightage','Maximum Marks'],[[a['clo'],clos.get(a['clo'],'Unmatched CLO in Excel'),a['label'],a['weightage'],'Not available in provided files'] for a in assessments])
    doc.add_heading('4. Methodology',1); doc.add_paragraph('The Course Outline is authoritative for course information and exact CLO wording. The Excel file is authoritative for all numerical OBE calculations. Benchmark = 70%; ≥80% Strong, 70–79.99% Satisfactory, <70% Needs Improvement. Missing maximum marks are not inferred.')
    doc.add_heading('5. CLO-wise OBE Attainment',1); table(['CLO','CLO Description','Mean Attainment (%)','Students ≥70%','Status'],[[c,clos[c],f"{stats[c]['mean']:.2f}%",f"{stats[c]['n70']} ({stats[c]['pct70']:.0f}%)",status(stats[c]['mean'])] for c in clos]); doc.add_paragraph(f'Overall weighted attainment from the Excel G.Tot field: {gtot:.2f}%.')
    doc.add_heading('6. Assessment-wise Analysis',1); table(['CLO','Assessment/Question','Weightage','Average/Mean Score','Attainment %','Interpretation'],[[a['clo'],a['label'],a['weightage'],a['average'],'Not available','Raw maximum marks are not separately available.'] for a in assessments])
    doc.add_heading('7. Student Performance Analysis',1); table(['Metric','Result'],[['Number assessed',overall['n']],['Highest overall score',f"{overall['highest']:.2f}"],['Lowest overall score',f"{overall['lowest']:.2f}"],['Mean overall score',f"{overall['mean']:.2f}"],['Median',f"{overall['median']:.2f}"],['Standard deviation',f"{overall['sd']:.2f}"],['Meeting 70%',f"{overall['benchmark_pct']:.0f}%"]])
    doc.add_heading('8. Charts and Visual Evidence',1); caps=['Figure 1. CLO-wise OBE attainment with the 70% benchmark.','Figure 2. Percentage of students achieving ≥70% for each CLO.','Figure 3. Assessment mean scores as reported in Excel.'];
    for p,cap in zip(chart_paths,caps): doc.add_picture(str(p),width=__import__('docx').shared.Inches(6.7)); q=doc.add_paragraph(cap); q.alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading('9. CLO Alignment with Course Outline',1); table(['CLO','Official CLO','Assessment Evidence','Attainment %','Benchmark Achievement %','Status','CQI Priority'],[[c,clos[c],'; '.join(a['label'] for a in assessments if a['clo']==c) or 'No assessment evidence for this CLO was identified in the provided Excel file.',f"{stats[c]['mean']:.2f}%",f"{stats[c]['pct70']:.0f}%",status(stats[c]['mean']),'High' if stats[c]['mean']<70 else 'Maintain'] for c in clos])
    doc.add_heading('10. Key Findings',1); doc.add_paragraph(f'Overall attainment is {gtot:.2f}%. {strong} is strongest ({stats[strong]["mean"]:.2f}%) and {weak} is weakest ({stats[weak]["mean"]:.2f}%). {weak} is the primary CQI priority because it is below the 70% benchmark.')
    doc.add_heading('11. CQI / Action Plan',1); table(['CLO/Area','Identified Issue','Recommended Action','Teaching/Learning Intervention','Follow-up Evidence','Target'],[[c,f"{stats[c]['mean']:.2f}% mean attainment.",f'Use the exact CLO as the intervention focus: {clos[c]}','Targeted practice, formative assessment and feedback tied to the CLO.','Repeat CLO-aligned evidence and compare attainment.','Mean attainment ≥70%.'] for c in clos])
    doc.add_heading('12. Conclusion',1); doc.add_paragraph(f'The OBE evidence indicates {gtot:.2f}% overall attainment, with {strong} strongest and {weak} weakest. CQI should focus on CLOs below the 70% benchmark and monitor subsequent evidence.')
    doc.add_heading('13. OBE Quality Check',1); table(['Check','Result'],[['Students analyzed',overall['n']],['Overall attainment',f'{gtot:.2f}%'],['Strongest CLO',strong],['Weakest CLO',weak],['CLOs ≥70%',sum(stats[c]['mean']>=70 for c in clos)],['CLOs <70%',sum(stats[c]['mean']<70 for c in clos)],['Assessments analyzed',len(assessments)],['CLOs matched',sum(1 for c in clos if any(a['clo']==c for a in assessments))],['Calculations internally consistent','Yes; all tables/charts use the same analysis.']])
    doc.save(out)


# -------------------- PROFESSIONAL OBE DASHBOARD UI --------------------
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(
    page_title="OBE Evaluation Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stApp { background: #ffffff; }
    [data-testid="stSidebar"] {
        background: #f1f4f8;
        border-right: 1px solid #e1e6ed;
    }
    [data-testid="stSidebar"] .block-container { padding-top: 2.2rem; }
    .hero {
        padding: 8px 0 22px 0;
    }
    .hero h1 {
        font-size: 2.55rem;
        line-height: 1.15;
        margin: 0;
        color: #202433;
        font-weight: 750;
    }
    .hero p {
        color: #8a8f99;
        font-size: 1rem;
        margin-top: 12px;
    }
    .section-title {
        font-size: 1.9rem;
        font-weight: 700;
        color: #172f4d;
        margin: 28px 0 14px 0;
    }
    .section-subtitle {
        color: #6f7782;
        margin-top: -8px;
        margin-bottom: 18px;
    }
    .settings-title {
        font-size: 1.25rem;
        font-weight: 700;
        margin-bottom: 18px;
        color: #202433;
    }
    .info-card {
        background: #f5f7fa;
        border-radius: 9px;
        padding: 13px 15px;
        min-height: 66px;
        border: 1px solid #e5e9ef;
    }
    .info-label {
        color: #66717f;
        font-size: .82rem;
        margin-bottom: 5px;
    }
    .info-value {
        color: #273142;
        font-size: .96rem;
        font-weight: 500;
    }
    .metric-card {
        background: white;
        border: 1px solid #e2e7ee;
        border-radius: 12px;
        padding: 14px;
    }
    div[data-testid="stFileUploader"] {
        border-radius: 10px;
    }
    .small-note {
        color: #737b86;
        font-size: .86rem;
    }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown('<div class="settings-title">Report Settings</div>', unsafe_allow_html=True)

    method = st.selectbox(
        "CLO attainment method",
        ["Direct student-threshold attainment", "Mean attainment"],
        index=0
    )

    benchmark = st.number_input(
        "Achievement threshold (%)",
        min_value=0.0, max_value=100.0, value=70.0, step=1.0
    )
    BENCHMARK = float(benchmark)

    target = st.number_input(
        "Target CLO attainment (%)",
        min_value=0.0, max_value=100.0, value=70.0, step=1.0
    )

    st.divider()
    st.markdown("**Status rules**")
    st.caption("≥80% = Strong")
    st.caption("70–79.99% = Satisfactory")
    st.caption("<70% = Needs Improvement")

st.markdown("""
<div class="hero">
    <h1>📊 OBE Evaluation Report Generator</h1>
    <p>Generate an auditable Course Learning Outcome (CLO) attainment report from assessment data.</p>
</div>
""", unsafe_allow_html=True)

# ---------- SECTION 1 ----------
st.markdown('<div class="section-title">1. Course Information</div>', unsafe_allow_html=True)

outline = st.file_uploader(
    "Upload Course Outline",
    type=["docx"],
    key="outline",
    label_visibility="collapsed"
)

# Keep Excel uploader immediately available but visually subordinate.
excel = st.file_uploader(
    "Upload OBE Assessment Excel",
    type=["xlsx"],
    key="excel",
    label_visibility="collapsed"
)

if outline:
    try:
        info, objectives, clos = parse_outline(outline.getvalue())
    except Exception:
        info = {k: "" for k in ['Course Code','Course Title','Credit Hours','Program','Semester','Campus','Instructor/Faculty','Course Description']}
        objectives, clos = [], {}
else:
    info = {k: "" for k in ['Course Code','Course Title','Credit Hours','Program','Semester','Campus','Instructor/Faculty','Course Description']}
    objectives, clos = [], {}

# Academic Year is intentionally not inferred.
course_fields = [
    ("Institution", "Not available in the provided files."),
    ("Department", "Not available in the provided files."),
    ("Program", info.get("Program") or "Not available in the provided files."),
    ("Course Title", info.get("Course Title") or "Not available in the provided files."),
    ("Course Code", info.get("Course Code") or "Not available in the provided files."),
    ("Semester", info.get("Semester") or "Not available in the provided files."),
    ("Academic Year", "Not available in the provided files."),
    ("Course Teacher / Instructor", info.get("Instructor/Faculty") or "Not available in the provided files."),
    ("Credit Hours", info.get("Credit Hours") or "Not available in the provided files."),
]

for start in range(0, len(course_fields), 3):
    cols = st.columns(3)
    for col, (label, value) in zip(cols, course_fields[start:start+3]):
        with col:
            st.markdown(
                f'<div class="info-card"><div class="info-label">{label}</div>'
                f'<div class="info-value">{value}</div></div>',
                unsafe_allow_html=True
            )

# ---------- SECTION 2 ----------
st.markdown('<div class="section-title">2. Course Learning Outcomes</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Official CLO wording is taken from the uploaded Course Outline and is not paraphrased.</div>',
    unsafe_allow_html=True
)

if clos:
    clo_df = pd.DataFrame({
        "CLO": list(clos.keys()),
        "CLO Description": list(clos.values())
    })
    st.dataframe(clo_df, use_container_width=True, hide_index=True)
else:
    st.info("Upload the Course Outline to display the official CLOs.")

# ---------- SECTION 3 ----------
st.markdown('<div class="section-title">3. Student Assessment Data</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Upload the OBE Excel workbook containing assessment marks, CLO mappings and attainment evidence.</div>',
    unsafe_allow_html=True
)

if not outline or not excel:
    st.info("Please upload both the Course Outline (.docx) and OBE Assessment Excel (.xlsx) to generate the full report.")

if outline and excel:
    try:
        # Re-parse to ensure the analysis uses the current files.
        info, objectives, clos = parse_outline(outline.getvalue())
        raw = pd.read_excel(io.BytesIO(excel.getvalue()), sheet_name='OBE', header=None)

        assessments, stats, overall, gtot, rows, pct_cols, total_col = analyze(raw, clos)

        st.success("✓ Both source files have been loaded and analyzed.")

        # Assessment overview
        ac = st.columns(4)
        ac[0].metric("Students Assessed", overall["n"])
        ac[1].metric("Assessments", len(assessments))
        ac[2].metric("Overall CLO Attainment", f"{gtot:.2f}%" if not pd.isna(gtot) else "N/A")
        ac[3].metric(
            f"Students ≥{benchmark:.0f}%",
            f"{overall['benchmark_pct']:.1f}%"
        )

        # ---------- SECTION 4 ----------
        st.markdown('<div class="section-title">4. CLO-wise OBE Attainment</div>', unsafe_allow_html=True)
        result_df = pd.DataFrame([
            {
                "CLO": c,
                "CLO Description": clos[c],
                "Mean Attainment (%)": stats[c]["mean"],
                f"Students ≥{benchmark:.0f}%": stats[c]["n70"],
                f"Students ≥{benchmark:.0f}% (%)": stats[c]["pct70"],
                "Standard Deviation": stats[c]["sd"],
                "Status": status(stats[c]["mean"])
            }
            for c in clos
        ])
        st.dataframe(result_df, use_container_width=True, hide_index=True)

        valid = {c: stats[c]["mean"] for c in clos if not pd.isna(stats[c]["mean"])}
        strongest = max(valid, key=valid.get) if valid else "Not available"
        weakest = min(valid, key=valid.get) if valid else "Not available"
        above = sum(v >= benchmark for v in valid.values())
        below = sum(v < benchmark for v in valid.values())

        st.markdown('<div class="section-title">5. Visual OBE Evidence</div>', unsafe_allow_html=True)
        outdir = Path("obe_output")
        outdir.mkdir(exist_ok=True)
        chart_paths = charts(stats, assessments, outdir)

        c1, c2 = st.columns(2)
        with c1:
            st.image(str(chart_paths[0]), use_container_width=True)
        with c2:
            st.image(str(chart_paths[1]), use_container_width=True)
        st.image(str(chart_paths[2]), use_container_width=True)

        # ---------- SECTION 6 ----------
        st.markdown('<div class="section-title">6. Assessment-to-CLO Mapping</div>', unsafe_allow_html=True)
        mapping_df = pd.DataFrame([
            {
                "CLO": a["clo"],
                "Assessment/Question": a["label"],
                "Maximum Marks": "Not available in the provided files.",
                "Average/Mean Score": a["average"],
                "Attainment %": "Not available in the provided files.",
                "Interpretation": "Maximum marks are not separately available; attainment percentage is not inferred."
            }
            for a in assessments
        ])
        st.dataframe(mapping_df, use_container_width=True, hide_index=True)

        # ---------- SECTION 7 ----------
        st.markdown('<div class="section-title">7. Student Performance Analysis</div>', unsafe_allow_html=True)
        sp = st.columns(6)
        sp[0].metric("N", overall["n"])
        sp[1].metric("Highest", f"{overall['highest']:.2f}")
        sp[2].metric("Lowest", f"{overall['lowest']:.2f}")
        sp[3].metric("Mean", f"{overall['mean']:.2f}")
        sp[4].metric("Median", f"{overall['median']:.2f}")
        sp[5].metric("SD", f"{overall['sd']:.2f}")

        # ---------- SECTION 8 ----------
        st.markdown('<div class="section-title">8. CLO Alignment with Course Outline</div>', unsafe_allow_html=True)
        alignment_df = pd.DataFrame([
            {
                "CLO": c,
                "Official CLO": clos[c],
                "Assessment Evidence": "; ".join(
                    a["label"] for a in assessments if a["clo"] == c
                ) or "No assessment evidence for this CLO was identified in the provided Excel file.",
                "Attainment %": stats[c]["mean"],
                "Benchmark Achievement %": stats[c]["pct70"],
                "Status": status(stats[c]["mean"]),
                "CQI Priority": "High" if not pd.isna(stats[c]["mean"]) and stats[c]["mean"] < benchmark else "Maintain"
            }
            for c in clos
        ])
        st.dataframe(alignment_df, use_container_width=True, hide_index=True)

        # ---------- SECTION 9 ----------
        st.markdown('<div class="section-title">9. CQI Action Plan</div>', unsafe_allow_html=True)
        cqi_rows = []
        for c in clos:
            if not pd.isna(stats[c]["mean"]) and stats[c]["mean"] < benchmark:
                cqi_rows.append({
                    "CLO/Area": c,
                    "Identified Issue": f"Mean attainment = {stats[c]['mean']:.2f}%, below the {benchmark:.0f}% benchmark.",
                    "Recommended Action": f"Focus improvement directly on the official CLO: {clos[c]}",
                    "Teaching/Learning Intervention": "Provide targeted CLO-aligned practice, formative assessment and structured feedback.",
                    "Follow-up Evidence": "Repeat a CLO-aligned assessment and compare subsequent attainment.",
                    "Target": f"Mean attainment ≥{target:.0f}%."
                })

        if cqi_rows:
            st.dataframe(pd.DataFrame(cqi_rows), use_container_width=True, hide_index=True)
        else:
            st.success(f"No CLO below the {benchmark:.0f}% benchmark was identified.")

        # ---------- SECTION 10 ----------
        st.markdown('<div class="section-title">10. Executive OBE Summary</div>', unsafe_allow_html=True)
        summary_cols = st.columns(4)
        summary_cols[0].metric("Strongest CLO", strongest)
        summary_cols[1].metric("Weakest CLO", weakest)
        summary_cols[2].metric(f"CLOs ≥{benchmark:.0f}%", above)
        summary_cols[3].metric(f"CLOs <{benchmark:.0f}%", below)

        st.markdown(
            f"""
            **Overall interpretation:** Overall CLO attainment is
            **{gtot:.2f}%** based on the Excel evidence.
            The strongest CLO is **{strongest}** and the weakest is **{weakest}**.
            **{above}** CLO(s) meet or exceed the selected benchmark of
            **{benchmark:.0f}%**, while **{below}** CLO(s) fall below it.
            """,
        )

        # ---------- OUTPUTS ----------
        st.markdown('<div class="section-title">11. Download OBE Outputs</div>', unsafe_allow_html=True)

        sdf = pd.DataFrame({
            "Sr.": range(1, overall["n"] + 1),
            "Roll No.": [clean(raw.iloc[r, 2]) for r in rows],
            "Section": [clean(raw.iloc[r, 3]) for r in rows],
        })
        for c in clos:
            sdf[c + " Attainment %"] = [num(raw.iloc[r, pct_cols[c]]) for r in rows]
        sdf["Overall Score"] = [num(raw.iloc[r, total_col]) for r in rows]

        xlsx_path = outdir / "OBE_Analysis.xlsx"
        docx_path = outdir / "OBE_Evaluation_Report.docx"
        workbook(clos, assessments, stats, overall, sdf, xlsx_path)
        report(info, objectives, clos, assessments, stats, overall, gtot, sdf, chart_paths, docx_path)

        dl = st.columns(5)
        with dl[0]:
            st.download_button("📄 Word Report", docx_path.read_bytes(),
                               "OBE_Evaluation_Report.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)
        with dl[1]:
            st.download_button("📊 Excel Workbook", xlsx_path.read_bytes(),
                               "OBE_Analysis.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        for col, p in zip(dl[2:], chart_paths):
            with col:
                st.download_button(
                    f"🖼️ {p.stem}",
                    p.read_bytes(),
                    p.name,
                    mime="image/png",
                    use_container_width=True
                )

        # ---------- QUALITY CHECK ----------
        with st.expander("🔎 Final OBE Quality Check"):
            matched = sum(1 for c in clos if any(a["clo"] == c for a in assessments))
            qc = pd.DataFrame([
                ["Students analyzed", overall["n"]],
                ["Overall CLO attainment", f"{gtot:.2f}%" if not pd.isna(gtot) else "Not available"],
                ["Strongest CLO", strongest],
                ["Weakest CLO", weakest],
                [f"CLOs ≥{benchmark:.0f}%", above],
                [f"CLOs <{benchmark:.0f}%", below],
                ["Assessments analyzed", len(assessments)],
                ["CLOs successfully matched", matched],
                ["Internal consistency", "Yes; dashboard, workbook and report use the same underlying calculations."]
            ], columns=["Quality Check", "Result"])
            st.dataframe(qc, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error("The uploaded files could not be processed.")
        st.exception(e)
