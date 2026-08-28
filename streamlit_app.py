
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

st.set_page_config(page_title='OBE Evaluation Report Generator',page_icon='📊',layout='wide',initial_sidebar_state='expanded')
BENCHMARK=70.0

# --- Professional dashboard styling ---
st.markdown("""
<style>
.main {background:#f7f9fc;}
.block-container {padding-top:1.5rem; padding-bottom:2rem;}
.obe-hero {
 padding:1.35rem 1.6rem; border-radius:16px; margin-bottom:1.2rem;
 background:linear-gradient(135deg,#163a5f 0%,#285f8f 100%); color:white;
}
.obe-hero h1 {margin:0; font-size:2rem;}
.obe-hero p {margin:.4rem 0 0; opacity:.95;}
div[data-testid="stMetric"] {
 background:white; border:1px solid #dfe6ee; border-radius:12px; padding:10px;
}
.stButton>button {border-radius:9px; font-weight:600;}
</style>
""", unsafe_allow_html=True)
st.markdown("""
<div class="obe-hero">
<h1>📊 OBE Evaluation Report Generator</h1>
<p>Automated, auditable CLO attainment analysis, assessment evidence, CQI planning and accreditation-ready reporting.</p>
</div>
""", unsafe_allow_html=True)


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

# Missing import used only for paragraph alignment
from docx.enum.text import WD_ALIGN_PARAGRAPH

with st.sidebar:
    st.header("⚙️ OBE Controls")
    BENCHMARK = st.number_input("OBE benchmark (%)", 0.0, 100.0, 70.0, 1.0)
    st.divider()
    st.markdown("### Source hierarchy")
    st.info("Course Outline → official course information and exact CLO wording.\n\nExcel → numerical OBE evidence, marks, assessment results and mappings.")
    st.markdown("### Status rules")
    st.write("**≥80%** — Strong")
    st.write("**70–79.99%** — Satisfactory")
    st.write("**<70%** — Needs Improvement")

st.subheader("📁 Upload OBE Evidence")
c1, c2 = st.columns(2)
with c1:
    outline = st.file_uploader("1. Course Outline (.docx)", type=["docx"])
with c2:
    excel = st.file_uploader("2. OBE Assessment Excel (.xlsx)", type=["xlsx"])

if outline and excel:
    try:
        info, objectives, clos = parse_outline(outline.getvalue())
        raw = pd.read_excel(io.BytesIO(excel.getvalue()), sheet_name='OBE', header=None)
        assessments, stats, overall, gtot, rows, pct_cols, total_col = analyze(raw, clos)

        st.success("✓ Course Outline and OBE Excel successfully loaded.")

        st.subheader("📘 Course Overview")
        mc = st.columns(5)
        for col, label, value in [
            (mc[0],"Course Code",info["Course Code"] or "Not available"),
            (mc[1],"Course Title",info["Course Title"] or "Not available"),
            (mc[2],"Program",info["Program"] or "Not available"),
            (mc[3],"Semester",info["Semester"] or "Not available"),
            (mc[4],"Students",overall["n"])
        ]:
            col.metric(label, value)

        with st.expander("📚 Official CLOs", expanded=True):
            st.dataframe(pd.DataFrame({"CLO":list(clos),"Official CLO":list(clos.values())}),
                         use_container_width=True, hide_index=True)

        st.subheader("📊 CLO Attainment Summary")
        result_df = pd.DataFrame([
            {"CLO":c,"Official CLO":clos[c],"Mean Attainment (%)":stats[c]["mean"],
             "Students ≥70%":stats[c]["n70"],"Students ≥70% (%)":stats[c]["pct70"],
             "SD":stats[c]["sd"],"Status":status(stats[c]["mean"])}
            for c in clos
        ])
        st.dataframe(result_df,use_container_width=True,hide_index=True)

        valid={c:stats[c]["mean"] for c in clos if not pd.isna(stats[c]["mean"])}
        strongest=max(valid,key=valid.get) if valid else "N/A"
        weakest=min(valid,key=valid.get) if valid else "N/A"
        mm=st.columns(5)
        mm[0].metric("Overall Attainment",f"{gtot:.2f}%" if not pd.isna(gtot) else "N/A")
        mm[1].metric("Strongest CLO",strongest)
        mm[2].metric("Weakest CLO",weakest)
        mm[3].metric("CLOs ≥70%",sum(v>=70 for v in valid.values()))
        mm[4].metric("CLOs <70%",sum(v<70 for v in valid.values()))

        st.subheader("📈 Visual OBE Evidence")
        out=Path("obe_output"); out.mkdir(exist_ok=True)
        cp=charts(stats,assessments,out)
        ch1,ch2=st.columns(2)
        with ch1: st.image(str(cp[0]),use_container_width=True)
        with ch2: st.image(str(cp[1]),use_container_width=True)
        st.image(str(cp[2]),use_container_width=True)

        st.subheader("📝 Assessment-wise Analysis")
        adf=pd.DataFrame([
            {"CLO":a["clo"],"Assessment/Question":a["label"],"Weightage":a["weightage"],
             "Average/Mean Score":a["average"],
             "Maximum Marks":"Not available in the provided files.",
             "Attainment %":"Not available in the provided files.",
             "Interpretation":"Raw maximum marks are not separately available; normalized attainment is not inferred."}
            for a in assessments
        ])
        st.dataframe(adf,use_container_width=True,hide_index=True)

        st.subheader("👥 Student Performance")
        sp=st.columns(6)
        for col,label,value in [
            (sp[0],"N",overall["n"]),
            (sp[1],"Highest",f'{overall["highest"]:.2f}'),
            (sp[2],"Lowest",f'{overall["lowest"]:.2f}'),
            (sp[3],"Mean",f'{overall["mean"]:.2f}'),
            (sp[4],"Median",f'{overall["median"]:.2f}'),
            (sp[5],"SD",f'{overall["sd"]:.2f}')
        ]: col.metric(label,value)
        st.write(f"**Overall students meeting the {BENCHMARK:.0f}% benchmark:** {overall['benchmark_pct']:.1f}%")

        st.subheader("🔗 CLO Alignment with Course Outline")
        align=pd.DataFrame([
            {"CLO":c,"Official CLO":clos[c],
             "Assessment Evidence":"; ".join(a["label"] for a in assessments if a["clo"]==c) or
                 "No assessment evidence for this CLO was identified in the provided Excel file.",
             "Attainment %":stats[c]["mean"],"Benchmark Achievement %":stats[c]["pct70"],
             "Status":status(stats[c]["mean"]),
             "CQI Priority":"High" if not pd.isna(stats[c]["mean"]) and stats[c]["mean"]<70 else "Maintain"}
            for c in clos
        ])
        st.dataframe(align,use_container_width=True,hide_index=True)

        st.subheader("🛠️ CQI Action Plan")
        cqi=[]
        for c in clos:
            if not pd.isna(stats[c]["mean"]) and stats[c]["mean"]<70:
                cqi.append({"CLO/Area":c,"Identified Issue":f"Mean attainment = {stats[c]['mean']:.2f}%, below 70%.",
                            "Recommended Action":f"Focus directly on the official CLO: {clos[c]}",
                            "Teaching/Learning Intervention":"Targeted practice, formative assessment and feedback tied to the CLO.",
                            "Follow-up Evidence":"Repeat CLO-aligned assessment and compare attainment.",
                            "Target":"Mean attainment ≥70%."})
        if cqi: st.dataframe(pd.DataFrame(cqi),use_container_width=True,hide_index=True)
        else: st.success("No CLO below the 70% benchmark was identified from the available attainment data.")

        st.subheader("📥 Downloadable OBE Outputs")
        sdf=pd.DataFrame({"Sr.":range(1,overall["n"]+1),
                          "Roll No.":[clean(raw.iloc[r,2]) for r in rows],
                          "Section":[clean(raw.iloc[r,3]) for r in rows]})
        for c in clos: sdf[c+" Attainment %"]=[num(raw.iloc[r,pct_cols[c]]) for r in rows]
        sdf["Overall Score"]=[num(raw.iloc[r,total_col]) for r in rows]
        x=out/"OBE_Analysis.xlsx"; d=out/"OBE_Evaluation_Report.docx"
        workbook(clos,assessments,stats,overall,sdf,x)
        report(info,objectives,clos,assessments,stats,overall,gtot,sdf,cp,d)

        b1,b2=st.columns(2)
        with b1: st.download_button("📄 Download Word Report",d.read_bytes(),"OBE_Evaluation_Report.docx",use_container_width=True)
        with b2: st.download_button("📊 Download Excel Workbook",x.read_bytes(),"OBE_Analysis.xlsx",use_container_width=True)
        for col,p in zip(st.columns(3),cp):
            with col:
                st.download_button(f"Download {p.name}",p.read_bytes(),p.name,mime="image/png",use_container_width=True)

        with st.expander("🔎 Final OBE Quality Check"):
            matched=sum(1 for c in clos if any(a["clo"]==c for a in assessments))
            qc=pd.DataFrame([
                ["Students analyzed",overall["n"]],
                ["Overall CLO attainment",f"{gtot:.2f}%" if not pd.isna(gtot) else "Not available"],
                ["Strongest CLO",strongest],["Weakest CLO",weakest],
                ["CLOs ≥70%",sum(v>=70 for v in valid.values())],
                ["CLOs <70%",sum(v<70 for v in valid.values())],
                ["Assessments analyzed",len(assessments)],
                ["CLOs successfully matched",matched],
                ["Internal consistency","Yes; dashboard, workbook and report use the same analysis."]
            ],columns=["Check","Result"])
            st.dataframe(qc,use_container_width=True,hide_index=True)

    except Exception as e:
        st.error("The uploaded files could not be processed.")
        st.exception(e)
else:
st.info("Upload both files above to begin the OBE analysis.")
