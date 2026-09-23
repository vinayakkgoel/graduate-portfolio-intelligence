import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

st.set_page_config(page_title='Graduate Portfolio Intelligence', page_icon='📊', layout='wide', initial_sidebar_state='expanded')
MODEL = Path(__file__).parent / 'model.xlsx'

@st.cache_data
def read(sheet, header=0):
    return pd.read_excel(MODEL, sheet_name=sheet, header=header, engine='openpyxl')

@st.cache_data
def load_model():
    a = read('Method A - Summary')
    b = read('Method B - Summary')
    rev = read('Revenue Outlook', 3)
    cac = read('CAC & LTR Data (Trial)', 6)
    dash = read('Trends & Risk Dashboard', None)
    for df in [a,b]: df['Code'] = df['Code'].astype(str).str.strip()
    out = a.merge(b[['Code','# Cohorts','# Mature Cohorts','Pooled S(T3)','Simple-Avg S(T3)','Difference (Pooled - Simple Avg)','Cross-check vs Method A','Sample Flag']], on='Code', how='left', suffixes=('','_B'))
    for c in ['New Student Retention','Returning Student Retention','LTR Factor (New x Returning)','Pooled S(T3)']:
        out[c] = pd.to_numeric(out[c], errors='coerce')
    out['New Retention %'] = out['New Student Retention']*100
    out['Returning Retention %'] = out['Returning Student Retention']*100
    out['LTR Factor %'] = out['LTR Factor (New x Returning)']*100
    out['Pooled T3 %'] = out['Pooled S(T3)']*100
    if 'Program' in rev.columns:
        rev['Code'] = rev['Program'].astype(str).str.extract(r'^(\w+)')[0]
        for src,dst in [('LTR\n(Program Total)','LTR'),('Revenue\n(Program Total)','Revenue'),('New Starts\n(Fall 2026)','New Starts'),('LTR as % of Tuition\n(Direct)','LTR % Tuition')]:
            rev[dst] = pd.to_numeric(rev.get(src), errors='coerce')
        rev['LTR % Tuition'] *= 100
    if 'Program' in cac.columns:
        cac['Code'] = cac['Program'].astype(str).str.extract(r'^(\w+)')[0]
        for src,dst in [('CAC\n($/student)','CAC'),('LTR/Student\n(Direct method)','LTR/Student'),('Revenue Multiple\n(Gross LTR ÷ CAC)','Revenue Multiple')]:
            cac[dst] = pd.to_numeric(cac.get(src), errors='coerce')
    return out, rev, cac, dash

model, rev, cac, dash = load_model()

def dash_row(n):
    try:
        r = dash.iloc[n-1]
        return pd.to_numeric(r.iloc[1], errors='coerce'), pd.to_numeric(r.iloc[2], errors='coerce')
    except Exception: return np.nan, np.nan

fall, fall_prior = dash_row(8); fy, fy_prior = dash_row(9); new, new_prior = dash_row(10); cont, cont_prior = dash_row(11)
usable = model[model['Sample Flag'].eq('OK')].copy()
programs = model[['Code','Program Name']].dropna().sort_values('Program Name')

st.sidebar.title('Graduate Portfolio Intelligence')
st.sidebar.caption('V2 working prototype • Excel remains the reference model')
page = st.sidebar.radio('Navigate', ['Executive Overview','Program Explorer','Scenario Lab','Model QA'])

if page == 'Executive Overview':
    st.title('Executive Overview')
    st.caption('A management view of enrollment, persistence, unit economics and portfolio risk.')
    cols = st.columns(4)
    for c,(label,val,prior) in zip(cols,[('Fall Enrollment',fall,fall_prior),('FY Enrollment',fy,fy_prior),('FY New Starts',new,new_prior),('FY Continuing',cont,cont_prior)]):
        delta = val-prior if pd.notna(val) and pd.notna(prior) else None
        c.metric(label, f'{val:,.0f}' if pd.notna(val) else '—', f'{delta:+,.0f}' if delta is not None else None)
    st.divider()
    c1,c2,c3,c4 = st.columns(4)
    c1.metric('Graduate Programs', f'{len(model):,}')
    c2.metric('Programs with mature data', f'{len(usable):,}')
    c3.metric('Programs needing data', f'{len(model)-len(usable):,}')
    c4.metric('Method B cross-checks', f"{(model['Cross-check vs Method A']=='Matches').sum():,}")
    left,right = st.columns([1.2,1])
    with left:
        st.subheader('Persistence by program')
        chart = usable[['Program Name','New Retention %','Returning Retention %']].set_index('Program Name').sort_values('New Retention %').tail(15)
        st.bar_chart(chart)
    with right:
        st.subheader('Portfolio watchlist')
        w = usable[['Code','Program Name','New Retention %','Returning Retention %','Floor Breach Term']].copy()
        w['Risk Score'] = (100-w['New Retention %']) + (100-w['Returning Retention %'])
        w = w.sort_values('Risk Score',ascending=False).head(10)
        w['Risk'] = np.select([w['Risk Score']>100,w['Risk Score']>70],['High','Watch'],'Lower')
        st.dataframe(w[['Code','Program Name','Risk','Floor Breach Term']],hide_index=True,use_container_width=True)
    st.subheader('What this tells leadership')
    st.info('Use the dashboard to identify where enrollment is changing, then drill into persistence and unit economics. The Scenario Lab is intentionally directional until the model is fully reconciled to the workbook.')

elif page == 'Program Explorer':
    st.title('Program Explorer')
    choice = st.selectbox('Select a graduate program', [f'{r.Code} — {r["Program Name"]}' for _,r in programs.iterrows()])
    code = choice.split(' — ',1)[0]
    row = model[model.Code.eq(code)].iloc[0]
    rr = rev[rev.Code.eq(code)] if 'Code' in rev.columns else pd.DataFrame()
    cc = cac[cac.Code.eq(code)] if 'Code' in cac.columns else pd.DataFrame()
    st.subheader(f'{code} — {row["Program Name"]}')
    cols=st.columns(5)
    vals=[('T1 mature cohort',row['N @ T1 (mature)'],lambda x:f'{x:,.0f}'),('New retention',row['New Retention %'],lambda x:f'{x:.1f}%'),('Returning retention',row['Returning Retention %'],lambda x:f'{x:.1f}%'),('LTR factor',row['LTR Factor %'],lambda x:f'{x:.1f}%'),('Floor breach',row['Floor Breach Term'],lambda x:str(x))]
    for c,(lab,v,fmt) in zip(cols,vals): c.metric(lab, fmt(v) if pd.notna(v) else '—')
    a,b = st.columns(2)
    with a:
        st.subheader('Persistence diagnostics')
        st.write(f"**Method B pooled T3:** {row['Pooled T3 %']:.1f}%" if pd.notna(row['Pooled T3 %']) else '**Method B pooled T3:** —')
        st.write(f"**Cross-check:** {row['Cross-check vs Method A']}")
        st.write(f"**Sample flag:** {row['Sample Flag']}")
        st.write(f"**Cohorts:** {row['# Cohorts']}")
    with b:
        st.subheader('Financial snapshot')
        if len(rr):
            r=rr.iloc[0]
            st.write(f"**Revenue:** ${r['Revenue']:,.0f}" if pd.notna(r['Revenue']) else '**Revenue:** —')
            st.write(f"**LTR:** ${r['LTR']:,.0f}" if pd.notna(r['LTR']) else '**LTR:** —')
            st.write(f"**LTR / tuition:** {r['LTR % Tuition']:.1f}%" if pd.notna(r['LTR % Tuition']) else '**LTR / tuition:** —')
        if len(cc):
            c=cc.iloc[0]
            st.write(f"**Directional CAC:** ${c['CAC']:,.0f}" if pd.notna(c['CAC']) else '**Directional CAC:** —')
            st.write(f"**LTR / student:** ${c['LTR/Student']:,.0f}" if pd.notna(c['LTR/Student']) else '**LTR / student:** —')
            st.write(f"**Revenue multiple:** {c['Revenue Multiple']:.1f}x" if pd.notna(c['Revenue Multiple']) else '**Revenue multiple:** —')
    st.divider()
    st.subheader('Portfolio comparison')
    comp = usable[['Code','Program Name','New Retention %','Returning Retention %','LTR Factor %']].copy()
    comp['Selected'] = np.where(comp.Code.eq(code),'← selected','')
    st.dataframe(comp.sort_values('LTR Factor %',ascending=False),hide_index=True,use_container_width=True)

elif page == 'Scenario Lab':
    st.title('Scenario Lab')
    st.caption('Directional planning tool. It does not overwrite the Excel model.')
    col1,col2,col3 = st.columns(3)
    start_adj = col1.slider('New starts change',-20,20,0,1,format='%d%%')
    persist_adj = col2.slider('Persistence change',-20,20,0,1,format='%d%%')
    cac_adj = col3.slider('CAC change',-20,20,0,1,format='%d%%')
    base_starts = float(new) if pd.notna(new) else 0
    base_ltr = float(pd.to_numeric(rev['LTR'],errors='coerce').sum()) if 'LTR' in rev.columns else np.nan
    base_cac = float(pd.to_numeric(cac['CAC'],errors='coerce').sum()) if 'CAC' in cac.columns else np.nan
    scenario_starts = base_starts*(1+start_adj/100)
    scenario_ltr = base_ltr*(1+start_adj/100)*(1+persist_adj/100) if pd.notna(base_ltr) else np.nan
    scenario_cac = base_cac*(1+cac_adj/100) if pd.notna(base_cac) else np.nan
    m=st.columns(4)
    m[0].metric('New starts',f'{scenario_starts:,.0f}',f'{scenario_starts-base_starts:+,.0f}')
    m[1].metric('Directional LTR',f'${scenario_ltr:,.0f}',f'${scenario_ltr-base_ltr:+,.0f}' if pd.notna(base_ltr) else None)
    m[2].metric('Directional CAC',f'${scenario_cac:,.0f}',f'${scenario_cac-base_cac:+,.0f}' if pd.notna(base_cac) else None)
    ratio = scenario_ltr/scenario_cac if pd.notna(scenario_ltr) and scenario_cac else np.nan
    m[3].metric('Directional LTR:CAC',f'{ratio:.1f}x' if pd.notna(ratio) else '—')
    st.warning('This is a sensitivity layer, not a validated forecast. Once Excel reconciliation is complete, these controls can be connected to the formal model engine.')

else:
    st.title('Model QA')
    st.caption('First-pass checks designed to keep the web model honest.')
    checks=[]
    checks.append(('Program count',len(model), 'Expected: 41', len(model)==41))
    checks.append(('Fall enrollment',fall, 'Workbook dashboard', pd.notna(fall)))
    checks.append(('FY enrollment',fy, 'Workbook dashboard', pd.notna(fy)))
    checks.append(('FY new starts',new, 'Workbook dashboard', pd.notna(new)))
    checks.append(('FY continuing',cont, 'Workbook dashboard', pd.notna(cont)))
    checks.append(('Revenue rows parsed',int(rev['Revenue'].notna().sum()) if 'Revenue' in rev.columns else 0,'Should be > 0', 'Revenue' in rev.columns and rev['Revenue'].notna().sum()>0))
    checks.append(('CAC rows parsed',int(cac['CAC'].notna().sum()) if 'CAC' in cac.columns else 0,'Directional / trial', 'CAC' in cac.columns and cac['CAC'].notna().sum()>0))
    qa=pd.DataFrame(checks,columns=['Check','Value','Reference','Pass'])
    qa['Status']=np.where(qa.Pass,'PASS','CHECK')
    st.dataframe(qa[['Status','Check','Value','Reference']],hide_index=True,use_container_width=True)
    st.subheader('Next reconciliation set')
    st.write('Validate 5–10 representative programs across New Retention, Returning Retention, LTR Factor, Revenue, LTR and directional CAC before production deployment.')

st.sidebar.divider()
st.sidebar.caption('Source: Persistence_Model_Graduate_Programs__9_23_2026__3.xlsx')
