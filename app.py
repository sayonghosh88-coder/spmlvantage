from pathlib import Path
from datetime import date
import streamlit as st
from model import *
from presentation import *

ROOT=Path(__file__).parent
st.set_page_config(page_title='SPML Vantage',page_icon='◈',layout='wide',initial_sidebar_state='collapsed')

@st.cache_data(show_spinner=False)
def parse(payload): return read_workbook(payload)

def go(code=None):
    if code: st.query_params['project']=code
    elif 'project' in st.query_params: del st.query_params['project']

def html(value): st.html(value)

html('<style>'+ (ROOT/'assets/fonts.css').read_text()+(ROOT/'styles.css').read_text()+'</style>')
payload=st.session_state.get('workbook_bytes',(ROOT/'data/project_data.xlsx').read_bytes())
try: data=parse(payload)
except DataError as exc:
    st.error(str(exc)); st.stop()
projects=ordered_projects(data)
review=st.session_state.get('review_date',data['latest'])
code=st.query_params.get('project')
p=next((p for p in projects if p['Project code']==code),None)
if code and not p:
    st.warning('This project is not in the loaded workbook. Showing the portfolio.')
    go()

html(f'''<header class="masthead" id="portfolio"><a class="wordmark" href="#portfolio">SPML <strong>Vantage</strong><span>EXECUTIVE PROJECT INTELLIGENCE</span></a><div class="timestamp">Updated {esc(fmt(data['latest']))}<br><small>Review as of {esc(fmt(review))}</small></div></header>''')
if data['fictional']:
    html('<div class="demo-note">DEMONSTRATION · All project figures, dates and actions are fictional.</div>')
for warning in data['warnings']: st.warning(warning)

def cash(code=None):
    invoices=rows(data,'Invoicing Status',code); liabilities=rows(data,'Liability Status',code)
    collect=total(invoices,'Outstanding balance (INR lakh)');pay=total(liabilities,'Outstanding liability (INR lakh)')
    html(heading('Cash snapshot','Outstanding workbook balances; not net cash or a bank balance.','cash'))
    html('<div class="cash-grid">'+''.join(f'<div class="cash-card"><span class="eyebrow">{label}</span><strong>{money(t["value"])}</strong><p>{t["known"]} / {t["count"]} records with amounts</p></div>' for label,t in [('To collect',collect),('To pay',pay)])+'</div>')

def deployment_summary(code=None):
    d=deployment_totals(data,code)
    dates=', '.join(fmt(v) for v in sorted(d['dates'])) or 'Date not entered'
    html(heading('People & machines',f'{d["sites"]} site snapshots · {d["projects"]}/{d["expected_projects"]} projects · {dates}'))
    html('<div class="deployment-strip"><div><span class="eyebrow">LABOUR DEPLOYED</span><p><strong>'+esc(fmt(d['actual']))+'</strong> / '+esc(fmt(d['planned']))+' planned</p></div><div><span class="eyebrow">MACHINES DEPLOYED</span><p><strong>'+esc(fmt(d['machines']))+'</strong> at sites</p></div></div>')
    if len(d['dates'])>1: html('<p class="small-note amber">Site snapshot dates differ; these are latest available counts, not a same-day portfolio total.</p>')

if not p:
    contract=total(projects,'Contract value (INR lakh)'); indices=portfolio_evm(projects);asks=decision_queue(data)
    html(f'<div class="page-intro"><div><span class="eyebrow">THE MANAGEMENT VIEW</span><h1>Portfolio overview</h1></div><p>{len(projects)} projects · <strong class="mono">{money(contract["value"])}</strong> contract value<br><small>{contract["known"]}/{contract["count"]} project values recorded</small></p></div>')
    html(kpis([('Awaiting MD',str(len(asks)),'Decisions to take','red' if asks else 'green'),('Portfolio SPI',fmt(indices['SPI'],'decimal'),'Schedule performance',index_rag(indices['SPI'])),('Portfolio CPI',fmt(indices['CPI'],'decimal'),'Cost performance',index_rag(indices['CPI']))]))
    html('<p class="small-note">EVM is indicative: '+esc(indices['note'])+' Validation: '+esc(', '.join(sorted({str(x.get('EVM validation status','Not entered')) for x in projects})))+'.</p>')
    html(heading('Needs your decision',f'{len(asks)} awaiting MD · critical projects, then result due date','decisions'))
    html('<div class="asks-grid">'+(''.join(decision_card(data,r,review) for r in asks[:3]) or '<p>No decisions awaiting MD.</p>')+'</div>')
    if len(asks)>3:
        html('<details class="more-asks"><summary>View all '+str(len(asks))+' awaiting decisions</summary><div class="asks-grid">'+''.join(decision_card(data,r,review) for r in asks[3:])+'</div></details>')
    html(heading('Project health','Critical projects first · tap a card for the complete record'))
    with st.container(key='projects-grid'):
        for project in projects:
            with st.container(key='project-'+project['Project code']):
                html(project_card(data,project))
                st.button('Open '+project['Project name'],key='open-'+project['Project code'],on_click=go,args=(project['Project code'],),width='stretch')
    cash()
    deployment_summary()
else:
    st.button('← Portfolio',on_click=go,key='back',type='tertiary')
    rag=project_rag(p);e=evm(p);s=slip_days(p)
    html(f'<div class="project-title"><div><span class="code">{esc(code)}</span><h1>{esc(p["Project name"])}</h1></div>{chip(LABEL[rag],rag)}</div>')
    slip='Completion dates not confirmed' if s is None else f'{s} days beyond approved completion' if s>0 else 'Forecast within approved completion'
    html(f'''<section class="hero">{progress('Physical progress',p.get('Overall planned progress'),p.get('Overall actual progress'))}<div class="hero-grid"><div><span class="eyebrow">COMPLETION FORECAST</span><h3>{esc(fmt(p.get('Forecast completion'),'date'))}</h3><p class="{'red' if s and s>30 else 'amber' if s and s>0 else 'green'}">{esc(slip)}</p><small>Approved: {esc(fmt(p.get('Latest approved completion'),'date'))}</small></div><div><span class="eyebrow">EOT</span><h3>{esc(p.get('EOT status'))}</h3><p>Requested up to {esc(fmt(p.get('EOT requested up to'),'date'))}</p></div><div><span class="eyebrow">NEXT MILESTONE</span><h3>{esc(p.get('Next critical milestone'))}</h3><p>Plan {esc(fmt(p.get('Milestone planned date'),'date'))}<br>Forecast {esc(fmt(p.get('Milestone forecast date'),'date'))}</p></div></div></section>''')
    html(heading('Discipline progress','Actual fill · planned marker'))
    html('<div class="discipline-grid">'+''.join(progress(d,p.get(d+' planned progress'),p.get(d+' actual progress')) for d in ['Engineering','Procurement','Construction'])+'</div>')
    html(heading('Earned value','Reporting date '+fmt(p.get('EVM reporting date'),'date')+' · '+str(p.get('EVM validation status'))))
    html(kpis([('CPI',fmt(e['CPI'],'decimal'),'Cost performance',index_rag(e['CPI'])),('SPI',fmt(e['SPI'],'decimal'),'Schedule performance',index_rag(e['SPI'])),('VAC',money(e['VAC']),'Budget less forecast','unknown' if e['VAC'] is None else 'red' if e['VAC']<0 else 'green')]))
    html(f'<div class="evm-note"><p><b>BAC</b> {money(e["BAC"])} <span class="muted">→</span> <b>EAC</b> {money(e["EAC"])}</p><p><b>Reason:</b> {esc(p.get("Reason for EVM variance"))}</p><p><b>Corrective action:</b> {esc(p.get("EVM corrective action"))}</p><p class="meta">{esc(p.get("EVM responsible owner"))} · Due {esc(fmt(p.get("EVM action due date"),"date"))}</p></div>')
    html(heading('Month goals',fmt(p.get('Goal month'),'month')))
    goals=sorted(range(1,6),key=lambda i:RANK.get(str(p.get(f'Goal {i} Status')).lower(),3))
    html('<div class="goals">'+''.join(f'<div class="goal">{chip("Goal "+str(i),str(p.get(f"Goal {i} Status")).lower() if str(p.get(f"Goal {i} Status")).lower() in RANK else "unknown")}<span>{esc(p.get("Project Goal "+str(i)))}</span><b class="mono">{esc(fmt(p.get(f"Goal {i} Completion"),"percent"))}</b></div>' for i in goals)+'</div>')
    html(full_fields(REGISTRY,p,'Complete project status & EVM record'))
    html('<div class="record-divider"><span>FULL PROJECT RECORD</span><p>Every register, every row. Expand a card to read every field.</p></div>')
    html('<nav class="register-nav" aria-label="Project registers">'+''.join(f'<a href="#{REGISTER_IDS[s]}">{esc(s)}</a>' for s in REGISTER_ORDER)+'</nav>')
    for sheet in REGISTER_ORDER:
        if sheet=='Invoicing Status': cash(code)
        if sheet=='Construction & Resources':
            snapshots,snapshot_date=deployment(data,code)
            deployment_summary(code)
            if snapshots:
                html('<details class="more-asks"><summary>View site labour and machine deployment</summary>'+''.join(full_fields('Snapshot',r,str(r.get('Site / scope'))) for r in snapshots)+'</details>')
            else: html('<p>No site deployment snapshots entered.</p>')
        html(register_html(data,sheet,code,review))
    next_p=projects[(projects.index(p)+1)%len(projects)]
    st.button('Next project review → '+next_p['Project name'],on_click=go,args=(next_p['Project code'],),key='next-project',type='primary',width='stretch')

with st.expander('Workbook & review settings'):
    st.caption('Loaded: '+st.session_state.get('workbook_name','Five-project fictional test workbook'))
    uploaded=st.file_uploader('Load an updated nine-sheet Excel workbook',type=['xlsx'],key='uploader')
    if uploaded is not None and st.button('Use this workbook'):
        try:
            candidate=uploaded.getvalue();fresh=parse(candidate)
            st.session_state.update(workbook_bytes=candidate,workbook_name=uploaded.name,review_date=fresh['latest'])
            go();st.rerun()
        except DataError as exc: st.error(str(exc)+' The previous workbook is still loaded.')
    st.date_input('Review date for overdue flags',value=review,key='review_date')
    st.download_button('Download loaded workbook',payload,file_name='SPML_Vantage_Input.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.caption('Uploads apply to your browser session. To update the shared dashboard, replace data/project_data.xlsx in the deployed app.')
    if st.button('Restore demonstration workbook'):
        for key in ['workbook_bytes','workbook_name','review_date']: st.session_state.pop(key,None)
        go();st.rerun()

html('''<details class="methodology"><summary>How the indicators are calculated</summary><p>Physical progress and discipline colours: green at no more than 3 percentage points behind plan; amber at no more than 8 points behind; red beyond 8. CPI and SPI: green ≥ 1.00; amber ≥ 0.90; red below 0.90.</p><p>Project health uses the most severe of physical progress, CPI, SPI and completion slip. A slip beyond the latest approved completion is amber for 1–30 days and red beyond 30. Missing inputs are shown as not confirmed. Register rows are ordered by their own status and overdue date; completed rows remain visible.</p><p>CPI = EV / AC; SPI = EV / PV; EAC = BAC / CPI; ETC = EAC − AC; VAC = BAC − EAC. EAC assumes current cost efficiency continues. Portfolio indices use summed EV, AC and PV, and require matching EVM reporting dates. Unverified EVM remains indicative. No division by zero is shown as a valid result.</p><p>Cash totals use the entered outstanding balances; financial impact on decision cards is not added to outstanding cash. Monetary inputs are INR lakh; 100 lakh = 1 crore. “Result due” is the workbook’s result due date, not an invented decision deadline. “Not yet occurred” is retained as a pending actual date.</p></details>''')
with st.container(key='bottom-nav'):
    a,b,c=st.columns(3)
    with a:
        if p: st.button('◈ Portfolio',on_click=go,key='nav-home',width='stretch')
        else: html('<a class="nav-tab active" href="#portfolio">◈ Portfolio</a>')
    with b: html(f'<a class="nav-tab" href="#{REGISTER_IDS["MD Decisions"] if p else "decisions"}">≡ Decisions</a>')
    with c: html('<a class="nav-tab" href="#cash">₹ Cash</a>')
