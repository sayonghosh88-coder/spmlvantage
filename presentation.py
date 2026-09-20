"""Escaped, accessible presentation shared by the app and checks."""
from html import escape
from datetime import date, datetime
from model import *

def esc(v):
    return escape(str(v if v is not None else 'Not entered'), quote=True)

def fmt(v, kind=None):
    if v is None or v == '': return 'Not entered'
    d=as_date(v)
    if isinstance(v,(date,datetime)) or (kind in ['date','month'] and d):
        return d.strftime('%b %Y' if kind=='month' else '%d %b %Y')
    n=number(v)
    if n is not None:
        if kind=='percent': return f'{n*100:,.1f}%'
        if kind=='money': return f'₹{n:,.2f} L'
        if kind=='decimal': return f'{n:,.2f}'
        return f'{n:,.0f}' if n.is_integer() else f'{n:,.2f}'
    return str(v)

def money(v, crore=True):
    n=number(v)
    return 'Not entered' if n is None else f'₹{n/100:,.2f} Cr' if crore else f'₹{n:,.2f} L'

def chip(text,rag='unknown'):
    return f'<span class="chip {rag}">{esc(text)}</span>'

def heading(title,subtitle='',id=''):
    return f'<div class="section-heading" id="{esc(id)}"><h2>{esc(title)}</h2><p>{esc(subtitle)}</p></div>'

def progress(label,plan,actual):
    rag=progress_rag(plan,actual)
    if rag=='unknown': return f'<div class="progress-block"><b>{esc(label)}</b><p>Progress not confirmed</p></div>'
    return f'''<div class="progress-block"><div class="spread"><span>{esc(label)}</span><span class="mono"><strong>{fmt(actual,'percent')}</strong><small> / {fmt(plan,'percent')} plan</small></span></div>
    <div class="track" role="img" aria-label="{esc(label)}: actual {fmt(actual,'percent')}, planned {fmt(plan,'percent')}"><div class="fill {rag}" style="width:{actual*100}%"></div><i style="left:clamp(1px,{plan*100}%,calc(100% - 2px))"></i></div></div>'''

def kpis(items):
    return '<div class="kpis">'+''.join(f'<div class="kpi {"money-kpi" if "₹" in value else ""}"><div class="eyebrow">{esc(label)}</div><div class="big {rag}">{esc(value)}</div><small>{esc(note)}</small></div>' for label,value,note,rag in items)+'</div>'

def flag(p):
    s=slip_days(p); e=evm(p)
    if s and s>0: return f'Forecast {s} days beyond approved completion.'
    if e['CPI'] is not None and e['CPI']<1: return 'Cost performance below budget; corrective action required.'
    if progress_rag(p.get('Overall planned progress'),p.get('Overall actual progress')) in ['red','amber']:
        return str(p.get('Main reason for slippage') or 'Physical progress behind plan.')
    return 'Progress and EVM indicators on track.' if project_rag(p)=='green' else 'Confirm missing performance inputs.'

def project_card(data,p):
    code=p['Project code']; rag=project_rag(p); e=evm(p)
    asks=len([r for r in rows(data,'MD Decisions',code) if r.get('Status')=='Awaiting decision'])
    opened=len([r for r in rows(data,'MD Decisions',code) if r.get('Status')!='Closed'])
    return f'''<article class="project-card"><div class="spread"><span class="eyebrow"><span class="dot {rag}"></span>{esc(LABEL[rag])}</span><span class="arrow">↗</span></div>
    <h3>{esc(p['Project name'])}</h3><div class="code">{esc(code)}</div>
    {progress('Physical progress',p.get('Overall planned progress'),p.get('Overall actual progress'))}
    <div class="chips">{chip('CPI '+fmt(e['CPI'],'decimal'),index_rag(e['CPI']))}{chip('SPI '+fmt(e['SPI'],'decimal'),index_rag(e['SPI']))}</div>
    <p class="flag">{esc(flag(p))}</p><div class="card-foot">{asks} awaiting MD <span>· {opened} decisions open</span></div></article>'''

def decision_card(data,r,review):
    p=next(p for p in rows(data,REGISTRY) if p['Project code']==r['Project code'])
    rag=project_rag(p); due=as_date(r.get('Result due date'))
    return f'''<article class="ask-card"><div class="spread"><span class="eyebrow"><span class="dot {rag}"></span>{esc(p['Project name'])}</span>{chip('Overdue' if due and due<review else 'Awaiting MD','red' if due and due<review else 'amber')}</div>
    <h3>{esc(r.get('Decision / support required'))}</h3><p>{esc(r.get('Issue / blocker'))}</p>
    <div class="ask-foot"><b class="mono">{esc(money(r.get('Financial Impact (INR lakh)'),False))} impact</b><span>Result due {esc(fmt(r.get('Result due date'),'date'))}</span></div></article>'''

def full_fields(sheet,r,label='Every field'):
    spec=SCHEMAS.get(sheet,{'headers':[k for k in r if not k.startswith('_')],'types':{}})
    return f'<details class="full-fields"><summary>{esc(label)} · {len(spec["headers"])} fields</summary><dl>'+''.join(f'<div><dt>{esc(k)}</dt><dd data-field="{esc(k)}">{esc(fmt(r.get(k),spec["types"].get(k)))}</dd></div>' for k in spec['headers'])+'</dl></details>'

FIELDS={
 'MD Decisions':('Decision / support required','Decision ID','Issue / blocker','Recommended action'),
 'Construction & Resources':('Work front / structure','Front / activity ID','Reason for low / no execution','Recovery action required'),
 'Engineering Status':('Drawing / document description','Drawing / document reference','Reason pending / blocker','Action / MD support required'),
 'Procurement Status':('Item / package description','Procurement package / item ID','Reason pending / blocker','Action / MD support required'),
 'Invoicing Status':('Client name','Invoice / RA bill reference','Reason pending','Action / MD support required'),
 'Liability Status':('Contractor / vendor name','Liability / bill reference','Work affected if delayed','Action / MD support required'),
 'Lookahead Plan':('Planned activity','Front / activity ID','Constraint / support required','MD support required'),
 'Daily MOM':('Discussion point / issue','Item / action ID','Decision / instruction given','Agreed action')}

def record_card(sheet,r,review):
    title,ref,detail,action=FIELDS[sheet];rag,status,due=record_state(sheet,r,review)
    meta=[fmt(r.get(ref)),fmt(r.get('Site / location')) if r.get('Site / location') else '',fmt(r.get('Responsible owner') or r.get('Activity owner'))]
    extra=''
    if sheet=='Construction & Resources':
        extra=f'<p class="mono">Output {esc(fmt(r.get("Actual quantity achieved")))} / {esc(fmt(r.get("Weekly target quantity")))} {esc(r.get("Unit"))} planned</p>'
    if sheet=='Invoicing Status':
        extra=f'<p class="mono">To collect {esc(money(r.get("Outstanding balance (INR lakh)"),False))}</p><p>Expected certification: <b>{esc(fmt(r.get("Expected certification date"),"date"))}</b></p>'
    if sheet=='Liability Status': extra=f'<p class="mono">To pay {esc(money(r.get("Outstanding liability (INR lakh)"),False))}</p>'
    if sheet=='Lookahead Plan': extra=f'<p>{esc(fmt(r.get("Planned start"),"date"))} → {esc(fmt(r.get("Planned finish"),"date"))} · {esc(fmt(r.get("Planned quantity for the week")))} {esc(r.get("Unit"))}</p>'
    if r.get('Next Status'): extra+=f'<p>Next: <b>{esc(r["Next Status"])}</b> · {esc(fmt(r.get("Next Status Target Date"),"date"))}</p>'
    return f'''<article class="record-card" data-record="{esc(r.get(ref))}"><div class="spread">{chip(status,rag)}<small>{'Overdue · ' if rag=='red' and due and due<review else ''}{esc(fmt(due,'date')) if due else 'Date not confirmed'}</small></div>
    <h3>{esc(r.get(title))}</h3><p class="meta">{esc(' · '.join(x for x in meta if x))}</p>{extra}<p>{esc(r.get(detail))}</p><p><b>Action:</b> {esc(r.get(action))}</p>{full_fields(sheet,r)}</article>'''

REGISTER_ORDER=['MD Decisions','Construction & Resources','Engineering Status','Procurement Status','Invoicing Status','Liability Status','Lookahead Plan','Daily MOM']
REGISTER_IDS={s:'register-'+str(i) for i,s in enumerate(REGISTER_ORDER)}

def register_html(data,sheet,code,review):
    rs=ordered_records(data,sheet,code,review)
    return heading(sheet,f'{len(rs)} records · all statuses included',REGISTER_IDS[sheet])+'<div class="records-grid">'+(''.join(record_card(sheet,r,review) for r in rs) if rs else '<p>No records entered.</p>')+'</div>'
