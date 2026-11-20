import html

from datetime import datetime

from zoneinfo import ZoneInfo

import pandas as pd

import streamlit as st

import plotly.express as px

st.set_page_config(page_title='Jaipuria | Marketing Intelligence',layout='wide',initial_sidebar_state='collapsed')



import re

from pathlib import Path

import pandas as pd

import streamlit as st

import gspread

from google.oauth2.service_account import Credentials



BOOKS={'2026–28':'ly','2027–29':'cy'}

ROOT=Path(__file__).resolve().parent

SCOPE=['https://www.googleapis.com/auth/spreadsheets.readonly']

def client():

    c=Credentials.from_service_account_info(dict(st.secrets['gcp_service_account']),scopes=SCOPE)

    return gspread.authorize(c)

def num(x):

    return pd.to_numeric(x.astype(str).str.replace(',','',regex=False).str.replace('%','',regex=False),errors='coerce')

def dates(x):

    s=x.astype(str).str.strip(); out=pd.Series(pd.NaT,index=s.index,dtype='datetime64[ns]')

    for fmt in ["%b'%y",'%m/%d/%y %H:%M','%m/%d/%Y %H:%M:%S','%d-%m-%Y %H:%M:%S','%d-%m-%Y %H:%M','%d-%m-%Y','%Y-%m-%d %H:%M:%S','%Y-%m-%d']:

        m=out.isna(); out.loc[m]=pd.to_datetime(s.loc[m],format=fmt,errors='coerce')

    m=out.isna(); out.loc[m]=pd.to_datetime(s.loc[m],format='mixed',errors='coerce',dayfirst=False)

    return out



def frame(rows):

    if not rows:return pd.DataFrame()

    h=[str(v).strip() or f'blank_{i}' for i,v in enumerate(rows[0])]; return pd.DataFrame([list(r[:len(h)])+['']*max(0,len(h)-len(r)) for r in rows[1:]],columns=h)

def col(d,k):

    if k not in d: raise ValueError(f'Missing required field: {k}')

    return d[k]

def clean(s):return s.astype(str).str.strip()

def transactions(rows,batch):

    d=frame(rows); out=pd.DataFrame({'Batch':batch,'Date':dates(col(d,'Transaction Date')),'Source':clean(col(d,'Marketing Portal Source')).str.casefold(),'Medium':clean(col(d,'Marketing Portal Source Medium')),'Campaign':clean(col(d,'Marketing Portal Source Campaign'))}); out=out[out.Source.isin(['emailer','whatsapp']) & out.Date.notna()].copy(); out.Source=out.Source.map({'emailer':'Emailer','whatsapp':'WhatsApp'}); out.Medium=out.Medium.replace('', 'Not specified').map(lambda x:{'automation':'Automation','bulk campaign':'Bulk Campaign','bulk':'Bulk Campaign'}.get(x.casefold(),x)); out.Campaign=out.Campaign.replace('','Not specified'); return out



def email(rows,batch,kind):

    d=frame(rows)

    if d.empty:return pd.DataFrame(columns=['Batch','Type','Date','Campaign','Subject','Delivered','Opened','Clicked'])

    auto=kind=='Automation'; date=dates(col(d,'Date' if auto else 'Scheduled On')); delivered=num(col(d,'Delivered')); opened=num(col(d,'Opens (sorted descending)' if auto else 'Opens')); clicks=num(col(d,'Clicks'))

    if auto: opened=opened*delivered/100; clicks=clicks*delivered/100

    out=pd.DataFrame({'Batch':batch,'Type':kind,'Date':date,'Campaign':clean(col(d,'Automation Name' if auto else 'Campaign Name')),'Subject':clean(col(d,'Subject')),'Delivered':delivered,'Opened':opened,'Clicked':clicks})

    return out[out.Delivered.notna() & out.Campaign.ne('') & ~out.Campaign.str.casefold().isin(['total','grand total','summary']) & out.Date.notna()].copy()



def wa(rows,batch,kind):

    if len(rows)<=3:return pd.DataFrame(columns=['Batch','Type','Date','Campaign','Triggered','Delivered only','Read','Clicked','Failed','Delivered'])

    h=[list(r)+['']*40 for r in rows[:3]]; labels=[str(v).strip().casefold() for v in h[1]]; first=[str(v).strip().casefold() for v in h[0]]; idx=lambda name:labels.index(name)

    trig=first.index('total triggered'); converse=kind=='Converse Chat'; result=[]

    for row in rows[3:]:

        r=list(row)+['']*40

        if not str(r[0]).strip():continue

        name='Converse Chat' if converse else ' / '.join(dict.fromkeys(str(v).strip() for v in r[1:3] if str(v).strip()))

        result.append([batch,kind,r[0],name,r[trig],r[idx('delivered')],r[idx('read')],r[idx('clicked')],r[idx('failed')]])

    d=pd.DataFrame(result,columns=['Batch','Type','Date','Campaign','Triggered','Delivered only','Read','Clicked','Failed']); d.Date=dates(d.Date)

    for c in ['Triggered','Delivered only','Read','Clicked','Failed']:d[c]=num(d[c])

    d['Delivered']=d[['Delivered only','Read','Clicked']].sum(axis=1,min_count=1)

    return d[d.Date.notna() & d.Triggered.notna()].copy()



def senders(rows,batch):

    d=frame(rows)

    if d.empty:return pd.DataFrame(columns=['Batch','Sender','Type','Sent','Opened','Clicked','Bounced'])

    out=pd.DataFrame({'Batch':batch,'Sender':clean(col(d,'From User')),'Type':clean(col(d,'Email Type')),'Sent':num(col(d,'# Emails')),'Opened':num(col(d,'# Opened Emails')),'Clicked':num(col(d,'# Clicked Emails')),'Bounced':num(col(d,'# Bounced Emails'))}); return out[out.Sent.notna() & out.Sender.ne('') & ~out.Sender.str.casefold().isin(['total','grand total','summary'])]



@st.cache_data(ttl=60,show_spinner=False)

def load():

    try:

        gc=client()

    except Exception as exc:

        raise RuntimeError(f'Credential setup: {type(exc).__name__}: {exc!s}' ) from exc

    all_data={k:[] for k in ['transactions','email','wa','chat','senders']}

    for batch, secret_name in BOOKS.items():
        key = str(st.secrets['marketing_sheets'][secret_name]).strip()

        try:

            b=gc.open_by_key(key); tabs={ws.title:ws for ws in b.worksheets()}

        except (PermissionError, gspread.exceptions.SpreadsheetNotFound) as exc:

            raise RuntimeError(f'Batch {batch}: spreadsheet access denied or unavailable. Share the sheet with the client_email in credentials.json as Viewer. ({type(exc).__name__})') from exc

        except Exception as exc:

            raise RuntimeError(f'Batch {batch}: Google Sheets API connection failed ({type(exc).__name__}: {exc!s})') from exc

        # Only required transaction columns: one header read and one batch request.

        t=tabs['Transaction Details']; headers=t.row_values(1); fields=['Transaction Date','Marketing Portal Source','Marketing Portal Source Medium','Marketing Portal Source Campaign']; ranges=[]

        for f in fields:

            j=headers.index(f)+1; letter=re.sub(r'\d','',gspread.utils.rowcol_to_a1(1,j)); ranges.append(f'{letter}2:{letter}{t.row_count}')

        values=t.batch_get(ranges); n=max(map(len,values),default=0); rows=[fields]+[[values[j][i][0] if i<len(values[j]) and values[j][i] else '' for j in range(4)] for i in range(n)]

        all_data['transactions'].append(transactions(rows,batch))

        names=['Email Automation','Email Bulk','Whatsapp Automation Summary','Whatsapp Bulk Summary','Converse Chat Summary','Total Emails Sent' if batch=='2026–28' else 'Total Email Sent']

        # One request for summary tabs; no student-level detail fetched.

        vals=b.values_batch_get([f"'{name}'!A:{gspread.utils.rowcol_to_a1(1,min(tabs[name].col_count,27))[:-1]}" for name in names],params={'valueRenderOption':'FORMATTED_VALUE'})['valueRanges']

        r=[v.get('values',[]) for v in vals]

        all_data['email'] += [email(r[0],batch,'Automation'),email(r[1],batch,'Bulk')]

        all_data['wa'] += [wa(r[2],batch,'Automation'),wa(r[3],batch,'Bulk')]

        all_data['chat'].append(wa(r[4],batch,'Converse Chat')); all_data['senders'].append(senders(r[5],batch))

    return {k:pd.concat(v,ignore_index=True) for k,v in all_data.items()}



LOGO="data:image/png;base64,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"

# Presentation layer: one deterministic responsive HTML grid, native Streamlit filters.
# Google Sheets reader and processing functions above are unchanged.
import math

B = list(BOOKS)
GREEN, PURPLE, ORANGE = '#204B39', '#8053AD', '#ECA13B'
EM = ['Delivered', 'Opened', 'Clicked']
WM = ['Triggered', 'Delivered', 'Read', 'Clicked', 'Failed']
SM = ['Sent', 'Opened', 'Clicked', 'Bounced']
PAGES = ['Overview', 'Transactions', 'Email', 'WhatsApp', 'Converse Chat', 'Senders']
ICONS = {
    'grid': '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
    'chart': '<path d="M3 20h18M6 20V12h3v8M11 20V8h3v12M16 20V4h3v16"/>',
    'mail': '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 6 9 7 9-7"/>',
    'chat': '<path d="M21 11.5a9 9 0 0 1-9 9 10 10 0 0 1-4-.9L3 21l1.4-4.6A9 9 0 1 1 21 11.5Z"/><path d="M8 12h.01M12 12h.01M16 12h.01" stroke-width="3"/>',
    'wa': '<path d="M21 11a9 9 0 0 1-9 10 10 10 0 0 1-4-.9L3 21l1.4-4.6A9 9 0 1 1 21 11Z"/><path d="M8 7c-2 2 2 9 7 10l2-2-3-2-1 1-3-3 1-1-2-3Z"/>',
    'users': '<circle cx="10" cy="7" r="3"/><path d="M3 21v-3a6 6 0 0 1 6-6h2a6 6 0 0 1 6 6v3ZM17 5a3 3 0 0 1 0 6m2 3a5 5 0 0 1 2 4v3"/>',
    'layers': '<path d="m12 3 10 5-10 5L2 8Zm-10 9 10 5 10-5M2 16l10 5 10-5"/>',
    'bulb': '<path d="M9 18h6m-5 3h4M9 15c0-2-4-3-4-7a7 7 0 0 1 14 0c0 4-4 5-4 7"/>',
    'send': '<path d="m22 2-7 20-4-9-9-4 20-7ZM11 13 22 2"/>',
    'check': '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
    'calendar': '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4m10-4v4M3 10h18"/>',
    'trophy': '<path d="M8 3h8v6a4 4 0 0 1-8 0ZM8 5H4v2a4 4 0 0 0 4 4m8-6h4v2a4 4 0 0 1-4 4m-4 2v6m-4 2h8"/>',
    'pie': '<path d="M12 3v9h9A9 9 0 1 1 12 3Z"/><path d="M15 2v7h7a8 8 0 0 0-7-7Z"/>',
    'table': '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M3 14h18M10 9v11"/>',
}
CSS = '''
:root{color-scheme:light;--ink:#172d24;--green:#204b39;--muted:#65766b;--cream:#f7f7f3;--line:#dfe5da;--purple:#8053ad;--orange:#eca13b;--lime:#edf2df}
[data-testid="stSidebar"],[data-testid="stSidebarCollapsedControl"],[data-testid="stHeader"],[data-testid="stToolbar"],#MainMenu,footer{display:none!important}
.stApp{background:var(--cream)!important;color:var(--ink)!important}.block-container{max-width:1680px!important;width:100%!important;padding:14px 28px 30px!important;box-sizing:border-box}
[data-testid="stVerticalBlock"]{gap:.65rem!important}[data-testid="stHorizontalBlock"]{gap:12px!important}[data-testid="stColumn"]{min-width:0!important}
.mi,.mi-top{font-family:Arial,Helvetica,sans-serif;color:var(--ink);font-size:14px;line-height:1.45;box-sizing:border-box;font-variant-numeric:tabular-nums}.mi *,.mi-top *{box-sizing:border-box}.mi p,.mi h2,.mi h3,.mi-top p{margin:0!important}.mi-icon{flex:none;vertical-align:middle}
.brand{display:flex;align-items:center;gap:22px;padding:4px 0 12px;border-bottom:2px solid #b9cc8a}.brand img{width:174px;height:auto;mix-blend-mode:multiply}.brand-text{flex:1;min-width:0;border-left:1px solid #d0d9ca;padding-left:22px}.brand-name{font-weight:750;font-size:clamp(23px,2.1vw,32px);letter-spacing:-.6px;line-height:1.2;color:#173b2c}.brand-sub{font-size:12px;color:var(--muted);margin-top:5px}.brand-badge{font-size:11px;font-weight:700;background:#e8eedb;border-radius:30px;padding:8px 13px;white-space:nowrap;color:var(--green)}
.mi-nav{display:flex;align-items:center;gap:7px;flex-wrap:wrap;padding:1px 0 3px}.mi-nav a{display:inline-flex;align-items:center;gap:7px;color:var(--green)!important;background:white;text-decoration:none!important;border:1px solid #d9e1d7;border-radius:7px;font-size:13px;font-weight:700;padding:8px 12px;line-height:1.3}.mi-nav a:hover,.mi-nav a:focus{background:#e6edda!important}.mi-nav a:first-child{background:#e8eedb!important}
[data-testid="stWidgetLabel"] p{color:#42584a!important;font-size:12px!important;font-weight:700!important}[data-baseweb="input"],[data-baseweb="select"]>div{background:white!important;border:1px solid #d9e1d7!important;border-radius:7px!important}[data-testid="stButton"] button{background:#204b39!important;border:1px solid #204b39!important;color:white!important;border-radius:7px!important;font-weight:650!important;min-height:38px!important}[data-testid="stButton"] button:hover{background:#2c6048!important}
.filter-note{font-size:11px;color:#64766b;margin:1px 0 3px;line-height:1.4}.statusline{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;color:#566e60;font-size:11px;padding:2px 0 8px;border-bottom:1px solid #dce4d8}.statusline b{color:#386c48}
.mi{display:flex;flex-direction:column;gap:22px;padding-top:8px}.section{scroll-margin-top:12px;min-width:0}.sectitle{display:flex;align-items:center;gap:9px;color:#173b2c!important;font-size:22px!important;font-weight:750!important;letter-spacing:-.4px;line-height:1.25;margin:0 0 11px!important;padding:0!important}.mi .sectitle{margin:0 0 12px!important}.section:not(:first-child){border-top:1px solid #dce3d8;padding-top:17px}.sectitle .mi-icon{color:var(--green)}
.section-grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(250px,26%);gap:14px;align-items:start}.leftstack,.rightstack{display:grid;gap:12px;min-width:0;align-content:start}.split{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;align-items:stretch}.txcharts{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(0,.85fr) minmax(0,1fr);gap:12px;align-items:stretch}.rankcharts{display:grid;grid-template-columns:minmax(0,1.25fr) minmax(0,1fr);gap:12px;align-items:stretch}.chartstack{display:grid;gap:12px;min-width:0}.panel{background:#fff;border:1px solid var(--line);border-radius:10px;padding:12px;min-width:0;box-shadow:0 1px 3px #193d2c07;overflow:hidden}.panel h3,.insight h3{display:flex;align-items:center;gap:8px;color:#173b2c!important;font-weight:750!important;font-size:15px!important;line-height:1.3;margin:0 0 9px!important;padding:0!important}.panel-sub{font-size:11px;color:var(--muted);margin:-2px 0 8px!important}.panel-fill{display:flex;flex-direction:column;justify-content:flex-start}.panel-fill .plot{margin-top:auto;margin-bottom:auto}
.kpis{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;min-width:0}.kpis.four{grid-template-columns:repeat(4,minmax(0,1fr))}.kpis.five{grid-template-columns:repeat(5,minmax(0,1fr))}.kpi{background:#fff;border:1px solid var(--line);border-radius:10px;padding:13px;min-width:0;min-height:102px;display:flex;flex-direction:column;justify-content:space-between;box-shadow:0 1px 3px #193d2c07}.kpi.primary{background:#204b39;color:#fff;border-color:#204b39}.kpi.accent{background:#f2eafa;border-color:#e6ddf0}.klabel{display:flex;gap:8px;align-items:center;font-weight:700;font-size:12px;line-height:1.25;min-height:25px;margin-bottom:8px!important}.kpair{display:grid;grid-template-columns:1fr 1fr;gap:8px}.kpair>div{min-width:0}.kpair>div+div{border-left:1px solid #dce4d8;padding-left:10px}.primary .kpair>div+div{border-color:#72917e}.kval{font-size:clamp(19px,1.9vw,30px);font-weight:750;letter-spacing:-.7px;line-height:1.08;white-space:nowrap}.kpair .kval{font-size:clamp(17px,1.65vw,26px)}.ksub{font-size:10px;line-height:1.2;margin-top:5px;color:#62766a}.primary .ksub{color:#dfeadf}.five .kpi,.four .kpi{min-height:89px;padding:11px}.five .kval,.four .kval{font-size:clamp(17px,1.55vw,25px)}.overview-grid{align-items:stretch}.overview-grid .kpis{grid-auto-rows:1fr}.overview-grid .kpi{min-height:108px}.overview-grid .leftstack{align-content:stretch}.overview-grid .rightstack{align-content:stretch}.overview-grid .insight{height:auto}.overview-grid .insight-row{padding:6px 0}.overview-grid .insight-row b{display:inline;margin-right:5px}.overview-grid .rightstack{gap:8px}
.insight{background:#edf2df;border:1px solid #dce6c8;border-radius:10px;padding:13px 15px;min-width:0;color:#213b2c}.insight h3{font-size:16px!important;padding-bottom:8px!important;border-bottom:1px solid #d4dec2}.insight-row{display:flex;align-items:flex-start;gap:8px;padding:8px 0;border-bottom:1px solid #d9e2cb;font-size:12px;line-height:1.38}.insight-row:last-child{border:0;padding-bottom:0}.insight-row b{display:block;font-size:12px;color:#193c2c;margin-bottom:2px}.insight-row div{min-width:0;overflow-wrap:anywhere}.scope{font-size:11px;color:#68547c;background:#f3edf8;border:1px solid #e7ddf0;border-radius:8px;padding:9px 11px;line-height:1.35}
.tablebox{overflow:auto;border-radius:6px;max-width:100%;background:#fff}.report{width:100%!important;border-collapse:collapse!important;border:0!important;border-spacing:0;font-size:12px!important;line-height:1.32;table-layout:auto;color:#263d31!important;margin:0!important;background:#fff}.report th,.report td{border:0!important;border-bottom:1px solid #e9ede6!important;padding:8px 7px!important;vertical-align:middle!important}.report th{font-size:10px!important;font-weight:750!important;color:#3a5844!important;background:#edf1e7!important;white-space:normal;line-height:1.2}.report .txt{text-align:left!important;white-space:normal;overflow-wrap:anywhere}.report .num{text-align:right!important;white-space:nowrap;font-variant-numeric:tabular-nums}.report tr:nth-child(even) td{background:#f7f9f5}.report tr.ly td{background:#eef3e1}.report tr.cy td{background:#f2eafa}.report tr.total td{background:#e2ebd1;font-weight:750}.report tr:last-child td{border-bottom:0!important}.report .namecell{min-width:112px}.report .secondary{display:block;color:#697970;font-size:10px;line-height:1.3;margin-top:3px;overflow-wrap:anywhere}.report .rank{width:35px;text-align:center!important}.ranknum{display:inline-block;background:#e8efdd;color:#204b39;border-radius:50%;min-width:23px;height:23px;text-align:center;line-height:23px;font-weight:700}.summary-table .report{font-size:11.5px!important}.summary-table .report th,.summary-table .report td{padding:9px 6px!important}.detail-row{margin-top:12px;display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.details{background:#fff;border:1px solid var(--line);border-radius:8px;min-width:0;overflow:hidden}.details summary{cursor:pointer;font-size:12px;font-weight:650;color:#335745;padding:10px 13px;line-height:1.4}.details summary:hover{background:#f0f3e9}.details[open]>summary{border-bottom:1px solid var(--line);background:#f4f6ef}.details .tablebox{max-height:320px;padding:1px 8px 8px}.details .report th{position:sticky;top:0;z-index:1}.details .empty{margin:8px}.empty{background:#f5f1f8;border:1px solid #e8e0ef;border-radius:7px;padding:11px;color:#68547c;font-size:12px;min-height:40px}
.plot{width:100%;min-width:0;overflow:hidden}.plot svg{display:block;width:100%;height:auto;overflow:visible}.plot text{font-family:Arial,Helvetica,sans-serif}.chartstack .plot svg{height:125px;width:100%}.rankcharts>.panel{align-self:stretch}.rankcharts .chartstack .panel{padding:10px 12px}.chart-legend{display:flex;gap:11px;flex-wrap:wrap;justify-content:center;color:#52675a;font-size:10px;margin:1px 0 5px}.legenditem{display:inline-flex;align-items:center;gap:5px}.legenddot{width:9px;height:9px;border-radius:2px;display:inline-block}.footnote{margin-top:8px;font-size:11px;color:#65766b;line-height:1.4}.footer{padding:12px 0 0;border-top:1px solid var(--line);font-size:11px;color:#65766b}
@media(max-width:1250px){.block-container{padding:12px 18px 24px!important}.section-grid{grid-template-columns:minmax(0,1fr) minmax(215px,25%);gap:11px}.panel{padding:10px}.txcharts,.split,.rankcharts,.kpis{gap:9px}.kval{font-size:23px}.kpair .kval{font-size:19px}.five .kval,.four .kval{font-size:18px}.report th,.report td{padding:6px 5px!important}.report{font-size:11px!important}.insight{padding:11px}.insight-row{font-size:11px}}
@media(max-width:950px){.section-grid{grid-template-columns:1fr}.overview-grid .insight{height:auto}.overview-grid .insight-row{padding:6px 0}.overview-grid .insight-row b{display:inline;margin-right:5px}.overview-grid .rightstack{gap:8px}.rightstack{grid-template-columns:1fr}.txcharts{grid-template-columns:1fr 1fr}.txcharts>.panel:first-child{grid-column:1/-1}.brand img{width:135px}.brand-badge{display:none}.kpis.five{grid-template-columns:repeat(3,minmax(0,1fr))}.rankcharts{grid-template-columns:1fr 1fr}}
@media(max-width:600px){.brand-text{min-width:0}.klabel{font-size:11px}.kval{font-size:20px}.kpair .kval{font-size:16px}.block-container{padding:10px 12px 20px!important}.brand{gap:12px;flex-wrap:wrap}.brand-text{border:0;padding:0}.brand-name{font-size:22px}.kpis,.kpis.four,.kpis.five{grid-template-columns:repeat(2,minmax(0,1fr))}.split,.rankcharts,.txcharts,.detail-row{grid-template-columns:1fr}.txcharts>.panel:first-child{grid-column:auto}.kpair{gap:4px}.kpair>div+div{padding-left:5px}.kpair .kval{font-size:17px}.sectitle{font-size:19px!important}}

/* Shared row alignment for report tables and insight panels. */
.mi{gap:18px;padding-top:5px}.section:not(:first-child){padding-top:14px}.mi .sectitle{margin:0 0 10px!important;font-size:20px!important;gap:9px!important;padding:7px 12px!important;border:1px solid #dce5d8;border-left:4px solid #eca13b;border-radius:8px;background:linear-gradient(90deg,#e9efdf,#fff 72%);line-height:1.25}.mi .sectitle .mi-icon{width:20px;height:20px}
.section-grid{grid-template-columns:minmax(0,1fr) minmax(250px,23%);align-items:stretch;gap:12px}.leftstack,.rightstack{display:flex;flex-direction:column;gap:10px;min-width:0;align-content:normal}.aligned-main{display:flex;flex-direction:column;gap:10px;flex:1;min-width:0}.aligned-main>.panel{flex:1;display:flex;flex-direction:column}.aligned-main>.panel>.tablebox{flex:1;display:flex;align-items:stretch}.aligned-main>.panel>.tablebox>.report{flex:1;width:100%;height:auto}.section-grid>.rightstack>.insight{flex:1}.section-before,.section-after{display:grid;gap:10px;min-width:0}.section-before{margin-bottom:10px}.section-after{margin-top:10px}.split,.txcharts{gap:10px}.panel{padding:11px;border-radius:9px;box-shadow:none}.panel h3,.insight h3{font-size:14px!important;margin:0 0 8px!important;gap:7px!important}.panel-sub{font-size:11px;margin:0 0 6px!important}.panel-fill .plot{margin:auto 0}.panel .plot svg{max-height:225px;width:100%}.txcharts .plot svg{height:190px;width:100%;max-height:190px}.section-after .plot svg{height:195px;width:100%;max-height:195px}.chart-legend{margin:1px 0 3px}.kpis{gap:10px}.kpi{border-radius:9px;box-shadow:none;padding:12px;min-height:91px}.klabel{min-height:20px;margin-bottom:7px!important}.overview-grid .leftstack{display:flex}.overview-grid .kpis{flex:1;grid-auto-rows:1fr}.overview-grid .kpi{min-height:100px}.overview-grid .rightstack{gap:0}.overview-grid .insight{height:auto}.four .kpi,.five .kpi{min-height:86px}
/* Independent reference-style insight tiles within a white outer frame. */
.insight{background:#fff;border:1px solid #dbe4d7;border-radius:9px;padding:10px;display:flex;flex-direction:column;gap:0;box-shadow:none;min-width:0}.insight h3{color:#204b39!important;border:0;padding:0 2px 2px!important;font-size:15px!important;margin:0 0 7px!important}.insight h3 .mi-icon{width:18px;height:18px}.insight-items{display:grid;grid-auto-rows:auto;align-content:stretch;gap:5px;flex:1;min-height:0}.insight-row,.overview-grid .insight-row{display:flex;align-items:center;gap:0;padding:6px 9px!important;border:1px solid #dfebdf!important;border-left:3px solid #39866b!important;border-radius:7px;background:#eef6ef;font-size:11.5px;line-height:1.3;color:#284636;min-width:0}.insight-row:nth-child(4n+2){background:#f7f2fa;border-color:#e8dbef!important;border-left-color:#8053ad!important}.insight-row:nth-child(4n+3){background:#fff6e9;border-color:#f1e2c8!important;border-left-color:#eca13b!important}.insight-row:nth-child(4n+4){background:#f5f7ed;border-color:#e2e9d4!important;border-left-color:#97ac65!important}.insight-row:last-child{padding-bottom:6px!important}.insight-row b,.overview-grid .insight-row b{display:block!important;font-size:10.5px;font-weight:750;color:#536c5c;margin:0 0 2px!important}.insight-row .insight-value{display:block;color:#203b32;font-size:11.5px;font-weight:600;line-height:1.32;overflow-wrap:anywhere}.insight-note{font-size:10px;color:#6b557f;background:#f5eff9;border-radius:5px;padding:5px 7px;margin-top:6px;line-height:1.3}.scope{padding:7px 9px;font-size:11px}
.tablebox{border:1px solid #e2e7df;border-radius:6px}.report{font-size:12px!important;line-height:1.3}.report th,.report td{padding:7px 8px!important;border-right:1px solid #e6ebe2!important;border-bottom:1px solid #e1e7dd!important}.report th:last-child,.report td:last-child{border-right:0!important}.report th{background:#e8eee2!important;font-size:11px!important;color:#284b38!important}.report tr.ly td{background:#edf4e6}.report tr.cy td{background:#f2eaf9}.report tr.total td{background:#e2ebce;font-weight:750;border-top:1px solid #b9caa8!important}.report .secondary{font-size:10px;line-height:1.25}.summary-table .report{font-size:12px!important}.summary-table .report th,.summary-table .report td{padding:7px 8px!important}.report .rank{width:35px}.ranknum{height:21px;min-width:21px;line-height:21px}.detail-row{margin-top:9px;gap:10px}.details summary{padding:8px 11px}.section-grid .panel{min-height:0}.overview-grid .kpis{height:100%}.aligned-main>.converse-pair{flex:1}.converse-pair>.panel{display:flex;flex-direction:column}.converse-pair>.panel>.tablebox{flex:1;display:flex;align-items:stretch}.converse-pair>.panel>.tablebox>.report{flex:1}.converse-pair .plot svg{height:195px}.converse-pair{grid-template-columns:minmax(0,.9fr) minmax(0,1.1fr)}.kpi .ksingle{margin-top:auto}
@media(max-width:1250px){.section-grid{grid-template-columns:minmax(0,1fr) minmax(220px,25%)}.insight{padding:8px}.insight-row,.overview-grid .insight-row{padding:5px 7px!important}.report th,.report td,.summary-table .report th,.summary-table .report td{padding:6px 5px!important}.report,.summary-table .report{font-size:11px!important}.report th{font-size:10px!important}}
@media(max-width:950px){.section-grid{grid-template-columns:1fr}.overview-grid .kpis{height:auto}.insight-items{grid-template-columns:repeat(2,minmax(0,1fr));grid-auto-rows:auto}.section-grid>.rightstack>.insight{flex:initial}.section-grid .aligned-main>.panel>.tablebox{display:block}.section-grid .aligned-main>.panel>.tablebox>.report{height:auto}}
@media(max-width:600px){.converse-pair{grid-template-columns:1fr}.insight-items{grid-template-columns:1fr}.mi .sectitle{font-size:17px!important}.report th,.report td{padding:5px!important}.report th{font-size:10px!important}.report .namecell{min-width:120px}}

'''
st.markdown('<style>'+CSS+'</style>', unsafe_allow_html=True)

def esc(v): return html.escape(str(v))
def fmt(v): return '—' if pd.isna(v) else f'{v:,.0f}'
def pct(v): return '—' if pd.isna(v) else f'{v:.2f}%'
def ratio(n,d): return n/d*100 if pd.notna(n) and pd.notna(d) and d>0 else float('nan')
def rate(n,d): return n.div(d.where(d>0))*100
def aggregate(d, keys, metrics, den='Delivered', numer='Opened'):
    if d.empty: return pd.DataFrame(columns=keys+metrics+['Rate','Click rate'])
    a=d.groupby(keys,dropna=False)[metrics].sum(min_count=1).reset_index()
    a['Rate']=rate(a[numer],a[den])
    if 'Clicked' in a: a['Click rate']=rate(a.Clicked,a[den])
    return a

def monthly(d): return d.assign(Month=pd.to_datetime(d.Date,errors='coerce').dt.strftime('%Y-%m'))
def monthlabel(v):
    try: return pd.to_datetime(v,format='%Y-%m').strftime('%b %Y')
    except (ValueError,TypeError): return str(v)
def pair(d,c): return d.groupby('Batch')[c].sum(min_count=1).reindex(B).tolist() if not d.empty else [float('nan')]*2
def pairrate(d,n,den): return [ratio(a,b) for a,b in zip(pair(d,n),pair(d,den))]
def total(d,c): return d[c].sum(min_count=1) if c in d else float('nan')
def change(a,b): return f'{b-a:+.2f} pp' if pd.notna(a) and pd.notna(b) else 'Comparison unavailable'
def icon(name,color=GREEN,size=20): return f'<svg class="mi-icon" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{ICONS[name]}</svg>'
def div(body,cls): return f'<div class="{cls}">{body}</div>'
def note(s): return div(esc(s),'scope')
def empty(s='No data available for this selection.'): return div(esc(s),'empty')
def panel(title,body,ico='chart',sub='',extra=''):
    return div('<h3>'+icon(ico,ORANGE if ico=='trophy' else GREEN,18)+esc(title)+'</h3>'+ (div(esc(sub),'panel-sub') if sub else '')+body,'panel '+extra)
def insight(items,heading='Key Insights',foot=''):
    tiles=''.join('<div class="insight-row"><div><b>'+esc(k)+'</b><span class="insight-value">'+esc(v)+'</span></div></div>' for k,v in items)
    return '<aside class="insight"><h3>'+icon('bulb',GREEN,18)+esc(heading)+'</h3>'+div(tiles,'insight-items')+(div(esc(foot),'insight-note') if foot else '')+'</aside>'
def kpi(label,values,ico='chart',style='',percent=False,sub='Overall'):
    f=pct if percent else fmt; co='#d5eca3' if style=='primary' else PURPLE if style=='accent' else ORANGE if ico=='mail' else GREEN
    def value(v,l): return div(esc(f(v)),'kval')+div(esc(l),'ksub')
    content=div(''.join(div(value(v,b),'') for b,v in zip(B,values)),'kpair') if isinstance(values,(list,tuple)) else div(value(values,sub),'ksingle')
    return div(div(icon(ico,co,22)+esc(label),'klabel')+content,'kpi '+style)
def cards(items,layout=''): return div(''.join(items),'kpis '+layout)
def split(*items,cls='split'): return div(''.join(items),cls)
def section(i,title,ico,left,right,details='',overview=False,before='',after=''):
    grid=div(div(div(left,'aligned-main') if not overview else left,'leftstack')+div(right,'rightstack'),'section-grid'+(' overview-grid' if overview else ''))
    return f'<section class="section" id="section-{i}"><h2 class="sectitle">{icon(ico,GREEN,20)}{esc(title)}</h2>'+ (div(before,'section-before') if before else '')+grid+(div(after,'section-after') if after else '')+(div(details,'detail-row') if details else '')+'</section>'
def count_table(t,keys): return t.groupby(keys+['Batch']).size().unstack(fill_value=0).reindex(columns=B,fill_value=0).reset_index()
def compare_table(a,kind,metrics):
    z=a[a.Type.eq(kind)].set_index('Batch').reindex(B).rename_axis('Batch').reset_index()
    return z[['Batch']+metrics].rename(columns={'Rate':'Open rate' if 'Opened' in metrics else 'Read rate'})
def ranked(d,keys,metrics,den='Delivered',numer='Opened'):
    r=aggregate(d,keys,metrics,den,numer).dropna(subset=['Rate']).sort_values(['Rate',den],ascending=False).reset_index(drop=True)
    r.insert(0,'Rank',range(1,len(r)+1));return r

TEXT={'batch','month','source','medium','campaign','subject','sender','type','campaign / template','campaign / subject'}
def table(d,compact=False):
    if d.empty:return empty()
    d=d.copy()
    def isnum(c): return c.casefold() not in TEXT and (pd.api.types.is_numeric_dtype(d[c]) or 'rate' in c.casefold() or c in B or c.casefold() in {'rank','triggered','delivered','delivered only','read','opened','clicked','failed','sent','bounced','transactions'})
    def cell(c,v):
        if isinstance(v,tuple):
            value=esc(v[0])+(f'<span class="secondary">{esc(v[1])}</span>' if v[1] else '')
            return f'<td class="txt namecell">{value}</td>'
        numeric=isnum(c);cl='num' if numeric else 'txt'
        if c=='Rank':return f'<td class="rank"><span class="ranknum">{fmt(v)}</span></td>'
        if c.casefold() in ('campaign','sender','subject','campaign / template','campaign / subject'):cl+=' namecell'
        val=pct(v) if 'rate' in c.casefold() else fmt(v) if numeric else '—' if pd.isna(v) else monthlabel(v) if c=='Month' else str(v)
        return f'<td class="{cl}">{esc(val)}</td>'
    h=''.join(f'<th class="{"num" if isnum(c) else "txt"}">{esc(c)}</th>' for c in d.columns)
    rows=[]
    for _,r in d.iterrows():
        cl='ly' if r.get('Batch')==B[0] else 'cy' if r.get('Batch')==B[1] else 'total' if str(r.iloc[0]).casefold() in ('total','overall') else ''
        rows.append(f'<tr class="{cl}">'+''.join(cell(c,v) for c,v in r.items())+'</tr>')
    return div('<table class="report"><thead><tr>'+h+'</tr></thead><tbody>'+''.join(rows)+'</tbody></table>','tablebox'+(' summary-table' if compact else ''))
def detail(title,d): return '<details class="details"><summary>'+esc(title)+'</summary>'+table(d)+'</details>'

def brush(v):
    if not pd.notna(v): return '—'
    if abs(v)>=1e6:return f'{v/1e6:.1f}m'
    if abs(v)>=1000:return f'{v/1000:.1f}k'
    return f'{v:g}'
def axis_max(v):
    if not pd.notna(v) or v<=0:return 1
    step=10**math.floor(math.log10(v/4)); result=math.ceil(v/step)*step
    for multiple in [1,2,2.5,5,10]:
        tick=step*multiple
        if math.ceil(v/tick)<=5:return math.ceil(v/tick)*tick
    return result

def svg_chart(d,x,y,group='Batch',kind='bar',width=480,height=210,secondary=None):
    """Responsive, accessible SVG chart. All chart coordinates derived from export data."""
    if d.empty or y not in d or d[y].notna().sum()==0:return empty('No chart data for this selection.')
    z=d.copy();z=z[z[x].notna()]; cats=sorted(z[x].unique()) if x=='Month' else list(dict.fromkeys(z[x])); groups=list(dict.fromkeys(z[group]));
    if group=='Batch':groups=[b for b in B if b in groups]
    colors={B[0]:GREEN,B[1]:PURPLE,'Delivered':GREEN,'Opened':ORANGE,'Read':ORANGE,'Sent':GREEN}
    colors.update({k:colors.get(k,[GREEN,PURPLE,ORANGE][i%3]) for i,k in enumerate(groups)})
    legends=''.join(f'<span class="legenditem"><i class="legenddot" style="background:{colors[k]}"></i>{esc(k)}</span>' for k in groups)
    if secondary is not None:legends+=f'<span class="legenditem"><i class="legenddot" style="background:{PURPLE}"></i>Read rate</span>'
    # Fixed viewbox with generous axis margins. No clipped tick labels at narrow widths.
    israte='rate' in y.casefold();left=43;right=46 if secondary is not None else 13;top=14; bottom=49 if x=='Month' and len(cats)>6 else 30
    plotw=width-left-right; ploth=height-top-bottom; y0=top+ploth
    maxv=float(z[y].max()); cap=axis_max(maxv*1.12 if kind=='bar' else maxv*1.08)
    if israte:cap=min(100,max(5,cap)) if maxv<=100 else cap
    cap=max(cap,1); bucket=plotw/max(len(cats),1)
    def xx(i):return left+bucket*(i+.5)
    def yy(v):return y0-float(v)/cap*ploth
    def text(xc,yc,t,anchor='middle',size=10,color='#52685a',weight='normal',transform=''):
        return f'<text x="{xc:.2f}" y="{yc:.2f}" text-anchor="{anchor}" font-size="{size}" font-weight="{weight}" fill="{color}" {transform}>{esc(t)}</text>'
    sv=[]
    for i in range(5):
        v=cap*i/4; pos=yy(v);sv.append(f'<line x1="{left}" y1="{pos:.2f}" x2="{width-right}" y2="{pos:.2f}" stroke="#e4eae1" stroke-width="1"/>');sv.append(text(left-6,pos+3,brush(v)+('%' if israte else ''),'end',9))
        if secondary is not None:sv.append(text(width-right+5,pos+3,f'{i*25}%','start',9))
    stride=max(1,math.ceil(len(cats)/10))
    for i,c in enumerate(cats):
        if i%stride and i!=len(cats)-1:continue
        lab=monthlabel(c) if x=='Month' else str(c)
        if x=='Month' and len(cats)>6:sv.append(text(xx(i),y0+13,lab,'end',9,transform=f'transform="rotate(-38 {xx(i):.2f} {y0+13:.2f})"'))
        else:sv.append(text(xx(i),y0+15,lab,'middle',9))
    sv.append(f'<line x1="{left}" y1="{y0}" x2="{width-right}" y2="{y0}" stroke="#aebdad"/>')
    look={(r[x],r[group]):r[y] for _,r in z.iterrows()}
    barw=min(55,bucket*.73/max(1,len(groups)))
    for j,g in enumerate(groups):
        vals=[look.get((c,g),float('nan')) for c in cats]; points=[]
        for i,v in enumerate(vals):
            if pd.isna(v):
                # Break line across unavailable observations, not a fabricated zero.
                if kind=='line' and points:
                    if len(points)>1:sv.append(f'<polyline points="{" ".join(points)}" fill="none" stroke="{colors[g]}" stroke-width="2.3" stroke-linejoin="round"/>')
                    points=[]
                continue
            cy=yy(v);cx=xx(i);tooltip=esc(f'{monthlabel(cats[i]) if x=="Month" else cats[i]} · {g}: {pct(v) if israte else fmt(v)}')
            if kind=='bar':
                bx=cx+(j-(len(groups)-1)/2)*barw-barw*.46
                sv.append(f'<rect x="{bx:.2f}" y="{cy:.2f}" width="{barw*.92:.2f}" height="{max(0,y0-cy):.2f}" rx="2" fill="{colors[g]}"><title>{tooltip}</title></rect>')
                if len(cats)<=3:sv.append(text(bx+barw*.46,cy-5,brush(float(v)) if x=='Month' and maxv>10000 else fmt(v),'middle',10,colors[g],'600'))
            else:
                points.append(f'{cx:.2f},{cy:.2f}');sv.append(f'<circle cx="{cx:.2f}" cy="{cy:.2f}" r="3.5" fill="white" stroke="{colors[g]}" stroke-width="2"><title>{tooltip}</title></circle>')
        if kind=='line' and len(points)>1:sv.append(f'<polyline points="{" ".join(points)}" fill="none" stroke="{colors[g]}" stroke-width="2.3" stroke-linejoin="round"/>')
    if secondary is not None:
        positions=[]
        for i,c in enumerate(cats):
            v=secondary.get(c,float('nan'))
            if pd.isna(v):continue
            a=xx(i);b=y0-float(v)/100*ploth;positions.append(f'{a:.2f},{b:.2f}');sv.append(f'<circle cx="{a:.2f}" cy="{b:.2f}" r="3.6" fill="white" stroke="{PURPLE}" stroke-width="2"/>');sv.append(text(a,max(top+5,b-8),pct(v),'middle',10,PURPLE,'700'))
        if len(positions)>1:sv.append(f'<polyline points="{" ".join(positions)}" fill="none" stroke="{PURPLE}" stroke-width="2"/>')
    svg=f'<svg viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="{esc(y)} by {esc(x)}">'+''.join(sv)+'</svg>'
    return div(div(legends,'chart-legend')+svg,'plot')

def donut(d):
    if d.empty or d.Transactions.sum()==0:return empty('No LY transactions.')
    nums=d.groupby('Source').Transactions.sum(); tot=nums.sum(); r=53; circ=2*math.pi*r; offset=0; sv=[];leg=[]
    for source in ['Emailer','WhatsApp']:
        n=nums.get(source,0); co=GREEN if source=='Emailer' else PURPLE; arc=circ*n/tot
        if arc:sv.append(f'<circle cx="100" cy="73" r="{r}" fill="none" stroke="{co}" stroke-width="23" stroke-dasharray="{arc:.3f} {circ-arc:.3f}" stroke-dashoffset="{-offset:.3f}" transform="rotate(-90 100 73)"><title>{esc(source)}: {fmt(n)} ({pct(n/tot*100)})</title></circle>')
        offset+=arc;leg.append(f'<span class="legenditem"><i class="legenddot" style="background:{co}"></i>{esc(source)} {pct(n/tot*100)}</span>')
    svg='<svg viewBox="0 0 200 148" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Last year channel share">'+''.join(sv)+f'<text x="100" y="71" font-size="25" font-weight="750" text-anchor="middle" fill="{GREEN}">{fmt(tot)}</text><text x="100" y="86" font-size="10" text-anchor="middle" fill="#65766b">Transactions</text></svg>'
    return div(svg+div(''.join(leg),'chart-legend'),'plot')

# Native controls are created once, above every report section.
def brand_html():
    return '<div class="mi-top"><div class="brand"><img src="'+LOGO+'" alt="Jaipuria Institute of Management"><div class="brand-text"><div class="brand-name">Email &amp; WhatsApp Marketing Dashboard</div><div class="brand-sub">PGDM Admissions · Full-cycle comparison · 2026–28 (LY) vs 2027–29 (CY)</div></div><span class="brand-badge">Google Sheets source</span></div></div>'
def nav_html():
    return '<div class="mi-top"><nav class="mi-nav">'+''.join(f'<a href="#section-{i}">{icon(k,GREEN,17)}{esc(n)}</a>' for i,(n,k) in enumerate(zip(PAGES,['grid','chart','mail','wa','chat','users']),1))+'</nav></div>'
def navigation():
    st.markdown(brand_html(),unsafe_allow_html=True)
    a,b=st.columns([7,1]);a.markdown(nav_html(),unsafe_allow_html=True)
    if b.button('↻ Refresh data',key='refresh_all',use_container_width=True):load.clear()
    a,b,c,d,e=st.columns([1,1,1,1,2.1])
    batch=a.selectbox('Batch',B,index=1,key='global_batch')
    kind=b.selectbox('Campaign type',['Automation','Bulk'],key='global_kind')
    start=c.date_input('Transaction start',value=None,key='global_start')
    end=d.date_input('Transaction end',value=None,key='global_end')
    query=e.text_input('Search campaign / subject / sender',key='global_search',placeholder='Search keyword…')
    st.markdown('<div class="filter-note">Batch selects campaign and sender rankings; campaign type selects Email/WhatsApp trends and rankings. Date range filters Transactions only. Search filters detail/rankings only. Overview and batch comparison tables remain full-cycle.</div>',unsafe_allow_html=True)
    return batch,kind,start,end,query

def overview(t,e,w,ch,same):
    counts=t.groupby('Batch').size().reindex(B,fill_value=0);ec=t[t.Source.eq('Emailer')].groupby('Batch').size().reindex(B,fill_value=0);wc=t[t.Source.eq('WhatsApp')].groupby('Batch').size().reindex(B,fill_value=0)
    er=pairrate(e,'Opened','Delivered');wr=pairrate(w[w.Type.eq('Automation')],'Read','Delivered');ar=pairrate(e[e.Type.eq('Automation')],'Opened','Delivered')
    chatrate=ratio(total(ch,'Read'),total(ch,'Delivered')) if same else pairrate(ch,'Read','Delivered')
    left=cards([kpi('Marketing transactions',counts.tolist(),'layers','primary'),kpi('Emailer transactions',ec.tolist(),'mail'),kpi('WhatsApp transactions',wc.tolist(),'wa'),kpi('Email overall open rate',er,'mail',percent=True),kpi('WhatsApp automation read rate',wr,'wa',percent=True),kpi('Converse Chat read rate',chatrate,'chat','accent',True,'Shared summary')])
    right=insight([('Transaction volume',f'{fmt(counts.iloc[0])} LY · {fmt(counts.iloc[1])} CY.'),('Leading channel',f'Emailer accounts for {pct(ratio(ec.iloc[0],counts.iloc[0]))} of LY marketing transactions.'),('Email automation',f'{pct(ar[0])} LY → {pct(ar[1])} CY ({change(*ar)}).'),('WhatsApp automation',f'{pct(wr[0])} LY → {pct(wr[1])} CY ({change(*wr)}).')],foot='Full available cycle exports cover unequal periods; not a matched-period YoY comparison.')
    return section(1,'Executive Overview','grid',left,right,overview=True)

def transaction_section(t,start,end,query):
    invalid=start and end and start>end
    if invalid:start=end=None
    if start:t=t[t.Date>=pd.Timestamp(start)]
    if end:t=t[t.Date<pd.Timestamp(end)+pd.Timedelta(days=1)]
    m=monthly(t).groupby(['Batch','Month']).size().reset_index(name='Transactions'); source=t.groupby(['Batch','Source']).size().reindex(pd.MultiIndex.from_product([B,['Emailer','WhatsApp']],names=['Batch','Source']),fill_value=0).reset_index(name='Transactions')
    charts=split(panel('Monthly transaction trend',svg_chart(m,'Month','Transactions',kind='line',width=390,height=193),'chart',extra='panel-fill'),panel('Channel share · LY',donut(t[t.Batch.eq(B[0])].groupby('Source').size().reset_index(name='Transactions')),'pie',extra='panel-fill'),panel('Source comparison',svg_chart(source,'Source','Transactions',width=300,height=185),'chart',extra='panel-fill'),cls='txcharts')
    summary=count_table(t,['Source','Medium']);summary=pd.concat([summary,pd.DataFrame([{'Source':'Total','Medium':'',**{b:int(summary[b].sum()) for b in B}}])],ignore_index=True)
    ly=t[t.Batch.eq(B[0])];cy=t[t.Batch.eq(B[1])];peak=m.sort_values('Transactions',ascending=False);unknown=int(t.Medium.eq('Not specified').sum());z=t[t.Campaign.str.contains(query,case=False,regex=False)] if query else t
    before=(note('Start date is after end date. Showing the full available transaction period.') if invalid else '')+charts
    left=panel('Transactions by source and medium',table(summary,True),'table')
    right=insight([('Total transactions',f'{len(ly):,} LY · {len(cy):,} CY (selected dates).'),('Leading channel',f'Emailer · {pct(ratio(ly.Source.eq("Emailer").sum(),len(ly)))} of LY transactions.'),('Peak month',f'{monthlabel(peak.iloc[0].Month)} · {peak.iloc[0].Batch} · {fmt(peak.iloc[0].Transactions)} transactions.' if len(peak) else 'No records.'),('Attribution coverage',f'{pct(ratio(len(t)-unknown,len(t)))} specified · {unknown:,} of {len(t):,} missing medium.')],foot='Exact source matching. Blank medium is not inferred.')
    details=detail('Complete campaign-wise transaction breakdown',count_table(z,['Source','Medium','Campaign']))+detail('Month-wise transaction detail',count_table(monthly(t),['Month','Source','Medium']))
    return section(2,'Transaction Performance','chart',left,right,details,before=before)

# Same layout for both channels, with complete comparison tables and dynamic right rail.
def channel_section(d,is_email,batch,kind,query):
    page='Email' if is_email else 'WhatsApp';idx=3 if is_email else 4;n='Opened' if is_email else 'Read';kr='open rate' if is_email else 'read rate';metrics=EM if is_email else WM;ico='mail' if is_email else 'wa'
    agg=aggregate(d,['Batch','Type'],metrics,numer=n);selected=d[d.Batch.eq(batch)&d.Type.eq(kind)];z=selected
    if query:z=z[z.Campaign.str.contains(query,case=False,regex=False)|(z.Subject.str.contains(query,case=False,regex=False) if is_email else False)]
    cols=['Delivered','Opened','Rate','Clicked','Click rate'] if is_email else ['Triggered','Delivered','Read','Rate']
    panels=[]
    for k in ['Automation','Bulk']:
        missing=set(B)-set(d[d.Type.eq(k)].Batch)
        # Keep both table panels aligned. Missing data is already visible as em dashes.
        panels.append(panel(page+' '+k,table(compare_table(agg,k,cols),True)+ (note('Unavailable: '+', '.join(sorted(missing))) if missing else ''),ico))
    before=''
    if not is_email:before+=cards([kpi('Triggered',total(selected,'Triggered'),'send','primary',sub=batch+' · '+kind),kpi('Delivered',total(selected,'Delivered'),'check',sub=batch+' · '+kind),kpi('Read',total(selected,'Read'),'chat',sub=batch+' · '+kind),kpi('Read rate',ratio(total(selected,'Read'),total(selected,'Delivered')),'pie','accent',True,batch+' · '+kind)],'four')
    before+=split(*panels)
    keys=['Campaign','Subject'] if is_email else ['Campaign'];r=ranked(z,keys,metrics,numer=n);short=r.head(5).copy()
    name='Campaign / Subject' if is_email else 'Campaign / Template'
    if is_email:short[name]=list(zip(short.Subject,short.Campaign))
    else:short[name]=[tuple(reversed(v.rsplit(' / ',1))) if ' / ' in v else (v,'') for v in short.Campaign]
    short=short[['Rank',name,'Delivered',n,'Rate']].rename(columns={'Rate':kr.capitalize()})
    m=aggregate(monthly(selected),['Month'],metrics,numer=n);trend=aggregate(monthly(d[d.Type.eq(kind)]),['Batch','Month'],metrics,numer=n)
    volume=svg_chart(m.melt(id_vars='Month',value_vars=['Delivered',n],var_name='Metric',value_name='Volume'),'Month','Volume','Metric',width=570,height=200)
    trendplot=svg_chart(trend,'Month','Rate',kind='line',width=570,height=200)
    left=panel('Top 5 campaigns by '+kr,table(short),'trophy',batch+' · '+kind)
    after=split(panel('Monthly delivered & '+n.lower(),volume,'chart',batch+' · '+kind,extra='panel-fill'),panel('Monthly '+kr+' trend',trendplot,'chart',kind+' · both batches',extra='panel-fill'))
    au=pairrate(d[d.Type.eq('Automation')],n,'Delivered');bu=pairrate(d[d.Type.eq('Bulk')],n,'Delivered');v=ratio(total(selected,n),total(selected,'Delivered'))
    items=[('Automation '+kr,f'{pct(au[0])} LY → {pct(au[1])} CY ({change(*au)}).'),('Bulk '+kr,f'{pct(bu[0])} LY → {pct(bu[1])} CY ({change(*bu)}).'),('Selected campaign type',f'{batch} · {kind}: {fmt(total(selected,"Delivered"))} delivered, {fmt(total(selected,n))} {n.lower()}; {pct(v)}.'),('Top ranked result',f'{pct(r.iloc[0].Rate)} from {fmt(r.iloc[0].Delivered)} deliveries. Low-volume campaigns can have high rates.' if not r.empty else 'No eligible campaign for this selection.')]
    if is_email:items.append(('Click engagement',f'{pct(ratio(total(selected,"Clicked"),total(selected,"Delivered")))} of deliveries clicked.'))
    elif B[1] not in set(d[d.Type.eq('Bulk')].Batch):items.append(('Data availability','Current-batch WhatsApp Bulk export is unavailable, not zero.'))
    right=insight(items,foot='Automation opens/clicks estimated from exported rates × delivered; Bulk uses counts.' if is_email else 'Delivered = Delivered-only + Read + Clicked. Read rate = Read ÷ Delivered.')
    details=detail('Complete campaign rankings · '+batch+' · '+kind,r)+detail('Monthly engagement detail · '+kind,trend)
    return section(idx,page+' Campaign Performance',ico,left,right,details,before=before,after=after)

def converse(ch,same,batch):
    if not same:ch=ch[ch.Batch.eq(batch)]
    label='Shared summary' if same else batch;delivered=total(ch,'Delivered');read=total(ch,'Read');rr=ratio(read,delivered)
    m=aggregate(monthly(ch),['Month'],WM,numer='Read');overall=aggregate(ch.assign(Month='Overall'),['Month'],WM,numer='Read')
    k=cards([kpi('Delivered',delivered,'send','primary',sub=label),kpi('Read',read,'chat',sub=label),kpi('Overall read rate',rr,'pie','accent',True,label)])
    data=pd.concat([m,overall],ignore_index=True)[['Month','Delivered','Read','Rate']].rename(columns={'Rate':'Read rate'})
    plot=svg_chart(m.melt(id_vars='Month',value_vars=['Delivered','Read'],var_name='Metric',value_name='Volume'),'Month','Volume','Metric',width=475,height=158,secondary=dict(zip(m.Month,m.Rate)))
    left=split(panel('Month-wise Converse Chat',table(data,True),'calendar'),panel('Delivered vs read',plot,'chart',extra='panel-fill'),cls='split converse-pair')
    before=k
    after=''
    valid=m.dropna(subset=['Rate']);best=valid.loc[valid.Rate.idxmax()] if len(valid) else None;worst=valid.loc[valid.Rate.idxmin()] if len(valid) else None
    right=insight([('Overall engagement',f'{fmt(read)} reads from {fmt(delivered)} delivered · {pct(rr)}.'),('Best month',f'{monthlabel(best.Month)} · {pct(best.Rate)}.' if best is not None else 'Unavailable.'),('Lowest month',f'{monthlabel(worst.Month)} · {pct(worst.Rate)}.' if worst is not None else 'Unavailable.'),('First to latest month',change(valid.iloc[0].Rate,valid.iloc[-1].Rate) if len(valid)>1 else 'At least two months required.')],'Key Observations',foot='Identical batch summaries shown once; read rate = Read ÷ successfully delivered.' if same else 'Selected batch summary; read rate = Read ÷ successfully delivered.')
    return section(5,'Converse Chat Performance','chat',left,right,before=before,after=after)

def sender_section(s,batch,query):
    summary=aggregate(s,['Batch'],SM,den='Sent');selected=s[s.Batch.eq(batch)];z=selected[selected.Sender.str.contains(query,case=False,regex=False)] if query else selected;ss=total(selected,'Sent');oo=total(selected,'Opened');cc=total(selected,'Clicked')
    k=cards([kpi('Emails sent',ss,'mail','primary',sub=batch),kpi('Opened',oo,'mail',sub=batch),kpi('Clicked',cc,'check',sub=batch),kpi('Open rate',ratio(oo,ss),'pie','accent',True,batch),kpi('Click rate',ratio(cc,ss),'chart',percent=True,sub=batch)],'five')
    r=aggregate(z,['Sender'],SM,den='Sent').sort_values('Sent',ascending=False).reset_index(drop=True);r.insert(0,'Rank',range(1,len(r)+1));short=r.head(5).copy();short['Sender']=[(v.rsplit(' (',1)[0],v.rsplit(' (',1)[1].rstrip(')')) if ' (' in v else (v,'') for v in short.Sender]
    left=panel('Top 5 senders by email volume',table(short[['Rank','Sender','Sent','Opened','Rate','Clicked','Click rate']].rename(columns={'Rate':'Open rate'})),'trophy',batch)
    before=k+panel('Overall email summary · both batches',table(summary.set_index('Batch').reindex(B).reset_index()[['Batch','Sent','Opened','Clicked','Rate','Click rate']].rename(columns={'Rate':'Open rate'}),True),'table')
    a,b=pair(s,'Sent');orates=pairrate(s,'Opened','Sent');crates=pairrate(s,'Clicked','Sent');top=r.iloc[0] if not r.empty else None
    right=insight([('Email volume',f'{fmt(a)} LY · {fmt(b)} CY.'),('Open rate',f'{pct(orates[0])} → {pct(orates[1])} ({change(*orates)}).'),('Click rate',f'{pct(crates[0])} → {pct(crates[1])} ({change(*crates)}).'),('Top matching sender',f'{top.Sender.split(" (")[0]} · {fmt(top.Sent)} emails · {pct(ratio(top.Sent,ss))} of batch volume.' if top is not None else 'No matching sender.'),('Top five contribution',pct(ratio(total(r.head(5),'Sent'),ss))+' of batch volume.')],foot='Rates use sent emails as denominator. No reliable monthly date field in sender exports.')
    details=detail('Complete sender ranking · '+batch,r)+detail('Sender × email type',aggregate(z,['Sender','Type'],SM,den='Sent'))
    return section(6,'Sender-wise Email Performance','users',left,right,details,before=before)

def render_report(data,batch,kind,start,end,query,checked):
    t,e,w,ch,s=[data[k] for k in ['transactions','email','wa','chat','senders']]
    ca=aggregate(monthly(ch),['Batch','Month'],WM,numer='Read');p=ca[ca.Batch.eq(B[0])].drop(columns='Batch').reset_index(drop=True);q=ca[ca.Batch.eq(B[1])].drop(columns='Batch').reset_index(drop=True);same=p.equals(q) and not p.empty;cv=ch[ch.Batch.eq(B[1])] if same else ch
    status='<div class="statusline"><span><b>● Google Sheets connected</b> · '+esc(checked)+'</span><span>Refresh: 60 seconds · '+esc(batch)+' selected · '+esc(kind)+'</span></div>'
    return '<div class="mi-top">'+status+'</div><div class="mi">'+overview(t,e,w,cv,same)+transaction_section(t,start,end,query)+channel_section(e,True,batch,kind,query)+channel_section(w,False,batch,kind,query)+converse(cv,same,batch)+sender_section(s,batch,query)+'<div class="footer">Read-only Google Sheets connection · Full available export periods · — = unavailable, not zero · No Excel runtime connection</div></div>'

@st.fragment(run_every='60s')
def dashboard(batch,kind,start,end,query):
    try:
        with st.spinner('Reading Google Sheets…'):data=load()
    except Exception as exc:
        st.error('Google Sheets connection or schema error. Please contact the dashboard administrator.');return
    checked=datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%d %b %Y, %I:%M:%S %p IST')
    # One report markup tree: browser grid controls all section alignment.
    st.markdown(render_report(data,batch,kind,start,end,query,checked),unsafe_allow_html=True)

filters=navigation()
dashboard(*filters)
