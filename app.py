import html

from datetime import datetime

from zoneinfo import ZoneInfo

import pandas as pd

import streamlit as st

st.set_page_config(page_title='Jaipuria | Marketing Intelligence',layout='wide',initial_sidebar_state='collapsed')



import re

from pathlib import Path

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

    # Some worksheet exports contain Excel numeric date serials (e.g. 46235.59).
    m=out.isna(); serial=pd.to_numeric(s.loc[m],errors='coerce').where(lambda z:z.between(30000,90000))
    out.loc[m]=pd.to_datetime(serial,unit='D',origin='1899-12-30',errors='coerce')

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

    if auto:
        # Google FORMATTED_VALUE may return '9.39%' while unformatted Excel
        # exports may return its numeric fraction 0.0939. Support both.
        oraw=col(d,'Opens (sorted descending)').astype(str)
        craw=col(d,'Clicks').astype(str)
        opened=opened.mask((~oraw.str.contains('%',regex=False)) & opened.between(0,1),opened*100)
        clicks=clicks.mask((~craw.str.contains('%',regex=False)) & clicks.between(0,1),clicks*100)
        opened=opened*delivered/100
        clicks=clicks*delivered/100

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



def voucher_records(rows, batch):
    """Unique paid applications per voucher. All sources qualify, not only Emailer/WhatsApp."""
    d = frame(rows)
    names = ['Voucher Student Code', 'Voucher Student Code Secondary', 'Payment Status',
             'Application Id', 'Transaction Id', 'Transaction Date']
    for name in names:
        col(d, name)
    if d.empty:
        return pd.DataFrame(columns=['Batch', 'Voucher', 'Application ID', 'Payment Status', 'Transaction ID', 'Date'])
    primary = clean(d['Voucher Student Code']).replace({'nan': '', 'None': ''})
    secondary = clean(d['Voucher Student Code Secondary']).replace({'nan': '', 'None': ''})
    code = primary.mask(primary.eq(''), secondary).str.strip()
    status = clean(d['Payment Status'])
    appid = clean(d['Application Id']).replace({'nan': '', 'None': ''})
    tid = clean(d['Transaction Id']).replace({'nan': '', 'None': ''})
    o = pd.DataFrame({'Batch': batch, 'Voucher': code, 'Application ID': appid,
                      'Payment Status': status, 'Transaction ID': tid,
                      'Date': dates(d['Transaction Date'])})
    o = o[(o['Voucher'] != '') & (o['Application ID'] != '') &
          (o['Payment Status'].str.casefold() == 'paid') & o['Date'].notna()].copy()
    o['Voucher'] = o['Voucher'].str.upper()
    # Same paid application may be repeated in exports. Count it once per batch & voucher.
    return o.drop_duplicates(subset=['Batch', 'Voucher', 'Application ID'])

def senders(rows,batch):

    d=frame(rows)

    if d.empty:return pd.DataFrame(columns=['Batch','Sender','Type','Sent','Opened','Clicked','Bounced'])

    out=pd.DataFrame({'Batch':batch,'Sender':clean(col(d,'From User')),'Type':clean(col(d,'Email Type')),'Sent':num(col(d,'# Emails')),'Opened':num(col(d,'# Opened Emails')),'Clicked':num(col(d,'# Clicked Emails')),'Bounced':num(col(d,'# Bounced Emails'))}); return out[out.Sent.notna() & out.Sender.ne('') & ~out.Sender.str.casefold().isin(['total','grand total','summary'])]




def _norm_head(v):
    return re.sub(r'[^a-z0-9]', '', str(v).strip().casefold())


def conversation_ids_from_detail(tabs, batch):
    """Read an actual conversation identifier from an optional Converse Chat detail tab.

    Message IDs and recipient IDs are intentionally NOT counted as conversations.
    Empty result means the source cannot substantiate a unique-chat KPI.
    """
    cols = ['Batch', 'Conversation ID', 'ID Available']
    id_heads = {'conversationid', 'chatid', 'chatthreadid', 'threadid',
                'conversationsessionid', 'conversationuuid', 'chatsessionid'}
    possible = []
    for name, ws in tabs.items():
        title = _norm_head(name)
        if 'converse' in title and 'chat' in title and 'summary' not in title:
            possible.append(ws)
    for ws in possible:
        try:
            ix = None; header_row = 1
            for row in (1,2,3):
                headers = ws.row_values(row)
                ix = next((i for i, value in enumerate(headers)
                           if _norm_head(value) in id_heads), None)
                if ix is not None:
                    header_row = row
                    break
            if ix is None:
                continue
            ids = pd.Series(ws.col_values(ix + 1)[header_row:], dtype='str').str.strip()
            ids = ids[ids.ne('') & ~ids.str.casefold().isin(['nan', 'none', 'null', 'n/a'])]
            return pd.DataFrame({'Batch': batch, 'Conversation ID': ids.drop_duplicates().tolist() or [''], 'ID Available': True})
        except Exception:
            # A supplemental KPI must not interrupt the core reporting dashboard.
            continue
    return pd.DataFrame([{'Batch': batch, 'Conversation ID':'', 'ID Available':False}], columns=cols)


def unique_chat_value(unique_data, scope_batches):
    if unique_data is None or unique_data.empty:
        return float('nan')
    scoped = unique_data[unique_data['Batch'].isin(scope_batches)]
    if scoped.empty:
        return float('nan')
    if not all((scoped.loc[scoped.Batch.eq(b), 'ID Available'] == True).any() for b in scope_batches):
        return float('nan')
    return float(scoped.loc[scoped['Conversation ID'].astype(str).str.strip().ne(''),'Conversation ID'].nunique())

@st.cache_data(ttl=60,show_spinner=False)

def load():

    try:

        gc=client()

    except Exception as exc:

        raise RuntimeError(f'Credential setup: {type(exc).__name__}: {exc!s}' ) from exc

    all_data={k:[] for k in ['transactions','email','wa','chat','senders','voucher','unique_chat']}

    for batch, secret_name in BOOKS.items():
        key = str(st.secrets['marketing_sheets'][secret_name]).strip()

        try:

            b=gc.open_by_key(key); tabs={ws.title:ws for ws in b.worksheets()}

        except (PermissionError, gspread.exceptions.SpreadsheetNotFound) as exc:

            raise RuntimeError(f'Batch {batch}: spreadsheet access denied or unavailable. Share the sheet with the client_email in credentials.json as Viewer. ({type(exc).__name__})') from exc

        except Exception as exc:

            raise RuntimeError(f'Batch {batch}: Google Sheets API connection failed ({type(exc).__name__}: {exc!s})') from exc

        # Only required transaction columns: one header read and one batch request.

        t=tabs['Transaction Details']; headers=t.row_values(1); fields=['Transaction Date','Marketing Portal Source','Marketing Portal Source Medium','Marketing Portal Source Campaign','Voucher Student Code','Voucher Student Code Secondary','Payment Status','Application Id','Transaction Id']; ranges=[]

        for f in fields:

            j=headers.index(f)+1; letter=re.sub(r'\d','',gspread.utils.rowcol_to_a1(1,j)); ranges.append(f'{letter}2:{letter}{t.row_count}')

        values=t.batch_get(ranges); n=max(map(len,values),default=0); rows=[fields]+[[values[j][i][0] if i<len(values[j]) and values[j][i] else '' for j in range(len(fields))] for i in range(n)]

        all_data['transactions'].append(transactions([fields[:4]]+[r[:4] for r in rows[1:]],batch))
        all_data['voucher'].append(voucher_records([fields[4:]+['Transaction Date']]+[r[4:]+[r[0]] for r in rows[1:]],batch))

        names=['Email Automation','Email Bulk','Whatsapp Automation Summary','Whatsapp Bulk Summary','Converse Chat Summary','Total Emails Sent' if batch=='2026–28' else 'Total Email Sent']

        # One request for summary tabs; no student-level detail fetched.

        vals=b.values_batch_get([f"'{name}'!A:{gspread.utils.rowcol_to_a1(1,min(tabs[name].col_count,27))[:-1]}" for name in names],params={'valueRenderOption':'FORMATTED_VALUE'})['valueRanges']

        r=[v.get('values',[]) for v in vals]

        all_data['email'] += [email(r[0],batch,'Automation'),email(r[1],batch,'Bulk')]

        all_data['wa'] += [wa(r[2],batch,'Automation'),wa(r[3],batch,'Bulk')]

        all_data['chat'].append(wa(r[4],batch,'Converse Chat')); all_data['senders'].append(senders(r[5],batch))
        all_data['unique_chat'].append(conversation_ids_from_detail(tabs, batch))

    return {k:pd.concat([x for x in v if not x.empty],ignore_index=True) if any(not x.empty for x in v) else v[0].iloc[0:0].copy() for k,v in all_data.items()}



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

def svg_chart(d,x,y,group='Batch',kind='bar',width=480,height=210,secondary=None,show_value_labels=False):
    """Responsive, accessible SVG chart. All chart coordinates derived from export data."""
    if d.empty or y not in d or d[y].notna().sum()==0:return empty('No chart data for this selection.')
    z=d.copy();z=z[z[x].notna()]; cats=sorted(z[x].unique()) if x in ('Month','CycleMonth') else list(dict.fromkeys(z[x])); groups=list(dict.fromkeys(z[group]));
    if group=='Batch':groups=[b for b in B if b in groups]
    colors={B[0]:GREEN,B[1]:PURPLE,'Delivered':GREEN,'Opened':ORANGE,'Read':ORANGE,'Sent':GREEN,'Emailer':GREEN,'WhatsApp':ORANGE}
    colors.update({k:colors.get(k,[GREEN,PURPLE,ORANGE][i%3]) for i,k in enumerate(groups)})
    legends=''.join(f'<span class="legenditem"><i class="legenddot" style="background:{colors[k]}"></i>{esc(k)}</span>' for k in groups)
    if secondary is not None:legends+=f'<span class="legenditem"><i class="legenddot" style="background:{PURPLE}"></i>Read rate</span>'
    # Fixed viewbox with generous axis margins. No clipped tick labels at narrow widths.
    israte='rate' in y.casefold();left=43;right=46 if secondary is not None else 13;top=14; bottom=49 if x in ('Month','CycleMonth') and len(cats)>6 else 30
    plotw=width-left-right; ploth=height-top-bottom; y0=top+ploth
    maxv=float(z[y].max()); cap=axis_max(maxv*1.12 if kind=='bar' else maxv*1.08)
    if y=='Transactions':
        # Counts must use whole-number axis ticks (not fractional transactions).
        cap=max(4, math.ceil(maxv/4)*4)
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
        lab=(CYCLE_MONTHS[int(c)] if x=='CycleMonth' else monthlabel(c) if x=='Month' else str(c))
        if x in ('Month','CycleMonth') and len(cats)>6:sv.append(text(xx(i),y0+13,lab,'end',9,transform=f'transform="rotate(-38 {xx(i):.2f} {y0+13:.2f})"'))
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
            cy=yy(v);cx=xx(i);displaycat=(CYCLE_MONTHS[int(cats[i])] if x=='CycleMonth' else monthlabel(cats[i]) if x=='Month' else cats[i])
            tooltip=esc(f'{displaycat} · {g}: {pct(v) if israte else fmt(v)}')
            if kind=='bar':
                bx=cx+(j-(len(groups)-1)/2)*barw-barw*.46
                sv.append(f'<rect x="{bx:.2f}" y="{cy:.2f}" width="{barw*.92:.2f}" height="{max(0,y0-cy):.2f}" rx="2" fill="{colors[g]}"><title>{tooltip}</title></rect>')
                if len(cats)<=3 or show_value_labels:
                    sv.append(text(bx+barw*.46,max(10,cy-5),brush(float(v)) if x=='Month' and maxv>10000 else fmt(v),'middle',10,colors[g],'600'))
            else:
                points.append(f'{cx:.2f},{cy:.2f}');sv.append(f'<circle cx="{cx:.2f}" cy="{cy:.2f}" r="3.5" fill="white" stroke="{colors[g]}" stroke-width="2"><title>{tooltip}</title></circle>')
        if kind=='line' and len(points)>1:sv.append(f'<polyline points="{" ".join(points)}" fill="none" stroke="{colors[g]}" stroke-width="2.3" stroke-linejoin="round"/>')
    if secondary is not None:
        positions=[]
        for i,c in enumerate(cats):
            v=secondary.get(c,float('nan'))
            if pd.isna(v):continue
            a=xx(i);b=y0-float(v)/100*ploth;positions.append(f'{a:.2f},{b:.2f}');sv.append(f'<circle cx="{a:.2f}" cy="{b:.2f}" r="3.6" fill="white" stroke="{PURPLE}" stroke-width="2"/>');sv.append(text(a+9,max(top+9,b-12),pct(v),'start',10,PURPLE,'700'))
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
    svg='<svg viewBox="0 0 200 148" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Channel contribution by source">'+''.join(sv)+f'<text x="100" y="71" font-size="25" font-weight="750" text-anchor="middle" fill="{GREEN}">{fmt(tot)}</text><text x="100" y="86" font-size="10" text-anchor="middle" fill="#65766b">Transactions</text></svg>'
    return div(svg+div(''.join(leg),'chart-legend'),'plot')

# Native controls are created once, above every report section.
def brand_html():
    return '<div class="mi-top"><div class="brand"><img src="'+LOGO+'" alt="Jaipuria Institute of Management"><div class="brand-text"><div class="brand-name">Email &amp; WhatsApp Marketing Dashboard</div><div class="brand-sub">PGDM Admissions · Full-cycle comparison · 2026–28 (LY) vs 2027–29 (CY)</div></div><span class="brand-badge">Google Sheets source</span></div></div>'
def nav_html():
    return '<div class="mi-top"><nav class="mi-nav">'+''.join(f'<a href="#section-{i}">{icon(k,GREEN,17)}{esc(n)}</a>' for i,(n,k) in enumerate(zip(PAGES,['grid','chart','mail','wa','chat','users']),1))+'</nav></div>'
# ----------------------------------------------------------------------
# 2026-11-23 correction requirements (10 points): professional report v3.
# All Google keys remain in Streamlit Secrets. Read-only, refreshed every 60s.
# ----------------------------------------------------------------------
PAGES = ['Overview', 'Transactions', 'Email', 'WhatsApp', 'Converse Chat', 'Senders', 'Vouchers']
CSS_V3 = """
/* Layout system: 76% report / 24% management insights, shared top/bottom edges. */
.mi{gap:22px}.section{scroll-margin-top:16px}.section:not(:first-child){padding-top:20px}
.mi .sectitle{font-size:21px!important;padding:10px 13px!important;margin:0 0 12px!important;background:linear-gradient(90deg,#ecf1e7 0%,white 80%);letter-spacing:-.3px}
.section-grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(255px,24%);align-items:stretch;gap:14px}
.leftstack{display:flex;flex-direction:column;gap:12px;min-width:0}.rightstack{display:flex;flex-direction:column;min-width:0}
.rightstack>.insight{flex:1}.rightstack .insight-items{display:flex;flex-direction:column;gap:7px;justify-content:flex-start;flex:0 0 auto}.rightstack .insight-note{margin-top:auto}
.insight-row{border-left:3px solid #6b9c7a!important;background:#f0f6ef!important;padding:10px 11px!important;min-height:56px;display:flex;align-items:center}
.insight-row:nth-child(2n){border-left-color:#9b70b6!important;background:#f7f1fa!important}
.insight-row:nth-child(3n){border-left-color:#eda438!important;background:#fff8ed!important}
.insight-row b{font-size:11px!important}.insight-value{font-size:12px;line-height:1.4;display:block;margin-top:4px}
.panel{border-radius:10px;padding:13px}.panel h3{font-size:14px!important}
.panel .tablebox{max-width:100%;overflow-x:auto}.panel .report th{font-size:10px!important;line-height:1.25}
.report .num{min-width:58px}.report .txt{overflow-wrap:anywhere}.report .namecell{min-width:140px}
.kpis{gap:11px}.kpis.four{grid-template-columns:repeat(4,minmax(0,1fr))}.kpi{min-height:100px}
.kpis.four .kpi{min-height:102px}.kpis.four .kpair .kval{font-size:clamp(17px,1.6vw,23px)}
.overview-bottom{display:grid;grid-template-columns:minmax(170px,.85fr) repeat(2,minmax(0,1.6fr));gap:11px;align-items:stretch}
.overview-bottom .panel{padding:11px}.overview-bottom .kpi{height:100%;min-height:132px}
.tx-top{display:grid;grid-template-columns:minmax(0,1.55fr) minmax(0,.75fr);gap:12px;align-items:stretch}
.tx-bottom{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:12px}
.reports-row{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;align-items:stretch}
.rank-pair{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;align-items:stretch}
.rank-pair .panel{padding:11px;min-width:0}
.rank-pair .tablebox{max-height:350px;overflow:auto;border:1px solid #e0e9e1}
.rank-pair .report th{position:sticky;top:0;z-index:3}
.rank-pair .report td{font-size:11px}.rank-pair .report .secondary{font-size:10px}
.report tr th:first-child,.report tr td:first-child{text-align:left}
.subhead{font-weight:800;color:#234835;font-size:13px;border-left:3px solid #e7a23c;margin:10px 0 8px!important;padding:6px 10px;background:#f3f5ef;border-radius:4px}
.quality-note{font-size:11px;color:#697c72;padding:8px 2px}
/* Single expandable hierarchy table, with aligned numeric columns. */
.hier{border:1px solid #e0e8df;border-radius:9px;overflow:hidden;background:white}
.hier-header,.hier-row{display:grid;grid-template-columns:minmax(0,1fr) 95px 95px;gap:10px;align-items:center;min-width:0}
.hier-header{background:#e7eee1;padding:10px 12px;font-size:11px;font-weight:800;color:#204b39}
.hier-row{padding:9px 12px;border-top:1px solid #e6ebe4;font-size:12px;color:#2c4437}
.hier-row .hier-name{display:flex;align-items:center;gap:7px;min-width:0;overflow-wrap:anywhere}
.hier .hnum{font-variant-numeric:tabular-nums;text-align:right;font-weight:650;white-space:nowrap}
.hier details summary{list-style:none;cursor:pointer}.hier details summary::-webkit-details-marker{display:none}
.hier details summary .hier-name::before{content:'▸';font-size:12px;color:#6f8f72;margin-right:2px}
.hier details[open]>summary .hier-name::before{content:'▾'}
.hier-source>summary{background:#eff5e9;font-weight:750}.hier-medium>summary{background:#fafbf8;font-weight:650}
.hier-medium>summary .hier-name{padding-left:16px}.hier-campaign>.hier-name{padding-left:39px;color:#58705f}
.hier-footer{background:#dfebd3;font-weight:800}
.hier-empty{padding:14px;color:#64796c;font-size:12px}
/* Wide comparisons use a compact bar section, no overstretched table rows. */
.month-table .tablebox{max-height:280px;overflow:auto}.month-table .report th{position:sticky;top:0;z-index:2}
.chart-note{font-size:10.5px;color:#698070;margin:6px 0 0}
.voucher-rank .tablebox{max-height:490px;overflow:auto}.voucher-rank .report th{position:sticky;top:0;z-index:2}
.voucher-rank .report .num{min-width:80px}.voucher-rank .report .txt{min-width:130px}
.section-wide{margin-top:12px;display:grid;gap:12px}.section-wide .rank-pair .tablebox{max-height:370px}
@media(max-width:1250px){.section-grid{grid-template-columns:minmax(0,1fr) minmax(205px,24%)}.rank-pair .report td{padding:6px 5px!important}.overview-bottom{grid-template-columns:1fr 1fr}.overview-bottom>.kpi{grid-column:1/-1}.tx-top{grid-template-columns:minmax(0,1.35fr) minmax(0,.8fr)}}
@media(max-width:1000px){.section-grid{grid-template-columns:1fr}.rightstack>.insight{flex:auto}.rightstack .insight-items{display:grid;grid-template-columns:repeat(2,minmax(0,1fr))}.overview-bottom{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:680px){.kpis.four{grid-template-columns:repeat(2,minmax(0,1fr))}.rank-pair,.reports-row,.tx-top,.tx-bottom,.overview-bottom{grid-template-columns:1fr}.rightstack .insight-items{grid-template-columns:1fr}.hier-header,.hier-row{grid-template-columns:minmax(0,1fr) 65px 65px;gap:5px}.brand-text{border-left:0;padding-left:0}.section-grid{gap:9px}}

/* Final desktop/mobile typography & density QA: never squeeze KPI labels or sender tables. */
.section-grid > .leftstack,.section-grid > .rightstack{min-width:0}
.panel .report th,.panel .report td{font-variant-numeric:tabular-nums}
.rightstack .insight{min-width:0}
.rank-pair .tablebox{scrollbar-width:thin;scrollbar-color:#c3d3c2 transparent}
.rank-pair .report{min-width:445px}
.rank-pair .report th{background:#e8eee2!important}
.voucher-rank .report{min-width:520px}
.hier-row:focus-visible,.hier summary:focus-visible{outline:2px solid #8053ad;outline-offset:-2px}
@media(max-width:1000px){
  .section-grid{grid-template-columns:minmax(0,1fr)}
  .rightstack>.insight{min-height:0}
  .rightstack .insight-items{display:grid;grid-template-columns:repeat(2,minmax(0,1fr))}
  .rightstack .insight-row{min-height:56px}
}
@media(max-width:680px){
  .block-container{padding:10px 13px 22px!important}
  .brand{gap:10px;flex-wrap:wrap}
  .brand img{width:128px}
  .brand-text{flex-basis:calc(100% - 140px);padding-left:0;border-left:0}
  .brand-name{font-size:18px;letter-spacing:-.2px}
  .brand-sub{font-size:10px}
  .brand-badge{display:none}
  .mi .sectitle{font-size:17px!important}
  .kpis,.kpis.four,.kpis.five{grid-template-columns:repeat(2,minmax(0,1fr))}
  .kpi{min-height:91px;padding:10px}
  .kpair .kval{font-size:17px!important}
  .klabel{font-size:11px}
  .overview-bottom,.reports-row,.rank-pair,.tx-top,.tx-bottom,.detail-row{grid-template-columns:1fr}
  .overview-bottom>.kpi{grid-column:auto}
  .rightstack .insight-items{grid-template-columns:1fr}
  .hier-header,.hier-row{grid-template-columns:minmax(0,1fr) 62px 62px;gap:3px;padding-left:7px;padding-right:7px}
  .hier-medium>summary .hier-name{padding-left:9px}
  .hier-campaign>.hier-name{padding-left:22px}
  .panel .report{font-size:11px!important}
  .mi-nav{gap:4px}
  .mi-nav a{font-size:11px;padding:6px 7px;gap:4px}
}
@media(max-width:370px){
  .kpis{grid-template-columns:1fr}
  .overview-bottom{grid-template-columns:1fr}
  .kpis.four,.kpis.five{grid-template-columns:repeat(2,minmax(0,1fr))}
}

"""
st.markdown('<style>'+CSS_V3+'</style>',unsafe_allow_html=True)


def nav_html():
    ico=['grid','chart','mail','wa','chat','users','trophy']
    return '<div class="mi-top"><nav class="mi-nav">'+''.join(
        f'<a href="#section-{i}">{icon(k,GREEN,17)}{esc(n)}</a>'
        for i,(n,k) in enumerate(zip(PAGES,ico),1))+'</nav></div>'


# Header/navigation polish only. All 7 report sections and calculations stay unchanged.
CSS_FILTER_TOOLBAR = """
/* One deliberate desktop row of filters, consistent native widget heights. */
.mi-nav{gap:6px!important;padding:4px 0 2px!important}
.mi-nav a{min-height:36px;display:inline-flex;align-items:center;padding:8px 11px!important;
 font-size:12px!important;white-space:nowrap}
[data-testid="stVerticalBlockBorderWrapper"]{
  background:#fff!important;border:1px solid #dce5d8!important;
  border-radius:10px!important;box-shadow:0 1px 3px rgba(22,53,38,.045)!important;
}
[data-testid="stVerticalBlockBorderWrapper"] > div{padding-top:2px!important;padding-bottom:2px!important}
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stWidgetLabel"] p{
  font-size:11px!important;font-weight:700!important;letter-spacing:0!important;
  white-space:nowrap;color:#35513f!important;
}
[data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="select"] > div,
[data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="input"],
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stDateInput"] input{
  min-height:38px!important;border-radius:7px!important;font-size:12px!important;
}
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stColumn"]{min-width:0!important}
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stColumn"] input{
  min-width:0!important;font-size:12px!important;
}
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stDateInput"] input{
  letter-spacing:-.2px!important;padding-left:8px!important;
}
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stSelectbox"] [data-baseweb="select"] span{
  font-size:12px!important;
}
/* Navigation buttons remain a separate, lightweight utility row. */
[data-testid="stButton"] button[kind="secondary"]{height:36px!important;min-height:36px!important}
@media(max-width:1250px){
 .mi-nav a{padding:7px 8px!important;font-size:11px!important}
 [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stWidgetLabel"] p{font-size:10px!important}
}
/* A single row remains on wide desktops; smaller screens wrap as a clean grid. */
@media(max-width:1180px){
 [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stHorizontalBlock"]{
   flex-wrap:wrap!important;align-items:flex-end!important;gap:10px!important;
 }
 [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]{
   flex:1 1 calc(25% - 10px)!important;width:calc(25% - 10px)!important;
   min-width:185px!important;
 }
}
@media(max-width:800px){
 [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]{
   flex:1 1 calc(50% - 10px)!important;width:calc(50% - 10px)!important;
   min-width:145px!important;
 }
 [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stWidgetLabel"] p{font-size:12px!important}
}
@media(max-width:420px){
 [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]{
   flex:1 1 100%!important;width:100%!important;min-width:0!important;
 }
}
"""
st.markdown('<style>'+CSS_FILTER_TOOLBAR+'</style>', unsafe_allow_html=True)

# Layout polish only: no reporting, credentials, or filter logic changes.
CSS_ALIGNMENT_FINAL = r"""
/* =============================================================
   NOV-2026 ALIGNMENT FIX — PRESENTATION ONLY
   1. Reduce unused Streamlit top space.
   2. Align reporting and insight columns at their top edge.
   3. Prevent right-side cards and tables from stretching to
      artificial heights, especially Sender and Voucher reports.
   4. Consistent chart/table rows, compact gaps, responsive layout.
   ============================================================= */
/* Streamlit layout: remove the default toolbar/header reservation. */
html,body{margin:0!important}
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
section.main{padding-top:0!important;margin-top:0!important}
[data-testid="stMainBlockContainer"],
[data-testid="stAppViewBlockContainer"],
section.main>.block-container,
.block-container{
  padding-top:4px!important;
  padding-bottom:22px!important;
}
/* Top branding/filter density: preserve same native controls. */
.brand{padding:0 0 8px!important;gap:18px!important}
.brand img{max-width:162px!important}
.brand-text{padding-left:18px!important}
.brand-name{line-height:1.15!important}
.mi-nav{padding:0!important;margin:0!important}
.mi-top{margin:0!important;padding:0!important}
[data-testid="stVerticalBlock"]{gap:.46rem!important}
[data-testid="stVerticalBlockBorderWrapper"]>div{padding-top:0!important;padding-bottom:0!important}
.filter-note{margin:0!important}
.statusline{padding:3px 0 6px!important}
/* Compact regular rhythm, without a large empty gap between sections. */
.mi{gap:15px!important;padding-top:3px!important}
.section:not(:first-child){padding-top:13px!important}
.sectitle,.mi .sectitle{margin-bottom:10px!important}
.section-wide{margin-top:10px!important;gap:10px!important}
/* Fixed proportional reporting/insights widths, natural heights. */
.section-grid{
  display:grid!important;
  grid-template-columns:minmax(0,1fr) minmax(285px,24%)!important;
  align-items:start!important;
  gap:12px!important;
}
.section-grid>.leftstack,
.section-grid>.rightstack{
  min-width:0!important;
  min-height:0!important;
  height:auto!important;
  align-self:start!important;
  display:flex!important;
  flex-direction:column!important;
  gap:10px!important;
}
.section-grid>.leftstack>*,.section-grid>.rightstack>*{flex-shrink:0}
.section-grid>.rightstack>.insight{
  flex:0 0 auto!important;
  height:auto!important;
  min-height:0!important;
  align-self:stretch!important;
  justify-content:flex-start!important;
  padding:11px!important;
}
.rightstack .insight-items{
  display:grid!important;
  grid-template-columns:minmax(0,1fr)!important;
  grid-auto-rows:max-content!important;
  align-content:start!important;
  gap:7px!important;
  flex:initial!important;
  min-height:0!important;
}
.rightstack .insight-row{
  display:block!important;
  height:auto!important;
  min-height:0!important;
  padding:9px 10px!important;
  line-height:1.35!important;
}
.rightstack .insight-row b{line-height:1.25!important;margin-bottom:3px!important}
.rightstack .insight-value{line-height:1.4!important;overflow-wrap:anywhere}
.rightstack .insight-note{margin-top:8px!important}
/* Tables/charts: independent content-based heights, rows match naturally. */
.section-grid .aligned-main,
.section-grid .leftstack>.aligned-main{flex:0 0 auto!important;height:auto!important}
.aligned-main>.panel,.aligned-main>.converse-pair{flex:0 0 auto!important;height:auto!important}
.aligned-main>.panel>.tablebox,
.converse-pair>.panel>.tablebox{flex:initial!important;display:block!important;height:auto!important}
.aligned-main>.panel>.tablebox>.report,
.converse-pair>.panel>.tablebox>.report{height:auto!important}
.tx-top,.tx-bottom,.reports-row,.rank-pair,.converse-pair{
  align-items:stretch!important;
  grid-auto-rows:auto!important;
  gap:11px!important;
}
.tx-top>.panel,.tx-bottom>.panel,.reports-row>.panel,
.rank-pair>.panel,.converse-pair>.panel{height:100%;min-width:0}
.panel{min-width:0!important;border-radius:9px!important;padding:12px!important}
.panel h3{margin-bottom:8px!important}
.panel-fill .plot{margin:0!important}
/* No artificially taller SVG charts. */
.chartstack .plot svg,.tx-top .plot svg{max-height:215px!important}
.reports-row .plot svg{max-height:225px!important}
/* Narrow table/card alignment and scroll contained inside the panel. */
.tablebox{max-width:100%!important;min-width:0!important}
.rank-pair .tablebox{max-height:370px!important;overflow:auto!important}
.voucher-rank .tablebox{max-height:480px!important;overflow:auto!important}
.month-table .tablebox{max-height:280px!important;overflow:auto!important}
.report th,.report td{vertical-align:middle!important}
.report .txt{text-align:left!important}
.report .num{text-align:right!important;font-variant-numeric:tabular-nums}
/* Preserve readable widths below desktop breakpoint; insight is below report. */
@media(max-width:1150px){
  .section-grid{grid-template-columns:minmax(0,1fr) minmax(225px,25%)!important;gap:10px!important}
  .brand img{max-width:142px!important}
}
@media(max-width:1000px){
  .section-grid{grid-template-columns:minmax(0,1fr)!important}
  .section-grid>.rightstack{align-self:stretch!important}
  .rightstack .insight-items{grid-template-columns:repeat(2,minmax(0,1fr))!important}
  .section-grid>.rightstack>.insight{height:auto!important}
}
@media(max-width:680px){
  [data-testid="stMainBlockContainer"],.block-container{padding:3px 12px 18px!important}
  .section-grid{gap:9px!important}
  .rightstack .insight-items{grid-template-columns:minmax(0,1fr)!important}
  .tx-top,.tx-bottom,.reports-row,.rank-pair,.converse-pair{grid-template-columns:minmax(0,1fr)!important}
  .panel{padding:10px!important}
  .brand{gap:10px!important}
  .brand-text{padding-left:0!important}
  .mi{gap:13px!important}
}
"""
st.markdown('<style>'+CSS_ALIGNMENT_FINAL+'</style>', unsafe_allow_html=True)

# Final structural alignment: a single grid row governs report + insight heights.
# Reporting calculations, filters, worksheet access, and 60-second refresh remain unchanged.
CSS_STRUCTURAL_ALIGNMENT = r"""
/* 01. Remove Streamlit's top header allowance, without shifting report content into the toolbar. */
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"]{
  display:none!important;height:0!important;min-height:0!important;max-height:0!important;
}
[data-testid="stMain"], [data-testid="stMainBlockContainer"],
[data-testid="stAppViewBlockContainer"], .block-container{
  padding-top:0!important;margin-top:0!important;
}
[data-testid="stAppViewContainer"] > .main,
[data-testid="stAppViewContainer"] > section {padding-top:0!important;margin-top:0!important}
/* Pull branded header up by the residual invisible header spacing in Streamlit.
   Limit to the hosted app, not offline HTML previews. */
.stApp .brand{margin-top:-18px!important}
.brand{padding-bottom:6px!important;gap:15px!important}
.brand img{max-width:156px!important}
.brand-name{font-size:clamp(22px,2vw,30px)!important}
[data-testid="stVerticalBlock"]{gap:.35rem!important}
.statusline{padding:2px 0 4px!important}
.mi{gap:12px!important;padding-top:2px!important}
.section{scroll-margin-top:4px!important}
.section:not(:first-child){padding-top:10px!important}
.sectitle,.mi .sectitle{padding:8px 11px!important;margin-bottom:9px!important;font-size:20px!important}
.section-wide{margin-top:8px!important}
/* 02. Every section: one grid row with shared height and aligned top and bottom. */
.section-grid{
  display:grid!important;
  grid-template-columns:minmax(0,1fr) minmax(285px,24%)!important;
  align-items:stretch!important;
  gap:12px!important;
}
.section-grid>.leftstack,
.section-grid>.rightstack{
  display:flex!important;
  flex-direction:column!important;
  min-width:0!important;
  min-height:0!important;
  height:100%!important;
  align-self:stretch!important;
  gap:10px!important;
}
/* 03. Insight panel fills the *same row*; evenly distribute cards,
   not an empty white box underneath tightly packed cards. */
.section-grid>.rightstack>.insight{
  display:flex!important;
  flex-direction:column!important;
  flex:1 1 auto!important;
  height:100%!important;
  min-height:0!important;
  margin:0!important;
  padding:10px 11px!important;
  align-self:stretch!important;
}
.section-grid .insight h3{flex:0 0 auto!important;margin-bottom:8px!important}
.section-grid .insight-items{
  display:grid!important;
  grid-template-columns:minmax(0,1fr)!important;
  grid-auto-rows:minmax(0,1fr)!important;
  gap:7px!important;
  flex:1 1 auto!important;
  min-height:0!important;
  align-content:stretch!important;
}
.section-grid .insight-row{
  display:flex!important;
  align-items:center!important;
  min-height:0!important;
  height:100%!important;
  padding:7px 9px!important;
  line-height:1.28!important;
}
.section-grid .insight-row b{font-size:10.5px!important;margin-bottom:3px!important}
.section-grid .insight-value{font-size:11.5px!important;line-height:1.31!important}
.section-grid .insight-note{flex:0 0 auto!important;margin-top:7px!important}
/* 04. If insights are taller, expand the final left reporting row; never leave
   a strip of unstyled background underneath the charts and tables. */
.section-grid .leftstack>.overview-bottom:last-child,
.section-grid .leftstack>.reports-row:last-child,
.section-grid .leftstack>.tx-top:last-child,
.section-grid .leftstack>.panel:last-child,
.section-grid .leftstack>.aligned-main:last-child{
  flex:1 1 auto!important;min-height:0!important;
}
.section-grid .leftstack>.reports-row:last-child>.panel,
.section-grid .leftstack>.overview-bottom:last-child>.panel,
.section-grid .leftstack>.overview-bottom:last-child>.kpi,
.section-grid .leftstack>.tx-top:last-child>.panel{
  height:100%!important;min-height:0!important;
}
.section-grid .reports-row>.panel,
.section-grid .tx-top>.panel,
.section-grid .tx-bottom>.panel,
.section-grid .overview-bottom>.panel{
  display:flex!important;flex-direction:column!important;min-width:0!important;
}
/* 05. Plot labels stay visible, plots center vertically in equal-height cards. */
.section-grid .panel .plot{
  display:flex!important;flex-direction:column!important;
  justify-content:center!important;flex:1 1 auto!important;min-width:0!important;
}
.section-grid .panel .plot svg{height:auto!important;max-height:215px!important;width:100%!important}
.section-grid .panel .chart-legend{flex:0 0 auto!important}
/* Overview comparison tables fill their aligned card height; no empty bottom strip. */
#section-1 .overview-bottom>.panel>.tablebox{
  display:flex!important;flex:1 1 auto!important;min-height:0!important;
}
#section-1 .overview-bottom>.panel>.tablebox>.report{
  flex:1 1 auto!important;height:100%!important;
}
/* Converse comparison: the month table uses the same available card height
   as its neighbouring chart; no unused white strip below the table. */
#section-5 .converse-pair>.panel:first-child>.tablebox{
  display:flex!important;flex:1 1 auto!important;align-items:stretch!important;
}
#section-5 .converse-pair>.panel:first-child>.tablebox>.report{
  flex:1 1 auto!important;height:100%!important;
}
/* 06. Tables keep their own scroll, not a tall blank panel. */
.section-grid .panel .tablebox{flex:0 0 auto!important;min-width:0!important}
.section-grid .report th,.section-grid .report td{vertical-align:middle!important}
.section-wide .rank-pair{gap:11px!important;align-items:stretch!important}
.section-wide .rank-pair>.panel{min-width:0!important}
.section-wide .rank-pair .tablebox{max-height:400px!important;overflow:auto!important}
.section-grid .voucher-rank .tablebox{max-height:500px!important;overflow:auto!important}
.section-grid .month-table .tablebox{max-height:290px!important;overflow:auto!important}
/* 07. Compact section details and remove extra dead space. */
.section-grid .panel{padding:11px!important}
.section-grid .kpi{min-height:94px!important}
.section-grid .kpis{gap:10px!important}
.section-grid .overview-bottom{gap:10px!important;align-items:stretch!important}
.section-grid .overview-bottom .panel{height:100%!important}
.detail-row{margin-top:8px!important}
.footer{margin-top:7px!important;padding-top:9px!important}
/* Smaller displays: retain clear one-column stacking and natural insight heights. */
@media (max-width:1120px){
 .section-grid{grid-template-columns:minmax(0,1fr) minmax(225px,25%)!important;gap:9px!important}
 .section-grid .insight-row{padding:6px 7px!important}
 .section-grid .insight-value{font-size:11px!important}
}
@media (max-width:950px){
 .section-grid{grid-template-columns:minmax(0,1fr)!important}
 .section-grid>.leftstack,.section-grid>.rightstack{height:auto!important}
 .section-grid>.rightstack>.insight{height:auto!important}
 .section-grid .insight-items{grid-template-columns:repeat(2,minmax(0,1fr))!important;grid-auto-rows:auto!important}
 .section-grid .insight-row{height:auto!important;min-height:54px!important}
}
@media (max-width:650px){
 .brand-name{font-size:21px!important}
 .section-grid .insight-items{grid-template-columns:minmax(0,1fr)!important}
 .section-grid .insight-row{min-height:0!important}
 .section-grid .reports-row,.section-grid .tx-top,.section-grid .tx-bottom,
 .section-grid .overview-bottom,.section-wide .rank-pair{grid-template-columns:minmax(0,1fr)!important}
 .sectitle,.mi .sectitle{font-size:18px!important}
 .section-grid .panel .plot svg{max-height:210px!important}
}
"""
st.markdown('<style>'+CSS_STRUCTURAL_ALIGNMENT+'</style>', unsafe_allow_html=True)





CSS_DOCX_CORRECTIONS = r"""
/* DOCX correction 1: two Converse Chat KPIs, not one over-stretched count. */
#section-1 .chat-double{min-height:138px!important;padding:12px 14px!important;gap:7px!important}
#section-1 .chat-double .klabel{font-size:11px;min-height:auto!important;margin:0!important}
.chat-pair{display:grid;grid-template-columns:1fr 1fr;gap:8px;align-items:end;margin-top:auto}
.chat-pair>div+div{border-left:1px solid #638773;padding-left:10px}
.chat-value{font-size:clamp(20px,1.7vw,29px);font-weight:760;line-height:1.1;color:white;font-variant-numeric:tabular-nums}
.chat-caption{color:#dbe9de;font-size:10px;line-height:1.3;margin-top:4px}
.chat-method{font-size:9px;line-height:1.25;color:#d3e2d6}
/* DOCX correction 3: two equal-width channel-specific monthly transaction tables. */
.tx-month-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;align-items:stretch}
.tx-month-grid>.panel{min-width:0;display:flex;flex-direction:column;gap:4px}
.tx-month-grid .tablebox{max-height:430px;overflow:auto}
.tx-month-grid .report th{position:sticky;top:0;z-index:2}
.tx-month-grid .report td:first-child,.tx-month-grid .report th:first-child{text-align:left!important}
.tx-month-grid .report tr:last-child td{background:#e2ebd1!important;font-weight:750}
/* DOCX 4-6: each campaign type has a clear section label and two graphs. */
#section-3 .chart-subhead,#section-4 .chart-subhead{font-weight:750;font-size:13px;margin:2px 0 0!important}
#section-3 .monthly-pair,#section-4 .monthly-pair{align-items:stretch;gap:12px}
#section-3 .monthly-pair>.panel,#section-4 .monthly-pair>.panel{min-height:238px!important;display:flex;flex-direction:column}
#section-3 .monthly-pair .plot,#section-4 .monthly-pair .plot{flex:1 1 auto;display:flex;flex-direction:column;justify-content:center}
#section-3 .monthly-pair .plot svg,#section-4 .monthly-pair .plot svg{max-height:206px!important;width:100%!important}
#section-3 .section-grid>.leftstack,#section-4 .section-grid>.leftstack{gap:8px!important}
#section-2 .tx-top .panel-sub{line-height:1.35}
@media(max-width:950px){.tx-month-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(max-width:680px){.tx-month-grid,#section-3 .monthly-pair,#section-4 .monthly-pair{grid-template-columns:minmax(0,1fr)!important}.chat-double .chat-pair{grid-template-columns:1fr 1fr}}
"""
st.markdown('<style>'+CSS_DOCX_CORRECTIONS+'</style>',unsafe_allow_html=True)

def navigation():
    st.markdown(brand_html(), unsafe_allow_html=True)

    # Navigation + refresh form one slim, consistently aligned row.
    nav_col, refresh_col = st.columns([8.6, 1.4], gap='small', vertical_alignment='center')
    nav_col.markdown(nav_html(), unsafe_allow_html=True)
    if refresh_col.button('↻ Refresh', key='refresh_all', use_container_width=True):
        load.clear()

    # Medium choices come from live data; unknown values remain selectable.
    try:
        t0 = load()['transactions']
        media = sorted(v for v in t0.Medium.dropna().unique() if str(v).strip())
    except Exception:
        media = ['Automation', 'Bulk Campaign', 'Not specified']

    # One horizontal filter bar: no disconnected second row or side note.
    with st.container(border=True):
        batch_col, type_col, source_col, medium_col, from_col, to_col, order_col, search_col = st.columns(
            [0.82, 0.94, 1.05, 1.12, 1.02, 1.02, 1.25, 1.75],
            gap='small', vertical_alignment='bottom'
        )
        batch = batch_col.selectbox('Batch', B, index=1, key='global_batch')
        kind = type_col.selectbox('Campaign type', ['Automation', 'Bulk'], key='global_kind')
        source = source_col.selectbox('Source', ['All sources', 'Emailer', 'WhatsApp'], key='tx_source')
        medium = medium_col.selectbox('Medium', ['All mediums'] + media, key='tx_medium')
        start = from_col.date_input('From date', value=None, key='global_start', format='DD/MM/YYYY')
        end = to_col.date_input('To date', value=None, key='global_end', format='DD/MM/YYYY')
        voucher_order = order_col.selectbox(
            'Voucher ranking', ['Highest to lowest', 'Lowest to highest'], key='voucher_order'
        )
        query = search_col.text_input(
            'Search reports', key='global_search', placeholder='Campaign, sender or voucher…'
        )
    return batch, kind, source, medium, start, end, query, voucher_order


def section(i,title,ico,left,right,details='',footer=''):
    result=(f'<section class="section" id="section-{i}"><h2 class="sectitle">'
            +icon(ico,GREEN,20)+esc(title)+'</h2>'
            +'<div class="section-grid"><div class="leftstack">'+left+'</div><div class="rightstack">'+right+'</div></div>')
    if footer:result+=div(footer,'footnote')
    if details:result+=div(details,'section-wide')
    return result+'</section>'


def pair_counts(d,colname):
    return [float(v) for v in d.groupby('Batch')[colname].sum(min_count=1).reindex(B,fill_value=0)] if not d.empty else [0,0]


def breakdown_table(d,channel):
    metric='Opened' if channel=='email' else 'Read'
    metrics=EM if channel=='email' else WM
    if d.empty:return empty('No engagement export records available.')
    a=aggregate(d,['Batch','Type'],metrics,numer=metric)
    if a.empty:return empty('No engagement export records available.')
    import itertools
    idx=pd.MultiIndex.from_product([B,['Automation','Bulk']],names=['Batch','Type'])
    a=a.set_index(['Batch','Type']).reindex(idx).reset_index()
    a=a[['Batch','Type','Rate','Click rate']].rename(
        columns={'Rate':'Open rate' if channel=='email' else 'Read rate'})
    return table(a,True)


def overview(t,e,w,ch,same,unique_chats=None,chat_batches=None):
    counts=pair_counts(t.assign(_n=1),'_n');em=pair_counts(t[t.Source.eq('Emailer')].assign(_n=1),'_n');wa_count=pair_counts(t[t.Source.eq('WhatsApp')].assign(_n=1),'_n')
    er=pairrate(e,'Opened','Delivered'); ecr=pairrate(e,'Clicked','Delivered')
    wr=pairrate(w,'Read','Delivered'); wcr=pairrate(w,'Clicked','Delivered')
    cc=total(ch,'Triggered')
    unique_n=unique_chat_value(unique_chats,chat_batches or B)
    row1=cards([kpi('Combined (Emailer + WhatsApp) Transactions',counts,'layers','primary'),kpi('Emailer Transactions',em,'mail'),kpi('WhatsApp Transactions',wa_count,'wa')])
    row2=cards([kpi('Email Overall Open Rate',er,'mail',percent=True),kpi('Email Overall Click Rate',ecr,'chart',percent=True),kpi('WhatsApp Overall Read Rate',wr,'wa',percent=True),kpi('WhatsApp Overall Click Rate',wcr,'chart','accent',True)],'four')
    # A two-KPI card replaces the old tall single-number panel; it remains aligned
    # with the neighbouring Automation-vs-Bulk summary cards.
    chat=div(
        '<div class="klabel">'+icon('chat','#cde9b0',21)+'Converse Chat · Messages & Conversations</div>'+
        '<div class="chat-pair"><div><div class="chat-value">'+fmt(cc)+'</div><div class="chat-caption">Total triggered messages</div></div>'+
        '<div><div class="chat-value">'+fmt(unique_n)+'</div><div class="chat-caption">Unique chats'+(' · unavailable' if pd.isna(unique_n) else '')+'</div></div></div>'+
        '<div class="chat-method">Distinct conversation IDs, when the detail export provides them</div>',
        'kpi primary chat-double')
    bottom=div(chat+panel('Email · Automation vs Bulk',breakdown_table(e,'email'),'mail')+panel('WhatsApp · Automation vs Bulk',breakdown_table(w,'wa'),'wa'),'overview-bottom')
    left=row1+row2+bottom
    anu=pairrate(e[e.Type.eq('Automation')],'Opened','Delivered');wau=pairrate(w[w.Type.eq('Automation')],'Read','Delivered')
    items=[('Total marketing transactions',f'{fmt(counts[0])} LY · {fmt(counts[1])} CY.'),
           ('Email engagement',f'Open {pct(er[0])} LY → {pct(er[1])} CY; click {pct(ecr[0])} → {pct(ecr[1])}.'),
           ('WhatsApp engagement',f'Read {pct(wr[0])} LY → {pct(wr[1])} CY; click {pct(wcr[0])} → {pct(wcr[1])}.'),
           ('Automation comparison',f'Email opens {pct(anu[0])} → {pct(anu[1])}; WhatsApp reads {pct(wau[0])} → {pct(wau[1])}.'),
           ('Converse Chat',f'{fmt(cc)} triggered messages · {fmt(unique_n)} unique chats.' if pd.notna(unique_n) else f'{fmt(cc)} triggered; unique count requires conversation IDs in a Converse Chat detail tab.'),
           ('Comparison scope','Full export periods differ across batches; not matched-period YoY.')]
    return section(1,'Executive Overview','grid',left,insight(items))

def nested_transaction_table(t):
    """A single Source → Medium → Campaign drill-down table; counts by row, not invented mediums."""
    if t.empty:return empty('No transactions match these filters.')
    total_by_batch=t.groupby('Batch').size().reindex(B,fill_value=0)
    html_=['<div class="hier" role="region" aria-label="Expandable transaction breakdown">',
           '<div class="hier-header"><span>Source / Medium / Campaign</span><span class="hnum">2026–28</span><span class="hnum">2027–29</span></div>']
    def nums(df):
        z=df.groupby('Batch').size().reindex(B,fill_value=0)
        return '<span class="hnum">'+fmt(z[B[0]])+'</span><span class="hnum">'+fmt(z[B[1]])+'</span>'
    for src,sd in t.groupby('Source',sort=True):
        html_.append('<details class="hier-source"><summary class="hier-row"><span class="hier-name">'+esc(src)+'</span>'+nums(sd)+'</summary>')
        for med,md in sd.groupby('Medium',sort=True):
            html_.append('<details class="hier-medium"><summary class="hier-row"><span class="hier-name">'+esc(med)+'</span>'+nums(md)+'</summary>')
            for camp,cd in md.groupby('Campaign',sort=True):
                html_.append('<div class="hier-row hier-campaign"><span class="hier-name">'+esc(camp)+'</span>'+nums(cd)+'</div>')
            html_.append('</details>')
        html_.append('</details>')
    html_.append('<div class="hier-row hier-footer"><span>Total</span><span class="hnum">'+fmt(total_by_batch[B[0]])+'</span><span class="hnum">'+fmt(total_by_batch[B[1]])+'</span></div></div>')
    return ''.join(html_)



CYCLE_MONTHS = ['Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug']

def cycle_month_index(date_series):
    dt=pd.to_datetime(date_series,errors='coerce')
    return (dt.dt.month-9)%12


def cycle_rate_data(d, metric):
    """Weighted rates for matching months, rather than disjoint year-month axes."""
    if d.empty:
        return pd.DataFrame(columns=['Batch','CycleMonth','Rate'])
    a=d.copy();a['CycleMonth']=cycle_month_index(a['Date'])
    a=a.dropna(subset=['CycleMonth'])
    if a.empty:return pd.DataFrame(columns=['Batch','CycleMonth','Rate'])
    a['CycleMonth']=a['CycleMonth'].astype(int)
    g=a.groupby(['Batch','CycleMonth'])[['Delivered',metric]].sum(min_count=1).reset_index()
    g['Rate']=rate(g[metric],g['Delivered'])
    return g[['Batch','CycleMonth','Rate']]


def monthly_channel_comparison(t, source):
    """Side-by-side LY and CY transaction counts for Sep–Aug reporting months."""
    a=t[t.Source.eq(source)].copy()
    if a.empty:return pd.DataFrame(columns=['Month']+B+['Total'])
    a['CycleMonth']=cycle_month_index(a.Date)
    a=a.dropna(subset=['CycleMonth'])
    a['CycleMonth']=a.CycleMonth.astype(int)
    g=a.groupby(['CycleMonth','Batch']).size().unstack(fill_value=0).reindex(columns=B,fill_value=0)
    # Show months present in either cycle only; don't invent future observations.
    g=g.sort_index().reset_index()
    g['Month']=g.CycleMonth.map(dict(enumerate(CYCLE_MONTHS)))
    g['Total']=g[B].sum(axis=1)
    out=g[['Month']+B+['Total']]
    tot=pd.DataFrame([{'Month':'Total',B[0]:int(out[B[0]].sum()),B[1]:int(out[B[1]].sum()),'Total':int(out.Total.sum())}])
    return pd.concat([out,tot],ignore_index=True)

def transaction_section(t, batch, source, medium, start, end, query):
    invalid=bool(start and end and start>end)
    if invalid:start=end=None
    tt=t.copy()
    if start:tt=tt[tt.Date>=pd.Timestamp(start)]
    if end:tt=tt[tt.Date<pd.Timestamp(end)+pd.Timedelta(days=1)]
    if source!='All sources':tt=tt[tt.Source.eq(source)]
    if medium!='All mediums':tt=tt[tt.Medium.eq(medium)]
    selected=tt[tt.Batch.eq(batch)]
    monthly_source=monthly(selected).groupby(['Month','Source']).size().reset_index(name='Transactions')
    if len(monthly_source):
        ix=pd.MultiIndex.from_product([sorted(monthly_source.Month.unique()),['Emailer','WhatsApp']],names=['Month','Source'])
        monthly_source=monthly_source.set_index(['Month','Source']).reindex(ix,fill_value=0).reset_index()
    ly=tt[tt.Batch.eq(B[0])];cy=tt[tt.Batch.eq(B[1])]
    pie=selected.groupby('Source').size().reset_index(name='Transactions')
    # Direct value labels + count-specific heading make very small transaction numbers legible.
    chart1=panel('Monthly Transactions · Emailer vs WhatsApp',
                 svg_chart(monthly_source,'Month','Transactions',group='Source',kind='bar',width=640,height=218,show_value_labels=True),
                 'chart',sub=f'{batch} · transaction count by month · Green = Emailer / Orange = WhatsApp',extra='panel-fill')
    chart2=panel('Channel Contribution',donut(pie),'pie',sub=f'{batch} · selected filter scope',extra='panel-fill')
    charts=div(chart1+chart2,'tx-top')
    email_month=monthly_channel_comparison(tt,'Emailer')
    wa_month=monthly_channel_comparison(tt,'WhatsApp')
    # Exactly two consistent monthwise tables as requested in the redlined document.
    detail_tables=div(
        panel('Emailer · Month-wise Transactions',table(email_month),'mail',sub='Sep–Aug aligned months · LY vs CY')+
        panel('WhatsApp · Month-wise Transactions',table(wa_month),'wa',sub='Sep–Aug aligned months · LY vs CY'),
        'tx-month-grid')
    search_t=tt[tt.Campaign.str.contains(query,case=False,regex=False)] if query else tt
    left=charts+panel('Transaction Attribution · Source → Medium → Campaign',nested_transaction_table(search_t),'table',sub='Expand source and medium to see actual campaign counts')
    lead=tt.Source.value_counts();leading=lead.index[0] if len(lead) else 'Unavailable'
    monthly_count=monthly(tt).groupby(['Month','Source']).size().reset_index(name='Transactions')
    peak=monthly_count.sort_values('Transactions',ascending=False) if len(monthly_count) else monthly_count
    unknown=int(tt.Medium.eq('Not specified').sum());coverage=ratio(len(tt)-unknown,len(tt))
    right=insight([('Total transactions',f'{fmt(len(ly))} LY · {fmt(len(cy))} CY.'),
                   ('Leading channel',f'{leading} · {fmt(lead.iloc[0])} transactions.' if len(lead) else 'No data.'),
                   ('Peak month',f'{monthlabel(peak.iloc[0].Month)} · {peak.iloc[0].Source} · {fmt(peak.iloc[0].Transactions)}.' if len(peak) else 'No data.'),
                   ('Medium completeness',f'{pct(coverage)} specified · {fmt(unknown)} not specified.'),
                   ('Monthwise comparison','Emailer and WhatsApp have independent LY/CY month tables below.')],
                  foot='Transactions use Transaction Date. Filters apply to charts and tables; Search applies to the hierarchy only.')
    return section(2,'Transaction Performance','chart',left,right,
                   details=detail_tables,
                   footer=note('Start date exceeds end date; showing all dates.') if invalid else '')

def ranked_all(d, metric):
    """All campaigns, weighted sums by campaign/subject; independent open/read vs click rankings."""
    if d.empty:
        return pd.DataFrame(columns=['Rank','Campaign / Subject','Delivered',metric,'Rate'])
    group=['Campaign','Subject'] if 'Subject' in d else ['Campaign']
    metric_fields=['Delivered',metric] if metric=='Clicked' else ['Delivered',metric,'Clicked']
    a=d.groupby(group,dropna=False)[metric_fields].sum(min_count=1).reset_index()
    a=a[a.Delivered.gt(0)].copy()
    a['Rate']=rate(a[metric],a.Delivered)
    a=a.dropna(subset=['Rate']).sort_values(['Rate','Delivered','Campaign'],ascending=[False,False,True],kind='stable').reset_index(drop=True)
    a.insert(0,'Rank',range(1,len(a)+1))
    if 'Subject' in a:
        a['Campaign / Subject']=[(sub,camp) for camp,sub in zip(a.Campaign,a.Subject)]
    else:
        a['Campaign / Subject']=[(val.rsplit(' / ',1)[-1],val.rsplit(' / ',1)[0]) if ' / ' in val else (val,'') for val in a.Campaign]
    return a[['Rank','Campaign / Subject','Delivered',metric,'Rate']].rename(columns={'Rate':('Click rate' if metric=='Clicked' else 'Open rate' if metric=='Opened' else 'Read rate')})


def channel_section(d, is_email, batch, kind, query):
    channel='Email' if is_email else 'WhatsApp';index=3 if is_email else 4
    metric='Opened' if is_email else 'Read';ico='mail' if is_email else 'wa';m=EM if is_email else WM
    summary=aggregate(d,['Batch','Type'],m,numer=metric)
    def compact_type(typ):
        z=summary[summary.Type.eq(typ)].set_index('Batch').reindex(B).rename_axis('Batch').reset_index()
        fields=['Batch','Delivered',metric,'Rate','Clicked','Click rate']
        if not is_email:fields.insert(1,'Triggered')
        return table(z[fields].rename(columns={'Rate':'Open rate' if is_email else 'Read rate'}),True)
    header=div(panel(channel+' Automation',compact_type('Automation'),ico)+panel(channel+' Bulk',compact_type('Bulk'),ico),'reports-row')
    selected=d[d.Batch.eq(batch)&d.Type.eq(kind)]
    if query:
        mask=selected.Campaign.str.contains(query,case=False,regex=False)
        if is_email:mask=mask|selected.Subject.str.contains(query,case=False,regex=False)
        selected=selected[mask]
    rank_html=[];ranks={}
    for typ in ['Automation','Bulk']:
        current=d[d.Batch.eq(batch)&d.Type.eq(typ)]
        if query:
            mask=current.Campaign.str.contains(query,case=False,regex=False)
            if is_email:mask=mask|current.Subject.str.contains(query,case=False,regex=False)
            current=current[mask]
        open_rank=ranked_all(current,metric);click_rank=ranked_all(current,'Clicked')
        ranks[typ]=(open_rank,click_rank)
        head=div(esc(typ+' · Full Campaign Rankings  |  '+batch+' · '+fmt(len(open_rank))+' campaigns'),'subhead')
        rank_html.append(head+div(
            panel('All campaigns · '+('Open Rate' if is_email else 'Read Rate'),table(open_rank),'trophy',sub='Complete ranking by delivered-based rate; check volume alongside rate')+
            panel('All campaigns · Click Rate',table(click_rank),'trophy',sub='Complete ranking by delivered-based click rate'), 'rank-pair'))

    # Each channel gets TWO pairs: Automation volume+rate and Bulk volume+rate.
    # The LY and CY lines are plotted on the SAME Sep–Aug month-of-cycle axis.
    chart_rows=[]
    for typ in ['Automation','Bulk']:
        typ_data=d[d.Type.eq(typ)]
        volume=aggregate(monthly(typ_data[typ_data.Batch.eq(batch)]),['Month'],m,numer=metric)
        if not volume.empty:
            melted=volume.melt(id_vars='Month',value_vars=['Delivered',metric],var_name='Metric',value_name='Volume')
            vplot=svg_chart(melted,'Month','Volume',group='Metric',kind='bar',width=520,height=210,show_value_labels=True)
        else:
            vplot=empty(f'{batch} {channel} {typ}: data unavailable for this export.')
        trend=cycle_rate_data(typ_data,metric)
        if trend.empty:
            tr=empty(f'No {typ} month-on-month rates are available for either batch.')
        else:
            tr=svg_chart(trend,'CycleMonth','Rate',group='Batch',kind='line',width=520,height=210)
        tag=div(typ+' · Monthly Engagement  |  '+batch,'subhead chart-subhead')
        row=div(
            panel('Monthly Delivered & '+metric,vplot,'chart',sub=f'{batch} · {typ} · delivered vs '+metric.lower(),extra='panel-fill')+
            panel('Monthly '+('Open Rate' if is_email else 'Read Rate')+' · LY vs CY',tr,'chart',sub='Sep–Aug · same calendar months across both batches',extra='panel-fill'),
            'reports-row monthly-pair')
        chart_rows.append(tag+row)
    left=header+''.join(chart_rows)
    overall=pairrate(d,metric,'Delivered');clicks=pairrate(d,'Clicked','Delivered')
    a=pairrate(d[d.Type.eq('Automation')],metric,'Delivered');b=pairrate(d[d.Type.eq('Bulk')],metric,'Delivered')
    toplist=ranks[kind][0];top=toplist.iloc[0] if not toplist.empty else None
    items=[('Overall '+('open' if is_email else 'read')+' rate',f'{pct(overall[0])} LY → {pct(overall[1])} CY ({change(*overall)}).'),
           ('Overall click rate',f'{pct(clicks[0])} LY → {pct(clicks[1])} CY ({change(*clicks)}).'),
           ('Automation '+('open' if is_email else 'read')+' rate',f'{pct(a[0])} LY → {pct(a[1])} CY.'),
           ('Bulk '+('open' if is_email else 'read')+' rate',f'{pct(b[0])} LY → {pct(b[1])} CY.'),
           ('Selected batch/type',f'{batch} · {kind}: {fmt(total(selected,"Delivered"))} delivered, {fmt(total(selected,metric))} '+metric.lower()+'; '+pct(ratio(total(selected,metric),total(selected,'Delivered')))+'.'),
           ('Highest '+('open' if is_email else 'read')+' rate',f'{top["Campaign / Subject"][0]} · {pct(top["Open rate" if is_email else "Read rate"])} on {fmt(top.Delivered)} deliveries.' if top is not None else 'Unavailable.'),
           ('Monthly comparison','Automation and Bulk each show delivered volume plus month-aligned LY/CY rates.')]
    if not is_email and not len(d[d.Batch.eq(B[1])&d.Type.eq('Bulk')]):
        items.append(('Bulk data coverage','CY WhatsApp Bulk export unavailable; reported as —, not 0.'))
    foot='Email Automation opens/clicks are estimated as delivered × exported %; Bulk uses counts.' if is_email else 'Read and Click rates use Delivered denominator; missing exports are unavailable.'
    return section(index,channel+' Campaign Performance',ico,left,insight(items,foot=foot),details=''.join(rank_html))

def converse(ch,same,batch,unique_chats=None,chat_batches=None):
    x=ch if same else ch[ch.Batch.eq(batch)]
    delivered=total(x,'Delivered');read=total(x,'Read');triggered=total(x,'Triggered')
    unique_n=unique_chat_value(unique_chats,chat_batches or B)
    m=aggregate(monthly(x),['Month'],WM,numer='Read')
    overall=aggregate(x.assign(Month='Overall'),['Month'],WM,numer='Read')
    k=cards([kpi('Triggered chat messages',triggered,'send','primary',sub='Messages, not unique chats'),
             kpi('Unique chats',unique_n,'users',sub='Distinct conversation IDs if available'),
             kpi('Delivered',delivered,'check',sub='Available export'),
             kpi('Read',read,'chat',sub='Available export'),
             kpi('Read rate',ratio(read,delivered),'pie','accent',True,sub='Read ÷ Delivered')],'five')
    data=pd.concat([m,overall],ignore_index=True)[['Month','Triggered','Delivered','Read','Rate']].rename(columns={'Rate':'Read rate'}) if len(m) else pd.DataFrame(columns=['Month','Triggered','Delivered','Read','Read rate'])
    melted=m.melt(id_vars='Month',value_vars=['Delivered','Read'],var_name='Metric',value_name='Volume') if len(m) else pd.DataFrame(columns=['Month','Metric','Volume'])
    curve=svg_chart(melted,'Month','Volume','Metric',kind='bar',width=540,height=210,secondary=dict(zip(m.Month,m.Rate)) if len(m) else None)
    main=div(panel('Month-wise Converse Chat',table(data),'calendar')+panel('Delivered vs Read · Monthly',curve,'chart'),'reports-row')
    valid=m.dropna(subset=['Rate'])
    best=valid.loc[valid.Rate.idxmax()] if len(valid) else None
    worst=valid.loc[valid.Rate.idxmin()] if len(valid) else None
    right=insight([('Triggered volume',f'{fmt(triggered)} messages triggered in this report.'),
                   ('Read performance',f'{fmt(read)} read / {fmt(delivered)} delivered · {pct(ratio(read,delivered))}.'),
                   ('Best month',f'{monthlabel(best.Month)} · {pct(best.Rate)}.' if best is not None else 'Unavailable.'),
                   ('Lowest month',f'{monthlabel(worst.Month)} · {pct(worst.Rate)}.' if worst is not None else 'Unavailable.'),
                   ('Unique chats',f'{fmt(unique_n)} distinct conversation IDs.' if pd.notna(unique_n) else 'Not available: no supported chat/conversation ID field in the detail export.'),
                   ('Metric definition','Triggered messages are not necessarily unique chat conversations.')],
                  foot='Same monthly summary across batches is counted only once.')
    return section(5,'Converse Chat Performance','chat',k+main,right)


def sender_section(s,batch,query):
    summary=aggregate(s,['Batch'],SM,den='Sent')
    selected=s[s.Batch.eq(batch)]
    z=selected[selected.Sender.str.contains(query,case=False,regex=False)] if query else selected
    ss=total(selected,'Sent');oo=total(selected,'Opened');cc=total(selected,'Clicked')
    k=cards([kpi('Emails sent',ss,'mail','primary',sub=batch),kpi('Opened',oo,'mail',sub=batch),kpi('Clicked',cc,'check',sub=batch),kpi('Open rate',ratio(oo,ss),'pie','accent',True,batch),kpi('Click rate',ratio(cc,ss),'chart',percent=True,sub=batch)],'five')
    r=aggregate(z,['Sender'],SM,den='Sent').sort_values('Sent',ascending=False).reset_index(drop=True)
    r.insert(0,'Rank',range(1,len(r)+1));short=r.head(5).copy()
    if len(short):short['Sender']=[(v.rsplit(' (',1)[0],v.rsplit(' (',1)[1].rstrip(')')) if ' (' in v else (v,'') for v in short.Sender]
    left=k+panel('Top Senders by Email Volume',table(short[['Rank','Sender','Sent','Opened','Rate','Clicked','Click rate']].rename(columns={'Rate':'Open rate'})),'users',sub=batch)+panel('Overall Email Summary · Both Batches',table(summary.set_index('Batch').reindex(B).reset_index()[['Batch','Sent','Opened','Clicked','Rate','Click rate']].rename(columns={'Rate':'Open rate'}),True),'table')
    a,b=pair(s,'Sent');orr=pairrate(s,'Opened','Sent');crr=pairrate(s,'Clicked','Sent')
    top=r.iloc[0] if len(r) else None
    right=insight([('Volume',f'{fmt(a)} LY · {fmt(b)} CY.'),('Open rate',f'{pct(orr[0])} → {pct(orr[1])} ({change(*orr)}).'),('Click rate',f'{pct(crr[0])} → {pct(crr[1])} ({change(*crr)}).'),('Top matching sender',f'{top.Sender.split(" (")[0]} · {fmt(top.Sent)} sent.' if top is not None else 'No matching sender.'),('Sender scope','Rates use Emails Sent as denominator; no monthly date column in export.')])
    details=detail('Complete Sender Rankings · '+batch,r)+detail('Sender × Email Type',aggregate(z,['Sender','Type'],SM,den='Sent'))
    return section(6,'Sender-wise Email Performance','users',left,right,div(details,'detail-row'))


def voucher_section(v,order,query):
    if query:v=v[v.Voucher.str.contains(query,case=False,regex=False)]
    if v.empty:
        data=pd.DataFrame(columns=['Student Voucher Code']+B+['Total Paid Applications'])
    else:
        data=v.groupby(['Voucher','Batch'])['Application ID'].nunique().unstack(fill_value=0).reindex(columns=B,fill_value=0)
        data['Total Paid Applications']=data.sum(axis=1)
        data=data.reset_index().rename(columns={'Voucher':'Student Voucher Code'})
        data=data.sort_values(['Total Paid Applications','Student Voucher Code'],ascending=[order=='Lowest to highest',True],kind='stable').reset_index(drop=True)
        data.insert(0,'Rank',range(1,len(data)+1))
    unique_ly=v[v.Batch.eq(B[0])]['Application ID'].nunique() if len(v) else 0
    unique_cy=v[v.Batch.eq(B[1])]['Application ID'].nunique() if len(v) else 0
    best=data.sort_values('Total Paid Applications',ascending=False).iloc[0] if len(data) else None
    suspicious=int(v.Voucher.astype(str).str.fullmatch(r'\d{5}(?:\.0)?').sum()) if len(v) else 0
    cards1=cards([kpi('Paid Voucher Applications',[unique_ly,unique_cy],'layers','primary'),
                  kpi('Voucher Codes Used',v.Voucher.nunique() if len(v) else 0,'trophy',sub='Across available batches'),
                  kpi('Top Voucher · Paid Applications',best['Total Paid Applications'] if best is not None else float('nan'),'check',sub=best['Student Voucher Code'] if best is not None else 'Unavailable')])
    left=cards1+panel('Voucher-wise Paid Application Rankings',table(data),'table',sub='Unique Application ID · Payment Status = Paid · sortable highest/lowest via top filter',extra='voucher-rank')
    right=insight([('Paid applications with vouchers',f'{fmt(unique_ly)} LY · {fmt(unique_cy)} CY.'),
                   ('Distinct codes',f'{fmt(v.Voucher.nunique())} codes in the filtered scope.'),
                   ('Highest-converting code',f'{best["Student Voucher Code"]} · {fmt(best["Total Paid Applications"])} paid applications.' if best is not None else 'No records.'),
                   ('Sort direction',order),
                   ('Counting rule','Unique Application ID per Batch + Voucher with Payment Status Paid and valid Transaction Date.'),
                   ('Voucher data quality',f'{fmt(suspicious)} paid records have numeric date-like voucher values; check source cell formatting.' if suspicious else 'No numeric date-like voucher codes detected.')],
                  foot='Voucher code primary field, fallback to secondary. All sources included, no duplicate paid applications.')
    return section(7,'Student Voucher Code Performance','trophy',left,right)


def render_report(data,batch,kind,source,medium,start,end,query,voucher_order,checked):
    t,e,w,ch,s,v=[data[k] for k in ['transactions','email','wa','chat','senders','voucher']]
    unique_chats=data.get('unique_chat',pd.DataFrame(columns=['Batch','Conversation ID','ID Available']))
    ca=aggregate(monthly(ch),['Batch','Month'],WM,numer='Read')
    p=ca[ca.Batch.eq(B[0])].drop(columns='Batch').reset_index(drop=True)
    q=ca[ca.Batch.eq(B[1])].drop(columns='Batch').reset_index(drop=True)
    same=p.equals(q) and not p.empty
    cv=ch[ch.Batch.eq(B[1])] if same else ch[ch.Batch.eq(batch)]
    status='<div class="statusline"><span><b>● Google Sheets connected</b> · '+esc(checked)+'</span><span>Auto-refresh 60s · '+esc(batch)+' · '+esc(kind)+'</span></div>'
    chat_batches=B if same else [batch]
    parts=[overview(t,e,w,cv,same,unique_chats,chat_batches),transaction_section(t,batch,source,medium,start,end,query),
           channel_section(e,True,batch,kind,query),channel_section(w,False,batch,kind,query),
           converse(cv,same,batch,unique_chats,chat_batches),sender_section(s,batch,query),voucher_section(v,voucher_order,query)]
    return '<div class="mi-top">'+status+'</div><div class="mi">'+''.join(parts)+ '<div class="footer">Source: Google Sheets only · Automatic refresh every 60s while open · Full-cycle LY/CY comparisons are not matched reporting periods · — = unavailable, not zero</div></div>'

@st.fragment(run_every='60s')
def dashboard(batch,kind,source,medium,start,end,query,voucher_order):
    try:
        with st.spinner('Reading Google Sheets…'):
            data=load()
    except Exception as exc:
        st.error('Google Sheets connection or schema error. Review worksheet names and permissions.');return
    checked=datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%d %b %Y, %I:%M:%S %p IST')
    st.markdown(render_report(data,batch,kind,source,medium,start,end,query,voucher_order,checked),unsafe_allow_html=True)

# Mount the navigation and dashboard only once per script execution.
# The module-level marker resets on normal Streamlit reruns, preserving auto refresh.
if not globals().get('_MARKETING_UI_MOUNTED', False):
    _MARKETING_UI_MOUNTED = True
    filters = navigation()
    dashboard(*filters)
