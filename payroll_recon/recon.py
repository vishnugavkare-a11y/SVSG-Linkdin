"""Payroll vs bank-file reconciliation.
Usage: python3 recon.py <summary.xlsx> <bank1.xlsx> [<bank2.xlsx> ...] -o out.xlsx
Month->tab mapping: edit TABS below, or pass --tab "Jul=other.xlsx::July 26" to take a month from another workbook."""
import sys, re, argparse, difflib, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

MONTHS=['Apr','May','Jun','Jul','Aug','Sep']
TABS={'Apr':'April 26_Vish','May':'May 26_Vish','Jun':'June 26_Vish','Jul':'July 26','Aug':'August 26','Sep':'September 26'}
# bank-file name -> summary name (normalised keys); add new aliases here
ALIAS={'priya patel bhumi':'bhumi mahesh patel','devesh pravn gurjar':'devesh pravin gurjar'}
TOL=10  # rupees tolerance
def norm(s): return re.sub(r'\s+',' ',str(s).replace('\t',' ')).strip().lower()
def match(name,known):
    n=norm(name); n=ALIAS.get(n,n)
    if n in known: return n
    t=n.split()
    for k in known:                       # first+last token match
        kt=k.split()
        if t and kt and t[0]==kt[0] and t[-1]==kt[-1]: return k
    m=difflib.get_close_matches(n,known,1,0.75)
    return m[0] if m else None
def month_of(narr,fallback):
    s=str(narr).lower()
    for m,full in zip(MONTHS,['april','may','june','july','august','sept']):
        if full in s or m.lower()+'2' in s.replace(' ',''): return m
    return fallback
def load_summary(path):
    rows=[]
    cache={path:pd.read_excel(path,sheet_name=None)}
    for m,tab in TABS.items():
        f=path
        if '::' in tab: f,tab=tab.split('::',1)       # month taken from another workbook
        if f not in cache: cache[f]=pd.read_excel(f,sheet_name=None)
        x=cache[f]
        if tab not in x: print('WARN missing tab',tab); continue
        d=x[tab]; d=d[d['Name'].notna()].copy()
        for _,r in d.iterrows():
            g=lambda c:(0 if c not in d or pd.isna(r[c]) else r[c])
            rows.append(dict(month=m,key=norm(r['Name']),name=str(r['Name']).replace('\t','').strip(),
              salary=g('Salary'),days=g('Total Days Payable'),payable=g('Salary Payable'),pt=g('PT'),loan=g('Loan'),
              net=g('Salary for the Month'),svsg=(r['SVSG'] if 'SVSG' in d and pd.notna(r['SVSG']) else None),
              fs=(r['FS'] if 'FS' in d and pd.notna(r['FS']) else None),entity=str(r.get(' Entity ','')) ))
    return pd.DataFrame(rows)
def load_bank(paths):
    rows=[]
    for p in paths:
        d=pd.read_excel(p)
        for _,r in d.iterrows():
            rows.append(dict(file=p.split('/')[-1],name=r['BNF_NAME'],amt=float(r['AMOUNT']),narr=r['CREDIT_NARR'],date=r['PYMT_DATE']))
    b=pd.DataFrame(rows)
    return b.drop_duplicates(subset=['name','amt','narr','date']).reset_index(drop=True)  # same file uploaded twice
def main():
    ap=argparse.ArgumentParser(); ap.add_argument('summary'); ap.add_argument('bank',nargs='+'); ap.add_argument('-o','--out',default='payroll_reconciliation.xlsx')
    ap.add_argument('--tab',action='append',default=[]); ap.add_argument('--extra',help='CSV of payments made outside the bank files: month,employee,amount,note'); a=ap.parse_args()
    for t in a.tab: k,v=t.split('=',1); TABS[k]=v
    S=load_summary(a.summary); B=load_bank(a.bank)
    known=sorted(S.key.unique())
    # assign each bank line to a pay month: the summary month it relates to = month named in narration, else file month
    if a.extra:
        E=pd.read_csv(a.extra); E=pd.DataFrame(dict(file='manual: '+E.note.fillna('').astype(str),name=E.employee,amt=E.amount.astype(float),narr='Paid separately',date=None,man=E.month))
        B=pd.concat([B,E],ignore_index=True)
    if 'man' not in B: B['man']=None
    B['key']=B.name.map(lambda n:match(n,known))
    B['month']=[month_of(n,None) for n in B.narr]
    # narration mismatch check: date-based expected month = month before payment date
    def pay_month(dt):
        d=pd.to_datetime(dt,dayfirst=True); i=(d.month-2)%12   # paid in Jun -> May salary
        return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][i]
    B['exp_month']=[pay_month(d) if d else None for d in B.date]
    # May's late file (22-Jun) and normal salary paid next month => use expected month from date
    B['month']=[m if isinstance(m,str) else e for m,e in zip(B.man,B.exp_month)]
    ex=[]; det=[]
    for m in MONTHS:
        sm=S[S.month==m]; bm=B[B.month==m]
        for k in sorted(set(sm.key)|set(bm.key.dropna())):
            s=sm[sm.key==k]; b=bm[bm.key==k]
            paid=b.amt.sum() if len(b) else 0
            if len(s):
                s=s.iloc[0]; exp=s.svsg if s.svsg is not None else None
                det.append([m,s['name'],s.salary,s.days,s.payable,s.pt,s.loan,s.net,exp,s.fs,paid,s.entity])
            else:
                det.append([m,b.name.iloc[0],None,None,None,None,None,None,None,None,paid,'NOT IN SUMMARY'])
                ex.append([m,b.name.iloc[0],'Paid but not in summary',paid])
        for _,r in bm[bm.key.isna()].iterrows(): ex.append([m,r['name'],'Bank name not matched to any employee',r.amt])
        for _,r in bm.iterrows():
            if not str(r.file).startswith('manual') and not re.search(r'conveyance',str(r.narr),re.I) and month_of(r.narr,r.exp_month)!=r.exp_month:
                ex.append([m,r['name'],f'Narration "{r.narr}" does not match pay month {r.exp_month}',r.amt])
    D=pd.DataFrame(det,columns=['Month','Employee','Salary','Days payable','Salary payable','PT','Loan','Net payable','SVSG share (expected via this bank file)','Finstrat share','Paid via bank file','Entity'])
    # rules
    for _,r in D.iterrows():
        m,n=r.Month,r.Employee
        e=r['SVSG share (expected via this bank file)']; p=r['Paid via bank file']
        if e is not None and pd.notna(e):
            if e>0 and p==0: ex.append([m,n,'SVSG share not paid in bank file',e])
            elif abs(p-e)>TOL and p>0: ex.append([m,n,f'Paid {p:,.0f} vs SVSG share {e:,.0f} (diff {p-e:+,.0f})',p-e])
        elif r['Net payable'] and p>0 and abs(p-r['Net payable'])>TOL and r.Entity!='NOT IN SUMMARY':
            ex.append([m,n,f'No SVSG/FS split in summary; paid {p:,.0f} vs net {r["Net payable"]:,.0f}',p-r['Net payable']])
        if r.PT not in (0,200,400) and pd.notna(r.PT): ex.append([m,n,f'Unusual PT value {r.PT:,.0f}',r.PT])
        for c in ('SVSG share (expected via this bank file)','Finstrat share'):
            if pd.notna(r[c]) and r[c]<0: ex.append([m,n,f'Negative {c.split(" ")[0]} amount',r[c]])
    P=S.sort_values('month',key=lambda s:s.map(MONTHS.index)).groupby('key')
    prev={}; gone=set()
    for m in MONTHS:
        for _,r in S[S.month==m].iterrows():
            if r.key in prev:
                o=prev[r.key]
                if r.salary!=o.salary: ex.append([m,r['name'],f'Salary changed {o.salary:,.0f} -> {r.salary:,.0f}',r.salary-o.salary])
                if r.loan!=o.loan and (r.loan or o.loan): ex.append([m,r['name'],f'Loan deduction changed {o.loan:,.0f} -> {r.loan:,.0f}',r.loan-o.loan])
            prev[r.key]=r
        cur=set(S[S.month==m].key)
        for k,o in list(prev.items()):
            if k not in cur and k not in gone: ex.append([m,o['name'],'Dropped off summary (exit?)',None]); gone.add(k)
    X=pd.DataFrame(ex,columns=['Month','Employee','Exception','Amount']); X['o']=X.Month.map(MONTHS.index); X=X.sort_values(['o','Employee']).drop(columns='o')
    # ---- write
    wb=Workbook(); hd=Font(bold=True,color='FFFFFF'); fill=PatternFill('solid',fgColor='1F3864'); red=PatternFill('solid',fgColor='F8CBAD')
    ws=wb.active; ws.title='Monthly summary'
    ws.append(['Month','Net payable (summary)','SVSG share','Finstrat share','Paid via SVSG bank file','Bank lines'])
    for i,m in enumerate(MONTHS,2):
        d=D[D.Month==m]
        ws.append([m,d['Net payable'].sum(),d['SVSG share (expected via this bank file)'].sum() or None,d['Finstrat share'].sum() or None,d['Paid via bank file'].sum(),int((B.month==m).sum())])
    n=len(MONTHS)+1; ws.append(['Total']+[f'=SUM({c}2:{c}{n})' for c in 'BCDEF'])
    ws=wb.create_sheet('Detail'); ws.append(list(D.columns)+['Variance (paid - SVSG share, or - net if no split)'])
    for i,row in enumerate(D.itertuples(index=False),2):
        ws.append([None if (isinstance(v,float) and pd.isna(v)) else v for v in row]+[f'=IF(I{i}="",IF(H{i}="","",K{i}-H{i}),K{i}-I{i})'])
    ws=wb.create_sheet('Employee-wise')
    cols=['Employee','Month','Days payable','Monthly salary','Salary payable','PT','Loan','Net salary','SVSG share','Finstrat share','Paid via SVSG bank','Entity']
    ws.append(cols)
    src=['Employee','Month','Days payable','Salary','Salary payable','PT','Loan','Net payable','SVSG share (expected via this bank file)','Finstrat share','Paid via bank file','Entity']
    E2=D.copy(); E2['k']=E2.Employee.map(lambda n:norm(ALIAS.get(norm(n),n)))
    for k in sorted(E2.k.unique()):
        g=E2[E2.k==k]; g=g.sort_values('Month',key=lambda s:s.map(MONTHS.index)); first=ws.max_row+1
        for _,r in g.iterrows():
            ws.append([None if (isinstance(r[c],float) and pd.isna(r[c])) else r[c] for c in src])
        last=ws.max_row; nm=g.Employee.iloc[0]
        ws.append([f'{nm} - TOTAL','',f'=SUM(C{first}:C{last})','']+[f'=SUM({c}{first}:{c}{last})' for c in 'EFGHIJK'])
        for c in ws[ws.max_row]: c.font=Font(bold=True); c.fill=PatternFill('solid',fgColor='D9E1F2')
        ws.append([])
    ws=wb.create_sheet('Exceptions'); ws.append(list(X.columns))
    for r in X.itertuples(index=False): ws.append([None if (isinstance(v,float) and pd.isna(v)) else v for v in r])
    ws=wb.create_sheet('Bank lines'); ws.append(['Pay month','Employee (bank name)','Matched to','Amount','Narration','Date','Source file'])
    for r in B.sort_values('month',key=lambda s:s.map(MONTHS.index)).itertuples(): ws.append([r.month,r.name,r.key,r.amt,r.narr,r.date,r.file])
    for w in wb:
        for c in w[1]: c.font=hd; c.fill=fill; c.alignment=Alignment(wrap_text=True,vertical='center')
        w.freeze_panes='A2'
        for i,col in enumerate(w.columns,1): w.column_dimensions[get_column_letter(i)].width=min(60,max(12,max(len(str(c.value or '')) for c in col)+2))
        for row in w.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value,(int,float)): c.number_format='#,##0;[Red]-#,##0'
    wb['Exceptions'].column_dimensions['C'].width=70
    wb.save(a.out); print('saved',a.out,'| exceptions:',len(X)); print(X.to_string())
main()
