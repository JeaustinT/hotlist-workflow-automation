"""Rebuild clean synthetic workbooks and fixtures; never reads operational files."""
from pathlib import Path
import csv
import json
import io
import os
import tempfile
from verify_embedded_vba import verify
from datetime import datetime, timezone
import xlsxwriter
from vba_container import make_project

ROOT = Path(__file__).resolve().parents[1]
VERSION = '0.3.2'
RAW = ['Location','Publication ID','Publication Name','Client ID','Client Name','Vendor ID','Vendor Name','Invoice Number','Line ID','Batch ID','Invoice Date','Amount','Input Date','Status','Hotlist Comment 1','Hotlist Comment 2','Activity Date','Agent','Agent Comment','Currency']
HEADERS = RAW + ['Assigned Analyst','Invoice Type','Run ID','Invoice Key','Group Key','Assignment Basis','Volume Class','Resolved Media','Workstream','Exclusion Reason','Run At UTC','Work Saved At UTC','Duplicate in Run']
ASSIGNED_HEADERS = ['Assigned Analyst', 'Invoice Type', 'Run ID', 'Location', 'Publication ID', 'Publication Name', 'Client ID', 'Invoice Number', 'Invoice Date', 'Amount', 'Vendor ID', 'Batch ID', 'Status', 'Hotlist Comment 1', 'Hotlist Comment 2', 'Agent', 'Agent Comment', 'Client Name', 'Vendor Name', 'Input Date', 'Activity Date', 'Currency', 'Line ID', 'Resolved Media', 'Workstream', 'Group Key', 'Assignment Basis', 'Volume Class', 'Duplicate in Run', 'Invoice Key', 'Exclusion Reason', 'Run At UTC', 'Work Saved At UTC']

TEAM = [('DEMO-ANL-A','Alex Morgan','alex.morgan@example.com','YES',1),('DEMO-ANL-B','Jordan Lee','jordan.lee@example.com','YES',.5),('DEMO-ANL-C','Taylor Smith','taylor.smith@example.com','YES',1),('DEMO-ANL-D','Casey Brown','casey.brown@example.com','NO',1),('DEMO-ANL-E','Morgan Davis','morgan.davis@example.com','NO',1)]
PRIOR_ID='20270110T090000Z-11111111-1111-4111-8111-111111111111'
PRIOR_TIME='2027-01-10T09:00:00Z'
RUN_TIME='2027-01-12T09:00:00Z'


def part(value):
    value = str(value or '').strip().upper()
    return f'{len(value)}:{value}'


def invoice(row):
    return ''.join(part(row[i]) for i in [0,5,3,7])


def group(row):
    return ''.join(part(row[i]) for i in [0,5,3])


def synthetic():
    rows=[]
    for n in range(60):
        g=0 if n<8 else 1 if n<12 else 2 if n<24 else 3+(n-24)//4
        loc='DEMO-LOC-A' if g%2==0 else 'DEMO-LOC-B'
        cli=f'DEMO-CLI-{g:02d}'; ven=f'DEMO-VEN-{g:02d}'
        pub='DEMO-PUB-D' if n%2==0 else 'DEMO-PUB-F'
        rows.append([loc,pub,'Demo Digital Publication',cli,f'Demo Client {g:02d}',ven,f'Demo Vendor {g:02d}',f'DEMO-INV-{n+1:03d}',f'DEMO-LINE-{n+1:03d}','DEMO-BAT-02','2027-01-05',round(125+7.25*n,2),'2027-01-12','Pending','','','','','','USD'])
    # New activity must balance rather than return to an owner, regardless of comments.
    for n in [5,6,7]:
        rows[n][14]='Demo follow-up recorded';rows[n][16]='2027-01-11'
    # Large new group is deliberately indivisible (rows 13 through 24).
    for n in range(48,52): rows[n][1]='DEMO-PUB-P';rows[n][2]='Demo Print Publication'
    for n in range(52,55): rows[n][0]='DEMO-LOC-X'
    for n in range(55,57): rows[n][13]='Released to Clear'
    for n in range(57,60): rows[n][1]='DEMO-PUB-U';rows[n][2]='Demo Unknown Publication'
    # Same invoice number, different invoice identity.
    rows[28][7]=rows[0][7]
    # Two source lines of one invoice in the large new group.
    rows[14][7]=rows[13][7]
    prior=[]
    for n in range(12):
        row=rows[n][:]
        row[12]='2027-01-10';row[14]='';row[15]='';row[16]=''
        owner='DEMO-ANL-A' if n<8 else 'DEMO-ANL-E'
        row[17]=owner if n==0 else ''
        row[18]='Demo prior note to demonstrate clearing' if n==0 else ''
        prior.append(row+[owner,'NEW',PRIOR_ID,invoice(row),group(row),'Weighted balanced allocation','NEW VOLUME','DIGITAL','DIGITAL','',PRIOR_TIME,'',False])
    return rows,prior


def reference_model(rows,history,team=TEAM):
    """Executable specification; this is not a VBA runtime test."""
    from collections import defaultdict, Counter
    people={x[0]:x for x in team if x[3]=='YES'}
    loads={x:0 for x in people}; groups=defaultdict(list); votes=defaultdict(Counter)
    latest={}
    for row in sorted(history,key=lambda r:(r[30],r[22])):latest[invoice(row)]=row
    assigned=[];exceptions=[]
    for raw in rows:
        r=raw[:]+['']*13
        key=invoice(r);g=group(r)
        r[23]=key;r[24]=g;r[21]='REPEATED' if key in latest else 'NEW'
        active=bool(r[16] and r[16]>'2027-01-10')
        r[26]='NEW VOLUME' if key not in latest or active else 'CARRYOVER / REVIEW'
        reason='';stream='REVIEW'
        if r[13]=='Released to Clear':reason='Released to Clear'
        elif r[0] not in ['DEMO-LOC-A','DEMO-LOC-B']:reason='Out-of-scope location'
        else:
            stream={'DEMO-PUB-D':'DIGITAL','DEMO-PUB-F':'DIGITAL','DEMO-PUB-P':'PRINT'}.get(r[1],'REVIEW')
            if stream=='PRINT':reason='Print media'
            if stream=='REVIEW':reason='Unresolved media classification'
        r[27]=r[28]=stream;r[29]=reason
        if reason:exceptions.append(r);continue
        groups[g].append(r)
        if key in latest and not active and not r[14] and not r[15] and latest[key][20] in people:
            votes[g][latest[key][20]]+=1
    for g in sorted(groups,key=lambda g:(-len(groups[g]),g)):
        batch=groups[g]
        if votes[g]:
            owner=min(votes[g],key=lambda p:(-votes[g][p],loads[p]/people[p][4],p));basis='Repeated untouched - previous owner'
        else:
            owner=min(people,key=lambda p:((loads[p]+len(batch))/people[p][4],loads[p],p));basis='Weighted balanced allocation'
        loads[owner]+=len(batch)
        for r in batch:r[20]=owner;r[25]=basis
    assigned=sorted((r for batch in groups.values() for r in batch),key=lambda r:r[8])
    for r in assigned:r[32]=sum(invoice(x)==invoice(r) for x in rows)>1
    return assigned,exceptions,loads


def csv_write(name,headers,rows):
    with (ROOT/'demo/sample-data'/name).open('w',newline='',encoding='utf-8') as f:
        w=csv.writer(f,lineterminator="\n");w.writerow(headers);w.writerows(rows)


def excel_value(value, header):
    if value and header in ['Invoice Date','Input Date','Activity Date','Run At UTC','Work Saved At UTC','Duplicate in Run']:
        return datetime.fromisoformat(str(value).replace('Z','+00:00')).replace(tzinfo=None)
    return value


def write_table(book,name,headers,rows,visible=True):
    sheet=book.add_worksheet(name)
    sheet.set_vba_name('Sheet'+str(len(book.worksheets())))
    sheet.freeze_panes(1,0);sheet.hide_gridlines(2)
    sheet.set_tab_color('#1F4E78' if visible else '#8A9AA8')
    sheet.set_column(0,len(headers)-1,20)
    sheet.set_row(0,32)
    date=book.add_format({'num_format':'yyyy-mm-dd'})
    stamp=book.add_format({'num_format':'yyyy-mm-dd hh:mm:ss'})
    money=book.add_format({'num_format':'#,##0.00'})
    clean=[[excel_value(v,h) for h,v in zip(headers,r)] for r in rows]
    columns=[]
    for h in headers:
        c={'header':h}
        if h in ['Invoice Date','Input Date','Activity Date']:c['format']=date
        if h in ['Run At UTC','Work Saved At UTC']:c['format']=stamp
        if h=='Amount':c['format']=money
        if name=='Assigned Work' and h in ['Agent','Agent Comment']:
            c['format']=book.add_format({'bg_color':'#FFF2CC','locked':False})
            c['header_format']=book.add_format({'bold':True,'font_color':'#17365D','bg_color':'#FFE699','text_wrap':True,'border':1})
        columns.append(c)
    sheet.add_table(0,0,max(1,len(clean)),len(headers)-1,{'columns':columns,'data':clean,'style':'Table Style Medium 2','autofilter':True})
    if name=='Assigned Work':
        # Freeze assignment and invoice identity; source fields retain canonical storage elsewhere.
        sheet.freeze_panes(1,8)
        sheet.set_zoom(85)
        sheet.set_column('A:A',18);sheet.set_column('B:B',12);sheet.set_column('C:C',16)
        sheet.set_column('D:E',14);sheet.set_column('F:F',20);sheet.set_column('G:G',14);sheet.set_column('H:H',18)
        sheet.set_column('I:I',12);sheet.set_column('J:J',12);sheet.set_column('K:M',15)
        sheet.set_column('N:O',27)
        progress=book.add_format({'bg_color':'#FFF2CC','locked':False})
        sheet.set_column('P:P',18,progress);sheet.set_column('Q:Q',38,progress)
        sheet.set_column('R:W',20,None,{'level':1})
        sheet.set_column('X:AG',20,None,{'hidden':True,'level':1})
    elif name in ['Exceptions & Review','Inventory History']:
        # Preserve readable operational columns, keep technical detail accessible.
        for first,last in [(2,2),(4,4),(6,6),(9,9),(19,19),(23,24),(27,28),(30,31)]:sheet.set_column(first,last,20,None,{'hidden':True})
        sheet.set_column(18,18,36)
    if not visible:sheet.hide()
    return sheet


def build():
    rows,history=synthetic();assigned,exceptions,loads=reference_model(rows,history)
    csv_write('raw_hotlist.csv',RAW,rows)
    csv_write('history_seed.csv',HEADERS,history)
    csv_write('expected_assignments.csv',['Line ID','Invoice Key','Group Key','Invoice Type','Volume Class','Assigned Analyst','Assignment Basis','Duplicate in Run'],[[r[i] for i in [8,23,24,21,26,20,25,32]] for r in assigned])
    csv_write('expected_exceptions.csv',['Line ID','Exclusion Reason'],[[r[8],r[29]] for r in exceptions])
    csv_write('availability_profiles.csv',['Profile','Analyst ID','Available','Weight'],[[profile,t[0],('YES' if t[0] in active else 'NO'),t[4]] for profile,active in [('DEFAULT',{'DEMO-ANL-A','DEMO-ANL-B','DEMO-ANL-C'}),('ALTERNATE',{'DEMO-ANL-B','DEMO-ANL-D','DEMO-ANL-E'}),('ALL',set(t[0] for t in TEAM))] for t in TEAM])
    modules={p.stem:(p.read_text(),False) for p in sorted((ROOT/'src/vba').glob('*.bas'))}
    for n in ['ThisWorkbook']+['Sheet'+str(i) for i in range(1,11)]:
        base='00020819' if n=='ThisWorkbook' else '00020820'
        source=f'Attribute VB_Name = "{n}"\nAttribute VB_Base = "0{{{base}-0000-0000-C000-000000000046}}"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\nOption Explicit\n'
        modules[n]=(source,True)
    project=make_project(modules)
    final_path=ROOT/'demo/Hotlist_Workflow_Automation_Demo.xlsm'
    with tempfile.TemporaryDirectory(prefix='.hotlist-build-',dir=final_path.parent) as stage:
        path=Path(stage)/final_path.name
        book=xlsxwriter.Workbook(path,{'strings_to_formulas':False,'strings_to_urls':False})
        book.set_properties({'title':'Hotlist Workflow Automation — Synthetic Demo','author':'Hotlist Demo','company':'Fictional Demonstration','comments':'Synthetic data only. Desktop validation pending.','created':datetime(2027,1,1)})
        book.set_vba_name('ThisWorkbook')
        if book.add_vba_project(io.BytesIO(project),is_stream=True)!=0:
            raise RuntimeError('VBA project embedding failed')
        control=book.add_worksheet('Control');control.set_vba_name('Sheet1');control.hide_gridlines(2)
        control.set_column('A:A',19);control.set_column('B:B',23);control.set_column('C:D',14);control.set_column('E:F',3);control.set_column('G:H',23);control.set_column('J:J',24);control.set_column('K:K',26)
        title=book.add_format({'bold':True,'font_size':22,'font_color':'#FFFFFF','bg_color':'#16324F'})
        label=book.add_format({'bold':True,'font_color':'#16324F'});note=book.add_format({'text_wrap':True,'font_color':'#43566B'})
        control.merge_range('A1:K2','HOTLIST WORKFLOW AUTOMATION',title)
        control.merge_range('A3:H3','SYNTHETIC DEMO 0.3.2 — VBA Desktop validation pending',label)
        for col,h in enumerate(['Analyst ID','Analyst Name','Available?','Weight']):control.write(3,col,h,label)
        for n,t in enumerate(TEAM,4):
            control.write_row(n,0,[t[0],t[1],t[3],t[4]])
        control.set_column('D:D',14,book.add_format({'num_format':'0%'}))
        control.data_validation('C5:C9',{'validate':'list','source':['YES','NO']})
        control.data_validation('D5:D9',{'validate':'decimal','criteria':'>','value':0})
        control.write('G4','Allowed locations',label);control.write('G5','DEMO-LOC-A');control.write('G6','DEMO-LOC-B')
        for row,caption,macro in [(10,'SAVE PROGRESS','SaveProgress'),(12,'PROCESS HOTLIST','ProcessHotlist'),(14,'CREATE OUTLOOK DRAFT','CreateOutlookDraft'),(16,'EXPORT HISTORY','ExportHistory')]:
            control.insert_button(row-1,6,{'macro':macro,'caption':caption,'width':245,'height':32})
        metadata=[('Demo version',VERSION),('Last completed run',PRIOR_ID),('Last run UTC',datetime(2027,1,10,9)),('Input signature',''),('',''),('Lock token',''),('Lock user / action',''),('Lock time UTC',''),('Run status','COMPLETED'),('',''),('',''),('Fault injection','')]
        for n,(k,v) in enumerate(metadata,3):
            control.write(n,9,k,label)
            if isinstance(v,datetime):control.write_datetime(n,10,v,book.add_format({'num_format':'yyyy-mm-dd hh:mm:ss'}))
            else:control.write(n,10,v)
        control.merge_range('A11:D15','Ready for a second synthetic run: 60 input rows. Assigned Work and Inventory History contain a 12-row seed, generated in Python, not executed by VBA. Edit progress, then process the new input.',note)
        control.merge_range('A17:D21','Keep a saved recovery copy. One coordinator runs actions while analysts pause edits. Lock recovery: Alt+F8 → RecoverStaleLock, after 30 minutes and coordinator confirmation.',note)
        control.freeze_panes(4,0)
        write_table(book,'Raw Hotlist',RAW,rows)
        write_table(book,'Assigned Work',ASSIGNED_HEADERS,[[row[HEADERS.index(h)] for h in ASSIGNED_HEADERS] for row in history])
        write_table(book,'Exceptions & Review',HEADERS,[])
        summary=book.add_worksheet('Assignment Summary');summary.set_vba_name('Sheet5');summary.hide_gridlines(2);summary.freeze_panes(1,0);summary.set_column('A:K',21)
        summary.write_row('A1',['Analyst ID','Analyst Name','Weight','New Volume','Carryover / Review','Assigned','Weighted Target','Difference'],label)
        counts={t[0]:sum(r[20]==t[0] for r in history) for t in TEAM}
        for n,t in enumerate([TEAM[0],TEAM[1],TEAM[4]],1):summary.write_row(n,0,[t[0],t[1],1,counts[t[0]],0,counts[t[0]],4,counts[t[0]]-4])
        summary.write('A5','TOTAL');summary.write('F5',12);summary.write('J2','Run ID');summary.write('K2',PRIOR_ID);summary.write('J3','Exceptions');summary.write('K3',0)
        write_table(book,'Inventory History',HEADERS,history,False)
        exact=sorted({(r[0],r[3],r[1],'DIGITAL' if r[1]=='DEMO-PUB-D' else 'PRINT') for r in rows if r[1] in ['DEMO-PUB-D','DEMO-PUB-P'] and r[0]!='DEMO-LOC-X'})
        write_table(book,'Media Reference',['Location','Client ID','Publication ID','Workstream'],exact,False)
        write_table(book,'Publication Fallback',['Publication ID','Workstream'],[['DEMO-PUB-D','PRINT'],['DEMO-PUB-F','DIGITAL'],['DEMO-PUB-P','PRINT']],False)
        write_table(book,'Team Directory',['Analyst ID','Analyst Name','Email'],[t[:3] for t in TEAM],False)
        instructions=book.add_worksheet('Instructions');instructions.set_vba_name('Sheet10');instructions.set_column('A:A',115);instructions.hide_gridlines(2)
        for n,text in enumerate(['Synthetic demo only; desktop compile and action validation required.','1. Save a recovery copy. Verify Control availability, weights, and allowed locations.','2. SAVE PROGRESS reconciles Agent and Agent Comment. Intentional clearing asks for confirmation.','3. PROCESS HOTLIST handles the 60 ready input rows; prior progress must reconcile.','4. Inspect Assigned Work, Exceptions & Review, and Assignment Summary.','5. CREATE OUTLOOK DRAFT opens a fictional-address draft; it never sends automatically.','6. EXPORT HISTORY writes a macro-free synthetic snapshot.','Assigned Work: I2 freezes ownership and invoice identity. Edit highlighted Agent and Agent Comment in P:Q.', '7. RecoverStaleLock is available through Alt+F8 only after 30 minutes and coordinator confirmation.','Activity Date is typed and displayed yyyy-mm-dd. New activity is strictly after the previous run calendar date.','Invoice identity: Location + Vendor + Client + Invoice Number. Line ID: unique source row. Group Key: Location + Vendor + Client.','The preloaded seed and expected CSVs are Python-calculated fixtures, not evidence of executed VBA.']):instructions.write(n,0,text)
        instructions.hide();control.activate()
        book.close()
        verify(path,ROOT/'src/vba',VERSION)
        os.replace(path,final_path)
    # Reopen the published artifact, not the in-memory project or staged inputs.
    receipt=verify(final_path,ROOT/'src/vba',VERSION)
    (ROOT/'demo/embedded-vba-manifest.json').write_text(json.dumps(receipt,indent=2)+'\n')
    sample=ROOT/'demo/sample-data/Sample_Raw_Hotlist.xlsx'
    book=xlsxwriter.Workbook(sample,{'strings_to_formulas':False,'strings_to_urls':False})
    book.set_properties({'author':'Hotlist Demo','title':'Synthetic Raw Hotlist','created':datetime(2027,1,1)})
    write_table(book,'Raw Hotlist',RAW,rows);book.close()
    print(json.dumps({'source_rows':len(rows),'seed_rows':len(history),'expected_assigned':len(assigned),'expected_exceptions':len(exceptions),'expected_loads':loads,'workbook':str(final_path),'workbook_sha256':receipt['workbook_sha256']},indent=2))

if __name__=='__main__':build()
