"""Structural and executable-specification checks; does not execute VBA.

Run with Python, openpyxl, XlsxWriter, and oletools installed.
"""
import sys
from pathlib import Path
import csv
import json
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from openpyxl import load_workbook
from oletools.olevba import VBA_Parser
import olefile
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from verify_embedded_vba import verify
from build_demo import synthetic, reference_model, invoice, group, TEAM, RAW, HEADERS, ASSIGNED_HEADERS, part, excel_value


def main():
    checks=[]
    def check(value,label):
        assert value,label
        checks.append(label)
    path=ROOT/'demo/Hotlist_Workflow_Automation_Demo.xlsm'
    w=load_workbook(path,read_only=False,keep_vba=True)
    check(w.sheetnames==['Control','Raw Hotlist','Assigned Work','Exceptions & Review','Assignment Summary','Inventory History','Media Reference','Publication Fallback','Team Directory','Instructions'],'worksheet order')
    check([s.title for s in w if s.sheet_state=='visible']==w.sheetnames[:5],'five visible sheets')
    rows,history=synthetic();a,e,loads=reference_model(rows,history)
    check(w['Raw Hotlist'].max_row==61 and w['Inventory History'].max_row==13,'fixture row counts')
    check([c.value for c in w['Raw Hotlist'][1]]==RAW,'raw schema')
    check([c.value for c in w['Inventory History'][1]]==HEADERS,'history schema')
    queue=w['Assigned Work']
    check([c.value for c in queue[1]]==ASSIGNED_HEADERS,'analyst-first Assigned Work schema')
    check(ASSIGNED_HEADERS[:8]==['Assigned Analyst','Invoice Type','Run ID','Location','Publication ID','Publication Name','Client ID','Invoice Number'],'frozen assignment and invoice context fields')
    check(queue.freeze_panes=='I2' and queue.sheet_view.pane.xSplit==8 and queue.sheet_view.pane.ySplit==1,'Assigned Work I2 freezes header and A:H')
    check(queue['P1'].value=='Agent' and queue['Q1'].value=='Agent Comment','adjacent progress fields at P:Q')
    check(all(not queue.cell(2,c).protection.locked and queue.cell(2,c).fill.fgColor.rgb=='FFFFF2CC' for c in [16,17]) and not queue.protection.sheet,'progress cells highlighted and editable')
    def hidden_at(column):
        return any(d.hidden and d.min<=column<=d.max for d in queue.column_dimensions.values())
    check(all(not hidden_at(c) for c in range(1,18)),'primary A:Q fields visible')
    check(all(hidden_at(c) for c in range(24,34)),'technical X:AG fields hidden to the right')
    check(queue['AC2'].data_type=='b' and queue['AC2'].number_format=='General','duplicate flag is boolean, not date-formatted')
    actual_canonical=[]
    for n,seed in enumerate(history,2):
        actual_canonical.append([queue.cell(n,ASSIGNED_HEADERS.index(h)+1).value for h in HEADERS])
    expected_seed=[[excel_value(v,h) if v!='' else None for h,v in zip(HEADERS,row)] for row in history]
    check(actual_canonical==expected_seed,'all seed fields round-trip through displayed column order')
    mapping=[ASSIGNED_HEADERS.index(h)+1 for h in HEADERS]
    utilities=(ROOT/'src/vba/DemoUtilities.bas').read_text()
    mapping_body=utilities.split('Public Function WorksheetColumn',1)[1].split('End Function',1)[0]
    vba_mapping=[int(n) for n in re.search(r'displayColumns = Array\(([^)]+)\)',mapping_body)[1].split(',')]
    check(vba_mapping==mapping and sorted(mapping)==list(range(1,34)),'VBA display mapping is complete and matches builder')
    queue['P2']='DEMO-ANL-C';queue['Q2']='Demo progress mapping check'
    edited=[queue.cell(2,mapping[i]).value for i in range(33)]
    check(edited[17:19]==['DEMO-ANL-C','Demo progress mapping check'],'progress edits map to canonical Agent fields')
    check(edited[22]==history[0][22] and invoice(edited)==history[0][23] and edited[8]==history[0][8],'progress key unchanged by analyst layout')
    queue['P2']='';queue['Q2']=''
    check([queue.cell(2,mapping[i]).value for i in [17,18]]==['',''],'intentional clearing maps to both progress values')
    # All mutations above are in-memory; the distributed workbook is never saved here.
    check(all([[row[HEADERS.index(h)] for h in ASSIGNED_HEADERS][column-1] for column in mapping]==row for row in a),'calculated assignment rows round-trip through display mapping')
    office=(ROOT/'src/vba/OfficeIntegration.bas').read_text()
    check('With exportWorkbook.Windows(1)' in office and '.SplitRow = 1' in office and '.SplitColumn = 0' in office and '.FreezePanes = True' in office,'history export source retains header-only freeze')

    check(w['Control']['K4'].value=='0.3.2' and w['Control']['K12'].value=='COMPLETED','version and seed run metadata')
    check(all(w['Control'][c].value is None for c in ['K9','K10','K11','K15']),'unlocked, fault injection off')
    check(not w.defined_names and not w._external_links,'no named ranges or external links')
    check(all(c.data_type!='f' for s in w for row in s for c in row),'no workbook formulas or formula injection')
    sample=load_workbook(ROOT/'demo/sample-data/Sample_Raw_Hotlist.xlsx',read_only=True)
    check(sample['Raw Hotlist'].max_row==61,'sample xlsx row count');sample.close()
    check(len({r[8] for r in rows})==60,'unique source Line IDs')
    seed_owners=defaultdict(set)
    for h in history:seed_owners[(h[22],h[24])].add(h[20])
    check(all(len(x)==1 for x in seed_owners.values()),'historical seed group integrity')
    check(all(invoice(h)==h[23] and group(h)==h[24] for h in history),'history identity consistency')
    check(len(a)==48 and len(e)==12 and len(a)+len(e)==60,'48 assigned and 12 exceptions reconcile')
    check(loads=={'DEMO-ANL-A':20,'DEMO-ANL-B':8,'DEMO-ANL-C':20},'default capacity allocation')
    owners=defaultdict(set)
    for r in a:owners[r[24]].add(r[20])
    check(all(len(x)==1 for x in owners.values()),'one owner per group')
    check(Counter(r[29] for r in e)=={'Print media':4,'Out-of-scope location':3,'Released to Clear':2,'Unresolved media classification':3},'exception reasons')
    check(sum(r[21]=='REPEATED' for r in a)==12,'repeat identities')
    check(sum(r[26]=='NEW VOLUME' for r in a)==39,'39 new-volume rows, including three active repeats')
    check(all(r[20]=='DEMO-ANL-A' and r[25]=='Repeated untouched - previous owner' for r in a if int(r[8][-3:])<=8),'available historical owner preserves group')
    check(all(r[25]=='Weighted balanced allocation' for r in a if 9<=int(r[8][-3:])<=12),'unavailable historical owner uses balancing')
    check(invoice(rows[0])!=invoice(rows[28]) and rows[0][7]==rows[28][7],'same invoice number remains distinct across vendors')
    check(invoice(rows[13])==invoice(rows[14]) and rows[13][8]!=rows[14][8],'multi-line invoice identity')
    check(sum(bool(r[32]) for r in a)==2,'duplicate invoice flagged without dropping distinct lines')
    check(max(Counter(r[24] for r in a).values())==12,'large group remains intact')
    # Simulate different past runs containing subsets of a group; vote priority is explicit.
    vote_history=[h[:] for h in history]
    for n,h in enumerate(vote_history):
        h[22]='DEMO-PAST-RUN-'+str(n)
        h[20]='DEMO-ANL-A' if n<3 else 'DEMO-ANL-B' if n<8 else 'DEMO-ANL-E'
    vote_a,_,_=reference_model(rows,vote_history)
    check(all(r[20]=='DEMO-ANL-A' for r in vote_a if int(r[8][-3:])<=8),'highest available-owner vote wins')
    # Remove all other group rows so active-repeat behavior is independently exposed.
    active=[r for r in rows[5:8]]
    active_result,_,_=reference_model(active,history)
    check(all(r[25]=='Weighted balanced allocation' for r in active_result),'active repeats use balancing independently')
    alternate=[(t[0],t[1],t[2],'YES' if t[0] in {'DEMO-ANL-B','DEMO-ANL-D','DEMO-ANL-E'} else 'NO',t[4]) for t in TEAM]
    alternate_a,_,alternate_loads=reference_model(rows,history,alternate)
    check(all(r[20] in alternate_loads for r in alternate_a),'alternate availability profile')
    check(invoice([str(v).lower() if isinstance(v,str) else v for v in rows[0]])==invoice(rows[0]),'case-normalized identity')
    check(part('01')!=part('1') and part('a:b')+part('c')!=part('a')+part('b:c'),'leading zeros and delimiter-safe keys')
    sig=''.join(part(v) for r in rows for v in r)
    # VBA dates are Excel serials: their representations are shorter than fixture ISO dates.
    check(len(sig)<32767,'small-demo exact signature fits a cell')
    with (ROOT/'demo/sample-data/expected_assignments.csv').open() as f:
        expected=list(csv.DictReader(f))
    check(len(expected)==48 and Counter(x['Assigned Analyst'] for x in expected)==Counter(loads),'expected assignment fixture')
    # These are authored domain assertions above, not claims of Excel execution.
    with zipfile.ZipFile(path) as z:
        check(not any(any(k in n.lower() for k in ['externallink','connection','pivotcache','embedding','customxml','comments']) for n in z.namelist()),'no connections, caches, external links, embeddings, or comments')
        check(z.testzip() is None,'zip package integrity')
        vb=z.read('xl/vbaProject.bin')
        compound=olefile.OleFileIO(io.BytesIO(vb))
        check(not any('__SRP' in '/'.join(n) for n in compound.listdir()),'no compiled SRP caches')
        check(compound.openstream('VBA/_VBA_PROJECT').read()==bytes.fromhex('cc61ffff000000'),'source-only project header')
        vml=''.join(z.read(n).decode() for n in z.namelist() if n.endswith('.vml'))
        check([vml.index(x) for x in ['SAVE PROGRESS','PROCESS HOTLIST','CREATE OUTLOOK DRAFT','EXPORT HISTORY']]==sorted(vml.index(x) for x in ['SAVE PROGRESS','PROCESS HOTLIST','CREATE OUTLOOK DRAFT','EXPORT HISTORY']),'button order')
        for macro in ['SaveProgress','ProcessHotlist','CreateOutlookDraft','ExportHistory']:check(macro in vml,'button binding '+macro)
    final_report=verify(path,ROOT/'src/vba','0.3.2')
    for name,entry in final_report['modules'].items():
        check(entry['match'] and entry['direct_decompression_match'],'final ZIP/CFB source verification '+name)
    check(final_report==json.loads((ROOT/'demo/embedded-vba-manifest.json').read_text()),'published workbook and module receipt match')
    p=VBA_Parser(str(path));extracted={name:code.replace('\r\n','\n') for _,_,name,code in p.extract_macros()}
    for module in (ROOT/'src/vba').glob('*.bas'):
        check(extracted[module.name]==module.read_text(),'embedded source matches '+module.name)
        check(all(len(line)<1024 for line in extracted[module.name].splitlines()),'VBA line-length bound '+module.name)
        # Targeted regression guards, not a VBA parser/compiler. Ignore strings/comments.
        code=re.sub(r'"(?:[^"\n]|"")*"|\x27[^\n]*', '', extracted[module.name])
        declarations=re.findall(r'\b([A-Za-z_]\w*)(?:\(\))?\s+As\s+(?:New\s+)?(\w+)',code,re.I)
        blocked={'input','line','append','output','name'}
        check(not any(name.lower() in blocked for name,_ in declarations),'file-statement keyword identifier guard '+module.name)
        valid_loops=True
        for procedure in re.findall(r'(?ims)^(?:Public|Private) (?:Sub|Function) .*?^End (?:Sub|Function)\s*$',code):
            types={name.lower():kind.lower() for name,kind in re.findall(r'\b([A-Za-z_]\w*)(?:\(\))?\s+As\s+(?:New\s+)?(\w+)',procedure,re.I)}
            for variable in re.findall(r'\bFor Each\s+(\w+)\s+In\b',procedure,re.I):
                # All current loops use Variants for arrays, dictionary keys, or stored rows.
                valid_loops=valid_loops and types.get(variable.lower())=='variant'
        check(valid_loops,'current For Each variable declaration guard '+module.name)

    check(len(extracted)==15,'four standard and eleven empty document modules')
    p.close();w.close()
    print('\n'.join('PASS '+x for x in checks))
    print(f'\n{len(checks)} checks passed. VBA compilation/execution, Office dialogs, rollback, and coauthoring NOT executed.')

if __name__=='__main__':main()
