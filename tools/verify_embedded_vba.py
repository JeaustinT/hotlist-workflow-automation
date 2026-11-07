"""Verify the final XLSM bytes, never just build inputs. Does not run VBA."""
from collections import Counter
from pathlib import Path
import argparse
import hashlib
import io
import json
import re
import zipfile
import xml.etree.ElementTree as ET
import olefile
from oletools.olevba import VBA_Parser, decompress_stream

EXPECTED = ('DemoUtilities', 'HistoryManagement', 'HotlistProcessing', 'OfficeIntegration')
REQUIRED_IDENTIFIERS = {
    'DemoUtilities': ('appendRows', 'sheetName'),
    'HistoryManagement': (),
    'HotlistProcessing': ('inputRows', 'lineId'),
    'OfficeIntegration': ('exportWorkbook',),
}


class EmbeddedSourceError(ValueError):
    pass


def normalized(source):
    # Keep every character, attribute, and space except platform line endings/BOM.
    return source.removeprefix('\ufeff').replace('\r\n', '\n').replace('\r', '\n')


def digest(data):
    return hashlib.sha256(data).hexdigest()


def require(condition, message):
    if not condition:
        raise EmbeddedSourceError(message)


def verify(workbook, source_dir, expected_version=None):
    workbook, source_dir = Path(workbook), Path(source_dir)
    final_bytes = workbook.read_bytes()  # Reopen actual final artifact on every call.
    with zipfile.ZipFile(io.BytesIO(final_bytes)) as package:
        counts = Counter(package.namelist())
        require(not any(n > 1 for n in counts.values()), 'Duplicate ZIP parts in final workbook')
        require(package.testzip() is None, 'Invalid workbook ZIP CRC')
        links = ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))
        projects = [r for r in links if r.attrib.get('Type', '').endswith('/vbaProject')]
        require(len(projects) == 1, 'Expected exactly one workbook VBA project relationship')
        link = projects[0]
        require(link.attrib.get('TargetMode') != 'External', 'External VBA project is forbidden')
        require(link.attrib['Target'] in ('vbaProject.bin', '/xl/vbaProject.bin'), 'Unexpected VBA project target')
        binary = package.read('xl/vbaProject.bin')
        namespace = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        control = ET.fromstring(package.read('xl/worksheets/sheet1.xml'))
        cell = control.find(".//s:c[@r='K4']", namespace)
        require(cell is not None, 'Missing Control version cell')
        value = cell.find('s:v', namespace)
        require(value is not None, 'Missing Control version value')
        if cell.attrib.get('t') == 's':
            strings = ET.fromstring(package.read('xl/sharedStrings.xml'))
            version = ''.join(t.text or '' for t in list(strings)[int(value.text)].iter() if t.tag.endswith('}t'))
        else:
            version = value.text
    if expected_version is not None:
        require(version == expected_version, f'Workbook version {version!r} != {expected_version!r}')

    # oletools follows the actual project directory to extract/decompress module source.
    parser = VBA_Parser(str(workbook), data=final_bytes)
    try:
        extracted = list(parser.extract_macros())
    finally:
        parser.close()
    standard = {}
    for _, stream, filename, source in extracted:
        if filename.lower().endswith('.bas'):
            name = Path(filename).stem
            require(name not in standard, f'Duplicate standard module: {name}')
            standard[name] = normalized(source)
    require(set(standard) == set(EXPECTED), f'Unexpected embedded standard modules: {sorted(standard)}')
    require({p.stem for p in source_dir.glob('*.bas')} == set(EXPECTED), 'Unexpected repository standard module set')

    # Independent direct stream decompression for this clean, source-offset-zero format.
    project = olefile.OleFileIO(io.BytesIO(binary))
    try:
        require(project.openstream('VBA/_VBA_PROJECT').read() == bytes.fromhex('cc61ffff000000'), 'Unexpected compiled-cache project header')
        require(not any('__SRP' in '/'.join(p) for p in project.listdir()), 'Compiled SRP caches are forbidden')
        declaration = project.openstream('PROJECT').read().decode('cp1252')
        listed = re.findall(r'^Module=(.+)\r?$', declaration, re.M)
        listed = [n.strip() for n in listed]
        require(len(listed) == len(EXPECTED) and set(listed) == set(EXPECTED), 'PROJECT module list mismatch')
        report = {}
        for name in EXPECTED:
            embedded = standard[name]
            direct = normalized(decompress_stream(project.openstream('VBA/' + name).read()).decode('cp1252'))
            require(direct == embedded, f'{name}: extractor/direct-stream disagreement')
            code = re.sub(r'"(?:[^"\n]|"")*"|\x27[^\n]*', '', embedded)
            declarations = re.findall(r'\b([A-Za-z_]\w*)(?:\(\))?\s+As\s+(?:New\s+)?(\w+)', code, re.I)
            blocked = {'input', 'line', 'append', 'output', 'name'}
            require(not any(n.lower() in blocked for n, _ in declarations), f'{name}: stale reserved identifier declaration')
            require(not re.search(r'\binput\s*\.\s*Add\b', code, re.I), f'{name}: stale input.Add')
            require(not re.search(r'\bDim\s+line\s+As\s+String\b', code, re.I), f'{name}: stale Dim line As String')
            if name == 'HotlistProcessing':
                require(bool(re.search(r'\bDim\s+inputRows\s+As\s+New\s+Collection\b', code, re.I)), 'Missing corrected inputRows declaration')
            for identifier in REQUIRED_IDENTIFIERS[name]:
                require(bool(re.search(r'\b' + identifier + r'\b', code)), f'{name}: missing corrected identifier {identifier}')
            expected = normalized((source_dir / (name + '.bas')).read_text(encoding='utf-8'))
            require(embedded == expected, f'{name}: final-binary source mismatch (embedded {digest(embedded.encode())}, repository {digest(expected.encode())})')
            report[name] = {'match': True, 'normalized_source_sha256': digest(expected.encode()), 'normalized_embedded_sha256': digest(embedded.encode()), 'direct_decompression_match': True}
    finally:
        project.close()
    # Static source equality never proves a desktop compile. Accept external evidence
    # only when it is bound to these exact artifact bytes and this demo version.
    receipt = {'demo_version': version, 'workbook_sha256': digest(final_bytes), 'vba_project_sha256': digest(binary), 'modules': report, 'stale_identifier_checks': 'PASS', 'excel_desktop_compile': 'NOT VERIFIED FOR THIS ARTIFACT'}
    evidence_path = source_dir.resolve().parents[1] / 'validation/desktop-acceptance.json'
    if evidence_path.exists():
        evidence = json.loads(evidence_path.read_text(encoding='utf-8'))
        if evidence.get('workbook_sha256') == receipt['workbook_sha256'] and evidence.get('demo_version') == version and evidence.get('excel_desktop_compile') == 'PASS':
            receipt['excel_desktop_compile'] = 'PASS'
            receipt['desktop_acceptance_evidence'] = 'validation/desktop-acceptance.json'
    return receipt


if __name__ == '__main__':
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument('workbook', type=Path)
    cli.add_argument('--source-dir', type=Path, default=Path(__file__).resolve().parents[1] / 'src/vba')
    cli.add_argument('--version')
    args = cli.parse_args()
    print(json.dumps(verify(args.workbook, args.source_dir, args.version), indent=2))
