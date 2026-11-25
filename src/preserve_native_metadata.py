"""Current native-metadata preservation pass.
Adapted 2026-11-20 from the archived helper to accept both prefixed and
unprefixed worksheet XML. All cell authoring remains in artifact-tool.
Usage: python preserve_native_metadata.py BASELINE_XLSX REBUILT_XLSX
"""
import io
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

if len(sys.argv) != 3:
    raise SystemExit("Usage: preserve_native_metadata.py BASELINE_XLSX REBUILT_XLSX")
original, delivered = map(Path, sys.argv[1:3])
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
ET.register_namespace('c', ns['c'])
ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
ET.register_namespace('r', ns['r'])

def sheet_paths(z):
    rels = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    result = {}
    for s in ET.fromstring(z.read('xl/workbook.xml')).find('m:sheets', ns):
        target = rels[s.attrib['{' + ns['r'] + '}id']]
        result[s.attrib['name']] = target.lstrip('/') if target.startswith('/') else 'xl/' + target
    return result

with zipfile.ZipFile(original) as src, zipfile.ZipFile(delivered) as dst:
    original_paths, output_paths = sheet_paths(src), sheet_paths(dst)
    contents = {name: dst.read(name) for name in dst.namelist()}
    restored = []
    # Preserve the source's A4 portrait settings for the four untouched sheets.
    for name in ['一年级', '二年级', '三年级', '五年级']:
        source_xml = src.read(original_paths[name]).decode('utf-8')
        output_xml = contents[output_paths[name]].decode('utf-8')
        setup = re.search(r'<(?:\w+:)?pageSetup\b[^>]*(?:/>|>.*?</(?:\w+:)?pageSetup>)', source_xml)
        if setup:
            prefix = re.search(r'<(\w+:)?worksheet\b', output_xml).group(1) or ''
            output_xml = re.sub(r'<(?:\w+:)?pageSetup\b[^>]*(?:/>|>.*?</(?:\w+:)?pageSetup>)', '', output_xml)
            margin = re.search(r'<(?:\w+:)?pageMargins\b[^>]*(?:/>|>.*?</(?:\w+:)?pageMargins>)', output_xml)
            pos = margin.end() if margin else output_xml.index('</' + prefix + 'worksheet>')
            fragment = re.sub(r'<(/?)(?:\w+:)?pageSetup', lambda m: '<' + m[1] + prefix + 'pageSetup', setup.group(0))
            output_xml = output_xml[:pos] + fragment + output_xml[pos:]
            contents[output_paths[name]] = output_xml.encode('utf-8')
            restored.append(name)
    # Restore charsets by the actual unchanged cells, since the same font name
    # can legitimately have different source charset values.
    sstyles = ET.fromstring(src.read('xl/styles.xml'))
    dstyles = ET.fromstring(contents['xl/styles.xml'])
    sfonts, dfonts = sstyles.find('m:fonts', ns), dstyles.find('m:fonts', ns)
    sxfs, dxfs = sstyles.find('m:cellXfs', ns), dstyles.find('m:cellXfs', ns)
    desired = {}
    for name in ['一年级', '二年级', '三年级', '五年级']:
        old_cells = ET.fromstring(src.read(original_paths[name])).findall('.//m:sheetData/m:row/m:c', ns)
        new_cells = {c.attrib['r']: c for c in ET.fromstring(contents[output_paths[name]]).findall('.//m:sheetData/m:row/m:c', ns)}
        for old_cell in old_cells:
            new_cell = new_cells.get(old_cell.attrib['r'])
            if new_cell is None:
                continue
            old_font = sfonts[int(sxfs[int(old_cell.attrib.get('s', 0))].attrib.get('fontId', 0))]
            new_font_id = int(dxfs[int(new_cell.attrib.get('s', 0))].attrib.get('fontId', 0))
            charset = old_font.find('m:charset', ns)
            if charset is not None:
                desired.setdefault(new_font_id, set()).add(charset.attrib['val'])
    for font_id, values in desired.items():
        if len(values) != 1:
            raise ValueError('Conflicting unchanged source charsets')
        target = dfonts[font_id]
        current = target.find('m:charset', ns)
        if current is None:
            current = ET.SubElement(target, '{' + ns['m'] + '}charset')
        current.attrib['val'] = next(iter(values))
    ET.register_namespace('x', ns['m'])
    text = contents['xl/styles.xml'].decode('utf-8')
    text = re.sub(r'<(?:\w+:)?fonts\b[^>]*>.*?</(?:\w+:)?fonts>', lambda m: ET.tostring(dfonts, encoding='unicode'), text, count=1)
    contents['xl/styles.xml'] = text.encode('utf-8')
    strings = []
    if 'xl/sharedStrings.xml' in contents:
        strings = [''.join(si.itertext()) for si in ET.fromstring(contents['xl/sharedStrings.xml']).findall('m:si', ns)]
    sheets = {}
    for name, path in output_paths.items():
        values = {}
        for cell in ET.fromstring(contents[path]).findall('.//m:sheetData/m:row/m:c', ns):
            value = cell.find('m:v', ns)
            if value is None or value.text is None:
                value = ''.join(cell.find('m:is', ns).itertext()) if cell.attrib.get('t') == 'inlineStr' else None
            elif cell.attrib.get('t') == 's':
                value = strings[int(value.text)]
            elif cell.attrib.get('t') in ('str', 'e'):
                value = value.text
            else:
                value = float(value.text)
            values[cell.attrib['r']] = value
        sheets[name] = values
    def ref_values(formula):
        sheet, address = formula.rsplit('!', 1)
        sheet = sheet.strip("'").replace("''", "'")
        address = address.replace('$', '')
        first, last = (address.split(':') if ':' in address else [address, address])
        a, b = re.fullmatch(r'([A-Z]+)(\d+)', first), re.fullmatch(r'([A-Z]+)(\d+)', last)
        if not a or not b or a[1] != b[1]:
            raise ValueError('Unsupported chart reference ' + formula)
        return [sheets[sheet].get(a[1] + str(r)) for r in range(int(a[2]), int(b[2]) + 1)]
    chart_report = []
    for path in list(contents):
        if not re.fullmatch(r'xl/(?:drawings/)?charts/chart\d+\.xml', path):
            continue
        tree = ET.fromstring(contents[path])
        caches = []
        for kind in ['strRef', 'numRef']:
            for node in tree.findall('.//c:' + kind, ns):
                f = node.find('c:f', ns)
                if f is None or not f.text:
                    continue
                values = ref_values(f.text)
                tag = 'strCache' if kind == 'strRef' else 'numCache'
                old = node.find('c:' + tag, ns)
                if old is not None:
                    node.remove(old)
                cache = ET.SubElement(node, '{' + ns['c'] + '}' + tag)
                if kind == 'numRef':
                    ET.SubElement(cache, '{' + ns['c'] + '}formatCode').text = 'General'
                ET.SubElement(cache, '{' + ns['c'] + '}ptCount', {'val': str(len(values))})
                for index, value in enumerate(values):
                    if value is None or value == '':
                        continue
                    if kind == 'numRef' and not isinstance(value, (int, float)):
                        raise ValueError('Non-numeric chart value ' + repr(value))
                    point = ET.SubElement(cache, '{' + ns['c'] + '}pt', {'idx': str(index)})
                    ET.SubElement(point, '{' + ns['c'] + '}v').text = str(value)
                caches.append({'formula': f.text, 'count': len(values)})
        contents[path] = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
        chart_report.append({'path': path, 'caches': caches})
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as final:
        for name, value in contents.items():
            final.writestr(name, value)
    delivered.write_bytes(buffer.getvalue())
print(json.dumps({'print_settings_restored': restored, 'chart_caches': chart_report}, ensure_ascii=False))
