from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path

path = Path(r'c:\Users\CristianJoseRamosCon\OneDrive - Grenergy Renovables\Escritorio\DC Code\20260901 Analisis estados de proyecto GR (1).xlsx')
ns = {
    'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

with ZipFile(path) as z:
    sst = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall('a:si', ns):
            texts = []
            for t in si.findall('.//a:t', ns):
                texts.append(t.text or '')
            sst.append(''.join(texts))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rel_map = {r.attrib['Id']: r.attrib['Target'] for r in rels}
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    for sheet in wb.find('a:sheets', ns):
        name = sheet.attrib['name']
        rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        sheet_path = 'xl/' + rel_map[rid]
        sh = ET.fromstring(z.read(sheet_path))
        print('\n###', name)
        rows = []
        for row in sh.findall('.//a:sheetData/a:row', ns):
            vals = []
            for c in row.findall('a:c', ns):
                ref = c.attrib.get('r')
                cell_type = c.attrib.get('t')
                v = c.find('a:v', ns)
                val = v.text if v is not None else ''
                if cell_type == 's' and val != '':
                    val = sst[int(val)]
                elif cell_type == 'inlineStr':
                    val = ''.join(t.text or '' for t in c.findall('.//a:t', ns))
                vals.append((ref, val))
            rows.append(vals)
        for row in rows[:12]:
            print(row)
