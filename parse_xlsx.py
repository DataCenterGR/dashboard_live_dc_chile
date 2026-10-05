from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path

path = Path(r'c:\Users\CristianJoseRamosCon\OneDrive - Grenergy Renovables\Escritorio\DC Code\20260901 Analisis estados de proyecto GR (1).xlsx')

with ZipFile(path) as z:
    print('FILES:', z.namelist())
    # workbook names
    workbook = ET.fromstring(z.read('xl/workbook.xml'))
    ns = {
        'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    relmap = {}
    for rel in rels:
        relmap[rel.attrib['Id']] = rel.attrib['Target']
    print('SHEETS:')
    for sheet in workbook.find('a:sheets', ns):
        name = sheet.attrib['name']
        rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        print(name, '->', relmap[rid])
        xml_path = 'xl/' + relmap[rid]
        print('  XML PATH', xml_path)
        # print up to first rows of sheet data
        sheet_xml = ET.fromstring(z.read(xml_path))
        for row in sheet_xml.findall('.//a:sheetData/a:row', ns)[:3]:
            vals = []
            for c in row.findall('a:c', ns):
                vals.append((c.attrib.get('r'), c.attrib.get('t'), c.text))
            print('  ROW', vals)
