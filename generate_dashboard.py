import json
import re
from collections import Counter, OrderedDict
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(r'c:\Users\CristianJoseRamosCon\OneDrive - Grenergy Renovables\Escritorio\DC Code')
XLSX_PATH = ROOT / '20260901 Analisis estados de proyecto GR (1).xlsx'
HTML_PATH = ROOT / 'dashboard.html'
JSON_PATH = ROOT / 'dashboard_data.json'

ns = {
    'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def normalize_text(value):
    if value is None:
        return ''
    value = str(value).strip()
    return value.replace('\u00a0', ' ').replace('\r', '').replace('\n', ' ')


def clean_number(value):
    value = normalize_text(value)
    if not value:
        return 0.0
    value = value.replace(' ', '')
    value = value.replace('%', '')
    value = value.replace('€', '')
    value = value.replace('US$', '')
    value = value.replace('$', '')
    if '(' in value and ')' in value:
        value = '-' + value.replace('(', '').replace(')', '')
    if 'E' in value.upper():
        try:
            return float(value)
        except ValueError:
            pass
    if ',' in value and '.' in value:
        value = value.replace('.', '').replace(',', '.')
    elif ',' in value:
        value = value.replace(',', '.')
    elif value.count('.') > 1:
        value = value.replace('.', '')
    try:
        return float(value)
    except ValueError:
        return 0.0


def read_xlsx(path):
    with ZipFile(path) as z:
        sst = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('a:si', ns):
                texts = [t.text or '' for t in si.findall('.//a:t', ns)]
                sst.append(''.join(texts))

        rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rel_map = {rel.attrib['Id']: rel.attrib['Target'] for rel in rels_root}
        wb_root = ET.fromstring(z.read('xl/workbook.xml'))
        workbook = {}

        for sheet in wb_root.find('a:sheets', ns):
            name = sheet.attrib['name']
            rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            sheet_path = 'xl/' + rel_map[rid]
            sh_root = ET.fromstring(z.read(sheet_path))
            rows = []
            for row in sh_root.findall('.//a:sheetData/a:row', ns):
                cells = {}
                for cell in row.findall('a:c', ns):
                    ref = cell.attrib.get('r', '')
                    col = re.match(r'([A-Z]+)', ref)
                    if not col:
                        continue
                    column = col.group(1)
                    cell_type = cell.attrib.get('t')
                    raw = cell.find('a:v', ns)
                    value = raw.text if raw is not None else ''
                    if cell_type == 's' and value not in ('', None):
                        value = sst[int(value)]
                    elif cell_type == 'inlineStr':
                        value = ''.join(t.text or '' for t in cell.findall('.//a:t', ns))
                    cells[column] = value
                ordered = []
                for column in sorted(cells, key=lambda x: column_index(x)):
                    ordered.append(cells[column])
                rows.append(ordered)
            workbook[name] = rows
        return workbook


def column_index(cell_ref):
    letters = ''.join(ch for ch in cell_ref if ch.isalpha())
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch.upper()) - 64)
    return value


def rows_to_dict(rows, start_index=0):
    if not rows or start_index >= len(rows):
        return []
    header = [normalize_text(v) for v in rows[start_index]]
    data = []
    for row in rows[start_index + 1:]:
        record = {}
        for idx, header_name in enumerate(header):
            if idx < len(row):
                record[header_name] = normalize_text(row[idx])
        if any(record.values()):
            data.append(record)
    return data


def records_from_sheet(rows, required_fields):
    for idx, row in enumerate(rows):
        normalized = [normalize_text(v) for v in row]
        if all(field in normalized for field in required_fields):
            return rows_to_dict(rows, idx)
    return rows_to_dict(rows)


def summarise_solicitudes(records):
    counts = Counter()
    region_counts = Counter()
    type_counts = Counter()
    cap_total = 0.0
    top_projects = []

    for r in records:
        status = r.get('Estado Solicitud', 'Sin dato') or 'Sin dato'
        region = r.get('Región', 'Sin dato') or 'Sin dato'
        project_type = r.get('Tipo Proyecto', 'Sin dato') or 'Sin dato'
        capacity = clean_number(r.get('Capacidad [MW]', '0'))
        project = r.get('Proyecto', 'Sin dato') or 'Sin dato'

        counts[status] += 1
        region_counts[region] += 1
        type_counts[project_type] += 1
        cap_total += capacity

        top_projects.append({'project': project, 'capacity': capacity, 'region': region, 'status': status})

    top_projects = sorted(top_projects, key=lambda x: x['capacity'], reverse=True)[:10]
    return {
        'total_projects': len(records),
        'total_capacity_mw': round(cap_total, 1),
        'status_counts': OrderedDict(sorted(counts.items(), key=lambda x: (-x[1], x[0]))),
        'region_counts': OrderedDict(sorted(region_counts.items(), key=lambda x: (-x[1], x[0]))),
        'type_counts': OrderedDict(sorted(type_counts.items(), key=lambda x: (-x[1], x[0]))),
        'top_projects': top_projects,
    }


def summarise_dc_portfolio(records):
    if not records:
        return {'total_projects': 0, 'total_investment': 0, 'region_counts': {}, 'state_counts': {}}
    count_total = 0
    investment_total = 0.0
    region_counts = Counter()
    state_counts = Counter()

    for r in records:
        project = normalize_text(r.get('Nombre del Proyecto', '') or r.get('Proyecto', ''))
        if not project or not re.search(r'[A-Za-zÁÉÍÓÚáéíóúñÑüÜ]', project):
            continue

        region = normalize_text(r.get('Región', '') or r.get('Region', ''))
        state = normalize_text(r.get('Estado del Proyecto', '') or r.get('Estado', ''))
        if not region or not re.search(r'[A-Za-zÁÉÍÓÚáéíóúñÑüÜ]', region):
            continue
        if not state or not re.search(r'[A-Za-zÁÉÍÓÚáéíóúñÑüÜ]', state):
            continue

        count_total += 1
        investment = clean_number(r.get('Inversión (MMU$)', r.get('CAPEX (MMUSD)', '0')))
        region_counts[region] += 1
        state_counts[state] += 1
        investment_total += investment

    return {
        'total_projects': count_total,
        'total_investment': round(investment_total, 1),
        'region_counts': OrderedDict(sorted(region_counts.items(), key=lambda x: (-x[1], x[0]))),
        'state_counts': OrderedDict(sorted(state_counts.items(), key=lambda x: (-x[1], x[0]))),
    }


def build_dashboard_data():
    workbook = read_xlsx(XLSX_PATH)
    solicitudes_rows = workbook.get('solicitudes', [])
    dc_rows = workbook.get('DC SEA  2026', [])

    solicitudes_records = records_from_sheet(solicitudes_rows, ['Proyecto', 'Estado Solicitud'])
    dc_records = records_from_sheet(dc_rows, ['Nombre del Proyecto', 'Estado del Proyecto'])

    solicitudes_summary = summarise_solicitudes(solicitudes_records)
    dc_summary = summarise_dc_portfolio(dc_records)

    # Extra insight from the 2026 access sheet.
    acceso_rows = workbook.get('DC Acceso abierto 2026', [])
    acceso_records = records_from_sheet(acceso_rows, ['Proyecto', 'Capacidad [MW]'])
    acceso_total = len(acceso_records)
    acceso_capacity = round(sum(clean_number(r.get('Capacidad [MW]', '0')) for r in acceso_records), 1)

    data = {
        'title': 'Dashboard proyecto GR',
        'generated_from': str(XLSX_PATH.name),
        'solicitudes': solicitudes_summary,
        'dc': dc_summary,
        'acceso_2026': {
            'total_projects': acceso_total,
            'total_capacity_mw': acceso_capacity,
        },
    }
    return data


def render_html(data):
    status_items = ', '.join(f"{k}: {v}" for k, v in data['solicitudes']['status_counts'].items())
    region_items = ', '.join(f"{k}: {v}" for k, v in data['solicitudes']['region_counts'].items())
    type_items = ', '.join(f"{k}: {v}" for k, v in data['solicitudes']['type_counts'].items())
    dc_states = ', '.join(f"{k}: {v}" for k, v in data['dc']['state_counts'].items())
    top_projects = data['solicitudes']['top_projects']
    top_projects_html = ''.join(
        f"<tr><td>{p['project']}</td><td>{p['capacity']}</td><td>{p['region']}</td><td>{p['status']}</td></tr>"
        for p in top_projects
    )

    html = f"""<!DOCTYPE html>
<html lang='es'>
<head>
  <meta charset='UTF-8' />
  <meta name='viewport' content='width=device-width, initial-scale=1.0' />
  <title>{data['title']}</title>
  <script src='https://cdn.jsdelivr.net/npm/chart.js'></script>
  <style>
    :root {{
      --bg: #0b1220;
      --panel: #121b2b;
      --panel-soft: #182538;
      --text: #e9eef9;
      --muted: #a7b7d4;
      --primary: #31c48d;
      --secondary: #3b82f6;
      --accent: #f59e0b;
      --danger: #ef4444;
      --shadow: rgba(15, 23, 42, 0.35);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0; font-family: Arial, sans-serif; background: linear-gradient(135deg, #0b1220, #101a2d 45%, #0d1728);
      color: var(--text);
    }}
    .container {{ max-width: 1400px; margin: 0 auto; padding: 32px 20px 48px; }}
    h1 {{ margin: 0 0 12px; font-size: 2.2rem; }}
    .subtitle {{ color: var(--muted); margin-bottom: 24px; }}
    .kpis {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 24px; }}
    .kpi {{ background: rgba(18, 27, 43, 0.9); border: 1px solid rgba(148, 163, 184, 0.18); border-radius: 16px; padding: 18px 20px; box-shadow: 0 10px 25px var(--shadow); }}
    .kpi-label {{ color: var(--muted); font-size: 0.8rem; letter-spacing: 0.08em; text-transform: uppercase; }}
    .kpi-value {{ margin-top: 12px; font-size: 2rem; font-weight: 700; }}
    .kpi-note {{ margin-top: 8px; color: var(--muted); font-size: 0.85rem; }}
    .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 22px; margin-top: 18px; }}
    .panel {{ background: rgba(18, 27, 43, 0.9); border: 1px solid rgba(148, 163, 184, 0.18); border-radius: 16px; padding: 18px 20px; box-shadow: 0 10px 25px var(--shadow); }}
    .panel h3 {{ margin-top: 0; }}
    canvas {{ max-height: 320px; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 8px; }}
    th, td {{ padding: 10px 12px; border-bottom: 1px solid rgba(148, 163, 184, 0.20); text-align: left; font-size: 0.95rem; }}
    th {{ color: var(--muted); font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; font-size: 0.76rem; }}
    .small {{ font-size: 0.82rem; color: var(--muted); }}
  </style>
</head>
<body>
  <div class='container'>
    <h1>{data['title']}</h1>
    <div class='subtitle'>Fuente: {data['generated_from']}</div>

    <div class='kpis'>
      <div class='kpi'>
        <div class='kpi-label'>Solicitudes totales</div>
        <div class='kpi-value'>{data['solicitudes']['total_projects']}</div>
        <div class='kpi-note'>Proyectos analizados en la hoja de solicitudes</div>
      </div>
      <div class='kpi'>
        <div class='kpi-label'>Capacidad acumulada</div>
        <div class='kpi-value'>{data['solicitudes']['total_capacity_mw']} MW</div>
        <div class='kpi-note'>Capacidad total solicitada</div>
      </div>
      <div class='kpi'>
        <div class='kpi-label'>Acceso abierto 2026</div>
        <div class='kpi-value'>{data['acceso_2026']['total_projects']}</div>
        <div class='kpi-note'>{data['acceso_2026']['total_capacity_mw']} MW acumulados</div>
      </div>
      <div class='kpi'>
        <div class='kpi-label'>Portfolio DC 2026</div>
        <div class='kpi-value'>{data['dc']['total_projects']}</div>
        <div class='kpi-note'>{data['dc']['total_investment']} MMU$ inversión</div>
      </div>
    </div>

    <div class='grid'>
      <div class='panel'>
        <h3>Estado de solicitud</h3>
        <canvas id='statusChart'></canvas>
      </div>
      <div class='panel'>
        <h3>Distribución por región</h3>
        <canvas id='regionChart'></canvas>
      </div>
      <div class='panel'>
        <h3>Tipo de proyecto</h3>
        <canvas id='typeChart'></canvas>
      </div>
      <div class='panel'>
        <h3>Top 10 proyectos por capacidad</h3>
        <table>
          <thead>
            <tr><th>Proyecto</th><th>MW</th><th>Región</th><th>Estado</th></tr>
          </thead>
          <tbody>
            {top_projects_html}
          </tbody>
        </table>
      </div>
      <div class='panel'>
        <h3>Portfolio DC - estados</h3>
        <canvas id='dcStateChart'></canvas>
      </div>
      <div class='panel'>
        <h3>Portfolio DC - regiones</h3>
        <table>
          <thead>
            <tr><th>Región</th><th>Proyectos</th></tr>
          </thead>
          <tbody>
            {''.join(f'<tr><td>{k}</td><td>{v}</td></tr>' for k, v in data['dc']['region_counts'].items())}
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <script>
    const statusData = {json.dumps(dict(data['solicitudes']['status_counts']))};
    const regionData = {json.dumps(dict(data['solicitudes']['region_counts']))};
    const typeData = {json.dumps(dict(data['solicitudes']['type_counts']))};
    const dcStateData = {json.dumps(dict(data['dc']['state_counts']))};

    new Chart(document.getElementById('statusChart'), {{
      type: 'doughnut',
      data: {{ labels: Object.keys(statusData), datasets: [{{ data: Object.values(statusData), backgroundColor: ['#3b82f6','#31c48d','#f59e0b','#ef4444','#8b5cf6','#14b8a6'] }}] }},
      options: {{ plugins: {{ legend: {{ position: 'bottom' }} }} }}
    }});

    new Chart(document.getElementById('regionChart'), {{
      type: 'bar',
      data: {{ labels: Object.keys(regionData), datasets: [{{ label: 'Proyectos', data: Object.values(regionData), backgroundColor: '#3b82f6' }}] }},
      options: {{ indexAxis: 'y', plugins: {{ legend: {{ display: false }} }}, scales: {{ x: {{ beginAtZero: true }} }} }}
    }});

    new Chart(document.getElementById('typeChart'), {{
      type: 'pie',
      data: {{ labels: Object.keys(typeData), datasets: [{{ data: Object.values(typeData), backgroundColor: ['#31c48d','#f59e0b','#3b82f6','#ef4444','#8b5cf6','#14b8a6'] }}] }},
      options: {{ plugins: {{ legend: {{ position: 'bottom' }} }} }}
    }});

    new Chart(document.getElementById('dcStateChart'), {{
      type: 'polarArea',
      data: {{ labels: Object.keys(dcStateData), datasets: [{{ data: Object.values(dcStateData), backgroundColor: ['#31c48d','#f59e0b','#3b82f6','#ef4444','#8b5cf6'] }}] }},
      options: {{ plugins: {{ legend: {{ position: 'bottom' }} }} }}
    }});
  </script>
</body>
</html>
"""
    return html


def main():
    data = build_dashboard_data()
    JSON_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    HTML_PATH.write_text(render_html(data), encoding='utf-8')
    print(f'Archivo generado: {HTML_PATH}')
    print(f'JSON generado: {JSON_PATH}')
    print(json.dumps({
        'total_solicitudes': data['solicitudes']['total_projects'],
        'total_capacity_mw': data['solicitudes']['total_capacity_mw'],
        'dc_total': data['dc']['total_projects'],
        'dc_investment': data['dc']['total_investment'],
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
