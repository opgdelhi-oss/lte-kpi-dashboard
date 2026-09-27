import zipfile
import xml.etree.ElementTree as ET
import sys

sys.stdout.reconfigure(encoding='utf-8')
file_path = r"C:\Users\OmPrakashGupta\Desktop\Final_24 hrs_Hourly 4G KPI Report 21-09-2026.xlsx"

with zipfile.ZipFile(file_path, 'r') as z:
    ss = []
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    if 'xl/sharedStrings.xml' in z.namelist():
        ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in ss_tree.findall('.//main:si', ns):
            ss.append(''.join([t.text or '' for t in si.findall('.//main:t', ns)]))
    ws_tree = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    for r in ws_tree.findall('.//main:row', ns)[:8]:
        row_idx = r.attrib.get('r')
        cells = {}
        for c in r.findall('main:c', ns):
            ref = c.attrib.get('r')
            letter = ''.join([ch for ch in ref if ch.isalpha()])
            t = c.attrib.get('t')
            v = c.find('main:v', ns)
            val = v.text if v is not None else ''
            if t == 's' and val.isdigit():
                val = ss[int(val)] if int(val) < len(ss) else val
            cells[letter] = val
        print(f"Row {row_idx}: A='{cells.get('A')}', B='{cells.get('B')}', C='{cells.get('C')}', D='{cells.get('D')}', E='{cells.get('E')}', F='{cells.get('F')}', G='{cells.get('G')}'")
