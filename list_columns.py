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
    row1 = ws_tree.find('.//main:row[@r="1"]', ns)
    cols = []
    for c in row1.findall('main:c', ns):
        cell_ref = c.attrib.get('r')
        cell_type = c.attrib.get('t')
        v = c.find('main:v', ns)
        val = v.text if v is not None else ''
        if cell_type == 's' and val.isdigit():
            val = ss[int(val)]
        cols.append((cell_ref, val))
    
    print(f"Total columns: {len(cols)}")
    for idx, (col_letter, name) in enumerate(cols):
        print(f"{idx+1}. [{col_letter}] {name}")

    # Also check what dates/hours are in column A
    rows = ws_tree.findall('.//main:row', ns)
    print(f"\nTotal rows in sheet: {len(rows)}")
    sample_dates = []
    for r in rows[1:100:10]:
        c_first = r.find('main:c', ns)
        v = c_first.find('main:v', ns)
        if v is not None:
            sample_dates.append(v.text)
    print("Sample values in Col A (Period start time):", sample_dates)
