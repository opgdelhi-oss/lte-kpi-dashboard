import zipfile
import xml.etree.ElementTree as ET
import sys

sys.stdout.reconfigure(encoding='utf-8')
file_path = r"C:\Users\OmPrakashGupta\Desktop\Final_24 hrs_Hourly 4G KPI Report 21-09-2026.xlsx"

operators = set()
bands = set()
dates = set()

with zipfile.ZipFile(file_path, 'r') as z:
    ss = []
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    if 'xl/sharedStrings.xml' in z.namelist():
        ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in ss_tree.findall('.//main:si', ns):
            ss.append(''.join([t.text or '' for t in si.findall('.//main:t', ns)]))
    
    with z.open('xl/worksheets/sheet1.xml') as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag.endswith('row'):
                r_idx = elem.attrib.get('r')
                if r_idx != '1':
                    cells = {}
                    for c in elem.findall('main:c', ns):
                        ref = c.attrib.get('r')
                        letter = ''.join([ch for ch in ref if ch.isalpha()])
                        t = c.attrib.get('t')
                        v = c.find('main:v', ns)
                        val = v.text if v is not None else ''
                        if t == 's' and val.isdigit():
                            val = ss[int(val)] if int(val) < len(ss) else val
                        cells[letter] = val
                    
                    if 'G' in cells:
                        operators.add(cells['G'])
                    if 'F' in cells:
                        bands.add(cells['F'])
                    if 'A' in cells:
                        dates.add(cells['A'])
                elem.clear()

print("Unique Operators in Excel:", operators)
print("Unique Bands in Excel:", bands)
print("Sample Dates count:", len(dates))
