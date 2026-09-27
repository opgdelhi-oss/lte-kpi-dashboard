import zipfile
import xml.etree.ElementTree as ET
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"C:\Users\OmPrakashGupta\Desktop\Final_24 hrs_Hourly 4G KPI Report 21-09-2026.xlsx"
print(f"Inspecting: {file_path}")

with zipfile.ZipFile(file_path, 'r') as z:
    # 1. Read Workbook to get Sheet Names
    wb_xml = z.read('xl/workbook.xml')
    wb_tree = ET.fromstring(wb_xml)
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    
    sheets = []
    for sheet in wb_tree.findall('.//main:sheet', ns):
        name = sheet.attrib.get('name')
        sheet_id = sheet.attrib.get('sheetId')
        r_id = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        sheets.append((name, sheet_id, r_id))
    
    print("Sheets found:")
    for s in sheets:
        print(f"  - {s[0]} (ID: {s[1]}, rId: {s[2]})")

    # 2. Read Shared Strings (if exists)
    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        print("Reading shared strings...")
        ss_xml = z.read('xl/sharedStrings.xml')
        ss_tree = ET.fromstring(ss_xml)
        for si in ss_tree.findall('.//main:si', ns):
            # Text can be in <t> or multiple <r><t>
            text_parts = [t.text or '' for t in si.findall('.//main:t', ns)]
            shared_strings.append(''.join(text_parts))
        print(f"Total shared strings: {len(shared_strings)}")
    
    # 3. Read first sheet rows
    # Look at sheet1.xml or worksheet corresponding to first sheet
    sheet_files = [f for f in z.namelist() if f.startswith('xl/worksheets/sheet') and f.endswith('.xml')]
    print("Worksheet XML files:", sheet_files[:5])

    for sf in sheet_files[:3]:
        print(f"\n--- Sample from {sf} ---")
        ws_xml = z.read(sf)
        ws_tree = ET.fromstring(ws_xml)
        rows = ws_tree.findall('.//main:row', ns)
        print(f"Total rows in {sf}: {len(rows)}")
        
        for r in rows[:10]:
            r_idx = r.attrib.get('r')
            cells = []
            for c in r.findall('main:c', ns):
                cell_ref = c.attrib.get('r')
                cell_type = c.attrib.get('t')
                v = c.find('main:v', ns)
                val = v.text if v is not None else ''
                
                if cell_type == 's' and val.isdigit():
                    val = shared_strings[int(val)] if int(val) < len(shared_strings) else val
                cells.append((cell_ref, val))
            
            row_vals = [f"{c[0]}: {c[1]}" for c in cells[:15]]
            print(f"Row {r_idx}: {row_vals}")
