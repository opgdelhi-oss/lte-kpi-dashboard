import zipfile
import xml.etree.ElementTree as ET
import json
import datetime
import sys

sys.stdout.reconfigure(encoding='utf-8')
file_path = r"C:\Users\OmPrakashGupta\Desktop\Gemini\lte-kpi-dashboard\Final_24 hrs_Hourly 4G KPI Report 24-09-2026.xlsx"
out_compact_js = r"C:\Users\OmPrakashGupta\Desktop\Gemini\lte-kpi-dashboard\userReportDataCompact.js"

print("Starting extraction of user modified report data...")

def excel_date_to_datetime(serial):
    try:
        f = float(serial)
        dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=f)
        return dt
    except Exception:
        return None

with zipfile.ZipFile(file_path, 'r') as z:
    ss = []
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    if 'xl/sharedStrings.xml' in z.namelist():
        print("Loading shared strings...")
        ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in ss_tree.findall('.//main:si', ns):
            ss.append(''.join([t.text or '' for t in si.findall('.//main:t', ns)]))
    
    print("Reading worksheet with iterparse...")
    cols_map = {} # col letter -> stripped header name
    records = []
    
    with z.open('xl/worksheets/sheet1.xml') as f:
        context = ET.iterparse(f, events=('end',))
        for event, elem in context:
            if elem.tag.endswith('row'):
                r_idx = elem.attrib.get('r')
                if r_idx == '1':
                    for c in elem.findall('main:c', ns):
                        ref = c.attrib.get('r')
                        col_letter = ''.join([ch for ch in ref if ch.isalpha()])
                        t = c.attrib.get('t')
                        v = c.find('main:v', ns)
                        val = v.text if v is not None else ''
                        if t == 's' and val.isdigit():
                            val = ss[int(val)]
                        cols_map[col_letter] = val.strip()
                    print(f"Mapped {len(cols_map)} header columns: {list(cols_map.values())[:10]}")
                else:
                    # Data row
                    row_data = {}
                    for c in elem.findall('main:c', ns):
                        ref = c.attrib.get('r')
                        col_letter = ''.join([ch for ch in ref if ch.isalpha()])
                        header = cols_map.get(col_letter)
                        if header:
                            t = c.attrib.get('t')
                            v = c.find('main:v', ns)
                            val = v.text if v is not None else ''
                            if t == 's' and val.isdigit():
                                val = ss[int(val)] if int(val) < len(ss) else val
                            # If header already exists (e.g. duplicate 'Date'), keep non-empty or prefer first
                            if header not in row_data or not row_data[header]:
                                row_data[header] = val
                    
                    if row_data:
                        serial_time = row_data.get('Date', '') or row_data.get('Period start time', '')
                        dt = excel_date_to_datetime(serial_time)
                        date_str = dt.strftime('%Y-%m-%d') if dt else '2026-09-21'
                        hour_val = dt.hour if dt else 0
                        time_str = dt.strftime('%Y-%m-%d %H:%M') if dt else f"{date_str} {hour_val:02d}:00"
                        
                        site_name = row_data.get('MRBTS/SBTS name', '').strip()
                        site_id = row_data.get('MRBTS', '').strip()
                        lncel = row_data.get('LNCEL', '').strip()
                        
                        operator = row_data.get('Operator', '').strip() or 'Unknown'
                        band = row_data.get('Band', '').strip() or 'FDD'

                        def to_float(k, default=0.0):
                            v = row_data.get(k, '')
                            try:
                                return round(float(v), 2)
                            except (ValueError, TypeError):
                                return default

                        rec = {
                            'Date': date_str,
                            'Hour': hour_val,
                            'Timestamp': time_str,
                            'Operator': operator,
                            'Band': band,
                            'Site_Name': site_name,
                            'Site_ID': site_id,
                            'LNCEL': lncel,
                            'Cell_Availability': to_float('Cell Availability LTE-4G_KPI_Report', 100.0),
                            'VoLTE_CSSR': to_float('Nokia_LTE_VoLTE Call Setup Success Rate', 100.0),
                            'E2E_CSSR': to_float('E2E Call Setup Success Rate - All Bearer New', 100.0),
                            'RRC_Setup_SR': to_float('RRC Setup Success Rate', 99.0),
                            'ERAB_Setup_SR': to_float('ERAB Setup Success Rate', 99.0),
                            'Total_Traffic_GB': to_float('NSN_Data Volume - Total_GB', 0.0),
                            'VoLTE_Traffic_Erl': to_float('Nokia_LTE_VoLTE Traffic Erl', 0.0),
                            'VoLTE_Drop_Rate': to_float('NSN_LTE_VoLTE_Drop_Call_Rate', 0.0),
                            'Drop_Rate': to_float('Service DCR Active E-RAB', 0.0),
                            'DL_Packet_Loss': to_float('Nokia_LTE_DL Packet loss', 0.0),
                            'UL_Packet_Loss': to_float('Nokia_LTE_UL Packet loss', 0.0),
                            'DL_PRB_Util': to_float('PDSCH utilization rate PRB DL', 0.0),
                            'UL_PRB_Util': to_float('PDSCH utilization rate PRB UL', 0.0),
                            'DL_User_Throughput_Mbps': to_float('NSN_LTE_DL Throughput Per UE (NA)', 0.0),
                            'UL_User_Throughput_Mbps': to_float('NSN_LTE_UL Throughput Per UE (NA)', 0.0),
                            'CQI': to_float('E-UTRAN Average CQI', 0.0),
                            'Max_Users': int(to_float('Max RRC connected users', 0)),
                            'Active_Users_Avg': int(to_float('Avg RRC conn UE', 0)),
                            'Handover_SR': to_float('Intra Frequency Handover Sucess Rate New', 99.0),
                            'RSSI': to_float('UL PUSCH RSSI', -115.0),
                            'SINR': to_float('SINR', 10.0)
                        }
                        records.append(rec)
                elem.clear()

print(f"Extracted {len(records)} records.")

# Check unique values
operators = sorted(list(set(r['Operator'] for r in records)))
bands = sorted(list(set(r['Band'] for r in records)))
dates = sorted(list(set(r['Date'] for r in records)))
sites = sorted(list(set(r['Site_Name'] for r in records)))
print(f"Operators: {operators}")
print(f"Bands: {bands}")
print(f"Dates: {dates}")
print(f"Total Sites: {len(sites)}")

keys = list(records[0].keys())
compact = {
    'columns': keys,
    'rows': [[r[k] for k in keys] for r in records]
}

with open(out_compact_js, 'w', encoding='utf-8') as f:
    f.write("// Compact User Report Data (Date, Operator, Band, KPIs)\n")
    f.write("window.UserLTEDataCompact = ")
    json.dump(compact, f)
    f.write(";\n")

print(f"Saved to {out_compact_js} successfully!")
