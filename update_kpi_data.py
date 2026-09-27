"""Append previously missing report dates from OneDrive KPI workbooks."""

import argparse
import datetime as dt
import json
import os
import re
import sys
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

SOURCE_DIR = Path(r"C:\OM Prakash Gupta Data\OneDrive - cloudextel.com\CloudExtel RAN KPI")
OUTPUT_DIR = Path(__file__).resolve().parent
HOURLY_OUTPUT = OUTPUT_DIR / "userReportDataCompact.js"
DAYWISE_OUTPUT = OUTPUT_DIR / "daywiseData.js"
METADATA_OUTPUT = OUTPUT_DIR / "kpiDataSources.json"
METADATA_JS_OUTPUT = OUTPUT_DIR / "kpiDataSources.js"
EXCEL_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
COMPACT_COLUMNS = [
    "Date", "Hour", "Timestamp", "Operator", "Band", "Site_Name", "Site_ID", "LNCEL",
    "Cell_Availability", "VoLTE_CSSR", "E2E_CSSR", "RRC_Setup_SR", "ERAB_Setup_SR",
    "Total_Traffic_GB", "VoLTE_Traffic_Erl", "VoLTE_Drop_Rate", "Drop_Rate",
    "DL_Packet_Loss", "UL_Packet_Loss", "DL_PRB_Util", "UL_PRB_Util",
    "DL_User_Throughput_Mbps", "UL_User_Throughput_Mbps", "CQI", "Max_Users",
    "Active_Users_Avg", "Handover_SR", "RSSI", "SINR",
]
LNCEL_BAND_MAP = {
    **{str(lncel): "FDD" for lncel in (1, 2, 3, 4, 5, 6, 7, 8, 21, 22, 23, 24, 25, 26, 27, 28, 29)},
    **{str(lncel): "TDD" for lncel in (*range(51, 62), *range(101, 120))},
}


def log(message):
    print(f"[KPI import] {message}", flush=True)


def excel_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time())
    try:
        return dt.datetime(1899, 12, 30) + dt.timedelta(days=float(value))
    except (TypeError, ValueError, OverflowError):
        text = str(value).strip()
        for date_format in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y"):
            try:
                return dt.datetime.strptime(text, date_format)
            except ValueError:
                pass
    return None


def to_float(value, default=0.0):
    try:
        if value is None or value == "":
            return default
        number = float(value)
        return number if number == number and abs(number) != float("inf") else default
    except (TypeError, ValueError):
        return default


def normalize_record(row, daily=False):
    date_value = row.get("Date") or row.get("Period start time")
    timestamp = excel_date(date_value)
    if timestamp is None:
        return None
    date_string = timestamp.strftime("%Y-%m-%d")
    hour = 24 if daily else timestamp.hour

    def value(*names, default=0.0):
        for name in names:
            if name in row and row[name] not in (None, ""):
                return to_float(row[name], default)
        return default

    site_name = str(row.get("MRBTS/SBTS name") or row.get("Site_Name") or "").strip()
    site_id = str(row.get("MRBTS") or row.get("Site_ID") or "").strip()
    lncel = str(row.get("LNCEL") or "").strip()
    normalized_lncel = re.sub(r"\.0+$", "", lncel)
    operator = str(row.get("Operator") or "Unknown").strip()
    band = LNCEL_BAND_MAP.get(normalized_lncel) or str(row.get("Band") or "Unknown").strip()

    record = {
        "Date": date_string,
        "Hour": hour,
        "Timestamp": date_string if daily else timestamp.strftime("%Y-%m-%d %H:%M"),
        "Operator": operator,
        "Band": band,
        "Site_Name": site_name,
        "Site_ID": site_id,
        "LNCEL": lncel,
        "Cell_Availability": value("Cell Availability LTE-4G_KPI_Report", "Cell_Availability", default=100.0),
        "VoLTE_CSSR": value("Nokia_LTE_VoLTE Call Setup Success Rate", "VoLTE_CSSR", default=100.0),
        "E2E_CSSR": value("E2E Call Setup Success Rate - All ******", "E2E_CSSR", default=100.0),
        "RRC_Setup_SR": value("RRC Setup Success Rate", "RRC_Setup_SR", default=99.0),
        "ERAB_Setup_SR": value("ERAB Setup Success Rate", "ERAB_Setup_SR", default=99.0),
        "Total_Traffic_GB": value("NSN_Data Volume - Total_GB", "Total_Traffic_GB"),
        "VoLTE_Traffic_Erl": value("Nokia_LTE_VoLTE Traffic Erl", "VoLTE_Traffic_Erl"),
        "VoLTE_Drop_Rate": value("NSN_LTE_VoLTE_Drop_Call_Rate", "VoLTE_Drop_Rate"),
        "Drop_Rate": value("Service DCR Active E-RAB", "Drop_Rate"),
        "DL_Packet_Loss": value("Nokia_LTE_DL Packet loss", "DL_Packet_Loss"),
        "UL_Packet_Loss": value("Nokia_LTE_UL Packet loss", "UL_Packet_Loss"),
        "DL_PRB_Util": value("PDSCH utilization rate PRB DL", "DL_PRB_Util"),
        "UL_PRB_Util": value("PDSCH utilization rate PRB UL", "UL_PRB_Util"),
        "DL_User_Throughput_Mbps": value("NSN_LTE_DL Throughput Per UE (NA)", "DL_User_Throughput_Mbps"),
        "UL_User_Throughput_Mbps": value("NSN_LTE_UL Throughput Per UE (NA)", "UL_User_Throughput_Mbps"),
        "CQI": value("E-UTRAN Average CQI", "CQI"),
        "Max_Users": int(value("Max RRC connected users", "Max_Users")),
        "Active_Users_Avg": int(value("Avg RRC conn UE", "Active_Users_Avg")),
        "Handover_SR": value("Intra Frequency Handover Sucess Rate New", "Handover_SR", default=99.0),
        "RSSI": value("UL PUSCH RSSI", "RSSI", default=-115.0),
        "SINR": value("SINR", "SNIR", "SINR", default=10.0),
    }
    return record


def read_xlsx_rows(path):
    with zipfile.ZipFile(path) as workbook:
        shared_strings = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(text.text or "" for text in item.findall(".//main:t", EXCEL_NS))
                for item in root.findall(".//main:si", EXCEL_NS)
            ]

        sheet_name = "xl/worksheets/sheet1.xml"
        if sheet_name not in workbook.namelist():
            raise ValueError(f"{path.name} has no first worksheet")

        def cell_value(cell):
            cell_type = cell.attrib.get("t")
            value = cell.find("main:v", EXCEL_NS)
            if cell_type == "inlineStr":
                return "".join(text.text or "" for text in cell.findall(".//main:t", EXCEL_NS))
            if value is None:
                return ""
            text = value.text or ""
            if cell_type == "s" and text.isdigit():
                index = int(text)
                return shared_strings[index] if index < len(shared_strings) else text
            return text

        headers = {}
        with workbook.open(sheet_name) as source:
            for _, element in ET.iterparse(source, events=("end",)):
                if not element.tag.endswith("row"):
                    continue
                row = {}
                for cell in element.findall("main:c", EXCEL_NS):
                    reference = cell.attrib.get("r", "")
                    column = "".join(character for character in reference if character.isalpha())
                    if not column:
                        continue
                    if element.attrib.get("r") == "1":
                        headers[column] = str(cell_value(cell)).strip()
                    else:
                        header = headers.get(column)
                        if header:
                            value = cell_value(cell)
                            if header not in row or row[header] in ("", None):
                                row[header] = value
                if element.attrib.get("r") != "1" and row:
                    yield row
                element.clear()


def extract_hourly(path):
    records = []
    for row in read_xlsx_rows(path):
        record = normalize_record(row)
        if record and record["Site_ID"] and record["LNCEL"]:
            records.append(record)
    if not records:
        raise ValueError(f"No hourly KPI rows could be read from {path.name}")
    return records


def extract_daywise(path):
    if path.suffix.lower() == ".xlsb":
        try:
            import pandas as pd
        except ImportError as error:
            raise RuntimeError("Day Wise XLSB import requires pandas and pyxlsb") from error
        rows = pd.read_excel(path, sheet_name=0, engine="pyxlsb").to_dict(orient="records")
    else:
        rows = read_xlsx_rows(path)
    records = []
    for row in rows:
        record = normalize_record(row, daily=True)
        if record and record["Site_ID"] and record["LNCEL"]:
            records.append(record)
    if not records:
        raise ValueError(f"No Day Wise KPI rows could be read from {path.name}")
    return records


def report_date(path):
    matches = re.findall(r"(?<!\d)(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})(?!\d)", path.stem)
    for day, month, year in reversed(matches):
        try:
            return dt.datetime(int(year), int(month), int(day)).date()
        except ValueError:
            pass
    matches = re.findall(r"(?<!\d)(\d{4})[-_.](\d{1,2})[-_.](\d{1,2})(?!\d)", path.stem)
    for year, month, day in reversed(matches):
        try:
            return dt.datetime(int(year), int(month), int(day)).date()
        except ValueError:
            pass
    return dt.date.min


def latest_report(kind):
    reports = find_reports(kind)
    if not reports:
        raise FileNotFoundError(f"No {kind} KPI Excel report found in {SOURCE_DIR}")
    return max(reports, key=lambda path: (report_date(path), path.stat().st_mtime_ns))


def find_reports(kind):
    extensions = {".xlsx", ".xlsb"}
    if not SOURCE_DIR.is_dir():
        raise FileNotFoundError(f"OneDrive KPI folder not found: {SOURCE_DIR}")
    candidates = []
    for path in SOURCE_DIR.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in extensions or path.name.startswith("~$"):
            continue
        normalized_name = re.sub(r"[^a-z0-9]+", " ", path.stem.lower())
        if kind == "hourly" and ("hourly" in normalized_name and "kpi report" in normalized_name):
            candidates.append(path)
        elif kind == "daywise" and (("daywise" in normalized_name or "day wise" in normalized_name) and "kpi report" in normalized_name):
            candidates.append(path)
    return candidates


def fingerprint(path):
    stat = path.stat()
    return {"name": path.name, "size": stat.st_size, "modified_ns": stat.st_mtime_ns}


def write_atomic(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as output:
            output.write(text)
            output.flush()
            os.fsync(output.fileno())
        os.replace(temporary_name, path)
    except Exception:
        try:
            os.unlink(temporary_name)
        except OSError:
            pass
        raise


def compact_js(variable, records, title):
    payload = {
        "columns": COMPACT_COLUMNS,
        "rows": [[record[column] for column in COMPACT_COLUMNS] for record in records],
    }
    return f"// {title}\nwindow.{variable} = {json.dumps(payload, separators=(',', ':'))};\n"


def load_compact(path, variable):
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8")
    prefix = f"window.{variable} = "
    if not text.startswith("//") or prefix not in text:
        raise ValueError(f"Unexpected compact data format in {path.name}")
    payload_start = text.index(prefix) + len(prefix)
    payload_text = text[payload_start:].strip()
    if payload_text.endswith(";"):
        payload_text = payload_text[:-1]
    payload = json.loads(payload_text)
    columns = payload["columns"]
    return [dict(zip(columns, values)) for values in payload["rows"]]


def record_key(record, daily):
    identity = (
        record["Date"],
        record["Operator"],
        record["Band"],
        record["Site_ID"] or record["Site_Name"],
        record["LNCEL"],
    )
    return identity if daily else identity + (record["Hour"],)


def append_missing_dates(existing_records, incoming_records, daily):
    existing_dates = {record["Date"] for record in existing_records}
    additions = {}
    for record in incoming_records:
        if record["Date"] in existing_dates:
            continue
        additions[record_key(record, daily)] = record
    merged = existing_records + list(additions.values())
    merged.sort(key=lambda record: (
        record["Date"],
        record["Hour"],
        record["Operator"],
        record["Site_Name"],
        record["LNCEL"],
    ))
    return merged, sorted({record["Date"] for record in additions.values()})


def read_source_fingerprints(hourly_paths, daywise_paths):
    sources = {"hourly": {}, "daywise": {}}
    for kind, paths in (("hourly", hourly_paths), ("daywise", daywise_paths)):
        for path in paths:
            sources[kind][str(path.relative_to(SOURCE_DIR))] = fingerprint(path)
    return sources


def run_update():
    hourly_paths = find_reports("hourly")
    daywise_paths = find_reports("daywise")
    if not hourly_paths or not daywise_paths:
        missing = [kind for kind, paths in (("hourly", hourly_paths), ("daywise", daywise_paths)) if not paths]
        raise FileNotFoundError(f"Missing {', '.join(missing)} KPI report(s) in {SOURCE_DIR}")
    sources = read_source_fingerprints(hourly_paths, daywise_paths)
    all_paths = hourly_paths + daywise_paths
    if any(time.time_ns() - path.stat().st_mtime_ns < 20_000_000_000 for path in all_paths):
        log("A workbook is still syncing; will retry on the next scan.")
        return False

    try:
        previous = json.loads(METADATA_OUTPUT.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        previous = {}
    if previous.get("sources") == sources and HOURLY_OUTPUT.exists() and DAYWISE_OUTPUT.exists():
        log("No new report changes found.")
        return False

    hourly_records = load_compact(HOURLY_OUTPUT, "UserLTEDataCompact")
    daywise_records = load_compact(DAYWISE_OUTPUT, "DaywiseLTEData")
    original_hourly_dates = {record["Date"] for record in hourly_records}
    original_daywise_dates = {record["Date"] for record in daywise_records}

    for kind, paths, existing_records, extractor, daily in (
        ("hourly", hourly_paths, hourly_records, extract_hourly, False),
        ("Day Wise", daywise_paths, daywise_records, extract_daywise, True),
    ):
        incoming = []
        for path in sorted(paths, key=lambda item: (report_date(item), item.stat().st_mtime_ns, str(item).lower())):
            log(f"Scanning {kind} report: {path.name}")
            incoming.extend(extractor(path))
        merged, added_dates = append_missing_dates(existing_records, incoming, daily)
        if kind == "hourly":
            hourly_records = merged
            log(f"Added {len(added_dates)} missing hourly date(s): {', '.join(added_dates) if added_dates else 'none'}")
        else:
            daywise_records = merged
            log(f"Added {len(added_dates)} missing Day Wise date(s): {', '.join(added_dates) if added_dates else 'none'}")

    metadata = {
        "sources": sources,
        "hourlyFile": latest_report("hourly").name,
        "daywiseFile": latest_report("daywise").name,
        "hourlyRecords": len(hourly_records),
        "daywiseRecords": len(daywise_records),
        "hourlyLatestDate": max(record["Date"] for record in hourly_records),
        "daywiseLatestDate": max(record["Date"] for record in daywise_records),
        "hourlyDatesAdded": sorted({record["Date"] for record in hourly_records} - original_hourly_dates),
        "daywiseDatesAdded": sorted({record["Date"] for record in daywise_records} - original_daywise_dates),
        "updatedAt": previous.get("updatedAt", dt.datetime.now().astimezone().isoformat(timespec="seconds")),
    }
    data_changed = (
        metadata["hourlyDatesAdded"]
        or metadata["daywiseDatesAdded"]
        or not HOURLY_OUTPUT.exists()
        or not DAYWISE_OUTPUT.exists()
        or not METADATA_JS_OUTPUT.exists()
    )
    if data_changed:
        metadata["updatedAt"] = dt.datetime.now().astimezone().isoformat(timespec="seconds")
        write_atomic(HOURLY_OUTPUT, compact_js("UserLTEDataCompact", hourly_records, "Compact hourly KPI report data"))
        write_atomic(DAYWISE_OUTPUT, compact_js("DaywiseLTEData", daywise_records, "Compact Day Wise KPI report data"))
        write_atomic(METADATA_JS_OUTPUT, "window.KPIDataSources = " + json.dumps(metadata, separators=(",", ":")) + ";\n")
        log(f"Published {len(hourly_records):,} hourly and {len(daywise_records):,} Day Wise records.")
    else:
        log("No missing dates found; existing date data was left unchanged.")
    write_atomic(METADATA_OUTPUT, json.dumps(metadata, indent=2) + "\n")
    return bool(data_changed)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--once", action="store_true", help="scan and update once, then exit")
    parser.add_argument("--interval", type=int, default=300, help="scan interval in seconds (default: 300)")
    args = parser.parse_args()

    if args.once:
        run_update()
        return
    while True:
        try:
            run_update()
        except Exception as error:
            log(f"Update failed: {error}")
        time.sleep(max(30, args.interval))


if __name__ == "__main__":
    main()
