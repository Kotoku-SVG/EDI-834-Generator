#!/usr/bin/env python3
"""
mock_834_generator_nolibs.py

Self-contained (stdlib-only) MOCK EDI 834 generator.
- No pandas
- No openpyxl
- No third-party libraries

It reads an .xlsx by parsing the underlying OpenXML (zip + XML).
It expects the workbook template layout:
  Sheets: Settings, Plans, Members, Dependents
  - Plans/Members/Dependents: row 1 = headers
  - Settings: labels in column A, values in column B (any row)

Usage:
  python mock_834_generator_nolibs.py --in mock_834_generator_template.xlsx --out out_834.txt

Notes:
- This is a simplified 834 for testing.
- Real carrier companion guides vary significantly.
"""

from __future__ import annotations
import argparse
import zipfile
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

SEG_TERM = "~"
ELEM_SEP = "*"

# -------------------- XLSX (OpenXML) minimal reader --------------------

NS = {
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "w": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

def _strip(s: str) -> str:
    return "" if s is None else str(s).strip()

def _col_letters_to_index(col: str) -> int:
    # A->1, B->2 ... Z->26, AA->27 ...
    col = col.upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n

def _cell_ref_to_rc(ref: str) -> tuple[int, int]:
    # "C12" -> (12, 3)
    m = re.match(r"^([A-Za-z]+)(\d+)$", ref)
    if not m:
        return (0, 0)
    c = _col_letters_to_index(m.group(1))
    r = int(m.group(2))
    return (r, c)

def read_xlsx_tables(xlsx_path: Path) -> dict[str, list[dict[str, str]]]:
    """
    Returns dict of sheet_name -> list[rows as dict(header->value)] for table-like sheets
    plus Settings returned as list of dicts with keys Field/Value if it matches.
    """
    with zipfile.ZipFile(xlsx_path, "r") as z:
        # Shared strings
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in xml.findall("w:si", NS):
                # concatenate all t nodes within si
                texts = [t.text or "" for t in si.findall(".//w:t", NS)]
                shared_strings.append("".join(texts))
        # Workbook: map sheet name -> r:id
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = []
        for sh in wb_xml.findall(".//w:sheets/w:sheet", NS):
            name = sh.attrib.get("name", "")
            rid = sh.attrib.get(f"{{{NS['r']}}}id", "")
            sheets.append((name, rid))

        # workbook rels: map r:id -> target (e.g., worksheets/sheet1.xml)
        rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {}
        for rel in rels_xml.findall("rel:Relationship", NS):
            rid_to_target[rel.attrib.get("Id","")] = rel.attrib.get("Target","")

        def parse_sheet(sheet_target: str) -> dict[tuple[int,int], str]:
            path = "xl/" + sheet_target.lstrip("/")
            xml = ET.fromstring(z.read(path))
            cells: dict[tuple[int,int], str] = {}
            for c in xml.findall(".//w:sheetData/w:row/w:c", NS):
                ref = c.attrib.get("r","")
                r, col = _cell_ref_to_rc(ref)
                t = c.attrib.get("t","")  # 's' for shared string
                v = c.find("w:v", NS)
                if v is None or v.text is None:
                    value = ""
                else:
                    raw = v.text
                    if t == "s":
                        try:
                            value = shared_strings[int(raw)]
                        except Exception:
                            value = raw
                    else:
                        value = raw
                cells[(r,col)] = value
            return cells

        # Convert cell map to row dicts using first row headers
        results: dict[str, list[dict[str,str]]] = {}

        for name, rid in sheets:
            target = rid_to_target.get(rid, "")
            if not target or not target.startswith("worksheets/"):
                continue
            cells = parse_sheet(target)

            # Determine max row/col used
            if not cells:
                results[name] = []
                continue

            max_row = max(r for (r,c) in cells.keys())
            max_col = max(c for (r,c) in cells.keys())

            # Settings: key/value in col 1/2 anywhere below header row (the template uses a title row)
            if name.lower() == "settings":
                kv = {}
                # Scan rows 1..max_row; treat col1 as key, col2 as value
                for r in range(1, max_row+1):
                    k = _strip(cells.get((r,1), ""))
                    v = _strip(cells.get((r,2), ""))
                    if k and k.lower() not in ("field", "mock 834 generator - settings"):
                        # Keep last occurrence
                        kv[k] = v
                # Return as a single-row dict for convenience
                results[name] = [kv]
                continue

            # Table sheets: row1 headers
            headers = []
            for c in range(1, max_col+1):
                h = _strip(cells.get((1,c), ""))
                headers.append(h)

            table = []
            for r in range(2, max_row+1):
                row = {}
                empty = True
                for c in range(1, max_col+1):
                    h = headers[c-1] if c-1 < len(headers) else ""
                    if not h:
                        continue
                    val = _strip(cells.get((r,c), ""))
                    if val != "":
                        empty = False
                    row[h] = val
                if not empty:
                    table.append(row)
            results[name] = table

        return results

# -------------------- EDI helpers --------------------

def seg(*elements: str) -> str:
    return ELEM_SEP.join([e if e is not None else "" for e in elements]) + SEG_TERM

def yyyymmdd(val: str) -> str:
    s = _strip(val)
    if not s:
        return ""
    # Already YYYYMMDD
    if len(s) == 8 and s.isdigit():
        return s
    # ISO date
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        return "".join(m.groups())
    # Excel numeric date sometimes appears (rare here). Try to parse as int days since 1899-12-30.
    if re.match(r"^\d+(\.\d+)?$", s):
        try:
            days = float(s)
            base = datetime(1899, 12, 30)
            dt = base + timedelta(days=days)
            return dt.strftime("%Y%m%d")
        except Exception:
            pass
    # Fallback: try common MM/DD/YYYY
    m2 = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m2:
        mm, dd, yy = m2.groups()
        return f"{yy}{int(mm):02d}{int(dd):02d}"
    raise ValueError(f"Bad date '{s}' (expected YYYYMMDD or YYYY-MM-DD)")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="Input Excel .xlsx (template)")
    ap.add_argument("--out", dest="out", required=True, help="Output 834 text file")
    ap.add_argument("--test", action="store_true", help="Print the first lines after writing")
    args = ap.parse_args()

    xlsx = Path(args.inp)
    outp = Path(args.out)

    tables = read_xlsx_tables(xlsx)

    settings = (tables.get("Settings") or [{}])[0]
    plans = tables.get("Plans") or []
    members = tables.get("Members") or []
    deps = tables.get("Dependents") or []

    # Core settings (safe defaults)
    sender = _strip(settings.get("Sender_ID (ISA06)", "SENDERID"))
    receiver = _strip(settings.get("Receiver_ID (ISA08)", "RECEIVERID"))
    icn = _strip(settings.get("Interchange_Control (ISA13)", "1")).zfill(9)
    gcn = _strip(settings.get("Group_Control (GS06)", "1"))
    tcn = _strip(settings.get("Transaction_Control (ST02)", "1"))
    sponsor_name = _strip(settings.get("Sponsor_Name (Employer)", "SPONSOR"))
    sponsor_id = _strip(settings.get("Sponsor_ID (Employer ID)", "000000000"))
    payer_name = _strip(settings.get("Payer_Name (Carrier)", "PAYER"))
    payer_id = _strip(settings.get("Payer_ID (Carrier ID)", "999999"))
    file_type = _strip(settings.get("File_Type", "FULL")).upper()
    as_of = yyyymmdd(_strip(settings.get("As_Of_Date", datetime.now().strftime("%Y-%m-%d"))))

    plan_map = {}
    for r in plans:
        key = _strip(r.get("Plan_Key"))
        if key:
            plan_map[key] = r

    now = datetime.now()
    isa_date = now.strftime("%y%m%d")
    isa_time = now.strftime("%H%M")
    gs_date = now.strftime("%Y%m%d")
    gs_time = now.strftime("%H%M")

    edi: list[str] = []
    edi.append(seg("ISA","00","          ","00","          ","ZZ",sender.ljust(15)[:15],
                   "ZZ",receiver.ljust(15)[:15],isa_date,isa_time,"^","00501",icn,"0","P",">"))
    edi.append(seg("GS","BE",sender,receiver,gs_date,gs_time,gcn,"X","005010X220A1"))
    edi.append(seg("ST","834",tcn,"005010X220A1"))
    edi.append(seg("BGN","00",tcn,gs_date,gs_time,"","",file_type))
    edi.append(seg("DTP","007","D8",as_of))
    edi.append(seg("N1","P5",sponsor_name,"FI",sponsor_id))
    edi.append(seg("N1","IN",payer_name,"FI",payer_id))

    # Index deps by subscriber for speed
    deps_by_sub: dict[str, list[dict[str,str]]] = {}
    for d in deps:
        sid = _strip(d.get("Subscriber_ID"))
        if not sid:
            continue
        deps_by_sub.setdefault(sid, []).append(d)

    for m in members:
        sub_id = _strip(m.get("Subscriber_ID"))
        if not sub_id:
            continue

        ssn = _strip(m.get("Subscriber_SSN"))
        last = _strip(m.get("Sub_Last"))
        first = _strip(m.get("Sub_First"))
        middle = _strip(m.get("Sub_Middle"))
        dob = yyyymmdd(_strip(m.get("Sub_DOB")))
        gender = _strip(m.get("Sub_Gender"))
        addr1 = _strip(m.get("Sub_Address1"))
        city = _strip(m.get("Sub_City"))
        state = _strip(m.get("Sub_State"))
        zipc = _strip(m.get("Sub_Zip"))
        emp_status = _strip(m.get("Employment_Status"))
        action = _strip(m.get("Action") or "ADD").upper()
        cov_start = yyyymmdd(_strip(m.get("Coverage_Start")))
        cov_end = yyyymmdd(_strip(m.get("Coverage_End")))
        plan_key = _strip(m.get("Plan_Key"))
        tier = _strip(m.get("Coverage_Tier_Code"))

        mtc = {"ADD":"001","CHG":"002","TERM":"024"}.get(action, "001")
        edi.append(seg("INS","Y","18",mtc,"XN","A","E","","",emp_status))

        if ssn.isdigit() and len(ssn) == 9:
            edi.append(seg("NM1","IL","1",last,first,middle,"","", "34", ssn))
        else:
            edi.append(seg("NM1","IL","1",last,first,middle,"","", "ZZ", sub_id))

        if addr1:
            edi.append(seg("N3",addr1))
        if city or state or zipc:
            edi.append(seg("N4",city,state,zipc))
        if dob or gender:
            edi.append(seg("DMG","D8",dob,gender))

        if cov_start:
            edi.append(seg("DTP","356","D8",cov_start))
        if cov_end:
            edi.append(seg("DTP","357","D8",cov_end))

        # Subscriber coverage
        if plan_key:
            if plan_key not in plan_map:
                raise KeyError(f"Plan_Key '{plan_key}' not found in Plans sheet.")
            pr = plan_map[plan_key]
            line = _strip(pr.get("HD_Insurance_Line_Code")) or _strip(pr.get("Benefit_Type_Code"))
            plan_desc = _strip(pr.get("HD_Plan_Coverage_Desc")) or plan_key
            edi.append(seg("HD","030","",line,plan_desc,tier))
            if cov_start:
                edi.append(seg("DTP","348","D8",cov_start))
            if cov_end:
                edi.append(seg("DTP","349","D8",cov_end))

        # Dependents
        for d in deps_by_sub.get(sub_id, []):
            dep_id = _strip(d.get("Dep_ID"))
            rel = _strip(d.get("Relationship"))
            d_ssn = _strip(d.get("Dep_SSN"))
            d_last = _strip(d.get("Dep_Last"))
            d_first = _strip(d.get("Dep_First"))
            d_mid = _strip(d.get("Dep_Middle"))
            d_dob = yyyymmdd(_strip(d.get("Dep_DOB")))
            d_gender = _strip(d.get("Dep_Gender"))
            d_action = _strip(d.get("Action") or "ADD").upper()
            d_start = yyyymmdd(_strip(d.get("Coverage_Start"))) or cov_start
            d_end = yyyymmdd(_strip(d.get("Coverage_End"))) or cov_end
            d_plan_key = _strip(d.get("Plan_Key")) or plan_key
            d_mtc = {"ADD":"001","CHG":"002","TERM":"024"}.get(d_action, "001")

            rel_code = {"SPO":"01","CHD":"19"}.get(rel, "34")
            edi.append(seg("INS","N",rel_code,d_mtc,"XN","A","E"))

            if d_ssn.isdigit() and len(d_ssn) == 9:
                edi.append(seg("NM1","IL","1",d_last,d_first,d_mid,"","", "34", d_ssn))
            else:
                comp = f"{sub_id}-{dep_id}" if dep_id else f"{sub_id}-DEP"
                edi.append(seg("NM1","IL","1",d_last,d_first,d_mid,"","", "ZZ", comp))

            if d_dob or d_gender:
                edi.append(seg("DMG","D8",d_dob,d_gender))

            if d_start:
                edi.append(seg("DTP","356","D8",d_start))
            if d_end:
                edi.append(seg("DTP","357","D8",d_end))

            if d_plan_key:
                if d_plan_key not in plan_map:
                    raise KeyError(f"Plan_Key '{d_plan_key}' not found in Plans sheet.")
                prd = plan_map[d_plan_key]
                line_d = _strip(prd.get("HD_Insurance_Line_Code")) or _strip(prd.get("Benefit_Type_Code"))
                plan_desc_d = _strip(prd.get("HD_Plan_Coverage_Desc")) or d_plan_key
                edi.append(seg("HD","030","",line_d,plan_desc_d,""))
                if d_start:
                    edi.append(seg("DTP","348","D8",d_start))
                if d_end:
                    edi.append(seg("DTP","349","D8",d_end))

    # SE count: from ST to SE inclusive
    # We'll append placeholder then replace
    edi.append("SE_PLACEHOLDER")
    edi.append(seg("GE","1",gcn))
    edi.append(seg("IEA","1",icn))

    st_index = next(i for i,s in enumerate(edi) if s.startswith("ST"+ELEM_SEP))
    se_index = next(i for i,s in enumerate(edi) if s == "SE_PLACEHOLDER")
    seg_count = (se_index - st_index) + 1
    edi[se_index] = seg("SE", str(seg_count), tcn)

    outp.write_text("\n".join(edi), encoding="utf-8")

    if args.test:
        print("Wrote:", outp)
        for line in edi[:12]:
            print(line)

if __name__ == "__main__":
    main()
