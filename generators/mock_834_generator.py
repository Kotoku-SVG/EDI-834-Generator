#!/usr/bin/env python3
"""
mock_834_generator.py

Reads the Excel template 'mock_834_generator_template.xlsx' and outputs a simplified EDI 834 file.

Usage:
  python mock_834_generator.py --in mock_834_generator_template.xlsx --out out_834.txt

Notes:
- This is a MOCK generator for testing. Real 834 implementations must follow each carrier's companion guide.
- Segment choices here are common but not exhaustive.
"""
from __future__ import annotations
import argparse
from datetime import datetime
from pathlib import Path
import pandas as pd

SEG_TERM = "~"
ELEM_SEP = "*"
COMP_SEP = ":"

def seg(*elements: str) -> str:
    return ELEM_SEP.join([e if e is not None else "" for e in elements]) + SEG_TERM

def yyyymmdd(s: str) -> str:
    s = str(s).strip()
    if not s:
        return ""
    # accept YYYYMMDD, YYYY-MM-DD, datetime-like
    if len(s) == 8 and s.isdigit():
        return s
    try:
        dt = pd.to_datetime(s)
        return dt.strftime("%Y%m%d")
    except Exception:
        raise ValueError(f"Bad date '{s}' (expected YYYYMMDD or parseable date)")

def read_settings(xlsx: Path) -> dict:
    df = pd.read_excel(xlsx, sheet_name="Settings", header=2, usecols=[0,1]).dropna()
    # 'Field' column contains same as value; use first col as keys
    settings = {}
    for _, row in df.iterrows():
        key = str(row.iloc[0]).strip()
        val = row.iloc[1]
        settings[key] = "" if pd.isna(val) else str(val).strip()
    return settings

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="Input Excel template")
    ap.add_argument("--out", dest="out", required=True, help="Output 834 text file")
    ap.add_argument("--test", action="store_true", help="Also print a few lines to stdout")
    args = ap.parse_args()

    xlsx = Path(args.inp)
    outp = Path(args.out)

    settings = read_settings(xlsx)

    # Pull core settings (with safe defaults)
    sender = settings.get("Sender_ID (ISA06)", "SENDERID")
    receiver = settings.get("Receiver_ID (ISA08)", "RECEIVERID")
    icn = settings.get("Interchange_Control (ISA13)", "1").zfill(9)
    gcn = settings.get("Group_Control (GS06)", "1")
    tcn = settings.get("Transaction_Control (ST02)", "0001")
    sponsor_name = settings.get("Sponsor_Name (Employer)", "SPONSOR")
    sponsor_id = settings.get("Sponsor_ID (Employer ID)", "000000000")
    payer_name = settings.get("Payer_Name (Carrier)", "PAYER")
    payer_id = settings.get("Payer_ID (Carrier ID)", "999999")
    file_type = settings.get("File_Type", "FULL").upper()
    as_of = yyyymmdd(settings.get("As_Of_Date", datetime.now().strftime("%Y-%m-%d")))

    plans = pd.read_excel(xlsx, sheet_name="Plans").fillna("")
    members = pd.read_excel(xlsx, sheet_name="Members").fillna("")
    deps = pd.read_excel(xlsx, sheet_name="Dependents").fillna("")

    plan_map = {str(r["Plan_Key"]).strip(): r for _, r in plans.iterrows() if str(r["Plan_Key"]).strip()}

    now = datetime.now()
    isa_date = now.strftime("%y%m%d")
    isa_time = now.strftime("%H%M")
    gs_date = now.strftime("%Y%m%d")
    gs_time = now.strftime("%H%M")

    # Build EDI
    edi = []
    # ISA: use fixed width elements where typical. This is simplified.
    edi.append(seg("ISA","00","          ","00","          ","ZZ",sender.ljust(15)[:15],"ZZ",receiver.ljust(15)[:15],isa_date,isa_time,"^","00501",icn,"0","P",">"))
    edi.append(seg("GS","BE",sender,receiver,gs_date,gs_time,gcn,"X","005010X220A1"))
    edi.append(seg("ST","834",tcn,"005010X220A1"))
    # BGN: 00=original
    edi.append(seg("BGN","00",tcn,gs_date,gs_time,"","",file_type))
    edi.append(seg("DTP","007","D8",as_of))  # effective/as-of

    # Sponsor
    edi.append(seg("N1","P5",sponsor_name,"FI",sponsor_id))
    # Payer
    edi.append(seg("N1","IN",payer_name,"FI",payer_id))

    # Member loops
    for _, m in members.iterrows():
        sub_id = str(m.get("Subscriber_ID","")).strip()
        if not sub_id:
            continue
        ssn = str(m.get("Subscriber_SSN","")).strip()
        last = str(m.get("Sub_Last","")).strip()
        first = str(m.get("Sub_First","")).strip()
        middle = str(m.get("Sub_Middle","")).strip()
        dob = yyyymmdd(m.get("Sub_DOB",""))
        gender = str(m.get("Sub_Gender","")).strip()
        addr1 = str(m.get("Sub_Address1","")).strip()
        city = str(m.get("Sub_City","")).strip()
        state = str(m.get("Sub_State","")).strip()
        zipc = str(m.get("Sub_Zip","")).strip()
        emp_status = str(m.get("Employment_Status","")).strip()
        action = str(m.get("Action","ADD")).strip().upper()
        cov_start = yyyymmdd(m.get("Coverage_Start",""))
        cov_end = yyyymmdd(m.get("Coverage_End",""))
        plan_key = str(m.get("Plan_Key","")).strip()
        tier = str(m.get("Coverage_Tier_Code","")).strip()

        # INS: member level
        # INS01: Y/N subscriber; INS02: 18=self (subscriber)
        # INS03: maintenance type code (001=add, 002=change, 024=cancel) - simplified mapping
        mtc = {"ADD":"001","CHG":"002","TERM":"024"}.get(action,"001")
        edi.append(seg("INS","Y","18",mtc,"XN","A","E","","",emp_status))

        # NM1: subscriber
        # NM108/109: identification code qualifier/ID (34=SSN, else use employee ID)
        if ssn and ssn.isdigit() and len(ssn)==9:
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
            edi.append(seg("DTP","356","D8",cov_start))  # eligibility begin
        if cov_end:
            edi.append(seg("DTP","357","D8",cov_end))    # eligibility end

        # Coverage (HD loop) for subscriber
        if plan_key:
            if plan_key not in plan_map:
                raise KeyError(f"Plan_Key '{plan_key}' not found in Plans sheet.")
            pr = plan_map[plan_key]
            line = str(pr.get("HD_Insurance_Line_Code","")).strip() or str(pr.get("Benefit_Type_Code","")).strip()
            plan_desc = str(pr.get("HD_Plan_Coverage_Desc","")).strip() or plan_key
            edi.append(seg("HD", "030", "", line, plan_desc, tier))
            if cov_start:
                edi.append(seg("DTP","348","D8",cov_start))  # benefit begin
            if cov_end:
                edi.append(seg("DTP","349","D8",cov_end))    # benefit end

        # Dependent loops tied to subscriber
        sub_deps = deps[deps["Subscriber_ID"].astype(str).str.strip() == sub_id]
        for _, d in sub_deps.iterrows():
            dep_id = str(d.get("Dep_ID","")).strip()
            rel = str(d.get("Relationship","")).strip()
            d_ssn = str(d.get("Dep_SSN","")).strip()
            d_last = str(d.get("Dep_Last","")).strip()
            d_first = str(d.get("Dep_First","")).strip()
            d_mid = str(d.get("Dep_Middle","")).strip()
            d_dob = yyyymmdd(d.get("Dep_DOB",""))
            d_gender = str(d.get("Dep_Gender","")).strip()
            d_action = str(d.get("Action","ADD")).strip().upper()
            d_start = yyyymmdd(d.get("Coverage_Start","")) or cov_start
            d_end = yyyymmdd(d.get("Coverage_End","")) or cov_end
            d_plan_key = str(d.get("Plan_Key","")).strip() or plan_key
            d_mtc = {"ADD":"001","CHG":"002","TERM":"024"}.get(d_action,"001")

            # INS: dependent (not subscriber -> N)
            # INS02: relationship code (19=child, 01=spouse commonly); simplified mapping:
            rel_code = {"SPO":"01","CHD":"19"}.get(rel,"34")  # 34=other adult
            edi.append(seg("INS","N",rel_code,d_mtc,"XN","A","E"))

            if d_ssn and d_ssn.isdigit() and len(d_ssn)==9:
                edi.append(seg("NM1","IL","1",d_last,d_first,d_mid,"","", "34", d_ssn))
            else:
                # Use ZZ + composite id subscriber+dep
                edi.append(seg("NM1","IL","1",d_last,d_first,d_mid,"","", "ZZ", f"{sub_id}-{dep_id}" if dep_id else f"{sub_id}-DEP"))

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
                line_d = str(prd.get("HD_Insurance_Line_Code","")).strip() or str(prd.get("Benefit_Type_Code","")).strip()
                plan_desc_d = str(prd.get("HD_Plan_Coverage_Desc","")).strip() or d_plan_key
                edi.append(seg("HD","030","",line_d,plan_desc_d,""))
                if d_start:
                    edi.append(seg("DTP","348","D8",d_start))
                if d_end:
                    edi.append(seg("DTP","349","D8",d_end))

    # SE count includes ST and SE
    # We'll compute after assembling
    # Placeholder, then replace
    edi.append("SE_PLACEHOLDER")
    edi.append(seg("GE","1",gcn))
    edi.append(seg("IEA","1",icn))

    # compute SE segment count: segments from ST to SE inclusive
    # find ST index
    st_index = next(i for i,s in enumerate(edi) if s.startswith("ST"+ELEM_SEP))
    # SE placeholder index is second to last before GE/IEA; find it
    se_index = next(i for i,s in enumerate(edi) if s == "SE_PLACEHOLDER")
    seg_count = (se_index - st_index) + 1
    edi[se_index] = seg("SE", str(seg_count), tcn)

    outp.write_text("\n".join(edi), encoding="utf-8")

    if args.test:
        print("Wrote:", outp)
        print("First 12 lines:")
        for line in edi[:12]:
            print(line)

if __name__ == "__main__":
    main()
