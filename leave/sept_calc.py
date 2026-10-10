"""September 2026 AL, LSL and TIL journals, worked exactly the way the
AL-LSL Provision workbook works them. Reads the August workbook for last
month's figures and the Employment Hero Leave balances report for September."""
import openpyxl, collections, json, sys

U = "/root/.claude/uploads/a91f56e6-b7da-5dad-8e0b-6d5f516f9360/"
AUG = U + "55bf3a59-2026-08_AL-LSL_Provision.xlsx"
SEP = U + "ea1b8da1-LeaveBalances_as_at_2026-09-30_1.xlsx"
TYPES = ("Annual Leave", "Long Service Leave", "Time In Lieu Taken")

def load_sep():
    ws = openpyxl.load_workbook(SEP, data_only=True)["Export"]
    out = []
    for r in ws.iter_rows(min_row=2, max_col=12, values_only=True):
        if r[5] in TYPES:
            out.append(list(r))
    return out

def full(r):
    return ("%s %s" % (r[2], r[3])).strip().lower()

def main():
    sep = load_sep()
    wv = openpyxl.load_workbook(AUG, data_only=True)
    res = {}

    # ---- Annual Leave: each tab row, SUMIFS by full name; difference to August
    al = wv["Annual Leave"]
    by = collections.defaultdict(float)
    for r in sep:
        if r[5] == "Annual Leave":
            by[full(r)] += r[9] or 0
    rows, matched = [], set()
    for i in range(3, 98):
        name = al.cell(row=i, column=1).value
        dept = al.cell(row=i, column=2).value
        aug = al.cell(row=i, column=60).value or 0          # BH
        s = by.get((name or "").strip().lower(), 0.0)
        if name: matched.add(name.strip().lower())
        rows.append((i, name, dept, aug, round(s, 2), round(s - aug, 2)))
    unmatched = {k: v for k, v in by.items() if k not in matched}
    tot_tab = sum(r[4] for r in rows); tot_data = sum(by.values())
    groups = {"OFFICE / ADMIN": ("Office/Admin", "Finance", "HR", "Management"),
              "PRODUCTION (excl VIDEO)": ("Production",), "ONSITE (including PRD backup)": ("Onsite",),
              "VIDEO DEPT": ("Video",), "INTEGRATION": ("Integration",), "Consulting": ("Consulting",)}
    al_lines = {g: round(sum(r[5] for r in rows if r[2] in ds), 2) for g, ds in groups.items()}
    res["AL"] = {"aug_total": round(sum(r[3] for r in rows), 2), "sep_total_tab": round(tot_tab, 2),
                 "sep_total_data": round(tot_data, 2), "unmatched": unmatched,
                 "movement": round(tot_tab - sum(r[3] for r in rows), 2), "lines": al_lines,
                 "lines_total": round(sum(al_lines.values()), 2)}

    # ---- Long Service Leave: the August block (JI names, JJ dept, JK value)
    ls = wv["Long Service Leave"]
    lby = collections.defaultdict(float)
    for r in sep:
        if r[5] == "Long Service Leave":
            lby[full(r)] += r[9] or 0
    lrows = []
    for i in range(2, 18):
        name = ls.cell(row=i, column=269).value; dept = ls.cell(row=i, column=270).value
        aug = ls.cell(row=i, column=271).value or 0
        s = lby.get((name or "").strip().lower(), 0.0) if name else 0.0
        lrows.append((i, name, dept, aug, round(s, 2), round(s - aug, 2)))
    # the extra Onsite term in JP32 reads an old block (IP/IR); carry its August value
    jp32 = ls.cell(row=32, column=276).value
    ons_now = sum(r[3] - 0 for r in []) # placeholder
    lsl_depts = ["OFFICE / ADMIN", "PRODUCTION (excl VIDEO)", "ONSITE (including PRD backup)",
                 "VIDEO DEPT", "INTEGRATION", "Consulting"]
    lsl_lines = {d: round(sum(r[5] for r in lrows if r[2] == d), 2) for d in lsl_depts}
    aug_ons_block = sum(ls.cell(row=i, column=272).value or 0 for i in range(2, 18)
                        if ls.cell(row=i, column=270).value == "ONSITE (including PRD backup)")
    ip_term = round(jp32 - aug_ons_block, 2)
    res["LSL"] = {"rows": lrows, "lines": lsl_lines, "ip_term_in_JP32": ip_term,
                  "movement": round(sum(r[5] for r in lrows), 2),
                  "sep_block_total": round(sum(r[4] for r in lrows), 2)}

    # ---- Time in Lieu: value by location
    til = collections.defaultdict(float); til_rows = []
    for r in sep:
        if r[5] == "Time In Lieu Taken":
            til[r[4]] += r[9] or 0; til_rows.append(r)
    res["TIL"] = {"by_location": {k: round(v, 2) for k, v in sorted(til.items())},
                  "total": round(sum(til.values()), 2), "rows": len(til_rows)}
    json.dump(res, open("/home/user/Claude/leave/sept_calc.json", "w"), indent=1, default=str)
    return res

if __name__ == "__main__":
    r = main()
    a = r["AL"]
    print("AL  Aug total %.2f  Sep tab %.2f  Sep data %.2f  movement %.2f" % (a["aug_total"], a["sep_total_tab"], a["sep_total_data"], a["movement"]))
    print("    unmatched in data:", a["unmatched"])
    for k, v in a["lines"].items(): print("    %-32s %10.2f" % (k, v))
    print("    lines total %.2f" % a["lines_total"])
    l = r["LSL"]
    print("LSL movement %.2f  Sep block total %.2f  IP term in JP32 %.2f" % (l["movement"], l["sep_block_total"], l["ip_term_in_JP32"]))
    for row in l["rows"]: print("    ", row)
    for k, v in l["lines"].items(): print("    %-32s %10.2f" % (k, v))
    t = r["TIL"]
    print("TIL total %.2f rows %d" % (t["total"], t["rows"]))
    for k, v in t["by_location"].items(): print("    %-34s %9.2f" % (k, v))
