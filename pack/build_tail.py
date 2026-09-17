
# ----------------------------------------------------------------------------
# One row skeleton shared by the P&L, both budget sheets, the prior year and
# the Summary, so a paste into any of them lines up with the rest.
# ----------------------------------------------------------------------------

def layout(subs):
    """(kind, label, dept, cat, sub) per row, in display order."""
    out = []
    for dept in DEPTS:
        out.append(("dept", dept, dept, None, None))
        for cat in CATEGORY_ORDER:
            cat_subs = [x for c, x in subs if c == cat]
            if not cat_subs:
                continue
            out.append(("cat", "    " + cat, dept, cat, None))
            for sub in cat_subs:
                out.append(("sub", "        " + sub, dept, cat, sub))
        out.append(("blank", "", None, None, None))
    return out


ROW0 = 6  # first body row on every reporting sheet


def row_map(rows):
    """dept row, category rows and their member subcategory rows, by department."""
    info = {}
    for i, (kind, _, dept, cat, sub) in enumerate(rows):
        r = ROW0 + i
        if kind == "dept":
            info[dept] = {"row": r, "cats": {}}
        elif kind == "cat":
            info[dept]["cats"][cat] = {"row": r, "subs": []}
        elif kind == "sub":
            info[dept]["cats"][cat]["subs"].append(r)
    return info


def paint(ws, rows, ncols):
    for i, (kind, label, *_rest) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        ws.cell(row=r, column=1, value=label)
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            if kind == "dept":
                cell.fill = HEAD_FILL
                cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
            elif kind == "cat":
                cell.fill = BAND_FILL
                cell.font = BOLD
            else:
                cell.font = BODY
        if kind == "sub":
            ws.row_dimensions[r].outlineLevel = 2
            ws.row_dimensions[r].hidden = True
        elif kind == "cat":
            ws.row_dimensions[r].outlineLevel = 1
    ws.sheet_properties.outlinePr.summaryBelow = False


def roll_up(ws, info, cols):
    """Category rows sum their subcategories; department rows sum their categories."""
    for dept, d in info.items():
        for cat, c in d["cats"].items():
            subs = c["subs"]
            for col in cols:
                ws[f"{col}{c['row']}"] = f"=SUM({col}{subs[0]}:{col}{subs[-1]})"
        for col in cols:
            parts = "+".join(f"{col}{c['row']}" for c in d["cats"].values())
            ws[f"{col}{d['row']}"] = "=" + parts


def risk(r, var_d, var_p, a, b):
    return (f'=IF(AND({a}{r}=0,{b}{r}=0),"No Activity",'
            f'IF(ABS({var_d}{r})<Setup!$B$14,"Low",'
            f'IF(OR(ABS({var_d}{r})>=Setup!$B$17,ABS({var_p}{r})>=Setup!$B$16),"High",'
            f'IF(ABS({var_p}{r})>=Setup!$B$15,"Medium","Low"))))')


def header(ws, labels, month_cols=True, widths=None):
    if month_cols:
        for i, (y, m) in enumerate(MONTHS):
            c = ws.cell(row=5, column=2 + i, value=date(y, m, 1))
            c.number_format = "mmm-yy"
    ws.cell(row=5, column=1, value="Line")
    for ref, text in labels:
        ws[f"{ref}5"] = text
    for c in ws[5]:
        if c.value is not None:
            c.font = HEAD
            c.fill = HEAD_FILL
            c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.column_dimensions["A"].width = 44
    for ref, w in (widths or []):
        ws.column_dimensions[ref].width = w
    ws.freeze_panes = "B6"
    ws.sheet_view.showGridLines = False


def titles(ws, note=None):
    ws["A1"] = "Profit & Loss"
    ws["A1"].font = TITLE
    ws["A2"] = ('=TEXT(Setup!$B$5,"mmm yyyy")&"   |   click the + and - buttons on the left to '
                'open a department, then a category, down to subcategory detail"')
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")
    if note:
        ws["A3"] = note
        ws["A3"].font = Font(name=FONT, size=9, italic=True, color="C00000")


MONTH_COLS = [get_column_letter(2 + i) for i in range(12)]
FY_COL = get_column_letter(14)          # N


def build_paste_sheet(wb, title, subs, note):
    """Budget FY27, Budget FY26 and Actual FY26 - same skeleton, yellow months."""
    ws = wb.create_sheet(title)
    titles(ws, note)
    header(ws, [(FY_COL, "FY total")],
           widths=[(c, 12) for c in MONTH_COLS] + [(FY_COL, 13)])
    rows = layout(subs)
    paint(ws, rows, 14)
    info = row_map(rows)
    for i, (kind, _, dept, cat, sub) in enumerate(rows):
        r = ROW0 + i
        if kind != "sub":
            continue
        for col in MONTH_COLS:
            c = ws[f"{col}{r}"]
            c.fill = YELLOW_FILL
            c.font = INPUT_FONT
            c.number_format = MONEY
    roll_up(ws, info, MONTH_COLS)
    for i, (kind, *_x) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        ws[f"{FY_COL}{r}"] = f"=SUM(B{r}:M{r})"
        for col in MONTH_COLS + [FY_COL]:
            ws[f"{col}{r}"].number_format = MONEY
    return ws


def sum_months(sheet, r, lo, hi):
    """Sum a row's month cells between two month indexes, inclusive, 0 based."""
    cols = MONTH_COLS[lo:hi + 1]
    if not cols:
        return "0"
    return f"SUM('{sheet}'!{cols[0]}{r}:{cols[-1]}{r})"


def ytd_expr(sheet, r):
    """Year to date is months 1 to the reporting period. SUMIF over the header row."""
    return (f"SUMIF('{sheet}'!$B$5:$M$5,\"<=\"&Setup!$B$5,'{sheet}'!$B{r}:$M{r})")


def qtr_expr(sheet, r):
    return (f"SUMIFS('{sheet}'!$B{r}:$M{r},'{sheet}'!$B$5:$M$5,\"<=\"&Setup!$B$5,"
            f"'{sheet}'!$B$5:$M$5,\">=\"&DATE(YEAR(Setup!$B$5),FLOOR(MONTH(Setup!$B$5)-1,3)+1,1))")


def build_pl(wb, subs, engine_index):
    """The P&L. Actual to the reporting month, budget after it, then her
    comparison blocks: budget, prior year, year on year and risk."""
    ws = wb.create_sheet("P&L FY27")
    titles(ws)
    labels = [(FY_COL, "FY total"), ("O", "YTD"), ("P", "YTD vs Budget"), ("Q", "Var $"),
              ("R", "Var %"), ("T", "PY"), ("U", "PY vs PY budget"), ("V", "Var $"),
              ("W", "Var %"), ("Y", "YOY"), ("Z", "YOY Budget"), ("AA", "Var $"),
              ("AB", "Var %"), ("AC", "Risk")]
    header(ws, labels,
           widths=[(c, 12) for c in MONTH_COLS] + [(FY_COL, 13)]
                  + [(r, 13) for r, _ in labels[1:]] + [("S", 3), ("X", 3), ("AC", 12)])
    rows = layout(subs)
    paint(ws, rows, 29)
    info = row_map(rows)

    for i, (kind, _, dept, cat, sub) in enumerate(rows):
        r = ROW0 + i
        if kind != "sub":
            continue
        eng = engine_index[(dept, sub)]
        for j, col in enumerate(MONTH_COLS):
            ecol = get_column_letter(5 + j)
            ws[f"{col}{r}"] = (f"=IF({col}$5<=Setup!$B$5,Engine!${ecol}{eng},"
                               f"'Budget FY27'!{col}{r})")
    roll_up(ws, info, MONTH_COLS)

    for i, (kind, *_x) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        ws[f"{FY_COL}{r}"] = f"=SUM(B{r}:M{r})"
        ws[f"O{r}"] = "=" + ytd_expr("P&L FY27", r)
        ws[f"P{r}"] = "=" + ytd_expr("Budget FY27", r)
        ws[f"Q{r}"] = f"=O{r}-P{r}"
        ws[f"R{r}"] = f"=IFERROR(Q{r}/ABS(P{r}),0)"
        ws[f"T{r}"] = "=" + ytd_expr("Actual FY26", r)
        ws[f"U{r}"] = "=" + ytd_expr("Budget FY26", r)
        ws[f"V{r}"] = f"=T{r}-U{r}"
        ws[f"W{r}"] = f"=IFERROR(V{r}/ABS(U{r}),0)"
        ws[f"Y{r}"] = f"=O{r}-T{r}"
        ws[f"Z{r}"] = f"=P{r}-U{r}"
        ws[f"AA{r}"] = f"=Y{r}-Z{r}"
        ws[f"AB{r}"] = f"=IFERROR(AA{r}/ABS(Z{r}),0)"
        ws[f"AC{r}"] = risk(r, "Q", "R", "O", "P")
        for col in MONTH_COLS + [FY_COL, "O", "P", "Q", "T", "U", "V", "Y", "Z", "AA"]:
            ws[f"{col}{r}"].number_format = MONEY
        for col in ("R", "W", "AB"):
            ws[f"{col}{r}"].number_format = PCT
        ws[f"AC{r}"].alignment = Alignment(horizontal="center")
    return ws


def build_summary(wb, subs):
    """Her Summary: three rolling months, YTD, quarter, year on year, risk."""
    ws = wb.create_sheet("Summary")
    titles(ws)
    labels = [("E", "YTD"), ("F", "YTD vs Rolling Forecast"), ("G", "YTD vs Budget"),
              ("H", "Var $"), ("I", "Var %"),
              ("K", "Quarterly"), ("L", "QTR vs QTR budget"), ("M", "Var $"), ("N", "Var %"),
              ("P", "YOY"), ("Q", "YOY Budget"), ("R", "Var $"), ("S", "Var %"),
              ("T", "Risk")]
    header(ws, labels, month_cols=False,
           widths=[("B", 13), ("C", 13), ("D", 13)]
                  + [(r, 14) for r, _ in labels] + [("J", 3), ("O", 3), ("T", 12)])
    # The three rolling months are the quarter the reporting month sits in.
    for col, off in (("B", 0), ("C", 1), ("D", 2)):
        c = ws[f"{col}5"]
        c.value = (f'=DATE(YEAR(Setup!$B$5),FLOOR(MONTH(Setup!$B$5)-1,3)+1+{off},1)')
        c.number_format = "mmm yyyy"
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")

    rows = layout(subs)
    paint(ws, rows, 20)
    info = row_map(rows)

    for i, (kind, *_x) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        for col in ("B", "C", "D"):
            ws[f"{col}{r}"] = (f"=SUMIF('P&L FY27'!$B$5:$M$5,{col}$5,'P&L FY27'!$B{r}:$M{r})")
        ws[f"E{r}"] = "=" + ytd_expr("P&L FY27", r)
        ws[f"F{r}"] = "=" + ytd_expr("Budget FY27", r)
        ws[f"G{r}"] = "=" + ytd_expr("Budget FY27", r)
        ws[f"H{r}"] = f"=E{r}-G{r}"
        ws[f"I{r}"] = f"=IFERROR(H{r}/ABS(G{r}),0)"
        ws[f"K{r}"] = "=" + qtr_expr("P&L FY27", r)
        ws[f"L{r}"] = "=" + qtr_expr("Budget FY27", r)
        ws[f"M{r}"] = f"=K{r}-L{r}"
        ws[f"N{r}"] = f"=IFERROR(M{r}/ABS(L{r}),0)"
        ws[f"P{r}"] = "=E%d-'Actual FY26'!O%d" % (r, r) if False else f"=E{r}-'Actual FY26'!O{r}"
        ws[f"Q{r}"] = f"=G{r}-'Budget FY26'!O{r}"
        ws[f"R{r}"] = f"=P{r}-Q{r}"
        ws[f"S{r}"] = f"=IFERROR(R{r}/ABS(Q{r}),0)"
        ws[f"T{r}"] = risk(r, "H", "I", "E", "G")
        for col in ("B", "C", "D", "E", "F", "G", "H", "K", "L", "M", "P", "Q", "R"):
            ws[f"{col}{r}"].number_format = MONEY
        for col in ("I", "N", "S"):
            ws[f"{col}{r}"].number_format = PCT
        ws[f"T{r}"].alignment = Alignment(horizontal="center")
    return ws


UTIL_DEPTS = ["Onsite", "Production", "Video", "Integration", "Consulting", "CTS"]
HOUR_TYPES = ["Chargeable", "Non-chargeable", "Leave"]
UCOLS = [get_column_letter(4 + i) for i in range(12)]   # D..O


def build_utilisation(wb):
    ws = wb.create_sheet("Utilisation")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Utilisation"
    ws["A1"].font = TITLE
    ws["A2"] = ("Working days and working hours calculate themselves from the calendar and the "
                "public holidays on Lists. Type the hours in the yellow cells.")
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")

    for i, (y, m) in enumerate(MONTHS):
        c = ws.cell(row=4, column=4 + i, value=date(y, m, 1))
        c.number_format = "mmm-yy"
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    ws["A4"] = "Calendar"
    ws["A4"].font = HEAD
    ws["A4"].fill = HEAD_FILL

    ws["A5"] = "Working days"
    ws["A6"] = "Working hours"
    for c in ("A5", "A6"):
        ws[c].font = BOLD
    for i, col in enumerate(UCOLS):
        ws[f"{col}5"] = f"=NETWORKDAYS({col}$4,EOMONTH({col}$4,0),Holidays)"
        ws[f"{col}5"].number_format = "#,##0"
        ws[f"{col}6"] = f"={col}5*Setup!$B$11"
        ws[f"{col}6"].number_format = "#,##0"
        for r in (5, 6):
            ws[f"{col}{r}"].font = BOLD
            ws[f"{col}{r}"].fill = GREY_FILL

    heads = [("A", "FY (year it ends)", 18), ("B", "Department", 16), ("C", "Hour Type", 18)]
    for ref, text, w in heads:
        ws[f"{ref}8"] = text
        ws[f"{ref}8"].font = HEAD
        ws[f"{ref}8"].fill = HEAD_FILL
        ws.column_dimensions[ref].width = w
    for i, (y, m) in enumerate(MONTHS):
        c = ws.cell(row=8, column=4 + i, value=date(y, m, 1))
        c.number_format = "mmm-yy"
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[UCOLS[i]].width = 11
    ws["P8"] = "FY total"
    ws["P8"].font = HEAD
    ws["P8"].fill = HEAD_FILL
    ws.column_dimensions["P"].width = 12
    ws.freeze_panes = "D9"

    r = 9
    for dept in UTIL_DEPTS:
        first = r
        for ht in HOUR_TYPES:
            ws[f"A{r}"] = "=Setup!$B$7"
            ws[f"B{r}"] = dept
            ws[f"C{r}"] = ht
            for col in UCOLS:
                c = ws[f"{col}{r}"]
                c.fill = YELLOW_FILL
                c.font = INPUT_FONT
                c.number_format = "#,##0.0"
            ws[f"P{r}"] = f"=SUM(D{r}:O{r})"
            ws[f"P{r}"].number_format = "#,##0.0"
            for ref in ("A", "B", "C"):
                ws[f"{ref}{r}"].font = BODY
            r += 1
        charge, noncharge = first, first + 1
        ws[f"C{r}"] = "Utilisation %"
        ws[f"C{r}"].font = BOLD
        for col in UCOLS:
            ws[f"{col}{r}"] = f"=IFERROR({col}{charge}/({col}{charge}+{col}{noncharge}),0)"
            ws[f"{col}{r}"].number_format = PCT
            ws[f"{col}{r}"].font = BOLD
            ws[f"{col}{r}"].fill = BAND_FILL
        ws[f"P{r}"] = f"=IFERROR(P{charge}/(P{charge}+P{noncharge}),0)"
        ws[f"P{r}"].number_format = PCT
        for ref in ("A", "B", "C", "P"):
            ws[f"{ref}{r}"].fill = BAND_FILL
        r += 1
        ws[f"C{r}"] = "FTE"
        ws[f"C{r}"].font = BOLD
        for col in UCOLS:
            ws[f"{col}{r}"] = f"=IFERROR(({col}{charge}+{col}{noncharge})/{col}$6,0)"
            ws[f"{col}{r}"].number_format = "0.00"
            ws[f"{col}{r}"].fill = BAND_FILL
        for ref in ("A", "B", "C", "P"):
            ws[f"{ref}{r}"].fill = BAND_FILL
        r += 2
    return ws


def main():
    accounts = load_accounts("/tmp/claude-0/-home-user-Claude/"
                             "a91f56e6-b7da-5dad-8e0b-6d5f516f9360/scratchpad/lists.json")
    subs = subcategories(accounts)
    print(f"accounts {len(accounts)}  subcategories {len(subs)}  departments {len(DEPTS)}")

    wb = Workbook()
    wb.remove(wb.active)

    build_setup(wb)
    _, defs = build_lists(wb, accounts)
    build_gl_paste(wb)
    build_pl_check(wb)
    build_cleanup(wb)
    _, engine_index, _ = build_engine(wb, subs)
    build_paste_sheet(wb, "Budget FY27", subs,
                      "This year's budget or rolling forecast. Paste into the yellow months.")
    build_paste_sheet(wb, "Actual FY26", subs,
                      "Last year's actuals. Paste once, at the start of the year.")
    build_paste_sheet(wb, "Budget FY26", subs,
                      "Last year's budget. Paste once, at the start of the year.")
    build_pl(wb, subs, engine_index)
    build_summary(wb, subs)
    build_utilisation(wb)
    build_readme(wb)

    from openpyxl.workbook.defined_name import DefinedName
    for name, ref in defs.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    yellow = {"Setup", "GL_Paste", "PL_Check", "Budget FY27", "Actual FY26",
              "Budget FY26", "Utilisation"}
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = YELLOW if ws.title in yellow else NAVY

    out = "/home/user/Claude/pack/CTS Financial Controller Pack.xlsx"
    wb.save(out)
    print("saved", out)


if __name__ == "__main__":
    main()
