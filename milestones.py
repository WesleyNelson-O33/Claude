"""Add a 'Milestones' tab to a tracker workbook (works on both the master and the HR copy, by header lookup)."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

ROWS = 90
F = "Arial"
f_norm = Font(name=F, size=10); f_bold = Font(name=F, size=10, bold=True); f_title = Font(name=F, size=14, bold=True)
f_hdr = Font(name=F, size=10, bold=True, color="FFFFFF"); f_input = Font(name=F, size=10, color="0000FF"); f_note = Font(name=F, size=9, italic=True, color="555555")
fill_hdr = PatternFill("solid", fgColor="1F4E78"); fill_hdr2 = PatternFill("solid", fgColor="2E75B6")
fill_key = PatternFill("solid", fgColor="FFFF00"); fill_calc = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF"); border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATE = "dd/mm/yyyy"


def add_milestones(wb, r0=5, r1=154):
    ws = wb["Staff"]
    col = {ws.cell(row=4, column=c).value: get_column_letter(c) for c in range(1, ws.max_column + 1) if ws.cell(row=4, column=c).value}
    need = ["Name", "Surname", "EmploymentType", "PT Hrs/Week", "Days", "h S", "h n dates", "h D2", "h D3", "h T1", "h T2", "h T3"]
    missing = [n for n in need if n not in col]
    if missing:
        raise SystemExit(f"Staff headers not found: {missing}")
    # key helper column (first empty column after the last header)
    keycol = get_column_letter(ws.max_column + 1)
    ws[f"{keycol}4"] = "h key"; ws[f"{keycol}4"].font = f_hdr; ws[f"{keycol}4"].fill = PatternFill("solid", fgColor="7F7F7F")
    for r in range(r0, r1 + 1):
        ws[f"{keycol}{r}"] = f'=IF(LEN($A{r}&$B{r})=0,"",$A{r}&" "&$B{r})'
    ws.column_dimensions[keycol].hidden = True
    keyrng = f"Staff!${keycol}${r0}:${keycol}${r1}"

    if "Milestones" in wb.sheetnames:
        del wb["Milestones"]
    ms = wb.create_sheet("Milestones")
    ms.sheet_properties.tabColor = "7030A0"
    for k, v in {"A": 2, "B": 30, "C": 16, "D": 2, "E": 30, "F": 16, "G": 2, "H": 4, "I": 9, "J": 11, "K": 12, "L": 13, "M": 10, "N": 9, "O": 9, "P": 2, "Q": 4, "R": 9, "S": 11, "T": 12, "U": 13, "V": 10, "W": 9, "X": 9}.items():
        ms.column_dimensions[k].width = v
    ms["B2"] = "MILESTONES - one person, accrual by accrual"; ms["B2"].font = f_title
    ms["B3"] = "Pick a name in the yellow cell. Each milestone date is listed with the employment type on that date and the days it earned. Earned days add up to the Days column on the Staff tab."; ms["B3"].font = f_note
    ms["B5"] = "Staff member"; ms["B5"].font = f_bold
    ms["C5"] = f"=Staff!{keycol}{r0}"; ms["C5"].font = f_input; ms["C5"].fill = fill_key; ms["C5"].border = border
    dv = DataValidation(type="list", formula1=f"={keyrng}", allow_blank=True); ms.add_data_validation(dv); dv.add("C5")
    ms["B6"] = "Found?"; ms["B6"].font = f_norm
    ms["C6"] = f'=IF(ISNUMBER(MATCH($C$5,{keyrng},0)),"Yes","NO - pick from the list")'; ms["C6"].fill = fill_calc; ms["C6"].border = border

    def pull(h):
        return f'IF($C$6<>"Yes","",INDEX(Staff!${col[h]}${r0}:${col[h]}${r1},MATCH($C$5,{keyrng},0)))'

    left = [(8, "Commencement for Bonus Leave", "h S", DATE), (9, "Employment type today", "EmploymentType", None), (10, "PT hours per week (blank = 24+)", "PT Hrs/Week", None),
            (11, "Type at commencement", "h T1", None), (12, "Change 1 date", "h D2", DATE), (13, "Change 1 new type", "h T2", None),
            (14, "Change 2 date", "h D3", DATE), (15, "Change 2 new type", "h T3", None), (16, "Number of type dates entered", "h n dates", None)]
    for r, label, h, fmt in left:
        ms[f"B{r}"] = label; ms[f"B{r}"].font = f_norm; ms[f"B{r}"].border = border
        c = ms[f"C{r}"]; c.value = "=" + pull(h); c.font = f_norm; c.fill = fill_calc; c.border = border
        if fmt: c.number_format = fmt
    right = [(8, "As-of date", "=AsOfDate", DATE), (9, "Days per policy (Staff tab)", "=" + pull("Days"), None),
             (10, "Days earned in the tables below", f"=IF($C$6<>\"Yes\",\"\",SUM($O$20:$O${19+ROWS})+SUM($X$20:$X${19+ROWS}))", None),
             (11, "Reconciles?", '=IF($C$6<>"Yes","",IF(ABS(N(F9)-N(F10))<0.001,"Yes","NO - check inputs"))', None),
             (12, "Hours per day", "=HoursPerDay", None), (13, "Hours earned to date", "=IF($C$6<>\"Yes\",\"\",F10*HoursPerDay)", None)]
    for r, label, formula, fmt in right:
        ms[f"E{r}"] = label; ms[f"E{r}"].font = f_norm; ms[f"E{r}"].border = border
        c = ms[f"F{r}"]; c.value = formula; c.font = f_bold if r in (9, 10, 11, 13) else f_norm; c.fill = fill_calc; c.border = border
        if fmt: c.number_format = fmt

    S, T1, D2, T2, D3, T3, N, TYPE, HRS = "$C$8", "$C$11", "$C$12", "$C$13", "$C$14", "$C$15", "$C$16", "$C$9", "$C$10"
    heads = ["#", "Months", "Milestone", "Date", "Type on that date", "Pathway", "Days if pathway matches", "Days earned"]
    ms.merge_cells("H18:O18"); ms["H18"] = "FULL-TIME milestones (earned only if full-time on that date)"
    ms.merge_cells("Q18:X18"); ms["Q18"] = "PART-TIME 24+ hrs milestones (earned only if eligible part-time on that date)"
    for c in ("H18", "Q18"):
        ms[c].font = f_hdr; ms[c].fill = fill_hdr2; ms[c].alignment = center
    for j, h in enumerate(heads):
        for base in (8, 17):
            cell = ms.cell(row=19, column=base + j, value=h); cell.font = f_hdr; cell.fill = fill_hdr; cell.alignment = center; cell.border = border
    ms.row_dimensions[19].height = 40

    def isFT(t): return f'OR(ISNUMBER(SEARCH("full",{t})),UPPER(TRIM({t}))="FT",UPPER(TRIM({t}))="FTE")'
    def isPT(t): return f'OR(ISNUMBER(SEARCH("part",{t})),UPPER(TRIM({t}))="PT",UPPER(TRIM({t}))="PTE")'

    for base, path in ((8, "FT"), (17, "PT")):
        n, off, ml, dte, typ, pth, dy, earned = [get_column_letter(base + k) for k in range(8)]
        for i in range(ROWS):
            r = 20 + i
            ms[f"{n}{r}"] = i + 1
            if path == "FT":
                ms[f"{off}{r}"] = f'=IF({i}<FT_PreY5,FT_First+FT_Int1*{i},FT_Y5+FT_Int2*({i}-FT_PreY5))'
                ms[f"{dy}{r}"] = f'=IF({off}{r}<FT_Y5,1,IF(MOD({off}{r}-FT_Y5,12)=0,FT_AnnivDays,1))'
            else:
                ms[f"{off}{r}"] = f'=PT_First+PT_Int*{i}'
                ms[f"{dy}{r}"] = 1
            ms[f"{ml}{r}"] = f'=INT({off}{r}/12)&"y "&MOD({off}{r},12)&"m"'
            ms[f"{dte}{r}"] = f'=IF(OR($C$6<>"Yes",{S}=""),"",EDATE({S},{off}{r}))'
            ms[f"{typ}{r}"] = (f'=IF({dte}{r}="","",IF({N}=0,{TYPE},IF(AND({D3}<>"",{dte}{r}>={D3}),{T3},IF(AND({D2}<>"",{dte}{r}>={D2}),{T2},{T1}))))')
            ms[f"{pth}{r}"] = (f'=IF({dte}{r}="","",IF({isFT(typ + str(r))},"FT",IF(AND({isPT(typ + str(r))},OR({HRS}="",N({HRS})>=PT_MinHours)),"PT","None")))')
            ms[f"{earned}{r}"] = f'=IF({dte}{r}="","",IF(AND({pth}{r}="{path}",{dte}{r}<=AsOfDate),{dy}{r},0))'
            for L in (n, off, ml, dte, typ, pth, dy, earned):
                c = ms[f"{L}{r}"]; c.font = f_norm; c.fill = fill_calc; c.border = border
            ms[f"{dte}{r}"].number_format = DATE
        ms.conditional_formatting.add(f"{n}20:{earned}{19+ROWS}", FormulaRule(formula=[f'N(${earned}20)>0'], fill=PatternFill("solid", fgColor="C6EFCE")))
        ms.conditional_formatting.add(f"{n}20:{earned}{19+ROWS}", FormulaRule(formula=[f'AND(${dte}20<>"",${dte}20>AsOfDate)'], font=Font(name=F, size=10, color="999999")))
    ms.freeze_panes = "A20"
    return ms


def add_all_staff(wb, r0=5, r1=154):
    """One line per person listing every earned milestone, chronological. Needs add_milestones() first."""
    ws = wb["Staff"]; ms = wb["Milestones"]
    col = {ws.cell(row=4, column=c).value: get_column_letter(c) for c in range(1, ws.max_column + 1) if ws.cell(row=4, column=c).value}
    # merged, sorted schedule helper on the Milestones tab: Z = all offsets, AA = sorted, AB = pathway, AC = days
    ms["Z18"] = "Merged schedule (helper - do not edit)"; ms["Z18"].font = f_note
    for j, h in enumerate(["All offsets", "Sorted", "Pathway", "Days"]):
        c = ms.cell(row=19, column=26 + j, value=h); c.font = f_hdr; c.fill = PatternFill("solid", fgColor="7F7F7F"); c.border = border
    for i in range(2 * ROWS):
        r = 20 + i
        ms[f"Z{r}"] = f"=I{20+i}" if i < ROWS else f"=R{20+i-ROWS}"
        ms[f"AA{r}"] = f"=SMALL($Z$20:$Z${19+2*ROWS},{i+1})"
        ms[f"AB{r}"] = f'=IF({i+1}-COUNTIF($Z$20:$Z${19+2*ROWS},"<"&AA{r})=1,IF(COUNTIF($I$20:$I${19+ROWS},AA{r})>0,"FT","PT"),"PT")'
        ms[f"AC{r}"] = f'=IF(AB{r}="FT",INDEX($N$20:$N${19+ROWS},MATCH(AA{r},$I$20:$I${19+ROWS},0)),1)'
        for L in ("Z", "AA", "AB", "AC"):
            ms[f"{L}{r}"].font = f_note; ms[f"{L}{r}"].fill = fill_calc
    for L in ("Z", "AA", "AB", "AC"): ms.column_dimensions[L].width = 9
    ms.column_dimensions.group("Z", "AC", hidden=True, outline_level=1)

    if "All Staff Breakdown" in wb.sheetnames:
        del wb["All Staff Breakdown"]
    a = wb.create_sheet("All Staff Breakdown"); a.sheet_properties.tabColor = "7030A0"
    a["A1"] = "ALL STAFF - how each person's days were earned"; a["A1"].font = f_title
    a["A2"] = '="As at "&TEXT(AsOfDate,"dd/mm/yyyy")&". Each entry is: milestone date, pathway on that date (FT or PT), days earned. Only milestones already reached are listed. Casual time and milestones that fell while casual or under 24 hrs earn nothing and are not shown."'
    a["A2"].font = f_note
    heads = [("A", "Name", 12), ("B", "Surname", 13), ("C", "Type today", 11), ("D", "Commencement", 12), ("E", "Days", 7), ("F", "Hours", 7), ("G", "How the days were earned (date  pathway  days)", 120)]
    for L, h, w in heads:
        c = a[f"{L}4"]; c.value = h; c.font = f_hdr; c.fill = fill_hdr; c.alignment = center; c.border = border; a.column_dimensions[L].width = w
    E = 2 * ROWS; per = 20; npieces = (E + per - 1) // per
    piece_cols = [get_column_letter(8 + p) for p in range(npieces)]
    hcols = [get_column_letter(8 + npieces + k) for k in range(7)]
    helpers = list(zip(hcols, ["h n dates", "h D2", "h D3", "h P1", "h P2", "h P3", "h Pc"]))
    for L in piece_cols:
        c = a[f"{L}4"]; c.value = "piece"; c.font = f_hdr; c.fill = PatternFill("solid", fgColor="7F7F7F")
    for L, h in helpers:
        c = a[f"{L}4"]; c.value = h; c.font = f_hdr; c.fill = PatternFill("solid", fgColor="7F7F7F")
    a.column_dimensions.group(piece_cols[0], hcols[-1], hidden=True, outline_level=1)
    hn, hd2, hd3, hp1, hp2, hp3, hpc = hcols
    for r in range(r0, r1 + 1):
        sr = r
        a[f"A{r}"] = f'=IF(Staff!$A{sr}="","",Staff!$A{sr})'
        a[f"B{r}"] = f'=IF(Staff!$A{sr}="","",Staff!$B{sr})'
        a[f"C{r}"] = f'=IF(Staff!$A{sr}="","",Staff!${col["EmploymentType"]}{sr})'
        a[f"D{r}"] = f'=IF(Staff!$A{sr}="","",Staff!${col["h S"]}{sr})'; a[f"D{r}"].number_format = DATE
        a[f"E{r}"] = f'=IF(Staff!$A{sr}="","",Staff!${col["Days"]}{sr})'
        a[f"F{r}"] = f'=IF(E{r}="","",E{r}*HoursPerDay)'
        for L, h in helpers:
            a[f"{L}{r}"] = f'=IF(Staff!$A{sr}="","",Staff!${col[h]}{sr})'
        for p, L in enumerate(piece_cols):
            pieces = []
            for e in range(p * per, min((p + 1) * per, E)):
                mr = 20 + e
                d = f'EDATE($D{r},Milestones!$AA${mr})'
                path_at = f'IF(${hn}{r}=0,${hpc}{r},IF(AND(${hd3}{r}<>"",{d}>=${hd3}{r}),${hp3}{r},IF(AND(${hd2}{r}<>"",{d}>=${hd2}{r}),${hp2}{r},${hp1}{r})))'
                pieces.append(f'IF(AND({d}<=AsOfDate,{path_at}=Milestones!$AB${mr}),TEXT({d},"dd/mm/yy")&" "&Milestones!$AB${mr}&" "&Milestones!$AC${mr}&";  ","")')
            a[f"{L}{r}"] = f'=IF(OR($A{r}="",$D{r}=""),"",' + "&".join(pieces) + ")"
        a[f"G{r}"] = f'=IF($A{r}="","",' + "&".join(f"{L}{r}" for L in piece_cols) + ")"
        for L in "ABCDEFG":
            c = a[f"{L}{r}"]; c.font = f_norm; c.fill = fill_calc; c.border = border
        a[f"G{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    a.freeze_panes = "C5"
    a.auto_filter.ref = f"A4:G{r1}"
    return a
