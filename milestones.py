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
            dy_adj = f'IF(AND("{path}"="FT",{off}{r}=FT_Y5,{dte}{r}<PolicyDate),1,{dy}{r})'
            ms[f"{earned}{r}"] = f'=IF({dte}{r}="","",IF(AND({pth}{r}="{path}",{dte}{r}<=AsOfDate,{dte}{r}>=SchemeStart),{dy_adj},0))'
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
    a["A2"].font = f_note
    heads = [("A", "Name", 12), ("B", "Surname", 13), ("C", "Type today", 11), ("D", "Commencement", 12), ("E", "Days", 7), ("F", "Hours", 7), ("G", "Earned this month (days)", 10), ("H", "How the days were earned (date  pathway  days)", 120)]
    for L, h, w in heads:
        c = a[f"{L}4"]; c.value = h; c.font = f_hdr; c.fill = fill_hdr; c.alignment = center; c.border = border; a.column_dimensions[L].width = w
    E = 2 * ROWS; per = 12; npieces = (E + per - 1) // per
    piece_cols = [get_column_letter(9 + p) for p in range(npieces)]
    hcols = [get_column_letter(9 + npieces + k) for k in range(7)]
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
        a[f"G{r}"] = f'=IF(Staff!$A{sr}="","",Staff!${col["Earned this month (days)"]}{sr})'
        for L, h in helpers:
            a[f"{L}{r}"] = f'=IF(Staff!$A{sr}="","",Staff!${col[h]}{sr})'
        for p, L in enumerate(piece_cols):
            pieces = []
            for e in range(p * per, min((p + 1) * per, E)):
                mr = 20 + e
                d = f'EDATE($D{r},Milestones!$AA${mr})'
                path_at = f'IF(${hn}{r}=0,${hpc}{r},IF(AND(${hd3}{r}<>"",{d}>=${hd3}{r}),${hp3}{r},IF(AND(${hd2}{r}<>"",{d}>=${hd2}{r}),${hp2}{r},${hp1}{r})))'
                dys = f'IF(AND(Milestones!$AB${mr}="FT",Milestones!$AA${mr}=FT_Y5,{d}<PolicyDate),1,Milestones!$AC${mr})'
                pieces.append(f'IF(AND({d}<=AsOfDate,{d}>=SchemeStart,{path_at}=Milestones!$AB${mr}),UPPER(TEXT({d},"dd mmm yy"))&" "&Milestones!$AB${mr}&" "&{dys}&";  ","")')
            a[f"{L}{r}"] = f'=IF(OR($A{r}="",$D{r}=""),"",' + "&".join(pieces) + ")"
        a[f"H{r}"] = f'=IF($A{r}="","",' + "&".join(f"{L}{r}" for L in piece_cols) + ")"
        for L in "ABCDEFGH":
            c = a[f"{L}{r}"]; c.font = f_norm; c.fill = fill_calc; c.border = border
        a[f"H{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    a.freeze_panes = "C5"
    a.auto_filter.ref = f"A4:H{r1}"
    a["A2"] = '="Accruing for "&UPPER(TEXT(AsOfDate,"mmmm yyyy"))&" (as at "&TEXT(AsOfDate,"dd/mm/yyyy")&"). Each entry is: milestone date, pathway on that date (FT or PT), days earned. Only milestones already reached are listed."'
    return a


NP = 20          # people columns on the tally tab
NM = 200         # month rows


def add_tally_check(wb, r0=5, r1=154, tally_rows=None, tally_names=None):
    """'HR Tally Check' tab: paste the HR tally (names in row 7, months down column A from row 8).
    Tracker days per person-month come from a helper sheet built off the merged milestone schedule."""
    ws = wb["Staff"]; ms = wb["Milestones"]
    col = {ws.cell(row=4, column=c).value: get_column_letter(c) for c in range(1, ws.max_column + 1) if ws.cell(row=4, column=c).value}
    keycol = col["h key"]; keyrng = f"Staff!${keycol}${r0}:${keycol}${r1}"
    E = 2 * ROWS
    for n in ("HR Tally Check", "Tally Helper"):
        if n in wb.sheetnames: del wb[n]
    t = wb.create_sheet("HR Tally Check"); t.sheet_properties.tabColor = "ED7D31"
    h = wb.create_sheet("Tally Helper"); h.sheet_properties.tabColor = "7F7F7F"; h.sheet_state = "hidden"

    # ---------- helper: per person 3 rows (date / pathway / earned) x E entries ----------
    h["A1"] = "Per person: milestone dates, pathway on that date, days earned. Built from the merged schedule on the Milestones tab."; h["A1"].font = f_note
    for j, lab in enumerate(["Person", "Key found", "S", "n", "D2", "D3", "P1", "P2", "P3", "Pc", "Row"]):
        h.cell(row=2, column=1 + j, value=lab).font = f_bold
    first_e_col = 12
    for p in range(NP):
        base = 3 + p * 3                 # date row; +1 pathway; +2 earned
        name_cell = f"'HR Tally Check'!{get_column_letter(2 + p)}$7"
        h[f"A{base}"] = f'=IF({name_cell}="","",{name_cell})'
        h[f"B{base}"] = f'=IF(A{base}="","",ISNUMBER(MATCH(A{base},{keyrng},0)))'
        def pull(hd): return f'IF(B{base}<>TRUE,"",INDEX(Staff!${col[hd]}${r0}:${col[hd]}${r1},MATCH($A{base},{keyrng},0)))'
        for L, hd in (("C", "h S"), ("D", "h n dates"), ("E", "h D2"), ("F", "h D3"), ("G", "h P1"), ("H", "h P2"), ("I", "h P3"), ("J", "h Pc")):
            h[f"{L}{base}"] = "=" + pull(hd)
        h[f"C{base}"].number_format = DATE; h[f"E{base}"].number_format = DATE; h[f"F{base}"].number_format = DATE
        h[f"K{base}"] = "date"; h[f"K{base+1}"] = "pathway"; h[f"K{base+2}"] = "earned"
        for e in range(E):
            cL = get_column_letter(first_e_col + e); mr = 20 + e
            d = f"{cL}{base}"
            h[d] = f'=IF($C{base}="","",EDATE($C{base},Milestones!$AA${mr}))'
            h[f"{cL}{base+1}"] = f'=IF({d}="","",IF($D{base}=0,$J{base},IF(AND($F{base}<>"",{d}>=$F{base}),$I{base},IF(AND($E{base}<>"",{d}>=$E{base}),$H{base},$G{base}))))'
            h[f"{cL}{base+2}"] = f'=IF({d}="",0,IF(AND({cL}{base+1}=Milestones!$AB${mr},{d}>=SchemeStart),IF(AND(Milestones!$AB${mr}="FT",Milestones!$AA${mr}=FT_Y5,{d}<PolicyDate),1,Milestones!$AC${mr}),0))'
            h[d].number_format = DATE
    lastE = get_column_letter(first_e_col + E - 1)

    # ---------- tally tab layout ----------
    t.column_dimensions["A"].width = 12
    for p in range(NP): t.column_dimensions[get_column_letter(2 + p)].width = 11
    t["A1"] = "HR TALLY CHECK - paste the HR tally grid here and compare it with the tracker"; t["A1"].font = f_title
    t["A2"] = "Row 7 = names exactly as on the Staff tab (Name Surname). Column A from row 8 = months (Jan-17 style text or a date). Grid = days HR tallied. The tracker's days for the same person and month appear in the TRACKER block to the right, and the DIFFERENCE block after that. Only months listed in column A are compared - add earlier month rows to see milestones before the tally began."; t["A2"].font = f_note
    t["A3"] = "Compare totals up to (month end):"; t["A3"].font = f_bold
    t["B3"] = "=EOMONTH(AsOfDate,-1)"; t["B3"].number_format = DATE; t["B3"].font = f_input; t["B3"].fill = fill_key; t["B3"].border = border
    t["C3"] = "Defaults to the end of last month. Change it to compare up to a different month."; t["C3"].font = f_note
    labels = {4: "HR tally total to that month", 5: "Tracker total to that month", 6: "Difference (tracker - HR)"}
    for r, lab in labels.items(): t[f"A{r}"] = lab; t[f"A{r}"].font = f_bold
    t["A7"] = "Month"; t["A7"].font = f_hdr; t["A7"].fill = fill_hdr; t["A7"].border = border
    # month-start helper in a hidden column (AZ) and month-end (BA)
    mcol, ecol = get_column_letter(2 + 3 * NP + 8), get_column_letter(2 + 3 * NP + 9)
    t.column_dimensions[mcol].hidden = True; t.column_dimensions[ecol].hidden = True
    t[f"{mcol}7"] = "month start"; t[f"{ecol}7"] = "month end"
    months_const = '{"JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"}'
    for i in range(NM):
        r = 8 + i
        t[f"{mcol}{r}"] = (f'=IF($A{r}="","",IF(ISNUMBER($A{r}),DATE(YEAR($A{r}),MONTH($A{r}),1),'
                           f'IFERROR(DATE(2000+VALUE(RIGHT(TRIM($A{r}),2)),MATCH(UPPER(LEFT(TRIM($A{r}),3)),{months_const},0),1),"")))')
        t[f"{ecol}{r}"] = f'=IF({mcol}{r}="","",EOMONTH({mcol}{r},0))'
        t[f"{mcol}{r}"].number_format = DATE; t[f"{ecol}{r}"].number_format = DATE
    # blocks: tracker at column X (24), difference at AP (42) -> use offsets
    tr0, df0 = 2 + NP + 2, 2 + 2 * NP + 4     # 24, 46
    t.cell(row=7, column=tr0 - 1, value="TRACKER").font = f_bold
    t.cell(row=7, column=df0 - 1, value="DIFFERENCE").font = f_bold
    for p in range(NP):
        inL = get_column_letter(2 + p); trL = get_column_letter(tr0 + p); dfL = get_column_letter(df0 + p)
        t.column_dimensions[trL].width = 10; t.column_dimensions[dfL].width = 10
        base = 3 + p * 3
        t[f"{trL}7"] = f'=IF({inL}7="","",{inL}7)'; t[f"{dfL}7"] = f'=IF({inL}7="","",{inL}7)'
        for c in (f"{inL}7", f"{trL}7", f"{dfL}7"):
            t[c].font = f_hdr; t[c].fill = fill_hdr; t[c].border = border; t[c].alignment = center
        t[f"{inL}7"].font = f_input; t[f"{inL}7"].fill = fill_key
        for i in range(NM):
            r = 8 + i
            t[f"{inL}{r}"].fill = fill_key; t[f"{inL}{r}"].font = f_input; t[f"{inL}{r}"].border = border
            t[f"{trL}{r}"] = (f'=IF(OR({inL}$7="",${mcol}{r}=""),"",IF(\'Tally Helper\'!$B${base}<>TRUE,"not on Staff",'
                              f'SUMIFS(\'Tally Helper\'!${get_column_letter(first_e_col)}${base+2}:${lastE}${base+2},'
                              f'\'Tally Helper\'!${get_column_letter(first_e_col)}${base}:${lastE}${base},">="&${mcol}{r},'
                              f'\'Tally Helper\'!${get_column_letter(first_e_col)}${base}:${lastE}${base},"<="&${ecol}{r})))')
            t[f"{dfL}{r}"] = f'=IF(OR({trL}{r}="",NOT(ISNUMBER({trL}{r}))),"",{trL}{r}-N({inL}{r}))'
            for c in (f"{trL}{r}", f"{dfL}{r}"):
                t[c].font = f_norm; c2 = t[c]; c2.fill = fill_calc; c2.border = border; c2.number_format = 'General;-General;"-"'
        # totals
        t[f"{inL}4"] = f'=IF({inL}7="","",SUMIFS({inL}8:{inL}{7+NM},${ecol}8:${ecol}{7+NM},"<="&$B$3))'
        t[f"{inL}5"] = f'=IF({inL}7="","",IF(\'Tally Helper\'!$B${base}<>TRUE,"not on Staff",SUMIFS({trL}8:{trL}{7+NM},${ecol}8:${ecol}{7+NM},"<="&$B$3)))'
        t[f"{inL}6"] = f'=IF(OR({inL}5="",NOT(ISNUMBER({inL}5))),"",{inL}5-N({inL}4))'
        for r in (4, 5, 6):
            t[f"{inL}{r}"].font = f_bold; t[f"{inL}{r}"].fill = fill_calc; t[f"{inL}{r}"].border = border; t[f"{inL}{r}"].number_format = 'General;-General;"-"'
    dfl0 = get_column_letter(df0); dfl1 = get_column_letter(df0 + NP - 1)
    t.conditional_formatting.add(f"{dfl0}8:{dfl1}{7+NM}", FormulaRule(formula=[f'AND(ISNUMBER({dfl0}8),{dfl0}8<>0)'], fill=PatternFill("solid", fgColor="FFC7CE")))
    t.conditional_formatting.add(f"B6:{get_column_letter(1+NP)}6", FormulaRule(formula=['AND(ISNUMBER(B6),B6<>0)'], fill=PatternFill("solid", fgColor="FFC7CE")))
    t.conditional_formatting.add(f"B5:{get_column_letter(1+NP)}5", FormulaRule(formula=['B5="not on Staff"'], fill=PatternFill("solid", fgColor="FFC7CE")))
    t.freeze_panes = "B8"
    # optional preload
    if tally_names:
        for p, nm in enumerate(tally_names): t[f"{get_column_letter(2+p)}7"] = nm
    if tally_rows:
        for i, (lab, vals) in enumerate(tally_rows):
            t[f"A{8+i}"] = lab
            for p, v in enumerate(vals):
                if v not in (None, ""): t[f"{get_column_letter(2+p)}{8+i}"] = v
    return t
