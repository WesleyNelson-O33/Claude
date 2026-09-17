def build_actual_py(wb, subs, engine_index):
    """Last year's actuals, straight off the same GL paste. Paste FY26 into
    GL_Paste alongside FY27 and this fills itself."""
    ws = wb.create_sheet("Actual FY26")
    titles(ws, "Calculated. Paste last year's GL into GL_Paste with this year's and this fills itself.")
    header(ws, [(FY_COL, "FY total")],
           widths=[(c, 12) for c in MONTH_COLS] + [(FY_COL, 13)])
    for i, (y, m) in enumerate(MONTHS_PY):
        c = ws.cell(row=5, column=2 + i, value=date(y, m, 1))
        c.number_format = "mmm-yy"
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    rows = layout(subs)
    paint(ws, rows, 14)
    info = row_map(rows)
    for i, (kind, _, dept, cat, sub) in enumerate(rows):
        r = ROW0 + i
        if kind != "sub":
            continue
        eng = engine_index[(dept, sub)]
        for j, col in enumerate(MONTH_COLS):
            ws[f"{col}{r}"] = f"=Engine!{get_column_letter(5 + j)}{eng}"
    roll_up(ws, info, MONTH_COLS)
    for i, (kind, *_x) in enumerate(rows):
        r = ROW0 + i
        if kind == "blank":
            continue
        ws[f"{FY_COL}{r}"] = f"=SUM(B{r}:M{r})"
        for col in MONTH_COLS + [FY_COL]:
            ws[f"{col}{r}"].number_format = MONEY
    return ws


def build_clients(wb):
    """Revenue by client off the Contact column: month, quarter, year, and the
    same periods a year earlier."""
    ws = wb.create_sheet("Clients")
    style_title(ws, "Top clients",
                "Revenue by client, from the Contact column on the GL. Type the client names in "
                "column B exactly as they read in Xero; everything to the right calculates.")
    ws.sheet_view.showGridLines = False
    last = GL_ROWS + 4
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 42
    for i in range(9):
        ws.column_dimensions[get_column_letter(3 + i)].width = 15

    heads = ["Rank", "Client", "Month", "Prior month", "MOM $", "Quarter",
             "Prior quarter", "QOQ $", "YTD", "Last year YTD", "YOY $"]
    for i, h in enumerate(heads):
        c = ws.cell(row=5, column=1 + i, value=h)
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.freeze_panes = "C6"

    qs = "DATE(YEAR(Setup!$B$5),FLOOR(MONTH(Setup!$B$5)-1,3)+1,1)"
    for i in range(30):
        r = 6 + i
        ws[f"A{r}"] = f'=IF($B{r}="","",RANK($I{r},$I$6:$I$35))'
        c = ws[f"B{r}"]
        c.fill = YELLOW_FILL
        c.font = INPUT_FONT
        base = (f'SUMIFS(Cleanup!$A$5:$A${last},Cleanup!$F$5:$F${last},$B{r},'
                f'Cleanup!$D$5:$D${last},"Income"')
        ws[f"C{r}"] = f'=IF($B{r}="","",{base},Cleanup!$B$5:$B${last},Setup!$B$5))'
        ws[f"D{r}"] = f'=IF($B{r}="","",{base},Cleanup!$B$5:$B${last},EDATE(Setup!$B$5,-1)))'
        ws[f"E{r}"] = f'=IF($B{r}="","",C{r}-D{r})'
        ws[f"F{r}"] = (f'=IF($B{r}="","",{base},Cleanup!$B$5:$B${last},">="&{qs},'
                       f'Cleanup!$B$5:$B${last},"<="&Setup!$B$5))')
        ws[f"G{r}"] = (f'=IF($B{r}="","",{base},Cleanup!$B$5:$B${last},">="&EDATE({qs},-3),'
                       f'Cleanup!$B$5:$B${last},"<"&{qs}))')
        ws[f"H{r}"] = f'=IF($B{r}="","",F{r}-G{r})'
        ws[f"I{r}"] = (f'=IF($B{r}="","",{base},Cleanup!$B$5:$B${last},">="&DATE(Setup!$B$7-1,7,1),'
                       f'Cleanup!$B$5:$B${last},"<="&Setup!$B$5))')
        ws[f"J{r}"] = (f'=IF($B{r}="","",{base},Cleanup!$B$5:$B${last},">="&DATE(Setup!$B$7-2,7,1),'
                       f'Cleanup!$B$5:$B${last},"<="&EDATE(Setup!$B$5,-12)))')
        ws[f"K{r}"] = f'=IF($B{r}="","",I{r}-J{r})'
        for col in "CDEFGHIJK":
            ws[f"{col}{r}"].number_format = MONEY
            ws[f"{col}{r}"].font = BODY

    ws["A40"] = "Top clients by department  -  year to date"
    ws["A40"].font = BOLD
    ws["B41"] = "Client"
    ws["B41"].font = HEAD
    ws["B41"].fill = HEAD_FILL
    for i, dept in enumerate(DEPTS):
        c = ws.cell(row=41, column=3 + i, value=dept)
        c.font = HEAD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    for i in range(30):
        r = 42 + i
        ws[f"B{r}"] = f'=IF($B{6 + i}="","",$B{6 + i})'
        ws[f"B{r}"].font = BODY
        for j, dept in enumerate(DEPTS):
            col = get_column_letter(3 + j)
            ws[f"{col}{r}"] = (
                f'=IF($B{r}="","",SUMIFS(Cleanup!$A$5:$A${last},'
                f'Cleanup!$F$5:$F${last},$B{r},Cleanup!$D$5:$D${last},"Income",'
                f'Cleanup!$C$5:$C${last},"{dept}",'
                f'Cleanup!$B$5:$B${last},">="&DATE(Setup!$B$7-1,7,1),'
                f'Cleanup!$B$5:$B${last},"<="&Setup!$B$5))')
            ws[f"{col}{r}"].number_format = MONEY
            ws[f"{col}{r}"].font = BODY
    return ws


PL_LAST = 232


def build_chart_data(wb):
    """The series behind every chart, so the Charts sheet stays clean."""
    ws = wb.create_sheet("Chart_Data")
    style_title(ws, "Chart data  -  calculated, never type here",
                "Each block is one chart on the Charts sheet.")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    for i in range(len(MONTHS)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 11

    def months_row(r):
        c = ws.cell(row=r, column=1, value="Month")
        c.font = HEAD
        c.fill = HEAD_FILL
        for i, (y, m) in enumerate(MONTHS):
            c = ws.cell(row=r, column=2 + i, value=date(y, m, 1))
            c.number_format = "mmm-yy"
            c.font = HEAD
            c.fill = HEAD_FILL

    def cat_sum(sheet, cat, col):
        return (f"SUMIF('{sheet}'!$A$6:$A${PL_LAST},\"    {cat}\","
                f"'{sheet}'!{col}$6:{col}${PL_LAST})")

    ws["A4"] = "1.  Revenue versus profit"
    ws["A4"].font = BOLD
    months_row(5)
    for j, label in enumerate(("Revenue", "Gross profit", "Net profit")):
        r = 6 + j
        ws.cell(row=r, column=1, value=label).font = BODY
        for i, col in enumerate(MONTH_COLS):
            if label == "Revenue":
                f = "=" + cat_sum("P&L FY27", "Income", col)
            elif label == "Gross profit":
                f = "=" + cat_sum("P&L FY27", "Income", col) + "+" + \\
                    cat_sum("P&L FY27", "Cost of Sales", col)
            else:
                f = "=" + "+".join(cat_sum("P&L FY27", c, col) for c in CATEGORY_ORDER)
            ws.cell(row=r, column=2 + i, value=f).number_format = MONEY

    ws["A10"] = "2.  Budget versus actual  -  net profit"
    ws["A10"].font = BOLD
    months_row(11)
    for j, (label, sheet) in enumerate((("Actual", "P&L FY27"), ("Budget", "Budget FY27"))):
        r = 12 + j
        ws.cell(row=r, column=1, value=label).font = BODY
        for i, col in enumerate(MONTH_COLS):
            f = "=" + "+".join(cat_sum(sheet, c, col) for c in CATEGORY_ORDER)
            ws.cell(row=r, column=2 + i, value=f).number_format = MONEY

    ws["A16"] = "3.  Utilisation month on month"
    ws["A16"].font = BOLD
    months_row(17)
    for j, dept in enumerate(UTIL_DEPTS):
        r = 18 + j
        ws.cell(row=r, column=1, value=dept).font = BODY
        for i, col in enumerate(UCOLS):
            ws.cell(row=r, column=2 + i, value=(
                f'=IFERROR(SUMIFS(Utilisation!{col}$9:{col}$60,Utilisation!$B$9:$B$60,$A{r},'
                f'Utilisation!$C$9:$C$60,"Chargeable")/'
                f'SUMIFS(Utilisation!{col}$9:{col}$60,Utilisation!$B$9:$B$60,$A{r},'
                f'Utilisation!$C$9:$C$60,"<>Leave"),0)')).number_format = PCT

    ws["A26"] = "4.  Days sales outstanding"
    ws["A26"].font = BOLD
    ws["C26"] = ("Trade receivables is not in a P&L export, so type the closing balance "
                 "each month in the yellow row. DSO then calculates.")
    ws["C26"].font = Font(name=FONT, size=9, italic=True, color="C00000")
    months_row(27)
    for label, r in (("Trade receivables (closing)", 28), ("Revenue", 29), ("DSO (days)", 30)):
        ws.cell(row=r, column=1, value=label).font = BOLD if r == 30 else BODY
    for i, col in enumerate(MONTH_COLS):
        c = ws.cell(row=28, column=2 + i)
        c.fill = YELLOW_FILL
        c.font = INPUT_FONT
        c.number_format = MONEY
        cl = get_column_letter(2 + i)
        ws.cell(row=29, column=2 + i, value=f"={cl}6").number_format = MONEY
        ws.cell(row=30, column=2 + i,
                value=f"=IFERROR({cl}28/{cl}29*DAY(EOMONTH({cl}$27,0)),0)").number_format = "#,##0"
    return ws


def build_charts(wb):
    from openpyxl.chart import BarChart, LineChart, Reference

    ws = wb.create_sheet("Charts")
    style_title(ws, "Charts",
                "Every series is a formula off the pack, so nothing is copied across each month.")
    ws.sheet_view.showGridLines = False
    cd = wb["Chart_Data"]
    ncol = 1 + len(MONTHS)

    def monthly(chart, title, first, lastrow, catrow, anchor, pct=False):
        chart.title = title
        chart.height, chart.width = 8, 19
        data = Reference(cd, min_col=1, max_col=ncol, min_row=first, max_row=lastrow)
        cats = Reference(cd, min_col=2, max_col=ncol, min_row=catrow, max_row=catrow)
        chart.add_data(data, titles_from_data=True, from_rows=True)
        chart.set_categories(cats)
        if pct:
            chart.y_axis.numFmt = "0%"
        ws.add_chart(chart, anchor)

    c = BarChart(); c.type = "col"; c.grouping = "clustered"
    monthly(c, "Revenue versus profit", 6, 8, 5, "A4")
    c = BarChart(); c.type = "col"; c.grouping = "clustered"
    monthly(c, "Budget versus actual  -  net profit", 12, 13, 11, "A21")
    c = LineChart()
    monthly(c, "Utilisation month on month", 18, 23, 17, "A38", pct=True)
    c = LineChart()
    monthly(c, "Days sales outstanding", 30, 30, 27, "A55")

    cl = wb["Clients"]
    def clients(chart, title, c1, c2, anchor, r1=5, r2=15, cr1=6, cr2=15, col=2):
        chart.title = title
        chart.height, chart.width = 10, 19
        chart.add_data(Reference(cl, min_col=c1, max_col=c2, min_row=r1, max_row=r2),
                       titles_from_data=True)
        chart.set_categories(Reference(cl, min_col=col, min_row=cr1, max_row=cr2))
        ws.add_chart(chart, anchor)

    c = BarChart(); c.type = "bar"; c.grouping = "clustered"
    clients(c, "Top 10 clients  -  year to date against last year", 9, 10, "L4")
    c = BarChart(); c.type = "col"; c.grouping = "clustered"
    clients(c, "Top 10 clients  -  month on month", 3, 4, "L24")
    c = BarChart(); c.type = "col"; c.grouping = "clustered"
    clients(c, "Top 10 clients  -  quarter on quarter", 6, 7, "L44")
    c = BarChart(); c.type = "col"; c.grouping = "stacked"; c.overlap = 100
    clients(c, "Top clients by department  -  year to date",
            3, 2 + len(DEPTS), "L64", r1=41, r2=51, cr1=42, cr2=51)
    return ws


