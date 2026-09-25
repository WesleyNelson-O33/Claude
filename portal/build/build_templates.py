"""Build the monthly data templates for the CTS Business Intelligence Portal.

Each template is shaped to the export it receives, header for header, so the
monthly job is a straight paste. The Build page in the portal reads these
files from the templates folder and writes the data files the portal runs on.

Three templates are placeholders until a real export arrives: Zoho Deals,
OnRent Orders and Qwilr Quotes. Their columns are the fields the portal needs,
laid out on the system's standard export where one is known, and each says so
on its Read me tab.
"""
import json
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path("/home/user/Claude")
OUT = ROOT / "portal/templates"
ACCOUNTS = json.loads((ROOT / "reporting/data/accounts.json").read_text())

ARIAL = "Arial"
NAVY = "181818"           # the brand charcoal
TEAL = "14A3A3"
INPUT_FONT = Font(name=ARIAL, size=10, color="0000FF")
BODY = Font(name=ARIAL, size=10)
BOLD = Font(name=ARIAL, size=10, bold=True)
HEAD = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
TITLE = Font(name=ARIAL, size=16, bold=True, color=NAVY)
SUB = Font(name=ARIAL, size=9, italic=True, color="595959")
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
YELLOW = PatternFill("solid", fgColor="FFFF00")
PASTE = PatternFill("solid", fgColor="FFF2CC")
GREY = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '$#,##0.00;($#,##0.00);-'
PCT = '0.0%'

FY = 27
MONTHS_FY27 = [date(2026, m, 1) for m in range(7, 13)] + [date(2027, m, 1) for m in range(1, 7)]
MONTHS_FY26 = [date(2025, m, 1) for m in range(7, 13)] + [date(2026, m, 1) for m in range(1, 7)]


def mlabel(d):
    return d.strftime("%b-%y")


# ------------------------------------------------------------- building blocks
def readme(wb, title, lines, placeholder=False):
    """A Read me tab: what the file is, what to paste, what not to touch."""
    ws = wb.create_sheet("Read me", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 108
    ws["B2"] = title
    ws["B2"].font = TITLE
    ws["B3"] = "CTS Business Intelligence Portal  |  monthly data template"
    ws["B3"].font = SUB
    r = 5
    if placeholder:
        ws[f"B{r}"] = ("PLACEHOLDER. The columns on the data tab are what the portal needs, laid out on "
                       "the system's standard export where one is known. Send one real export and the "
                       "tab will be matched to it exactly, so the monthly paste needs no reshaping.")
        ws[f"B{r}"].font = Font(name=ARIAL, size=10, bold=True, color="C00000")
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 44
        r += 2
    for kind, text in lines:
        c = ws[f"B{r}"]
        c.value = text
        if kind == "h":
            c.font = Font(name=ARIAL, size=12, bold=True, color=NAVY)
        else:
            c.font = BODY
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = max(15, 15 * (1 + len(text) // 105))
        r += 1
    r += 1
    ws[f"B{r}"] = "Legend"
    ws[f"B{r}"].font = Font(name=ARIAL, size=12, bold=True, color=NAVY)
    r += 1
    for fill, font, text in ((PASTE, INPUT_FONT, "Cream cells with blue text: paste or type here."),
                             (YELLOW, INPUT_FONT, "Yellow cells: a setting you choose."),
                             (GREY, BODY, "Grey cells: calculated or fixed. Do not type here.")):
        ws[f"B{r}"] = text
        ws[f"B{r}"].font = font
        ws[f"B{r}"].fill = fill
        r += 1
    return ws


def data_tab(wb, name, headers, widths, example, note, paste_rows=2000, title_row=None):
    """A paste target: the export's own headers across row 1 (or row 2 under a
    title row, where the export has one), one example row, then cream cells."""
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    hr = 1
    if title_row:
        ws["A1"] = title_row
        ws["A1"].font = BOLD
        hr = 2
    for i, hd in enumerate(headers, start=1):
        c = ws.cell(row=hr, column=i, value=hd)
        c.font = HEAD
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1] if i - 1 < len(widths) else 14
    ws.row_dimensions[hr].height = 30
    for i, v in enumerate(example, start=1):
        c = ws.cell(row=hr + 1, column=i, value=v)
        c.font = INPUT_FONT
        c.fill = PASTE
        c.border = BOX
        if isinstance(v, date):
            c.number_format = "dd/mm/yyyy"
        elif isinstance(v, float):
            c.number_format = "#,##0.00"
    ws.cell(row=hr + 1, column=1).comment = Comment(
        "Example row showing the expected format. Paste over it; it is not read as data "
        "unless you leave it here.", "CTS Portal")
    for r in range(hr + 2, hr + 2 + paste_rows):
        for i in range(1, len(headers) + 1):
            c = ws.cell(row=r, column=i)
            c.fill = PASTE
            c.font = INPUT_FONT
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(headers))}{hr + 1 + paste_rows}"
    if note:
        ws.cell(row=hr, column=len(headers) + 2, value=note).font = SUB
    return ws


def list_tab(wb, name, headers, rows, widths, editable_cols=None, note=None):
    """A reference list you maintain by hand: header, rows, cream where editable."""
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    for i, hd in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=hd)
        c.font = HEAD
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1] if i - 1 < len(widths) else 14
    ws.row_dimensions[1].height = 30
    editable = set(editable_cols or range(1, len(headers) + 1))
    for r, row in enumerate(rows, start=2):
        for i, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.border = BOX
            if i in editable:
                c.font = INPUT_FONT
                c.fill = PASTE
            else:
                c.font = BODY
                c.fill = GREY
            if isinstance(v, float) and 0 <= v <= 1 and "%" in headers[i - 1]:
                c.number_format = PCT
            if isinstance(v, date):
                c.number_format = "dd/mm/yyyy"
    ws.freeze_panes = "A2"
    if note:
        ws.cell(row=1, column=len(headers) + 2, value=note).font = SUB
    return ws


def save(wb, name):
    wb.remove(wb["Sheet"]) if "Sheet" in wb.sheetnames else None
    p = OUT / name
    wb.save(p)
    print("  %-40s %5.0f KB  tabs: %s" % (name, p.stat().st_size / 1024, ", ".join(wb.sheetnames)))


OUT.mkdir(parents=True, exist_ok=True)
print("templates:")

# ------------------------------------------------------------ 01 Xero P&L
wb = Workbook()
readme(wb, "01 Xero Profit and Loss", [
    ("h", "What this is"),
    ("p", "The profit and loss by account by month. This is what the P&L pages read and what the "
          "P&L Control page checks the ledger against."),
    ("h", "Each month"),
    ("p", "1. In Xero run Profit and Loss for the financial year to date, one column per month, "
          "accrual basis. Export to Excel."),
    ("p", "2. Copy the account rows and paste them into the P&L tab from cell A2, account names "
          "in column A and one month per column. The month headings across row 1 are fixed; "
          "match your paste to them."),
    ("p", "3. Re-paste every prior month, not just the new one. Earlier months can move: an Aware "
          "Super adjustment retrospectively changed a July figure once, and anyone working off "
          "the old paste never saw it."),
    ("h", "What the portal does with it"),
    ("p", "Rows are matched to the chart of accounts on the account name, exactly as Xero spells "
          "it. Section headings and total rows such as Total Income and Gross Profit are ignored, "
          "so pasting the whole report is fine. A row whose name is not in the chart is reported "
          "on the Build page rather than silently dropped."),
])
heads = ["Account"] + [mlabel(d) for d in MONTHS_FY26] + [mlabel(d) for d in MONTHS_FY27]
ws = data_tab(wb, "P&L", heads, [40] + [11] * 24,
              ["Contract Support Staff"] + [0.0] * 24,
              "FY26 in columns B to M, FY27 in N to Y. Income positive, costs positive as Xero shows them.",
              paste_rows=260)
for r in range(2, 262):
    for c in range(2, 26):
        ws.cell(row=r, column=c).number_format = MONEY
# pre-list the chart so the paste has something to line up against
for i, a in enumerate(ACCOUNTS):
    ws.cell(row=2 + i, column=1, value=a["account"])
save(wb, "01 Xero P&L.xlsx")

# ------------------------------------------------------ 02 Xero GL transactions
wb = Workbook()
readme(wb, "02 Xero GL Transactions", [
    ("h", "What this is"),
    ("p", "One row per general ledger transaction. This is what every departmental figure in the "
          "portal is built from, because the department is read off each transaction's cost "
          "centre rather than the account."),
    ("h", "Each month"),
    ("p", "1. In Xero open the saved custom report Account transactions for P&L analysis "
          "(All Reports, then Custom). Run it year to date from 1 July, sorted by account code."),
    ("p", "2. Export to Excel. Convert the values to number before you copy, or the totals will "
          "not add up."),
    ("p", "3. Paste into the GL tab from cell A2, all thirteen columns in the order across row 1. "
          "Paste the whole year to date each time."),
    ("h", "Before you paste"),
    ("p", "Stripe fee rows arrive with no cost centre because Xero splits a card payment three "
          "ways. Set them to 9000 - OFFICE / ADMIN [CTS]. Revenue rows from manual journals "
          "arrive with no contact; fill them in, and spell each client one way."),
])
data_tab(wb, "GL",
         ["Index", "Account Code", "Account Name", "Source", "Date", "Contact", "Debit", "Credit",
          "Job Numbers", "Invoice Number", "Reference", "Description", "Cost Centres"],
         [7, 12, 34, 14, 12, 30, 13, 13, 28, 14, 16, 40, 30],
         [1, "42005", "Contract Support Staff", "Invoice", date(2026, 8, 14),
          "Meridian Bank Group", 0.0, 12450.0, "J26-0412 [ONS]", "INV-24501", "",
          "Onsite support, August", "1200 - ONSITE"],
         "Thirteen columns, the same order as the Xero report.", paste_rows=30000)
save(wb, "02 Xero GL Transactions.xlsx")

# ---------------------------------------------------- 03 Employment Hero earnings
wb = Workbook()
readme(wb, "03 Employment Hero Earnings", [
    ("h", "What this is"),
    ("p", "The Earnings Details report, one row per pay line. It gives the portal hours by "
          "department, chargeable against non-chargeable, leave, public holidays, labour cost, "
          "and casual against permanent for every person."),
    ("h", "Each month"),
    ("p", "1. In Employment Hero run Earnings Details for each pay run in the month. Export to Excel."),
    ("p", "2. Paste into the Earnings tab from cell A1, title row and all. Paste one pay run "
          "under the last, leaving the extra title and header rows in; the portal skips them."),
    ("p", "3. Check the Locations tab: any new location the export carries is listed on the Build "
          "page. Give it a department and say whether it is client work."),
    ("h", "The two tabs you maintain"),
    ("p", "Locations: which department each Employment Hero location belongs to and whether "
          "time on it is chargeable. The department is guessed from the bracket tag on the "
          "name, as in [ONS]; correct it here and it stays corrected."),
    ("p", "Pay categories: what each pay category counts as. Worked time, leave, public holiday, "
          "or pay only. Leave and public holiday are paid but not worked and stay out of "
          "utilisation. Pay only carries cost but no hours."),
    ("p", "Adjustments: a pay-only line that was actually hours worked, a back pay for "
          "instance. Give it the hours and the cost centre and it counts as worked time "
          "against that person and department."),
])
data_tab(wb, "Earnings",
         ["Employee Id", "Employee External Id", "Employee Name", "Pay Category Id",
          "Pay Category External Id", "Pay Category Name", "Units", "Unit Type", "Location Id",
          "Location External Id", "Location Name", "Notes", "Rate", "Rate Type", "Gross Earnings",
          "Taxable Earnings", "SG Super"],
         [11, 10, 22, 12, 12, 26, 8, 9, 11, 12, 34, 44, 10, 10, 13, 13, 10],
         [4745780, 298, "Abel SAVIGE", 2856297, 2286634, "Casual Ordinary Hours", 4.25, "Hours",
          755271, 26012301, "PWC IT Support - Brisbane [ONS]",
          "7/09/2026 - 08:45 to 13:00. 07-09-2026 [05:39 pm]: Finish 0515 regular shift  (Abel SAVIGE)",
          39.4618, "per Hour", 167.71, 167.71, 20.13],
         "Paste from A1 with the Earnings Details title row; it is skipped.",
         paste_rows=12000, title_row="Earnings Details")
list_tab(wb, "Locations",
         ["Location Name", "Department", "Client work"],
         [["PWC IT Support - Brisbane [ONS]", "ONSITE", "Chargeable"],
          ["CTS Head Office [CTS]", "ADMIN", "Non-chargeable"]],
         [40, 16, 16],
         note="Department: ONSITE, PRODUCTION, VIDEO, INTEGRATION, CONSULTING, ADMIN. Client work: Chargeable or Non-chargeable.")
dv = DataValidation(type="list", formula1='"ONSITE,PRODUCTION,VIDEO,INTEGRATION,CONSULTING,ADMIN"', allow_blank=True)
wb["Locations"].add_data_validation(dv); dv.add("B2:B500")
dv2 = DataValidation(type="list", formula1='"Chargeable,Non-chargeable"', allow_blank=True)
wb["Locations"].add_data_validation(dv2); dv2.add("C2:C500")
for r in range(4, 502):
    for c in (1, 2, 3):
        wb["Locations"].cell(row=r, column=c).fill = PASTE
        wb["Locations"].cell(row=r, column=c).font = INPUT_FONT
list_tab(wb, "Pay categories",
         ["Pay Category Name", "Counts as"],
         [["Casual Ordinary Hours", "Worked time"], ["Permanent Ordinary Hours", "Worked time"],
          ["Casual Overtime x 1.75", "Worked time"], ["Permanent Overtime x 1.25", "Worked time"],
          ["Annual Leave", "Leave"], ["Personal/Carer's Leave", "Leave"],
          ["Time In Lieu Taken", "Leave"], ["Bonus Leave", "Leave"],
          ["Public Holiday", "Public holiday"], ["Leave Without Pay", "Pay only"],
          ["Back Payment", "Pay only"], ["Per Diem", "Pay only"],
          ["Unused leave payment (normal termination)", "Pay only"]],
         [40, 18],
         note="Counts as: Worked time, Leave, Public holiday, Pay only. Any category not listed is guessed from its name and reported on the Build page.")
dv3 = DataValidation(type="list", formula1='"Worked time,Leave,Public holiday,Pay only"', allow_blank=True)
wb["Pay categories"].add_data_validation(dv3); dv3.add("B2:B300")
for r in range(15, 302):
    for c in (1, 2):
        wb["Pay categories"].cell(row=r, column=c).fill = PASTE
        wb["Pay categories"].cell(row=r, column=c).font = INPUT_FONT
list_tab(wb, "Adjustments",
         ["Employee Name", "Pay Category Name", "Note contains", "Hours worked", "Post to location", "Counts as"],
         [["C Diaz", "Back Payment", "correction to PR14", 6.0, "Ashfield AGM Event [PRD]", "Worked time"]],
         [22, 26, 30, 12, 34, 16],
         note="One row per pay-only line you want counted as hours. Matched on employee, category and a phrase from the note.")
for r in range(3, 302):
    for c in range(1, 7):
        wb["Adjustments"].cell(row=r, column=c).fill = PASTE
        wb["Adjustments"].cell(row=r, column=c).font = INPUT_FONT
save(wb, "03 Employment Hero Earnings.xlsx")

# ----------------------------------------------------------- 04 Zoho deals
wb = Workbook()
readme(wb, "04 Zoho CRM Deals", [
    ("h", "What this is"),
    ("p", "Open and recently closed deals from Zoho CRM. This is the pipeline: expected revenue "
          "by month and by department, win rate, and what the Revenue Schedule's forecast "
          "column should read instead of last year's shares."),
    ("h", "Each month"),
    ("p", "1. In Zoho CRM open the Deals module, choose all deals or a view covering the last "
          "twelve months, and export to Excel."),
    ("p", "2. Paste into the Deals tab from cell A2."),
    ("h", "Department"),
    ("p", "The portal needs to know which department a deal belongs to. If your Zoho has a "
          "field for it, keep it in the Department column; otherwise the Build page lets you "
          "map Type or Lead Source to a department once."),
], placeholder=True)
data_tab(wb, "Deals",
         ["Deal Name", "Account Name", "Stage", "Amount", "Closing Date", "Deal Owner",
          "Probability (%)", "Expected Revenue", "Type", "Lead Source", "Contact Name",
          "Created Time", "Modified Time", "Department"],
         [32, 28, 18, 13, 13, 18, 12, 14, 18, 16, 22, 16, 16, 14],
         ["Boardroom refresh, level 12", "Meridian Bank Group", "Proposal/Price Quote", 48500.0,
          date(2026, 10, 15), "Jordan", 60, 29100.0, "New Business", "Referral", "A Contact",
          date(2026, 8, 3), date(2026, 9, 12), "INTEGRATION"],
         "Zoho's standard Deals export, plus a Department column.", paste_rows=3000)
save(wb, "04 Zoho Deals.xlsx")

# ---------------------------------------------------------- 05 OnRent orders
wb = Workbook()
readme(wb, "05 OnRent Events Orders", [
    ("h", "What this is"),
    ("p", "Confirmed and quoted jobs from OnRent Events. This is Production's order book: what "
          "is booked by month, which is the forward view of the seasonality the monthly "
          "report reads."),
    ("h", "Each month"),
    ("p", "1. In OnRent export the orders or jobs list for the coming and recent months."),
    ("p", "2. Paste into the Orders tab from cell A2."),
    ("h", "What the portal needs"),
    ("p", "A job number, the client, the event start and end dates, the department, a status "
          "that says whether it is a quote, confirmed, completed or cancelled, and the value "
          "excluding GST. Anything else in the export is left alone."),
], placeholder=True)
data_tab(wb, "Orders",
         ["Order No", "Client", "Job Title", "Event Start", "Event End", "Department", "Status",
          "Value ex GST", "Owner", "Created"],
         [12, 28, 34, 12, 12, 14, 12, 14, 16, 12],
         ["ORD-2318", "Ashfield Capital Partners", "AGM, Sydney", date(2026, 11, 12),
          date(2026, 11, 12), "PRODUCTION", "Confirmed", 38200.0, "M Held", date(2026, 8, 20)],
         "Status: Quote, Confirmed, Completed, Cancelled.", paste_rows=3000)
save(wb, "05 OnRent Orders.xlsx")

# ----------------------------------------------------------- 06 Qwilr quotes
wb = Workbook()
readme(wb, "06 Qwilr Quotes", [
    ("h", "What this is"),
    ("p", "Quotes issued through Qwilr and what became of them: sent, viewed, accepted, "
          "declined. This gives quote conversion by department and by month."),
    ("h", "Each month"),
    ("p", "1. In Qwilr export the projects or quotes list for the last twelve months."),
    ("p", "2. Paste into the Quotes tab from cell A2."),
    ("h", "What the portal needs"),
    ("p", "A quote reference, the client, when it was sent, its status, when it was accepted "
          "if it was, its value excluding GST, who owns it, and the department."),
], placeholder=True)
data_tab(wb, "Quotes",
         ["Quote Ref", "Client", "Title", "Sent", "Status", "Accepted", "Value ex GST", "Owner",
          "Department"],
         [12, 28, 34, 12, 12, 12, 14, 16, 14],
         ["Q-1042", "Trentham Logistics", "Warehouse AV integration", date(2026, 8, 22),
          "Accepted", date(2026, 9, 3), 64000.0, "W Corrigan", "INTEGRATION"],
         "Status: Draft, Sent, Viewed, Accepted, Declined, Expired.", paste_rows=3000)
save(wb, "06 Qwilr Quotes.xlsx")

# ---------------------------------------------------------------- 07 Budget
wb = Workbook()
readme(wb, "07 Budget", [
    ("h", "What this is"),
    ("p", "The budget by account by month, in the same layout as the P&L template so the two "
          "line up row for row. This drives every budget column and every risk flag, and it "
          "fills the months after the reporting month on the P&L."),
    ("h", "Once a year, and when the forecast moves"),
    ("p", "Type or paste the budget into the Budget tab against each account. FY26 in columns "
          "B to M, FY27 in N to Y. Income positive, costs positive."),
    ("p", "Revise when the rolling forecast moves. The portal reads whatever is here."),
])
ws = data_tab(wb, "Budget", heads, [40] + [11] * 24, ["Contract Support Staff"] + [0.0] * 24,
              "Same layout as the P&L template.", paste_rows=260)
for r in range(2, 262):
    for c in range(2, 26):
        ws.cell(row=r, column=c).number_format = MONEY
for i, a in enumerate(ACCOUNTS):
    ws.cell(row=2 + i, column=1, value=a["account"])
save(wb, "07 Budget.xlsx")

# ---------------------------------------------------------------- 08 Config
wb = Workbook()
readme(wb, "08 Config", [
    ("h", "What this is"),
    ("p", "Everything the portal is set up with, in one place: the reporting month, the "
          "departments, the overhead split bases, public holidays, who sees which tabs, and "
          "who receives the monthly email. Change it here and rebuild; nothing has to be "
          "set in a browser."),
    ("h", "Each month"),
    ("p", "Change the reporting month on the Reporting tab. That is the one control everything "
          "else follows from."),
    ("h", "The split bases"),
    ("p", "The 3 Way percentages are confirmed from the budget bridge. Staff and Office Dept "
          "are placeholders until the live percentages from the budget workbook are typed in. "
          "Every allocated departmental figure rests on them, and the portal flags any figure "
          "that does until they are marked Confirmed."),
])
list_tab(wb, "Reporting",
         ["Setting", "Value", "Note"],
         [["Reporting month", date(2026, 8, 1), "The one control you change each month"],
          ["Financial year starts in month", 7, "July"],
          ["Hours in a working day", 8, "Used by utilisation"],
          ["State for working days", "VIC", "Public holidays differ by state"],
          ["Materiality floor ($)", 5000, "Below this a variance is always Low"],
          ["Medium risk (% variance)", 0.10, ""],
          ["High risk (% variance)", 0.25, ""],
          ["High risk ($ variance)", 50000, "Either test alone makes it High"],
          ["Portal version", "2.0.0", ""]],
         [32, 16, 44], editable_cols=[2])
wb["Reporting"]["B2"].number_format = "mmm yyyy"
wb["Reporting"]["B7"].number_format = PCT
wb["Reporting"]["B8"].number_format = PCT
list_tab(wb, "Departments",
         ["Code", "Short name", "Long name", "Order", "Revenue dept", "Utilisation target",
          "GM norm", "NP norm", "Tags", "Account suffix", "Manager", "Note"],
         [["ONSITE", "Onsite", "Onsite and managed services", 1, "Yes", 0.65, 0.45, 0.10, "ONSITE, ONS", "ONS", "",
           "Support and managed services: always profitable, GM around 45%, NP above 10%"],
          ["PRODUCTION", "Production", "Event production", 2, "Yes", 0.65, 0.55, None, "PRODUCTION, PRD", "PRD", "",
           "Very seasonal. GM 55% or better; in the 40s is worth investigating"],
          ["VIDEO", "Video", "Video production", 3, "Yes", 0.65, None, None, "VIDEO, VID", "VID", "", ""],
          ["INTEGRATION", "Integration", "Systems integration", 4, "Yes", 0.65, 0.15, None, "INTEGRATION, INT", "INTEGRATION", "",
           "GM at least 15%. NP skewed by how few people work in it"],
          ["CONSULTING", "Consulting", "Technology consulting", 5, "Yes", 0.65, None, None, "CONSULTING, CONS", "CONS", "",
           "Frequently below budget; a loss in most months is normal. Runs mid forties on utilisation"],
          ["ADMIN", "Admin", "Office and administration", 6, "No", None, None, None, "CTS, ADMIN", "ADMIN", "",
           "Carries the overheads that are pushed out"]],
         [13, 12, 28, 6, 10, 11, 9, 9, 20, 12, 16, 60],
         note="Utilisation target: only consulting's is documented (about 65%). The rest are assumed.")
list_tab(wb, "Split bases",
         ["Basis", "Production + Video", "Onsite", "Consulting + Integration", "Status", "Note"],
         [["3 Way", 0.34, 0.33, 0.33, "Confirmed", "Split roughly evenly three ways. From the budget bridge."],
          ["Staff", 0.40, 0.47, 0.13, "Placeholder", "By headcount. Type the live percentages from the budget workbook and set Confirmed."],
          ["Office Dept", 0.20, 0.65, 0.15, "Placeholder", "Weighted to one department. Type the live percentages and set Confirmed."],
          ["", None, None, None, "", ""],
          ["Production share of Production + Video", 0.79, None, None, "Confirmed", "FY23 actuals on the PRD tab"],
          ["Integration share of Consulting + Integration", 0.72, None, None, "Confirmed", "FY23 actuals on the CONS tab"]],
         [40, 14, 12, 18, 12, 60], editable_cols=[2, 3, 4, 5])
for r in (2, 3, 4, 6, 7):
    for c in (2, 3, 4):
        wb["Split bases"].cell(row=r, column=c).number_format = PCT
dv4 = DataValidation(type="list", formula1='"Confirmed,Placeholder"', allow_blank=True)
wb["Split bases"].add_data_validation(dv4); dv4.add("E2:E7")
subs = sorted({a["account"] for a in ACCOUNTS if a["group"] == "Expenses"})
# subcategory = the account name with the division suffix stripped, as the portal derives it
import re
SUFFIX = re.compile(r"\s*-\s*(ONS|PRD|VID|CONS|INTEGRATION|ADMIN)$")
subcats = sorted({SUFFIX.sub("", n) for n in subs if not SUFFIX.search(n) or True})
subcats = sorted({SUFFIX.sub("", n) for n in subs})
OFFICE = {"Rent", "Rental Outgoings", "Office Rentals", "Electricity", "Office Cleaning", "Storage Fees",
          "Internet", "Office Phones", "Office Supplies", "Printing", "Postage", "Repairs & Maintenance"}
STAFF = {"Superannuation", "Payroll Tax", "Workers' Compensation", "Recruitment", "Staff Amenities",
         "Staff Entertainment", "Training Material & Courses", "Payroll Processing Fee",
         "Leave expense (Annual)", "Fringe Benefits Tax", "Mobile Phones", "Travel & Per Diems",
         "Taxis/Parking", "IT Network Service & Support"}
list_tab(wb, "Split assignment",
         ["Overhead subcategory", "Basis"],
         [[s, "Office Dept" if s in OFFICE else "Staff" if s in STAFF else "3 Way"] for s in subcats],
         [40, 14], editable_cols=[2],
         note="Which basis each Admin overhead line is pushed out on. Derived from the account; correct it to match the labels on your allocation tab.")
dv5 = DataValidation(type="list", formula1='"3 Way,Staff,Office Dept"', allow_blank=True)
wb["Split assignment"].add_data_validation(dv5); dv5.add("B2:B200")

# public holidays: the fixed and Easter-derived ones, plus the state ones marked to verify
import sys
sys.path.insert(0, str(ROOT / "portal/build"))
hol_rows = []
try:
    from datetime import timedelta
    def easter(year):
        a, b, c = year % 19, year // 100, year % 100
        d, e = b // 4, b % 4
        f = (b + 8) // 25; g = (b - f + 1) // 3
        h = (19 * a + b - d - g + 15) % 30
        i, k = c // 4, c % 4
        l = (32 + 2 * e + 2 * i - h - k) % 7
        m = (a + 11 * h + 22 * l) // 451
        month = (h + l - 7 * m + 114) // 31
        day = ((h + l - 7 * m + 114) % 31) + 1
        return date(year, month, day)
    def nth(y, m, wd, n):
        d = date(y, m, 1); d += timedelta(days=(wd - d.weekday()) % 7); return d + timedelta(days=7 * (n - 1))
    def sub(d):
        return d + timedelta(days=2) if d.weekday() == 5 else d + timedelta(days=1) if d.weekday() == 6 else d
    for y in (2026, 2027):
        e = easter(y)
        hol_rows += [[sub(date(y, 1, 1)), "All", "New Year's Day", "Fixed"],
                     [sub(date(y, 1, 26)), "All", "Australia Day", "Fixed"],
                     [e - timedelta(days=2), "All", "Good Friday", "Fixed"],
                     [e + timedelta(days=1), "All", "Easter Monday", "Fixed"],
                     [date(y, 4, 25), "All", "Anzac Day", "Fixed"],
                     [sub(date(y, 12, 25)), "All", "Christmas Day", "Fixed"],
                     [sub(date(y, 12, 26)), "All", "Boxing Day", "Fixed"],
                     [nth(y, 3, 0, 2), "VIC", "Labour Day", "Verify"],
                     [nth(y, 6, 0, 2), "VIC", "King's Birthday", "Verify"],
                     [nth(y, 11, 1, 1), "VIC", "Melbourne Cup", "Verify"],
                     [nth(y, 10, 0, 1), "NSW", "Labour Day", "Verify"],
                     [nth(y, 6, 0, 2), "NSW", "King's Birthday", "Verify"],
                     [nth(y, 5, 0, 1), "QLD", "Labour Day", "Verify"],
                     [nth(y, 10, 0, 1), "QLD", "King's Birthday", "Verify"]]
except Exception as ex:
    print("holiday generation failed:", ex)
hol_rows.sort(key=lambda r: (r[0], r[1]))
list_tab(wb, "Holidays",
         ["Date", "State", "Holiday", "Status"],
         hol_rows, [13, 8, 22, 10],
         note="All means every state. Fixed dates and Easter are certain; Verify means a state proclamation that moves. Check those against the official list. Add a row for any other state you employ people in.")
list_tab(wb, "Users",
         ["Name", "Role", "Department", "Email", "Note"],
         [["Danica Nelson", "Finance head", "", "", "Holds the finance head role until handed over"],
          ["Duncan", "Executive", "", "", "Approves payments, reads the monthly report"],
          ["Graham", "Executive", "", "", "Approves pay runs"],
          ["Jordan", "Department head", "CONSULTING", "", "Consulting"],
          ["Production manager", "Department head", "PRODUCTION", "", ""]],
         [22, 16, 14, 30, 44],
         note="Role: Finance head, Finance, Executive, Department head, Viewer. Department only matters for a Department head.")
dv6 = DataValidation(type="list", formula1='"Finance head,Finance,Executive,Department head,Viewer"', allow_blank=True)
wb["Users"].add_data_validation(dv6); dv6.add("B2:B200")
for r in range(7, 202):
    for c in range(1, 6):
        wb["Users"].cell(row=r, column=c).fill = PASTE
        wb["Users"].cell(row=r, column=c).font = INPUT_FONT
PAGES = [("home", "Dashboard"), ("context", "Business Context"), ("pnl", "P&L"),
         ("pnl-dept", "P&L by Department"), ("pnl-spread", "P&L Spread"),
         ("allocation", "Overhead Allocation"), ("control", "P&L Control"),
         ("bva", "Budget vs Actual"), ("actions", "Actions"), ("rev-summary", "Revenue Summary"),
         ("rev-schedule", "Revenue Schedule"), ("pipeline", "Pipeline"), ("clients", "Top Clients"),
         ("client-dept", "Clients by Department"), ("util", "Utilisation"),
         ("profit-fte", "Profitability per FTE"), ("staff-profit", "Profitability by Employee"),
         ("ledger", "Detail Records"), ("charts", "Charts"), ("setup", "Setup"),
         ("loaders", "Data Loaders"), ("build", "Build"), ("distribution", "Distribution"),
         ("config", "Config & Variables"), ("access", "Access Control"), ("about", "About")]
ROLE_PAGES = {
    "Finance": {p for p, _ in PAGES} - {"access"},
    "Executive": {"home", "context", "pnl", "pnl-dept", "pnl-spread", "bva", "actions", "rev-summary",
                  "rev-schedule", "pipeline", "clients", "client-dept", "util", "profit-fte",
                  "staff-profit", "charts", "about"},
    "Department head": {"home", "context", "pnl-dept", "rev-schedule", "pipeline", "clients", "util",
                        "charts", "about"},
    "Viewer": {"home", "context", "about"},
}
list_tab(wb, "Tabs by role",
         ["Tab id", "Tab", "Finance head", "Finance", "Executive", "Department head", "Viewer"],
         [[pid, title, "Always",
           "Yes" if pid in ROLE_PAGES["Finance"] else "",
           "Yes" if pid in ROLE_PAGES["Executive"] else "",
           "Yes" if pid in ROLE_PAGES["Department head"] else "",
           "Yes" if pid in ROLE_PAGES["Viewer"] else ""] for pid, title in PAGES],
         [14, 26, 12, 10, 11, 15, 9], editable_cols=[4, 5, 6, 7],
         note="Yes shows the tab to that role. This decides what people are shown; it does not protect anything.")
list_tab(wb, "Distribution",
         ["Name", "Email", "Tier", "Department", "Send", "Note"],
         [["Duncan", "duncan@example.com.au", 1, "", "Yes", "Full result, cost and margin, commentary"],
          ["Graham", "graham@example.com.au", 1, "", "Yes", ""],
          ["Jordan", "jordan@example.com.au", 2, "CONSULTING", "Yes", "Own department in depth, one line on the rest"],
          ["Production manager", "", 2, "PRODUCTION", "No", "No address yet"],
          ["Danica Nelson", "danica@example.com.au", 3, "", "Yes", "Finance: the controls and data quality"]],
         [22, 30, 6, 14, 6, 44],
         note="Tier 1 executive, tier 2 department head, tier 3 finance. Send: Yes or No. Addresses here are placeholders; replace them.")
dv7 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
wb["Distribution"].add_data_validation(dv7); dv7.add("E2:E200")
for r in range(7, 202):
    for c in range(1, 7):
        wb["Distribution"].cell(row=r, column=c).fill = PASTE
        wb["Distribution"].cell(row=r, column=c).font = INPUT_FONT
save(wb, "08 Config.xlsx")
print("done")
