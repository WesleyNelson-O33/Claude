"""Build the CTS Business Intelligence Portal data files.

Every figure that could be sourced is sourced. The FY26 control totals and the
July 2025 month come from reporting/data/validation.json, which was pulled from
Xero on 3 September 2026. The August 2026 category totals come from the figures
pack/loadgl.py wrote into the Controller Pack's PL_Check sheet. The chart of
accounts is reporting/data/accounts.json, read from Xero on the same date.

Everything between those anchors is modelled: the seed reconciles to the real
totals to the cent, but the shape inside them comes from the seasonality and
margin norms written down in the Monthly Reporting manual, not from real
transactions. Every generated file carries seed:true so the portal can say so
on screen.
"""
import json
import math
from datetime import date, timedelta
from pathlib import Path

ROOT = Path("/home/user/Claude")
REP = ROOT / "reporting/data"
OUT = ROOT / "portal/data"

ACCOUNTS = json.loads((REP / "accounts.json").read_text())
VALID = json.loads((REP / "validation.json").read_text())

BUILT = "2026-09-21"

# ----------------------------------------------------------------- determinism
class Rand:
    """Small LCG. Deterministic across runs and across machines."""

    def __init__(self, seed=20260921):
        self.s = seed

    def next(self):
        self.s = (1103515245 * self.s + 12345) % (1 << 31)
        return self.s / (1 << 31)

    def pick(self, seq):
        return seq[int(self.next() * len(seq)) % len(seq)]

    def jitter(self, lo, hi):
        return lo + self.next() * (hi - lo)


R = Rand()

# ------------------------------------------------------------------ department
# The six reporting departments and the tags that resolve to them, taken from
# DEPT_TAGS in pack/build_pack.py.
DEPTS = ["ONSITE", "PRODUCTION", "VIDEO", "INTEGRATION", "CONSULTING", "ADMIN"]

DEPT_TAGS = [
    ("ONSITE", "ONSITE"), ("PRODUCTION", "PRODUCTION"), ("VIDEO", "VIDEO"),
    ("INTEGRATION", "INTEGRATION"), ("CONSULTING", "CONSULTING"), ("CTS", "ADMIN"),
    ("ONS", "ONSITE"), ("PRD", "PRODUCTION"), ("VID", "VIDEO"),
    ("INT", "INTEGRATION"), ("CONS", "CONSULTING"), ("ADMIN", "ADMIN"),
]

# Division suffix on an account name, from reporting/build_accounts.py.
SUFFIX_TO_DEPT = {"ONS": "ONSITE", "PRD": "PRODUCTION", "VID": "VIDEO",
                  "CONS": "CONSULTING", "INTEGRATION": "INTEGRATION", "ADMIN": "ADMIN"}

COST_CENTRE_CODE = {"ONSITE": "1200", "PRODUCTION": "2100", "VIDEO": "2400",
                    "INTEGRATION": "3100", "CONSULTING": "3600", "ADMIN": "9000"}
JOB_TAG = {"ONSITE": "ONS", "PRODUCTION": "PRD", "VIDEO": "VID",
           "INTEGRATION": "INT", "CONSULTING": "CONS", "ADMIN": "CTS"}

# Which departments may carry an account that has no division in its name.
# Derived from the account's meaning, not from the data, so it is configurable
# in the portal and flagged there as derived.
UNSUFFIXED_HINTS = {
    "Ad Hoc Service Labour": ["ONSITE"],
    "Contract Help Desk": ["ONSITE"],
    "Contract Maintenance": ["ONSITE"],
    "Contract Support Staff": ["ONSITE"],
    "Installation Labour": ["ONSITE", "INTEGRATION"],
    "Production Labour": ["PRODUCTION"],
    "Equipment Hires": ["PRODUCTION", "VIDEO"],
    "Tech Event Management": ["PRODUCTION"],
    "Video Labour": ["VIDEO"],
    "Tech Consulting/Project Manage": ["CONSULTING"],
    "Equipment Sales": ["INTEGRATION"],
    "Conference Call Charge": ["INTEGRATION"],
    "Mobile Phone Plans": ["INTEGRATION"],
    "Subscriptions & Licences - Income": ["INTEGRATION", "ONSITE"],
    "Miscellaneous Income": ["ADMIN"],
    "Opening Work in Progress": ["ADMIN"],
    "Closing Work in Progress": ["ADMIN"],
}

# Cost of sales accounts with no division take the department of the job.
COS_ANY = ["Damage Waiver", "Discounts Given", "Equipment - Hires", "Equipment - Purchase",
           "Equipment -Maintenance Support", "Service - hires", "Stamp Duty",
           "Subscriptions & Licences - Expense"]


def account_dept(name):
    for suffix, dept in SUFFIX_TO_DEPT.items():
        if name.endswith("- " + suffix) or name.endswith("-" + suffix):
            return dept
    return None


def subcategory(name):
    """Account name with the division suffix stripped.

    This reproduces the subcategory names used by the budget bridge
    (Direct Salaries, Direct Superannuation, Equipment - Purchase and the
    rest), and it gives the P&L its middle level: category, subcategory,
    account.
    """
    for suffix in SUFFIX_TO_DEPT:
        for sep in (" - ", "-"):
            if name.endswith(sep + suffix):
                return name[: -len(sep + suffix)].strip()
    return name


# ---------------------------------------------------------------- the calendar
# Australian financial year, 1 July to 30 June. FY is named for the year it ends.
MONTH_ORDER = [7, 8, 9, 10, 11, 12, 1, 2, 3, 4, 5, 6]
MONTH_NAME = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
STATES = ["VIC", "NSW", "QLD", "WA", "SA", "TAS", "ACT", "NT"]


def fy_months(fy):
    """[(year, month)] for a financial year, July first."""
    return [(fy - 1 if m >= 7 else fy, m) for m in MONTH_ORDER]


def mkey(y, m):
    return "%04d-%02d" % (y, m)


def easter(year):
    """Anonymous Gregorian algorithm. Good Friday and Easter Monday follow."""
    a, b, c = year % 19, year // 100, year % 100
    d, e = b // 4, b % 4
    f, g = (b + 8) // 25, 0
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def nth_weekday(year, month, weekday, n):
    d = date(year, month, 1)
    d += timedelta(days=(weekday - d.weekday()) % 7)
    return d + timedelta(days=7 * (n - 1))


def substitute(d):
    """A fixed-date holiday falling on a weekend moves to the next Monday."""
    if d.weekday() == 5:
        return d + timedelta(days=2)
    if d.weekday() == 6:
        return d + timedelta(days=1)
    return d


def holidays(year):
    """Public holidays for one calendar year.

    certain=True are the ones that are fixed by date or computable exactly.
    certain=False are the state-specific ones whose dates are proclaimed and
    do move; the portal shows them amber and asks you to verify them against
    your state's official list before relying on the utilisation figures.
    """
    out = []
    e = easter(year)

    def add(d, name, states, certain):
        out.append({"date": d.isoformat(), "name": name, "states": states,
                    "certain": certain})

    add(substitute(date(year, 1, 1)), "New Year's Day", STATES, True)
    add(substitute(date(year, 1, 26)), "Australia Day", STATES, True)
    add(e - timedelta(days=2), "Good Friday", STATES, True)
    add(e + timedelta(days=1), "Easter Monday", STATES, True)
    add(date(year, 4, 25), "Anzac Day", STATES, True)
    add(substitute(date(year, 12, 25)), "Christmas Day", STATES, True)
    add(substitute(date(year, 12, 26)), "Boxing Day", STATES, True)

    add(nth_weekday(year, 3, 0, 2), "Labour Day", ["VIC", "TAS"], False)
    add(nth_weekday(year, 3, 0, 1), "Labour Day", ["WA"], False)
    add(nth_weekday(year, 5, 0, 1), "Labour Day", ["QLD", "NT"], False)
    add(nth_weekday(year, 10, 0, 1), "Labour Day", ["NSW", "SA", "ACT"], False)
    add(nth_weekday(year, 6, 0, 2), "King's Birthday",
        ["VIC", "NSW", "SA", "TAS", "ACT", "NT"], False)
    add(nth_weekday(year, 9, 0, 4), "King's Birthday", ["WA"], False)
    add(nth_weekday(year, 10, 0, 1), "King's Birthday", ["QLD"], False)
    add(nth_weekday(year, 11, 1, 1), "Melbourne Cup", ["VIC"], False)
    return out


def working_days(y, m, state, hol_index):
    d = date(y, m, 1)
    days = 0
    while d.month == m:
        if d.weekday() < 5 and (d.isoformat(), state) not in hol_index:
            days += 1
        d += timedelta(days=1)
    return days


def build_calendar():
    hol = []
    for year in (2025, 2026, 2027):
        hol += holidays(year)
    index = {(h["date"], s) for h in hol for s in h["states"]}

    months = []
    for fy in (26, 27):
        for i, (y, m) in enumerate(fy_months(2000 + fy)):
            months.append({
                "key": mkey(y, m), "y": y, "m": m, "fy": fy,
                "label": "%s-%02d" % (MONTH_NAME[m - 1], y % 100),
                "long": "%s %d" % (MONTH_NAME[m - 1], y),
                "period": i + 1, "quarter": i // 3 + 1,
                "wd": {s: working_days(y, m, s, index) for s in STATES},
                # weekdays only, no holidays deducted. This is the denominator
                # that makes a year land on 2,080 or 2,088 hours, which is the
                # figure Monthly Reporting Part 2 says to check against.
                "bd": working_days(y, m, "__none__", set()),
                "days": (date(y + (m == 12), m % 12 + 1, 1) - date(y, m, 1)).days,
            })
    return {"fyStartMonth": 7, "states": STATES, "months": months,
            "holidays": hol,
            "quarters": [{"n": 1, "label": "Q1 Jul-Sep"}, {"n": 2, "label": "Q2 Oct-Dec"},
                         {"n": 3, "label": "Q3 Jan-Mar"}, {"n": 4, "label": "Q4 Apr-Jun"}],
            "meta": {"built": BUILT,
                     "note": "Financial year 1 July to 30 June, named for the year it ends. "
                             "Holidays flagged certain:false are state proclamations and must "
                             "be verified before the utilisation figures are relied on."}}


CAL = build_calendar()
MONTHS26 = [m["key"] for m in CAL["months"] if m["fy"] == 26]
MONTHS27 = [m["key"] for m in CAL["months"] if m["fy"] == 27]
ALL_MONTHS = MONTHS26 + MONTHS27

# --------------------------------------------------------------- the P&L shape
# Revenue mix and gross margin by department. Modelled: the blend lands on the
# real FY26 gross margin of 42.8%, and the per-department margins follow the
# norms written down in Monthly Reporting Part 3.
DEPT_MIX = {
    "ONSITE":      {"rev": 0.380, "gm": 0.45},
    "PRODUCTION":  {"rev": 0.300, "gm": 0.57},
    "VIDEO":       {"rev": 0.080, "gm": 0.50},
    "INTEGRATION": {"rev": 0.170, "gm": 0.16},
    "CONSULTING":  {"rev": 0.060, "gm": 0.35},
    "ADMIN":       {"rev": 0.010, "gm": 0.00},
}

# Monthly shape, July first. Onsite is contract work and close to flat.
# Production and video follow the calendar in Part 3: results season in August
# and September, the AGM peak from October to December, quiet in January and
# March. Integration and consulting peak in May and June as the year closes.
SHAPE = {
    "ONSITE":      [.082, .085, .085, .087, .087, .080, .072, .085, .085, .082, .085, .085],
    "PRODUCTION":  [.055, .090, .095, .115, .120, .085, .045, .090, .060, .075, .085, .085],
    "VIDEO":       [.055, .090, .095, .115, .120, .085, .045, .090, .060, .075, .085, .085],
    "INTEGRATION": [.060, .070, .075, .080, .085, .070, .050, .080, .080, .085, .130, .135],
    "CONSULTING":  [.060, .070, .075, .080, .085, .070, .050, .080, .080, .085, .130, .135],
    "ADMIN":       [.083, .083, .084, .083, .084, .083, .083, .084, .083, .083, .084, .083],
}
FLAT = [1 / 12.0] * 12


def cents(x):
    return int(round(x * 100))


def profile(weights, anchor_frac=None):
    """Normalise a 12-month weight vector, optionally pinning July exactly."""
    total = sum(weights)
    p = [w / total for w in weights]
    if anchor_frac is None:
        return p
    rest = sum(p[1:])
    scale = (1 - anchor_frac) / rest
    return [anchor_frac] + [w * scale for w in p[1:]]


def split_exact(total_c, fracs):
    """Split an integer number of cents by fractions, exactly."""
    raw = [total_c * f for f in fracs]
    out = [int(math.floor(v)) for v in raw]
    short = total_c - sum(out)
    order = sorted(range(len(raw)), key=lambda i: raw[i] - out[i], reverse=True)
    for i in range(short):
        out[order[i % len(order)]] += 1
    return out


FY26 = VALID["periods"]["FY26"]
JUL25 = VALID["periods"]["2025-07"]

# The one real FY27 anchor: the August 2026 category totals that loadgl.py
# wrote into PL_Check as the worked example. Costs are shown there with the
# pack's sign convention; taken here as positive magnitudes.
AUG26 = {"income": 603466.54, "cos": 335516.00, "expenses": 208889.82,
         "other_income": 861.35, "other_expenses": 0.00}

BLEND = [sum(DEPT_MIX[d]["rev"] * SHAPE[d][i] for d in DEPTS) for i in range(12)]

TOTALS = {
    "income": (FY26["total_income"], JUL25["total_income"], BLEND),
    "cos": (FY26["total_cost_of_sales"], JUL25["total_cost_of_sales"], BLEND),
    "other_income": (FY26["total_other_income"], JUL25["total_other_income"], FLAT),
    # Overheads barely move month to month; July is the real anchor.
    "expenses": (FY26["total_expenses"], JUL25["total_expenses"], FLAT),
}

MONTH_TOTALS = {}
for cat, (fy_total, jul, shape) in TOTALS.items():
    frac = profile(shape, anchor_frac=jul / fy_total)
    MONTH_TOTALS[cat] = dict(zip(MONTHS26, split_exact(cents(fy_total), frac)))

# FY27 actuals: July and August only, because the reporting month is August
# 2026 and everything after it is forecast from budget. August is the real
# anchor. July is modelled as the prior July grown 6%.
GROWTH = 1.06
for cat in TOTALS:
    MONTH_TOTALS[cat][MONTHS27[0]] = int(round(MONTH_TOTALS[cat][MONTHS26[0]] * GROWTH))
MONTH_TOTALS["income"][MONTHS27[1]] = cents(AUG26["income"])
MONTH_TOTALS["cos"][MONTHS27[1]] = cents(AUG26["cos"])
MONTH_TOTALS["expenses"][MONTHS27[1]] = cents(AUG26["expenses"])
MONTH_TOTALS["other_income"][MONTHS27[1]] = cents(AUG26["other_income"])

ACTUAL_MONTHS = MONTHS26 + MONTHS27[:2]

print("FY26 income check :", sum(MONTH_TOTALS["income"][m] for m in MONTHS26) / 100,
      "vs", FY26["total_income"])
print("FY26 cos check    :", sum(MONTH_TOTALS["cos"][m] for m in MONTHS26) / 100,
      "vs", FY26["total_cost_of_sales"])
print("FY26 exp check    :", sum(MONTH_TOTALS["expenses"][m] for m in MONTHS26) / 100,
      "vs", FY26["total_expenses"])
print("Jul-25 income     :", MONTH_TOTALS["income"][MONTHS26[0]] / 100,
      "vs", JUL25["total_income"])
print("Aug-26 income     :", MONTH_TOTALS["income"][MONTHS27[1]] / 100,
      "vs", AUG26["income"])

(OUT / "..").resolve()
json.dump({"cal": CAL}, open("/tmp/cal_check.json", "w"))


# -------------------------------------------------------- department splitting
DEPT_REV = {d: DEPT_MIX[d]["rev"] for d in DEPTS}
DEPT_COS = {d: DEPT_MIX[d]["rev"] * (1 - DEPT_MIX[d]["gm"]) for d in DEPTS}
_cos_total = sum(DEPT_COS.values())
DEPT_COS = {d: v / _cos_total for d, v in DEPT_COS.items()}

# Departments carry their own indirect labour; everything else sits in Admin
# and is pushed out by the overhead split. That is the whole point of the
# allocation, so Admin has to hold the bulk of it.
DEPT_EXP = {"ONSITE": .16, "PRODUCTION": .14, "VIDEO": .04,
            "INTEGRATION": .06, "CONSULTING": .05, "ADMIN": .55}
DEPT_OI = {"ADMIN": 1.0}


def dept_split(cat, mi):
    """Fractions by department for one month index, shaped by seasonality."""
    if cat == "income":
        base = DEPT_REV
    elif cat == "cos":
        base = DEPT_COS
    elif cat == "expenses":
        base = DEPT_EXP
    else:
        return DEPT_OI
    if cat == "expenses":
        return base
    w = {d: base[d] * SHAPE[d][mi] for d in DEPTS}
    t = sum(w.values())
    return {d: v / t for d, v in w.items()}


# --------------------------------------------------------- account eligibility
CAT_OF_GROUP = {"Income": "income", "Cost of Sales": "cos",
                "Other Income": "other_income", "Expenses": "expenses"}

BY_NAME = {a["account"]: a for a in ACCOUNTS}
for a in ACCOUNTS:
    a["cat"] = CAT_OF_GROUP[a["group"]]
    a["sub"] = subcategory(a["account"])
    a["dept"] = account_dept(a["account"])

# Relative size of a subcategory within its category. Anything not named here
# takes the default, which keeps the long tail of small accounts alive.
WEIGHT = {
    # income
    "Contract Support Staff": 34, "Contract Help Desk": 16, "Contract Maintenance": 14,
    "Installation Labour": 9, "Ad Hoc Service Labour": 6,
    "Production Labour": 30, "Equipment Hires": 16, "Tech Event Management": 10,
    "Video Labour": 26, "Equipment Sales": 30, "Subscriptions & Licences - Income": 12,
    "Conference Call Charge": 4, "Mobile Phone Plans": 5,
    "Tech Consulting/Project Manage": 30, "Miscellaneous Income": 3,
    "Credit Card Collected": 1.2, "Freight & Travel": 2.0, "Meal & Travel": 1.6,
    "Credit Card Fees": 0.6, "Opening Work in Progress": 2, "Closing Work in Progress": 2,
    # cost of sales
    "Direct Salaries": 30, "Direct Wages": 22, "Sub-Contract Labour": 18,
    "Equipment - Purchase": 16, "Equipment - Hires": 9, "Direct Superannuation": 6,
    "Direct Travel": 4, "Direct Freight": 3, "Supplier Per Diems": 2.5,
    "Direct Mobile": 1.6, "Direct Workers Comp": 1.4, "Damage Waiver": 1.0,
    "Equipment -Maintenance Support": 6, "Subscriptions & Licences - Expense": 7,
    "Service - hires": 3, "Discounts Given": 0.8, "Stamp Duty": 0.4,
    "Credit Card Paid": 1.0,
    # expenses, departmental
    "Indirect Salaries": 24, "Indirect Wages": 14,
    # expenses, the Admin overhead pool
    "Rent": 26, "Superannuation": 22, "Payroll Tax": 10, "Insurance": 7,
    "IT Network Service & Support": 9, "Office Rentals": 6, "Electricity": 3,
    "Accounting Fees": 4, "Dues & Subscriptions": 4, "Legal Fees": 3,
    "Mobile Phones": 3, "Internet": 2.5, "Office Supplies": 2.5, "Printing": 1.2,
    "Depreciation": 8, "Depreciation - ROU": 5, "Interest Expense": 4,
    "Workers' Compensation": 3.5, "Recruitment": 2.5, "Staff Amenities": 2,
    "Staff Entertainment": 1.8, "Client Entertainment": 1.5, "Travel & Per Diems": 3,
    "Taxis/Parking": 1.6, "Repairs & Maintenance": 2, "Office Cleaning": 1.6,
    "Bank Charges": 1.2, "Stripe Fees": 2.2, "Credit Card Fees Paid": 1.8,
    "Payroll Processing Fee": 1.4, "Training Material & Courses": 1.6,
    "Advertising": 1.8, "Professional Indemnity": 1.5, "Other Insurance": 1.5,
    "Business Advisory": 2, "Income Tax Expense": 6, "Fringe Benefits Tax": 1.2,
    "Storage Fees": 1.4, "Office Phones": 1.2, "Freight Paid": 1.2,
    # other income
    "Interest Received": 30, "Interest Income": 20, "Rental Income": 24,
    "Income Other": 14, "Employee FBT Contribtution": 8,
}
DEFAULT_WEIGHT = 0.5

# Accounts that only ever carry a historic balance. Kept in the chart, given no
# weight, so they appear in the P&L at zero exactly as they do in Xero.
DORMANT = {"Cash Flow Boost", "JobKeeper Subsidy", "JobSaver Subsidy",
           "Gain on lease termination", "Bad Debts", "SGC Penalties", "Pager", "Fax",
           "1300 Number", "Custodian Vaults", "Instant Asset Write Off",
           "Disposal of fixed assets", "Dividend paid [93000]", "Bank Revaluations",
           "Foreign Currency Translation*", "Realised Currency Gains",
           "Unrealised Currency Gains", "Under/(Over) provision for tax",
           "Leave expense (Long Service)", "Filing Fee", "Donations / Charity",
           "Replacements", "Stamp Duty (Indirect)", "Travel Insurance",
           "Other Telco Expenses", "Business Combined", "Other Expenses [64900]",
           "Office Supplies [61500]"}


def weight(acc):
    if acc["account"] in DORMANT:
        return 0.0
    return WEIGHT.get(acc["sub"], DEFAULT_WEIGHT)


def eligible(cat, dept):
    """Accounts that a transaction in this category and department may hit."""
    out = []
    for a in ACCOUNTS:
        if a["cat"] != cat or weight(a) == 0:
            continue
        if a["dept"] == dept:
            out.append(a)
        elif a["dept"] is None:
            hint = UNSUFFIXED_HINTS.get(a["account"])
            if hint is not None:
                if dept in hint:
                    out.append(a)
            elif cat == "cos" and a["account"] in COS_ANY and dept != "ADMIN":
                out.append(a)
            elif cat == "expenses" and dept == "ADMIN":
                out.append(a)
            elif cat == "other_income" and dept == "ADMIN":
                out.append(a)
    return out


ELIGIBLE = {(c, d): eligible(c, d) for c in TOTALS for d in DEPTS}
for (c, d), accs in ELIGIBLE.items():
    if not accs and c != "other_income":
        print("  no accounts for", c, d)

# ------------------------------------------------------------------- the ledger
CLIENTS = [
    ("Meridian Bank Group", ["ONSITE"], 16), ("Parkin Vale Advisory", ["ONSITE"], 11),
    ("Southgate Mutual", ["ONSITE"], 8), ("Alderway Legal", ["ONSITE"], 6),
    ("Corell Industries", ["ONSITE", "INTEGRATION"], 7),
    ("Ashfield Capital Partners", ["PRODUCTION"], 13),
    ("Northbridge Energy", ["PRODUCTION"], 11),
    ("Vantage Resources Ltd", ["PRODUCTION"], 9),
    ("Keelson Group", ["PRODUCTION", "VIDEO"], 7),
    ("Lyndhurst Property Trust", ["PRODUCTION"], 6),
    ("Brayford Health", ["PRODUCTION", "VIDEO"], 5),
    ("Tessaro Beverages", ["VIDEO"], 6), ("Calder & Rowe", ["VIDEO"], 4),
    ("Harlow Education Group", ["VIDEO", "PRODUCTION"], 4),
    ("Trentham Logistics", ["INTEGRATION"], 12),
    ("Wexford Councils Alliance", ["INTEGRATION"], 9),
    ("Bellamy Retail Holdings", ["INTEGRATION"], 8),
    ("Orenda Systems", ["INTEGRATION"], 6),
    ("Marchmont Insurance", ["CONSULTING", "ONSITE"], 7),
    ("Quillon Software", ["CONSULTING"], 5),
    ("Farrow Institute", ["CONSULTING"], 4),
    ("Dunmore Civil", ["INTEGRATION", "PRODUCTION"], 5),
    ("Sable Point Hotels", ["PRODUCTION"], 4),
    ("Westcliffe Media", ["VIDEO"], 3),
    ("Ingram Foods Australia", ["ONSITE"], 5),
    ("Rothbury Mining", ["PRODUCTION"], 4),
    ("Pellier Group", ["ONSITE", "CONSULTING"], 3),
    ("Aveline Studios", ["VIDEO"], 3),
]
SUPPLIERS = ["Aldridge AV Supply", "Norcross Freight", "Peninsula Staffing",
             "Redgate Equipment Hire", "Sumner Technical Services", "Halloway Rentals",
             "Kestrel Travel", "Bright Lane Logistics", "Ormond Crew Services",
             "Verity Cabling", "Lockhart Audio", "Grange Power Systems",
             "Telstra Business", "Origin Energy", "Officeworks", "CBA Merchant",
             "Employsure", "Eastern Bell Legal", "Stripe Payments Australia"]

INDUSTRY = {
    "Meridian Bank Group": "Financial Services", "Southgate Mutual": "Financial Services",
    "Ashfield Capital Partners": "Financial Services", "Marchmont Insurance": "Insurance",
    "Parkin Vale Advisory": "Professional Services", "Alderway Legal": "Professional Services",
    "Farrow Institute": "Education", "Harlow Education Group": "Education",
    "Northbridge Energy": "Energy & Resources", "Vantage Resources Ltd": "Energy & Resources",
    "Rothbury Mining": "Energy & Resources", "Grange Power Systems": "Energy & Resources",
    "Trentham Logistics": "Transport & Logistics", "Bright Lane Logistics": "Transport & Logistics",
    "Wexford Councils Alliance": "Government", "Bellamy Retail Holdings": "Retail",
    "Ingram Foods Australia": "Retail", "Tessaro Beverages": "Retail",
    "Brayford Health": "Health", "Lyndhurst Property Trust": "Property",
    "Sable Point Hotels": "Hospitality", "Keelson Group": "Industrial",
    "Corell Industries": "Industrial", "Dunmore Civil": "Construction",
    "Orenda Systems": "Technology", "Quillon Software": "Technology",
    "Westcliffe Media": "Media", "Aveline Studios": "Media",
    "Calder & Rowe": "Professional Services", "Pellier Group": "Professional Services",
}
ENGAGEMENT = {
    "ONSITE": "Contracted support", "PRODUCTION": "Event production",
    "VIDEO": "Video production", "INTEGRATION": "Project integration",
    "CONSULTING": "Consulting engagement", "ADMIN": "Other",
}

SOURCES = {"income": ["Invoice", "Invoice", "Invoice", "Manual Journal"],
           "cos": ["Bill", "Bill", "Bill", "Spend Money", "Manual Journal"],
           "expenses": ["Bill", "Bill", "Spend Money", "Manual Journal"],
           "other_income": ["Receive Money", "Manual Journal"]}

# Roughly how many ledger lines the seed emits a month. The real export runs
# near 1,900 a month (3,771 lines across July and August 2026), so the seed is
# a reduction: the same shape at about an eighth of the line count, to keep the
# data file small. The engine is sized for the real volume.
LINES_PER_MONTH = 250

gl_rows = []
fin = {}          # account -> month -> cents, signed the way Xero reports it
by_dept = {}      # (dept, cat) -> month -> cents
inv = 24500
job = 1400


_client_cache = {}


def clients_for(dept):
    """Clients for a department, with a cumulative weight ladder.

    Picking by weight rather than uniformly keeps each client's annual share
    stable while leaving the month to month lumpiness that event work actually
    has. Production revenue really does swing on whether a client ran an event
    that month, so the variance at client level is the point, not noise.
    """
    if dept in _client_cache:
        return _client_cache[dept]
    pool = [c for c in CLIENTS if dept in c[1]] or CLIENTS
    total = sum(c[2] for c in pool)
    ladder, run = [], 0.0
    for c in pool:
        run += c[2] / total
        ladder.append((run, c[0]))
    _client_cache[dept] = ladder
    return ladder


def pick_client(ladder):
    x = R.next()
    for edge, name in ladder:
        if x <= edge:
            return name
    return ladder[-1][1]


def emit(month, dept, cat, acc, amount_c, nlines):
    """Split one account-month into ledger lines and record them."""
    global inv, job
    if amount_c == 0:
        return
    y, m = int(month[:4]), int(month[5:])
    parts = split_exact(amount_c, [1.0 / nlines] * nlines) if nlines > 1 else [amount_c]
    pool = clients_for(dept)
    for p in parts:
        if p == 0:
            continue
        # spread across the month, weekdays only
        day = 1 + int(R.next() * (CAL_INDEX[month]["days"] - 1))
        d = date(y, m, day)
        while d.weekday() > 4:
            d -= timedelta(days=1)
        if cat in ("income", "other_income"):
            contact = pick_client(pool) if cat == "income" else ""
            credit, debit = p / 100.0, 0.0
        else:
            contact = R.pick(SUPPLIERS)
            credit, debit = 0.0, p / 100.0
        inv += 1
        job += 1 if R.next() > 0.7 else 0
        tag = JOB_TAG[dept]
        jobno = "J%02d-%04d [%s]" % (y % 100, job, tag)
        cc = "%s - %s" % (COST_CENTRE_CODE[dept], dept)
        roll = R.next()
        if roll < 0.004:
            cc, jobno = "", ""                      # lands in UNALLOCATED
        elif roll < 0.05:
            cc = ""                                 # falls back to the job tag
        gl_rows.append([d.isoformat(), acc["account"], contact,
                        round(debit, 2), round(credit, 2),
                        R.pick(SOURCES[cat]), jobno,
                        "INV-%05d" % inv if cat == "income" else "",
                        cc, ""])


CAL_INDEX = {m["key"]: m for m in CAL["months"]}

for month in ACTUAL_MONTHS:
    mi = MONTH_ORDER.index(CAL_INDEX[month]["m"])
    month_lines = 0
    for cat in ("income", "cos", "expenses", "other_income"):
        total_c = MONTH_TOTALS[cat][month]
        splits = dept_split(cat, mi)
        dept_c = split_exact(total_c, [splits.get(d, 0.0) for d in DEPTS])
        for dept, amt in zip(DEPTS, dept_c):
            if amt == 0:
                continue
            by_dept.setdefault((dept, cat), {})[month] = \
                by_dept.setdefault((dept, cat), {}).get(month, 0) + amt
            accs = ELIGIBLE[(cat, dept)]
            if not accs:
                continue
            ws = [weight(a) * R.jitter(0.85, 1.15) for a in accs]
            tw = sum(ws)
            acc_c = split_exact(amt, [w / tw for w in ws])
            for a, c in zip(accs, acc_c):
                if c == 0:
                    continue
                sign = -1 if cat in ("cos", "expenses") else 1
                fin.setdefault(a["account"], {})
                fin[a["account"]][month] = fin[a["account"]].get(month, 0) + sign * c
                n = max(1, min(8, int(abs(c) / 1800000) + 1))
                month_lines += n
                emit(month, dept, cat, a, c, n)
    # keep the line count near the target by merging nothing; report instead
print("ledger lines:", len(gl_rows), "over", len(ACTUAL_MONTHS), "months",
      "=", len(gl_rows) // len(ACTUAL_MONTHS), "a month")


# ---------------------------------------------------------------------- budget
def acct_factor(name, spread=0.12):
    """Stable per-account budget factor, so a rebuild gives the same budget."""
    h = 0
    for ch in name:
        h = (h * 131 + ord(ch)) % 100003
    return 1.0 - spread + (h % 1000) / 1000.0 * (2 * spread)


# Budget bias by department, from Part 3: consulting is frequently below budget
# (so it was budgeted high), and one department came in well above budget
# because its pipeline was understated (production).
BUDGET_BIAS = {"ONSITE": 1.01, "PRODUCTION": 0.93, "VIDEO": 1.00,
               "INTEGRATION": 1.04, "CONSULTING": 1.22, "ADMIN": 1.00}

budget = {}


def set_budget(acct, month, c):
    budget.setdefault(acct, {})[month] = int(round(c))


for a in ACCOUNTS:
    name = a["account"]
    if weight(a) == 0:
        continue
    dept = a["dept"] or "ADMIN"
    bias = BUDGET_BIAS.get(dept, 1.0) if a["cat"] in ("income", "cos") else 1.0
    f = acct_factor(name) * bias
    actual26 = fin.get(name, {})
    fy26_total = sum(actual26.get(m, 0) for m in MONTHS26)
    if fy26_total == 0:
        continue
    for i, m in enumerate(MONTHS26):
        set_budget(name, m, actual26.get(m, 0) * f)
    # FY27 budget: the FY26 shape, grown, smoothed towards a flat twelfth so a
    # budget is not simply last year's noise repeated.
    for i, m in enumerate(MONTHS27):
        share26 = (actual26.get(MONTHS26[i], 0) / fy26_total) if fy26_total else 1 / 12
        share = 0.75 * share26 + 0.25 / 12
        set_budget(name, m, fy26_total * GROWTH * f * share)

print("budget accounts:", len(budget))

# ----------------------------------------------------------------- utilisation
# Full time equivalents by department. Part 2 works an example at nineteen and
# a half employees across the business and says production has six full timers,
# and that consulting shows almost no FTE because the work is done by a casual
# engaged ad hoc. Those two are the anchors; the rest is modelled to the total.
FTE = {"ONSITE": 6.5, "PRODUCTION": 6.0, "VIDEO": 1.5,
       "INTEGRATION": 2.0, "CONSULTING": 0.5, "ADMIN": 3.0}
# Chargeable share of worked hours. Part 3 describes a team running in the mid
# forties against a target of around sixty five; consulting is that team.
UTIL = {"ONSITE": 0.72, "PRODUCTION": 0.68, "VIDEO": 0.62,
        "INTEGRATION": 0.58, "CONSULTING": 0.46, "ADMIN": 0.08}
LEAVE_SHARE = {"ONSITE": .075, "PRODUCTION": .065, "VIDEO": .07,
               "INTEGRATION": .08, "CONSULTING": .05, "ADMIN": .085}
# Leave clusters in January and drops away in the busy months.
LEAVE_SHAPE = [1.1, 0.7, 0.7, 0.6, 0.6, 1.4, 2.2, 0.8, 0.9, 1.2, 0.9, 0.9]

util_rows = []
for month in ALL_MONTHS:
    info = CAL_INDEX[month]
    mi = MONTH_ORDER.index(info["m"])
    # A person is paid for every weekday. Public holidays are paid but not
    # worked, so they come out first and are their own bucket, the way the
    # Employment Hero export reports them. What is left splits into leave and
    # worked time, and worked time splits into chargeable and non-chargeable.
    paid_hours = info["bd"] * 8
    holiday_hours = (info["bd"] - info["wd"]["VIC"]) * 8
    workable = paid_hours - holiday_hours
    for dept in DEPTS:
        ph = holiday_hours * FTE[dept]
        remaining = workable * FTE[dept]
        leave = remaining * LEAVE_SHARE[dept] * LEAVE_SHAPE[mi]
        worked = remaining - leave
        # utilisation tracks the department's own season
        season = SHAPE[dept][mi] * 12
        chg = worked * min(0.95, UTIL[dept] * (0.82 + 0.18 * season))
        util_rows.append([month, dept, round(chg, 1), round(worked - chg, 1),
                          round(leave, 1), round(ph, 1), FTE[dept]])

# ---------------------------------------------------------------- the clients
client_meta = []
for name, depts, size in CLIENTS:
    client_meta.append({
        "contact": name,
        # Part 2: the display name on the chart is the listing name, not the
        # trading name, and it is matched by hand. Seeded equal so the mapping
        # is visible and editable rather than hidden.
        "display": name,
        "industry": INDUSTRY.get(name, "Other"),
        "engagement": ENGAGEMENT[depts[0]],
        "primaryDept": depts[0],
    })

# ------------------------------------------------------------ revenue schedule
# Part 3: onsite is budgeted per role rather than per client; the other
# departments are budgeted per client, with anything unlisted grouped into a
# single new business line.
ONSITE_ROLES = [
    ("Meridian Bank Group", "Team Leader, level 3"),
    ("Meridian Bank Group", "Support technician x2"),
    ("Meridian Bank Group", "Boardroom AV support"),
    ("Parkin Vale Advisory", "Onsite technician"),
    ("Parkin Vale Advisory", "Help desk, shared"),
    ("Southgate Mutual", "Managed service, tier 2"),
    ("Alderway Legal", "Onsite technician, part week"),
    ("Corell Industries", "Maintenance contract"),
    ("Ingram Foods Australia", "Help desk, shared"),
    ("Marchmont Insurance", "Onsite technician"),
]

sched = []
for dept in ["ONSITE", "PRODUCTION", "VIDEO", "INTEGRATION", "CONSULTING"]:
    if dept == "ONSITE":
        for client, role in ONSITE_ROLES:
            sched.append({"dept": dept, "kind": "role", "client": client, "line": role})
    else:
        for name, depts, size in CLIENTS:
            if depts[0] == dept:
                sched.append({"dept": dept, "kind": "client", "client": name, "line": name})
        sched.append({"dept": dept, "kind": "newbiz", "client": "",
                      "line": "New business, not separately listed"})

# ------------------------------------------------------------------- write out
def js(varname, payload, note):
    body = json.dumps(payload, separators=(",", ":"))
    return ("// %s\n// Built %s by portal/build/build_data.py. Do not edit by hand.\n"
            "window.%s = %s;\n" % (note, BUILT, varname, body))


OUT.mkdir(parents=True, exist_ok=True)

SEED_NOTE = ("Seed data. The FY26 totals, the July 2025 month and the August 2026 "
             "category totals are the real Xero figures. Everything inside them is "
             "modelled from the seasonality and margin norms in the Monthly Reporting "
             "manual. Replace it by loading the real Xero export on the Data Loaders page.")

(OUT / "CTS_cal_data.js").write_text(
    js("CTS_CAL", CAL, "CTS financial calendar, 1 July to 30 June"))

(OUT / "CTS_accounts_data.js").write_text(js("CTS_ACCOUNTS", {
    "meta": {"source": "Xero chart of accounts, read 3 September 2026",
             "count": len(ACCOUNTS), "built": BUILT, "seed": False},
    "accounts": [{"name": a["account"], "group": a["group"], "cat": a["cat"],
                  "sub": a["sub"], "dept": a["dept"], "division": a["division"]}
                 for a in ACCOUNTS],
}, "CTS chart of accounts, 190 accounts read from Xero"))

(OUT / "CTS_gl_data.js").write_text(js("CTS_GL", {
    "meta": {"seed": True, "note": SEED_NOTE, "built": BUILT,
             "rows": len(gl_rows), "months": ACTUAL_MONTHS,
             "realLineRate": "about 1,900 a month in the live export; the seed runs "
                             "about 220 to keep the file small"},
    "cols": ["date", "account", "contact", "debit", "credit", "source",
             "jobNo", "invoiceNo", "costCentre", "description"],
    "rows": gl_rows,
}, "CTS general ledger transactions"))

(OUT / "CTS_fin_data.js").write_text(js("CTS_FIN", {
    "meta": {"seed": True, "note": SEED_NOTE, "built": BUILT,
             "basis": "accrual", "currency": "AUD",
             "control": {"FY26": FY26, "2025-07": JUL25, "2026-08": AUG26}},
    "actual": {k: v for k, v in fin.items()},
    "budget": budget,
    "months": ALL_MONTHS,
}, "CTS profit and loss by account by month"))

(OUT / "CTS_util_data.js").write_text(js("CTS_UTIL", {
    "meta": {"seed": True, "built": BUILT, "hoursPerDay": 8, "state": "VIC",
             "note": "Hours by department by month, split the four ways the "
                     "Employment Hero export splits them. Anchors: about 19.5 full time "
                     "equivalents across the business and six in production, both "
                     "from Monthly Reporting Part 2; consulting running in the mid "
                     "forties against a 65% target, from Part 3."},
    "cols": ["month", "dept", "chargeable", "nonChargeable", "leave",
             "publicHoliday", "fte"],
    "rows": util_rows,
}, "CTS utilisation hours by department"))

(OUT / "CTS_clients_data.js").write_text(js("CTS_CLIENTS", {
    "meta": {"seed": True, "built": BUILT,
             "note": "Client names in the seed are fictional. Real client names "
                     "arrive with the real GL export. The display name is the one "
                     "shown on the charts and is matched by hand, per Part 2."},
    "clients": client_meta,
    "schedule": sched,
}, "CTS client reference data and the revenue schedule lines"))

print("wrote", len(list(OUT.glob("*.js"))), "data files to", OUT)
for f in sorted(OUT.glob("*.js")):
    print("   %-28s %6.0f KB" % (f.name, f.stat().st_size / 1024))
