"""Capture the CTS chart of accounts and derive each account's division."""
import json
import re
from pathlib import Path

INCOME = [
    "Ad Hoc Service Labour", "Closing Work in Progress", "Conference Call Charge",
    "Contract Help Desk", "Contract Maintenance", "Contract Support Staff",
    "Credit Card Collected - CONS", "Credit Card Collected - ONS",
    "Credit Card Collected - PRD", "Credit Card Collected - VID",
    "Credit Card Collected-INTEGRATION", "Credit Card Fees - ADMIN",
    "Equipment Hires", "Equipment Sales", "Freight & Travel - ADMIN",
    "Freight & Travel - CONS", "Freight & Travel - INTEGRATION",
    "Freight & Travel - ONS", "Freight & Travel - PRD", "Freight & Travel - VID",
    "Installation Labour", "Meal & Travel - ADMIN", "Meal & Travel - CONS",
    "Meal & Travel - INTEGRATION", "Meal & Travel - ONS", "Meal & Travel - PRD",
    "Meal & Travel - VID", "Miscellaneous Income", "Mobile Phone Plans",
    "Opening Work in Progress", "Production Labour", "Subscriptions & Licences - Income",
    "Tech Consulting/Project Manage", "Tech Event Management", "Video Labour",
]

COST_OF_SALES = [
    "Credit Card Paid - ADMIN", "Credit Card Paid - CONS", "Credit Card Paid - INTEGRATION",
    "Credit Card Paid - ONS", "Credit Card Paid - PRD", "Credit Card Paid - VID",
    "Damage Waiver", "Direct Freight - ADMIN", "Direct Freight - CONS",
    "Direct Freight - INTEGRATION", "Direct Freight - ONS", "Direct Freight - PRD",
    "Direct Freight - VID", "Direct Mobile - CONS", "Direct Mobile - INTEGRATION",
    "Direct Mobile - ONS", "Direct Mobile - PRD", "Direct Mobile - VID",
    "Direct Salaries - ADMIN", "Direct Salaries - CONS", "Direct Salaries - INTEGRATION",
    "Direct Salaries - ONS", "Direct Salaries - PRD", "Direct Salaries - VID",
    "Direct Superannuation - ADMIN", "Direct Superannuation - CONS",
    "Direct Superannuation - ONS", "Direct Superannuation - PRD",
    "Direct Superannuation - VID", "Direct Superannuation-INTEGRATION",
    "Direct Travel - ADMIN", "Direct Travel - CONS", "Direct Travel - INTEGRATION",
    "Direct Travel - ONS", "Direct Travel - PRD", "Direct Travel - VID",
    "Direct Wages - CONS", "Direct Wages - INTEGRATION", "Direct Wages - ONS",
    "Direct Wages - PRD", "Direct Wages - VID", "Direct Workers Comp - ADMIN",
    "Direct Workers Comp - CONS", "Direct Workers Comp - INTEGRATION",
    "Direct Workers Comp - ONS", "Direct Workers Comp - PRD", "Direct Workers Comp - VID",
    "Discounts Given", "Equipment - Hires", "Equipment - Purchase",
    "Equipment -Maintenance Support", "Service - hires", "Stamp Duty",
    "Sub-Contract Labour - ADMIN", "Sub-Contract Labour - CONS",
    "Sub-Contract Labour - INTEGRATION", "Sub-Contract Labour - ONS",
    "Sub-Contract Labour - PRD", "Sub-Contract Labour - VID",
    "Subscriptions & Licences - Expense", "Supplier Per Diems - ADMIN",
    "Supplier Per Diems - CONS", "Supplier Per Diems - INTEGRATION",
    "Supplier Per Diems - ONS", "Supplier Per Diems - PRD", "Supplier Per Diems - VID",
]

OTHER_INCOME = [
    "Cash Flow Boost", "Employee FBT Contribtution", "Gain on lease termination",
    "Income Other", "Interest Income", "Interest Received", "JobKeeper Subsidy",
    "JobSaver Subsidy", "Rental Income",
]

EXPENSES = [
    "1300 Number", "Accounting Fees", "Advertising", "Bad Debts", "Bank Charges",
    "Bank Revaluations", "Business Advisory", "Business Combined", "Client Entertainment",
    "Credit Card Fees Paid", "Custodian Vaults", "Depreciation", "Depreciation - ROU",
    "Discounts Taken", "Disposal of fixed assets", "Dividend paid [93000]",
    "Donations / Charity", "Dues & Subscriptions", "Electricity", "Fax", "Filing Fee",
    "Foreign Currency Translation*", "Freight Paid", "Fringe Benefits Tax",
    "IT Network Service & Support", "Income Tax Expense", "Indirect Salaries - ADMIN",
    "Indirect Salaries - CONS", "Indirect Salaries - INTEGRATION",
    "Indirect Salaries - ONS", "Indirect Salaries - PRD", "Indirect Salaries - VID",
    "Indirect Wages - ADMIN", "Indirect Wages - CONS", "Indirect Wages - INTEGRATION",
    "Indirect Wages - ONS", "Indirect Wages - PRD", "Indirect Wages - VID",
    "Instant Asset Write Off", "Insurance", "Interest Expense", "Internet",
    "Leave expense (Annual)", "Leave expense (Long Service)", "Legal Fees",
    "Mobile Phones", "Office Cleaning", "Office Phones", "Office Rentals",
    "Office Supplies", "Office Supplies [61500]", "Other Expenses [64900]",
    "Other Insurance", "Other Telco Expenses", "Pager", "Payroll Processing Fee",
    "Payroll Tax", "Postage", "Printing", "Professional Indemnity",
    "Realised Currency Gains", "Recruitment", "Rent", "Rental Outgoings",
    "Repairs & Maintenance", "Replacements", "SGC Penalties", "Staff Amenities",
    "Staff Entertainment", "Stamp Duty (Indirect)", "Storage Fees", "Stripe Fees",
    "Superannuation", "Taxis/Parking", "Training Material & Courses",
    "Travel & Per Diems", "Travel Insurance", "Under/(Over) provision for tax",
    "Unrealised Currency Gains", "Workers' Compensation",
]

DIVISIONS = {
    "ONS": "Onsite", "PRD": "Production", "VID": "Video",
    "CONS": "Consulting", "INTEGRATION": "Integration", "ADMIN": "Admin",
}
# Suffix may be " - CODE" or "-CODE"; anchored to the end so "Direct Travel - CONS"
# is not mistaken for something else.
SUFFIX = re.compile(r"\s*-\s*(" + "|".join(DIVISIONS) + r")$")


def division(name):
    m = SUFFIX.search(name)
    return DIVISIONS[m.group(1)] if m else "Unallocated"


accounts = []
for group, names in (("Income", INCOME), ("Cost of Sales", COST_OF_SALES),
                     ("Other Income", OTHER_INCOME), ("Expenses", EXPENSES)):
    for name in sorted(names):
        accounts.append({"account": name, "group": group, "division": division(name)})

out = Path("/home/user/Claude/reporting/data/accounts.json")
out.write_text(json.dumps(accounts, indent=2) + "\n")

from collections import Counter
print(f"{len(accounts)} accounts")
print("by group:     ", dict(Counter(a['group'] for a in accounts)))
print("by division:  ", dict(Counter(a['division'] for a in accounts)))
