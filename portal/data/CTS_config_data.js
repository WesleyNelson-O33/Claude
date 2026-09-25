// CTS Business Intelligence Portal - configuration
//
// This is the one data file meant to be edited by hand. Everything else in
// portal/data is generated. Each block records where its numbers came from, so
// a figure that was sourced can be told apart from one that is a placeholder.
window.CTS_CONFIG = {
  ORG: {
    name: "Corporate Technology Services Pty Ltd",
    short: "CTS",
    currency: "AUD",
    fyStartMonth: 7,                 // 1 July to 30 June
    basis: "accrual",
    reportingMonth: "2026-08",       // the one control that moves each month
    priorFY: 26,
    currentFY: 27,
    portalUrl: "",   // the SharePoint link or synced folder path, from the Config Reporting tab; goes in every email
  },

  // ------------------------------------------------------------ departments
  // Order, tags and the ADMIN mapping are confirmed from DEPT_TAGS in
  // pack/build_pack.py. Margin norms are from Monthly Reporting Part 3.
  // Utilisation targets: only one is documented (around 65%, Part 3), so the
  // rest carry that same figure and are flagged as assumed.
  DEPARTMENTS: [
    { code: "ONSITE", short: "Onsite", long: "Onsite and managed services", order: 1,
      isRevenue: true, utilTarget: 0.65, utilTargetSource: "documented",
      gmNorm: 0.45, npNorm: 0.10,
      normLabel: "Support and managed services",
      normNote: "Part 3 calls this support and managed services: always profitable, gross margin around 45 per cent, net profit above 10 per cent. Mapped to Onsite here; confirm that mapping.",
      normAssumed: true,
      tags: ["ONSITE", "ONS"], suffix: "ONS" },

    { code: "PRODUCTION", short: "Production", long: "Event production", order: 2,
      isRevenue: true, utilTarget: 0.65, utilTargetSource: "assumed",
      gmNorm: 0.55, npNorm: null,
      normNote: "Very seasonal. Gross margin should be 55 per cent or better; in the 40s is worth investigating, usually equipment or labour.",
      normAssumed: false,
      tags: ["PRODUCTION", "PRD"], suffix: "PRD" },

    { code: "VIDEO", short: "Video", long: "Video production", order: 3,
      isRevenue: true, utilTarget: 0.65, utilTargetSource: "assumed",
      gmNorm: null, npNorm: null,
      normNote: "No norm is written down for video. Part 3 treats production and video separately in the revenue schedule but gives the margin norm for production only.",
      normAssumed: true,
      tags: ["VIDEO", "VID"], suffix: "VID" },

    { code: "INTEGRATION", short: "Integration", long: "Systems integration", order: 4,
      isRevenue: true, utilTarget: 0.65, utilTargetSource: "assumed",
      gmNorm: 0.15, npNorm: null,
      normNote: "Gross margin should be at least 15 per cent. Net profit is skewed by how few people work in it.",
      normAssumed: false,
      tags: ["INTEGRATION", "INT"], suffix: "INTEGRATION" },

    { code: "CONSULTING", short: "Consulting", long: "Technology consulting", order: 5,
      isRevenue: true, utilTarget: 0.65, utilTargetSource: "documented",
      gmNorm: null, npNorm: null,
      normNote: "Frequently below budget. A loss in most months is normal, not a red flag. Runs in the mid forties on utilisation against a target of around 65 per cent.",
      normAssumed: false,
      tags: ["CONSULTING", "CONS"], suffix: "CONS" },

    { code: "ADMIN", short: "Admin", long: "Office and administration", order: 6,
      isRevenue: false, utilTarget: null, utilTargetSource: null,
      gmNorm: null, npNorm: null,
      normNote: "Carries the overheads that are pushed out to the other departments. Its own result is only meaningful before allocation.",
      normAssumed: false,
      tags: ["CTS", "ADMIN"], suffix: "ADMIN" },

    { code: "UNALLOCATED", short: "Unallocated", long: "No department on the line", order: 7,
      isRevenue: false, utilTarget: null, gmNorm: null, npNorm: null,
      normNote: "A GL line whose Cost Centres column is blank and whose job number carries no bracket tag. Should be zero. Anything here is a coding gap to fix in Xero.",
      normAssumed: false,
      tags: [], suffix: null },
  ],

  // How a GL line resolves to a department. Cost Centres first, then the
  // bracket tag on the job number. Confirmed from the Cleanup sheet formulas.
  DEPT_RESOLUTION: {
    order: ["costCentre", "jobTag"],
    note: "Department comes from the Cost Centres column, which carries the department name outright. If that is blank the pack falls back to the bracket tag on the Job Numbers column. Unmatched lines read UNALLOCATED.",
  },

  // -------------------------------------------------------- overhead splits
  // Three bases, each overhead line labelled with the one that applies.
  // The 3 Way percentages are confirmed from pack/build_bridge.py (34/33/33).
  // Staff and Office Dept are read live off the budget workbook in the real
  // process and are NOT in this repository, so they are placeholders here and
  // the portal flags every figure that depends on them.
  SPLIT_BASES: [
    { code: "3WAY", label: "3 Way", source: "confirmed",
      note: "Split roughly evenly three ways, near enough a third each. Confirmed at 34/33/33 from the budget bridge.",
      groups: { "PRD/VID": 0.34, "ONS": 0.33, "CONS/INT": 0.33 } },
    { code: "STAFF", label: "Staff", source: "placeholder",
      note: "Split by headcount, so it follows where the people actually are. The live percentages sit in the budget workbook and are not in this repository. These are placeholders weighted to reflect that consulting and integration together carry only a fraction, because almost nobody is employed there.",
      groups: { "PRD/VID": 0.40, "ONS": 0.47, "CONS/INT": 0.13 } },
    { code: "OFFICE", label: "Office Dept", source: "placeholder",
      note: "Weighted heavily to one department, for costs that mostly belong there. Live percentages are in the budget workbook and are not in this repository.",
      groups: { "PRD/VID": 0.20, "ONS": 0.65, "CONS/INT": 0.15 } },
  ],

  // The budget groups the splits are set against are three, not six. These two
  // sub-splits break them out. Defaults are the FY23 actuals on the PRD and
  // CONS tabs, per pack/build_bridge.py.
  SUB_SPLITS: {
    productionShareOfPrdVid: 0.79,
    integrationShareOfConsInt: 0.72,
    source: "confirmed",
    note: "Production takes 0.79 of the production and video group; integration takes 0.72 of the consulting and integration group. Video and consulting take the rest.",
  },

  // Which basis each overhead subcategory is allocated on. In the live process
  // the basis is labelled against every row on the allocation tab. This map is
  // derived from the account meaning and is meant to be corrected in place.
  SPLIT_ASSIGNMENT: {
    source: "derived",
    default: "3WAY",
    bySub: {
      "Rent": "OFFICE", "Rental Outgoings": "OFFICE", "Office Rentals": "OFFICE",
      "Electricity": "OFFICE", "Office Cleaning": "OFFICE", "Storage Fees": "OFFICE",
      "Internet": "OFFICE", "Office Phones": "OFFICE", "Office Supplies": "OFFICE",
      "Printing": "OFFICE", "Postage": "OFFICE", "Repairs & Maintenance": "OFFICE",
      "Superannuation": "STAFF", "Payroll Tax": "STAFF", "Workers' Compensation": "STAFF",
      "Recruitment": "STAFF", "Staff Amenities": "STAFF", "Staff Entertainment": "STAFF",
      "Training Material & Courses": "STAFF", "Payroll Processing Fee": "STAFF",
      "Leave expense (Annual)": "STAFF", "Fringe Benefits Tax": "STAFF",
      "Mobile Phones": "STAFF", "Travel & Per Diems": "STAFF", "Taxis/Parking": "STAFF",
      "IT Network Service & Support": "STAFF",
    },
  },

  // Lines the reporting manual says are handled by hand every month.
  MANUAL_LINES: [
    { sub: "Legal Fees", label: "Legal fees split",
      note: "One GL code, two lines in the report. Employsure is a fixed $800 a month retainer and is always there; Eastern Bell is the actual lawyers and varies. The retainer is entered cumulatively on the year to date and quarter tabs, which is the part that goes wrong.",
      fixed: { "Employsure": 800, "Eastern Bell": null } },
  ],

  // ----------------------------------------------------------- risk flagging
  // Confirmed from the risk() formula in pack/build_pack.py. The materiality
  // floor is tested before the percentage, so a large percentage on a small
  // dollar variance is still Low.
  RISK: {
    materialityFloor: 5000,
    mediumPct: 0.10,
    highPct: 0.25,
    highDollar: 50000,
    source: "confirmed",
  },

  // --------------------------------------------------------- utilisation
  UTIL: {
    hoursPerDay: 8,
    state: "VIC",
    source: "confirmed",
    note: "Utilisation is chargeable over chargeable plus non-chargeable, with leave excluded from both sides. Full time equivalent is worked hours over available hours, where available is working days for the selected state times hours per day. The monthly reporting graphs file instead works on a rolling twelve months against 2,080 or 2,088 hours, so the two figures are expected to differ.",
    rollingMonths: 12,
    annualHoursCheck: [2080, 2088],
  },

  // ---------------------------------------------------------- what is normal
  // Monthly Reporting Part 3. This is the part that is hardest to pick up from
  // a spreadsheet and it is what tells you whether a number is worth a comment.
  SEASON: [
    { months: [8, 9], label: "Results season",
      note: "August and September. Production should be profitable; if not, that is worth a comment." },
    { months: [10, 11, 12], label: "AGM peak",
      note: "October to December, driven by annual general meeting season. Production's most profitable quarter." },
    { months: [12, 1], label: "Quiet", note: "December and January are quiet." },
    { months: [2], label: "Picks up", note: "February picks up again." },
    { months: [3], label: "Quiet", note: "March is quiet." },
    { months: [5, 6], label: "Year end peak",
      note: "May and June are the peak for integration and consulting, as the financial year closes." },
  ],

  // --------------------------------------------------------------- reporting
  COMMENTARY: {
    topClientHeadings: ["Revenue increased", "Revenue decreased",
                        "New to the list", "Dropped off the list"],
    note: "Compare dollars, not percentages: management prefer dollar value against dollar value, because percentages read as more precise than they are. Where consulting revenue did not come in, say whether the deal was lost or moved to a later month.",
  },

  CATEGORY_ORDER: ["Income", "Cost of Sales", "Expenses", "Other Income", "Other Expenses"],

  // ------------------------------------------------------- the monthly email
  // Who gets which tier. Rebuilt from the Distribution tab of 08 Config.xlsx;
  // addresses here are placeholders. Tier 1 executive, 2 department head,
  // 3 finance. Nothing is sent by the portal itself: see docs/AUTOMATION.md.
  DISTRIBUTION: [
    { name: "Duncan", email: "", tier: 1, dept: null, send: true, note: "Full result, cost and margin, commentary" },
    { name: "Graham", email: "", tier: 1, dept: null, send: true, note: "" },
    { name: "Jordan", email: "", tier: 2, dept: "CONSULTING", send: true, note: "Own department in depth, one line on the rest" },
    { name: "Production manager", email: "", tier: 2, dept: "PRODUCTION", send: false, note: "No address yet" },
    { name: "Danica Nelson", email: "", tier: 3, dept: null, send: true, note: "Finance: the controls and data quality" },
  ],

  VERSION: "1.0.0",
};
