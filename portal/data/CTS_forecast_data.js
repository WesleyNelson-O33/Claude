// CTS forecast settings. Built from 09 Forecast.xlsx on Build; this is the
// starting set: defaults by category, no growth, no overrides.
window.CTS_FORECAST = {
  meta: { built: "2026-09-25", source: "default", version: 1,
          note: "Methods by category. Income: same month last year. Cost of sales: percentage of department revenue, labour lines on run rate. Overheads: run rate of the last three months." },
  assumptions: { horizonMonths: 12, runRateMonths: 3, enabled: true, growth: { "default": 0 },
                 pipeline: { enabled: true, nearMonths: 3, cancellationRate: 0, countWon: false,
                             never: { "default": 0.3, ONSITE: 0.6, PRODUCTION: 0.3, VIDEO: 0.3, INTEGRATION: 0.1, CONSULTING: 0.2, ADMIN: 1 },
                             leadMonths: { "default": 0, INTEGRATION: 1 },
                             orderStatuses: "Confirmed, Booked, In progress, Completed", quoteStatuses: "Accepted", probability: {} } },
  methods: [
    { match: "Category: Income",         method: "seasonal",    param: null, note: "Same month last year, grown by the department's growth assumption" },
    { match: "Category: Cost of Sales",  method: "revenue_pct", param: null, note: "Measured share of department revenue over the last twelve months" },
    { match: "Subcategory: Direct Salaries",       method: "runrate", param: null, note: "Permanent labour follows time, not revenue" },
    { match: "Subcategory: Direct Wages",          method: "runrate", param: null, note: "" },
    { match: "Subcategory: Direct Superannuation", method: "runrate", param: null, note: "" },
    { match: "Subcategory: Direct Workers Comp",   method: "runrate", param: null, note: "" },
    { match: "Category: Expenses",       method: "runrate",     param: null, note: "Average of the last three months" },
    { match: "Category: Other Income",   method: "runrate",     param: 12,   note: "Average of the last twelve months" },
    { match: "Category: Other Expenses", method: "runrate",     param: 12,   note: "" },
  ],
  overrides: [],
};
