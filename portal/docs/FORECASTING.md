# Forecasting: the design

CTS does not forecast today. This is the plan for building it into the portal
so that every page that shows a budget also shows a forecast, and cash is
forecast off the same numbers. It is a design to agree, not a build that has
happened. The parts are ordered so each one is useful on its own.

## The one idea

A forecast is a third series beside Actual and Budget.

| Series | What it is | When it changes |
|---|---|---|
| Actual | What Xero says happened | Every month, for the month just closed |
| Budget | The target set before the year started | Once a year, with the odd reforecast |
| Forecast | The best current view of the months still to come | Every month, at Build |

Today the portal fills the months after the reporting month with budget and
calls it the full year view. With a forecast, those months come from the
forecast instead, and budget stays as the thing to compare against. Nothing
else about the portal changes shape: the P&L, the department P&Ls, the
revenue schedule and the bridge all keep their columns and gain one.

The forecast is built the same way everything else in the portal is built:
from a template finance controls, at Build, with every figure carrying a note
of where it came from. A forecast line that rests on a run rate is flagged the
way an allocated figure that rests on a placeholder split base is flagged now.

## Revenue forecast, by department by month

Three layers, each covering the horizon the others cannot.

**1. Confirmed work.** Orders in OnRent with an event date, and Qwilr quotes
accepted with a delivery date. These are revenue in the month the work lands.
Counted in full, less the historical cancellation rate once there is history
to measure it.

**2. Weighted pipeline.** Open deals in Zoho, at the deal's amount times its
probability, landing in the month of its expected close. The probability by
stage is Zoho's until there are twelve months of won and lost history in the
portal, then it is measured: what share of deals at each stage actually
closed, and how long after the stated close date. The Pipeline page becomes
the input to this, which is what it is for.

**3. Baseline.** Revenue that never goes through the CRM or a quote: repeat
onsite work, small jobs, the run of business. Measured as the trailing twelve
months by department, shaped by the seasonal profile already in the config,
grown by a percentage typed on the forecast template. This is what the
forecast rests on beyond the pipeline horizon.

How they combine, by month out from the reporting month:

| Horizon | Forecast is |
|---|---|
| Months 1 to 3 | Confirmed plus weighted pipeline, plus the baseline share that history says is never in the pipeline that far out |
| Months 4 to 12 | Baseline, replaced by confirmed plus weighted wherever that is higher |
| Any month | A typed override wins, and is flagged as typed |

The "never in the pipeline" share is one number per department, measured
once from history and typed on the template. It stops the near months
double counting and stops them looking empty just because a department books
its work late.

## Cost forecast

Costs follow revenue or follow time, and the template says which per account.

| Cost | Method | Where the driver comes from |
|---|---|---|
| Permanent labour | Headcount times monthly cost, month by month | The staff list from Earnings, plus a tab for starters, leavers and pay changes |
| Casual and contract labour | Percentage of department revenue | Trailing twelve months, per department |
| Other cost of sales (hire, freight, subcontractors) | Percentage of department revenue | Trailing twelve months, per department |
| Fixed overheads (rent, insurance, software, leases) | The contract amount on its dates | The commitments tab of the cash template |
| Variable overheads | Trailing three month run rate | Actuals |
| Anything | A typed amount | The overrides tab, flagged as typed |

The result is a P&L forecast by account by department by month. It goes
through the same overhead allocation as the actuals, so the department
forecasts after the split reconcile to the company forecast, the same way
the actuals do now.

## Cash flow forecast

Cash is where this pays off, and CTS has nothing today. Two horizons, one
model, and they agree where they overlap.

**Thirteen weeks, direct.** Week by week: opening bank, receipts, payments,
closing bank. This is the one that answers "can we make payroll and the BAS in
the same fortnight".

| Line | Comes from |
|---|---|
| Opening bank | Xero bank balances as at the reporting date |
| Receipts from debtors | Xero Aged Receivables, invoice by invoice, on the due date shifted by how late each client actually pays (measured from paid invoice history, defaulting to terms) |
| Receipts from forecast sales | The revenue forecast, invoiced on the department's invoicing pattern (onsite at the event, integration in deposits and milestones, consulting monthly) and collected on the client's days to pay |
| Payments to creditors | Xero Aged Payables, bill by bill, on the due date |
| Payroll | The pay cycle and the current run cost from Earnings |
| Superannuation | SG super from Earnings, paid on its due dates |
| BAS | GST collected on forecast receipts less GST on forecast payments, and PAYG withholding, on the lodgement dates |
| Commitments | Rent, loans, leases, insurance, subscriptions, from the commitments tab, each with its amount, frequency, next date and GST treatment |
| Capital spending and tax | Typed on the same tab |

**Twelve months, indirect.** Month by month from the P&L forecast: profit,
add back depreciation, less the movement in debtors and creditors on the
measured days to pay, less capital spending, loan repayments, tax and
distributions. This is the one that answers "what does cash look like at
Christmas and at year end".

The first three months of the indirect forecast are made to agree with the
thirteen weeks, so there is one closing cash line and not two.

What the Cash page shows: the closing cash line across the horizon with the
low point marked, the receipts and payments that make it, the ten largest
receipts still to come with their expected week, and headroom against any
facility. The Dashboard gets a closing cash tile beside the profit ones.

## What it needs from the systems

Four more templates, in the same shape as the rest: a Read me tab that says
what to run, a data tab that is a straight paste.

| Template | Run in | Feeds |
|---|---|---|
| 09 Xero Aged Receivables | Xero, Aged Receivables Detail, as at month end | Debtor receipts, days to pay |
| 10 Xero Aged Payables | Xero, Aged Payables Detail, as at month end | Creditor payments |
| 11 Cash and Commitments | Typed, standing | Bank balances at month end, the commitments list, the tax calendar, the invoicing pattern per department, any facility |
| 12 Forecast | Typed, standing, revised monthly | Method per account, the baseline growth and never in pipeline percentages, headcount changes, overrides |

The three pipeline exports already in the plan (Zoho, OnRent, Qwilr) carry
the revenue side. Days to pay by client needs one more Xero run, Receivable
Invoice Detail for the last twelve months with paid dates, done once and then
yearly. Each of these is a report Xero already has; none of them needs a new
report built. The exact column layout of each is matched to the real export
the same way the others will be, from one sample.

## What it does to the portal

New pages, in the same style as the rest:

- **Forecast.** The company and department P&L with Actual, Forecast and
  Budget across the year, the full year landing, and a flag on every forecast
  line saying which method made it.
- **Revenue Forecast.** Department by month, stacked by layer: confirmed,
  weighted pipeline, baseline, typed. The gap between forecast and budget by
  month is the sales conversation.
- **Cash Flow.** The thirteen week and twelve month views, the closing cash
  line, the low point, the large receipts still to come.
- **Forecast Accuracy.** Each month's forecast against what then happened.
  The archive makes this free: every month's folder holds the forecast that
  stood at the time, so after three months the portal can say how far out the
  three month forecast tends to be, by department. That is what makes a
  forecast get better.

Existing pages that gain a column or a line: Budget vs Actual (forecast beside
budget, and the risk flag also tests forecast against budget), Revenue Summary
and Revenue Schedule (the run to year end comes from the forecast, not the
budget), P&L Spread, Dashboard (full year forecast tile, closing cash tile),
and the tier one and tier three emails (full year landing, cash low point,
what moved since last month's forecast).

Governance comes from what is already there. The forecast is rebuilt at Build
and frozen in the archive, so there is always exactly one forecast for a
month and it is the one that was reported. The build log records which
methods and overrides were in force. A "what moved" walk from last month's
forecast to this month's is the first thing the Forecast page shows.

## The order to build it

Each step stands on its own and needs only what is listed.

1. **P&L forecast on run rate and seasonality, with overrides.** Needs the
   12 Forecast template, which starts with sensible defaults typed in. No new
   system exports. Gives the Forecast page, the forecast column everywhere,
   and the full year landing. This is the step that changes how the monthly
   report reads, and it is two to three days of build.
2. **Twelve month cash on the indirect method.** Needs the 11 Cash and
   Commitments template, the opening bank balances and the commitments list.
   Gives the Cash Flow page's monthly view and the cash tile.
3. **Pipeline driven revenue.** Needs the three real exports from Zoho, OnRent
   and Qwilr matched to their templates, which is already the next input on
   the list. Gives the Revenue Forecast page and replaces the baseline in the
   near months with what is actually booked and quoted.
4. **Thirteen week direct cash.** Needs the aged receivables and payables
   pastes and one run of paid invoice history. Gives the weekly view and the
   low point.
5. **Forecast accuracy.** Needs nothing but three months of archive.

Step 3 can run ahead of step 2 if the exports arrive first; nothing in the
order is forced except that 1 comes before 2 and 4 comes after 1.

## Assumptions to correct

These are the defaults the build will start from unless told otherwise.

- The forecast horizon is rolling twelve months from the reporting month,
  and the full year figure is the current financial year.
- Payroll is fortnightly, super is paid quarterly by the 28th, BAS is
  quarterly on the standard dates, and there is no PAYG instalment. All four
  are settings on the cash template.
- Invoicing patterns: onsite and production invoice on completion of the
  event, integration takes a deposit on acceptance and the balance on
  handover, video and consulting invoice monthly. Terms are thirty days.
- There is no overdraft or facility. If there is, its limit goes on the cash
  template and headroom is shown against it.
- Budget stays fixed for the year. A reforecast that replaces the budget is a
  paste into the Budget template, which is how it works today.
