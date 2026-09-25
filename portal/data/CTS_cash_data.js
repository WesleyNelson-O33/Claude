// CTS cash settings and balances. Built from 10 Cash and Commitments.xlsx on
// Build. This starting set is SEED: fictional balances so the Cash Flow page
// has something to show. Replace it with the real month end figures.
window.CTS_CASH = {
  meta: { built: "2026-09-25", source: "seed", seed: true, version: 1 },
  settings: { debtorDays: 45, creditorDays: 30, basFrequency: "Quarterly", superTiming: "With each pay",
              paygInstalment: 0, facility: 0, openingAR: null, openingAP: null, cardsPaidInFull: "Yes" },
  balances: [
    { month: "2026-07", account: "Operating account", balance: 812400.00, note: "seed" },
    { month: "2026-07", account: "Savings account", balance: 60000.00, note: "seed" },
    { month: "2026-08", account: "Operating account", balance: 794150.00, note: "seed" },
    { month: "2026-08", account: "Savings account", balance: 60000.00, note: "seed" },
  ],
  cards: [
    { month: "2026-08", card: "Company card 1", holder: "Director", limit: 30000, balance: 8420.55, paymentDay: 15, note: "seed" },
    { month: "2026-08", card: "Company card 2", holder: "Operations", limit: 15000, balance: 3190.20, paymentDay: 15, note: "seed" },
  ],
  commitments: [
    { name: "Office rent", category: "Rent", amount: 28500, gst: "Yes", frequency: "Monthly", next: "2026-09-22", end: null, inPnl: "Yes", note: "seed" },
    { name: "Business insurance", category: "Insurance", amount: 42000, gst: "Yes", frequency: "Annual", next: "2027-03-01", end: null, inPnl: "Yes", note: "seed, paid once a year" },
    { name: "Equipment loan", category: "Loan", amount: 6500, gst: "No", frequency: "Monthly", next: "2026-09-28", end: "2028-06-28", inPnl: "No", note: "seed, principal only; interest is in the P&L" },
  ],
};
