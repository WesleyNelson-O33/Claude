// CTS Business Intelligence Portal - people and access
//
// WHAT THIS IS AND IS NOT
//
// This decides which tabs a person is shown. It does not protect anything.
// The portal is a file that runs in the browser with no server behind it, so
// everything in portal/data is readable by anyone who can open the folder,
// whatever their role says. Treat this as deciding what people are shown, not
// what they can reach.
//
// If a figure genuinely must not be seen by someone, the answer is a separate
// build that does not contain it. portal/build/build_data.py can write a
// restricted data set; see the note on the Access Control page.
//
// Identification is by picking your name. There is no password and no single
// sign on, because CTS has neither wired to this portal.
window.CTS_USERS = {
  meta: {
    version: 2,
    updated: "2026-09-23",
    owner: "adminRole",
    note: "Visibility only. Not a security boundary.",
  },

  // Every page the portal can show. sensitive:true marks the ones that carry
  // cost, pay or transaction level detail, so whoever sets this up can see at
  // a glance what they are handing out.
  PAGES: [
    { id: "home",         section: "Dashboard",   title: "Dashboard" },
    { id: "context",      section: "Dashboard",   title: "Business Context" },
    { id: "pnl",          section: "Finance",     title: "P&L", sensitive: true },
    { id: "pnl-dept",     section: "Finance",     title: "P&L by Department", sensitive: true },
    { id: "pnl-spread",   section: "Finance",     title: "P&L Spread", sensitive: true },
    { id: "allocation",   section: "Finance",     title: "Overhead Allocation", sensitive: true },
    { id: "control",      section: "Finance",     title: "P&L Control", sensitive: true },
    { id: "bva",          section: "Budget",      title: "Budget vs Actual", sensitive: true },
    { id: "actions",      section: "Budget",      title: "Actions", sensitive: true },
    { id: "forecast",     section: "Budget",      title: "Forecast", sensitive: true },
    { id: "cash",         section: "Cash",        title: "Cash Flow", sensitive: true },
    { id: "rev-summary",  section: "Revenue",     title: "Revenue Summary" },
    { id: "rev-schedule", section: "Revenue",     title: "Revenue Schedule" },
    { id: "pipeline",     section: "Revenue",     title: "Pipeline" },
    { id: "clients",      section: "Revenue",     title: "Top Clients" },
    { id: "client-dept",  section: "Revenue",     title: "Clients by Department" },
    { id: "util",         section: "Utilisation", title: "Utilisation" },
    { id: "profit-fte",   section: "Utilisation", title: "Profitability per FTE", sensitive: true },
    { id: "staff-profit", section: "Utilisation", title: "Profitability by Employee", sensitive: true },
    { id: "ledger",       section: "Ledger",      title: "Detail Records", sensitive: true },
    { id: "charts",       section: "Ledger",      title: "Charts" },
    { id: "setup",        section: "Admin",       title: "Setup" },
    { id: "loaders",      section: "Admin",       title: "Data Loaders", sensitive: true },
    { id: "build",        section: "Admin",       title: "Build", sensitive: true },
    { id: "distribution", section: "Admin",       title: "Distribution", sensitive: true },
    { id: "config",       section: "Admin",       title: "Config & Variables" },
    { id: "access",       section: "Admin",       title: "Access Control", sensitive: true },
    { id: "about",        section: "Admin",       title: "About" },
  ],

  // Roles, in order of how much they see. `pages` is the allow list.
  // `ownDeptOnly` restricts every departmental figure to the person's own
  // department. `admin` is the only role that can change any of this.
  ROLES: [
    {
      id: "admin", label: "Finance head", admin: true, ownDeptOnly: false,
      note: "Everything, including who sees what. Intended for one person.",
      pages: ["home","context","pnl","pnl-dept","pnl-spread","allocation","control",
              "bva","actions","forecast","cash","rev-summary","rev-schedule","pipeline","clients","client-dept",
              "util","profit-fte","staff-profit","ledger","charts","setup","loaders","build",
              "distribution","config","access","about"],
    },
    {
      id: "finance", label: "Finance", admin: false, ownDeptOnly: false,
      note: "Everything except changing who sees what. The person who runs the month.",
      pages: ["home","context","pnl","pnl-dept","pnl-spread","allocation","control",
              "bva","actions","forecast","cash","rev-summary","rev-schedule","pipeline","clients","client-dept",
              "util","profit-fte","staff-profit","ledger","charts","setup","loaders","build",
              "distribution","config","about"],
    },
    {
      id: "exec", label: "Executive", admin: false, ownDeptOnly: false,
      note: "The whole company result, but not the ledger, the loaders or the workings.",
      pages: ["home","context","pnl","pnl-dept","pnl-spread","bva","actions","forecast","cash",
              "rev-summary","rev-schedule","pipeline","clients","client-dept","util",
              "profit-fte","staff-profit","charts","about"],
    },
    {
      id: "deptHead", label: "Department head", admin: false, ownDeptOnly: true,
      note: "Their own department in detail. No company P&L, no other department's numbers.",
      pages: ["home","context","pnl-dept","forecast","rev-schedule","pipeline","clients","util","charts","about"],
    },
    {
      id: "viewer", label: "Viewer", admin: false, ownDeptOnly: false,
      note: "The headline result and the context behind it, nothing else.",
      pages: ["home","context","about"],
    },
  ],

  // People. `dept` only matters for a role with ownDeptOnly.
  // Seeded from the names the payroll and reporting manuals actually use;
  // correct them on the Access Control page.
  USERS: [
    { name: "Danica Nelson",  role: "admin",    dept: null,
      note: "Set up the portal. Holds the finance head role until it is handed over." },
    { name: "Duncan",         role: "exec",     dept: null,
      note: "Approves the payment run and reads the monthly report." },
    { name: "Graham",         role: "exec",     dept: null,
      note: "Approves the pay run." },
    { name: "Jordan",         role: "deptHead", dept: "CONSULTING",
      note: "Confirms whether consulting revenue was lost or deferred." },
    { name: "Production manager", role: "deptHead", dept: "PRODUCTION",
      note: "Approves overtime and flags team travel." },
  ],

  // What someone sees if they have not identified themselves.
  DEFAULT_ROLE: "viewer",
};
