/* CTS Business Intelligence Portal - the Build module
 * ---------------------------------------------------------------------------
 * Turns the monthly Excel templates into the data files the portal runs on.
 *
 * Where it runs: in the browser, from the synced folder, with no server.
 * The page asks for the portal folder once (File System Access API, Edge or
 * Chrome), reads templates/*.xlsx with SheetJS, converts, checks, and writes
 * data/*.js, moving the previous set into data/_previous first. That is the
 * same mechanism O33's O33Dir library uses, and it is why nothing needs
 * installing.
 *
 * This file is the canonical converter. The paste boxes on Data Loaders are
 * the ad hoc path and share the parsing helpers in CTS.parse.
 */
(function () {
  "use strict";
  var CTS = (window.CTS = window.CTS || {});
  var B = (CTS.build = {});

  var TEMPLATES = [
    { key: "pnl",      file: "01 Xero P&L.xlsx",                  sheet: "P&L",      out: "CTS_fin_data.js",      need: true },
    { key: "gl",       file: "02 Xero GL Transactions.xlsx",      sheet: "GL",       out: "CTS_gl_data.js",       need: true },
    { key: "earnings", file: "03 Employment Hero Earnings.xlsx",  sheet: "Earnings", out: "CTS_util_data.js",     need: false },
    { key: "deals",    file: "04 Zoho Deals.xlsx",                sheet: "Deals",    out: "CTS_pipeline_data.js", need: false },
    { key: "orders",   file: "05 OnRent Orders.xlsx",             sheet: "Orders",   out: "CTS_pipeline_data.js", need: false, placeholder: true },
    { key: "quotes",   file: "06 Qwilr Quotes.xlsx",              sheet: "Quotes",   out: "CTS_pipeline_data.js", need: false, placeholder: true },
    { key: "budget",   file: "07 Budget.xlsx",                    sheet: "Budget",   out: "CTS_fin_data.js",      need: true },
    { key: "config",   file: "08 Config.xlsx",                    sheet: null,       out: "CTS_config_data.js",   need: true },
    { key: "forecast", file: "09 Forecast.xlsx",                  sheet: "Methods",  out: "CTS_forecast_data.js", need: false },
    { key: "cash",     file: "10 Cash and Commitments.xlsx",      sheet: "Bank balances", out: "CTS_cash_data.js", need: false },
    { key: "ar",       file: "11 Xero Aged Receivables.xlsx",     sheet: "Receivables", out: "CTS_aged_data.js", need: false },
    { key: "ap",       file: "12 Xero Aged Payables.xlsx",        sheet: "Payables", out: "CTS_aged_data.js",   need: false },
    { key: "schedule", file: "13 Revenue Schedule.xlsx",          sheet: "Schedule", out: "CTS_clients_data.js",  need: false },
  ];
  B.TEMPLATES = TEMPLATES;

  /* ================================================== reading a workbook */
  function P() { return CTS.parse; }
  function E() { return CTS.engine; }

  /** A sheet as {head, rows} with dates as ISO strings, the same shape the
   *  paste parser produces, so one converter serves both. Leading rows with
   *  fewer than two filled cells (a title such as "Earnings Details") are
   *  skipped, and a repeated header row inside the data (one pay run pasted
   *  under another) is dropped. */
  B.table = function (wb, sheetName, opts) {
    opts = opts || {};
    var ws = wb.Sheets[sheetName];
    if (!ws) return null;
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    function cell(v) {
      if (v instanceof Date) {
        // An Excel date is a day, not an instant. SheetJS hands it back as a
        // Date at midnight in one timezone or another, and read in the wrong
        // one a 1 November transaction lands in October. Push to midday and
        // read the UTC day, which is the same day whichever way it leaned.
        var t = new Date(v.getTime() + 12 * 3600 * 1000);
        var y = t.getUTCFullYear(), m = t.getUTCMonth() + 1, d = t.getUTCDate();
        return y + "-" + (m < 10 ? "0" : "") + m + "-" + (d < 10 ? "0" : "") + d;
      }
      if (typeof v === "string") return v.trim();
      return v;
    }
    var rows = aoa.map(function (r) { return r.map(cell); });
    if (opts.findHead) {
      // scan the first rows for the one the caller recognises as the header
      var at = -1;
      for (var i = 0; i < Math.min(rows.length, 40); i++) { if (opts.findHead(rows[i].map(function (x) { return String(x).trim(); }))) { at = i; break; } }
      if (at < 0) return null;
      rows = rows.slice(at);
    } else {
      while (rows.length && rows[0].filter(function (x) { return x !== "" && x != null; }).length < 2) rows.shift();
    }
    if (!rows.length) return { head: [], rows: [] };
    var head = rows[0].map(function (x) { return String(x).trim(); });
    var headKey = head.join("\u0001");
    var body = rows.slice(1).filter(function (r) {
      if (!r.some(function (x) { return x !== "" && x != null; })) return false;
      if (r.map(function (x) { return String(x).trim(); }).join("\u0001") === headKey) return false;
      var filled = r.filter(function (x) { return x !== "" && x != null; }).length;
      return filled >= (opts.minFilled || 2);
    });
    return { head: head, rows: body };
  };

  function colIndex(head, names) {
    var normed = head.map(function (h) { return String(h).toLowerCase().replace(/[^a-z0-9]/g, ""); });
    for (var i = 0; i < names.length; i++) {
      var k = normed.indexOf(names[i]);
      if (k >= 0) return k;
    }
    return -1;
  }
  B.colIndex = colIndex;

  var MON = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
  /** A month heading to its key, without looking at the calendar, so a
   *  column for a year the calendar has not seen yet is still read. Accepts
   *  2027-07, Jul-27, Jul 2027, July 2027, 07/2027 and an Excel date. */
  function monthOfLabel(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date && !isNaN(v)) {
      var dd = new Date(v.getTime() + 12 * 3600 * 1000);
      return dd.getUTCFullYear() + "-" + (dd.getUTCMonth() < 9 ? "0" : "") + (dd.getUTCMonth() + 1);
    }
    var s = String(v).trim().toLowerCase();
    var m = s.match(/^(\d{4})-(\d{2})/);
    if (m) return m[1] + "-" + m[2];
    m = s.match(/^([a-z]{3})[a-z]*[\s\-\/']*(\d{2}|\d{4})$/);
    if (m && MON.indexOf(m[1]) >= 0) {
      var y = m[2].length === 2 ? 2000 + +m[2] : +m[2];
      return y + "-" + (MON.indexOf(m[1]) < 9 ? "0" : "") + (MON.indexOf(m[1]) + 1);
    }
    m = s.match(/^(\d{1,2})[\/\-](\d{4})$/);
    if (m && +m[1] >= 1 && +m[1] <= 12) return m[2] + "-" + (+m[1] < 10 ? "0" : "") + +m[1];
    return null;
  }
  B.monthOfLabel = monthOfLabel;

  /** The calendar the build is working against. Set while a build runs so
   *  the converters see the months of this build, not of the last one. */
  var ACTIVE = { monthIdx: null };
  function monthIdx() { return ACTIVE.monthIdx || E().monthIdx; }

  /** Build the financial calendar for a span of month keys, padded out to
   *  whole financial years, with working days net of the holidays given.
   *  This is what lets the portal roll into a new financial year: the span
   *  comes from the templates, so a Jul-27 column makes FY28 exist. */
  B.calendar = function (keys, holidays, states, meta) {
    var STATES = states || ["VIC", "NSW", "QLD", "WA", "SA", "TAS", "ACT", "NT"];
    var NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    keys = keys.filter(Boolean).sort();
    function fyOf(y, mo) { return (mo >= 7 ? y + 1 : y) % 100; }
    var first = keys[0].split("-").map(Number), last = keys[keys.length - 1].split("-").map(Number);
    var fyA = fyOf(first[0], first[1]), fyZ = fyOf(last[0], last[1]);
    var index = {};
    (holidays || []).forEach(function (h) { (h.states || STATES).forEach(function (s) { index[h.date + "|" + s] = 1; }); });
    var months = [];
    for (var fy = fyA; fy <= fyZ; fy++) {
      for (var i = 0; i < 12; i++) {
        var mo = ((6 + i) % 12) + 1, y = 2000 + fy - (mo >= 7 ? 1 : 0);
        var bd = 0, wd = {}, days = new Date(y, mo, 0).getDate();
        STATES.forEach(function (s) { wd[s] = 0; });
        for (var d = 1; d <= days; d++) {
          var dow = new Date(y, mo - 1, d).getDay();
          if (dow === 0 || dow === 6) continue;
          bd++;
          var iso = y + "-" + (mo < 10 ? "0" : "") + mo + "-" + (d < 10 ? "0" : "") + d;
          STATES.forEach(function (s) { if (!index[iso + "|" + s]) wd[s]++; });
        }
        months.push({ key: y + "-" + (mo < 10 ? "0" : "") + mo, y: y, m: mo, fy: fy,
                      label: NAMES[mo - 1] + "-" + String(y).slice(2), long: NAMES[mo - 1] + " " + y,
                      period: i + 1, quarter: Math.floor(i / 3) + 1, wd: wd, bd: bd, days: days });
      }
    }
    return { fyStartMonth: 7, states: STATES, months: months, holidays: holidays || [],
             quarters: [{ n: 1, label: "Q1 Jul-Sep" }, { n: 2, label: "Q2 Oct-Dec" }, { n: 3, label: "Q3 Jan-Mar" }, { n: 4, label: "Q4 Apr-Jun" }],
             meta: Object.assign({ note: "Financial year 1 July to 30 June, named for the year it ends." }, meta || {}) };
  };

  /* =========================================================== converters */
  var C = (B.convert = {});

  /** Account rows down, months across. Section headings and totals are
   *  skipped by name; anything else not in the chart is reported. */
  function accountGrid(t, label) {
    var months = t.head.slice(1).map(monthOfLabel);
    var out = {}, unknown = [], skippedTotals = 0, matched = 0;
    var SKIP = /^(total |gross profit|net profit|income$|cost of sales$|other income$|expenses$|other expenses$|less |plus )/i;
    t.rows.forEach(function (r) {
      var name = String(r[0] || "").trim();
      if (!name) return;
      if (SKIP.test(name)) { skippedTotals++; return; }
      if (!E().acct[name]) { unknown.push(name); return; }
      matched++;
      var a = E().acct[name];
      var sign = (a.cat === "cos" || a.cat === "expenses" || a.cat === "other_expenses") ? -1 : 1;
      months.forEach(function (mk, i) {
        if (!mk) return;
        var v = P().num(r[i + 1]);
        if (v) (out[name] = out[name] || {})[mk] = Math.round(v * 100) * sign;
      });
    });
    var used = {};
    Object.keys(out).forEach(function (a) { Object.keys(out[a]).forEach(function (mk) { used[mk] = 1; }); });
    return { grid: out, months: Object.keys(used).sort(), unknown: unknown,
             skippedTotals: skippedTotals, matched: matched, label: label };
  }
  C.pnl = function (t) { return accountGrid(t, "P&L"); };
  C.budget = function (t) { return accountGrid(t, "Budget"); };

  /** The Departments tab of 07 Budget: one row per department group and
   *  category, months across, the way the budget workbook sets income and
   *  direct costs per group and the overheads for Admin. Costs are stored
   *  negative, as the account grid stores them. */
  C.budgetDept = function (t) {
    var months = t.head.slice(3).map(monthOfLabel), rows = [], bad = [];
    var CAT = { income: "income", "direct costs": "cos", "cost of sales": "cos", cos: "cos", overheads: "expenses", expenses: "expenses", "other income": "other_income", "other expenses": "other_expenses" };
    t.rows.forEach(function (r) {
      var grp = String(r[0] || "").trim(); if (!grp) return;
      var depts = String(r[1] || "").split(/[,;]/).map(function (x) { return x.trim().toUpperCase(); }).filter(Boolean);
      var cat = CAT[String(r[2] || "").trim().toLowerCase()];
      if (!depts.length || !cat) { bad.push(grp); return; }
      var sign = (cat === "cos" || cat === "expenses" || cat === "other_expenses") ? -1 : 1, grid = {};
      months.forEach(function (mk, i) { if (!mk) return; var v = P().num(r[i + 3]); if (v) grid[mk] = Math.round(v * 100) * sign; });
      rows.push({ group: grp, depts: depts, cat: cat, months: grid });
    });
    return { rows: rows, bad: bad };
  };

  /** 13 Revenue Schedule: the Schedule tab (one line per role or client,
   *  budget months across) and the Clients tab (display name, industry,
   *  engagement, primary department, aliases). */
  C.schedule = function (t, clientRows) {
    var months = t.head.slice(5).map(monthOfLabel), lines = [], bad = 0;
    var KIND = { client: "client", role: "role", "new business": "newbiz", newbiz: "newbiz", new: "newbiz" };
    t.rows.forEach(function (r) {
      var dept = String(r[0] || "").trim().toUpperCase(); if (!dept) return;
      var kind = KIND[String(r[1] || "").trim().toLowerCase()] || "client";
      var client = String(r[2] || "").trim(), line = String(r[3] || "").trim() || client;
      if (kind !== "newbiz" && !client) { bad++; return; }
      var grid = {};
      months.forEach(function (mk, i) { if (!mk) return; var v = P().num(r[i + 5]); if (v) grid[mk] = Math.round(v * 100); });
      var l = { dept: dept, kind: kind, client: client, line: line };
      if (String(r[4] || "").trim()) l.industry = String(r[4]).trim();
      if (Object.keys(grid).length) l.budget = grid;
      lines.push(l);
    });
    var clients = [], seen = {};
    (clientRows || []).forEach(function (r) {
      var contact = String(r[0] || "").trim(); if (!contact || seen[contact]) return; seen[contact] = 1;
      var c = { contact: contact, display: String(r[1] || "").trim() || contact, industry: String(r[2] || "").trim(),
                engagement: String(r[3] || "").trim(), primaryDept: String(r[4] || "").trim().toUpperCase() || null };
      var al = String(r[5] || "").split(";").map(function (x) { return x.trim(); }).filter(Boolean);
      if (al.length) c.aliases = al;
      clients.push(c);
    });
    // every scheduled client is a client, even if the Clients tab misses it
    lines.forEach(function (l) {
      if (!l.client || seen[l.client]) return; seen[l.client] = 1;
      clients.push({ contact: l.client, display: l.client, industry: l.industry || "", engagement: "", primaryDept: l.dept });
    });
    return { payload: { meta: { seed: false, built: today(), note: "Built from 13 Revenue Schedule.xlsx. Budget per line in cents by month; actual is read off the ledger by contact." },
                        clients: clients, schedule: lines },
             bad: bad, typed: lines.filter(function (l) { return l.budget; }).length };
  };

  /** The thirteen column Xero transactions report. */
  C.gl = function (t) {
    var h = t.head;
    var ci = {
      code: colIndex(h, ["accountcode"]), name: colIndex(h, ["accountname", "account"]),
      source: colIndex(h, ["source"]), date: colIndex(h, ["date"]), contact: colIndex(h, ["contact"]),
      debit: colIndex(h, ["debit"]), credit: colIndex(h, ["credit"]),
      job: colIndex(h, ["jobnumbers", "jobnumber"]), inv: colIndex(h, ["invoicenumber"]),
      ref: colIndex(h, ["reference"]), desc: colIndex(h, ["description"]),
      cc: colIndex(h, ["costcentres", "costcentre"]),
    };
    if (ci.name < 0 || ci.date < 0) throw new Error("GL: could not find Account Name and Date columns.");
    var rows = [], bad = 0, unknown = {}, autoCoded = 0, noContact = 0;
    t.rows.forEach(function (r) {
      var iso = P().isoDate(r[ci.date]);
      var name = String(r[ci.name] || "").trim();
      if (!iso || !name) { bad++; return; }
      if (!E().acct[name]) unknown[name] = (unknown[name] || 0) + 1;
      var cc = ci.cc >= 0 ? String(r[ci.cc] || "") : "";
      var contact = ci.contact >= 0 ? String(r[ci.contact] || "").trim() : "";
      // Stripe fee lines come across with no cost centre; the analysis file
      // codes them to Admin by hand every month, so the build does it here
      if (!cc.trim() && !(ci.job >= 0 && /\[/.test(String(r[ci.job] || ""))) && /stripe/i.test(contact + " " + (ci.desc >= 0 ? String(r[ci.desc] || "") : ""))) {
        cc = "9000 - OFFICE / ADMIN [CTS]"; autoCoded++;
      }
      rows.push([iso, name, contact,
                 P().num(r[ci.debit]), P().num(r[ci.credit]),
                 ci.source >= 0 ? String(r[ci.source] || "") : "",
                 ci.job >= 0 ? String(r[ci.job] || "") : "",
                 ci.inv >= 0 ? String(r[ci.inv] || "") : "",
                 cc,
                 ci.desc >= 0 ? String(r[ci.desc] || "") : ""]);
      var a = E().acct[name];
      if (a && a.cat === "income" && !contact && (P().num(r[ci.credit]) || P().num(r[ci.debit]))) noContact++;
    });
    return { payload: { meta: { seed: false, built: today(), rows: rows.length, note: "Built from 02 Xero GL Transactions.xlsx" },
                        cols: ["date", "account", "contact", "debit", "credit", "source", "jobNo", "invoiceNo", "costCentre", "description"],
                        rows: rows },
             bad: bad, unknown: unknown, autoCoded: autoCoded, noContact: noContact };
  };

  /** Xero's standard General Ledger Detail report, as a fallback for the
   *  custom Account transactions report. It is grouped by account, with the
   *  account name on its own row above its lines and a Total row below, and
   *  it carries no cost centre, job number or contact column: the contact is
   *  what the Description holds on invoice and bill lines. Departments cannot
   *  be coded from it, which the build says plainly. */
  C.glDetail = function (t) {
    var h = t.head;
    var ci = { date: colIndex(h, ["date"]), source: colIndex(h, ["source"]), desc: colIndex(h, ["description"]),
               ref: colIndex(h, ["reference"]), debit: colIndex(h, ["debit"]), credit: colIndex(h, ["credit"]) };
    if (ci.date < 0 || ci.debit < 0 || ci.credit < 0) throw new Error("GL Detail: could not find Date, Debit and Credit columns.");
    var rows = [], bad = 0, account = null, skipped = {}, inChart = 0, noContact = 0;
    var NOT_CONTACT = /^(manual journal|bank transfer|adjustment|spend money|receive money)$/i;
    t.rows.forEach(function (r) {
      var first = String(r[0] || "").trim();
      var iso = P().isoDate(r[ci.date]);
      if (!iso) {
        // a heading row names the account; a Total row closes it
        if (first && !/^total /i.test(first) && !/^(opening|closing) balance/i.test(first)) account = first;
        return;
      }
      if (!account) { bad++; return; }
      if (!E().acct[account]) { skipped[account] = (skipped[account] || 0) + 1; return; }
      inChart++;
      var src = ci.source >= 0 ? String(r[ci.source] || "") : "";
      var desc = ci.desc >= 0 ? String(r[ci.desc] || "").trim() : "";
      var contact = NOT_CONTACT.test(src) ? "" : desc;
      var a = E().acct[account];
      if (a.cat === "income" && !contact) noContact++;
      rows.push([iso, account, contact, P().num(r[ci.debit]), P().num(r[ci.credit]), src, "", "",
                 "", ci.ref >= 0 ? String(r[ci.ref] || "") : ""]);
    });
    return { payload: { meta: { seed: false, built: today(), rows: rows.length, layout: "General Ledger Detail",
                                note: "Built from Xero's General Ledger Detail report. No cost centres or job numbers: departments are not coded." },
                        cols: ["date", "account", "contact", "debit", "credit", "source", "jobNo", "invoiceNo", "costCentre", "description"],
                        rows: rows },
             bad: bad, unknown: {}, autoCoded: 0, noContact: noContact, skipped: skipped, detail: true };
  };

  /** The Xero Aged Receivables and Aged Payables Detail reports: one row per
   *  open invoice or bill with its ageing buckets. Contact grouping and total
   *  rows that come with a paste are skipped on having no date and no total. */
  C.aged = function (t) {
    var h = t.head, p = P();
    var ci = { contact: colIndex(h, ["contact", "contactname", "name"]), date: colIndex(h, ["date", "invoicedate", "billdate"]),
               due: colIndex(h, ["duedate", "due"]), no: colIndex(h, ["number", "invoicenumber", "reference"]), ref: colIndex(h, ["reference", "description"]),
               cur: colIndex(h, ["current"]), b1: colIndex(h, ["1month", "lessthan1month", "1monthold"]), b2: colIndex(h, ["1month", "1monthold"]),
               b3: colIndex(h, ["2months", "2monthsold"]), b4: colIndex(h, ["3months", "3monthsold"]), older: colIndex(h, ["older", "olderthan3months"]), total: colIndex(h, ["total"]) };
    // "< 1 Month" and "1 Month" both normalise to 1month: take them by position
    var lt = -1, one = -1;
    h.forEach(function (x, i) { var k = String(x).toLowerCase().replace(/\s+/g, ""); if (k === "<1month") lt = i; else if (k === "1month") one = i; });
    if (lt >= 0) ci.b1 = lt; if (one >= 0) ci.b2 = one;
    if (ci.contact < 0 || ci.total < 0) throw new Error("Aged: need Contact and Total columns.");
    var rows = [], skipped = 0, carried = "";
    t.rows.forEach(function (r) {
      var contact = String(r[ci.contact] || "").trim();
      var iso = ci.date >= 0 ? p.isoDate(r[ci.date]) : null, total = p.money(r[ci.total]);
      if (!iso && !total) { if (contact && !/^total/i.test(contact)) carried = contact; skipped++; return; }
      if (!contact) contact = carried;
      if (/^total/i.test(contact)) { skipped++; return; }
      function c(i) { return i >= 0 ? Math.round(p.money(r[i]) * 100) : 0; }
      rows.push({ contact: contact, date: iso, due: ci.due >= 0 ? p.isoDate(r[ci.due]) : null,
                  number: ci.no >= 0 ? String(r[ci.no] || "") : "", ref: ci.ref >= 0 ? String(r[ci.ref] || "") : "",
                  current: c(ci.cur), lt1: c(ci.b1), m1: c(ci.b2), m2: c(ci.b3), m3: c(ci.b4), older: c(ci.older), total: Math.round(total * 100) });
    });
    return { rows: rows, skipped: skipped };
  };

  /** The Earnings Details report, with the Locations, Pay categories and
   *  Adjustments tabs from the same workbook. Same rules as the paste loader. */
  /** A location not on the Locations tab: the department from the bracket
   *  tag, or failing that a department tag anywhere in the name ("ONS DTTL
   *  (non-chargeable)", "PRODUCTION"), and chargeable unless the name says
   *  otherwise or the department is Admin. */
  function guessLocation(name) {
    var g = E().resolveDept("", name);
    if (g === "UNALLOCATED") {
      var up = String(name || "").toUpperCase();
      var best = null;
      (E().cfg.DEPARTMENTS || E().depts || []).forEach(function (d) {
        (d.tags || []).forEach(function (tag) {
          var re = new RegExp("(^|[^A-Z])" + tag + "([^A-Z]|$)");
          if (re.test(up) && (!best || tag.length > best.tag.length)) best = { code: d.code, tag: tag };
        });
      });
      if (best) g = best.code;
    }
    return { dept: g, chargeable: g !== "ADMIN" && g !== "UNALLOCATED" && !/non[- ]?ch/i.test(String(name || "")), guessed: true };
  }
  C.guessLocation = guessLocation;
  C.earnings = function (t, locs, cats, adjs, fallbackMonth) {
    var h = t.head, p = P();
    var ci = {
      empId: colIndex(h, ["employeeid"]), empExt: colIndex(h, ["employeeexternalid"]),
      empName: colIndex(h, ["employeename"]), catName: colIndex(h, ["paycategoryname"]),
      units: colIndex(h, ["units"]), unitType: colIndex(h, ["unittype"]),
      locName: colIndex(h, ["locationname"]), notes: colIndex(h, ["notes"]),
      gross: colIndex(h, ["grossearnings"]), sup: colIndex(h, ["sgsuper"]),
    };
    if (ci.catName < 0 || ci.locName < 0 || ci.units < 0) throw new Error("Earnings: need Pay Category Name, Location Name and Units.");

    var locMap = {}, catMap = {};
    (locs || []).forEach(function (r) {
      if (!r[0]) return;
      locMap[String(r[0]).trim()] = { dept: String(r[1] || "").trim().toUpperCase() || null,
                                      chargeable: /^charge/i.test(String(r[2] || "")) };
    });
    (cats || []).forEach(function (r) {
      if (!r[0]) return;
      var k = String(r[1] || "").toLowerCase();
      catMap[String(r[0]).trim()] = /leave/.test(k) ? "leave" : /public/.test(k) ? "publicHoliday"
                                   : /pay only|exclude/.test(k) ? "exclude" : "work";
    });
    var adjList = (adjs || []).filter(function (r) { return r[0] || r[1]; }).map(function (r) {
      var kind = String(r[5] || "").toLowerCase();
      return { who: String(r[0] || "").trim().toLowerCase(), cat: String(r[1] || "").trim().toLowerCase(),
               note: String(r[2] || "").trim().toLowerCase(), hours: p.num(r[3]),
               loc: String(r[4] || "").trim(),
               kind: /leave/.test(kind) ? "leave" : /public/.test(kind) ? "publicHoliday" : "work" };
    });

    var bucket = {}, staff = {}, labour = {}, newLocs = {}, newCats = {}, usedFallback = 0;
    var reallocated = 0, payOnly = 0, lines = 0, spread = 0;
    // Pay runs pasted one under the other: each starts with its title row.
    // A line whose note carries no date (salaried staff, allowances) is
    // spread over the months its pay run covers, on weekdays, rather than
    // dropped whole into the reporting month.
    var blocks = [], cur = null;
    function isTitle(r) { return /^earnings details$/i.test(String(r[0] || "").trim()) && r.slice(1).every(function (x) { return x === "" || x == null; }); }
    t.rows.forEach(function (r) {
      if (isTitle(r) || !cur) { cur = { rows: [], dates: {}, min: null, max: null }; blocks.push(cur); if (isTitle(r)) return; }
      cur.rows.push(r);
      var iso = ci.notes >= 0 ? p.noteDate(r[ci.notes]) : null;
      if (iso) cur.dates[iso] = (cur.dates[iso] || 0) + 1;
    });
    // the pay period is the fortnight holding most of the block's dated lines;
    // a back-dated correction or leave booked ahead does not stretch it
    blocks.forEach(function (bk) {
      var ds = Object.keys(bk.dates).sort(); if (!ds.length) return;
      var best = null;
      ds.forEach(function (end) {
        var start = new Date(new Date(end + "T12:00:00Z").getTime() - 13 * 86400000).toISOString().slice(0, 10), n = 0;
        ds.forEach(function (d) { if (d >= start && d <= end) n += bk.dates[d]; });
        if (!best || n > best.n) best = { n: n, start: start, end: end };
      });
      bk.min = best.start; bk.max = best.end;
    });
    function weekdayShare(from, to) {
      var out = {}, total = 0, d = new Date(from + "T12:00:00Z"), end = new Date(to + "T12:00:00Z");
      while (d <= end) {
        var dow = d.getUTCDay();
        if (dow !== 0 && dow !== 6) { var mk = d.toISOString().slice(0, 7); out[mk] = (out[mk] || 0) + 1; total++; }
        d = new Date(d.getTime() + 86400000);
      }
      Object.keys(out).forEach(function (mk) { out[mk] = out[mk] / total; });
      return out;
    }
    var share = null;
    blocks.forEach(function (bk) {
      share = null;
      if (bk.min && bk.max) share = weekdayShare(bk.min, bk.max);
      bk.rows.forEach(function (r) {
        var iso = ci.notes >= 0 ? p.noteDate(r[ci.notes]) : null;
        if (iso) addLine(r, iso.slice(0, 7), 1);
        else if (share) { usedFallback++; spread++; Object.keys(share).forEach(function (mk) { addLine(r, mk, share[mk]); }); }
        else { usedFallback++; addLine(r, fallbackMonth, 1); }
      });
    });
    function addLine(r, mk, frac) {
      var locName = String(r[ci.locName] || "").trim() || "(blank)";
      var loc = locMap[locName];
      if (!loc) {
        loc = guessLocation(locName);
        newLocs[locName] = loc;
      }
      if (!loc.dept) loc.dept = "UNALLOCATED";
      var catName = String(r[ci.catName] || "").trim() || "(blank)";
      var kind = catMap[catName];
      if (!kind) { kind = p.guessCategory(catName); newCats[catName] = kind; }
      if (!monthIdx()[mk]) return;
      var isHours = ci.unitType < 0 || /hour/i.test(String(r[ci.unitType]));
      var units = (isHours ? p.num(r[ci.units]) : 0) * frac;
      var gross = (ci.gross >= 0 ? p.money(r[ci.gross]) : 0) * frac;
      var sup = (ci.sup >= 0 ? p.money(r[ci.sup]) : 0) * frac;
      var who = ci.empName >= 0 ? String(r[ci.empName] || "").trim() : "";
      var note = ci.notes >= 0 ? String(r[ci.notes] || "").toLowerCase() : "";

      if (kind === "exclude" || units === 0) {
        payOnly++;
        var a = adjList.filter(function (x) {
          return (!x.who || who.toLowerCase() === x.who) && (!x.cat || catName.toLowerCase() === x.cat) &&
                 (!x.note || note.indexOf(x.note) >= 0);
        })[0];
        if (a && a.hours > 0) {
          units = a.hours * frac; kind = a.kind; reallocated++;
          if (a.loc && locMap[a.loc]) { locName = a.loc; loc = locMap[a.loc]; }
        }
      }
      if (frac === 1 || mk === Object.keys(share || {})[0]) lines++;
      var key = mk + "|" + loc.dept;
      var b = bucket[key] || (bucket[key] = { month: mk, dept: loc.dept, chargeable: 0, nonChargeable: 0, leave: 0, publicHoliday: 0 });
      if (kind === "work") b[loc.chargeable ? "chargeable" : "nonChargeable"] += units;
      else if (kind === "leave") b.leave += units;
      else if (kind === "publicHoliday") b.publicHoliday += units;
      var lb = labour[key] || (labour[key] = { gross: 0, sup: 0 });
      lb.gross += gross; lb.sup += sup;

      var pid = (ci.empExt >= 0 && String(r[ci.empExt]).trim()) || (ci.empId >= 0 && String(r[ci.empId]).trim()) || who;
      var sk = mk + "|" + pid + "|" + loc.dept;
      var s = staff[sk] || (staff[sk] = { month: mk, empId: pid, name: who || pid, dept: loc.dept,
                chargeable: 0, nonChargeable: 0, leave: 0, publicHoliday: 0, gross: 0, sup: 0, employment: null });
      if (kind === "work") s[loc.chargeable ? "chargeable" : "nonChargeable"] += units;
      else if (kind === "leave") s.leave += units;
      else if (kind === "publicHoliday") s.publicHoliday += units;
      s.gross += gross; s.sup += sup;
      var emp = p.guessEmployment(catName);
      if (emp) s.employment = emp;
    }
    function r2(v) { return Math.round(v * 100) / 100; }
    var utilRows = Object.keys(bucket).sort().map(function (k) {
      var b = bucket[k];
      return [b.month, b.dept, r2(b.chargeable), r2(b.nonChargeable), r2(b.leave), r2(b.publicHoliday), 0];
    });
    var staffRows = Object.keys(staff).sort().map(function (k) {
      var s = staff[k];
      return [s.month, s.empId, s.name, s.dept, r2(s.chargeable), r2(s.nonChargeable), r2(s.leave),
              r2(s.publicHoliday), Math.round(s.gross * 100), Math.round(s.sup * 100), s.employment];
    });
    return {
      util: { meta: { seed: false, built: today(), source: "03 Employment Hero Earnings.xlsx" },
              cols: ["month", "dept", "chargeable", "nonChargeable", "leave", "publicHoliday", "fte"], rows: utilRows },
      staff: { meta: { seed: false, built: today(), source: "03 Employment Hero Earnings.xlsx" },
               cols: ["month", "empId", "name", "dept", "chargeable", "nonChargeable", "leave", "publicHoliday", "grossCents", "superCents", "employment"], rows: staffRows },
      lines: lines, newLocations: newLocs, newCategories: newCats, usedFallback: usedFallback, spread: spread,
      blocks: blocks.map(function (b) { return { from: b.min, to: b.max, lines: b.rows.length }; }),
      reallocated: reallocated, payOnly: payOnly,
    };
  };

  /** Pipeline: deals, orders and quotes. Placeholder shapes until real
   *  exports arrive; every column is matched by name so a reorder is fine. */
  function deptOf(v) {
    var s = String(v || "").trim();
    if (!s) return "UNALLOCATED";
    var d = E().resolveDept(s, s);
    return d;
  }
  /** Zoho's "All Deals by Stage" report, pasted whole. The stage is a group
   *  heading with a count after it ("Closed Won ( 483 )") on a row of its
   *  own, and the deals sit under it with the stage cell blank, so the stage
   *  is carried down. CTS DEPARTMENT is Zoho's own field, mapped to the
   *  portal's departments on the template's Departments tab. Opportunity
   *  Number is the job number, which is what OnRent, Qwilr and the Xero job
   *  tracking category also carry, and what stops a deal being counted twice. */
  var ZOHO_DEPT = { "event production": "PRODUCTION", "production": "PRODUCTION", "support": "ONSITE", "onsite": "ONSITE",
                    "video production": "VIDEO", "video": "VIDEO", "integration": "INTEGRATION", "consulting": "CONSULTING",
                    "multi-service": "UNALLOCATED", "multi service": "UNALLOCATED", "admin": "ADMIN", "cts": "ADMIN" };
  C.deals = function (t, deptRows) {
    var h = t.head, p = P();
    var ci = { stage: colIndex(h, ["stage"]), name: colIndex(h, ["dealname"]), account: colIndex(h, ["accountname"]),
               amount: colIndex(h, ["amountrecordcurrency", "amount"]), dept: colIndex(h, ["ctsdepartment", "department"]),
               close: colIndex(h, ["closingdate", "closedate"]), job: colIndex(h, ["opportunitynumber", "jobnumber"]),
               owner: colIndex(h, ["dealowner", "owner"]), prob: colIndex(h, ["probability"]), expected: colIndex(h, ["expectedrevenue"]),
               type: colIndex(h, ["type"]), source: colIndex(h, ["leadsource"]), created: colIndex(h, ["createdtime", "created"]) };
    if (ci.name < 0 || ci.stage < 0) throw new Error("Deals: need Stage and Deal Name.");
    var map = Object.assign({}, ZOHO_DEPT);
    (deptRows || []).forEach(function (r) { if (r[0]) map[String(r[0]).trim().toLowerCase()] = String(r[1] || "").trim().toUpperCase(); });
    var codes = {}; (E().depts || []).forEach(function (d) { codes[d.code] = 1; });
    var stage = "", unmapped = {}, out = [], groups = 0;
    t.rows.forEach(function (r) {
      var st = String(r[ci.stage] || "").trim(), name = String(r[ci.name] || "").trim();
      if (st) { stage = st.replace(/\s*\(\s*\d+\s*\)\s*$/, "").trim(); groups++; }
      if (!name) return;
      var amt = p.money(r[ci.amount]), prob = ci.prob >= 0 ? p.num(r[ci.prob]) : null;
      var zd = ci.dept >= 0 ? String(r[ci.dept] || "").trim() : "";
      var code = map[zd.toLowerCase()];
      if (!code || !codes[code]) { if (zd) unmapped[zd] = (unmapped[zd] || 0) + 1; code = (zd && map[zd.toLowerCase()] === undefined) ? deptOf(zd) : "UNALLOCATED"; }
      out.push({ name: name, account: ci.account >= 0 ? String(r[ci.account] || "").trim() : "",
                 stage: stage || st, amount: Math.round(amt * 100),
                 close: ci.close >= 0 ? p.isoDate(r[ci.close]) : null,
                 job: ci.job >= 0 ? String(r[ci.job] || "").trim() : "",
                 owner: ci.owner >= 0 ? String(r[ci.owner] || "") : "",
                 prob: prob == null ? null : (prob > 1 ? prob / 100 : prob),
                 expected: ci.expected >= 0 ? Math.round(p.money(r[ci.expected]) * 100) : (prob != null ? Math.round(amt * (prob > 1 ? prob / 100 : prob) * 100) : null),
                 type: ci.type >= 0 ? String(r[ci.type] || "") : "", source: ci.source >= 0 ? String(r[ci.source] || "") : "",
                 created: ci.created >= 0 ? p.isoDate(r[ci.created]) : null, zohoDept: zd, dept: code });
    });
    out.unmapped = unmapped; out.groups = groups;
    return out;
  };
  C.orders = function (t) {
    var h = t.head, p = P();
    var ci = { no: colIndex(h, ["orderno", "ordernumber", "jobno"]), client: colIndex(h, ["client", "customer"]),
               title: colIndex(h, ["jobtitle", "title"]), start: colIndex(h, ["eventstart", "startdate", "start"]),
               end: colIndex(h, ["eventend", "enddate", "end"]), dept: colIndex(h, ["department"]),
               status: colIndex(h, ["status"]), value: colIndex(h, ["valueexgst", "value", "total"]),
               owner: colIndex(h, ["owner"]), created: colIndex(h, ["created"]) };
    if (ci.client < 0 || ci.start < 0) throw new Error("Orders: need Client and Event Start.");
    return t.rows.map(function (r) {
      return { no: ci.no >= 0 ? String(r[ci.no] || "") : "", client: String(r[ci.client] || ""),
               title: ci.title >= 0 ? String(r[ci.title] || "") : "",
               start: p.isoDate(r[ci.start]), end: ci.end >= 0 ? p.isoDate(r[ci.end]) : null,
               dept: ci.dept >= 0 ? deptOf(r[ci.dept]) : "PRODUCTION",
               status: ci.status >= 0 ? String(r[ci.status] || "") : "",
               value: Math.round(p.money(r[ci.value]) * 100),
               owner: ci.owner >= 0 ? String(r[ci.owner] || "") : "",
               created: ci.created >= 0 ? p.isoDate(r[ci.created]) : null };
    }).filter(function (o) { return o.client && o.start; });
  };
  C.quotes = function (t) {
    var h = t.head, p = P();
    var ci = { ref: colIndex(h, ["quoteref", "reference", "quoteno"]), client: colIndex(h, ["client", "customer"]),
               title: colIndex(h, ["title"]), sent: colIndex(h, ["sent", "sentdate", "created"]),
               status: colIndex(h, ["status"]), accepted: colIndex(h, ["accepted", "accepteddate"]),
               value: colIndex(h, ["valueexgst", "value", "total"]), owner: colIndex(h, ["owner"]),
               dept: colIndex(h, ["department"]) };
    if (ci.client < 0 || ci.status < 0) throw new Error("Quotes: need Client and Status.");
    return t.rows.map(function (r) {
      return { ref: ci.ref >= 0 ? String(r[ci.ref] || "") : "", client: String(r[ci.client] || ""),
               title: ci.title >= 0 ? String(r[ci.title] || "") : "",
               sent: ci.sent >= 0 ? p.isoDate(r[ci.sent]) : null, status: String(r[ci.status] || ""),
               accepted: ci.accepted >= 0 ? p.isoDate(r[ci.accepted]) : null,
               value: Math.round(p.money(r[ci.value]) * 100),
               owner: ci.owner >= 0 ? String(r[ci.owner] || "") : "",
               dept: ci.dept >= 0 ? deptOf(r[ci.dept]) : "UNALLOCATED" };
    }).filter(function (q) { return q.client; });
  };

  /** The Config workbook becomes both CTS_config_data.js and CTS_users_data.js. */
  C.config = function (wb) {
    var base = window.CTS_CONFIG || {};
    var cfg = JSON.parse(JSON.stringify(base));
    var p = P();
    function tab(name) { var t = B.table(wb, name); return t ? t.rows : []; }
    function bool(v) { return /^y/i.test(String(v || "")); }

    // Reporting
    var rep = {};
    tab("Reporting").forEach(function (r) { rep[String(r[0] || "").toLowerCase()] = r[1]; });
    var rm = p.isoDate(rep["reporting month"]);
    cfg.ORG = cfg.ORG || {};
    if (rm) cfg.ORG.reportingMonth = rm.slice(0, 7);
    if (rep["portal link"] != null) cfg.ORG.portalUrl = String(rep["portal link"] || "").trim();
    cfg.UTIL = cfg.UTIL || {};
    if (rep["hours in a working day"]) cfg.UTIL.hoursPerDay = p.num(rep["hours in a working day"]);
    if (rep["state for working days"]) cfg.UTIL.state = String(rep["state for working days"]).toUpperCase();
    cfg.RISK = cfg.RISK || {};
    if (rep["materiality floor ($)"] != null) cfg.RISK.materialityFloor = p.num(rep["materiality floor ($)"]);
    if (rep["medium risk (% variance)"] != null) cfg.RISK.mediumPct = p.num(rep["medium risk (% variance)"]);
    if (rep["high risk (% variance)"] != null) cfg.RISK.highPct = p.num(rep["high risk (% variance)"]);
    if (rep["high risk ($ variance)"] != null) cfg.RISK.highDollar = p.num(rep["high risk ($ variance)"]);
    if (rep["portal version"]) cfg.VERSION = String(rep["portal version"]);
    cfg.RISK.source = "08 Config.xlsx";

    // Departments
    var depts = tab("Departments").filter(function (r) { return r[0]; }).map(function (r) {
      var util = r[5] === "" || r[5] == null ? null : p.num(r[5]);
      return { code: String(r[0]).toUpperCase(), short: String(r[1] || r[0]), long: String(r[2] || ""),
               order: p.num(r[3]) || 99, isRevenue: bool(r[4]),
               utilTarget: util == null ? null : (util > 1 ? util / 100 : util),
               utilTargetSource: "08 Config.xlsx",
               gmNorm: r[6] === "" || r[6] == null ? null : (p.num(r[6]) > 1 ? p.num(r[6]) / 100 : p.num(r[6])),
               npNorm: r[7] === "" || r[7] == null ? null : (p.num(r[7]) > 1 ? p.num(r[7]) / 100 : p.num(r[7])),
               tags: String(r[8] || "").split(",").map(function (x) { return x.trim().toUpperCase(); }).filter(Boolean),
               suffix: String(r[9] || "") || null, manager: String(r[10] || ""),
               normNote: String(r[11] || ""), normAssumed: false };
    });
    if (depts.length) {
      depts.push({ code: "UNALLOCATED", short: "Unallocated", long: "No department on the line", order: 99,
                   isRevenue: false, utilTarget: null, gmNorm: null, npNorm: null, tags: [], suffix: null,
                   normNote: "A line with no cost centre and no bracket tag.", normAssumed: false });
      cfg.DEPARTMENTS = depts;
    }

    // Split bases and sub splits
    var sb = tab("Split bases");
    var bases = [], subs = {};
    sb.forEach(function (r) {
      var name = String(r[0] || "").trim();
      if (!name) return;
      var v = p.num(r[1]); if (v > 1) v /= 100;
      if (/production share/i.test(name)) subs.prd = v;
      else if (/integration share/i.test(name)) subs.int = v;
      else {
        var a = p.num(r[1]), b = p.num(r[2]), c = p.num(r[3]);
        if (a > 1) { a /= 100; b /= 100; c /= 100; }
        bases.push({ code: /3 ?way/i.test(name) ? "3WAY" : /staff/i.test(name) ? "STAFF" : "OFFICE",
                     label: name, source: /confirm/i.test(String(r[4] || "")) ? "confirmed" : "placeholder",
                     note: String(r[5] || ""), groups: { "PRD/VID": a, "ONS": b, "CONS/INT": c } });
      }
    });
    if (bases.length) cfg.SPLIT_BASES = bases;
    cfg.SUB_SPLITS = cfg.SUB_SPLITS || {};
    if (subs.prd) cfg.SUB_SPLITS.productionShareOfPrdVid = subs.prd;
    if (subs.int) cfg.SUB_SPLITS.integrationShareOfConsInt = subs.int;
    var assign = {};
    tab("Split assignment").forEach(function (r) {
      if (!r[0]) return;
      var b = String(r[1] || "");
      assign[String(r[0]).trim()] = /staff/i.test(b) ? "STAFF" : /office/i.test(b) ? "OFFICE" : "3WAY";
    });
    if (Object.keys(assign).length) cfg.SPLIT_ASSIGNMENT = { source: "08 Config.xlsx", default: "3WAY", bySub: assign };

    // Holidays: written into the calendar file rather than config
    var holidays = tab("Holidays").filter(function (r) { return r[0]; }).map(function (r) {
      var iso = p.isoDate(r[0]);
      var st = String(r[1] || "All").toUpperCase();
      return iso ? { date: iso, name: String(r[2] || ""), states: st === "ALL" ? null : st.split(",").map(function (x) { return x.trim(); }),
                     certain: /fixed/i.test(String(r[3] || "")) } : null;
    }).filter(Boolean);

    // Distribution
    var dist = tab("Distribution").filter(function (r) { return r[0]; }).map(function (r) {
      return { name: String(r[0]), email: String(r[1] || ""), tier: p.num(r[2]) || 3,
               dept: String(r[3] || "").toUpperCase() || null, send: bool(r[4]), note: String(r[5] || ""),
               cadence: /fort/i.test(String(r[6] || "")) ? (/both|month/i.test(String(r[6])) ? "both" : "fortnightly") : "monthly" };
    });
    cfg.DISTRIBUTION = dist;

    // Users and tabs by role
    var ROLE_ID = { "finance head": "admin", "finance": "finance", "executive": "exec",
                    "department head": "deptHead", "viewer": "viewer" };
    var users = tab("Users").filter(function (r) { return r[0]; }).map(function (r) {
      return { name: String(r[0]), role: ROLE_ID[String(r[1] || "").toLowerCase()] || "viewer",
               dept: String(r[2] || "").toUpperCase() || null, email: String(r[3] || ""), note: String(r[4] || "") };
    });
    var tbr = tab("Tabs by role");
    var pages = [], rolePages = { admin: [], finance: [], exec: [], deptHead: [], viewer: [] };
    tbr.forEach(function (r) {
      var id = String(r[0] || "").trim();
      if (!id) return;
      pages.push({ id: id, title: String(r[1] || id) });
      rolePages.admin.push(id);
      if (bool(r[3])) rolePages.finance.push(id);
      if (bool(r[4])) rolePages.exec.push(id);
      if (bool(r[5])) rolePages.deptHead.push(id);
      if (bool(r[6])) rolePages.viewer.push(id);
    });
    var baseUsers = window.CTS_USERS || { ROLES: [], PAGES: [] };
    var roles = (baseUsers.ROLES || []).map(function (ro) {
      return Object.assign({}, ro, { pages: rolePages[ro.id] && rolePages[ro.id].length ? rolePages[ro.id] : ro.pages });
    });
    var pageMeta = {};
    (baseUsers.PAGES || []).forEach(function (pg) { pageMeta[pg.id] = pg; });
    var usersOut = {
      meta: { version: 3, updated: today(), note: "Built from 08 Config.xlsx. Visibility only, not a security boundary." },
      PAGES: pages.length ? pages.map(function (pg) { return Object.assign({ section: (pageMeta[pg.id] || {}).section || "", sensitive: (pageMeta[pg.id] || {}).sensitive }, pg); }) : baseUsers.PAGES,
      ROLES: roles, USERS: users.length ? users : baseUsers.USERS, DEFAULT_ROLE: "viewer",
    };
    cfg.META = { source: "08 Config.xlsx", built: today() };
    return { config: cfg, users: usersOut, holidays: holidays };
  };


  /** The forecast settings: Assumptions, Methods and Overrides tabs. Nothing
   *  is calculated here; the portal makes the forecast from these at run
   *  time, so the archive of a month is enough to remake its forecast. */
  C.forecast = function (wb) {
    var p = P(), e = E();
    function tab(name) { var t = B.table(wb, name); return t ? t.rows : []; }
    var ass = { horizonMonths: 12, runRateMonths: 3, enabled: true, growth: {} };
    var deptCodes = (e.depts || []).map(function (d) { return d.code; });
    tab("Assumptions").forEach(function (r) {
      var k = String(r[0] || "").toLowerCase().trim(), v = r[1];
      if (!k) return;
      if (/^horizon/.test(k)) ass.horizonMonths = +p.num(v) || 12;
      else if (/^run ?rate/.test(k)) ass.runRateMonths = +p.num(v) || 3;
      else if (/^forecast (is )?on|^enabled|^use the forecast/.test(k)) ass.enabled = !/^n|^off|^false|^0/.test(String(v || "").toLowerCase());
      else if (/^growth/.test(k)) {
        var code = deptCodes.filter(function (c) { return k.indexOf(c.toLowerCase()) >= 0; })[0] || "default";
        ass.growth[code] = +p.num(v) || 0;
      }
    });
    var pipe = { enabled: true, nearMonths: 3, cancellationRate: 0, countWon: false, never: {}, leadMonths: {}, probability: {},
                 orderStatuses: "Confirmed, Booked, In progress, Completed", quoteStatuses: "Accepted" };
    tab("Pipeline").forEach(function (r) {
      var k = String(r[0] || "").toLowerCase().trim(), v = r[1];
      if (!k) return;
      function deptIn(key) { return deptCodes.filter(function (c) { return key.indexOf(c.toLowerCase()) >= 0; })[0] || "default"; }
      if (/^use the pipeline/.test(k)) pipe.enabled = !/^n|^off|^false|^0/.test(String(v || "").toLowerCase());
      else if (/^near months/.test(k)) pipe.nearMonths = +p.num(v);
      else if (/^cancellation/.test(k)) pipe.cancellationRate = +p.num(v) || 0;
      else if (/^count won/.test(k)) pipe.countWon = /^y/i.test(String(v || ""));
      else if (/^never in the pipeline/.test(k)) pipe.never[deptIn(k)] = +p.num(v) || 0;
      else if (/^months from close/.test(k)) pipe.leadMonths[deptIn(k)] = +p.num(v) || 0;
      else if (/^confirmed order statuses/.test(k)) pipe.orderStatuses = String(v || pipe.orderStatuses);
      else if (/^accepted quote statuses/.test(k)) pipe.quoteStatuses = String(v || pipe.quoteStatuses);
      else if (/^probability:/.test(k)) { var st = String(r[0]).replace(/^probability:\s*/i, "").trim(); if (st && v !== "" && v != null) pipe.probability[st] = +p.num(v); }
    });
    ass.pipeline = pipe;
    var methods = [], badMethod = [];
    tab("Methods").forEach(function (r) {
      var match = String(r[0] || "").trim(); if (!match) return;
      var method = e.methodKey ? e.methodKey(r[1]) : null;
      if (!method) { badMethod.push(match + " (" + String(r[1] || "blank") + ")"); return; }
      var pv = r[2] === "" || r[2] == null ? null : p.num(r[2]);
      methods.push({ match: match, method: method, param: pv, note: String(r[3] || "").trim() });
    });
    var overrides = [], badOverride = [];
    var blankOverrides = 0;
    tab("Overrides").forEach(function (r) {
      var acct = String(r[1] || "").trim(); if (!acct && !r[3]) return;
      if (r[3] === "" || r[3] == null) { blankOverrides++; return; }   // a row with no amount is a note, not an override
      var mk = B.monthOfLabel(r[2]) || (p.isoDate(r[2]) || "").slice(0, 7) || null;
      if (!acct || !mk) { badOverride.push((acct || "(no account)") + " " + String(r[2] || "(no month)")); return; }
      overrides.push({ dept: String(r[0] || "").trim().toUpperCase(), account: acct, month: mk,
                       amount: p.num(r[3]), note: String(r[4] || "").trim() });
    });
    return { data: { meta: { built: today(), source: "09 Forecast.xlsx", version: 1 }, assumptions: ass, methods: methods, overrides: overrides },
             badMethod: badMethod, badOverride: badOverride, blankOverrides: blankOverrides };
  };


  /** Cash and commitments: Settings, Bank balances, Credit cards and
   *  Commitments tabs. Balances and cards keep every month typed, so the
   *  history of actual cash builds up beside the forecast. */
  C.cash = function (wb) {
    var p = P();
    function tab(name) { var t = B.table(wb, name); return t ? t.rows : []; }
    function monthOf(v) { return B.monthOfLabel(v) || (p.isoDate(v) || "").slice(0, 7) || null; }
    var settings = {};
    tab("Settings").forEach(function (r) {
      var k = String(r[0] || "").toLowerCase().trim(), v = r[1];
      if (!k) return;
      if (/^debtor/.test(k)) settings.debtorDays = p.num(v);
      else if (/^creditor/.test(k)) settings.creditorDays = p.num(v);
      else if (/^bas/.test(k)) settings.basFrequency = String(v || "");
      else if (/^super/.test(k)) settings.superTiming = String(v || "");
      else if (/^payg/.test(k)) settings.paygInstalment = p.num(v);
      else if (/^facility|^overdraft/.test(k)) settings.facility = p.num(v);
      else if (/^receivable|^opening receivable|^debtors owed/.test(k)) settings.openingAR = v === "" || v == null ? null : p.num(v);
      else if (/^payable|^opening payable|^creditors owed/.test(k)) settings.openingAP = v === "" || v == null ? null : p.num(v);
      else if (/^credit cards? paid/.test(k)) settings.cardsPaidInFull = String(v || "Yes");
    });
    var balances = [], badBal = 0;
    tab("Bank balances").forEach(function (r) {
      if (!r[0] && !r[1]) return;
      var mk = monthOf(r[0]); if (!mk) { badBal++; return; }
      balances.push({ month: mk, account: String(r[1] || "").trim(), balance: p.num(r[2]), note: String(r[3] || "").trim() });
    });
    var cards = [], badCard = 0;
    tab("Credit cards").forEach(function (r) {
      if (!r[0] && !r[1]) return;
      var mk = monthOf(r[0]); if (!mk) { badCard++; return; }
      cards.push({ month: mk, card: String(r[1] || "").trim(), holder: String(r[2] || "").trim(), limit: p.num(r[3]),
                   balance: p.num(r[4]), paymentDay: p.num(r[5]) || null, note: String(r[6] || "").trim() });
    });
    var commitments = [], badCom = [];
    tab("Commitments").forEach(function (r) {
      var name = String(r[0] || "").trim(); if (!name) return;
      var next = p.isoDate(r[5]);
      if (!next) { badCom.push(name); return; }
      commitments.push({ name: name, category: String(r[1] || "").trim(), amount: p.num(r[2]), gst: String(r[3] || ""),
                         frequency: String(r[4] || "Monthly"), next: next, end: p.isoDate(r[6]) || null,
                         inPnl: String(r[7] == null || r[7] === "" ? "Yes" : r[7]), note: String(r[8] || "").trim() });
    });
    return { data: { meta: { built: today(), source: "10 Cash and Commitments.xlsx", version: 1 }, settings: settings,
                     balances: balances, cards: cards, commitments: commitments },
             badBal: badBal, badCard: badCard, badCom: badCom };
  };

  /* ================================================ folder and building */
  function today() { return new Date().toISOString().slice(0, 10); }
  B.today = today;

  B.supported = function () { return typeof window.showDirectoryPicker === "function"; };

  B.pickFolder = async function () {
    var dir = await window.showDirectoryPicker({ id: "cts-portal", mode: "readwrite" });
    var ok = false;
    try { await dir.getFileHandle("CTS_bi_bundle.js"); await dir.getDirectoryHandle("templates"); ok = true; }
    catch (e) { ok = false; }
    return { handle: dir, ok: ok };
  };

  B.readTemplates = async function (dir) {
    var tdir = await dir.getDirectoryHandle("templates");
    var out = {}, missing = [];
    for (var i = 0; i < TEMPLATES.length; i++) {
      var t = TEMPLATES[i];
      try {
        var fh = await tdir.getFileHandle(t.file);
        var f = await fh.getFile();
        out[t.key] = { buffer: await f.arrayBuffer(), modified: f.lastModified };
      } catch (e) { missing.push(t.file); }
    }
    return { buffers: out, missing: missing };
  };

  /** Everything after reading, so a test can hand in buffers directly. */
  B.runFromBuffers = function (buffers, opts) {
    opts = opts || {};
    var results = {}, checks = [], outputs = {};
    function wbOf(key) {
      var b = buffers[key];
      return b ? XLSX.read(b.buffer || b, { type: "array", cellDates: true }) : null;
    }
    function check(level, area, text) { checks.push({ level: level, area: area, text: text }); }

    // Config first, so the month and departments used below are this month's
    var cfgWb = wbOf("config");
    var cfgOut = null;
    if (cfgWb) {
      try {
        cfgOut = C.config(cfgWb);
        outputs["CTS_config_data.js"] = "window.CTS_CONFIG = " + JSON.stringify(cfgOut.config) + ";\n";
        outputs["CTS_users_data.js"] = "window.CTS_USERS = " + JSON.stringify(cfgOut.users) + ";\n";
        var ph = (cfgOut.config.SPLIT_BASES || []).filter(function (b) { return b.source !== "confirmed"; });
        if (ph.length) check("warn", "Config", "Split bases still placeholder: " + ph.map(function (b) { return b.label; }).join(", ") + ". Every allocated figure rests on them.");
        else check("ok", "Config", "All three split bases confirmed.");
        check("ok", "Config", "Reporting month " + cfgOut.config.ORG.reportingMonth + ", " + (cfgOut.config.DEPARTMENTS || []).length + " departments, " + (cfgOut.users.USERS || []).length + " users, " + (cfgOut.config.DISTRIBUTION || []).length + " on the distribution list.");
      } catch (e) { check("fail", "Config", e.message); }
    } else check("warn", "Config", "08 Config.xlsx not found; keeping the current settings.");
    var reportingMonth = cfgOut ? cfgOut.config.ORG.reportingMonth : E().reportingMonth();

    // P&L and budget
    var finActual = null, finBudget = null, budgetDept = null;
    var pnlWb = wbOf("pnl");
    if (pnlWb) {
      try {
        var t = B.table(pnlWb, "P&L");
        var g = C.pnl(t);
        finActual = g.grid;
        check(g.matched ? "ok" : "fail", "P&L", g.matched + " accounts matched across " + g.months.length + " months" + (g.skippedTotals ? ", " + g.skippedTotals + " total rows skipped" : "") + ".");
        if (g.unknown.length) check("warn", "P&L", g.unknown.length + " rows not in the chart: " + g.unknown.slice(0, 5).join(", ") + (g.unknown.length > 5 ? " and others" : "") + ".");
      } catch (e) { check("fail", "P&L", e.message); }
    } else check("fail", "P&L", "01 Xero P&L.xlsx not found.");
    var budWb = wbOf("budget");
    if (budWb) {
      try {
        var tb = B.table(budWb, "Budget");
        var gb = C.budget(tb);
        finBudget = gb.grid;
        check(gb.matched ? "ok" : "warn", "Budget", gb.matched + " accounts budgeted across " + gb.months.length + " months.");
        var td = B.table(budWb, "Departments");
        if (td && td.rows.length) {
          var bd = C.budgetDept(td);
          budgetDept = bd.rows;
          var bdMonths = {}; bd.rows.forEach(function (r) { Object.keys(r.months).forEach(function (mk) { bdMonths[mk] = 1; }); });
          check(bd.rows.length ? "ok" : "warn", "Budget", bd.rows.length + " department group rows across " + Object.keys(bdMonths).length + " months on the Departments tab" + (bd.bad.length ? "; skipped " + bd.bad.join(", ") + " (no departments or category)" : "") + ".");
          // the groups should add to the account budget by category
          var diffs = [];
          ["income", "cos", "expenses"].forEach(function (cat) {
            Object.keys(bdMonths).forEach(function (mk) {
              var g = 0, a = 0;
              bd.rows.forEach(function (r) { if (r.cat === cat) g += r.months[mk] || 0; });
              Object.keys(gb.grid).forEach(function (n) { if ((E().acct[n] || {}).cat === cat) a += gb.grid[n][mk] || 0; });
              if (a && Math.abs(g - a) >= 100) diffs.push(mk + " " + cat + " " + ((g - a) / 100).toFixed(0));
            });
          });
          if (diffs.length) check("warn", "Budget", "Departments tab and account budget disagree by category: " + diffs.slice(0, 4).join("; ") + (diffs.length > 4 ? " and more" : "") + ". The departmental pages read the Departments tab; the P&L reads the accounts.");
        } else check("info", "Budget", "No Departments tab rows; departmental budgets are derived from last year's ledger share.");
      } catch (e) { check("fail", "Budget", e.message); }
    } else check("warn", "Budget", "07 Budget.xlsx not found; keeping the current budget.");
    // Calendar: the span of months the templates carry, padded to whole
    // financial years, so a new year appears when its first column does.
    var cal = null;
    ACTIVE.monthIdx = null;
    try {
      var span = {};
      function spanOf(grid) { Object.keys(grid || {}).forEach(function (a) { Object.keys(grid[a]).forEach(function (mk) { if (grid[a][mk]) span[mk] = 1; }); }); }
      spanOf(finActual); spanOf(finBudget);
      if (reportingMonth) span[reportingMonth] = 1;
      var keys = Object.keys(span);
      if (keys.length) {
        var baseCal = window.CTS_CAL || {};
        var hol = (cfgOut && cfgOut.holidays.length) ? cfgOut.holidays.map(function (h) { return { date: h.date, name: h.name, states: h.states || baseCal.states, certain: h.certain }; })
                                                     : (baseCal.holidays || []);
        cal = B.calendar(keys, hol, baseCal.states, { built: today(),
                         source: (cfgOut && cfgOut.holidays.length) ? "08 Config.xlsx Holidays tab" : "holidays carried from the previous calendar" });
        ACTIVE.monthIdx = {};
        cal.months.forEach(function (m) { ACTIVE.monthIdx[m.key] = m; });
        outputs["CTS_cal_data.js"] = "window.CTS_CAL = " + JSON.stringify(cal) + ";\n";
        results.cal = cal;
        var fys = []; cal.months.forEach(function (m) { if (fys.indexOf(m.fy) < 0) fys.push(m.fy); });
        var bare = fys.filter(function (fy) {
          return !hol.some(function (h) { var k = h.date.slice(0, 7); return ACTIVE.monthIdx[k] && ACTIVE.monthIdx[k].fy === fy; });
        });
        check("ok", "Calendar", "FY" + fys[0] + (fys.length > 1 ? " to FY" + fys[fys.length - 1] : "") + ", " + cal.months.length + " months, " + hol.length + " public holidays" + (cfgOut && cfgOut.holidays.length ? " from the Holidays tab." : " carried over."));
        if (bare.length) check("warn", "Calendar", "No public holidays listed for FY" + bare.join(", FY") + " on the Holidays tab. Working days in those months are plain weekdays until you add them, which overstates capacity and understates utilisation.");
        var rmM = ACTIVE.monthIdx[reportingMonth];
        if (rmM && rmM.m === 7) check("info", "Calendar", "Reporting month is July: FY" + rmM.fy + " has started. Budget for FY" + rmM.fy + " should be in 07 Budget.xlsx, and the P&L paste is now July only.");
      }
    } catch (e) { check("fail", "Calendar", e.message); }

    if (finActual || finBudget) {
      var cur = E().fin || {};
      var fin = { meta: { seed: false, built: today(), basis: "accrual", currency: "AUD", source: "01 Xero P&L.xlsx and 07 Budget.xlsx" },
                  actual: finActual || cur.actual || {}, budget: finBudget || cur.budget || {},
                  budgetDept: budgetDept || cur.budgetDept || [],
                  months: (cal ? cal.months : E().months).map(function (m) { return m.key; }) };
      outputs["CTS_fin_data.js"] = "window.CTS_FIN = " + JSON.stringify(fin) + ";\n";
      results.fin = fin;
    }

    // Ledger
    var glWb = wbOf("gl");
    if (glWb) {
      try {
        var gl;
        function detailHead(h) { return colIndex(h, ["runningbalance"]) >= 0 && colIndex(h, ["date"]) >= 0 && colIndex(h, ["debit"]) >= 0; }
        var td = B.table(glWb, "GL", { minFilled: 1, findHead: detailHead });
        var tg = td ? null : B.table(glWb, "GL");
        if (td) {
          gl = C.glDetail(td);
          var bs = Object.keys(gl.skipped);
          check("warn", "Ledger", "This is Xero's General Ledger Detail report, not the Account transactions for P&L analysis report. It carries no cost centres or job numbers, so nothing can be coded to a department: every departmental page reads as unallocated until the custom report is pasted instead. " + bs.length + " balance sheet accounts skipped.");
        } else gl = C.gl(tg);
        outputs["CTS_gl_data.js"] = "window.CTS_GL = " + JSON.stringify(gl.payload) + ";\n";
        results.gl = gl.payload;
        check(gl.payload.rows.length ? "ok" : "fail", "Ledger", gl.payload.rows.length + " lines read" + (gl.bad ? ", " + gl.bad + " skipped for no date or account" : "") + ".");
        var unk = Object.keys(gl.unknown);
        if (unk.length) check("warn", "Ledger", unk.length + " account names not in the chart: " + unk.slice(0, 5).join(", ") + (unk.length > 5 ? " and others" : "") + ". Add them to the chart or fix the spelling.");
        // the P&L control: ledger against the P&L by category for the months both cover
        if (finActual) {
          var byCat = {};
          gl.payload.rows.forEach(function (r) {
            var a = E().acct[r[1]]; if (!a) return;
            var mk = r[0].slice(0, 7);
            (byCat[mk] = byCat[mk] || {})[a.cat] = (byCat[mk][a.cat] || 0) + Math.round(r[4] * 100) - Math.round(r[3] * 100);
          });
          var plCat = {};
          Object.keys(finActual).forEach(function (n) {
            var a = E().acct[n]; if (!a) return;
            Object.keys(finActual[n]).forEach(function (mk) {
              (plCat[mk] = plCat[mk] || {})[a.cat] = (plCat[mk][a.cat] || 0) + finActual[n][mk];
            });
          });
          var bad = [];
          Object.keys(plCat).sort().forEach(function (mk) {
            if (!byCat[mk]) return;
            ["income", "cos", "expenses", "other_income"].forEach(function (cat) {
              var d = (byCat[mk][cat] || 0) - (plCat[mk][cat] || 0);
              if (Math.abs(d) >= 100) bad.push(mk + " " + cat + " " + (d / 100).toFixed(2));
            });
          });
          if (bad.length) check("warn", "Control", "Ledger and P&L disagree: " + bad.slice(0, 4).join("; ") + (bad.length > 4 ? " and more" : "") + ".");
          else check("ok", "Control", "Ledger agrees with the P&L in every month both cover.");
        }
        if (gl.autoCoded) check("info", "Ledger", gl.autoCoded + " Stripe fee lines with no cost centre coded to Admin, the way the analysis file does it by hand.");
        if (gl.noContact) check("warn", "Ledger", gl.noContact + " revenue lines carry no contact, so they cannot reach a client on the schedule or the top clients list. They count in the department's new business line.");
        var noDept = gl.payload.rows.filter(function (r) { return E().resolveDept(r[8], r[6]) === "UNALLOCATED"; }).length;
        if (noDept && !gl.detail) check("warn", "Ledger", noDept + " lines carry no department. Set the cost centre in Xero.");
        results.glRows = gl.payload.rows;
      } catch (e) { check("fail", "Ledger", e.message); }
    } else check("fail", "Ledger", "02 Xero GL Transactions.xlsx not found.");

    // Earnings
    var ehWb = wbOf("earnings");
    if (ehWb) {
      try {
        var te = B.table(ehWb, "Earnings", { minFilled: 1 });
        var locs = (B.table(ehWb, "Locations") || { rows: [] }).rows;
        var cats = (B.table(ehWb, "Pay categories") || { rows: [] }).rows;
        var adjs = (B.table(ehWb, "Adjustments") || { rows: [] }).rows;
        var eh = C.earnings(te, locs, cats, adjs, reportingMonth);
        outputs["CTS_util_data.js"] = "window.CTS_UTIL = " + JSON.stringify(eh.util) + ";\n";
        outputs["CTS_staff_data.js"] = "window.CTS_STAFF = " + JSON.stringify(eh.staff) + ";\n";
        results.util = eh.util; results.staff = eh.staff;
        check(eh.lines ? "ok" : "warn", "Earnings", eh.lines + " pay lines, " + eh.util.rows.length + " department months, " + eh.staff.rows.length + " people months" + (eh.reallocated ? ", " + eh.reallocated + " back pays reallocated to hours" : "") + ".");
        var nl = Object.keys(eh.newLocations);
        if (nl.length) check("warn", "Earnings", nl.length + " locations not on the Locations tab, guessed from the bracket tag: " + nl.slice(0, 4).join("; ") + (nl.length > 4 ? " and others" : "") + ". Add them to the tab to make the guess stick.");
        var nc = Object.keys(eh.newCategories);
        if (nc.length) check("warn", "Earnings", nc.length + " pay categories not on the Pay categories tab, guessed from the name: " + nc.slice(0, 4).join("; ") + ".");
        if (eh.blocks.length) check("info", "Earnings", eh.blocks.length + " pay run" + (eh.blocks.length > 1 ? "s" : "") + ": " + eh.blocks.map(function (b) { return (b.from || "?") + " to " + (b.to || "?"); }).join(", ") + ".");
        if (eh.spread) check("info", "Earnings", eh.spread + " lines had no date in their note (salaried hours, allowances) and were spread over their pay run's weekdays by month.");
        if (eh.usedFallback > eh.spread) check("warn", "Earnings", (eh.usedFallback - eh.spread) + " undated lines in a pay run with no dated lines at all were put in " + reportingMonth + ".");
      } catch (e) { check("fail", "Earnings", e.message); }
    } else check("info", "Earnings", "03 Employment Hero Earnings.xlsx not found; keeping the current hours.");

    // Pipeline
    var pipe = { meta: { seed: false, built: today(), placeholder: true,
                         note: "Deals, orders and quotes from the placeholder templates. Column shapes are provisional until a real export from each system is matched." },
                 deals: [], orders: [], quotes: [] };
    var any = false;
    [["deals", "Deals", C.deals], ["orders", "Orders", C.orders], ["quotes", "Quotes", C.quotes]].forEach(function (x) {
      var wb = wbOf(x[0]);
      if (!wb) return;
      try {
        if (x[0] === "deals") {
          var td = B.table(wb, "Deals", { minFilled: 1 });
          var dm = (B.table(wb, "Departments") || { rows: [] }).rows;
          var deals = C.deals(td, dm);
          pipe.deals = deals; any = true;
          var won = 0, lost = 0, open = 0; deals.forEach(function (d) { if (/won/i.test(d.stage)) won++; else if (/lost/i.test(d.stage)) lost++; else open++; });
          check("ok", "Pipeline", deals.length + " deals read from the Zoho report: " + won + " won, " + lost + " lost, " + open + " open across " + deals.groups + " stage groups.");
          var um = Object.keys(deals.unmapped);
          if (um.length) check("warn", "Pipeline", "Zoho departments with no portal department, left unallocated: " + um.map(function (k) { return k + " (" + deals.unmapped[k] + ")"; }).join(", ") + ". Set them on the Departments tab of 04 Zoho Deals.xlsx.");
          var noJob = deals.filter(function (d) { return !d.job; }).length; if (noJob) check("info", "Pipeline", noJob + " deals carry no Opportunity Number, so they cannot be matched to an order or quote.");
          return;
        }
        var tt = B.table(wb, x[1]);
        pipe[x[0]] = x[2](tt);
        any = true;
        check("ok", "Pipeline", pipe[x[0]].length + " " + x[0] + " read from the placeholder template.");
      } catch (e) { check("warn", "Pipeline", x[1] + ": " + e.message); }
    });
    pipe.meta.placeholder = !(pipe.deals.length && !pipe.orders.length && !pipe.quotes.length) && !!(pipe.orders.length || pipe.quotes.length);
    pipe.meta.note = pipe.meta.placeholder ? pipe.meta.note : "Deals from the Zoho All Deals by Stage report. Orders and quotes templates are still placeholder shapes until a real export from each is matched.";
    if (any) { outputs["CTS_pipeline_data.js"] = "window.CTS_PIPELINE = " + JSON.stringify(pipe) + ";\n"; results.pipeline = pipe; }

    // Forecast settings
    var fcWb = wbOf("forecast");
    if (fcWb) {
      try {
        var fcr = C.forecast(fcWb);
        outputs["CTS_forecast_data.js"] = "window.CTS_FORECAST = " + JSON.stringify(fcr.data) + ";\n";
        results.forecast = fcr.data;
        var unknownAcct = fcr.data.overrides.filter(function (o) { return !E().acct[o.account]; }).map(function (o) { return o.account; });
        check("ok", "Forecast", (fcr.data.assumptions.enabled ? "On. " : "Off. ") + fcr.data.methods.length + " method lines, " + fcr.data.overrides.length + " typed overrides, horizon " + fcr.data.assumptions.horizonMonths + " months, run rate " + fcr.data.assumptions.runRateMonths + " months. Pipeline layer " + (fcr.data.assumptions.pipeline.enabled ? "on, near months " + fcr.data.assumptions.pipeline.nearMonths : "off") + ".");
        if (fcr.badMethod.length) check("warn", "Forecast", fcr.badMethod.length + " method lines not understood and skipped: " + fcr.badMethod.slice(0, 4).join("; ") + ". Use Seasonal, Run rate, % of revenue, Fixed, Budget or Zero.");
        if (fcr.blankOverrides) check("info", "Forecast", fcr.blankOverrides + " override row" + (fcr.blankOverrides > 1 ? "s" : "") + " with no amount ignored.");
        if (fcr.badOverride.length) check("warn", "Forecast", fcr.badOverride.length + " overrides without an account or a month, skipped: " + fcr.badOverride.slice(0, 4).join("; ") + ".");
        if (unknownAcct.length) check("warn", "Forecast", "Override accounts not in the chart: " + unknownAcct.slice(0, 4).join("; ") + ". Spell them as Xero does.");
      } catch (e) { check("fail", "Forecast", e.message); }
    } else check("info", "Forecast", "09 Forecast.xlsx not found; keeping the current forecast settings.");

    // Cash and commitments
    var cashWb = wbOf("cash");
    if (cashWb) {
      try {
        var cr = C.cash(cashWb);
        outputs["CTS_cash_data.js"] = "window.CTS_CASH = " + JSON.stringify(cr.data) + ";\n";
        results.cash = cr.data;
        var atRm = cr.data.balances.filter(function (b) { return b.month === reportingMonth; });
        var opening = atRm.reduce(function (a, b) { return a + (b.balance || 0); }, 0);
        check(atRm.length ? "ok" : "warn", "Cash", atRm.length
          ? atRm.length + " bank balance" + (atRm.length > 1 ? "s" : "") + " at " + reportingMonth + " totalling " + opening.toFixed(2) + "."
          : "No bank balance typed for " + reportingMonth + " on the Bank balances tab; the cash flow will start from zero.");
        var cardsRm = cr.data.cards.filter(function (b) { return b.month === reportingMonth; });
        check("ok", "Cash", cardsRm.length + " credit card" + (cardsRm.length === 1 ? "" : "s") + " at " + reportingMonth + ", " + cr.data.commitments.length + " commitments, debtor days " + (cr.data.settings.debtorDays || 45) + ", creditor days " + (cr.data.settings.creditorDays || 30) + ".");
        if (cr.badBal || cr.badCard) check("warn", "Cash", (cr.badBal + cr.badCard) + " balance rows without a readable month were skipped.");
        if (cr.badCom.length) check("warn", "Cash", "Commitments without a next due date were skipped: " + cr.badCom.slice(0, 4).join("; ") + ".");
      } catch (e) { check("fail", "Cash", e.message); }
    } else check("info", "Cash", "10 Cash and Commitments.xlsx not found; keeping the current cash settings.");

    // Aged receivables and payables
    var arWb = wbOf("ar"), apWb = wbOf("ap");
    if (arWb || apWb) {
      try {
        var aged = { meta: { seed: false, built: today(), asAt: reportingMonth, note: "Built from 11 Xero Aged Receivables.xlsx and 12 Xero Aged Payables.xlsx, as at the reporting month end." }, receivables: [], payables: [] };
        var cur = E().aged || {};
        if (arWb) { var ta = B.table(arWb, "Receivables"); var ar = C.aged(ta); aged.receivables = ar.rows; } else aged.receivables = cur.receivables || [];
        if (apWb) { var tp = B.table(apWb, "Payables"); var ap = C.aged(tp); aged.payables = ap.rows; } else aged.payables = cur.payables || [];
        outputs["CTS_aged_data.js"] = "window.CTS_AGED = " + JSON.stringify(aged) + ";\n";
        results.aged = aged;
        function sum(rows, k) { return rows.reduce(function (a, r) { return a + (r[k] || 0); }, 0) / 100; }
        var arTot = sum(aged.receivables, "total"), arOver = arTot - sum(aged.receivables, "current");
        var apTot = sum(aged.payables, "total"), apOver = apTot - sum(aged.payables, "current");
        check("ok", "Aged", aged.receivables.length + " open invoices totalling " + arTot.toFixed(2) + " (" + arOver.toFixed(2) + " past due); " + aged.payables.length + " open bills totalling " + apTot.toFixed(2) + " (" + apOver.toFixed(2) + " past due).");
      } catch (e) { check("fail", "Aged", e.message); }
    } else check("info", "Aged", "No aged receivables or payables template found; the revenue risk page has no collection view.");

    // Revenue schedule and the client list
    var schWb = wbOf("schedule");
    if (schWb) {
      try {
        var ts = B.table(schWb, "Schedule"), tc = B.table(schWb, "Clients");
        var sr = C.schedule(ts, tc ? tc.rows : []);
        outputs["CTS_clients_data.js"] = "window.CTS_CLIENTS = " + JSON.stringify(sr.payload) + ";\n";
        results.clients = sr.payload;
        var byDept = {}; sr.payload.schedule.forEach(function (l) { byDept[l.dept] = (byDept[l.dept] || 0) + 1; });
        check(sr.payload.schedule.length ? "ok" : "warn", "Schedule", sr.payload.schedule.length + " schedule lines (" + Object.keys(byDept).map(function (d) { return d + " " + byDept[d]; }).join(", ") + "), " + sr.typed + " with a typed budget, " + sr.payload.clients.length + " clients" + (sr.bad ? "; " + sr.bad + " rows skipped for no client or role" : "") + ".");
        if (!sr.typed) check("info", "Schedule", "No monthly amounts typed yet, so the line budgets are derived from the department budget on last year's share. Type the sales forecast to replace that.");
        var noNew = ["CONSULTING", "INTEGRATION", "PRODUCTION", "VIDEO"].filter(function (d) { return byDept[d] && !sr.payload.schedule.some(function (l) { return l.dept === d && l.kind === "newbiz"; }); });
        if (noNew.length) check("warn", "Schedule", "No New business line for " + noNew.join(", ") + ": revenue from unlisted clients has nowhere to land.");
        // ledger contacts with revenue who are not on the schedule, and the reverse
        if (results.glRows) {
          var alias = {}; sr.payload.clients.forEach(function (c) { (c.aliases || []).forEach(function (a) { alias[a.toLowerCase()] = c.contact; }); });
          var onSched = {}; sr.payload.schedule.forEach(function (l) { if (l.client) onSched[l.client.toLowerCase()] = 1; });
          var seenRev = {}, unlisted = {}, matched = 0;
          results.glRows.forEach(function (r) {
            var a = E().acct[r[1]]; if (!a || a.cat !== "income" || !r[2]) return;
            var k = (alias[r[2].toLowerCase()] || r[2]).toLowerCase();
            if (onSched[k]) { seenRev[k] = 1; matched++; } else unlisted[r[2]] = (unlisted[r[2]] || 0) + Math.round(r[4] * 100) - Math.round(r[3] * 100);
          });
          var ul = Object.keys(unlisted).sort(function (a, b) { return unlisted[b] - unlisted[a]; });
          var unseen = Object.keys(onSched).filter(function (k) { return !seenRev[k]; }).length;
          check("info", "Schedule", matched + " revenue lines matched to a scheduled client; " + ul.length + " contacts with revenue are not on the schedule and count as new business" + (ul.length ? ": " + ul.slice(0, 5).join(", ") + (ul.length > 5 ? " and others" : "") : "") + ". " + unseen + " scheduled clients have no ledger revenue in the file; check the spelling or add an alias on the Clients tab.");
        }
      } catch (e) { check("fail", "Schedule", e.message); }
    } else check("info", "Schedule", "13 Revenue Schedule.xlsx not found; keeping the current schedule.");

    var log = { built: new Date().toISOString(), reportingMonth: reportingMonth, checks: checks,
                files: Object.keys(outputs), version: (cfgOut && cfgOut.config.VERSION) || (E().cfg && E().cfg.VERSION) };
    outputs["_build_log.json"] = JSON.stringify(log, null, 2);
    ACTIVE.monthIdx = null;
    return { checks: checks, outputs: outputs, results: results, log: log };
  };

  B.run = async function (dir) {
    var read = await B.readTemplates(dir);
    var r = B.runFromBuffers(read.buffers);
    read.missing.forEach(function (f) {
      var t = TEMPLATES.filter(function (x) { return x.file === f; })[0];
      if (t && t.need) r.checks.unshift({ level: "fail", area: "Templates", text: f + " is missing from the templates folder." });
    });
    r.missing = read.missing;
    r.buffers = read.buffers;
    return r;
  };

  /** Write the outputs. Each existing file is copied into data/_previous
   *  first, so one folder holds the rollback: copy it back over data.
   *  With an archive {month, buffers} the same outputs and the templates
   *  they came from are also copied into archive/YYYY-MM, which is the
   *  permanent record of what the portal showed for that month. */
  B.write = async function (dir, outputs, progress, archive) {
    var dataDir = await dir.getDirectoryHandle("data", { create: true });
    var prev = await dataDir.getDirectoryHandle("_previous", { create: true });
    var names = Object.keys(outputs), done = [];
    async function put(folder, name, content) {
      var fh = await folder.getFileHandle(name, { create: true });
      var w = await fh.createWritable();
      await w.write(content); await w.close();
    }
    for (var i = 0; i < names.length; i++) {
      var name = names[i];
      try {
        var old = await dataDir.getFileHandle(name);
        var of = await old.getFile();
        await put(prev, name, await of.text());
      } catch (e) { /* nothing to keep */ }
      await put(dataDir, name, outputs[name]);
      done.push(name);
      if (progress) progress(name, i + 1, names.length);
    }
    if (archive && archive.month) {
      var arc = await dir.getDirectoryHandle("archive", { create: true });
      var md = await arc.getDirectoryHandle(archive.month, { create: true });
      var ad = await md.getDirectoryHandle("data", { create: true });
      for (var j = 0; j < names.length; j++) await put(ad, names[j], outputs[names[j]]);
      var bufs = archive.buffers || {};
      var td = await md.getDirectoryHandle("templates", { create: true });
      var keys = Object.keys(bufs), copied = [];
      for (var k = 0; k < keys.length; k++) {
        var t = TEMPLATES.filter(function (x) { return x.key === keys[k]; })[0];
        var b = bufs[keys[k]];
        if (!t || !b) continue;
        await put(td, t.file, b.buffer || b);
        copied.push(t.file);
      }
      await put(md, "_archive.json", JSON.stringify({ month: archive.month, written: new Date().toISOString(),
                                                       data: names, templates: copied }, null, 2));
    }
    return done;
  };

  /** Publish one data file from the portal (commentary, recipients): the
   *  previous copy goes to data/_previous, and OneDrive carries the new one. */
  B.writeDataFile = async function (dir, name, global, data) {
    var dataDir = await dir.getDirectoryHandle("data", { create: true });
    var prev = await dataDir.getDirectoryHandle("_previous", { create: true });
    try {
      var old = await dataDir.getFileHandle(name); var of = await old.getFile();
      var ph = await prev.getFileHandle(name, { create: true }); var pw = await ph.createWritable();
      await pw.write(await of.text()); await pw.close();
    } catch (e) { /* first publish */ }
    var fh = await dataDir.getFileHandle(name, { create: true });
    var w = await fh.createWritable();
    await w.write("window." + global + " = " + JSON.stringify(data, null, 1) + ";\n"); await w.close();
    return name;
  };
  B.writeCommentary = function (dir, data) { return B.writeDataFile(dir, "CTS_commentary_data.js", "CTS_COMMENTARY", data); };
  B.writeRecipients = function (dir, list) {
    return B.writeDataFile(dir, "CTS_distribution_data.js", "CTS_DISTRIBUTION", { meta: { published: new Date().toISOString(), note: "Recipients edited in the portal. Wins over the Config template's Distribution tab." }, list: list });
  };

  /** Write the month's emails into outbox/YYYY-MM, one JSON and one HTML per
   *  recipient, which is what the Power Automate flow reads. */
  B.writeBinary = async function (dir, month, name, bytes) {
    var ob = await dir.getDirectoryHandle("outbox", { create: true });
    var md = await ob.getDirectoryHandle(month, { create: true });
    var fh = await md.getFileHandle(name, { create: true });
    var w = await fh.createWritable(); await w.write(bytes); await w.close();
    return name;
  };
  B.writeOutbox = async function (dir, month, messages, folder, prefix) {
    var ob = await dir.getDirectoryHandle("outbox", { create: true });
    var md = await ob.getDirectoryHandle(folder || month, { create: true });
    var written = [];
    for (var i = 0; i < messages.length; i++) {
      var m = messages[i];
      var base = (prefix || "") + String(m.to || m.name || ("recipient" + i)).replace(/[^a-z0-9]+/gi, "_").toLowerCase();
      var jh = await md.getFileHandle(base + ".json", { create: true });
      var jw = await jh.createWritable();
      await jw.write(JSON.stringify({ to: m.to, name: m.name, tier: m.tier, subject: m.subject,
                                      html: m.html, text: m.text, month: month, written: new Date().toISOString(),
                                      kind: m.kind || "monthly", attachments: m.attachments || [],
                                      status: "pending" }, null, 2));
      await jw.close();
      var hh = await md.getFileHandle(base + ".html", { create: true });
      var hw = await hh.createWritable();
      await hw.write(m.html); await hw.close();
      written.push(base);
    }
    var ih = await md.getFileHandle((prefix || "") + "_manifest.json", { create: true });
    var iw = await ih.createWritable();
    await iw.write(JSON.stringify({ month: month, count: messages.length, files: written,
                                    written: new Date().toISOString() }, null, 2));
    await iw.close();
    return written;
  };
})();
