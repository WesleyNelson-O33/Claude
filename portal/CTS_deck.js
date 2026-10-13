/* CTS Business Intelligence Portal - the Monthly Business Review deck
 * ---------------------------------------------------------------------------
 * Builds the Monthly Business Review as a PowerPoint file from the same
 * engine the pages read, in the layout of the September 2026 template:
 * cover, agenda, four sections (financial performance, customers and cash,
 * people and utilisation, summary and actions), each slide with its
 * commentary panel. Commentary is what was typed on the matching pages;
 * where nothing was typed the portal's proposed commentary stands in,
 * marked as proposed. One click on Distribution makes it; a second writes
 * it into the outbox so the flow sends it. Rendering is PptxGenJS
 * (vendor/pptxgen.bundle.js, MIT).
 */
(function () {
  "use strict";
  var CTS = (window.CTS = window.CTS || {});
  var D = (CTS.deck = {});
  var E = function () { return CTS.engine; };
  var F = function () { return CTS.fmt; };

  var C = { black: "121820", white: "FFFFFF", grey: "5F6670", greyLight: "C9CED6", panel: "F4F4F4", line: "E3E6EA",
            navy: "122B82", mint: "3FD0C9", lavender: "6D71FF", rose: "F33844", teal: "005358", burgundy: "A8052E",
            ink: "1A1A1A", comment: "E8F9F8", good: "0A6B2C", bad: "A0161F" };
  var DEPT = { ONSITE: "F33A45", PRODUCTION: "0E9E99", VIDEO: "B7850A", INTEGRATION: "2E4DA6", CONSULTING: "6C70FF", ADMIN: "018301", UNALLOCATED: "7B8391" };
  var FONT = "Calibri";
  var W = 10, H = 5.625;
  var ADDBACK = ["Depreciation", "Depreciation - ROU", "Instant Asset Write Off", "Interest Expense", "Income Tax Expense", "Under/(Over) provision for tax"];

  D.supported = function () { return typeof window.PptxGenJS === "function"; };
  D.fileName = function () {
    var e = E(), rm = e.reportingMonth();
    return "CTS Monthly Business Review - " + ((e.monthIdx[rm] || {}).long || rm).replace(/\s+/g, " ") + ".pptx";
  };

  function dollars(c) { return F().dollars(c); }
  function money(c) { return F().money(c); }
  function pct(x, dp) { return F().pct(x, dp); }
  /** A$ thousands, brackets for negative, the way the template shows them. */
  function k(c) { var n = Math.round((c || 0) / 100000); var s = Math.abs(n).toLocaleString("en-AU"); return n < 0 ? "(" + s + ")" : s; }
  function kSigned(c) { var n = Math.round((c || 0) / 100000); return (n > 0 ? "+" : n < 0 ? "-" : "") + Math.abs(n).toLocaleString("en-AU"); }
  function pSigned(x, dp) { if (x == null || !isFinite(x)) return ""; var v = (x * 100).toFixed(dp == null ? 1 : dp); return (x > 0 ? "+" : "") + v + "%"; }
  function kNum(c) { return Math.round((c || 0) / 100000); }
  function perHr(c) { var n = Math.round((c || 0) / 100); return n < 0 ? "($" + Math.abs(n) + ")" : "$" + n; }

  /* ---- slide furniture -------------------------------------------------- */
  var _pptx = null, _n = 0;
  function pptx() { return _pptx; }
  function number(s) { _n++; s.addText(String(_n), { x: 8.85, y: 5.21, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 9, color: C.grey, align: "right", isTextBox: true, margin: 0 }); }
  function frame(section, title) {
    var s = pptx().addSlide();
    s.background = { color: C.white };
    s.addText(section.toUpperCase(), { x: 0.55, y: 0.4, w: 7.0, h: 0.3, fontFace: FONT, fontSize: 10, bold: true, color: C.navy, charSpacing: 2, isTextBox: true, margin: 0 });
    s.addText(title, { x: 0.55, y: 0.7, w: 7.6, h: 0.75, fontFace: FONT, fontSize: 24, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
    var b = window.CTS_BRAND;
    if (b && b.logo && b.logo.light) s.addImage({ data: b.logo.light, x: W - 1.55, y: 0.35, w: 1.0, h: 0.68 });
    number(s);
    return s;
  }
  function dark(s) { s.background = { color: C.black }; }
  function divider(n, title, sub) {
    var s = pptx().addSlide(); dark(s);
    s.addText(n, { x: 0.55, y: 2.05, w: 1.4, h: 1.4, fontFace: FONT, fontSize: 60, bold: true, color: C.mint, isTextBox: true, margin: 0 });
    s.addText(title, { x: 2.3, y: 2.05, w: 6.6, h: 0.9, fontFace: FONT, fontSize: 32, bold: true, color: C.white, isTextBox: true, margin: 0 });
    s.addText(sub, { x: 2.3, y: 2.98, w: 6.6, h: 0.5, fontFace: FONT, fontSize: 14, color: C.greyLight, isTextBox: true, margin: 0 });
    _n++; s.addText(String(_n), { x: 8.85, y: 5.21, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 9, color: C.greyLight, align: "right", isTextBox: true, margin: 0 });
    return s;
  }
  function stat(s, x, value, label, colour, w) {
    var ww = w || 1.9;
    s.addShape(pptx().ShapeType.rect, { x: x, y: 1.6, w: ww, h: 0.7, fill: { color: C.panel }, line: { color: C.panel } });
    s.addText(value, { x: x + 0.15, y: 1.7, w: ww - 0.3, h: 0.36, fontFace: FONT, fontSize: 18, bold: true, color: colour || C.ink, isTextBox: true, margin: 0 });
    s.addText(label, { x: x + 0.15, y: 2.03, w: ww - 0.3, h: 0.24, fontFace: FONT, fontSize: 9, color: C.grey, isTextBox: true, margin: 0 });
  }
  function caption(s, text, y) {
    s.addText(text, { x: 0.55, y: y, w: 6.1, h: 0.22, fontFace: FONT, fontSize: 9.5, bold: true, color: C.grey, isTextBox: true, margin: 0 });
  }
  function footnote(s, text, y, w) {
    s.addText(text, { x: 0.55, y: y == null ? 4.8 : y, w: w || 6.1, h: 0.35, fontFace: FONT, fontSize: 8.5, color: C.grey, italic: true, isTextBox: true, margin: 0, valign: "top" });
  }
  /** The commentary panel on the right: what was typed on the pages for the
   *  month, or the portal's proposed text when nothing was. */
  function commentary(s, pageIds, prompts) {
    var e = E(), rm = e.reportingMonth(), items = [];
    pageIds.forEach(function (pg) { (e.commentaryFor ? e.commentaryFor(pg, rm, { all: true }) : []).forEach(function (c) { items.push(c); }); });
    s.addShape(pptx().ShapeType.rect, { x: 6.85, y: 1.6, w: 2.6, h: 3.45, fill: { color: C.comment }, line: { color: C.comment } });
    s.addText("COMMENTARY", { x: 7.05, y: 1.78, w: 2.2, h: 0.25, fontFace: FONT, fontSize: 9, bold: true, color: C.teal, charSpacing: 2, isTextBox: true, margin: 0 });
    var runs = [];
    if (items.length) {
      items.slice(0, 5).forEach(function (c) {
        runs.push({ text: c.text.replace(/\s+/g, " "), options: { color: C.grey, fontSize: 9, breakLine: true, paraSpaceAfter: 4, bullet: true } });
      });
    } else {
      var proposed = null;
      for (var i = 0; i < pageIds.length && !proposed; i++) { try { proposed = e.proposeCommentary ? e.proposeCommentary(pageIds[i]) : null; } catch (x) { proposed = null; } }
      if (proposed) {
        runs.push({ text: "Proposed from the data; edit on the " + ((CTS.pages[pageIds[0]] || {}).title || pageIds[0]) + " page.", options: { color: C.teal, fontSize: 8, italic: true, breakLine: true, paraSpaceAfter: 4 } });
        proposed.replace(/\s*\[[^\]]*\]\s*$/, "").split(/(?<=\.)\s+(?=[A-Z0-9])/).slice(0, 5).forEach(function (t) { if (t.trim()) runs.push({ text: t.trim(), options: { color: C.grey, fontSize: 9, breakLine: true, paraSpaceAfter: 4, bullet: true } }); });
      }
      (prompts || []).forEach(function (t) { runs.push({ text: t, options: { color: C.greyLight, fontSize: 8.5, italic: true, breakLine: true, paraSpaceAfter: 3, bullet: true } }); });
    }
    if (runs.length) { runs[runs.length - 1].options.breakLine = false; s.addText(runs, { x: 7.05, y: 2.08, w: 2.2, h: 2.85, fontFace: FONT, valign: "top", isTextBox: true, margin: 0 }); }
    return items.length;
  }
  function tableOpts(extra) {
    return Object.assign({ fontFace: FONT, fontSize: 9, color: C.ink, border: { type: "solid", pt: 0.5, color: C.line }, autoPage: false, rowH: 0.24, margin: 0.04 }, extra || {});
  }
  function hdr(cols) { return cols.map(function (c) { return { text: c.label, options: { bold: true, color: C.grey, fill: { color: C.panel }, fontSize: 8.5, align: c.align || "right" } }; }); }
  function cell(v, o) { return { text: v == null ? "" : String(v), options: Object.assign({ align: "right" }, o || {}) }; }
  function kCell(c, bold) { return cell(k(c), { bold: !!bold }); }
  function varCell(c, bold) { return cell(kSigned(c), { color: c < 0 ? C.bad : c > 0 ? C.good : C.ink, bold: !!bold }); }
  function totalRow(cells) { return cells.map(function (x) { x.options = Object.assign({}, x.options, { bold: true, fill: { color: "EEF0F3" } }); return x; }); }
  function chartOpts(extra) {
    return Object.assign({ catAxisLabelFontSize: 8, valAxisLabelFontSize: 8, catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, catAxisLabelColor: C.grey, valAxisLabelColor: C.grey,
                           valGridLine: { color: C.line, size: 0.5 }, catGridLine: { style: "none" }, legendFontSize: 8, legendFontFace: FONT, legendColor: C.grey, legendPos: "b",
                           dataLabelFontFace: FONT, dataLabelFontSize: 7, dataLabelColor: C.grey }, extra || {});
  }
  function sumAcct(grid, names, keys) { var t = 0; names.forEach(function (n) { t += E().sumMonths(grid[n], keys); }); return t; }
  /** EBITDA: net profit with depreciation, the instant asset write off,
   *  interest and tax added back. The chart of accounts has no other
   *  below the line items. */
  function ebitda(keys, which) {
    var e = E();
    if (which === "budget") { var b = e.pnlBudget(keys).totals; return b.netProfit - sumAcct(e.finBudget, ADDBACK, keys); }
    var a = e.pnlActual(keys).totals; return a.netProfit - sumAcct(e.idx.acctMonth, ADDBACK, keys);
  }
  function deptIncome(code, keys) { return E().sumMonths((E().idx.deptCatMonth[code] || {}).income, keys); }

  /* ---- the deck ----------------------------------------------------------- */
  D.build = function () { return E().withoutView ? E().withoutView(buildInner) : buildInner(); };
  function buildInner() {
    if (!D.supported()) throw new Error("PptxGenJS did not load (vendor/pptxgen.bundle.js).");
    var e = E(), rm = e.reportingMonth(), mi = e.monthIdx[rm] || {}, mLong = mi.long || rm, mShort = mi.label || rm;
    var fy = e.currentFY(), ytd = e.ytd(), pyMonth = e.priorYearMonths([rm]), pyYtd = e.priorYearMonths(ytd);
    var ytdLabel = "YTD (" + (e.monthIdx[ytd[0]] || {}).label.split("-")[0] + " to " + mShort.split("-")[0] + ")";
    var P = new window.PptxGenJS(); _pptx = P; _n = 0;
    P.layout = "LAYOUT_16x9"; P.author = "CTS Business Intelligence Portal"; P.company = "Corporate Technology Services"; P.title = "CTS Monthly Business Review, " + mLong;
    var b = window.CTS_BRAND;
    var who = (e.identity && e.identity()) || "Finance";
    var revDepts = e.depts.filter(function (d) { return d.isRevenue && d.code !== "ADMIN"; });

    // 1 cover
    var s = P.addSlide(); dark(s); _n++;
    if (b && b.logo && b.logo.dark) s.addImage({ data: b.logo.dark, x: 0.55, y: 0.5, w: 1.6, h: 1.09 });
    s.addText("MONTHLY BUSINESS REVIEW", { x: 0.55, y: 2.0, w: 6.2, h: 0.35, fontFace: FONT, fontSize: 12, bold: true, color: C.mint, charSpacing: 4, isTextBox: true, margin: 0 });
    s.addText(mLong, { x: 0.55, y: 2.35, w: 6.2, h: 1.35, fontFace: FONT, fontSize: 44, bold: true, color: C.white, isTextBox: true, margin: 0 });
    s.addText("Results, risk revenue, customers, cash and utilisation", { x: 0.55, y: 3.7, w: 6.2, h: 0.6, fontFace: FONT, fontSize: 14, color: C.greyLight, isTextBox: true, margin: 0 });
    s.addText(who, { x: 0.55, y: 4.65, w: 4.0, h: 0.55, fontFace: FONT, fontSize: 11, color: C.greyLight, isTextBox: true, margin: 0 });
    s.addText("Built " + new Date().toISOString().slice(0, 10) + " from the portal", { x: 6.5, y: 4.65, w: 2.95, h: 0.55, fontFace: FONT, fontSize: 11, color: C.greyLight, align: "right", isTextBox: true, margin: 0 });

    // 2 agenda
    s = frame("Agenda", "What we will cover");
    var agenda = [
      ["01", "Financial performance", "P&L summary, P&L per department, revenue and contribution, risk revenue profile"],
      ["02", "Customers and cash", "Top 10 customers, receivables and DSO, top customers year on year"],
      ["03", "People and utilisation", "Utilisation by department, profit based on utilisation"],
      ["04", "Summary and actions", "The commentary from each section, with an owner and a date"],
    ];
    agenda.forEach(function (a, i) {
      var y = 1.65 + i * 0.82;
      s.addShape(P.ShapeType.rect, { x: 0.55, y: y, w: 8.9, h: 0.7, fill: { color: C.panel }, line: { color: C.panel } });
      s.addText(a[0], { x: 0.75, y: y + 0.12, w: 0.7, h: 0.46, fontFace: FONT, fontSize: 22, bold: true, color: C.mint, isTextBox: true, margin: 0 });
      s.addText(a[1], { x: 1.55, y: y + 0.1, w: 7.7, h: 0.28, fontFace: FONT, fontSize: 13, bold: true, color: C.ink, isTextBox: true, margin: 0 });
      s.addText(a[2], { x: 1.55, y: y + 0.38, w: 7.7, h: 0.26, fontFace: FONT, fontSize: 9.5, color: C.grey, isTextBox: true, margin: 0 });
    });

    // 3 divider
    divider("01", "Financial performance", "P&L for " + mLong + " and " + ytdLabel);

    // 4 P&L summary
    var act = e.pnlActual([rm]).totals, bud = e.pnlBudget([rm]).totals, py = e.pnlActual(pyMonth).totals;
    var ytdA = e.pnlActual(ytd).totals, ytdB = e.pnlBudget(ytd).totals;
    var eb = { m: ebitda([rm]), mb: ebitda([rm], "budget"), y: ebitda(ytd), yb: ebitda(ytd, "budget"), py: ebitda(pyMonth) };
    s = frame("Financial performance", "P&L summary");
    var rows = [hdr([{ label: "A$ thousands", align: "left" }, { label: mShort.replace("-", " ") }, { label: "Budget" }, { label: "Var" }, { label: "YTD" }, { label: "Budget" }, { label: "Var" }, { label: "PY " + mShort.replace("-", " ") }])];
    function pnlRow(label, key, bold) { return [cell(label, { align: "left", bold: !!bold }), kCell(act[key], bold), kCell(bud[key]), varCell(act[key] - bud[key]), kCell(ytdA[key], bold), kCell(ytdB[key]), varCell(ytdA[key] - ytdB[key]), kCell(py[key])]; }
    function pctRow(label, num, den) {
      function r(n, d) { return d ? n / d : null; }
      var m = r(num.m, den.m), mb = r(num.mb, den.mb), y = r(num.y, den.y), yb = r(num.yb, den.yb), p = r(num.py, den.py);
      return [cell(label, { align: "left", italic: true, color: C.grey }), cell(pct(m), { italic: true, color: C.grey }), cell(pct(mb), { italic: true, color: C.grey }), cell(pSigned(m != null && mb != null ? m - mb : null), { italic: true, color: C.grey }),
              cell(pct(y), { italic: true, color: C.grey }), cell(pct(yb), { italic: true, color: C.grey }), cell(pSigned(y != null && yb != null ? y - yb : null), { italic: true, color: C.grey }), cell(pct(p), { italic: true, color: C.grey })];
    }
    var inc = { m: act.income, mb: bud.income, y: ytdA.income, yb: ytdB.income, py: py.income }, gp = { m: act.grossProfit, mb: bud.grossProfit, y: ytdA.grossProfit, yb: ytdB.grossProfit, py: py.grossProfit };
    rows.push(pnlRow("Revenue", "income", true));
    rows.push(pnlRow("Cost of sales", "cos"));
    rows.push(pnlRow("Gross profit", "grossProfit", true));
    rows.push(pctRow("Gross margin %", gp, inc));
    rows.push(pnlRow("Overheads", "expenses"));
    rows.push([cell("EBITDA", { align: "left", bold: true }), kCell(eb.m, true), kCell(eb.mb), varCell(eb.m - eb.mb), kCell(eb.y, true), kCell(eb.yb), varCell(eb.y - eb.yb), kCell(eb.py)]);
    rows.push(pctRow("EBITDA margin %", eb, inc));
    s.addTable(rows, tableOpts({ x: 0.55, y: 1.6, w: 6.1, colW: [1.45, 0.66, 0.66, 0.6, 0.66, 0.66, 0.6, 0.81] }));
    var last6 = e.actualMonths.filter(function (m) { return m <= rm; }).slice(-6);
    caption(s, "Revenue and EBITDA, last six months", 3.88);
    s.addChart(P.ChartType.bar, [
      { name: "Revenue", labels: last6.map(function (m) { return e.monthIdx[m].label; }), values: last6.map(function (m) { return kNum(e.pnlActual([m]).totals.income); }) },
      { name: "EBITDA", labels: last6.map(function (m) { return e.monthIdx[m].label; }), values: last6.map(function (m) { return kNum(ebitda([m])); }) },
    ], chartOpts({ x: 0.55, y: 4.08, w: 6.1, h: 1.05, barDir: "col", barGapWidthPct: 80, chartColors: [C.navy, C.mint], showLegend: true, legendPos: "r", valAxisLabelFormatCode: "#,##0", valAxisHidden: false }));
    commentary(s, ["home", "pnl", "bva"], ["What drove the month against budget", "YTD position and the run rate to full year", "One off items to call out"]);

    // 5 P&L per department
    var alloc = e.allocate([rm]), budInc = e.budgetByDept([rm], "income");
    s = frame("Financial performance", "P&L per department, " + mShort.replace("-", " "));
    rows = [hdr([{ label: "A$ thousands", align: "left" }, { label: "Revenue" }, { label: "Budget" }, { label: "Var" }, { label: "Gross profit" }, { label: "GM %" }, { label: "Overheads" }, { label: "Contribution" }])];
    var tot = { income: 0, budget: 0, gp: 0, own: 0 };
    alloc.rows.filter(function (r) { return r.dept.isRevenue && r.code !== "ADMIN"; }).forEach(function (r) {
      var bd = budInc[r.code] || 0, contrib = r.grossProfit + r.ownExpenses;
      tot.income += r.income; tot.budget += bd; tot.gp += r.grossProfit; tot.own += r.ownExpenses;
      rows.push([cell(r.dept.short, { align: "left", bold: true }), kCell(r.income), kCell(bd), varCell(r.income - bd), kCell(r.grossProfit), cell(pct(r.gmPct)), kCell(r.ownExpenses), kCell(contrib, true)]);
    });
    var admin = alloc.rows.filter(function (r) { return r.code === "ADMIN"; })[0];
    var adminOwn = admin ? admin.ownExpenses + admin.grossProfit : 0;
    rows.push([cell("Admin", { align: "left", bold: true }), kCell(admin ? admin.income : 0), kCell(budInc.ADMIN || 0), varCell((admin ? admin.income : 0) - (budInc.ADMIN || 0)), kCell(admin ? admin.grossProfit : 0), cell(""), kCell(admin ? admin.ownExpenses : 0), kCell(adminOwn, true)]);
    var unalloc = alloc.rows.filter(function (r) { return r.code === "UNALLOCATED" && (r.income || r.ownExpenses); })[0];
    if (unalloc) rows.push([cell("Unallocated", { align: "left", color: C.grey }), kCell(unalloc.income), kCell(0), varCell(unalloc.income), kCell(unalloc.grossProfit), cell(pct(unalloc.gmPct)), kCell(unalloc.ownExpenses), kCell(unalloc.grossProfit + unalloc.ownExpenses)]);
    var cInc = act.income, cGP = act.grossProfit, cOv = act.expenses;
    rows.push(totalRow([cell("CTS", { align: "left" }), kCell(cInc), kCell(bud.income), varCell(cInc - bud.income), kCell(cGP), cell(pct(cInc ? cGP / cInc : null)), kCell(cOv), kCell(cGP + cOv)]));
    s.addTable(rows, tableOpts({ x: 0.55, y: 1.6, w: 6.1, colW: [1.2, 0.7, 0.7, 0.6, 0.85, 0.6, 0.75, 0.7] }));
    footnote(s, "Overheads are what is coded to the department; Admin carries the shared overheads. Contribution is gross profit less those overheads, and the CTS row is operating profit before other income. Push the Admin pool out on the P&L by Department page to see the split.", 4.2, 6.1);
    commentary(s, ["pnl-dept", "allocation"], ["Which departments beat budget and why", "Margin against each department's norm", "Overheads moving against plan"]);

    // 6 revenue and contribution by department
    s = frame("Financial performance", "Revenue and contribution by department");
    var dr = alloc.rows.filter(function (r) { return r.dept.isRevenue && r.code !== "ADMIN"; });
    s.addChart(P.ChartType.bar, [
      { name: "Revenue", labels: dr.map(function (r) { return r.dept.short; }), values: dr.map(function (r) { return kNum(r.income); }) },
      { name: "Contribution", labels: dr.map(function (r) { return r.dept.short; }), values: dr.map(function (r) { return kNum(r.grossProfit + r.ownExpenses); }) },
    ], chartOpts({ x: 0.55, y: 1.6, w: 6.1, h: 3.2, barDir: "col", barGapWidthPct: 60, chartColors: [C.navy, C.mint], showLegend: true, showValue: true, dataLabelPosition: "outEnd", valAxisLabelFormatCode: "#,##0", showTitle: true, title: mShort.replace("-", " ") + ", A$ thousands", titleFontSize: 9, titleFontFace: FONT, titleColor: C.grey }));
    footnote(s, "Contribution is gross profit less the department's own overheads, before any share of Admin.", 4.85, 6.1);
    commentary(s, ["rev-summary", "pnl-dept"], ["Where the mix of revenue is shifting", "Departments running below contribution expectations", "What changes next month"]);

    // 7 risk revenue profile
    s = frame("Financial performance", "Risk revenue profile");
    var lastFY = e.monthsOfFY(fy - 1), lines = [], contracted = 0, nonContract = 0;
    revDepts.forEach(function (d) {
      var mix = e.revenueMix([rm], d.code), income = deptIncome(d.code, [rm]), rec = mix.recurring, non = income - rec;
      var pyMix = e.revenueMix(lastFY, d.code), pyInc = deptIncome(d.code, lastFY), share = pyInc ? pyMix.recurring / pyInc : 0;
      var bd = budInc[d.code] || 0, bRec = Math.round(bd * share), bNon = bd - bRec;
      contracted += rec; nonContract += non;
      if (rec || bRec) lines.push({ label: d.short + ": Contract income", actual: rec, budget: bRec });
      if (non || bNon) lines.push({ label: d.short + ": Non contract income", actual: non, budget: bNon });
    });
    lines.forEach(function (l) { l.v = l.actual - l.budget; l.vp = l.budget ? l.v / Math.abs(l.budget) : null; l.risk = e.risk(l.actual, l.budget); });
    lines.sort(function (a, b) { return Math.abs(b.v) - Math.abs(a.v); });
    var flagged = lines.filter(function (l) { return l.risk === "High" || l.risk === "Medium"; }).length;
    var totM = contracted + nonContract;
    stat(s, 0.55, pct(totM ? contracted / totM : null, 0), "Contracted revenue, secured", C.teal);
    stat(s, 2.6, pct(totM ? nonContract / totM : null, 0), "Non-contract revenue, at risk", C.burgundy);
    stat(s, 4.65, String(flagged), "Revenue lines flagged Medium or High", C.ink, 2.0);
    rows = [hdr([{ label: "Revenue line vs budget, A$ thousands", align: "left" }, { label: "Actual" }, { label: "Budget" }, { label: "Var" }, { label: "Var %" }, { label: "Risk" }])];
    lines.slice(0, 8).forEach(function (l) { rows.push([cell(l.label, { align: "left" }), kCell(l.actual), kCell(l.budget), varCell(l.v), cell(pSigned(l.vp)), cell(l.risk, { color: l.risk === "High" ? C.bad : l.risk === "Medium" ? "B7850A" : C.grey, bold: l.risk === "High" })]); });
    s.addTable(rows, tableOpts({ x: 0.55, y: 2.5, w: 6.1, colW: [2.5, 0.75, 0.75, 0.7, 0.7, 0.7] }));
    footnote(s, "Contract income is fixed fee support, maintenance and licence accounts. Its budget share per department is last year's recurring share of that department's revenue. Risk follows the Controller Pack rule: the materiality floor first, then the percentage.", 4.6, 6.1);
    commentary(s, ["rev-risk", "rev-forecast"], ["How much of next quarter is already contracted", "Non-contract revenue: pipeline cover and conversion", "Flagged lines: cause and what changes"]);

    // 8 divider
    divider("02", "Customers and cash", "Top customers, receivables and days sales outstanding");

    // 9 top 10 customers
    var ytdRows = e.clientRows(ytd, pyYtd), monthRows = e.clientRows([rm]), byContact = {};
    monthRows.forEach(function (r) { byContact[r.contact] = r.amount; });
    var ytdTot = ytdA.income || 1, top10 = ytdRows.filter(function (r) { return r.amount > 0; }).slice(0, 10);
    s = frame("Customers and cash", "Top 10 customers");
    rows = [hdr([{ label: "Client, A$ thousands", align: "left" }, { label: mShort.replace("-", " ") }, { label: "YTD" }, { label: "Share of YTD" }, { label: "FY" + (fy - 1) + " YTD" }, { label: "YoY" }])];
    var t10 = { m: 0, y: 0, p: 0 };
    top10.forEach(function (r, i) {
      t10.m += byContact[r.contact] || 0; t10.y += r.amount; t10.p += r.prior;
      rows.push([cell((i + 1) + ".  " + (r.display || r.contact), { align: "left" }), kCell(byContact[r.contact] || 0), kCell(r.amount), cell(pct(r.amount / ytdTot, 0)), kCell(r.prior), cell(r.prior ? pSigned((r.amount - r.prior) / Math.abs(r.prior), 0) : "New", { color: r.prior && r.amount < r.prior ? C.bad : C.good })]);
    });
    rows.push(totalRow([cell("Top 10", { align: "left" }), kCell(t10.m), kCell(t10.y), cell(pct(t10.y / ytdTot, 0)), kCell(t10.p), cell(t10.p ? pSigned((t10.y - t10.p) / Math.abs(t10.p), 0) : "")]));
    s.addTable(rows, tableOpts({ x: 0.55, y: 1.6, w: 6.1, colW: [2.6, 0.7, 0.7, 0.75, 0.75, 0.6], rowH: 0.22, fontSize: 8.5 }));
    footnote(s, "Top 10 customers are " + pct(t10.y / ytdTot, 0) + " of " + ytdLabel + " revenue. Client names are the Xero contact, folded by the aliases on the Clients tab of the revenue schedule.", 4.6, 6.1);
    commentary(s, ["clients", "client-dept"], ["Clients growing year on year and what is driving it", "Clients in decline and the recovery plan", "New logos replacing lost revenue"]);

    // 10 receivables and DSO
    s = frame("Customers and cash", "Receivables and days sales outstanding");
    var aged = e.aged, arRows = aged ? (aged.receivables || []) : [];
    var buckets = { current: 0, lt1: 0, m1: 0, m2: 0, m3: 0, older: 0, total: 0 };
    arRows.forEach(function (r) { Object.keys(buckets).forEach(function (bk) { buckets[bk] += r[bk] || 0; }); });
    var last3 = e.actualMonths.filter(function (m) { return m <= rm; }).slice(-3);
    var rev3 = e.pnlActual(last3).totals.income, days3 = last3.reduce(function (a, m) { return a + ((e.monthIdx[m] || {}).days || 30); }, 0);
    var dso = rev3 ? buckets.total / (rev3 / days3) : null;
    var target = (e.cash && e.cash.settings && e.cash.settings.debtorDays) || 30;
    if (aged) {
      stat(s, 0.55, dso != null ? Math.round(dso) + " days" : "n/a", "DSO, target " + target + " days", dso != null && dso > target ? C.burgundy : C.teal);
      stat(s, 2.6, dollars(buckets.total), "Trade receivables at month end", C.ink);
      stat(s, 4.65, dollars(buckets.m2 + buckets.m3 + buckets.older), "Over 60 days", buckets.m2 + buckets.m3 + buckets.older > 0 ? C.burgundy : C.teal, 2.0);
      var bk = [["Current", buckets.current], ["1 to 30 days", buckets.lt1], ["31 to 60 days", buckets.m1], ["61 to 90 days", buckets.m2], ["Over 90 days", buckets.m3 + buckets.older]];
      caption(s, "Aged receivables at " + mShort.replace("-", " ") + ", by due date", 2.5);
      s.addChart(P.ChartType.bar, [{ name: "Owed", labels: bk.map(function (x) { return x[0]; }), values: bk.map(function (x) { return kNum(x[1]); }) }],
        chartOpts({ x: 0.55, y: 2.72, w: 3.0, h: 1.95, barDir: "col", chartColors: [C.navy], showLegend: false, showValue: true, dataLabelPosition: "outEnd", valAxisLabelFormatCode: "#,##0", valAxisHidden: true }));
      rows = [hdr([{ label: "Bucket", align: "left" }, { label: "A$k" }, { label: "%" }])];
      bk.forEach(function (x) { rows.push([cell(x[0], { align: "left" }), kCell(x[1]), cell(pct(buckets.total ? x[1] / buckets.total : null, 0))]); });
      rows.push(totalRow([cell("Total", { align: "left" }), kCell(buckets.total), cell("100%")]));
      s.addTable(rows, tableOpts({ x: 3.75, y: 2.72, w: 2.9, colW: [1.5, 0.7, 0.7] }));
      footnote(s, "DSO = receivables / (last three months' revenue / " + days3 + " days). The six month DSO trend builds as each month end is archived.", 4.75, 6.1);
    } else {
      s.addText("Fill 11 Xero Aged Receivables.xlsx and run Build to see receivables, the ageing and days sales outstanding here.", { x: 0.55, y: 1.6, w: 6.1, h: 0.8, fontFace: FONT, fontSize: 11, color: C.grey, isTextBox: true, margin: 0 });
    }
    commentary(s, ["cash", "rev-risk"], ["Largest overdue accounts and collection status", "Why DSO moved this month", "Cash actions before next month end"]);

    // 11 top customers year on year
    s = frame("Customers and cash", "Top customers, FY" + fy + " vs FY" + (fy - 1) + " year to date");
    s.addChart(P.ChartType.bar, [
      { name: "FY" + fy + " YTD", labels: top10.map(function (r) { return r.display || r.contact; }), values: top10.map(function (r) { return kNum(r.amount); }) },
      { name: "FY" + (fy - 1) + " YTD", labels: top10.map(function (r) { return r.display || r.contact; }), values: top10.map(function (r) { return kNum(r.prior); }) },
    ], chartOpts({ x: 0.55, y: 1.6, w: 6.1, h: 3.15, barDir: "bar", barGapWidthPct: 50, chartColors: [C.navy, C.greyLight], showLegend: true, valAxisLabelFormatCode: "#,##0", catAxisLabelFontSize: 7.5 }));
    footnote(s, "A$ thousands, same period last year. Customers ranked by current year to date.", 4.85, 6.1);
    commentary(s, ["clients"], ["Customers growing year on year and what is driving it", "Customers in decline and the recovery plan", "New logos replacing lost revenue"]);

    // 12 divider
    divider("03", "People and utilisation", "Chargeable hours, utilisation and the profit they produce");

    // 13 utilisation by department
    var util = e.utilFor([rm]), prevKey = e.months.filter(function (m) { return m.key < rm; }).slice(-1)[0], prevUtil = prevKey ? e.utilFor([prevKey.key]) : null;
    var prevBy = {}; if (prevUtil) prevUtil.rows.forEach(function (r) { prevBy[r.dept.code] = r.util; });
    s = frame("People and utilisation", "Utilisation by department, " + mShort.replace("-", " "));
    rows = [hdr([{ label: "Hours", align: "left" }, { label: "FTE" }, { label: "Chargeable" }, { label: "Non-chargeable" }, { label: "Leave" }, { label: "Util %" }, { label: "Prior" }, { label: "vs target" }])];
    var uRows = util.rows.filter(function (r) { return r.dept.code !== "UNALLOCATED" || r.worked; });
    uRows.forEach(function (r) {
      var gap = r.target != null && r.util != null ? r.util - r.target : null;
      rows.push([cell(r.dept.short, { align: "left", bold: true }), cell(r.fte != null ? r.fte.toFixed(0) : ""), cell(F().hours(r.chargeable)), cell(F().hours(r.nonChargeable)), cell(F().hours(r.leave)),
                 cell(pct(r.util), { bold: true, color: gap != null && gap < -0.05 ? C.bad : C.ink }), cell(pct(prevBy[r.dept.code])), cell(gap == null ? "" : pSigned(gap), { color: gap != null && gap < 0 ? C.bad : C.good })]);
    });
    rows.push(totalRow([cell("CTS", { align: "left" }), cell(util.total.fte != null ? util.total.fte.toFixed(0) : ""), cell(F().hours(util.total.chargeable)), cell(F().hours(util.total.nonChargeable)), cell(F().hours(util.total.leave)), cell(pct(util.total.util)), cell(prevUtil ? pct(prevUtil.total.util) : ""), cell("")]));
    s.addTable(rows, tableOpts({ x: 0.55, y: 1.6, w: 6.1, colW: [1.1, 0.5, 0.85, 1.05, 0.65, 0.65, 0.6, 0.7] }));
    footnote(s, "Utilisation is chargeable hours over hours worked, leave excluded, from the Employment Hero pay runs. Prior is " + (prevKey ? prevKey.label : "the month before") + ". Target " + pct((revDepts[0] || {}).utilTarget || 0.65, 0) + " unless set otherwise per department.", 4.3, 6.1);
    commentary(s, ["util", "profit-fte"], ["Departments under target and the reason: bench time, leave, training", "Overtime or casual hours masking low utilisation", "Hiring or roster changes planned"]);

    // 14 profit based on utilisation
    var staff = e.staffFor([rm]), costBy = {};
    staff.rows.forEach(function (r) { costBy[r.dept] = (costBy[r.dept] || 0) + r.cost; });
    s = frame("People and utilisation", "Profit based on utilisation");
    rows = [hdr([{ label: "A$ thousands", align: "left" }, { label: "Util %" }, { label: "Charge hrs" }, { label: "Rate" }, { label: "Labour rev" }, { label: "Labour cost" }, { label: "Labour profit" }, { label: "Per hr" }, { label: "+5 pts" }])];
    var pt = { chg: 0, rev: 0, cost: 0, plus: 0, worked: 0 };
    uRows.filter(function (r) { return r.dept.isRevenue && r.dept.code !== "ADMIN"; }).forEach(function (r) {
      var labourRev = e.revenueMix([rm], r.dept.code).labour, rate = r.chargeable ? labourRev / r.chargeable : 0, cost = costBy[r.dept.code] || 0, profit = labourRev - cost;
      var plus = Math.round(0.05 * r.worked * rate);
      pt.chg += r.chargeable; pt.rev += labourRev; pt.cost += cost; pt.plus += plus; pt.worked += r.worked;
      rows.push([cell(r.dept.short, { align: "left", bold: true }), cell(pct(r.util)), cell(F().hours(r.chargeable)), cell(rate ? perHr(rate) : ""), kCell(labourRev), kCell(-cost), kCell(profit, true), cell(r.chargeable ? perHr(profit / r.chargeable) : "", { color: profit < 0 ? C.bad : C.ink }), varCell(plus)]);
    });
    var cRate = pt.chg ? pt.rev / pt.chg : 0;
    rows.push(totalRow([cell("CTS", { align: "left" }), cell(pct(pt.worked ? pt.chg / pt.worked : null)), cell(F().hours(pt.chg)), cell(cRate ? perHr(cRate) : ""), kCell(pt.rev), kCell(-pt.cost), kCell(pt.rev - pt.cost), cell(pt.chg ? perHr((pt.rev - pt.cost) / pt.chg) : ""), varCell(pt.plus)]));
    s.addTable(rows, tableOpts({ x: 0.55, y: 1.6, w: 6.1, colW: [1.0, 0.55, 0.7, 0.5, 0.75, 0.8, 0.8, 0.5, 0.5], fontSize: 8.5 }));
    s.addShape(P.ShapeType.rect, { x: 0.55, y: 3.75, w: 6.1, h: 0.7, fill: { color: C.comment }, line: { color: C.comment } });
    s.addText("+" + dollars(pt.plus).replace("$", "$") + " a month", { x: 0.7, y: 3.82, w: 2.4, h: 0.36, fontFace: FONT, fontSize: 18, bold: true, color: C.teal, isTextBox: true, margin: 0 });
    s.addText("Five more points of utilisation at current charge rates, before any extra cost", { x: 3.1, y: 3.85, w: 3.45, h: 0.5, fontFace: FONT, fontSize: 9.5, color: C.grey, isTextBox: true, margin: 0, valign: "middle" });
    footnote(s, "Labour revenue is the labour accounts (consulting, production, installation, support labour, event management) coded to the department. Rate is labour revenue over chargeable hours. Labour cost is the department's gross pay and super from the pay runs.", 4.6, 6.1);
    commentary(s, ["staff-profit", "profit-fte"], ["Where an extra chargeable hour is worth most", "Rates against cost per hour", "Who is below recovery and why"]);

    // 15 divider
    divider("04", "Summary and actions", "The commentary from each section, with an owner and a date");

    // 16 summary and actions
    s = frame("Summary and actions", "Commentary and actions from this review");
    var areas = [["P&L", ["home", "pnl", "bva", "pnl-dept", "allocation", "rev-summary"]], ["Risk revenue", ["rev-risk", "rev-forecast", "rev-schedule", "pipeline", "forecast"]], ["Customers and DSO", ["clients", "client-dept", "cash"]], ["Utilisation", ["util", "profit-fte", "staff-profit"]]];
    rows = [hdr([{ label: "Area", align: "left" }, { label: "Commentary and agreed action", align: "left" }, { label: "Owner", align: "left" }, { label: "Due", align: "left" }])];
    areas.forEach(function (a) {
      var items = [];
      a[1].forEach(function (pg) { (e.commentaryFor ? e.commentaryFor(pg, rm, { all: true }) : []).forEach(function (c) { items.push(c); }); });
      if (items.length) items.slice(0, 3).forEach(function (c, i) { rows.push([cell(i ? "" : a[0], { align: "left", bold: true }), cell(c.text.replace(/\s+/g, " "), { align: "left" }), cell(c.author || "", { align: "left" }), cell("", { align: "left" })]); });
      else rows.push([cell(a[0], { align: "left", bold: true }), cell("Type the commentary agreed in the meeting", { align: "left", color: C.greyLight, italic: true }), cell("Owner", { align: "left", color: C.greyLight, italic: true }), cell("Date", { align: "left", color: C.greyLight, italic: true })]);
    });
    s.addTable(rows, tableOpts({ x: 0.55, y: 1.6, w: 8.9, colW: [1.3, 5.5, 1.2, 0.9], rowH: 0.3, fontSize: 9 }));
    footnote(s, "Rows come from the commentary typed on the portal pages for " + mLong + ". Owner is who wrote it; the due date is agreed in the meeting.", 4.85, 8.9);

    // 17 thank you
    s = P.addSlide(); dark(s); _n++;
    if (b && b.logo && b.logo.dark) s.addImage({ data: b.logo.dark, x: 0.55, y: 0.5, w: 1.6, h: 1.09 });
    s.addText("Thank you", { x: 0.55, y: 2.1, w: 8.9, h: 0.9, fontFace: FONT, fontSize: 40, bold: true, color: C.white, isTextBox: true, margin: 0 });
    var url = (e.cfg.ORG && e.cfg.ORG.portalUrl) || "";
    s.addText(url ? [{ text: "Every figure in this deck is live in the ", options: {} }, { text: "CTS Business Intelligence Portal", options: { hyperlink: { url: url }, color: C.mint, bold: true } }, { text: ", with its workings, the forecast and the commentary." }]
                  : "Every figure in this deck is live in the CTS Business Intelligence Portal, with its workings, the forecast and the commentary. Open it from the synced CTS Business Portal folder.",
              { x: 0.55, y: 3.05, w: 8.0, h: 0.9, fontFace: FONT, fontSize: 13, color: C.greyLight, isTextBox: true, margin: 0, valign: "top" });
    s.addText("Corporate Technology Services  |  Seamless AV  |  Confidential", { x: 0.55, y: 4.85, w: 8, h: 0.3, fontFace: FONT, fontSize: 9, color: C.greyLight, isTextBox: true, margin: 0 });
    return P;
  }
  D.bytes = function () { return D.build().write({ outputType: "arraybuffer" }); };
  D.download = function () { return D.build().writeFile({ fileName: D.fileName() }); };

  /** The note that travels with the deck: short, branded, pointing at the portal. */
  D.messages = function () {
    var M = CTS.email, e = E(), rm = e.reportingMonth(), mLabel = (e.monthIdx[rm] || {}).long || rm;
    var act = e.pnlActual([rm]).totals, bud = e.pnlBudget([rm]).totals;
    var file = D.fileName();
    return M.recipients().filter(function (r) { return +r.tier !== 2; }).map(function (r) {
      var subject = "CTS Monthly Business Review, " + mLabel + " (deck attached)";
      var body = '<p style="font-family:Segoe UI,Calibri,Arial,sans-serif;font-size:13.5px;color:#1E2530;line-height:1.5">Hi ' + r.name.split(" ")[0] + ",<br><br>The Monthly Business Review for " + mLabel + " is attached: revenue " + F().dollars(act.income) + " against " + F().dollars(bud.income) + " budgeted, net profit " + F().money(act.netProfit) + " against " + F().money(bud.netProfit) + ". The deck covers the P&amp;L, the departments, the risk revenue profile, top customers, receivables and DSO, utilisation and the agreed actions.</p>";
      var html = M.shellFor ? M.shellFor(subject, "For " + r.name + " (" + M.TIERS[r.tier] + "). Built " + new Date().toISOString().slice(0, 10) + ".", body) : body;
      return { to: r.email || "", name: r.name, tier: r.tier, subject: subject, html: html, text: "The Monthly Business Review for " + mLabel + " is attached. See the full result in the portal.", attachments: [file], kind: "deck" };
    });
  };
})();
