/* CTS Business Intelligence Portal - the month end deck
 * ---------------------------------------------------------------------------
 * Builds the month end P&L pack as a PowerPoint file, in the CTS livery,
 * from the same engine the pages read. One click on Distribution makes it;
 * a second writes it into the outbox with a note per executive recipient so
 * the flow sends it. Rendering is PptxGenJS (vendor/pptxgen.bundle.js, MIT).
 */
(function () {
  "use strict";
  var CTS = (window.CTS = window.CTS || {});
  var D = (CTS.deck = {});
  var E = function () { return CTS.engine; };
  var F = function () { return CTS.fmt; };

  var C = { black: "121820", white: "FFFFFF", grey: "5F6670", greyLight: "C9CED6", panel: "F6F7F9", line: "DDE1E7",
            navy: "122B82", mint: "3FD0C9", lavender: "6D71FF", rose: "F33844", teal: "005358", ink: "1E2530",
            good: "0A6B2C", bad: "A0161F" };
  var DEPT = { ONSITE: "F33A45", PRODUCTION: "0E9E99", VIDEO: "B7850A", INTEGRATION: "2E4DA6", CONSULTING: "6C70FF", ADMIN: "018301", UNALLOCATED: "7B8391" };
  var FONT = "Segoe UI";
  var W = 13.333, H = 7.5;

  D.supported = function () { return typeof window.PptxGenJS === "function"; };
  D.fileName = function () {
    var e = E(), rm = e.reportingMonth();
    return "CTS Month End " + ((e.monthIdx[rm] || {}).long || rm).replace(/\s+/g, " ") + ".pptx";
  };

  function dollars(c) { return F().dollars(c); }
  function money(c) { return F().money(c); }
  function pct(x, dp) { return F().pct(x, dp); }
  function cents(c) { return Math.round(c / 100); }

  /* ---- slide furniture -------------------------------------------------- */
  function frame(pptx, title, eyebrow, accent) {
    var s = pptx.addSlide();
    s.background = { color: C.white };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.08, fill: { color: accent || C.mint }, line: { color: accent || C.mint } });
    s.addText((eyebrow || "CTS  |  Seamless AV").toUpperCase(), { x: 0.6, y: 0.3, w: 8, h: 0.3, fontFace: FONT, fontSize: 10, bold: true, color: C.navy, charSpacing: 3 });
    s.addText(title, { x: 0.6, y: 0.55, w: 10, h: 0.7, fontFace: FONT, fontSize: 26, bold: true, color: C.black });
    var b = window.CTS_BRAND;
    if (b && b.logo && b.logo.light) s.addImage({ data: b.logo.light, x: W - 1.9, y: 0.32, w: 1.3, h: 0.88 });
    s.addText("Corporate Technology Services  |  Seamless AV  |  Confidential", { x: 0.6, y: H - 0.42, w: 8, h: 0.3, fontFace: FONT, fontSize: 9, color: C.grey });
    s.addText("CTS Business Intelligence Portal", { x: W - 4.6, y: H - 0.42, w: 4, h: 0.3, fontFace: FONT, fontSize: 9, color: C.grey, align: "right" });
    return s;
  }
  function note(s, text, y) {
    s.addText(text, { x: 0.6, y: y == null ? H - 0.95 : y, w: W - 1.2, h: 0.45, fontFace: FONT, fontSize: 10.5, color: C.grey, valign: "top" });
  }
  function tiles(s, items, y) {
    var n = items.length, gap = 0.18, w = (W - 1.2 - gap * (n - 1)) / n;
    items.forEach(function (t, i) {
      var x = 0.6 + i * (w + gap), col = t.tone === "good" ? C.good : t.tone === "bad" ? C.bad : C.black;
      s.addShape(pptx().ShapeType.rect, { x: x, y: y, w: w, h: 1.3, fill: { color: C.panel }, line: { color: C.line, width: 0.5 } });
      s.addShape(pptx().ShapeType.rect, { x: x, y: y, w: w, h: 0.07, fill: { color: t.accent || C.navy }, line: { color: t.accent || C.navy } });
      s.addText(t.label.toUpperCase(), { x: x + 0.15, y: y + 0.15, w: w - 0.3, h: 0.25, fontFace: FONT, fontSize: 9, bold: true, color: C.grey, charSpacing: 2 });
      s.addText(t.value, { x: x + 0.15, y: y + 0.42, w: w - 0.3, h: 0.5, fontFace: FONT, fontSize: 22, bold: true, color: col });
      if (t.sub) s.addText(t.sub, { x: x + 0.15, y: y + 0.92, w: w - 0.3, h: 0.3, fontFace: FONT, fontSize: 9.5, color: C.grey });
    });
    return y + 1.3;
  }
  /** The commentary for a page, on its slide: published and the finance
   *  head's own drafts, since it is the finance head who builds the deck. */
  function commentBox(s, pageIds, y, hgt) {
    var e = E(), rm = e.reportingMonth(), items = [];
    pageIds.forEach(function (pg) { (e.commentaryFor ? e.commentaryFor(pg, rm, { all: true }) : []).forEach(function (c) { items.push(c); }); });
    if (!items.length) return;
    var yy = y == null ? 5.55 : y, hh = hgt || 1.0;
    s.addShape(pptx().ShapeType.rect, { x: 0.6, y: yy, w: W - 1.2, h: hh, fill: { color: "EEEEFF" }, line: { color: "EEEEFF" } });
    s.addShape(pptx().ShapeType.rect, { x: 0.6, y: yy, w: 0.06, h: hh, fill: { color: C.lavender }, line: { color: C.lavender } });
    s.addText([{ text: "COMMENTARY  ", options: { bold: true, color: C.navy, fontSize: 9, charSpacing: 2 } }].concat(items.map(function (c, i) {
      return { text: (i ? "  " : "") + c.text.replace(/\s+/g, " ") + (c.dept ? " (" + ((e.deptOf[c.dept] || {}).short || c.dept) + ")" : ""), options: { color: C.ink, fontSize: 10 } };
    })), { x: 0.75, y: yy + 0.05, w: W - 1.5, h: hh - 0.1, fontFace: FONT, valign: "top" });
  }
  var _pptx = null;
  function pptx() { return _pptx; }
  function tableOpts(extra) {
    return Object.assign({ fontFace: FONT, fontSize: 10, color: C.ink, border: { type: "solid", pt: 0.5, color: C.line }, autoPage: false, rowH: 0.3 }, extra || {});
  }
  function hdr(cols) { return cols.map(function (c) { return { text: c.label, options: { bold: true, color: C.grey, fill: { color: C.panel }, fontSize: 9, align: c.align || "right" } }; }); }
  function cell(v, o) { return { text: v == null ? "" : String(v), options: Object.assign({ align: "right" }, o || {}) }; }
  function moneyCell(c, bold) { return cell(money(c), { color: c < 0 ? C.bad : C.ink, bold: !!bold }); }

  /* ---- the deck ----------------------------------------------------------- */
  D.build = function () { return E().withoutView ? E().withoutView(buildInner) : buildInner(); };
  function buildInner() {
    if (!D.supported()) throw new Error("PptxGenJS did not load (vendor/pptxgen.bundle.js).");
    var e = E(), p = CTS.period(), rm = e.reportingMonth(), mLabel = (e.monthIdx[rm] || {}).long || rm;
    var fy = e.currentFY(), fyKeys = e.monthsOfFY(fy), ytd = e.ytd();
    var P = new window.PptxGenJS(); _pptx = P;
    P.layout = "LAYOUT_WIDE"; P.author = "CTS Business Intelligence Portal"; P.company = "Corporate Technology Services"; P.title = "CTS month end result, " + mLabel;
    var b = window.CTS_BRAND;

    // 1 cover
    var s = P.addSlide(); s.background = { color: C.black };
    s.addShape(P.ShapeType.ellipse, { x: 8.6, y: -2.2, w: 5.4, h: 5.4, fill: { color: C.mint, transparency: 15 }, line: { color: C.mint, transparency: 100 } });
    s.addShape(P.ShapeType.ellipse, { x: 10.4, y: 0.6, w: 4.8, h: 4.8, fill: { color: C.lavender, transparency: 25 }, line: { color: C.lavender, transparency: 100 } });
    s.addShape(P.ShapeType.ellipse, { x: 8.9, y: 3.2, w: 4.8, h: 4.8, fill: { color: C.rose, transparency: 30 }, line: { color: C.rose, transparency: 100 } });
    if (b && b.logo && b.logo.dark) s.addImage({ data: b.logo.dark, x: 0.75, y: 0.6, w: 2.2, h: 1.5 });
    s.addText("MONTH END RESULT", { x: 0.75, y: 3.3, w: 7, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: C.mint, charSpacing: 4 });
    s.addText(mLabel, { x: 0.75, y: 3.7, w: 8, h: 1.1, fontFace: FONT, fontSize: 54, bold: true, color: C.white });
    s.addText("FY" + fy + " to " + ((e.monthIdx[rm] || {}).label || rm) + ", with the forecast to year end", { x: 0.75, y: 4.8, w: 8, h: 0.5, fontFace: FONT, fontSize: 18, color: C.greyLight });
    s.addText([{ text: "Finance", options: { bold: true, color: C.white, breakLine: true } }, { text: "Corporate Technology Services" }], { x: 0.75, y: 6.3, w: 6, h: 0.7, fontFace: FONT, fontSize: 12, color: C.greyLight });
    s.addText("Built " + new Date().toISOString().slice(0, 10), { x: W - 4.5, y: 6.5, w: 3.8, h: 0.4, fontFace: FONT, fontSize: 12, color: C.greyLight, align: "right" });

    // 2 headline
    var act = e.pnlActual([rm]).totals, bud = e.pnlBudget([rm]).totals, ytdA = e.pnlActual(ytd).totals, ytdB = e.pnlBudget(ytd).totals;
    var util = e.utilFor(ytd), cash = e.cashEnabled() && e.forecastEnabled() ? e.cashflow() : null;
    var land = e.forecastEnabled() ? e.pnlBlend(fyKeys).totals : null, fb = e.pnlBudget(fyKeys).totals;
    s = frame(P, "The month at a glance", "Headline, " + mLabel, C.mint);
    var y = tiles(s, [
      { label: "Revenue", value: dollars(act.income), sub: (act.income - bud.income >= 0 ? "+" : "−") + dollars(Math.abs(act.income - bud.income)) + " vs budget", tone: act.income >= bud.income ? "good" : "bad", accent: C.navy },
      { label: "Gross profit", value: dollars(act.grossProfit), sub: pct(act.income ? act.grossProfit / act.income : null) + " margin, budget " + pct(bud.income ? bud.grossProfit / bud.income : null), accent: C.mint },
      { label: "Net profit", value: dollars(act.netProfit), sub: "budget " + dollars(bud.netProfit), tone: act.netProfit >= bud.netProfit ? "good" : "bad", accent: C.rose },
      { label: "Year to date net profit", value: dollars(ytdA.netProfit), sub: "budget " + dollars(ytdB.netProfit), tone: ytdA.netProfit >= ytdB.netProfit ? "good" : "bad", accent: C.lavender },
      { label: "Utilisation YTD", value: pct(util.total.util), sub: util.total.fte ? util.total.fte.toFixed(1) + " FTE" : "", accent: C.teal },
    ], 1.5);
    var points = CTS.pages && CTS.pages.story && CTS.pages.story.points ? CTS.pages.story.points() : [];
    var hasHome = e.commentaryFor && e.commentaryFor("home", rm, { all: true }).concat(e.commentaryFor("story", rm, { all: true })).length;
    if (hasHome) commentBox(s, ["home", "story"], y + 0.2, 1.0);
    if (points.length) {
      s.addText(points.slice(0, 6).map(function (pt, i) { return [{ text: (i + 1) + ".  " + pt.head, options: { bold: true, color: C.black, breakLine: true, fontSize: 12.5 } }, { text: "      " + pt.body, options: { color: C.grey, fontSize: 10, breakLine: true, paraSpaceAfter: 5 } }]; }).reduce(function (a, x) { return a.concat(x); }, []),
                { x: 0.6, y: y + (hasHome ? 1.35 : 0.3), w: W - 1.2, h: H - y - (hasHome ? 2.4 : 1.4), fontFace: FONT, valign: "top" });
    }

    // 3 company P&L
    s = frame(P, "Profit and loss", "Company, " + mLabel + " and FY" + fy + " to date", C.navy);
    var lines = [["Income", "income"], ["Cost of sales", "cos"], ["Gross profit", "grossProfit"], ["Overheads", "expenses"], ["Other income", "otherIncome"], ["Net profit", "netProfit"]];
    var rows = [hdr([{ label: "", align: "left" }, { label: mLabel }, { label: "Budget" }, { label: "Variance" }, { label: "Year to date" }, { label: "Budget" }, { label: "Variance" }].concat(land ? [{ label: "FY" + fy + " landing" }, { label: "FY budget" }] : []))];
    lines.forEach(function (l) {
      var k = l[1], bold = /profit/i.test(l[0]);
      rows.push([cell(l[0], { align: "left", bold: bold }), moneyCell(act[k], bold), moneyCell(bud[k]), moneyCell(act[k] - bud[k]), moneyCell(ytdA[k], bold), moneyCell(ytdB[k]), moneyCell(ytdA[k] - ytdB[k])].concat(land ? [moneyCell(land[k], bold), moneyCell(fb[k])] : []));
    });
    s.addTable(rows, tableOpts({ x: 0.6, y: 1.45, w: W - 1.2, colW: land ? [2.2, 1.25, 1.25, 1.25, 1.35, 1.25, 1.25, 1.35, 1.2] : [2.6, 1.6, 1.6, 1.6, 1.6, 1.6, 1.6] }));
    var full = e.pnlBlend(fyKeys), labels = fyKeys.map(function (k) { return e.monthIdx[k].label; });
    s.addChart(P.ChartType.bar, [
      { name: "Revenue", labels: labels, values: full.income.map(cents) },
      { name: "Gross profit", labels: labels, values: full.grossProfit.map(cents) },
      { name: "Net profit", labels: labels, values: full.netProfit.map(cents) },
    ], { x: 0.6, y: 3.75, w: W - 1.2, h: 2.75, barDir: "col", barGapWidthPct: 60, chartColors: [C.navy, "0E9E99", "B7850A"], showLegend: true, legendPos: "b", legendFontSize: 9, legendFontFace: FONT,
         catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, valAxisLabelFormatCode: "$#,##0", catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, valGridLine: { color: C.line, size: 0.5 }, catGridLine: { style: "none" } });
    commentBox(s, ["pnl", "bva"], 6.5, 0.55);
    note(s, "Costs shown as the ledger holds them, so gross profit is income plus cost of sales. Months after " + ((e.monthIdx[rm] || {}).label || rm) + " are " + (e.forecastEnabled() ? "forecast" : "budget") + ".", 3.4);

    // 4 departments after the split
    var alloc = e.allocate(p.keys), dRows = alloc.rows.filter(function (r) { return r.dept.isRevenue && (r.income || r.expenses); });
    var budD = e.budgetByDept(p.keys, "income");
    s = frame(P, "Departments after the overhead split", p.label, C.lavender);
    rows = [hdr([{ label: "Department", align: "left" }, { label: "Revenue" }, { label: "Budget" }, { label: "GM %" }, { label: "Norm" }, { label: "Own overhead" }, { label: "Share of Admin" }, { label: "Net profit" }])];
    dRows.forEach(function (r) { rows.push([cell(r.dept.short, { align: "left", bold: true }), moneyCell(r.income), moneyCell(budD[r.code] || 0), cell(pct(r.gmPct)), cell(r.dept.gmNorm != null ? pct(r.dept.gmNorm, 0) : "n/a"), moneyCell(r.ownExpenses), moneyCell(r.allocated), moneyCell(r.netProfit, true)]); });
    s.addTable(rows, tableOpts({ x: 0.6, y: 1.45, w: 7.4, colW: [1.35, 1.0, 1.0, 0.7, 0.6, 0.95, 0.95, 0.85], fontSize: 9.5 }));
    s.addChart(P.ChartType.bar, [
      { name: "Revenue", labels: dRows.map(function (r) { return r.dept.short; }), values: dRows.map(function (r) { return cents(r.income); }) },
      { name: "Net profit after split", labels: dRows.map(function (r) { return r.dept.short; }), values: dRows.map(function (r) { return cents(r.netProfit); }) },
    ], { x: 8.2, y: 1.45, w: 4.6, h: 4.2, barDir: "col", chartColors: [C.navy, "B7850A"], showLegend: true, legendPos: "b", legendFontSize: 9, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, valAxisLabelFormatCode: "$#,##0", valGridLine: { color: C.line, size: 0.5 } });
    var src = e.allocationIsSourced();
    commentBox(s, ["pnl-dept"], 5.8, 0.75);
    note(s, (src.ok ? "The Admin pool is pushed out on the confirmed split bases. " : "Allocated figures rest on placeholder split bases (" + src.placeholders.join(", ") + "). ") + "Own overhead and the share of Admin are kept apart so nothing is counted twice.");

    // 5 revenue by department by month
    s = frame(P, "Revenue by department", "FY" + fy + ", " + e.blendLabel(), C.rose);
    var depts = e.revenueDepts.filter(function (d) { return d.code !== "ADMIN"; });
    s.addChart(P.ChartType.bar, depts.map(function (d) {
      return { name: d.short, labels: labels, values: fyKeys.map(function (k) {
        var srcK = e.blendSource(k);
        var v = srcK === "actual" ? e.sumMonths((e.idx.deptCatMonth[d.code] || {}).income, [k]) : srcK === "forecast" ? (e.forecastByDept([k], "income")[d.code] || 0) : (e.budgetByDept([k], "income")[d.code] || 0);
        return cents(v); }) };
    }), { x: 0.6, y: 1.45, w: 7.6, h: 4.6, barDir: "col", barGrouping: "stacked", chartColors: depts.map(function (d) { return DEPT[d.code] || C.grey; }), showLegend: true, legendPos: "b", legendFontSize: 9, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, valAxisLabelFormatCode: "$#,##0", valGridLine: { color: C.line, size: 0.5 } });
    var prior = e.priorYearMonths(p.keys);
    rows = [hdr([{ label: "Department", align: "left" }, { label: p.id === "month" ? "Month" : "Period" }, { label: "Last year" }, { label: "Budget" }, { label: "vs budget" }])];
    depts.forEach(function (d) { var now = e.sumMonths((e.idx.deptCatMonth[d.code] || {}).income, p.keys), was = e.sumMonths((e.idx.deptCatMonth[d.code] || {}).income, prior), bd = budD[d.code] || 0; rows.push([cell(d.short, { align: "left", bold: true }), moneyCell(now), moneyCell(was), moneyCell(bd), moneyCell(now - bd)]); });
    s.addTable(rows, tableOpts({ x: 8.4, y: 1.45, w: 4.4, colW: [1.2, 0.85, 0.85, 0.8, 0.7], fontSize: 9 }));
    commentBox(s, ["rev-summary", "rev-schedule", "rev-forecast"], 6.2, 0.7);

    // 6 top clients
    var clients = e.clientRows(p.keys, prior).slice(0, 12);
    s = frame(P, "Top clients", p.label, C.teal);
    rows = [hdr([{ label: "Client", align: "left" }, { label: "Revenue" }, { label: "Same period last year" }, { label: "Change" }, { label: "Share" }])];
    var totInc = act.income || 1; var perTot = e.pnlActual(p.keys).totals.income || 1;
    clients.forEach(function (c) { rows.push([cell(c.display || c.contact, { align: "left" }), moneyCell(c.amount), moneyCell(c.prior || 0), moneyCell(c.delta), cell(pct(c.amount / perTot))]); });
    s.addTable(rows, tableOpts({ x: 0.6, y: 1.45, w: W - 1.2, colW: [5.1, 1.7, 2.2, 1.6, 1.5], fontSize: 10 }));
    commentBox(s, ["clients", "client-dept"], 5.6, 0.7);
    note(s, "Revenue lines by contact off the ledger. A client that appears here and not in the revenue schedule is new business.");

    // 7 utilisation
    s = frame(P, "Utilisation", p.label + ", chargeable over worked hours, leave excluded", C.mint);
    rows = [hdr([{ label: "Department", align: "left" }, { label: "Chargeable" }, { label: "Non chargeable" }, { label: "Leave" }, { label: "Utilisation" }, { label: "Target" }, { label: "FTE" }])];
    util.rows.forEach(function (r) { rows.push([cell(r.dept.short, { align: "left", bold: true }), cell(F().hours(r.chargeable)), cell(F().hours(r.nonChargeable)), cell(F().hours(r.leave)), cell(pct(r.util), { color: r.target != null && r.util != null && r.util < r.target - 0.05 ? C.bad : C.ink, bold: true }), cell(r.target != null ? pct(r.target, 0) : "n/a"), cell(r.fte != null ? r.fte.toFixed(1) : "")]); });
    rows.push([cell("Total", { align: "left", bold: true }), cell(F().hours(util.total.chargeable), { bold: true }), cell(F().hours(util.total.nonChargeable)), cell(F().hours(util.total.leave)), cell(pct(util.total.util), { bold: true }), cell(""), cell(util.total.fte != null ? util.total.fte.toFixed(1) : "")]);
    s.addTable(rows, tableOpts({ x: 0.6, y: 1.45, w: 7.4, colW: [1.5, 1.0, 1.15, 0.85, 1.0, 0.9, 0.7], fontSize: 10 }));
    s.addChart(P.ChartType.bar, [{ name: "Utilisation", labels: util.rows.map(function (r) { return r.dept.short; }), values: util.rows.map(function (r) { return Math.round((r.util || 0) * 1000) / 10; }) },
                                 { name: "Target", labels: util.rows.map(function (r) { return r.dept.short; }), values: util.rows.map(function (r) { return Math.round((r.target || 0) * 1000) / 10; }) }],
               { x: 8.2, y: 1.45, w: 4.6, h: 4.2, barDir: "col", chartColors: ["0E9E99", C.greyLight], showLegend: true, legendPos: "b", legendFontSize: 9, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, valAxisLabelFormatCode: "0\"%\"", valAxisMaxVal: 100, valGridLine: { color: C.line, size: 0.5 } });
    commentBox(s, ["util", "profit-fte", "staff-profit"], 5.8, 0.75);
    note(s, "Hours from the Employment Hero pay runs. Read a small negative before flagging it: a team whose hours are almost all chargeable usually just took leave.");

    // 8 forecast
    if (land) {
      var fcKeys = fyKeys.filter(function (k) { return e.blendSource(k) === "forecast"; }), actKeys = fyKeys.filter(function (k) { return e.blendSource(k) === "actual"; });
      var aT = e.pnlActual(actKeys).totals, fT = e.pnlForecast(fcKeys).totals;
      s = frame(P, "Full year landing", "FY" + fy + ", actual to " + ((e.monthIdx[rm] || {}).label || rm) + " then forecast", C.lavender);
      rows = [hdr([{ label: "", align: "left" }, { label: "Actual to date" }, { label: "Forecast to year end" }, { label: "Full year" }, { label: "Budget" }, { label: "Variance" }, { label: "Risk" }])];
      lines.forEach(function (l) { var k = l[1], bold = /profit/i.test(l[0]); rows.push([cell(l[0], { align: "left", bold: bold }), moneyCell(aT[k]), moneyCell(fT[k]), moneyCell(land[k], bold), moneyCell(fb[k]), moneyCell(land[k] - fb[k]), cell(e.risk(land[k], fb[k]), { color: e.risk(land[k], fb[k]) === "High" ? C.bad : C.grey })]); });
      s.addTable(rows, tableOpts({ x: 0.6, y: 1.45, w: W - 1.2, colW: [2.4, 1.8, 2.0, 1.8, 1.8, 1.6, 0.73] }));
      var fInc = e.forecastByDept(fcKeys, "income"), bInc = e.budgetByDept(fyKeys, "income");
      var dl = depts.map(function (d) { var ai = e.sumMonths((e.idx.deptCatMonth[d.code] || {}).income, actKeys); return { d: d, full: ai + (fInc[d.code] || 0), bud: bInc[d.code] || 0 }; });
      s.addChart(P.ChartType.bar, [{ name: "Full year", labels: dl.map(function (x) { return x.d.short; }), values: dl.map(function (x) { return cents(x.full); }) }, { name: "Budget", labels: dl.map(function (x) { return x.d.short; }), values: dl.map(function (x) { return cents(x.bud); }) }],
                 { x: 0.6, y: 3.9, w: W - 1.2, h: 2.6, barDir: "col", chartColors: [C.lavender, C.greyLight], showLegend: true, legendPos: "b", legendFontSize: 9, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, valAxisLabelFormatCode: "$#,##0", valGridLine: { color: C.line, size: 0.5 } });
      commentBox(s, ["forecast", "rev-forecast", "pipeline"], 6.55, 0.5);
      note(s, "Income on booked work, weighted deals and the baseline; costs on their methods; typed overrides on top. Budget is the comparison.", 3.45);
    }

    // 9 cash
    if (cash) {
      s = frame(P, "Cash flow, twelve months ahead", "From " + ((e.monthIdx[rm] || {}).label || rm) + ", indirect method", C.mint);
      tiles(s, [
        { label: "Cash now", value: dollars(cash.opening), sub: cash.balances.length + " accounts", accent: C.navy },
        { label: "Low point", value: dollars(cash.low.closing), sub: cash.low.label, tone: cash.low.closing < 0 ? "bad" : "good", accent: C.rose },
        { label: "In twelve months", value: dollars(cash.closing), sub: e.monthIdx[cash.months[cash.months.length - 1]].label, accent: C.mint },
        { label: "Credit cards owing", value: dollars(cash.cardOwing), sub: "limit " + dollars(cash.cardLimit), accent: C.lavender },
      ], 1.45);
      s.addChart(P.ChartType.line, [{ name: "Closing cash", labels: cash.rows.map(function (r) { return r.label; }), values: cash.rows.map(function (r) { return cents(r.closing); }) }],
                 { x: 0.6, y: 3.0, w: W - 1.2, h: 3.5, chartColors: ["0E9E99"], lineSize: 2.5, lineDataSymbolSize: 7, showLegend: false, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, valAxisLabelFormatCode: "$#,##0", valGridLine: { color: C.line, size: 0.5 } });
      commentBox(s, ["cash"], 6.55, 0.5);
      note(s, "Receipts on " + cash.settings.debtorDays + " days to pay, payments on " + cash.settings.creditorDays + ", wages in the month, GST on the BAS, commitments on their dates, cards cleared next month.", 2.8);
    }

    // 10 commentary
    var comments = e.commentaryFor ? e.commentaryFor(null, rm, { all: true }).filter(function (c) { return !c.draft; }) : [];
    if (comments.length) {
      s = frame(P, "Commentary", mLabel, C.lavender);
      s.addText(comments.map(function (c) { return [{ text: ((CTS.pages[c.page] || {}).title || c.page) + (c.dept ? ", " + ((e.deptOf[c.dept] || {}).short || c.dept) : ""), options: { bold: true, color: C.navy, fontSize: 12, breakLine: true } }, { text: c.text, options: { color: C.ink, fontSize: 11, breakLine: true } }, { text: c.author, options: { color: C.grey, fontSize: 9, breakLine: true, paraSpaceAfter: 8 } }]; }).reduce(function (a, x) { return a.concat(x); }, []),
                { x: 0.6, y: 1.45, w: W - 1.2, h: H - 2.5, fontFace: FONT, valign: "top" });
    }

    // 11 actions
    var items = [];
    var actP = e.pnlActual(p.keys), budP = e.pnlBudget(p.keys);
    actP.cats.forEach(function (cat) { var bcat = budP.by[cat.key]; cat.children.forEach(function (sub) { var bsub = (bcat.children || []).filter(function (x) { return x.sub === sub.sub; })[0]; var bv = bsub ? bsub.total : 0; if (e.risk(sub.total, bv) === "High") items.push({ kind: "Variance", what: cat.label + ": " + sub.sub, detail: dollars(sub.total) + " against budget " + dollars(bv), amount: sub.total - bv }); }); });
    dRows.forEach(function (r) { if (r.dept.gmNorm != null && r.gmPct != null && r.gmPct < r.dept.gmNorm) items.push({ kind: "Below norm", what: r.dept.short + " gross margin " + pct(r.gmPct), detail: "Norm " + pct(r.dept.gmNorm, 0), amount: Math.round((r.dept.gmNorm - r.gmPct) * r.income) }); });
    util.rows.forEach(function (r) { if (r.target != null && r.util != null && r.util < r.target - 0.05) items.push({ kind: "Under target", what: r.dept.short + " utilisation " + pct(r.util), detail: "Target " + pct(r.target, 0), amount: 0 }); });
    items.sort(function (a, b) { return Math.abs(b.amount) - Math.abs(a.amount); });
    s = frame(P, "Worth a comment", p.label, C.rose);
    if (items.length) {
      rows = [hdr([{ label: "Kind", align: "left" }, { label: "What", align: "left" }, { label: "Detail", align: "left" }, { label: "Effect" }])];
      items.slice(0, 14).forEach(function (it) { rows.push([cell(it.kind, { align: "left", color: C.grey }), cell(it.what, { align: "left", bold: true }), cell(it.detail, { align: "left" }), it.amount ? moneyCell(it.amount) : cell("")]); });
      s.addTable(rows, tableOpts({ x: 0.6, y: 1.45, w: W - 1.2, colW: [1.4, 4.2, 5.0, 1.5], fontSize: 10 }));
    } else s.addText("Nothing above the thresholds this period.", { x: 0.6, y: 1.5, w: 10, h: 0.5, fontFace: FONT, fontSize: 12, color: C.grey });
    note(s, "The Controller Pack's rule: the materiality floor is checked before the percentage, so a big percentage on a small variance is still Low. Commentary explains the variance rather than restating it.");

    // 12 closing
    s = P.addSlide(); s.background = { color: C.navy };
    s.addShape(P.ShapeType.ellipse, { x: 9.2, y: 3.6, w: 5.6, h: 5.6, fill: { color: C.lavender, transparency: 55 }, line: { color: C.lavender, transparency: 100 } });
    if (b && b.logo && b.logo.dark) s.addImage({ data: b.logo.dark, x: 0.75, y: 0.6, w: 2.2, h: 1.5 });
    s.addText("The full result is in the portal", { x: 0.75, y: 3.0, w: 9, h: 0.9, fontFace: FONT, fontSize: 36, bold: true, color: C.white });
    var url = (e.cfg.ORG && e.cfg.ORG.portalUrl) || "";
    s.addText(url ? [{ text: "Open ", options: {} }, { text: "CTS Business Intelligence Portal", options: { hyperlink: { url: url }, color: C.mint, bold: true } }, { text: ", pick your name, and every figure in this deck is live with its workings, the forecast and the commentary." }]
                  : "Open CTS Business Intelligence Portal.html from the synced CTS Business Portal folder, pick your name, and every figure in this deck is live with its workings, the forecast and the commentary.",
              { x: 0.75, y: 4.0, w: 8.5, h: 1.2, fontFace: FONT, fontSize: 16, color: C.greyLight, valign: "top" });
    s.addText("Corporate Technology Services  |  Seamless AV  |  Confidential", { x: 0.75, y: H - 0.7, w: 8, h: 0.4, fontFace: FONT, fontSize: 10, color: C.greyLight });
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
      var subject = "CTS month end result, " + mLabel + " (deck attached)";
      var body = '<p style="font-family:Segoe UI,Calibri,Arial,sans-serif;font-size:13.5px;color:#1E2530;line-height:1.5">Hi ' + r.name.split(" ")[0] + ",<br><br>The month end pack for " + mLabel + " is attached: revenue " + F().dollars(act.income) + " against " + F().dollars(bud.income) + " budgeted, net profit " + F().money(act.netProfit) + " against " + F().money(bud.netProfit) + ". The deck has the P&amp;L, the departments after the overhead split, revenue, top clients, utilisation, the full year landing, cash and the commentary.</p>";
      var html = M.shellFor ? M.shellFor(subject, "For " + r.name + " (" + M.TIERS[r.tier] + "). Built " + new Date().toISOString().slice(0, 10) + ".", body) : body;
      return { to: r.email || "", name: r.name, tier: r.tier, subject: subject, html: html, text: "The month end pack for " + mLabel + " is attached. See the full result in the portal.", attachments: [file], kind: "deck" };
    });
  };
})();
