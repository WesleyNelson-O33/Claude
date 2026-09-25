/* CTS Business Intelligence Portal - the monthly email
 * ---------------------------------------------------------------------------
 * Renders the month as email, one per recipient, in three tiers:
 *
 *   1  Executive         the whole company: result, departments after the
 *                        overhead split, top clients under the four headings,
 *                        utilisation, and what is worth a comment
 *   2  Department head   their own department in depth, one line on the rest
 *   3  Finance           the controls: does the ledger tie, what is uncoded,
 *                        what rests on a placeholder, what the build did
 *
 * Sending is deliberately separate from rendering. The portal writes the
 * messages into outbox/YYYY-MM as JSON and HTML, and whatever sends them
 * (a Power Automate flow, later Graph) reads that folder. Nothing goes out
 * without somebody having clicked Build first, which is the same decision
 * O33 made: automated scheduling risks distributing unreviewed data.
 */
(function () {
  "use strict";
  var CTS = (window.CTS = window.CTS || {});
  var M = (CTS.email = {});
  var F = function () { return CTS.fmt; };
  var E = function () { return CTS.engine; };

  M.TIERS = { 1: "Executive", 2: "Department head", 3: "Finance" };

  M.recipients = function () {
    var e = E();
    return (e.recipients ? e.recipients() : (e.cfg.DISTRIBUTION || [])).filter(function (r) { return r.send !== false; });
  };
  M.CADENCE = { monthly: "Monthly", fortnightly: "Fortnightly", both: "Monthly and fortnightly" };

  /* ---- email safe HTML in the CTS livery: tables and inline styles ------ */
  var FONT = "font-family:'Segoe UI',Calibri,Arial,Helvetica,sans-serif;";
  var C = { black: "#121820", ink: "#1E2530", grey: "#5F6670", line: "#DDE1E7", panel: "#F6F7F9", navy: "#122B82", mint: "#3FD0C9", lavender: "#6D71FF", rose: "#F33844", teal: "#005358", good: "#0A6B2C", bad: "#A0161F", warn: "#7D5600" };
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]; }); }
  function h1(t) { return '<h1 style="' + FONT + 'font-size:22px;color:#FFFFFF;margin:0 0 4px;letter-spacing:-.01em">' + esc(t) + "</h1>"; }
  function h2(t) { return '<h2 style="' + FONT + 'font-size:15px;color:' + C.navy + ';margin:24px 0 8px;border-bottom:2px solid ' + C.mint + ';padding-bottom:4px;text-transform:uppercase;letter-spacing:.06em">' + esc(t) + "</h2>"; }
  function p(t, muted) { return '<p style="' + FONT + 'font-size:13.5px;color:' + (muted ? C.grey : C.ink) + ';margin:6px 0;line-height:1.5">' + t + "</p>"; }
  function table(cols, rows) {
    var out = '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;' + FONT + 'font-size:12.5px;color:' + C.ink + '">';
    out += "<tr>" + cols.map(function (c) {
      return '<th style="text-align:' + (c.align || "right") + ';padding:7px 8px;background:' + C.panel + ';color:' + C.grey + ';font-size:11px;text-transform:uppercase;letter-spacing:.04em;border-bottom:1px solid ' + C.line + '">' + esc(c.label) + "</th>";
    }).join("") + "</tr>";
    rows.forEach(function (r) {
      out += '<tr>' + cols.map(function (c) {
        var v = typeof c.value === "function" ? c.value(r) : r[c.key];
        var txt = c.fmt ? c.fmt(v) : esc(v == null ? "-" : v);
        var neg = typeof v === "number" && v < 0 && c.fmt;
        return '<td style="text-align:' + (c.align || "right") + ';padding:6px 8px;border-bottom:1px solid ' + C.line + ';' +
               (r._bold ? "font-weight:bold;background:" + C.panel + ";" : "") + (r._muted ? "color:" + C.grey + ";font-style:italic;" : "") + (neg ? "color:" + C.bad + ";" : "") + '">' + txt + "</td>";
      }).join("") + "</tr>";
    });
    return out + "</table>";
  }
  function money(c) { return esc(F().money(c)); }
  function pct(x) { return esc(F().pct(x)); }
  function chip(t, tone) {
    var bg = { good: "#DDF3E4", warn: "#FBEFC9", bad: "#FBDFE1", muted: C.panel }[tone || "muted"];
    var ink = { good: C.good, warn: C.warn, bad: C.bad, muted: C.grey }[tone || "muted"];
    return '<span style="display:inline-block;padding:1px 8px;border-radius:9px;background:' + bg + ';color:' + ink + ';font-size:11px;font-weight:bold;' + FONT + '">' + esc(t) + "</span>";
  }
  /** A row of stat tiles, four or five across, the way the portal shows them. */
  function tiles(items) {
    var out = '<table cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:6px 0;width:100%;margin:10px -6px"><tr>';
    items.forEach(function (t) {
      var col = t.tone === "good" ? C.good : t.tone === "bad" ? C.bad : t.tone === "warn" ? C.warn : C.black;
      out += '<td style="vertical-align:top;background:' + C.panel + ';border-top:4px solid ' + (t.accent || C.navy) + ';padding:9px 10px;width:' + Math.floor(100 / items.length) + '%">' +
        '<div style="' + FONT + 'font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:' + C.grey + ';font-weight:bold">' + esc(t.label) + '</div>' +
        '<div style="' + FONT + 'font-size:20px;font-weight:bold;color:' + col + ';margin:3px 0 2px">' + esc(t.value) + '</div>' +
        (t.sub ? '<div style="' + FONT + 'font-size:11.5px;color:' + C.grey + '">' + esc(t.sub) + '</div>' : "") + "</td>";
    });
    return out + "</tr></table>";
  }
  function callout(title, body, accent) {
    return '<div style="background:' + C.panel + ';border-left:4px solid ' + (accent || C.mint) + ';padding:10px 14px;margin:14px 0">' +
      '<div style="' + FONT + 'font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:' + C.navy + ';font-weight:bold;margin-bottom:4px">' + esc(title) + '</div>' + body + "</div>";
  }
  function bullets(items) {
    return '<ul style="' + FONT + 'font-size:13px;color:' + C.ink + ';margin:4px 0 4px 18px;padding:0;line-height:1.5">' + items.map(function (t) { return "<li style=\"margin:4px 0\">" + t + "</li>"; }).join("") + "</ul>";
  }
  /** Where to open the portal: the link on the Config Reporting tab, else
   *  the folder path. Every email says it, because the email is the summary
   *  and the portal is the result. */
  function portalLink() {
    var e = E(), url = (e.cfg && e.cfg.ORG && e.cfg.ORG.portalUrl) || "";
    var inner = '<div style="' + FONT + 'font-size:13px;color:' + C.ink + ';line-height:1.5"><b>See the full result in the portal.</b> Every figure here, with the workings, the departments, the forecast and the commentary, is live in the CTS Business Intelligence Portal. ' +
      (url ? '<a href="' + esc(url) + '" style="color:' + C.navy + ';font-weight:bold">Open the portal</a>' : 'Open <b>CTS Business Intelligence Portal.html</b> from the synced <b>CTS Business Portal</b> folder in OneDrive') +
      ', pick your name, and the Dashboard is the same month this email is about.</div>';
    return '<div style="margin:14px 0 4px;padding:12px 14px;background:#EEEEFF;border-left:4px solid ' + C.lavender + '">' + inner + "</div>";
  }
  M.portalLink = portalLink;
  M.shellFor = function (title, sub, body, opts) { return shell(title, sub, body, opts); };
  function shell(title, sub, body, opts) {
    opts = opts || {};
    var accent = opts.accent || C.mint;
    return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>' + esc(title) + '</title></head>' +
      '<body style="margin:0;padding:0;background:#F3F5F8"><div style="max-width:780px;margin:0 auto;padding:18px 14px">' +
      '<div style="background:' + C.black + ';padding:22px 24px 18px;border-radius:10px 10px 0 0">' +
      '<div style="' + FONT + 'font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:' + accent + ';font-weight:bold;margin-bottom:8px">cts &nbsp;|&nbsp; Seamless AV' + (opts.division ? ' &nbsp;|&nbsp; ' + esc(opts.division) : "") + '</div>' +
      h1(title) + '<div style="' + FONT + 'font-size:13px;color:#C9CED6">' + esc(sub) + '</div></div>' +
      '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse"><tr><td style="height:4px;background:' + C.mint + '"></td><td style="height:4px;background:' + C.lavender + '"></td><td style="height:4px;background:' + C.rose + '"></td></tr></table>' +
      '<div style="background:#FFFFFF;padding:16px 24px 22px;border-radius:0 0 10px 10px">' + portalLink() + body + portalLink() +
      '<p style="' + FONT + 'font-size:11px;color:' + C.grey + ';margin-top:26px;border-top:1px solid ' + C.line + ';padding-top:10px">' +
      "Corporate Technology Services &nbsp;|&nbsp; Seamless AV &nbsp;|&nbsp; Confidential. Generated by the CTS Business Intelligence Portal from the templates in the portal folder; figures are as at the build shown above. Open the portal for the live view.</p></div></div></body></html>";
  }

  /* ---- the pieces --------------------------------------------------------- */
  function companyBlock(keys, label) {
    var e = E(), pnl = e.pnlActual(keys), bud = e.pnlBudget(keys), prior = e.pnlActual(e.priorYearMonths(keys));
    var rows = [
      ["Revenue", pnl.totals.income, bud.totals.income, prior.totals.income],
      ["Cost of sales", pnl.totals.cos, bud.totals.cos, prior.totals.cos],
      ["Gross profit", pnl.totals.grossProfit, bud.totals.grossProfit, prior.totals.grossProfit, true],
      ["Overheads", pnl.totals.expenses, bud.totals.expenses, prior.totals.expenses],
      ["Other income", pnl.totals.otherIncome, bud.totals.otherIncome, prior.totals.otherIncome],
      ["Net profit", pnl.totals.netProfit, bud.totals.netProfit, prior.totals.netProfit, true],
    ].map(function (r) { return { l: r[0], a: r[1], b: r[2], py: r[3], v: r[1] - r[2], _bold: !!r[4] }; });
    var gm = pnl.totals.income ? pnl.totals.grossProfit / pnl.totals.income : null;
    var outlook = "";
    if (e.forecastEnabled && e.forecastEnabled()) {
      var fyK = e.monthsOfFY(e.currentFY()), land = e.pnlBlend(fyK).totals, fb = e.pnlBudget(fyK).totals;
      outlook = p("<b>Full year FY" + e.currentFY() + " outlook:</b> revenue " + money(land.income) + " against budget " + money(fb.income) +
                  ", net profit " + money(land.netProfit) + " against budget " + money(fb.netProfit) +
                  ". Actual to date then forecast, line by line on the methods set for the month.");
      if (e.cashEnabled && e.cashEnabled()) {
        var c = e.cashflow();
        if (c) outlook += p("<b>Cash:</b> " + money(c.opening) + " at month end, low point " + money(c.low.closing) + " in " + esc(c.low.label) +
                            ", " + money(c.closing) + " in twelve months. Credit cards owing " + money(c.cardOwing) + ".");
      }
    }
    return h2("Company result, " + label) +
      table([{ key: "l", label: "", align: "left" }, { key: "a", label: "Actual", fmt: money },
             { key: "b", label: "Budget", fmt: money }, { key: "v", label: "Variance", fmt: money },
             { key: "py", label: "Last year", fmt: money }], rows) +
      p("Gross margin " + pct(gm) + ". Compared in dollars rather than percentages, as management prefer.", true) + outlook;
  }

  function deptBlock(keys, label, only) {
    var e = E(), alloc = e.allocate(keys);
    var rows = alloc.rows.filter(function (r) { return r.dept.isRevenue && (r.income || r.expenses); })
      .filter(function (r) { return !only || r.code === only; })
      .map(function (r) {
        var norm = r.dept.gmNorm;
        return { d: r.dept.short, rev: r.income, gm: r.gmPct, own: r.ownExpenses, share: r.allocated, np: r.netProfit,
                 flag: norm != null && r.gmPct != null && r.gmPct < norm ? chip("below " + F().pct(norm, 0) + " norm", "warn") : chip("ok", "good") };
      });
    var s = e.allocationIsSourced();
    return h2((only ? "Your department" : "Departments after the overhead split") + ", " + label) +
      table([{ key: "d", label: "Department", align: "left" }, { key: "rev", label: "Revenue", fmt: money },
             { key: "gm", label: "GM %", fmt: pct }, { key: "own", label: "Own overhead", fmt: money },
             { key: "share", label: "Share of Admin", fmt: money }, { key: "np", label: "Net profit", fmt: money },
             { key: "flag", label: "", align: "left", fmt: function (v) { return v; } }], rows) +
      (s.ok ? "" : p(chip("placeholder", "warn") + " The " + esc(s.placeholders.join(" and ")) + " split percentages are placeholders. Allocated figures move when they are confirmed.", true));
  }

  function oneLiners(keys, except) {
    var e = E(), alloc = e.allocate(keys);
    var lines = alloc.rows.filter(function (r) { return r.dept.isRevenue && r.code !== except && r.income; })
      .map(function (r) { return esc(r.dept.short) + " " + money(r.income) + " revenue, " + pct(r.gmPct) + " margin"; });
    return h2("The other departments, one line each") + p(lines.join("<br>"));
  }

  function clientsBlock(keys) {
    var e = E(), rows = e.clientRows(keys, e.priorYearMonths(keys)).filter(function (r) { return r.amount || r.prior; });
    var head = (e.cfg.COMMENTARY || {}).topClientHeadings || ["Revenue increased", "Revenue decreased", "New to the list", "Dropped off the list"];
    var inc = rows.filter(function (r) { return r.prior && r.delta > 0; }).slice(0, 5);
    var dec = rows.filter(function (r) { return r.prior && r.delta < 0; }).sort(function (a, b) { return a.delta - b.delta; }).slice(0, 5);
    var nw = rows.filter(function (r) { return r.amount && !r.prior; }).slice(0, 5);
    var dr = rows.filter(function (r) { return !r.amount && r.prior; }).slice(0, 5);
    function list(t, items, key) {
      return "<b>" + esc(t) + "</b><br>" + (items.length ? items.map(function (r) { return esc(r.display) + " " + money(r[key]); }).join("<br>") : "None") + "<br><br>";
    }
    return h2("Top clients, against the same period last year") +
      table([{ key: "rank", label: "#", align: "left" }, { key: "display", label: "Client", align: "left" },
             { key: "amount", label: "This period", fmt: money }, { key: "prior", label: "Last year", fmt: money },
             { key: "delta", label: "Variance", fmt: money }], rows.slice(0, 10)) +
      p(list(head[0], inc, "delta") + list(head[1], dec, "delta") + list(head[2], nw, "amount") + list(head[3], dr, "prior"));
  }

  function utilBlock(keys, only) {
    var e = E(), u = e.utilFor(keys), roll = e.utilRolling(keys[keys.length - 1]);
    var rows = u.rows.filter(function (r) { return r.total && (!only || r.code === only); }).map(function (r) {
      return { d: r.dept.short, c: r.chargeable, n: r.nonChargeable, u: r.util, t: r.target,
               gap: r.gap, fte: r.fte, flag: r.gap != null && r.gap < -0.05 ? chip("under target", "warn") : "" };
    });
    return h2("Utilisation" + (only ? "" : ", all departments")) +
      table([{ key: "d", label: "Department", align: "left" }, { key: "c", label: "Chargeable", fmt: function (v) { return esc(F().hours(v)); } },
             { key: "n", label: "Non-charge", fmt: function (v) { return esc(F().hours(v)); } }, { key: "u", label: "Util", fmt: pct },
             { key: "t", label: "Target", fmt: function (v) { return v == null ? "-" : esc(F().pct(v, 0)); } },
             { key: "fte", label: "FTE", fmt: function (v) { return v == null ? "-" : v.toFixed(2); } },
             { key: "flag", label: "", align: "left", fmt: function (v) { return v; } }], rows) +
      p("Rolling twelve months " + pct(roll.total.util) + " on " + esc(F().num(roll.annualHours)) + " hours" +
        (roll.annualCheck ? ", which lands on the 2,080 or 2,088 the graphs file expects." : ", which does not land on 2,080 or 2,088; check the Days tab.") +
        " Read a small negative before flagging it: for a team whose hours are almost all chargeable it usually means leave.", true);
  }

  function actionsBlock(keys) {
    var e = E(), act = e.pnlActual(keys), bud = e.pnlBudget(keys), items = [];
    act.cats.forEach(function (cat) {
      var bcat = bud.by[cat.key];
      cat.children.forEach(function (sub) {
        var bs = (bcat.children || []).filter(function (x) { return x.sub === sub.sub; })[0];
        var b = bs ? bs.total : 0;
        if (e.risk(sub.total, b) === "High") items.push({ w: cat.label + ": " + sub.sub, a: sub.total, b: b, v: sub.total - b });
      });
    });
    items.sort(function (a, b) { return Math.abs(b.v) - Math.abs(a.v); });
    return h2("Worth a comment") + (items.length ?
      table([{ key: "w", label: "Line", align: "left" }, { key: "a", label: "Actual", fmt: money },
             { key: "b", label: "Budget", fmt: money }, { key: "v", label: "Variance", fmt: money }], items.slice(0, 12)) +
      p("High risk on the Controller Pack's thresholds. Commentary explains the variance rather than restating it; for consulting, say whether a deal was lost or moved.", true)
      : p("Nothing above the high risk thresholds this period.", true));
  }

  function controlsBlock(keys, log) {
    var e = E(), pc = e.plCheck(keys), q = e.quality, s = e.allocationIsSourced();
    var rows = pc.map(function (r) { return { c: r.label, gl: r.gl, pl: r.pl, d: r.diff, s: r.ok ? chip("OK", "good") : chip("CHECK", "bad") }; });
    var body = h2("Does the ledger agree with the P&L") +
      table([{ key: "c", label: "Category", align: "left" }, { key: "gl", label: "Ledger", fmt: money },
             { key: "pl", label: "P&L", fmt: money }, { key: "d", label: "Difference", fmt: money },
             { key: "s", label: "", align: "left", fmt: function (v) { return v; } }], rows) +
      h2("Coding quality") +
      p("Ledger lines " + esc(F().num(q.rows)) + ", " + esc(F().date(q.minDate)) + " to " + esc(F().date(q.maxDate)) + ".<br>" +
        "Lines with no department: <b>" + q.noDept + "</b>. On an unknown account: <b>" + q.unknownAcct + "</b>. Revenue lines with no contact: <b>" + q.noContact + "</b>.") +
      h2("What the figures rest on") +
      p(s.ok ? chip("confirmed", "good") + " All three split bases are confirmed."
             : chip("placeholder", "warn") + " " + esc(s.placeholders.join(" and ")) + " split percentages are placeholders. Every allocated departmental figure depends on them.");
    if (log && log.checks) {
      body += h2("The build") + p((log.checks || []).map(function (c) {
        return chip(c.level, c.level === "ok" ? "good" : c.level === "fail" ? "bad" : c.level === "warn" ? "warn" : "muted") + " <b>" + esc(c.area) + "</b> " + esc(c.text);
      }).join("<br>"));
    }
    return body;
  }

  /* ---- one message per recipient ------------------------------------------ */
  /** Published commentary for the month, for the pages a tier reads. */
  function commentaryBlock(pages, dept) {
    var e = E(), rm = e.reportingMonth();
    if (!e.commentaryFor) return "";
    var items = [];
    pages.forEach(function (pg) {
      e.commentaryFor(pg, rm, { all: true }).forEach(function (c) {
        if (c.draft) return;
        if (dept && c.dept && c.dept !== dept) return;
        items.push(Object.assign({ pageTitle: (window.CTS.pages && window.CTS.pages[pg] ? window.CTS.pages[pg].title : pg) }, c));
      });
    });
    if (!items.length) return "";
    return h2("Commentary") + items.map(function (c) {
      return p("<b>" + esc(c.pageTitle) + (c.dept ? ", " + esc((e.deptOf[c.dept] || {}).short || c.dept) : "") + "</b> " +
               esc(c.text).replace(/\n/g, "<br>") + ' <span style="color:#595959">' + esc(c.author) + "</span>");
    }).join("");
  }
  M.commentaryBlock = commentaryBlock;

  M.render = function (recipient, opts) { return E().withoutView ? E().withoutView(function () { return renderInner(recipient, opts); }) : renderInner(recipient, opts); };
  function renderInner(recipient, opts) {
    opts = opts || {};
    var e = E(), keys = opts.keys || e.ytd(), rm = e.reportingMonth();
    var label = "FY" + e.currentFY() + " to " + ((e.monthIdx[rm] || {}).label || rm);
    var monthLong = (e.monthIdx[rm] || {}).long || rm;
    var tier = +recipient.tier || 3, body = "", subject;
    if (tier === 1) {
      subject = "CTS monthly result, " + monthLong;
      body = companyBlock(keys, label) + commentaryBlock(["home", "pnl", "forecast", "cash", "rev-summary", "rev-forecast", "bva"]) + deptBlock(keys, label) + clientsBlock(keys) + utilBlock(keys) + actionsBlock(keys);
    } else if (tier === 2) {
      var d = recipient.dept || null;
      var dn = d && e.deptOf[d] ? e.deptOf[d].short : "your department";
      subject = "CTS " + dn + " result, " + monthLong;
      body = deptBlock(keys, label, d) + commentaryBlock(["pnl-dept", "rev-schedule", "rev-forecast", "util", "pipeline"], d) + utilBlock(keys, d) + oneLiners(keys, d);
    } else {
      subject = "CTS month end controls, " + monthLong;
      body = controlsBlock(keys, opts.log) + companyBlock(keys, label) + commentaryBlock(Object.keys(window.CTS.pages || {}));
    }
    var html = shell(subject, "For " + recipient.name + " (" + M.TIERS[tier] + "). Built " + new Date().toISOString().slice(0, 10) + ".", body, { accent: tier === 2 ? C.lavender : C.mint, division: tier === 2 && recipient.dept && e.deptOf[recipient.dept] ? "CTS " + e.deptOf[recipient.dept].short : null });
    var text = html.replace(/<style[\s\S]*?<\/style>/g, "").replace(/<\/(p|tr|h1|h2|div)>/g, "\n").replace(/<[^>]+>/g, "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/\n{3,}/g, "\n\n").trim();
    return { to: recipient.email || "", name: recipient.name, tier: tier, subject: subject, html: html, text: text };
  };


  /* ---- the fortnightly update ------------------------------------------ */
  function fnHeadline(x, name) {
    var f = x.current, bench = x.benchmark, kind = x.benchmarkKind;
    var vs = x.mtd - x.proRata, pctv = x.proRata ? vs / x.proRata : null;
    var s = [];
    s.push(f.label + " " + (f.complete ? "closed at " : "is running at ") + F().dollars(f.revenue) + " for " + name + ".");
    if (bench) s.push("Month to date " + F().dollars(x.mtd) + " against " + F().dollars(x.proRata) + " pro rata of the " + F().dollars(bench) + " " + kind + " for the month (" + (vs >= 0 ? "+" : "−") + F().dollars(Math.abs(vs)) + (pctv != null ? ", " + (pctv >= 0 ? "+" : "") + (pctv * 100).toFixed(1) + "%" : "") + ").");
    if (x.util != null) s.push("Utilisation " + F().pct(x.util) + (x.target != null ? (x.util >= x.target ? ", at or above" : ", below") + " the " + F().pct(x.target, 0) + " target" : "") + " on " + F().hours(x.chargeable) + " chargeable hours" + (x.hoursProRated ? " (pro rated from the " + ((E().monthIdx[x.hoursMonth] || {}).label || x.hoursMonth) + " pay runs)" : "") + ".");
    if (x.landingTotal) s.push(F().dollars(x.landingTotal) + " of confirmed work is booked to land this month and next.");
    return s.join(" ");
  }
  function fnCalls(x, name) {
    var out = [], vs = x.mtd - x.proRata, tone = vs >= 0 ? "good" : "bad";
    out.push("<b>" + esc(name) + " revenue</b> " + money(x.mtd) + " month to date, " + chip((vs >= 0 ? "+" : "−") + F().dollars(Math.abs(vs)) + " vs pro rata " + x.benchmarkKind, tone) + (x.gm != null ? ". Gross margin to date " + pct(x.gm) + "." : "."));
    if (x.util != null) out.push("<b>Utilisation</b> " + pct(x.util) + (x.target != null ? " against " + pct(x.target) + " target " + chip(x.util >= x.target ? "on target" : x.util >= x.target - 0.05 ? "close" : "under", x.util >= x.target ? "good" : x.util >= x.target - 0.05 ? "warn" : "bad") : "") + ", sell rate " + (x.sellRate != null ? F().dollars(x.sellRate) + "/h" : "n/a") + " on " + x.headcount + " people.");
    if (x.topClients.length) out.push("<b>Largest this fortnight</b> " + x.topClients.slice(0, 3).map(function (c) { return esc(c.client) + " " + money(c.amount); }).join(", ") + ".");
    if (x.landing.length) out.push("<b>Landing next</b> " + x.landing.slice(0, 3).map(function (l) { return esc(l.client) + " " + money(l.value) + " (" + ((E().monthIdx[l.month] || {}).label || l.month) + ")"; }).join(", ") + ".");
    return out;
  }
  function fnTable(x) {
    var rows = x.fortnights.map(function (f) {
      return { w: f.label + " (" + f.range + ")", state: f.complete ? chip("actual", "good") : f.started ? chip("in progress", "warn") : chip("schedule", "muted"),
               rev: f.complete || f.started ? f.revenue : null, sched: f.schedule, hrs: f.chargeable, sell: f.sellRate, util: f.util, _muted: !f.started };
    });
    rows.push({ w: "Month to date", state: "", rev: x.mtd, sched: x.fortnights.reduce(function (a, f) { return a + f.schedule; }, 0), hrs: x.chargeable, sell: x.sellRate, util: x.util, _bold: true });
    rows.push({ w: "Month " + x.benchmarkKind, state: "", rev: null, sched: x.benchmark, hrs: null, sell: null, util: null, _muted: true });
    return table([
      { key: "w", label: "Fortnight", align: "left" }, { key: "state", label: "", align: "left", fmt: function (v) { return v || ""; } },
      { key: "rev", label: "Revenue", fmt: function (v) { return v == null ? "-" : money(v); } },
      { key: "sched", label: "Schedule", fmt: function (v) { return v == null ? "-" : money(v); } },
      { key: "hrs", label: "Chargeable hrs", fmt: function (v) { return v == null ? "-" : esc(F().hours(v)); } },
      { key: "sell", label: "Sell rate", fmt: function (v) { return v == null ? "-" : esc(F().dollars(v)) + "/h"; } },
      { key: "util", label: "Util %", fmt: function (v) { return v == null ? "-" : pct(v); } },
    ], rows);
  }
  function fnTiles(x, accent) {
    var vs = x.mtd - x.proRata;
    return tiles([
      { label: "Revenue MTD", value: F().dollars(x.mtd), sub: (vs >= 0 ? "+" : "−") + F().dollars(Math.abs(vs)) + " vs pro rata", tone: vs >= 0 ? "good" : "bad", accent: accent },
      { label: x.current.label, value: F().dollars(x.current.revenue), sub: x.current.complete ? "closed" : "to " + F().date(x.snapshot), accent: accent },
      { label: "Month " + x.benchmarkKind, value: F().dollars(x.benchmark), sub: "budget " + F().dollars(x.budget), accent: accent },
      { label: "Sell rate", value: x.sellRate != null ? F().dollars(x.sellRate) + "/h" : "n/a", sub: "revenue over chargeable hours", accent: accent },
      { label: "Utilisation", value: F().pct(x.util), sub: x.target != null ? "target " + F().pct(x.target, 0) : "", tone: x.target != null && x.util != null ? (x.util >= x.target ? "good" : x.util >= x.target - 0.05 ? "warn" : "bad") : null, accent: accent },
    ]) + tiles([
      { label: "Headcount", value: String(x.headcount), sub: (x.fte ? x.fte.toFixed(1) + " FTE" : ""), accent: accent },
      { label: "Chargeable hours", value: F().hours(x.chargeable), sub: x.hoursProRated ? "pro rated to date" : "month", accent: accent },
      { label: "Worked hours", value: F().hours(x.worked), sub: "leave excluded", accent: accent },
      { label: "Gross margin MTD", value: F().pct(x.gm), sub: "off the ledger", accent: accent },
      { label: "Confirmed ahead", value: x.landingTotal ? F().dollars(x.landingTotal) : "none yet", sub: "this month and next", accent: accent },
    ]);
  }
  function fnOthers(except) {
    var e = E(), rows = e.revenueDepts.filter(function (d) { return d.code !== "ADMIN" && d.code !== except; }).map(function (d) {
      var x = e.fortnightFor(d.code); return { d: d.short, rev: x.mtd, bench: x.benchmark, vs: x.mtd - x.proRata, sell: x.sellRate, util: x.util };
    });
    return table([
      { key: "d", label: "Department", align: "left" }, { key: "rev", label: "Revenue MTD", fmt: money }, { key: "bench", label: "Month " + "benchmark", fmt: money },
      { key: "vs", label: "vs pro rata", fmt: money }, { key: "sell", label: "Sell rate", fmt: function (v) { return v == null ? "-" : esc(F().dollars(v)) + "/h"; } },
      { key: "util", label: "Util %", fmt: function (v) { return v == null ? "-" : pct(v); } },
    ], rows);
  }
  function fnForward(x) {
    var nextF = x.fortnights.filter(function (f) { return !f.complete; })[0];
    var rows = [];
    if (nextF) rows.push({ w: nextF.label + " (" + nextF.range + ")", sched: nextF.schedule - (nextF.started ? nextF.revenue : 0), conf: x.landing.filter(function (l) { return l.month === x.month; }).reduce(function (a, l) { return a + l.value; }, 0) });
    var nk = (function (k) { var y = +k.slice(0, 4), m = +k.slice(5, 7); return m === 12 ? (y + 1) + "-01" : y + "-" + (m + 1 < 10 ? "0" : "") + (m + 1); })(x.month);
    var e = E(), nextBench = 0; (x.dept ? [x.dept] : e.revenueDepts.map(function (d) { return d.code; })).forEach(function (d) { nextBench += e.forecastEnabled() && e.hasForecast(nk) ? (e.forecastByDept([nk], "income")[d] || 0) : (e.budgetByDept([nk], "income")[d] || 0); });
    rows.push({ w: (e.monthIdx[nk] || {}).long || nk, sched: nextBench, conf: x.landing.filter(function (l) { return l.month === nk; }).reduce(function (a, l) { return a + l.value; }, 0) });
    var t = table([{ key: "w", label: "Forward schedule", align: "left" }, { key: "sched", label: "Forecast", fmt: money }, { key: "conf", label: "Confirmed in the systems", fmt: money },
                   { key: "cover", label: "Covered", align: "right", value: function (r) { return r.sched ? r.conf / r.sched : null; }, fmt: function (v) { return v == null ? "-" : pct(v); } }], rows);
    var text = x.remaining ? F().dollars(x.remaining) + " is still to come this month to reach the " + x.benchmarkKind + ". " : "The month " + x.benchmarkKind + " is already reached. ";
    text += x.landing.length ? "Confirmed work in OnRent and Qwilr covers " + F().dollars(x.landingTotal) + " across this month and next; the largest is " + esc(x.landing[0].client) + " at " + money(x.landing[0].value) + "." : "Nothing is confirmed in the systems yet for this month and next, so the schedule rests on the forecast.";
    return callout("Forward view", p(text), C.lavender) + t;
  }
  function fnLanding(x) {
    if (!x.landing.length) return "";
    return table([
      { key: "month", label: "Lands", align: "left", value: function (r) { return (E().monthIdx[r.month] || {}).label || r.month; } },
      { key: "client", label: "Client", align: "left" }, { key: "title", label: "What", align: "left" },
      { key: "kind", label: "From", align: "left", value: function (r) { return r.kind === "order" ? "OnRent" : r.kind === "quote" ? "Qwilr" : "Zoho"; } },
      { key: "value", label: "Value", fmt: money },
    ], x.landing.slice(0, 8));
  }

  M.renderFortnightly = function (recipient, opts) { return E().withoutView ? E().withoutView(function () { return fortnightlyInner(recipient, opts); }) : fortnightlyInner(recipient, opts); };
  function fortnightlyInner(recipient, opts) {
    opts = opts || {};
    var e = E(), tier = +recipient.tier || 3, dept = tier === 2 ? (recipient.dept || null) : null;
    var x = e.fortnightFor(dept, opts.month, opts.fortnight);
    var dn = dept && e.deptOf[dept] ? e.deptOf[dept].short : "CTS";
    var brand = window.CTS_BRAND || {}, sub = dept && brand.subBrandOf ? brand.subBrandOf[dept] : null;
    var accent = sub === "production" ? C.mint : sub === "consulting" ? C.lavender : sub === "support" ? C.rose : C.mint;
    var when = x.monthLabel + " • " + x.current.label + " of 2" + (x.current.complete ? " complete" : " in progress") + " • Snapshot " + F().date(x.snapshot);
    var subject, body;
    if (dept) {
      subject = "CTS " + dn + " fortnightly update, " + x.current.label + " " + x.monthLabel;
      body = fnTiles(x, accent) + callout("Headline", p(fnHeadline(x, dn)), accent) +
        h2("Fortnight breakdown") + fnTable(x) + p("Green is timesheet and ledger data; schedule is the forecast for the month spread over the working days still to run. Hours " + (x.hoursProRated ? "are pro rated from the last pay runs to the snapshot date" : "are the month's pay runs") + ".", true) +
        h2("Team, " + x.headcount + " people") + p(esc(x.team.join(", "))) +
        h2("Key calls") + bullets(fnCalls(x, dn)) + (M.commentaryBlock ? M.commentaryBlock(["pnl-dept", "rev-schedule", "rev-forecast", "util", "pipeline"], dept) : "") +
        h2("Other departments, month to date") + fnOthers(dept) +
        h2("Forward") + fnForward(x) + fnLanding(x);
    } else {
      subject = "CTS fortnightly update, " + x.current.label + " " + x.monthLabel;
      var deptRows = e.revenueDepts.filter(function (d) { return d.code !== "ADMIN"; }).map(function (d) { var y = e.fortnightFor(d.code, opts.month, opts.fortnight); return { d: d.short, code: d.code, x: y }; });
      body = fnTiles(x, accent) + callout("Headline", p(fnHeadline(x, "the company")), accent) +
        h2("Fortnight breakdown, company") + fnTable(x) +
        h2("Departments, month to date") + table([
          { key: "d", label: "Department", align: "left" },
          { key: "rev", label: "Revenue MTD", value: function (r) { return r.x.mtd; }, fmt: money },
          { key: "cur", label: x.current.label, value: function (r) { return r.x.current.revenue; }, fmt: money },
          { key: "bench", label: "Month " + x.benchmarkKind, value: function (r) { return r.x.benchmark; }, fmt: money },
          { key: "vs", label: "vs pro rata", value: function (r) { return r.x.mtd - r.x.proRata; }, fmt: money },
          { key: "sell", label: "Sell rate", value: function (r) { return r.x.sellRate; }, fmt: function (v) { return v == null ? "-" : esc(F().dollars(v)) + "/h"; } },
          { key: "util", label: "Util %", value: function (r) { return r.x.util; }, fmt: function (v) { return v == null ? "-" : pct(v); } },
          { key: "hc", label: "People", value: function (r) { return r.x.headcount; } },
        ], deptRows) +
        h2("Key calls by department") + bullets(deptRows.map(function (r) { return fnCalls(r.x, r.d)[0]; })) +
        (M.commentaryBlock ? M.commentaryBlock(["home", "story", "forecast", "cash", "rev-forecast"]) : "") +
        h2("Forward") + fnForward(x) + fnLanding(x);
    }
    var html = shell(subject, when + ". For " + recipient.name + " (" + M.TIERS[tier] + "). Built " + new Date().toISOString().slice(0, 10) + ".", body, { accent: accent, division: dept ? "CTS " + dn : null });
    var text = html.replace(/<style[\s\S]*?<\/style>/g, "").replace(/<\/(p|tr|h1|h2|div|li)>/g, "\n").replace(/<[^>]+>/g, "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&nbsp;/g, " ").replace(/\n{3,}/g, "\n\n").trim();
    return { to: recipient.email || "", name: recipient.name, tier: tier, subject: subject, html: html, text: text, kind: "fortnightly", month: x.month, fortnight: x.current.n };
  };
  M.renderAllFortnightly = function (opts) {
    return M.recipients().filter(function (r) { return /fortnight|both/i.test(r.cadence || "monthly"); }).map(function (r) { return M.renderFortnightly(r, opts); });
  };

  M.renderAll = function (opts) {
    return M.recipients().map(function (r) { return M.render(r, opts); });
  };

  /** The one dispatch that needs nothing: hand the message to the mail
   *  client. mailto carries plain text only and browsers cap its length, so
   *  the body is a summary with a pointer to the outbox HTML. */
  M.mailto = function (msg) {
    var body = msg.text.length > 1800 ? msg.text.slice(0, 1800) + "\n\n[Full report in the portal outbox]" : msg.text;
    return "mailto:" + encodeURIComponent(msg.to) + "?subject=" + encodeURIComponent(msg.subject) + "&body=" + encodeURIComponent(body);
  };
})();
