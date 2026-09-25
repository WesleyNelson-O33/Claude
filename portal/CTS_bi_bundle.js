/* CTS Business Intelligence Portal
 * ---------------------------------------------------------------------------
 * Client side only. No backend, no network calls at runtime. Every figure on
 * screen is re-aggregated in the browser on load from the window.CTS_* data
 * files, the same way the Controller Pack re-aggregates a GL paste.
 *
 * Sign convention, everywhere inside this file: amount = credit - debit, in
 * whole cents. Income is positive, costs are negative. That is the Controller
 * Pack's convention, and it is why gross profit here is income PLUS cost of
 * sales rather than minus. Display helpers flip costs to positive magnitudes
 * at the last moment; nothing else does.
 */
(function () {
  "use strict";

  var CTS = (window.CTS = window.CTS || {});

  /* ======================================================== formatting */
  var F = (CTS.fmt = {
    money: function (c, dp) {
      if (c == null || isNaN(c)) return "-";
      // Table columns call fmt(value, row); anything that is not a number of
      // decimal places is ignored rather than handed to toLocaleString.
      dp = typeof dp === "number" ? dp : 0;
      var v = c / 100;
      if (v === 0) return "-";
      var s = Math.abs(v).toLocaleString("en-AU", {
        minimumFractionDigits: dp, maximumFractionDigits: dp,
      });
      return v < 0 ? "(" + s + ")" : s;
    },
    money0: function (c) { return F.money(c, 0); },
    dollars: function (c) {
      if (c == null || isNaN(c)) return "-";
      return "$" + F.money(c, 0);
    },
    k: function (c) {
      if (c == null || isNaN(c) || c === 0) return "-";
      var v = c / 100000;
      var s = Math.abs(v).toFixed(Math.abs(v) < 100 ? 1 : 0) + "k";
      return v < 0 ? "(" + s + ")" : s;
    },
    pct: function (x, dp) {
      if (x == null || isNaN(x) || !isFinite(x)) return "-";
      dp = typeof dp === "number" ? dp : 1;
      var s = (Math.abs(x) * 100).toFixed(dp) + "%";
      return x < 0 ? "(" + s + ")" : s;
    },
    hours: function (h) {
      if (h == null || isNaN(h)) return "-";
      return h.toLocaleString("en-AU", { maximumFractionDigits: 0 });
    },
    num: function (n, dp) {
      if (n == null || isNaN(n)) return "-";
      dp = typeof dp === "number" ? dp : 0;
      return n.toLocaleString("en-AU", {
        minimumFractionDigits: dp, maximumFractionDigits: dp });
    },
    date: function (iso) {
      if (!iso) return "";
      var p = iso.split("-");
      return p[2] + " " + ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][+p[1] - 1] + " " + p[0];
    },
  });

  /* ============================================================== DOM */
  function h(tag, attrs, kids) {
    var parts = tag.split(/([.#])/);
    var el = document.createElement(parts[0] || "div");
    for (var i = 1; i < parts.length; i += 2) {
      if (parts[i] === ".") el.classList.add(parts[i + 1]);
      else el.id = parts[i + 1];
    }
    if (attrs && (attrs.nodeType || typeof attrs === "string" || Array.isArray(attrs))) {
      kids = attrs; attrs = null;
    }
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        var v = attrs[k];
        if (v == null || v === false) return;
        if (k === "class") el.className += (el.className ? " " : "") + v;
        else if (k === "text") el.textContent = v;
        else if (k === "html") el.innerHTML = v;
        else if (k.slice(0, 2) === "on") el.addEventListener(k.slice(2), v);
        else if (k === "style" && typeof v === "object") Object.assign(el.style, v);
        else el.setAttribute(k, v === true ? "" : v);
      });
    }
    (Array.isArray(kids) ? kids : kids == null ? [] : [kids]).forEach(function (kid) {
      if (kid == null || kid === false) return;
      el.appendChild(kid.nodeType ? kid : document.createTextNode(String(kid)));
    });
    return el;
  }
  CTS.h = h;

  function svg(tag, attrs, kids) {
    var el = document.createElementNS("http://www.w3.org/2000/svg", tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      if (attrs[k] == null || attrs[k] === false) return;
      if (k.slice(0, 2) === "on") el.addEventListener(k.slice(2), attrs[k]);
      else el.setAttribute(k, attrs[k]);
    });
    (Array.isArray(kids) ? kids : kids == null ? [] : [kids]).forEach(function (kid) {
      if (kid == null || kid === false) return;
      el.appendChild(kid.nodeType ? kid : document.createTextNode(String(kid)));
    });
    return el;
  }
  CTS.svg = svg;

  /* ============================================================ store */
  var store = (CTS.store = {
    get: function (k, d) {
      try { var v = localStorage.getItem("cts.bi." + k); return v == null ? d : JSON.parse(v); }
      catch (e) { return d; }
    },
    set: function (k, v) {
      try { localStorage.setItem("cts.bi." + k, JSON.stringify(v)); } catch (e) {}
    },
  });

  /* =========================================================== engine */
  var E = (CTS.engine = {});

  E.init = function () {
    var cfg = (E.cfg = window.CTS_CONFIG);
    var cal = (E.cal = window.CTS_CAL);
    E.months = cal.months.slice();
    E.monthIdx = {};
    E.months.forEach(function (m, i) { E.monthIdx[m.key] = m; m.i = i; });

    E.depts = cfg.DEPARTMENTS.slice().sort(function (a, b) { return a.order - b.order; });
    E.deptOf = {};
    E.depts.forEach(function (d) { E.deptOf[d.code] = d; });
    E.revenueDepts = E.depts.filter(function (d) { return d.isRevenue; });
    E.postingDepts = E.depts.filter(function (d) { return d.code !== "UNALLOCATED"; });

    // tag -> department, longest tag first so ONSITE beats ONS
    E.tagMap = {};
    E.depts.forEach(function (d) {
      (d.tags || []).forEach(function (t) { E.tagMap[t.toUpperCase()] = d.code; });
    });
    E.tagsByLength = Object.keys(E.tagMap).sort(function (a, b) { return b.length - a.length; });

    E.accounts = (window.CTS_ACCOUNTS.accounts || []).slice();
    E.acct = {};
    E.accounts.forEach(function (a) { E.acct[a.name] = a; });
    E.subsOf = {};
    E.accounts.forEach(function (a) {
      (E.subsOf[a.cat] = E.subsOf[a.cat] || {});
      (E.subsOf[a.cat][a.sub] = E.subsOf[a.cat][a.sub] || []).push(a.name);
    });

    E.loadGL(window.CTS_GL);
    E.loadFin(window.CTS_FIN);
    E.loadUtil(window.CTS_UTIL);
    E.loadStaff(window.CTS_STAFF);
    E.loadPipeline(window.CTS_PIPELINE);
    E.loadForecast(window.CTS_FORECAST);
    E.loadCash(window.CTS_CASH);
    E.loadCommentary(window.CTS_COMMENTARY);
    E.loadUsers(window.CTS_USERS);
    E.clients = window.CTS_CLIENTS || { clients: [], schedule: [] };
    E.clientMeta = {};
    (E.clients.clients || []).forEach(function (c) { E.clientMeta[c.contact] = c; });

    E.seed = !!(window.CTS_GL.meta && window.CTS_GL.meta.seed);
    E.restoreSnapshots();
    E.ready = true;
    return E;
  };

  /* ---- GL hydration -------------------------------------------------- */
  E.resolveDept = function (costCentre, jobNo) {
    var cc = (costCentre || "").trim().toUpperCase();
    if (cc) {
      for (var i = 0; i < E.tagsByLength.length; i++) {
        if (cc.indexOf(E.tagsByLength[i]) >= 0) return E.tagMap[E.tagsByLength[i]];
      }
    }
    var j = (jobNo || "").toUpperCase();
    var open = j.indexOf("["), close = j.indexOf("]");
    if (open >= 0 && close > open) {
      var tag = j.slice(open + 1, close).trim();
      if (E.tagMap[tag]) return E.tagMap[tag];
    }
    return "UNALLOCATED";
  };

  function blank() { return {}; }
  function bump(obj, k, m, v) {
    var t = obj[k] || (obj[k] = {});
    t[m] = (t[m] || 0) + v;
  }
  function bump2(obj, a, b, m, v) {
    var t = obj[a] || (obj[a] = {});
    bump(t, b, m, v);
  }

  E.loadGL = function (data) {
    var cols = data.cols, rows = data.rows || [];
    var ci = {}; cols.forEach(function (c, i) { ci[c] = i; });

    var gl = (E.gl = new Array(rows.length));
    var idx = (E.idx = {
      acctMonth: blank(), deptCatMonth: blank(), deptSubMonth: blank(),
      deptAcctMonth: blank(), catMonth: blank(), clientMonth: blank(),
      clientDeptMonth: blank(), subMonth: blank(),
    });
    var q = (E.quality = { rows: rows.length, noDept: 0, unknownAcct: 0,
                           noContact: 0, minDate: null, maxDate: null });

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var date = r[ci.date];
      var name = r[ci.account];
      var a = E.acct[name];
      var amount = Math.round((+r[ci.credit] || 0) * 100) - Math.round((+r[ci.debit] || 0) * 100);
      var dept = E.resolveDept(r[ci.costCentre], r[ci.jobNo]);
      var month = date.slice(0, 7);
      var cat = a ? a.cat : "unknown";
      var sub = a ? a.sub : "UNKNOWN";
      var contact = (r[ci.contact] || "").trim();

      gl[i] = { date: date, month: month, account: name, contact: contact,
                dept: dept, cat: cat, sub: sub, amount: amount,
                source: r[ci.source], jobNo: r[ci.jobNo],
                invoiceNo: r[ci.invoiceNo], costCentre: r[ci.costCentre],
                description: r[ci.description] };

      if (dept === "UNALLOCATED") q.noDept++;
      if (!a) q.unknownAcct++;
      if (!q.minDate || date < q.minDate) q.minDate = date;
      if (!q.maxDate || date > q.maxDate) q.maxDate = date;

      bump(idx.acctMonth, name, month, amount);
      bump(idx.catMonth, cat, month, amount);
      bump(idx.subMonth, sub, month, amount);
      bump2(idx.deptCatMonth, dept, cat, month, amount);
      bump2(idx.deptSubMonth, dept, sub, month, amount);
      bump2(idx.deptAcctMonth, dept, name, month, amount);
      if (cat === "income") {
        if (!contact) q.noContact++;
        else {
          bump(idx.clientMonth, contact, month, amount);
          bump2(idx.clientDeptMonth, contact, dept, month, amount);
        }
      }
    }
    E.glMonths = Object.keys(idx.catMonth.income || {}).sort();
    return E;
  };

  /* ---- the P&L paste and the budget ---------------------------------- */
  E.loadFin = function (data) {
    E.fin = data || { actual: {}, budget: {} };
    E.finActual = E.fin.actual || {};
    E.finBudget = E.fin.budget || {};
    E.control = (E.fin.meta && E.fin.meta.control) || {};
    // months that carry a posted actual, in order
    var seen = {};
    Object.keys(E.finActual).forEach(function (a) {
      Object.keys(E.finActual[a]).forEach(function (m) { seen[m] = 1; });
    });
    E.actualMonths = Object.keys(seen).sort();
    return E;
  };

  E.loadUtil = function (data) {
    var cols = data.cols, ci = {};
    cols.forEach(function (c, i) { ci[c] = i; });
    E.utilRows = (data.rows || []).map(function (r) {
      return { month: r[ci.month], dept: r[ci.dept], chargeable: +r[ci.chargeable],
               nonChargeable: +r[ci.nonChargeable], leave: +r[ci.leave],
               // Public holiday is its own bucket. Like leave it is paid but
               // not worked, so it is excluded from the utilisation
               // denominator rather than counted as non-chargeable.
               publicHoliday: ci.publicHoliday == null ? 0 : (+r[ci.publicHoliday] || 0),
               fte: +r[ci.fte] };
    });
    E.utilIdx = {};
    E.utilRows.forEach(function (r) {
      (E.utilIdx[r.dept] = E.utilIdx[r.dept] || {})[r.month] = r;
    });
    E.utilMeta = data.meta || {};
    return E;
  };

  /* ---- access ---------------------------------------------------------
   *
   * This decides which tabs a person is shown. It is not a security boundary
   * and the portal says so wherever it comes up: the whole thing is a file
   * running in the browser with no server behind it, so anyone who can open
   * the folder can read every data file whatever their role says. It is here
   * because showing a department head twenty two tabs they do not want is a
   * real problem worth solving, not because it locks anything.
   */
  E.loadUsers = function (data) {
    if (!data) {
      // No access file. Since none of this protects anything, the safe
      // failure is an open portal with a note, not one nobody can get into.
      data = { meta: { missing: true }, PAGES: [], USERS: [], DEFAULT_ROLE: "open",
               ROLES: [{ id: "open", label: "Everything", admin: true,
                         ownDeptOnly: false, pages: null,
                         note: "No access file was loaded." }] };
    }
    E.users = data;
    E.pageMeta = {};
    (E.users.PAGES || []).forEach(function (p) { E.pageMeta[p.id] = p; });
    return E;
  };

  /** Role page lists, with any change the finance head has made on top. */
  E.rolePages = function (roleId) {
    var r = (E.users.ROLES || []).filter(function (x) { return x.id === roleId; })[0];
    if (r && r.pages === null) return Object.keys(CTS.pages || {});   // the open fallback
    var over = store.get("accessRoles", {});
    if (over[roleId]) return over[roleId];
    return r ? r.pages.slice() : [];
  };
  E.setRolePages = function (roleId, pages) {
    var over = store.get("accessRoles", {});
    over[roleId] = pages;
    store.set("accessRoles", over);
  };
  E.role = function (id) {
    return (E.users.ROLES || []).filter(function (r) { return r.id === id; })[0] || null;
  };
  E.userList = function () {
    return store.get("accessUsers", null) || (E.users.USERS || []).slice();
  };
  E.setUserList = function (list) { store.set("accessUsers", list); };

  E.identity = function () { return store.get("identity", null); };
  E.setIdentity = function (name) {
    store.set("identity", name);
    store.set("previewRole", null);
  };
  E.currentUser = function () {
    var n = E.identity();
    if (!n) return null;
    return E.userList().filter(function (u) { return u.name === n; })[0] || null;
  };
  /** The role actually in force, which may be one the admin is previewing. */
  E.currentRole = function () {
    var preview = store.get("previewRole", null);
    if (preview && E.isRealAdmin()) return E.role(preview);
    var u = E.currentUser();
    return E.role(u ? u.role : E.users.DEFAULT_ROLE) || E.role(E.users.DEFAULT_ROLE);
  };
  E.isRealAdmin = function () {
    var u = E.currentUser();
    var r = u ? E.role(u.role) : null;
    return !!(r && r.admin);
  };
  E.isAdmin = function () {
    var r = E.currentRole();
    return !!(r && r.admin);
  };
  E.previewRole = function () { return store.get("previewRole", null); };
  E.setPreviewRole = function (id) { store.set("previewRole", id); };

  E.can = function (pageId) {
    var r = E.currentRole();
    if (!r) return false;
    return E.rolePages(r.id).indexOf(pageId) >= 0;
  };

  /** A department head sees their own department only. Everyone else sees all
   *  of them. Applied wherever a page lists departments. */
  E.deptScope = function () {
    var r = E.currentRole();
    if (!r || !r.ownDeptOnly) return null;
    var u = E.currentUser();
    return (u && u.dept) || null;
  };
  E.visibleDepts = function (list) {
    var scope = E.deptScope();
    list = list || E.depts;
    if (!scope) return list;
    return list.filter(function (d) { return d.code === scope; });
  };

  /* ---- keeping a load -------------------------------------------------
   *
   * A load used to last until the page was reloaded, and the only way to make
   * it stick was to download a file and drop it in portal/data. That does not
   * work everywhere: a page served in a sandboxed frame has downloads blocked
   * outright. So a load is kept in the browser instead, which needs no
   * download, survives a reload, and works the same whether the portal is
   * opened from a folder or from a link.
   *
   * It is per browser and per person. It is not shared, and it is not a
   * substitute for putting the file in portal/data when everyone should get it.
   */
  var SNAP = { gl: "snap.gl", fin: "snap.fin", util: "snap.util", staff: "snap.staff" };

  E.snapshot = function (kind, payload) {
    var key = SNAP[kind];
    if (!key) return { ok: false, why: "unknown data set" };
    try {
      localStorage.setItem("cts.bi." + key, JSON.stringify(
        { saved: new Date().toISOString(), payload: payload }));
      return { ok: true };
    } catch (err) {
      // usually the 5MB per origin quota, which a full year of ledger can pass
      return { ok: false, why: /quota/i.test(String(err && err.name || err))
        ? "it is too big for what the browser will keep, which is about five megabytes"
        : String(err && err.message || err) };
    }
  };

  E.snapshots = function () {
    var out = {};
    Object.keys(SNAP).forEach(function (k) {
      var raw = store.get(SNAP[k], null);
      if (raw && raw.payload) out[k] = raw;
    });
    return out;
  };

  E.clearSnapshots = function (kind) {
    Object.keys(SNAP).forEach(function (k) {
      if (kind && k !== kind) return;
      try { localStorage.removeItem("cts.bi." + SNAP[k]); } catch (e) {}
    });
  };

  E.restoreSnapshots = function () {
    var snaps = E.snapshots();
    var restored = [];
    if (snaps.gl) { E.loadGL(snaps.gl.payload); E._deptShare = null; restored.push("ledger"); }
    if (snaps.fin) { E.fin = snaps.fin.payload; E.loadFin(E.fin); restored.push("profit and loss"); }
    if (snaps.util) { E.loadUtil(snaps.util.payload); restored.push("utilisation"); }
    if (snaps.staff) { E.loadStaff(snaps.staff.payload); restored.push("people"); }
    if (restored.length) E.seed = false;
    E.restored = restored;
    E.restoredAt = (snaps.util || snaps.gl || snaps.fin || snaps.staff || {}).saved || null;
    return restored;
  };

  /* ---- pipeline: deals, orders, quotes ---------------------------------- */
  E.loadPipeline = function (data) {
    E.pipeline = data || { meta: {}, deals: [], orders: [], quotes: [] };
    E.pipeline.deals = E.pipeline.deals || []; E.pipeline.orders = E.pipeline.orders || [];
    E.pipeline.quotes = E.pipeline.quotes || [];
    E.hasPipeline = !!(E.pipeline.deals.length || E.pipeline.orders.length || E.pipeline.quotes.length);
    return E;
  };
  function inKeys(iso, keys) { return !!iso && keys.indexOf(String(iso).slice(0, 7)) >= 0; }
  E.pipelineFor = function (fyKeys, periodKeys) {
    var pl = E.pipeline || { deals: [], orders: [], quotes: [] };
    var won = /closed won|won/i, lost = /closed lost|lost/i;
    var open = pl.deals.filter(function (d) { return !won.test(d.stage) && !lost.test(d.stage); });
    var wonD = pl.deals.filter(function (d) { return won.test(d.stage) && inKeys(d.close, fyKeys); });
    var lostD = pl.deals.filter(function (d) { return lost.test(d.stage) && inKeys(d.close, fyKeys); });
    var byStage = {};
    open.forEach(function (d) {
      var b = byStage[d.stage] || (byStage[d.stage] = { stage: d.stage, n: 0, amount: 0, expected: 0 });
      b.n++; b.amount += d.amount || 0; b.expected += d.expected || 0;
    });
    var expectedByMonth = {}, wonByMonth = {};
    fyKeys.forEach(function (k) { expectedByMonth[k] = 0; wonByMonth[k] = 0; });
    open.forEach(function (d) { if (inKeys(d.close, fyKeys)) expectedByMonth[d.close.slice(0, 7)] += d.expected || 0; });
    wonD.forEach(function (d) { wonByMonth[d.close.slice(0, 7)] += d.amount || 0; });
    var orders = pl.orders.filter(function (o) { return inKeys(o.start, fyKeys); });
    var bookByMonthDept = {};
    orders.forEach(function (o) {
      if (!/confirm|complete/i.test(o.status)) return;
      var mk = o.start.slice(0, 7);
      (bookByMonthDept[mk] = bookByMonthDept[mk] || {})[o.dept] = (bookByMonthDept[mk][o.dept] || 0) + (o.value || 0);
    });
    var quotes = pl.quotes.filter(function (q) { return inKeys(q.sent, periodKeys || fyKeys); });
    var qByDept = {};
    quotes.forEach(function (q) {
      var b = qByDept[q.dept] || (qByDept[q.dept] = { dept: q.dept, sent: 0, sentValue: 0, accepted: 0, acceptedValue: 0 });
      b.sent++; b.sentValue += q.value || 0;
      if (/accept/i.test(q.status)) { b.accepted++; b.acceptedValue += q.value || 0; }
    });
    return {
      open: open, won: wonD, lost: lostD, byStage: Object.keys(byStage).map(function (k) { return byStage[k]; }),
      openAmount: open.reduce(function (a, d) { return a + (d.amount || 0); }, 0),
      openExpected: open.reduce(function (a, d) { return a + (d.expected || 0); }, 0),
      wonAmount: wonD.reduce(function (a, d) { return a + (d.amount || 0); }, 0),
      winRate: (wonD.length + lostD.length) ? wonD.length / (wonD.length + lostD.length) : null,
      expectedByMonth: expectedByMonth, wonByMonth: wonByMonth,
      orders: orders, bookByMonthDept: bookByMonthDept,
      quotes: quotes, quotesByDept: Object.keys(qByDept).map(function (k) { return qByDept[k]; }),
    };
  };

  /* ---- people ---------------------------------------------------------- */
  E.loadStaff = function (data) {
    var d = data || { cols: [], rows: [] };
    var ci = {};
    (d.cols || []).forEach(function (c, i) { ci[c] = i; });
    E.staffRows = (d.rows || []).map(function (r) {
      return { month: r[ci.month], empId: r[ci.empId], name: r[ci.name], dept: r[ci.dept],
               chargeable: +r[ci.chargeable] || 0, nonChargeable: +r[ci.nonChargeable] || 0,
               leave: +r[ci.leave] || 0, publicHoliday: +r[ci.publicHoliday] || 0,
               gross: +r[ci.grossCents] || 0, sup: +r[ci.superCents] || 0,
               employment: r[ci.employment] || null };
    });
    E.staffMeta = d.meta || {};
    E.hasStaff = E.staffRows.length > 0;
    return E;
  };

  /** Profitability per person.
   *
   *  Revenue is not billed per person anywhere, so it is attributed: each
   *  department's revenue for the period is spread across its people on their
   *  share of that department's chargeable hours. That is a fair way to read
   *  who is carrying the work, and it is not the same as what someone earned.
   *  Anyone with no chargeable hours receives none, which is why support staff
   *  show a loss and should be read as a cost, not a performance.
   */
  E.staffFor = function (monthKeys) {
    var byEmp = {}, deptChg = {}, deptRev = {};
    var keys = {};
    monthKeys.forEach(function (k) { keys[k] = 1; });

    E.staffRows.forEach(function (r) {
      if (!keys[r.month]) return;
      var e = byEmp[r.empId] || (byEmp[r.empId] = {
        empId: r.empId, name: r.name, dept: r.dept, employment: r.employment,
        chargeable: 0, nonChargeable: 0, leave: 0, publicHoliday: 0, gross: 0, sup: 0,
        months: {} });
      e.chargeable += r.chargeable; e.nonChargeable += r.nonChargeable;
      e.leave += r.leave; e.publicHoliday += r.publicHoliday;
      e.gross += r.gross; e.sup += r.sup;
      e.months[r.month] = 1;
      if (r.employment) e.employment = r.employment;
      deptChg[r.dept] = (deptChg[r.dept] || 0) + r.chargeable;
    });

    E.postingDepts.forEach(function (d) {
      deptRev[d.code] = E.sumMonths((E.idx.deptCatMonth[d.code] || {}).income, monthKeys);
    });

    var rows = Object.keys(byEmp).map(function (id) {
      var e = byEmp[id];
      var worked = e.chargeable + e.nonChargeable;
      var cost = e.gross + e.sup;
      var share = deptChg[e.dept] ? e.chargeable / deptChg[e.dept] : 0;
      var revenue = Math.round((deptRev[e.dept] || 0) * share);
      var margin = revenue - cost;
      return {
        empId: e.empId, name: e.name, dept: e.dept, deptShort: (E.deptOf[e.dept] || {}).short || e.dept,
        employment: e.employment, months: Object.keys(e.months).length,
        chargeable: e.chargeable, nonChargeable: e.nonChargeable,
        leave: e.leave, publicHoliday: e.publicHoliday, worked: worked,
        util: worked ? e.chargeable / worked : null,
        gross: e.gross, sup: e.sup, cost: cost,
        revenue: revenue, margin: margin,
        marginPct: revenue ? margin / revenue : null,
        recovery: cost ? revenue / cost : null,
        chargeRate: e.chargeable ? Math.round(revenue / e.chargeable) : null,
        costRate: worked ? Math.round(cost / worked) : null,
        billable: e.chargeable > 0,
      };
    });
    rows.sort(function (a, b) { return b.margin - a.margin; });
    return { rows: rows, deptRevenue: deptRev, deptChargeable: deptChg };
  };

  /* ---- period helpers ------------------------------------------------ */
  E.monthsOfFY = function (fy) {
    return E.months.filter(function (m) { return m.fy === fy; }).map(function (m) { return m.key; });
  };
  E.reportingMonth = function () {
    return store.get("reportingMonth", E.cfg.ORG.reportingMonth);
  };
  E.setReportingMonth = function (k) { store.set("reportingMonth", k); E._fc = null; E._cash = null; };
  E.currentFY = function () {
    var m = E.monthIdx[E.reportingMonth()];
    return m ? m.fy : E.cfg.ORG.currentFY;
  };
  E.ytd = function (fy) {
    var rm = E.reportingMonth();
    return E.monthsOfFY(fy || E.currentFY()).filter(function (k) { return k <= rm; });
  };
  E.quarterMonths = function () {
    var m = E.monthIdx[E.reportingMonth()];
    if (!m) return [];
    var q = m.quarter;
    return E.months.filter(function (x) { return x.fy === m.fy && x.quarter === q && x.key <= m.key; })
      .map(function (x) { return x.key; });
  };
  E.priorYearMonths = function (keys) {
    return keys.map(function (k) {
      var y = +k.slice(0, 4) - 1;
      return y + k.slice(4);
    });
  };
  E.hasActual = function (monthKey) {
    return E.actualMonths.indexOf(monthKey) >= 0 && monthKey <= E.reportingMonth();
  };

  /* ---- summing ------------------------------------------------------- */
  function sumMonths(map, keys) {
    if (!map) return 0;
    var t = 0;
    for (var i = 0; i < keys.length; i++) t += map[keys[i]] || 0;
    return t;
  }
  E.sumMonths = sumMonths;

  E.acctActual = function (name, keys) { return sumMonths(E.finActual[name], keys); };
  E.acctBudget = function (name, keys) { return sumMonths(E.finBudget[name], keys); };

  /** A P&L month cell: actual up to the reporting month, budget after it.
   *  This is the Controller Pack's rule, so the twelve months always read as a
   *  full year with the shortfall ahead visible. */
  E.acctBlend = function (name, monthKey) {
    if (E.hasActual(monthKey)) return (E.finActual[name] || {})[monthKey] || 0;
    if (E.hasForecast && E.hasForecast(monthKey)) return E.acctForecast(name, monthKey);
    return (E.finBudget[name] || {})[monthKey] || 0;
  };
  /** Where a blended month comes from: actual to the reporting month, then
   *  the forecast where there is one, then budget. */
  E.blendSource = function (monthKey) {
    if (E.hasActual(monthKey)) return "actual";
    if (E.hasForecast && E.hasForecast(monthKey)) return "forecast";
    return "budget";
  };
  E.blendLabel = function () { return E.forecastEnabled && E.forecastEnabled() ? "actual then forecast" : "actual then budget"; };

  /* ---- the P&L ------------------------------------------------------- */
  var CAT_KEYS = ["income", "cos", "expenses", "other_income", "other_expenses"];
  var CAT_LABEL = { income: "Income", cos: "Cost of Sales", expenses: "Expenses",
                    other_income: "Other Income", other_expenses: "Other Expenses" };
  E.CAT_KEYS = CAT_KEYS; E.CAT_LABEL = CAT_LABEL;

  /** Build the three level P&L: category, subcategory, account.
   *  `pick(accountName, monthKey)` supplies the cents for one cell, so the same
   *  builder serves actual, budget and the blended forecast view. */
  E.buildPnl = function (monthKeys, pick) {
    var cats = [];
    CAT_KEYS.forEach(function (cat) {
      var subs = E.subsOf[cat] || {};
      var subRows = [];
      Object.keys(subs).sort().forEach(function (sub) {
        var accRows = subs[sub].slice().sort().map(function (name) {
          var cells = monthKeys.map(function (m) { return pick(name, m); });
          return { kind: "account", label: name, account: name, cells: cells,
                   total: cells.reduce(function (a, b) { return a + b; }, 0) };
        });
        var live = accRows.filter(function (r) { return r.total !== 0 || r.cells.some(function (c) { return c !== 0; }); });
        var cells = monthKeys.map(function (m, i) {
          return accRows.reduce(function (a, r) { return a + r.cells[i]; }, 0);
        });
        subRows.push({ kind: "sub", label: sub, sub: sub, cells: cells,
                       total: cells.reduce(function (a, b) { return a + b; }, 0),
                       children: accRows, liveChildren: live.length });
      });
      var live = subRows.filter(function (r) { return r.total !== 0; });
      var cells = monthKeys.map(function (m, i) {
        return subRows.reduce(function (a, r) { return a + r.cells[i]; }, 0);
      });
      cats.push({ kind: "cat", key: cat, label: CAT_LABEL[cat], cells: cells,
                  total: cells.reduce(function (a, b) { return a + b; }, 0),
                  children: live.length ? live : subRows, liveChildren: live.length });
    });
    var by = {};
    cats.forEach(function (c) { by[c.key] = c; });
    function vec(key) { return by[key] ? by[key].cells : monthKeys.map(function () { return 0; }); }
    function add() {
      var args = arguments;
      return monthKeys.map(function (_, i) {
        var t = 0;
        for (var a = 0; a < args.length; a++) t += args[a][i];
        return t;
      });
    }
    var income = vec("income"), cos = vec("cos"), exp = vec("expenses");
    var oi = vec("other_income"), oe = vec("other_expenses");
    var gp = add(income, cos);
    var np = add(gp, oi, exp, oe);
    function tot(v) { return v.reduce(function (a, b) { return a + b; }, 0); }
    return {
      months: monthKeys, cats: cats, by: by,
      income: income, cos: cos, expenses: exp, otherIncome: oi, otherExpenses: oe,
      grossProfit: gp, netProfit: np,
      totals: { income: tot(income), cos: tot(cos), expenses: tot(exp),
                otherIncome: tot(oi), otherExpenses: tot(oe),
                grossProfit: tot(gp), netProfit: tot(np) },
      gmPct: gp.map(function (v, i) { return income[i] ? v / income[i] : null; }),
      npPct: np.map(function (v, i) { return income[i] ? v / income[i] : null; }),
    };
  };

  /** The department in view: a department head's own, else the one picked
   *  at the top of the page, else none (the company). When one is in view
   *  the P&L functions read the ledger by department instead of the company
   *  P&L paste, the budget on each account's department split, and the
   *  forecast by department, so every page becomes that department's. */
  E.viewDept = function () {
    var scope = E.deptScope(); if (scope) return scope;
    if (E._noView) return null;
    var f = store.get("focusDept", null);
    return f && E.deptOf[f] ? f : null;
  };
  E.withoutView = function (fn) { var was = E._noView; E._noView = true; try { return fn(); } finally { E._noView = was; } };
  function deptActual(d) { return function (n, m) { return (((E.idx.deptAcctMonth[d] || {})[n] || {})[m]) || 0; }; }
  function deptBudget(d) { return function (n, m) { var b = (E.finBudget[n] || {})[m] || 0; return b ? Math.round(b * (E.deptShareOfAccount(n)[d] || 0)) : 0; }; }
  function deptForecast(d) { return function (n, m) { var f = E.forecast ? E.forecast() : null; return f ? ((((f.byDeptAcct[d] || {})[n] || {})[m]) || 0) : 0; }; }
  E.pnlActual = function (keys) {
    var d = E.viewDept();
    return E.buildPnl(keys, d ? deptActual(d) : function (n, m) { return (E.finActual[n] || {})[m] || 0; });
  };
  E.pnlBudget = function (keys) {
    var d = E.viewDept();
    return E.buildPnl(keys, d ? deptBudget(d) : function (n, m) { return (E.finBudget[n] || {})[m] || 0; });
  };
  E.pnlBlend = function (keys) {
    var d = E.viewDept();
    if (!d) return E.buildPnl(keys, E.acctBlend);
    var a = deptActual(d), b = deptBudget(d), f = deptForecast(d);
    return E.buildPnl(keys, function (n, m) { var src = E.blendSource(m); return src === "actual" ? a(n, m) : src === "forecast" ? f(n, m) : b(n, m); });
  };
})();

/* ===================================================================== *
 * Engine part 2: departmental P&L, overhead allocation, utilisation,
 * clients and the revenue schedule.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, E = CTS.engine;

  /* ---- exact cent splitting ------------------------------------------ */
  function splitExact(total, fracs) {
    var raw = fracs.map(function (f) { return total * f; });
    var out = raw.map(function (v) { return Math.floor(v); });
    var short = total - out.reduce(function (a, b) { return a + b; }, 0);
    var order = raw.map(function (v, i) { return i; }).sort(function (a, b) {
      return (raw[b] - out[b]) - (raw[a] - out[a]);
    });
    for (var i = 0; i < Math.abs(short); i++) {
      out[order[i % order.length]] += short > 0 ? 1 : -1;
    }
    return out;
  }
  E.splitExact = splitExact;

  /* ---- the three split bases, expanded to six departments ------------ */
  E.basis = function (code) {
    var b = E.cfg.SPLIT_BASES.filter(function (x) { return x.code === code; })[0];
    return b || E.cfg.SPLIT_BASES[0];
  };

  /** A basis is set against three budget groups, not six departments. The two
   *  sub splits break them out: production takes 0.79 of production and video,
   *  integration takes 0.72 of consulting and integration. */
  E.basisWeights = function (code) {
    var b = E.basis(code), g = b.groups, s = E.subSplits();
    var w = {};
    w.ONSITE = g["ONS"] || 0;
    w.PRODUCTION = (g["PRD/VID"] || 0) * s.prd;
    w.VIDEO = (g["PRD/VID"] || 0) * (1 - s.prd);
    w.INTEGRATION = (g["CONS/INT"] || 0) * s.int;
    w.CONSULTING = (g["CONS/INT"] || 0) * (1 - s.int);
    w.ADMIN = 0;
    return w;
  };
  E.subSplits = function () {
    var d = E.cfg.SUB_SPLITS;
    return {
      prd: store.get("subsplit.prd", d.productionShareOfPrdVid),
      int: store.get("subsplit.int", d.integrationShareOfConsInt),
    };
  };
  var store = CTS.store;

  E.basisFor = function (sub) {
    var a = E.cfg.SPLIT_ASSIGNMENT;
    var over = store.get("splitAssign", {});
    return over[sub] || a.bySub[sub] || a.default;
  };
  E.setBasisFor = function (sub, code) {
    var over = store.get("splitAssign", {});
    over[sub] = code;
    store.set("splitAssign", over);
  };

  /* ---- departmental P&L, before and after allocation ----------------- */
  var ALLOC_TARGETS = ["ONSITE", "PRODUCTION", "VIDEO", "INTEGRATION", "CONSULTING"];

  /** Raw departmental result, straight off the ledger, nothing pushed out. */
  E.deptRaw = function (monthKeys) {
    var out = {};
    E.depts.forEach(function (d) {
      var cat = E.idx.deptCatMonth[d.code] || {};
      var row = { code: d.code, dept: d };
      E.CAT_KEYS.forEach(function (k) {
        row[k] = E.sumMonths(cat[k], monthKeys);
      });
      row.grossProfit = row.income + row.cos;
      row.netProfit = row.grossProfit + row.other_income + row.expenses + row.other_expenses;
      out[d.code] = row;
    });
    return out;
  };

  /** Push the Admin overhead pool out across the departments.
   *
   *  A department's allocated result is its OWN overhead plus its share of the
   *  pool, kept as two separate numbers. The Monthly Reporting manual records a
   *  real bug where a departmental tab picked up its own overhead twice by
   *  referencing a cell that already contained it; keeping the two components
   *  apart is what makes that impossible to do silently here.
   */
  E.allocate = function (monthKeys) {
    var raw = E.deptRaw(monthKeys);
    var subs = E.idx.deptSubMonth["ADMIN"] || {};
    var workings = [];
    var received = {};
    ALLOC_TARGETS.forEach(function (d) { received[d] = 0; });

    Object.keys(subs).sort().forEach(function (sub) {
      var amount = E.sumMonths(subs[sub], monthKeys);
      if (!amount) return;
      var names = (E.subsOf["expenses"] || {})[sub];
      if (!names) return;                       // only expenses are pushed out
      var code = E.basisFor(sub);
      var w = E.basisWeights(code);
      var fr = ALLOC_TARGETS.map(function (d) { return w[d] || 0; });
      var tot = fr.reduce(function (a, b) { return a + b; }, 0) || 1;
      var parts = splitExact(amount, fr.map(function (f) { return f / tot; }));
      var split = {};
      ALLOC_TARGETS.forEach(function (d, i) {
        split[d] = parts[i];
        received[d] += parts[i];
      });
      workings.push({ sub: sub, basis: code, basisLabel: E.basis(code).label,
                      basisSource: E.basis(code).source, amount: amount, split: split });
    });

    var pool = workings.reduce(function (a, w) { return a + w.amount; }, 0);
    var rows = E.depts.map(function (d) {
      var r = raw[d.code];
      var own = r.expenses;
      var share = d.code === "ADMIN" ? -pool : (received[d.code] || 0);
      var expenses = own + share;
      var gp = r.grossProfit;
      var np = gp + r.other_income + expenses + r.other_expenses;
      return {
        code: d.code, dept: d, income: r.income, cos: r.cos, grossProfit: gp,
        ownExpenses: own, allocated: share, expenses: expenses,
        otherIncome: r.other_income, otherExpenses: r.other_expenses, netProfit: np,
        gmPct: r.income ? gp / r.income : null,
        npPct: r.income ? np / r.income : null,
      };
    });
    return { rows: rows, byCode: rows.reduce(function (m, r) { m[r.code] = r; return m; }, {}),
             workings: workings, pool: pool, raw: raw };
  };

  /** Anything that depends on a placeholder split basis is flagged, so a
   *  number that rests on a figure we do not actually have says so. */
  E.allocationIsSourced = function () {
    var used = {};
    Object.keys((E.idx.deptSubMonth["ADMIN"] || {})).forEach(function (sub) {
      if ((E.subsOf["expenses"] || {})[sub]) used[E.basisFor(sub)] = true;
    });
    var bad = Object.keys(used).filter(function (c) { return E.basis(c).source !== "confirmed"; });
    return { ok: bad.length === 0, placeholders: bad.map(function (c) { return E.basis(c).label; }) };
  };

  /* ---- budget by department ------------------------------------------ */
  /** The budget is held per account, and most accounts carry no department in
   *  their name. Each account's budget is therefore spread across departments
   *  on that account's own prior year ledger split. Derived, not sourced. */
  E.deptShareOfAccount = function (name) {
    var cached = E._deptShare || (E._deptShare = {});
    if (cached[name]) return cached[name];
    var tot = 0, by = {};
    E.postingDepts.forEach(function (d) {
      var v = 0, m = (E.idx.deptAcctMonth[d.code] || {})[name] || {};
      Object.keys(m).forEach(function (k) { v += m[k]; });
      if (v) { by[d.code] = v; tot += v; }
    });
    var out = {};
    if (tot) Object.keys(by).forEach(function (d) { out[d] = by[d] / tot; });
    else {
      var a = E.acct[name];
      out[(a && a.dept) || "ADMIN"] = 1;
    }
    return (cached[name] = out);
  };

  E.budgetByDept = function (monthKeys, cat) {
    var out = {};
    E.depts.forEach(function (d) { out[d.code] = 0; });
    Object.keys(E.finBudget).forEach(function (name) {
      var a = E.acct[name];
      if (!a || (cat && a.cat !== cat)) return;
      var amt = E.sumMonths(E.finBudget[name], monthKeys);
      if (!amt) return;
      var share = E.deptShareOfAccount(name);
      Object.keys(share).forEach(function (d) {
        out[d] = (out[d] || 0) + amt * share[d];
      });
    });
    Object.keys(out).forEach(function (d) { out[d] = Math.round(out[d]); });
    return out;
  };

  /* ---- risk flag ------------------------------------------------------ */
  /** Exactly the Controller Pack's rule, including the order of the tests: the
   *  materiality floor is checked BEFORE the percentage, so a large percentage
   *  on a small dollar variance is still Low. */
  E.risk = function (actual, budget) {
    var R = E.cfg.RISK;
    if (actual === 0 && budget === 0) return "No Activity";
    var varD = Math.abs(actual - budget) / 100;
    var varP = budget ? Math.abs((actual - budget) / budget) : (actual ? 1 : 0);
    if (varD < R.materialityFloor) return "Low";
    if (varD >= R.highDollar || varP >= R.highPct) return "High";
    if (varP >= R.mediumPct) return "Medium";
    return "Low";
  };
  E.riskRank = { "High": 3, "Medium": 2, "Low": 1, "No Activity": 0 };

  /* ---- utilisation ---------------------------------------------------- */
  E.utilState = function () { return store.get("utilState", E.cfg.UTIL.state); };
  E.hoursPerDay = function () { return store.get("hoursPerDay", E.cfg.UTIL.hoursPerDay); };

  E.workingDays = function (monthKeys, state) {
    state = state || E.utilState();
    return monthKeys.reduce(function (a, k) {
      var m = E.monthIdx[k];
      return a + (m ? (m.wd[state] || 0) : 0);
    }, 0);
  };
  /** Weekdays only, no holidays deducted. A full year of these times eight is
   *  the 2,080 or 2,088 hours Part 2 says to check the denominator against.
   *  The two counts are deliberately kept apart: the pack nets holidays off,
   *  the graphs file does not, which is one reason the two utilisation figures
   *  in the business never agree. */
  E.businessDays = function (monthKeys) {
    return monthKeys.reduce(function (a, k) {
      var m = E.monthIdx[k];
      return a + (m ? (m.bd || 0) : 0);
    }, 0);
  };

  E.utilFor = function (monthKeys) {
    var wd = E.workingDays(monthKeys);
    var perHead = wd * E.hoursPerDay();
    var rows = E.postingDepts.map(function (d) {
      var chg = 0, non = 0, lv = 0, ph = 0;
      monthKeys.forEach(function (k) {
        var r = (E.utilIdx[d.code] || {})[k];
        if (r) {
          chg += r.chargeable; non += r.nonChargeable; lv += r.leave;
          ph += r.publicHoliday || 0;
        }
      });
      var worked = chg + non;
      return {
        code: d.code, dept: d, chargeable: chg, nonChargeable: non, leave: lv,
        publicHoliday: ph, worked: worked, total: worked + lv + ph,
        util: worked ? chg / worked : null,
        fte: perHead ? worked / perHead : null,
        target: d.utilTarget,
        gap: worked && d.utilTarget != null ? chg / worked - d.utilTarget : null,
      };
    });
    var view = E.viewDept();
    if (view) rows = rows.filter(function (r) { return r.code === view; });
    var tc = rows.reduce(function (a, r) { return a + r.chargeable; }, 0);
    var tn = rows.reduce(function (a, r) { return a + r.nonChargeable; }, 0);
    var tw = rows.reduce(function (a, r) { return a + r.worked; }, 0);
    var tl = rows.reduce(function (a, r) { return a + r.leave; }, 0);
    var tp = rows.reduce(function (a, r) { return a + (r.publicHoliday || 0); }, 0);
    return {
      rows: rows, workingDays: wd, hoursPerHead: perHead, state: E.utilState(),
      total: { chargeable: tc, nonChargeable: tn, worked: tw, leave: tl,
               publicHoliday: tp, total: tw + tl + tp,
               util: tw ? tc / tw : null, fte: perHead ? tw / perHead : null },
    };
  };

  /** The graphs file method: a rolling twelve months, checked against the
   *  2,080 or 2,088 hour year that Part 2 says to verify rather than assume. */
  E.utilRolling = function (endMonth) {
    var all = E.months.map(function (m) { return m.key; });
    var i = all.indexOf(endMonth);
    if (i < 0) i = all.length - 1;
    var keys = all.slice(Math.max(0, i - 11), i + 1);
    var u = E.utilFor(keys);
    u.months = keys;
    u.complete = keys.length === 12;
    u.businessDays = E.businessDays(keys);
    u.annualHours = u.businessDays * E.hoursPerDay();
    u.annualCheck = E.cfg.UTIL.annualHoursCheck.indexOf(Math.round(u.annualHours)) >= 0;
    // the graphs file counts every logged hour, leave included, against the
    // weekday year; the pack counts worked hours against the holiday adjusted
    // month. Both are shown, because both are in use.
    u.rows.forEach(function (r) {
      r.fteRolling = u.annualHours ? r.total / u.annualHours : null;
    });
    u.total.fteRolling = u.annualHours
      ? u.rows.reduce(function (a, r) { return a + r.total; }, 0) / u.annualHours : null;
    return u;
  };

  /* ---- clients -------------------------------------------------------- */
  E.clientRows = function (monthKeys, priorKeys) {
    var names = Object.keys(E.idx.clientMonth);
    var rows = names.map(function (c) {
      var meta = E.clientMeta[c] || {};
      var byDept = {};
      var d = E.idx.clientDeptMonth[c] || {};
      Object.keys(d).forEach(function (k) { byDept[k] = E.sumMonths(d[k], monthKeys); });
      var amount = E.sumMonths(E.idx.clientMonth[c], monthKeys);
      var prior = priorKeys ? E.sumMonths(E.idx.clientMonth[c], priorKeys) : 0;
      return {
        contact: c, display: meta.display || c, industry: meta.industry || "Other",
        engagement: meta.engagement || "Other", primaryDept: meta.primaryDept || null,
        amount: amount, prior: prior, delta: amount - prior,
        pct: prior ? (amount - prior) / Math.abs(prior) : null,
        byDept: byDept, mapped: !!E.clientMeta[c],
      };
    });
    rows.sort(function (a, b) { return b.amount - a.amount; });
    rows.forEach(function (r, i) { r.rank = i + 1; });
    return rows;
  };

  /* ---- the revenue schedule ------------------------------------------- */
  /** Part 3: onsite is budgeted per role, the rest per client, with anything
   *  unlisted grouped into one new business line. The budget per line is
   *  derived: the department's budgeted income for the period, spread across
   *  its lines on each line's share of the prior year. */
  E.schedule = function (dept, monthKeys) {
    var lines = (E.clients.schedule || []).filter(function (s) { return s.dept === dept; });
    // Shares come off the whole prior financial year, not the same months a
    // year earlier: two months of last year is too thin a base to spread a
    // budget on, and a client that happened to bill nothing in those two
    // months would otherwise be budgeted at zero.
    var thisFY = (E.monthIdx[monthKeys[monthKeys.length - 1]] || {}).fy || E.currentFY();
    var priorKeys = E.monthsOfFY(thisFY - 1);
    var budgetTotal = E.budgetByDept(monthKeys, "income")[dept] || 0;

    var priors = lines.map(function (l) {
      if (l.kind === "newbiz") return 0;
      var d = E.idx.clientDeptMonth[l.client] || {};
      return E.sumMonths(d[dept], priorKeys);
    });
    // roles share one client's prior year evenly between that client's roles
    var roleCount = {};
    lines.forEach(function (l) { if (l.kind === "role") roleCount[l.client] = (roleCount[l.client] || 0) + 1; });
    priors = priors.map(function (p, i) {
      var l = lines[i];
      return l.kind === "role" ? p / (roleCount[l.client] || 1) : p;
    });
    var priorSum = priors.reduce(function (a, b) { return a + b; }, 0);
    // the new business line carries what the named lines do not
    var NEWBIZ = 0.08;

    var listed = [];
    lines.forEach(function (l, i) {
      var share = priorSum ? (priors[i] / priorSum) * (1 - NEWBIZ) : 0;
      if (l.kind === "newbiz") share = NEWBIZ;
      var budget = Math.round(budgetTotal * share);
      var actual;
      if (l.kind === "newbiz") actual = null;      // filled after the named lines
      else {
        var d = E.idx.clientDeptMonth[l.client] || {};
        actual = E.sumMonths(d[dept], monthKeys);
        if (l.kind === "role") actual = Math.round(actual / (roleCount[l.client] || 1));
      }
      listed.push({ kind: l.kind, line: l.line, client: l.client,
                    budget: budget, actual: actual, prior: Math.round(priors[i]) });
    });

    var deptActual = E.sumMonths((E.idx.deptCatMonth[dept] || {}).income, monthKeys);
    var named = listed.reduce(function (a, r) { return a + (r.actual || 0); }, 0);
    listed.forEach(function (r) { if (r.kind === "newbiz") r.actual = deptActual - named; });

    listed.forEach(function (r) {
      r.variance = r.actual - r.budget;
      r.pct = r.budget ? r.variance / Math.abs(r.budget) : null;
      r.risk = E.risk(r.actual, r.budget);
    });
    return {
      dept: dept, lines: listed, budgetTotal: budgetTotal, actualTotal: deptActual,
      variance: deptActual - budgetTotal,
      priorMonths: priorKeys,
      note: "Budget per line is derived: the department's budgeted income for the period, spread on each line's share of the whole prior financial year, with 8 per cent held back for new business. The real schedule is set per role for onsite and per client elsewhere, and is maintained by hand.",
    };
  };

  /* ---- the P&L control ------------------------------------------------ */
  /** PL_Check: does the ledger agree with the P&L that was pasted in.
   *  Under $1 reads OK, the same tolerance the Controller Pack uses. */
  E.plCheck = function (monthKeys) {
    return E.CAT_KEYS.map(function (cat) {
      var gl = 0;
      E.postingDepts.concat([{ code: "UNALLOCATED" }]).forEach(function (d) {
        gl += E.sumMonths((E.idx.deptCatMonth[d.code] || {})[cat], monthKeys);
      });
      var pl = 0;
      Object.keys(E.finActual).forEach(function (n) {
        var a = E.acct[n];
        if (a && a.cat === cat) pl += E.sumMonths(E.finActual[n], monthKeys);
      });
      return { cat: cat, label: E.CAT_LABEL[cat], gl: gl, pl: pl,
               diff: gl - pl, ok: Math.abs(gl - pl) < 100 };
    });
  };
})();

/* ===================================================================== *
 * Engine part 3: the forecast. A third series beside actual and budget.
 *
 * Every forecast cell is made by one method, chosen per account on the
 * 09 Forecast template (account, then subcategory, then category, then a
 * default by category), and every cell knows which method made it. Typed
 * overrides always win and are flagged as typed. The history a method
 * reads is the ledger by department, so the departmental forecasts add up
 * to the company one the same way the actuals do.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, E = CTS.engine;

  var METHODS = {
    seasonal:    { label: "Same month last year, grown",      short: "seasonal",     param: "growth, as a fraction (0.03 is 3%)" },
    runrate:     { label: "Run rate of recent months",        short: "run rate",     param: "months to average" },
    revenue_pct: { label: "Percentage of department revenue", short: "% of revenue", param: "fraction of revenue (0.35 is 35%); blank measures it from the last twelve months" },
    fixed:       { label: "Fixed amount each month",          short: "fixed",        param: "dollars per month, positive as the P&L shows it" },
    budget:      { label: "Budget as the forecast",           short: "budget",       param: "none" },
    zero:        { label: "Nothing expected",                 short: "zero",         param: "none" },
    pipeline:    { label: "Booked work, weighted deals and the baseline", short: "pipeline", param: "set on the Pipeline tab" },
  };
  E.FC_METHODS = METHODS;
  var DEFAULTS = { income: "seasonal", cos: "revenue_pct", expenses: "runrate", other_income: "runrate", other_expenses: "runrate" };
  // cost of sales lines that follow time, not revenue: permanent labour
  var LABOUR = /salar|wage|superann|workers/i;

  E.loadForecast = function (data) { E.fcData = data || null; E._fc = null; };

  function assumptions() {
    var a = (E.fcData && E.fcData.assumptions) || {};
    return { horizon: Math.max(1, +a.horizonMonths || 12), runRate: Math.max(1, +a.runRateMonths || 3),
             growth: a.growth || {}, enabled: !!E.fcData && a.enabled !== false };
  }
  E.forecastAssumptions = assumptions;
  E.forecastEnabled = function () { return assumptions().enabled; };

  function catOf(label) {
    var l = String(label).toLowerCase().replace(/[^a-z]/g, "");
    return { income: "income", revenue: "income", costofsales: "cos", cos: "cos", directcosts: "cos",
             expenses: "expenses", overheads: "expenses", otherincome: "other_income",
             otherexpenses: "other_expenses" }[l] || l;
  }
  E.methodKey = function (label) {
    var l = String(label || "").toLowerCase();
    if (!l) return null;
    if (METHODS[l]) return l;
    if (/season|same month|last year/.test(l)) return "seasonal";
    if (/run ?rate|average|trailing/.test(l)) return "runrate";
    if (/revenue|% ?of|percent/.test(l)) return "revenue_pct";
    if (/fixed|contract|flat/.test(l)) return "fixed";
    if (/budget/.test(l)) return "budget";
    if (/zero|none|nothing/.test(l)) return "zero";
    return null;
  };

  /** The rule for one account: the most specific line on the template wins. */
  E.forecastRule = function (name) {
    var a = E.acct[name] || {}, rules = (E.fcData && E.fcData.methods) || [];
    var hit = null, rank = 0;
    rules.forEach(function (r) {
      var m = String(r.match || "").trim(), method = E.methodKey(r.method);
      if (!m || !method) return;
      var score = 0;
      if (m.toLowerCase() === String(name).toLowerCase()) score = 3;
      else if (/^sub(category)?\s*:/i.test(m) && m.replace(/^sub(category)?\s*:\s*/i, "").toLowerCase() === String(a.sub || "").toLowerCase()) score = 2;
      else if (/^cat(egory)?\s*:/i.test(m) && catOf(m.replace(/^cat(egory)?\s*:\s*/i, "")) === a.cat) score = 1;
      if (score > rank) { rank = score; hit = Object.assign({}, r, { method: method }); }
    });
    if (hit) {
      var pv = hit.param === "" || hit.param == null || isNaN(+hit.param) ? null : +hit.param;
      return { method: hit.method, param: pv, source: rank === 3 ? "account" : rank === 2 ? "subcategory" : "category", note: hit.note || "" };
    }
    var def = DEFAULTS[a.cat] || "runrate";
    if (a.cat === "cos" && LABOUR.test(a.sub || "")) def = "runrate";
    return { method: def, param: null, source: "default", note: "" };
  };

  /** Months the forecast covers: after the reporting month, for the horizon,
   *  and never short of the end of the current financial year. */
  E.forecastMonths = function () {
    var rm = E.reportingMonth(), a = assumptions();
    var start = E.monthIdx[rm] ? E.monthIdx[rm].i : -1;
    var fy = E.monthsOfFY(E.currentFY()), fyLast = fy[fy.length - 1] || rm;
    var out = [];
    for (var k = start + 1; k < E.months.length; k++) {
      var key = E.months[k].key;
      if (out.length >= a.horizon && key > fyLast) break;
      out.push(key);
    }
    return out;
  };
  E.hasForecast = function (monthKey) {
    if (!E.forecastEnabled()) return false;
    return build().monthSet[monthKey] === 1;
  };

  /* ---- the pipeline layer on income ------------------------------------
   * Confirmed work (orders with an event date, accepted quotes), the open
   * deals weighted by probability, and the baseline for the revenue that
   * never goes through a system. Near months are what is booked plus the
   * never-in-pipeline share; far months never fall below the baseline. */
  function pipelineSettings() {
    var p = ((E.fcData && E.fcData.assumptions) || {}).pipeline || {};
    var never = Object.assign({ "default": 0.3, ADMIN: 1, UNALLOCATED: 1 }, p.never || {});
    return {
      enabled: p.enabled !== false && !!E.hasPipeline,
      near: Math.max(0, +p.nearMonths >= 0 ? +p.nearMonths : 3),
      cancel: Math.max(0, Math.min(1, +p.cancellationRate || 0)),
      never: never, lead: Object.assign({ "default": 0 }, p.leadMonths || {}),
      orderStatuses: p.orderStatuses || "Confirmed, Booked, In progress, Completed",
      quoteStatuses: p.quoteStatuses || "Accepted",
      probability: p.probability || {},
      countWon: !!p.countWon,
    };
  }
  E.pipelineSettings = pipelineSettings;
  function listRe(s) { var parts = String(s).split(/[,;]/).map(function (x) { return x.trim(); }).filter(Boolean); return parts.length ? new RegExp("^(" + parts.map(function (x) { return x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }).join("|") + ")$", "i") : /^$/; }
  function shift(iso, n) { if (!iso) return null; var y = +iso.slice(0, 4), m = +iso.slice(5, 7) - 1 + n; return (y + Math.floor(m / 12)) + "-" + ((((m % 12) + 12) % 12) + 1 < 10 ? "0" : "") + ((((m % 12) + 12) % 12) + 1); }
  function norm(s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }

  /** What the systems say lands in each month, by department. */
  E.pipelineTargets = function (months) {
    var S = pipelineSettings(), pl = E.pipeline || { deals: [], orders: [], quotes: [] };
    var set = {}; months.forEach(function (m) { set[m] = 1; });
    var out = {}, seen = {}, wonUncounted = [], counted = { orders: 0, quotes: 0, deals: 0, dupQuotes: 0 };
    function slot(dept, m) { return ((out[dept] = out[dept] || {})[m] = out[dept][m] || { confirmed: 0, weighted: 0, items: [], deals: [] }); }
    var okOrder = listRe(S.orderStatuses), okQuote = listRe(S.quoteStatuses);
    var won = /closed won|^won$/i, lost = /closed lost|^lost$/i;
    function leadFor(dept) { return S.lead[dept] != null ? +S.lead[dept] : +S.lead["default"] || 0; }
    pl.orders.forEach(function (o) {
      if (!okOrder.test(String(o.status || "").trim())) return;
      var m = String(o.start || "").slice(0, 7); if (!set[m]) return;
      var dept = o.dept || "UNALLOCATED", key = dept + "|" + norm(o.client) + "|" + norm(o.title);
      seen[key] = 1; counted.orders++;
      var s = slot(dept, m); s.confirmed += o.value || 0;
      s.items.push({ kind: "order", ref: o.no, client: o.client, title: o.title, month: m, value: o.value || 0, status: o.status });
    });
    pl.quotes.forEach(function (q) {
      if (!okQuote.test(String(q.status || "").trim())) return;
      var dept = q.dept || "UNALLOCATED", base = String(q.accepted || q.sent || "").slice(0, 7);
      var m = shift(base, leadFor(dept)); if (!m || !set[m]) return;
      var key = dept + "|" + norm(q.client) + "|" + norm(q.title);
      if (seen[key]) { counted.dupQuotes++; return; }
      seen[key] = 1; counted.quotes++;
      var s = slot(dept, m); s.confirmed += q.value || 0;
      s.items.push({ kind: "quote", ref: q.ref, client: q.client, title: q.title, month: m, value: q.value || 0, status: q.status });
    });
    pl.deals.forEach(function (d) {
      var dept = d.dept || "UNALLOCATED", m = shift(String(d.close || "").slice(0, 7), leadFor(dept));
      if (!m || !set[m]) return;
      var key = dept + "|" + norm(d.account) + "|" + norm(d.name);
      if (won.test(d.stage || "")) {
        if (seen[key]) return;
        if (S.countWon) { seen[key] = 1; var sw = slot(dept, m); sw.confirmed += d.amount || 0; sw.items.push({ kind: "deal won", ref: "", client: d.account, title: d.name, month: m, value: d.amount || 0, status: d.stage }); }
        else wonUncounted.push({ client: d.account, title: d.name, month: m, value: d.amount || 0 });
        return;
      }
      if (lost.test(d.stage || "")) return;
      var prob = S.probability[d.stage] != null ? +S.probability[d.stage] : (d.prob != null ? +d.prob : (d.amount ? (d.expected || 0) / d.amount : 0));
      prob = Math.max(0, Math.min(1, prob || 0));
      var s = slot(dept, m); var w = Math.round((d.amount || 0) * prob); s.weighted += w; counted.deals++;
      s.deals.push({ client: d.account, title: d.name, stage: d.stage, month: m, amount: d.amount || 0, prob: prob, weighted: w });
    });
    return { byDept: out, wonUncounted: wonUncounted, counted: counted, settings: S,
             deptsWithData: Object.keys(out).filter(function (d) { return Object.keys(out[d]).length; }) };
  };

  function priorKey(m) { return (+m.slice(0, 4) - 1) + m.slice(4); }
  function signOf(name) {
    var c = (E.acct[name] || {}).cat;
    return (c === "cos" || c === "expenses" || c === "other_expenses") ? -1 : 1;
  }

  function build() {
    var rm = E.reportingMonth();
    if (E._fc && E._fc.rm === rm) return E._fc;
    if (E.viewDept() && !E._noView) return E.withoutView(build);
    var a = assumptions();
    var months = E.forecastMonths();
    var fc = { rm: rm, months: months, monthSet: {}, byDeptAcct: {}, byAcct: {}, rules: {},
               deptIncome: {}, overrides: [], unapplied: [], history: {} };
    months.forEach(function (m) { fc.monthSet[m] = 1; });
    var glHist = (E.glMonths || []).filter(function (m) { return m <= rm; });
    var last12 = glHist.slice(-12), lastN = glHist.slice(-a.runRate);
    fc.history = { months: glHist.length, last12: last12, runRate: lastN };
    var depts = E.depts.map(function (d) { return d.code; });

    function val(dept, name, m) { return (((E.idx.deptAcctMonth[dept] || {})[name] || {})[m]) || 0; }
    function sum(dept, name, keys) { var t = 0; keys.forEach(function (k) { t += val(dept, name, k); }); return t; }
    function growthFor(dept, rule) {
      if (rule.param != null) return rule.param;
      var g = a.growth; var v = g[dept] != null ? g[dept] : g["default"];
      return +v || 0;
    }
    function cell(name, dept, m, rule) {
      switch (rule.method) {
        case "zero": return 0;
        case "budget": {
          var b = (E.finBudget[name] || {})[m] || 0, sh = E.deptShareOfAccount(name);
          return Math.round(b * (sh[dept] || 0));
        }
        case "fixed": {
          // a Fixed line with no amount typed is a run rate, not a zero
          if (rule.param == null) return lastN.length ? Math.round(sum(dept, name, lastN) / lastN.length) : 0;
          var s2 = E.deptShareOfAccount(name);
          return Math.round(rule.param * 100 * signOf(name) * (s2[dept] || 0));
        }
        case "runrate": {
          var keys = rule.param != null && rule.param >= 1 ? glHist.slice(-Math.round(rule.param)) : lastN;
          return keys.length ? Math.round(sum(dept, name, keys) / keys.length) : 0;
        }
        case "seasonal": {
          var pk = priorKey(m);
          if (glHist.indexOf(pk) >= 0) return Math.round(val(dept, name, pk) * (1 + growthFor(dept, rule)));
          return lastN.length ? Math.round(sum(dept, name, lastN) / lastN.length) : 0;
        }
        case "revenue_pct": {
          var inc = (fc.deptIncome[dept] || {})[m] || 0;
          if (!inc) return 0;
          var pct = rule.param;
          if (pct == null) {
            var hi = 0; last12.forEach(function (k) { hi += (((E.idx.deptCatMonth[dept] || {}).income || {})[k]) || 0; });
            pct = hi ? sum(dept, name, last12) / hi : 0;       // carries the cost's own sign
          } else pct = -Math.abs(pct) * (signOf(name) === -1 ? 1 : -1);
          return Math.round(inc * pct);
        }
      }
      return 0;
    }

    // pass 1: income, which the revenue percentage methods need
    var incomeAccts = E.accounts.filter(function (x) { return x.cat === "income"; }).map(function (x) { return x.name; });
    var otherAccts = E.accounts.filter(function (x) { return x.cat !== "income"; }).map(function (x) { return x.name; });
    function run(names) {
      names.forEach(function (name) {
        var rule = E.forecastRule(name);
        if (rule.method === "revenue_pct" && (E.acct[name] || {}).cat === "income") rule = { method: "runrate", param: null, source: "default", note: "" };
        fc.rules[name] = rule;
        depts.forEach(function (dept) {
          months.forEach(function (m) {
            var v = cell(name, dept, m, rule);
            if (!v) return;
            ((fc.byDeptAcct[dept] = fc.byDeptAcct[dept] || {})[name] = fc.byDeptAcct[dept][name] || {})[m] = v;
            if ((E.acct[name] || {}).cat === "income") (fc.deptIncome[dept] = fc.deptIncome[dept] || {})[m] = (fc.deptIncome[dept][m] || 0) + v;
          });
        });
      });
    }
    run(incomeAccts);

    // the pipeline layer: the systems' view of income replaces the baseline
    // where they know more, department by department
    fc.layers = {}; fc.pipeline = null;
    var PS = pipelineSettings();
    if (PS.enabled) {
      var pt = E.pipelineTargets(months);
      fc.pipeline = pt;
      var incomeNames = incomeAccts.slice();
      depts.forEach(function (dept) {
        var hasData = pt.deptsWithData.indexOf(dept) >= 0;
        var never = PS.never[dept] != null ? +PS.never[dept] : +PS.never["default"];
        if (!hasData) never = 1;
        months.forEach(function (m, i) {
          var base = (fc.deptIncome[dept] || {})[m] || 0;
          var L = (pt.byDept[dept] || {})[m] || { confirmed: 0, weighted: 0, items: [], deals: [] };
          var booked = Math.round((L.confirmed + L.weighted) * (1 - PS.cancel));
          var share = Math.round(base * never);
          var near = i < PS.near;
          var target = near ? booked + share : Math.max(base, booked + share);
          var floor = target - booked - share;
          (fc.layers[dept] = fc.layers[dept] || {})[m] = { baseline: base, confirmed: L.confirmed, weighted: L.weighted, booked: booked, never: share, floor: floor,
                                                            target: target, rule: hasData ? (near ? "near" : "far") : "baseline", items: L.items, deals: L.deals };
          if (target === base || !hasData) return;
          var cells = fc.byDeptAcct[dept] || (fc.byDeptAcct[dept] = {});
          if (base) {
            var running = 0, lastName = null;
            incomeNames.forEach(function (n) {
              var v = (cells[n] || {})[m]; if (!v) return;
              var nv = Math.round(v * target / base); cells[n][m] = nv; running += nv; lastName = n;
            });
            if (lastName && running !== target) cells[lastName][m] += target - running;
          } else {
            // nothing in the history for this month: put it on the department's biggest income line
            var best = null, bestV = -1;
            incomeNames.forEach(function (n) { var v = Math.abs(sum(dept, n, last12)); if (v > bestV) { bestV = v; best = n; } });
            if (best) (cells[best] = cells[best] || {})[m] = target;
          }
          (fc.deptIncome[dept] = fc.deptIncome[dept] || {})[m] = target;
        });
      });
      incomeNames.forEach(function (n) { if (fc.rules[n] && fc.rules[n].method === "seasonal") fc.rules[n] = Object.assign({}, fc.rules[n], { method: "pipeline", source: "pipeline" }); });
    }
    run(otherAccts);

    // typed overrides: department level replaces the cell; company level is
    // spread on the method's own split, or lands on the account's home
    // department when the method gave nothing
    ((E.fcData && E.fcData.overrides) || []).forEach(function (o) {
      var name = o.account, m = o.month, why = null;
      if (!E.acct[name]) why = "account not in the chart";
      else if (!fc.monthSet[m]) why = "month outside the forecast";
      else if (o.dept && depts.indexOf(o.dept) < 0) why = "unknown department";
      if (why) { fc.unapplied.push(Object.assign({ why: why }, o)); return; }
      var cents = Math.round((+o.amount || 0) * 100) * signOf(name);
      var targets = o.dept ? [o.dept] : depts;
      var cur = 0; targets.forEach(function (d) { cur += ((fc.byDeptAcct[d] || {})[name] || {})[m] || 0; });
      var parts = {};
      if (o.dept) parts[o.dept] = cents;
      else if (cur) targets.forEach(function (d) { parts[d] = Math.round(cents * ((((fc.byDeptAcct[d] || {})[name] || {})[m]) || 0) / cur); });
      else parts[(E.acct[name].dept) || "ADMIN"] = cents;
      Object.keys(parts).forEach(function (d) {
        ((fc.byDeptAcct[d] = fc.byDeptAcct[d] || {})[name] = fc.byDeptAcct[d][name] || {})[m] = parts[d];
      });
      fc.overrides.push(Object.assign({ cents: cents, was: cur }, o));
    });

    depts.forEach(function (d) {
      Object.keys(fc.byDeptAcct[d] || {}).forEach(function (name) {
        Object.keys(fc.byDeptAcct[d][name]).forEach(function (m) {
          (fc.byAcct[name] = fc.byAcct[name] || {})[m] = (fc.byAcct[name][m] || 0) + fc.byDeptAcct[d][name][m];
        });
      });
    });
    return (E._fc = fc);
  }
  E.forecast = function () { return E.forecastEnabled() ? build() : null; };
  E.forecastLayers = function () { var f = E.forecast(); return f ? { layers: f.layers || {}, pipeline: f.pipeline, months: f.months } : null; };

  E.acctForecast = function (name, m) {
    if (!E.forecastEnabled()) return 0;
    return ((build().byAcct[name] || {})[m]) || 0;
  };
  E.pnlForecast = function (keys) {
    var d = E.viewDept();
    return E.buildPnl(keys, function (n, m) {
      if (!E.hasForecast(m)) return 0;
      if (!d) return E.acctForecast(n, m);
      var f = build(); return (((f.byDeptAcct[d] || {})[n] || {})[m]) || 0;
    });
  };
  /** Forecast by department for a category over months, the shape budgetByDept has. */
  E.forecastByDept = function (monthKeys, cat) {
    var out = {}; E.depts.forEach(function (d) { out[d.code] = 0; });
    if (!E.forecastEnabled()) return out;
    var fc = build();
    Object.keys(fc.byDeptAcct).forEach(function (d) {
      Object.keys(fc.byDeptAcct[d]).forEach(function (name) {
        var a = E.acct[name]; if (!a || (cat && a.cat !== cat)) return;
        out[d] = (out[d] || 0) + E.sumMonths(fc.byDeptAcct[d][name], monthKeys);
      });
    });
    return out;
  };
  /** One row per subcategory saying how its accounts are forecast. */
  E.forecastLines = function (monthKeys) {
    if (!E.forecastEnabled()) return [];
    var fc = build(), rows = [];
    E.CAT_KEYS.forEach(function (cat) {
      var subs = E.subsOf[cat] || {};
      Object.keys(subs).sort().forEach(function (sub) {
        var names = subs[sub], methods = {}, sources = {}, total = 0, params = {};
        names.forEach(function (n) {
          var r = fc.rules[n]; if (!r) return;
          methods[r.method] = 1; sources[r.source] = 1;
          if (r.param != null) params[r.param] = 1;
          total += E.sumMonths(fc.byAcct[n], monthKeys);
        });
        var mk = Object.keys(methods), sk = Object.keys(sources), pk = Object.keys(params);
        rows.push({ cat: E.CAT_LABEL[cat], catKey: cat, sub: sub, accounts: names.length,
                    method: mk.length === 1 ? mk[0] : mk.length ? "mixed" : "none",
                    methodLabel: mk.length === 1 ? METHODS[mk[0]].short : mk.length ? "mixed" : "",
                    source: sk.length === 1 ? sk[0] : "mixed", param: pk.length === 1 ? +pk[0] : null, total: total });
      });
    });
    return rows;
  };
})();

/* ===================================================================== *
 * Engine part 4: cash. Twelve months on the indirect method, off the P&L
 * forecast: profit, less what is not cash, less the working capital the
 * revenue ties up, less the tax, the commitments that are not in the P&L
 * and the credit card balances. Opening cash, the commitments and the
 * cards come from 10 Cash and Commitments.xlsx.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, E = CTS.engine;

  var LABOUR = /salar|wage|superann|workers|payroll tax|fringe/i;
  var SUPER = /superann/i;
  var NONCASH = /deprec|amortis|write off|leave expense|unrealised|revaluation|disposal|provision for tax|income tax expense|gain on lease/i;
  var GSTFREE = /salar|wage|superann|workers|payroll tax|fringe|bank charges|interest|stamp duty|filing fee|donation|currency|bad debts|dividend|deprec|amortis/i;
  var GST = 0.1;

  E.loadCash = function (data) { E.cashData = data || null; E._cash = null; };
  E.cashEnabled = function () { return !!E.cashData; };

  function settings() {
    var s = (E.cashData && E.cashData.settings) || {};
    return {
      debtorDays: +s.debtorDays > 0 ? +s.debtorDays : 45,
      creditorDays: +s.creditorDays > 0 ? +s.creditorDays : 30,
      basFrequency: /month/i.test(s.basFrequency || "") ? "monthly" : "quarterly",
      superTiming: /quarter/i.test(s.superTiming || "") ? "quarterly" : "payday",
      paygInstalment: Math.round((+s.paygInstalment || 0) * 100),
      facility: Math.round((+s.facility || 0) * 100),
      openingAR: s.openingAR != null && s.openingAR !== "" ? Math.round(+s.openingAR * 100) : null,
      openingAP: s.openingAP != null && s.openingAP !== "" ? Math.round(+s.openingAP * 100) : null,
      cardsPaidInFull: !/^n/i.test(String(s.cardsPaidInFull == null ? "yes" : s.cardsPaidInFull)),
    };
  }
  E.cashSettings = settings;

  /** Balances and cards at the latest month on or before the reporting month. */
  function latest(list, rm) {
    var months = {}; (list || []).forEach(function (r) { if (r.month && r.month <= rm) months[r.month] = 1; });
    var keys = Object.keys(months).sort(); var at = keys[keys.length - 1] || null;
    return { month: at, rows: at ? list.filter(function (r) { return r.month === at; }) : [] };
  }
  E.cashBalances = function () { return latest((E.cashData || {}).balances, E.reportingMonth()); };
  E.cashCards = function () { return latest((E.cashData || {}).cards, E.reportingMonth()); };

  function addMonths(iso, n) {
    var y = +iso.slice(0, 4), m = +iso.slice(5, 7) - 1 + n, d = +iso.slice(8, 10) || 1;
    var dt = new Date(Date.UTC(y + Math.floor(m / 12), ((m % 12) + 12) % 12, 1));
    var last = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 0)).getUTCDate();
    dt.setUTCDate(Math.min(d, last));
    return dt.toISOString().slice(0, 10);
  }
  var FREQ = { weekly: 7, fortnightly: 14, monthly: 1, quarterly: 3, annual: 12, yearly: 12, once: 0 };
  function freqOf(s) { s = String(s || "monthly").toLowerCase(); return Object.keys(FREQ).filter(function (k) { return s.indexOf(k.slice(0, 4)) === 0; })[0] || "monthly"; }

  /** The dates one commitment falls due inside [from, to], and its P&L accrual a month. */
  function occurrences(c, from, to) {
    var f = freqOf(c.frequency), out = [], d = c.next || from;
    if (!d) return { dates: out, perMonth: 0 };
    var end = c.end && c.end < to ? c.end : to, guard = 0;
    if (f === "once") { if (d >= from && d <= end) out.push(d); }
    else if (FREQ[f] === 7 || FREQ[f] === 14) {
      while (d <= end && guard++ < 400) { if (d >= from) out.push(d); var t = new Date(d + "T00:00:00Z"); t.setUTCDate(t.getUTCDate() + FREQ[f]); d = t.toISOString().slice(0, 10); }
    } else {
      while (d <= end && guard++ < 400) { if (d >= from) out.push(d); d = addMonths(d, FREQ[f]); }
    }
    var perMonth = f === "once" ? 0 : f === "weekly" ? 52 / 12 : f === "fortnightly" ? 26 / 12 : 1 / FREQ[f];
    return { dates: out, perMonth: perMonth };
  }

  E.cashflow = function () {
    if (E.viewDept() && !E._noView) return E.withoutView(function () { return E.cashflow(); });
    var rm = E.reportingMonth();
    if (E._cash && E._cash.rm === rm) return E._cash;
    if (!E.cashEnabled() || !E.forecastEnabled()) return null;
    var S = settings(), months = E.forecastMonths().slice(0, 12);
    if (!months.length) return null;
    var lastKey = months[months.length - 1];
    var from = months[0] + "-01", to = lastKey + "-" + new Date(+lastKey.slice(0, 4), +lastKey.slice(5, 7), 0).getDate();

    // account classes
    function cls(name) {
      var a = E.acct[name] || {}, key = (a.sub || "") + " " + name;
      return { cat: a.cat, labour: LABOUR.test(key), sup: SUPER.test(key), noncash: NONCASH.test(key), gstFree: GSTFREE.test(key) };
    }
    // a month's P&L, actual or forecast, grouped the way cash needs it
    function pl(m) {
      var actual = E.hasActual(m);
      var r = { income: 0, otherIncome: 0, otherIncomeGst: 0, costs: 0, costsGst: 0, labour: 0, sup: 0, noncash: 0, otherExpenses: 0, netProfit: 0 };
      E.accounts.forEach(function (a) {
        var v = actual ? ((E.finActual[a.name] || {})[m] || 0) : E.acctForecast(a.name, m);
        if (!v) return;
        var c = cls(a.name);
        r.netProfit += v;
        if (c.cat === "income") r.income += v;
        else if (c.cat === "other_income") { r.otherIncome += v; if (!c.gstFree) r.otherIncomeGst += v; }
        else if (c.noncash) r.noncash += v;
        else if (c.labour) { r.labour += v; if (c.sup) r.sup += v; }
        else { r.costs += v; if (!c.gstFree) r.costsGst += v; }
      });
      return r;
    }
    var plOf = {};
    function P(m) { return plOf[m] || (plOf[m] = pl(m)); }

    // receivables and payables balances a month end carries, on days to pay
    function priorMonths(m, n) { var out = [], i = E.monthIdx[m] ? E.monthIdx[m].i : -1; for (var k = 0; k < n; k++) if (i - k >= 0) out.push(E.months[i - k].key); return out; }
    function balance(m, days, pick) {
      var k = days / 30, keys = priorMonths(m, Math.ceil(k) + 1), t = 0;
      keys.forEach(function (key, i) { var w = Math.max(0, Math.min(1, k - i)); if (w) t += pick(P(key)) * w; });
      return Math.round(t);
    }
    var arPick = function (p) { return p.income * (1 + GST) + p.otherIncome + p.otherIncomeGst * GST; };
    var apPick = function (p) { return -(p.costs + p.costsGst * GST); };

    // commitments
    var commitments = (E.cashData.commitments || []).map(function (c) {
      var occ = occurrences(c, from, to), amt = Math.round((+c.amount || 0) * 100), gst = /^y/i.test(String(c.gst || ""));
      var inPnl = !/^n/i.test(String(c.inPnl == null ? "yes" : c.inPnl));
      var byMonth = {}; occ.dates.forEach(function (d) { byMonth[d.slice(0, 7)] = (byMonth[d.slice(0, 7)] || 0) + amt; });
      // the accrual for timing is GST inclusive, like the payment: the P&L
      // carries it ex GST and the payments line already adds the GST back
      var accrual = Math.round(amt * occ.perMonth);
      return Object.assign({}, c, { cents: amt, gstFlag: gst, inPnl: inPnl, byMonth: byMonth, accrual: accrual, dates: occ.dates, freq: freqOf(c.frequency) });
    });

    var opening = E.cashBalances(), cards = E.cashCards();
    var openCash = opening.rows.reduce(function (a, r) { return a + Math.round((+r.balance || 0) * 100); }, 0);
    var cardOwing = cards.rows.reduce(function (a, r) { return a + Math.round((+r.balance || 0) * 100); }, 0);
    var cardLimit = cards.rows.reduce(function (a, r) { return a + Math.round((+r.limit || 0) * 100); }, 0);

    // A typed receivables or payables balance is better than an assumed
    // number of days: measure the days from it, so the walk starts from the
    // real balance and moves at the pace that balance implies.
    function solveDays(target, pick) {
      var lo = 0, hi = 180;
      for (var i = 0; i < 40; i++) { var mid = (lo + hi) / 2; if (balance(rm, mid, pick) < target) lo = mid; else hi = mid; }
      return Math.round((lo + hi) / 2);
    }
    S = Object.assign({}, S, { debtorSource: "typed days", creditorSource: "typed days" });
    if (S.openingAR != null && S.openingAR > 0) { S.debtorDays = Math.max(1, solveDays(S.openingAR, arPick)); S.debtorSource = "measured from the receivables balance"; }
    if (S.openingAP != null && S.openingAP > 0) { S.creditorDays = Math.max(1, solveDays(S.openingAP, apPick)); S.creditorSource = "measured from the payables balance"; }
    var arPrev = S.openingAR != null ? S.openingAR : balance(rm, S.debtorDays, arPick);
    var apPrev = S.openingAP != null ? S.openingAP : balance(rm, S.creditorDays, apPick);
    var arModel0 = balance(rm, S.debtorDays, arPick), apModel0 = balance(rm, S.creditorDays, apPick);

    // GST and super owed at the start: what the current quarter has built up
    var q = E.monthIdx[rm] ? E.monthIdx[rm].quarter : 1;
    var gstOwed = 0, superOwed = 0;
    E.monthsOfFY(E.currentFY()).filter(function (k) { return k <= rm && E.monthIdx[k].quarter === q; }).forEach(function (k) {
      var p = P(k); gstOwed += Math.round((p.income + p.otherIncomeGst + p.costsGst) * GST); superOwed += -p.sup;
    });
    if (S.basFrequency === "monthly") { gstOwed = Math.round((P(rm).income + P(rm).otherIncomeGst + P(rm).costsGst) * GST); }

    var gstOpen = gstOwed, superOpen = superOwed;
    var rows = [], cash = openCash, low = null;
    months.forEach(function (m, i) {
      var p = P(m), mo = E.monthIdx[m].m;
      var ar = balance(m, S.debtorDays, arPick), ap = balance(m, S.creditorDays, apPick);
      var receipts = Math.round(arPick(p)) + arPrev - ar;
      var payments = Math.round(apPick(p)) + apPrev - ap;
      var wages = -(p.labour - (S.superTiming === "quarterly" ? p.sup : 0));
      var gstMonth = Math.round((p.income + p.otherIncomeGst + p.costsGst) * GST);
      var basMonth = S.basFrequency === "monthly" || [10, 1, 4, 7].indexOf(mo) >= 0;
      var gstPaid = 0, superPaid = 0, payg = 0;
      if (basMonth) { gstPaid = gstOwed; gstOwed = 0; payg = S.basFrequency === "monthly" ? Math.round(S.paygInstalment / 3) : S.paygInstalment;
                      if (S.superTiming === "quarterly") { superPaid = superOwed; superOwed = 0; } }
      gstOwed += gstMonth;
      if (S.superTiming === "quarterly") superOwed += -p.sup;
      var timing = 0, outside = 0, outsideItems = [];
      commitments.forEach(function (c) {
        var paid = c.byMonth[m] || 0;
        if (c.inPnl) timing += c.accrual - paid;
        else if (paid) { outside += paid; outsideItems.push(c.name + " " + F_money(paid)); }
      });
      var cardPay = i === 0 && S.cardsPaidInFull ? cardOwing : 0;
      var other = p.otherExpenses;
      var net = receipts - payments - wages - gstPaid - superPaid - payg + timing - outside - cardPay + other;
      var closing = cash + net;
      rows.push({ month: m, label: E.monthIdx[m].label, opening: cash, income: p.income, netProfit: p.netProfit, noncash: -p.noncash,
                  receipts: receipts, payments: -payments, wages: -wages, gst: -gstPaid, sup: -superPaid, payg: -payg,
                  timing: timing, outside: -outside, outsideItems: outsideItems, cards: -cardPay, net: net, closing: closing,
                  ar: ar, ap: ap, gstOwed: gstOwed, superOwed: superOwed, headroom: closing + S.facility, bas: basMonth });
      if (low == null || closing < rows[low].closing) low = rows.length - 1;
      cash = closing; arPrev = ar; apPrev = ap;
    });
    function F_money(c) { return CTS.fmt.money(c); }
    return (E._cash = {
      rm: rm, months: months, rows: rows, opening: openCash, openingMonth: opening.month, balances: opening.rows,
      cards: cards.rows, cardsMonth: cards.month, cardOwing: cardOwing, cardLimit: cardLimit,
      closing: rows[rows.length - 1].closing, low: rows[low], settings: S, commitments: commitments,
      openingAR: S.openingAR != null ? S.openingAR : arModel0, openingAP: S.openingAP != null ? S.openingAP : apModel0,
      arModelled: S.openingAR == null, apModelled: S.openingAP == null, gstOpening: gstOpen, superOpening: superOpen,
    });
  };
})();

/* ===================================================================== *
 * Engine part 5: commentary. What people say about the numbers, kept
 * beside them. Published commentary travels in data/CTS_commentary_data.js
 * like every other data file; a comment typed in the portal is a draft in
 * this browser until finance publishes it to the folder.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, E = CTS.engine, store = CTS.store;

  /** Who gets the emails: the list edited in the portal (this browser), then
   *  the list published to the folder, then the Config template. */
  E.recipients = function () {
    var local = store.get("distribution.list", null);
    if (Array.isArray(local)) return local;
    var pub = window.CTS_DISTRIBUTION && Array.isArray(window.CTS_DISTRIBUTION.list) ? window.CTS_DISTRIBUTION.list : null;
    return pub || (E.cfg.DISTRIBUTION || []);
  };
  E.setRecipients = function (list) { store.set("distribution.list", list); };
  E.recipientsLocal = function () { return Array.isArray(store.get("distribution.list", null)); };
  E.clearRecipientsLocal = function () { store.set("distribution.list", null); };

  E.loadCommentary = function (data) {
    E.commentaryData = data && Array.isArray(data.items) ? data : { meta: {}, items: [] };
  };
  function drafts() { return store.get("commentary.drafts", []) || []; }
  function removed() { return store.get("commentary.removed", []) || []; }
  E.canComment = function () {
    var r = E.currentRole();
    return !!E.identity() && !!r && !!r.admin;
  };
  E.canPublishCommentary = function () { var r = E.currentRole(); return !!r && (r.admin || r.id === "finance"); };

  /** Items for a page and month, published then drafts, scoped to the
   *  person's department where they only see their own. */
  E.commentaryFor = function (page, month, opts) {
    opts = opts || {};
    var scope = opts.all ? null : E.deptScope();
    var gone = {}; removed().forEach(function (id) { gone[id] = 1; });
    var pub = (E.commentaryData.items || []).filter(function (c) { return !gone[c.id] && (!page || c.page === page) && (!month || c.month === month); })
      .map(function (c) { return Object.assign({}, c, { draft: false }); });
    var dr = drafts().filter(function (c) { return (!page || c.page === page) && (!month || c.month === month); })
      .map(function (c) { return Object.assign({}, c, { draft: true }); });
    return pub.concat(dr).filter(function (c) { return !scope || !c.dept || c.dept === scope; })
      .sort(function (a, b) { return String(a.at).localeCompare(String(b.at)); });
  };
  E.addCommentary = function (item) {
    var list = drafts();
    var it = { id: "c" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), month: item.month || E.reportingMonth(),
               page: item.page, dept: item.dept || E.deptScope() || "", text: String(item.text || "").trim(),
               author: E.identity() || "Unknown", at: new Date().toISOString() };
    if (!it.text) return null;
    list.push(it); store.set("commentary.drafts", list); return it;
  };
  E.updateCommentary = function (id, text) {
    var list = drafts(), hit = false;
    list.forEach(function (c) { if (c.id === id) { c.text = String(text || "").trim(); c.at = new Date().toISOString(); hit = true; } });
    if (hit) store.set("commentary.drafts", list.filter(function (c) { return c.text; }));
    return hit;
  };
  E.removeCommentary = function (id) {
    var list = drafts();
    if (list.some(function (c) { return c.id === id; })) { store.set("commentary.drafts", list.filter(function (c) { return c.id !== id; })); return; }
    var r = removed(); if (r.indexOf(id) < 0) { r.push(id); store.set("commentary.removed", r); }
  };
  E.commentaryDraftCount = function () { return drafts().length + removed().length; };
  /** Everything that would be in the file if it were published now. */
  E.commentaryExport = function () {
    var gone = {}; removed().forEach(function (id) { gone[id] = 1; });
    var items = (E.commentaryData.items || []).filter(function (c) { return !gone[c.id]; }).concat(drafts());
    return { meta: { published: new Date().toISOString(), by: E.identity() || "", count: items.length,
                     note: "Commentary typed in the portal and published to the folder. One item per page per month per author." }, items: items };
  };
  E.clearCommentaryDrafts = function () { store.set("commentary.drafts", []); store.set("commentary.removed", []); };
})();

/* ===================================================================== *
 * Engine part 6: the fortnight. Revenue by date off the ledger for the
 * month in progress, hours off the pay runs, the forecast for the month
 * as the benchmark, and the confirmed work landing next. This is what
 * the fortnightly update reads.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, E = CTS.engine;

  function dim(key) { return new Date(+key.slice(0, 4), +key.slice(5, 7), 0).getDate(); }
  function weekdays(key, from, to) {
    var y = +key.slice(0, 4), m = +key.slice(5, 7) - 1, n = 0;
    for (var d = from; d <= to; d++) { var dow = new Date(y, m, d).getDay(); if (dow !== 0 && dow !== 6) n++; }
    return n;
  }
  function nextKey(key) { var y = +key.slice(0, 4), m = +key.slice(5, 7); return m === 12 ? (y + 1) + "-01" : y + "-" + (m + 1 < 10 ? "0" : "") + (m + 1); }

  /** The month the fortnightly update is about: the month after the
   *  reporting month when the ledger already carries lines in it. */
  E.progressMonth = function () {
    var rm = E.reportingMonth(), later = (E.glMonths || []).filter(function (k) { return k > rm; });
    return later.length ? later[later.length - 1] : rm;
  };
  E.snapshotDate = function () { return (E.quality && E.quality.maxDate) || null; };

  E.fortnightFor = function (dept, monthKey, which) {
    var key = monthKey || E.progressMonth(), days = dim(key), snap = E.snapshotDate() || (key + "-" + days);
    var snapDay = snap.slice(0, 7) === key ? +snap.slice(8, 10) : (snap > key ? days : 0);
    var closed = E.hasActual(key) || snap.slice(0, 7) > key;
    if (closed) snapDay = days;
    var F = [{ n: 1, from: 1, to: 15 }, { n: 2, from: 16, to: days }];
    F.forEach(function (f) {
      f.label = "Fortnight " + f.n; f.range = f.from + "–" + f.to + " " + (E.monthIdx[key] || {}).label;
      f.complete = snapDay >= f.to; f.started = snapDay >= f.from;
      f.weekdays = weekdays(key, f.from, f.to);
      f.weekdaysDone = f.started ? weekdays(key, f.from, Math.min(f.to, snapDay)) : 0;
      f.revenue = 0; f.clients = {};
    });
    var depts = dept ? [dept] : E.revenueDepts.map(function (d) { return d.code; });
    var inDept = {}; depts.forEach(function (d) { inDept[d] = 1; });
    (E.gl || []).forEach(function (r) {
      if (r.cat !== "income" || r.month !== key || !inDept[r.dept]) return;
      var day = +r.date.slice(8, 10), f = day <= 15 ? F[0] : F[1];
      f.revenue += r.amount; if (r.contact) f.clients[r.contact] = (f.clients[r.contact] || 0) + r.amount;
    });
    var mtd = F[0].revenue + F[1].revenue;
    // hours: the pay runs dated in the month, else the last month with any
    var hoursKey = key, rows = E.staffRows.filter(function (r) { return r.month === key && inDept[r.dept]; });
    if (!rows.length) { var prev = E.months.filter(function (m) { return m.key < key; }).map(function (m) { return m.key; }).reverse()
      .filter(function (k) { return E.staffRows.some(function (r) { return r.month === k && inDept[r.dept]; }); })[0]; if (prev) { hoursKey = prev; rows = E.staffRows.filter(function (r) { return r.month === prev && inDept[r.dept]; }); } }
    var people = {}, chg = 0, worked = 0, leave = 0;
    rows.forEach(function (r) { people[r.empId] = r.name; chg += r.chargeable; worked += r.chargeable + r.nonChargeable; leave += r.leave; });
    var wdTotal = F[0].weekdays + F[1].weekdays, wdDone = F[0].weekdaysDone + F[1].weekdaysDone;
    var share = hoursKey === key && closed ? 1 : (wdTotal ? wdDone / wdTotal : 0);
    var chgToDate = chg * share, workedToDate = worked * share;
    F.forEach(function (f) { var sh = wdTotal ? f.weekdaysDone / wdTotal : 0; f.chargeable = chg * sh; f.worked = worked * sh; f.sellRate = f.chargeable ? Math.round(f.revenue / f.chargeable) : null; f.util = f.worked ? f.chargeable / f.worked : null; });
    // benchmarks
    var forecast = 0, budget = 0, actualMonth = null;
    depts.forEach(function (d) {
      budget += E.budgetByDept([key], "income")[d] || 0;
      if (E.hasActual(key)) { actualMonth = (actualMonth || 0) + E.sumMonths((E.idx.deptCatMonth[d] || {}).income, [key]); }
      else if (E.forecastEnabled() && E.hasForecast(key)) forecast += E.forecastByDept([key], "income")[d] || 0;
    });
    var benchmark = actualMonth != null ? actualMonth : (forecast || budget);
    var benchmarkKind = actualMonth != null ? "actual" : forecast ? "forecast" : "budget";
    var remaining = Math.max(0, benchmark - mtd);
    var open = F.filter(function (f) { return !f.complete; });
    var openWd = open.reduce(function (a, f) { return a + (f.weekdays - f.weekdaysDone); }, 0);
    F.forEach(function (f) { f.schedule = f.complete ? f.revenue : f.revenue + (openWd ? Math.round(remaining * (f.weekdays - f.weekdaysDone) / openWd) : 0); });
    var proRata = wdTotal ? Math.round(benchmark * wdDone / wdTotal) : 0;
    // gross margin to date, off the ledger
    var cos = 0; (E.gl || []).forEach(function (r) { if (r.cat === "cos" && r.month === key && inDept[r.dept] && +r.date.slice(8, 10) <= snapDay) cos += r.amount; });
    // what lands next, from the pipeline layer
    var landing = [], L = E.forecastLayers ? E.forecastLayers() : null;
    if (L) [key, nextKey(key)].forEach(function (mk) { depts.forEach(function (d) { var x = (L.layers[d] || {})[mk]; if (!x) return; (x.items || []).forEach(function (it) { landing.push(Object.assign({ dept: d }, it)); }); }); });
    landing.sort(function (a, b) { return b.value - a.value; });
    var util = E.utilFor([hoursKey]);
    var uRow = dept ? util.rows.filter(function (r) { return r.dept.code === dept; })[0] : null;
    var current = which ? F[which - 1] : (F[1].started ? F[1] : F[0]);
    return {
      month: key, monthLabel: (E.monthIdx[key] || {}).long || key, snapshot: snap, snapDay: snapDay, days: days, closed: closed,
      fortnights: F, current: current, mtd: mtd, benchmark: benchmark, benchmarkKind: benchmarkKind, budget: budget, proRata: proRata,
      remaining: remaining, cos: cos, gm: mtd ? (mtd + cos) / mtd : null,
      headcount: Object.keys(people).length, team: Object.keys(people).map(function (id) { return people[id]; }).sort(),
      hoursMonth: hoursKey, hoursProRated: hoursKey !== key || !closed, chargeable: chgToDate, worked: workedToDate, leave: leave * share,
      sellRate: chgToDate ? Math.round(mtd / chgToDate) : null, util: workedToDate ? chgToDate / workedToDate : null,
      target: uRow ? uRow.target : (E.cfg.UTIL && E.cfg.UTIL.target) || null, fte: uRow ? uRow.fte : util.total.fte,
      landing: landing.slice(0, 12), landingTotal: landing.reduce(function (a, i) { return a + i.value; }, 0),
      topClients: Object.keys(current.clients).map(function (c) { return { client: c, amount: current.clients[c] }; }).sort(function (a, b) { return b.amount - a.amount; }).slice(0, 5),
      dept: dept || null,
    };
  };
})();

/* ===================================================================== *
 * Engine part 7: proposed commentary. A first draft written from the
 * numbers for the page in view, for the finance head to edit. It says
 * what moved and against what; the person adds why.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, E = CTS.engine, F = CTS.fmt;
  function d(c) { return F.dollars(Math.abs(c)); }
  function vs(a, b, what) { var v = a - b; return what + " " + (v >= 0 ? "ahead of" : "behind") + " budget by " + d(v) + (b ? " (" + (v >= 0 ? "+" : "−") + (Math.abs(v / b) * 100).toFixed(1) + "%)" : ""); }
  function monthLabel(rm) { return (E.monthIdx[rm] || {}).long || rm; }
  var GEN = {
    home: function (rm, p) {
      var a = E.pnlActual(p.keys).totals, b = E.pnlBudget(p.keys).totals, py = E.pnlActual(E.priorYearMonths(p.keys)).totals, u = E.utilFor(p.keys);
      return [vs(a.income, b.income, "Revenue " + d(a.income) + " for " + p.label + ","), "against " + d(py.income) + " for the same period last year.",
              "Gross margin " + F.pct(a.income ? a.grossProfit / a.income : null) + " against " + F.pct(b.income ? b.grossProfit / b.income : null) + " budgeted.",
              vs(a.netProfit, b.netProfit, "Net profit " + F.money(a.netProfit) + ",") + ".", "Utilisation " + F.pct(u.total.util) + " across " + (u.total.fte ? u.total.fte.toFixed(1) + " FTE" : "the business") + ".",
              "[Why: what drove the revenue, what moved margin, what is one off.]"].join(" ");
    },
    pnl: function (rm, p) {
      var a = E.pnlActual(p.keys), b = E.pnlBudget(p.keys), out = [];
      a.cats.forEach(function (cat) { var bc = b.by[cat.key]; cat.children.forEach(function (sub) { var bs = (bc.children || []).filter(function (x) { return x.sub === sub.sub; })[0]; var bv = bs ? bs.total : 0; if (E.risk(sub.total, bv) === "High") out.push(cat.label + ", " + sub.sub + ": " + F.money(sub.total) + " against " + F.money(bv) + " budget."); }); });
      return (out.length ? "Lines flagged High for " + p.label + ": " + out.slice(0, 5).join(" ") : "No line is flagged High for " + p.label + ".") + " [Say whether each is timing, a one off, or a trend.]";
    },
    "pnl-dept": function (rm, p) {
      var alloc = E.allocate(p.keys), bd = E.budgetByDept(p.keys, "income");
      return alloc.rows.filter(function (r) { return r.dept.isRevenue && r.dept.code !== "ADMIN" && E.visibleDepts().some(function (x) { return x.code === r.code; }); }).map(function (r) {
        return r.dept.short + ": revenue " + d(r.income) + " " + (r.income >= (bd[r.code] || 0) ? "ahead of" : "behind") + " budget by " + d(r.income - (bd[r.code] || 0)) + ", gross margin " + F.pct(r.gmPct) + (r.dept.gmNorm != null ? (r.gmPct != null && r.gmPct < r.dept.gmNorm ? " below" : " at or above") + " the " + F.pct(r.dept.gmNorm, 0) + " norm" : "") + ", net profit after the split " + F.money(r.netProfit) + ".";
      }).join(" ") + " [Name the jobs behind the variance.]";
    },
    bva: function (rm, p) { return GEN.pnl(rm, p); },
    "rev-summary": function (rm, p) {
      var prior = E.priorYearMonths(p.keys), bd = E.budgetByDept(p.keys, "income");
      return E.visibleDepts().filter(function (x) { return x.isRevenue && x.code !== "ADMIN"; }).map(function (x) {
        var now = E.sumMonths((E.idx.deptCatMonth[x.code] || {}).income, p.keys), was = E.sumMonths((E.idx.deptCatMonth[x.code] || {}).income, prior);
        return x.short + " " + d(now) + ", " + (now >= was ? "up" : "down") + " " + d(now - was) + " on last year, " + (now >= (bd[x.code] || 0) ? "ahead of" : "behind") + " budget by " + d(now - (bd[x.code] || 0)) + ".";
      }).join(" ") + " [Which clients moved it.]";
    },
    "rev-schedule": function (rm, p) {
      var dept = CTS.store.get("schedDept", "PRODUCTION"), s = E.schedule(dept, p.keys);
      var misses = s.lines.filter(function (l) { return l.risk === "High"; }).slice(0, 5);
      return (E.deptOf[dept] || {}).short + " schedule " + d(s.actualTotal) + " against " + d(s.budgetTotal) + " budgeted. " + (misses.length ? "High variances: " + misses.map(function (l) { return (l.client || l.line) + " " + F.money(l.variance); }).join(", ") + "." : "No line is flagged High.") + " [For each miss: lost, or moved to a later month?]";
    },
    forecast: function (rm, p) {
      if (!E.forecastEnabled()) return "The forecast is not loaded.";
      var fy = E.monthsOfFY(p.fy), land = E.pnlBlend(fy).totals, fb = E.pnlBudget(fy).totals, f = E.forecast();
      return "FY" + p.fy + " is landing at revenue " + d(land.income) + " (" + (land.income >= fb.income ? "ahead of" : "behind") + " budget by " + d(land.income - fb.income) + ") and net profit " + F.money(land.netProfit) + " against " + F.money(fb.netProfit) + " budgeted. " + (f.overrides.length ? f.overrides.length + " typed override" + (f.overrides.length > 1 ? "s" : "") + " in force. " : "") + "[What would change the landing: the pipeline, a hire, a cost.]";
    },
    "rev-forecast": function (rm, p) {
      var L = E.forecastLayers(); if (!L) return "The forecast is not loaded.";
      var m = L.months[0], conf = 0, wgt = 0; Object.keys(L.layers).forEach(function (dpt) { var x = L.layers[dpt][m]; if (x) { conf += x.confirmed; wgt += x.weighted; } });
      return "Next month has " + d(conf) + " confirmed in the systems and " + d(wgt) + " of weighted open deals" + (L.pipeline && L.pipeline.wonUncounted.length ? ", with " + L.pipeline.wonUncounted.length + " won deals not yet carried by an order or quote" : "") + ". [Which deals are the swing.]";
    },
    cash: function (rm) {
      var c = E.cashEnabled() && E.forecastEnabled() ? E.cashflow() : null; if (!c) return "The cash flow is not loaded.";
      return "Cash " + d(c.opening) + " at " + monthLabel(rm) + ", low point " + d(c.low.closing) + " in " + c.low.label + ", " + d(c.closing) + " in twelve months. Credit cards owing " + d(c.cardOwing) + ". [Anything large falling due that the commitments tab does not carry.]";
    },
    util: function (rm, p) {
      var u = E.utilFor(p.keys), low = u.rows.filter(function (r) { return r.target != null && r.util != null && r.util < r.target - 0.05; });
      return "Utilisation " + F.pct(u.total.util) + " for " + p.label + ". " + (low.length ? low.map(function (r) { return r.dept.short + " " + F.pct(r.util) + " against " + F.pct(r.target, 0); }).join(", ") + " under target." : "Every department is within five points of target.") + " [Leave, training, or a genuine gap in work.]";
    },
    clients: function (rm, p) {
      var rows = E.clientRows(p.keys, E.priorYearMonths(p.keys)).slice(0, 3);
      return "Top clients for " + p.label + ": " + rows.map(function (r) { return r.display + " " + d(r.amount) + " (" + (r.delta >= 0 ? "up" : "down") + " " + d(r.delta) + " on last year)"; }).join(", ") + ". [Any client at risk or new.]";
    },
    "staff-profit": function (rm, p) {
      var st = E.staffFor(p.keys), neg = st.rows.filter(function (r) { return r.billable && r.margin < 0; });
      return neg.length ? neg.length + " billable people carried a negative margin for " + p.label + ": " + neg.slice(0, 4).map(function (r) { return r.name; }).join(", ") + ". [Rate, hours, or a quiet month.]" : "Every billable person covered their cost for " + p.label + ".";
    },
  };
  GEN["profit-fte"] = GEN.util; GEN.pipeline = GEN["rev-forecast"]; GEN["client-dept"] = GEN.clients; GEN.story = GEN.home; GEN["pnl-spread"] = GEN.pnl; GEN.actions = GEN.pnl;
  E.proposeCommentary = function (pageId) {
    var rm = E.reportingMonth(), p = CTS.period();
    var g = GEN[pageId];
    try { return g ? g(rm, p) : GEN.home(rm, p); } catch (e) { return "Could not draft from the data: " + (e && e.message || e); }
  };
  E.proposablePages = function () { return Object.keys(GEN).filter(function (k) { return CTS.pages[k] && E.can(k); }); };
})();

/* ===================================================================== *
 * UI primitives and charts.
 *
 * Charts are hand written SVG, no library, so the portal has no runtime
 * dependency and works from a file:// path with no network. Every chart is
 * laid out the way the graphs file lays them out: a labelled table beside the
 * picture, never a picture on its own. That is also what discharges the
 * contrast relief rule for the three lighter series colours.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, h = CTS.h, svg = CTS.svg, F = CTS.fmt, E = CTS.engine;

  var U = (CTS.ui = {});

  /* Categorical slots in fixed order, one per department, never cycled.
   * Validated with the dataviz palette validator: all checks pass in light
   * mode (three slots below 3:1 contrast, relieved by the paired table) and
   * all checks pass in dark mode. */
  U.SERIES = ["ONSITE", "PRODUCTION", "VIDEO", "INTEGRATION", "CONSULTING", "ADMIN", "UNALLOCATED"];
  U.colourOf = function (code) {
    var i = U.SERIES.indexOf(code);
    return i < 0 ? "var(--ink-muted)" : "var(--series-" + (i + 1) + ")";
  };
  U.RISK_CLASS = { "High": "critical", "Medium": "warning", "Low": "good", "No Activity": "muted" };

  /* ---- small pieces --------------------------------------------------- */
  U.spark = function (values, colour) {
    var vals = (values || []).filter(function (v) { return typeof v === "number" && isFinite(v); });
    if (vals.length < 2) return null;
    var w = 120, hh = 28, lo = Math.min.apply(null, vals.concat([0])), hi = Math.max.apply(null, vals.concat([0]));
    var span = hi - lo || 1, n = vals.length;
    var pts = vals.map(function (v, i) { return [(i / (n - 1)) * (w - 2) + 1, hh - 1 - ((v - lo) / span) * (hh - 2)]; });
    var d = pts.map(function (p, i) { return (i ? "L" : "M") + p[0].toFixed(1) + " " + p[1].toFixed(1); }).join(" ");
    var zero = hh - 1 - ((0 - lo) / span) * (hh - 2);
    var last = pts[pts.length - 1];
    return svg("svg", { class: "spark", viewBox: "0 0 " + w + " " + hh, "aria-hidden": "true", focusable: "false" }, [
      svg("path", { d: d + " L" + last[0].toFixed(1) + " " + zero.toFixed(1) + " L1 " + zero.toFixed(1) + " Z", fill: colour || "var(--sec)", opacity: ".12" }),
      lo < 0 ? svg("line", { x1: 1, x2: w - 1, y1: zero, y2: zero, stroke: "var(--line-strong)", "stroke-width": 1 }) : null,
      svg("path", { d: d, fill: "none", stroke: colour || "var(--sec)", "stroke-width": 1.8, "stroke-linejoin": "round", "stroke-linecap": "round" }),
      svg("circle", { cx: last[0], cy: last[1], r: 2.6, fill: colour || "var(--sec)" }),
    ]);
  };
  /** A tile. With href it is a link to the page behind the number; with
   *  spark it carries the last months as a line. */
  U.tile = function (opts) {
    var tag = opts.href ? "a.tile" : "div.tile";
    var attrs = opts.href ? { href: opts.href, title: opts.title || "Open " + (opts.label || "").toLowerCase() } : null;
    return h(tag + (opts.tone ? ".tone-" + opts.tone : ""), attrs, [
      h("div.tile-label", opts.label),
      h("div.tile-value", opts.value),
      opts.sub ? h("div.tile-sub", opts.sub) : null,
      opts.spark ? U.spark(opts.spark, opts.sparkColour) : null,
      opts.note ? h("div.tile-note", opts.note) : null,
      opts.href ? h("span.tile-go", { "aria-hidden": "true" }, "→") : null,
    ]);
  };
  /** A dark cover band for one department, carrying its division mark. */
  U.cover = function (code) {
    var d = E.deptOf[code], b = window.CTS_BRAND || {};
    if (!d) return null;
    var sub = (b.subBrandOf || {})[code], mark = sub && b.marks && b.marks[sub];
    return h("div.cover", { style: { "--cover": U.colourOf(code) } }, [
      h("div.cover-text", [h("div.eyebrow", sub ? "CTS " + sub.charAt(0).toUpperCase() + sub.slice(1) : "CTS"), h("h2", d.name || d.short), d.normNote ? h("p", d.normNote) : null]),
      mark ? h("img", { src: mark.onDark, alt: mark.alt }) : null,
    ]);
  };

  U.flag = function (kind, text, title) {
    return h("span.flag.flag-" + kind, { title: title || "" }, text);
  };

  U.note = function (text, kind) {
    return h("p.note" + (kind ? ".note-" + kind : ""), text);
  };

  U.source = function (text) {
    return h("p.source", [h("span.source-tag", "Source"), " " + text]);
  };

  U.slug = function (s) { return String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60); };
  U.chev = function (cls) {
    return svg("svg", { class: cls || "chev", viewBox: "0 0 16 16", "aria-hidden": "true", focusable: "false" },
      [svg("path", { d: "M3 6l5 5 5-5", fill: "none", stroke: "currentColor", "stroke-width": "1.8", "stroke-linecap": "round", "stroke-linejoin": "round" })]);
  };
  U.icon = function (name) {
    var P = { speaker: "M3 6h3l4-3v10l-4-3H3z M12 5.5a3.5 3.5 0 0 1 0 5 M14 3.5a6 6 0 0 1 0 9",
              stop: "M4 4h8v8H4z", collapse: "M3 5h10M3 8h10M3 11h10", expand: "M8 2v12M2 8h12",
              menu: "M2 4h12M2 8h12M2 12h12", theme: "M8 2a6 6 0 0 0 0 12z M8 2a6 6 0 0 1 0 12",
              text: "M3 4h10M8 4v9M5.5 13h5", contrast: "M8 2a6 6 0 1 0 0 12 6 6 0 0 0 0-12z M8 2v12" };
    return svg("svg", { viewBox: "0 0 16 16", "aria-hidden": "true", focusable: "false" },
      [svg("path", { d: P[name] || "", fill: name === "stop" ? "currentColor" : "none", stroke: "currentColor", "stroke-width": "1.6", "stroke-linecap": "round", "stroke-linejoin": "round" })]);
  };

  U.h1 = function (title, sub) {
    return h("header.page-head", [
      h("div.titles", [
        CTS.router && CTS.router.current && CTS.pages[CTS.router.current] ? h("div.eyebrow", CTS.pages[CTS.router.current].section) : null,
        h("h1", title),
        sub || E.viewDept() ? h("p.page-sub", (sub || "") + (E.viewDept() ? (sub ? " \u00b7 " : "") + ((E.deptOf[E.viewDept()] || {}).short || E.viewDept()) + " only" : "")) : null,
      ]),
      h("div.page-tools", { "data-tools": "1" }),
    ]);
  };

  /** A collapsible card. The open or closed state is remembered per page
   *  and section in this browser, so a person who closes the workings keeps
   *  them closed next month. */
  U.section = function (title, kids, note, opts) {
    opts = opts || {};
    var page = (CTS.router && CTS.router.current) || "";
    var key = "sec." + page + "." + U.slug(title);
    var collapsed = !!CTS.store.get(key, !!opts.collapsed);
    var id = "sec-" + page + "-" + U.slug(title);
    var sec = h("section.block", { "data-collapsed": collapsed ? "1" : "0", "aria-labelledby": id + "-h" });
    var body = h("div.block-body", { id: id }, [note ? U.note(note) : null].concat(Array.isArray(kids) ? kids : [kids]));
    var head = h("button.block-head", {
      type: "button", "aria-expanded": collapsed ? "false" : "true", "aria-controls": id,
      onclick: function () {
        var now = sec.getAttribute("data-collapsed") !== "1";
        sec.setAttribute("data-collapsed", now ? "1" : "0");
        head.setAttribute("aria-expanded", now ? "false" : "true");
        CTS.store.set(key, now);
      },
    }, [h("h2", { id: id + "-h" }, title), U.chev("chev")]);
    sec.appendChild(head); sec.appendChild(body);
    return sec;
  };
  U.setAllSections = function (collapsed) {
    Array.prototype.forEach.call(document.querySelectorAll("#main section.block"), function (sec) {
      sec.setAttribute("data-collapsed", collapsed ? "1" : "0");
      var head = sec.querySelector(".block-head"); if (head) head.setAttribute("aria-expanded", collapsed ? "false" : "true");
      var page = (CTS.router && CTS.router.current) || "", title = head && head.querySelector("h2") ? head.querySelector("h2").textContent : "";
      CTS.store.set("sec." + page + "." + U.slug(title), collapsed);
    });
  };

  /** Table builder. cols: [{key,label,align,fmt,cls,width}] */
  U.table = function (cols, rows, opts) {
    opts = opts || {};
    var thead = h("thead", h("tr", cols.map(function (c) {
      return h("th", { class: (c.align || "right") + (c.cls ? " " + c.cls : ""),
                       style: c.width ? { width: c.width } : null,
                       title: c.title || "" }, c.label);
    })));
    var body = h("tbody", rows.map(function (r) {
      var tr = h("tr" + (r._cls ? "." + r._cls : ""), cols.map(function (c) {
        var v = typeof c.value === "function" ? c.value(r) : r[c.key];
        if (v && v.nodeType) return h("td", { class: c.align || "right" }, v);
        var txt = c.fmt ? c.fmt(v, r) : (v == null ? "-" : String(v));
        return h("td", { class: (c.align || "right") + (c.cell ? " " + c.cell(r) : "") }, txt);
      }));
      if (r._onclick) { tr.classList.add("clickable"); tr.addEventListener("click", r._onclick); }
      return tr;
    }));
    var t = h("table.grid" + (opts.dense ? ".dense" : ""), [thead, body]);
    return opts.scroll === false ? t : h("div.tablewrap", t);
  };

  U.money = function (v) { return F.money(v); };
  U.moneyCell = function (v) { return v < 0 ? "neg" : ""; };

  /* ---- chart chrome --------------------------------------------------- */
  function niceTicks(lo, hi, count) {
    if (lo === hi) { hi = lo + 1; }
    var span = hi - lo;
    var step = Math.pow(10, Math.floor(Math.log10(span / (count || 5))));
    [1, 2, 2.5, 5, 10].some(function (m) {
      if (span / (step * m) <= (count || 5)) { step = step * m; return true; }
      return false;
    });
    var start = Math.floor(lo / step) * step;
    var out = [];
    for (var v = start; v <= hi + step / 2; v += step) out.push(v);
    return out;
  }

  function frame(opts) {
    var w = opts.width || 720, hgt = opts.height || 260;
    var pad = Object.assign({ l: 62, r: 16, t: 12, b: 34 }, opts.pad || {});
    var node = svg("svg", { class: "chart", viewBox: "0 0 " + w + " " + hgt,
                            preserveAspectRatio: "xMidYMid meet",
                            role: "img", "aria-label": opts.label || "chart" });
    return { node: node, w: w, h: hgt, pad: pad,
             iw: w - pad.l - pad.r, ih: hgt - pad.t - pad.b };
  }

  function yAxis(f, lo, hi, fmt) {
    var ticks = niceTicks(lo, hi, 5);
    var g = svg("g", { class: "axis" });
    ticks.forEach(function (t) {
      var y = f.pad.t + f.ih - ((t - lo) / (hi - lo)) * f.ih;
      g.appendChild(svg("line", { x1: f.pad.l, x2: f.pad.l + f.iw, y1: y, y2: y,
                                  class: t === 0 ? "zero" : "grid" }));
      g.appendChild(svg("text", { x: f.pad.l - 8, y: y + 4, class: "tick", "text-anchor": "end" },
                        fmt ? fmt(t) : String(t)));
    });
    f.node.appendChild(g);
    return function (v) { return f.pad.t + f.ih - ((v - lo) / (hi - lo)) * f.ih; };
  }

  function xLabels(f, labels, step) {
    var g = svg("g", { class: "axis" });
    var bw = f.iw / labels.length;
    labels.forEach(function (l, i) {
      if (step && i % step !== 0) return;
      g.appendChild(svg("text", { x: f.pad.l + bw * (i + 0.5), y: f.h - 12,
                                  class: "tick", "text-anchor": "middle" }, l));
    });
    f.node.appendChild(g);
  }

  function tip(f) {
    var box = h("div.charttip", { style: { display: "none" } });
    return {
      el: box,
      bind: function (node, html) {
        node.addEventListener("mouseenter", function (ev) {
          box.innerHTML = html;
          box.style.display = "block";
        });
        node.addEventListener("mousemove", function (ev) {
          var r = box.parentNode.getBoundingClientRect();
          box.style.left = Math.min(r.width - 170, Math.max(0, ev.clientX - r.left + 12)) + "px";
          box.style.top = Math.max(0, ev.clientY - r.top - 10) + "px";
        });
        node.addEventListener("mouseleave", function () { box.style.display = "none"; });
      },
    };
  }

  U.legend = function (items) {
    return h("div.legend", items.map(function (it) {
      return h("span.legend-item", [
        h("span.swatch", { style: { background: it.colour } }),
        it.label,
      ]);
    }));
  };

  /** A chart with its table underneath, never beside it: a chart squeezed
   *  into half the width squashes its labels and the note beside it. */
  U.figure = function (title, chart, table, note) {
    return h("figure.fig", [
      h("figcaption", title),
      h("div.fig-body", [
        h("div.fig-chart", [chart]),
        table ? h("div.fig-table", [table]) : null,
      ]),
      note ? U.note(note) : null,
    ]);
  };

  /** The commentary block under a page's heading: what has been said about
   *  this page for the reporting month, and a place to say more. */
  U.commentary = function (pageId) {
    var month = E.reportingMonth(), items = E.commentaryFor(pageId, month);
    var can = E.canComment();
    var label = (E.monthIdx[month] || {}).long || month;
    var body = h("div.commentary-body");
    function item(c) {
      var view = h("div.commentary-text-view", c.text);
      var meta = h("div.commentary-meta", [
        h("strong", c.author), h("span", F.date(String(c.at).slice(0, 10))),
        c.dept ? h("span.chip.chip-muted", (E.deptOf[c.dept] || {}).short || c.dept) : null,
        c.draft ? h("span.chip.chip-warning", "draft, this browser only") : h("span.chip.chip-good", "published"),
      ]);
      var row = h("div.commentary-item", [meta, view]);
      if (can && (c.draft || E.canPublishCommentary())) {
        var tools = h("div.row", [
          c.draft ? h("button.btn.btn-quiet.small", { type: "button", onclick: function () {
            var ta = h("textarea.commentary-text", c.text);
            var form = h("div.commentary-form", [ta, h("div.row", [
              h("button.btn.small", { type: "button", onclick: function () { E.updateCommentary(c.id, ta.value); CTS.router.reload(); } }, "Save"),
              h("button.btn.btn-quiet.small", { type: "button", onclick: function () { CTS.router.reload(); } }, "Cancel"),
            ])]);
            row.replaceChild(form, view); tools.remove();
          } }, "Edit") : null,
          h("button.btn.btn-quiet.small", { type: "button", onclick: function () { E.removeCommentary(c.id); CTS.router.reload(); } }, c.draft ? "Delete draft" : "Remove"),
        ]);
        row.appendChild(tools);
      }
      return row;
    }
    if (items.length) items.forEach(function (c) { body.appendChild(item(c)); });
    else body.appendChild(h("div.commentary-empty", can ? "Nothing written about this page for " + label + " yet. Explain the variance rather than restate it; management prefer dollars against dollars." : "Nothing written about this page for " + label + ". The finance head writes the commentary."));
    if (can) {
      var ta = h("textarea.commentary-text", { placeholder: "Add commentary for " + label + "…", "aria-label": "Commentary for " + label });
      var form = h("div.commentary-form", { hidden: true }, [ta, h("div.row", [
        h("button.btn.small", { type: "button", onclick: function () { if (E.addCommentary({ page: pageId, month: month, text: ta.value })) CTS.router.reload(); } }, "Save draft"),
        h("button.btn.btn-quiet.small", { type: "button", onclick: function () { form.hidden = true; addBtn.hidden = false; } }, "Cancel"),
        h("span.muted", "Saved in this browser until published from the Commentary page. Square brackets mark where the proposal needs your why."),
      ])]);
      var addBtn = h("button.btn.btn-quiet.small", { type: "button", onclick: function () { form.hidden = false; addBtn.hidden = true; ta.focus(); } }, "Add commentary");
      var propBtn = h("button.btn.small", { type: "button", title: "Write a first draft from this page's numbers, for you to edit",
        onclick: function () { form.hidden = false; addBtn.hidden = true; propBtn.hidden = true; ta.value = E.proposeCommentary(pageId); ta.focus(); } }, "Propose from the data");
      body.appendChild(h("div.row", [addBtn, propBtn]));
      body.appendChild(form);
    }
    var key = "sec." + pageId + ".commentary", collapsed = !!CTS.store.get(key, false);
    var wrap = h("section.commentary", { "data-collapsed": collapsed ? "1" : "0", "aria-label": "Commentary" });
    var head = h("div.commentary-head", [
      h("h2", "Commentary, " + label + (items.length ? " (" + items.length + ")" : "")),
      E.commentaryDraftCount() && E.canPublishCommentary() ? h("a.chip.chip-warning", { href: "#/commentary" }, E.commentaryDraftCount() + " to publish") : null,
    ]);
    wrap.appendChild(head); wrap.appendChild(body);
    return wrap;
  };

  /* ---- grouped / stacked columns -------------------------------------- */
  /** series: [{key,label,colour,values:[]}], labels: x labels */
  U.columns = function (opts) {
    var f = frame(opts);
    var wrap = h("div.chartwrap");
    var t = tip(f);
    var series = opts.series, labels = opts.labels;
    var stacked = !!opts.stacked;
    var n = labels.length;

    var lo = 0, hi = 0;
    if (stacked) {
      labels.forEach(function (_, i) {
        var pos = 0, neg = 0;
        series.forEach(function (s) { var v = s.values[i] || 0; if (v >= 0) pos += v; else neg += v; });
        hi = Math.max(hi, pos); lo = Math.min(lo, neg);
      });
    } else {
      series.forEach(function (s) {
        s.values.forEach(function (v) { hi = Math.max(hi, v || 0); lo = Math.min(lo, v || 0); });
      });
    }
    if (hi === lo) hi = lo + 1;
    var y = yAxis(f, lo, hi, opts.yfmt || F.k);
    var bw = f.iw / n;
    var inner = stacked ? bw * 0.62 : (bw * 0.7) / series.length;
    var gap = 2;                                   // surface gap between fills

    var g = svg("g");
    labels.forEach(function (lab, i) {
      var accPos = 0, accNeg = 0;
      series.forEach(function (s, si) {
        var v = s.values[i] || 0;
        if (!v) return;
        var x = stacked
          ? f.pad.l + bw * i + (bw - inner) / 2
          : f.pad.l + bw * i + bw * 0.15 + si * inner;
        var y0, y1;
        if (stacked) {
          if (v >= 0) { y0 = y(accPos + v); y1 = y(accPos); accPos += v; }
          else { y0 = y(accNeg); y1 = y(accNeg + v); accNeg += v; }
        } else { y0 = y(Math.max(0, v)); y1 = y(Math.min(0, v)); }
        var height = Math.max(1, y1 - y0 - (stacked ? gap : 0));
        var rect = svg("rect", {
          x: x, y: y0, width: Math.max(1, inner - gap), height: height,
          rx: 3, fill: s.colour, class: "mark",
        });
        t.bind(rect, "<b>" + lab + "</b><br>" + s.label + ": " +
               (opts.tipfmt || F.dollars)(v));
        if (opts.onPick) { rect.classList.add("clickable"); rect.setAttribute("tabindex", "0"); rect.setAttribute("role", "button");
          rect.addEventListener("click", function () { opts.onPick(i, s, v); });
          rect.addEventListener("keydown", function (ev) { if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); opts.onPick(i, s, v); } }); }
        g.appendChild(rect);
      });
    });
    f.node.appendChild(g);
    xLabels(f, labels, opts.labelStep || (n > 14 ? 2 : 1));
    wrap.appendChild(f.node);
    wrap.appendChild(t.el);
    return h("div", [
      series.length > 1 ? U.legend(series.map(function (s) {
        return { label: s.label, colour: s.colour }; })) : null,
      wrap,
    ]);
  };

  /* ---- lines ----------------------------------------------------------- */
  U.lines = function (opts) {
    var f = frame(opts);
    var wrap = h("div.chartwrap");
    var t = tip(f);
    var series = opts.series, labels = opts.labels;
    var lo = Infinity, hi = -Infinity;
    series.forEach(function (s) {
      s.values.forEach(function (v) {
        if (v == null) return;
        lo = Math.min(lo, v); hi = Math.max(hi, v);
      });
    });
    if (!isFinite(lo)) { lo = 0; hi = 1; }
    if (opts.zero) lo = Math.min(0, lo);
    var pad = (hi - lo) * 0.1 || 1;
    var y = yAxis(f, lo - pad, hi + pad, opts.yfmt || F.k);
    var bw = f.iw / labels.length;
    var xs = labels.map(function (_, i) { return f.pad.l + bw * (i + 0.5); });

    if (opts.target != null) {
      f.node.appendChild(svg("line", {
        x1: f.pad.l, x2: f.pad.l + f.iw, y1: y(opts.target), y2: y(opts.target),
        class: "targetline" }));
      // sits just inside the left edge, clear of the series at the right
      f.node.appendChild(svg("text", { x: f.pad.l + 6, y: y(opts.target) - 6,
                                       class: "tick", "text-anchor": "start" },
                             opts.targetLabel || "target"));
    }

    series.forEach(function (s) {
      var d = "", started = false;
      s.values.forEach(function (v, i) {
        if (v == null) { started = false; return; }
        d += (started ? "L" : "M") + xs[i] + " " + y(v);
        started = true;
      });
      f.node.appendChild(svg("path", { d: d, fill: "none", stroke: s.colour,
                                       "stroke-width": 2, "stroke-linejoin": "round",
                                       "stroke-linecap": "round", class: "mark" }));
      s.values.forEach(function (v, i) {
        if (v == null) return;
        var c = svg("circle", { cx: xs[i], cy: y(v), r: 4.5, fill: s.colour,
                                stroke: "var(--surface-1)", "stroke-width": 2, class: "mark" });
        if (opts.onPick) { c.classList.add("clickable"); c.addEventListener("click", function () { opts.onPick(i, s, v); }); }
        t.bind(c, "<b>" + labels[i] + "</b><br>" + s.label + ": " +
               (opts.tipfmt || F.dollars)(v));
        f.node.appendChild(c);
      });
    });
    xLabels(f, labels, opts.labelStep || (labels.length > 14 ? 2 : 1));
    wrap.appendChild(f.node);
    wrap.appendChild(t.el);
    return h("div", [
      series.length > 1 ? U.legend(series.map(function (s) {
        return { label: s.label, colour: s.colour }; })) : null,
      wrap,
    ]);
  };

  /* ---- horizontal bars (rankings) -------------------------------------- */
  U.bars = function (opts) {
    var rows = opts.rows;                      // [{label, value, value2?, colour}]
    var hgt = Math.max(120, rows.length * 26 + 30);
    var f = frame({ width: opts.width || 520, height: hgt,
                    pad: { l: opts.labelWidth || 160, r: 64, t: 8, b: 22 } });
    var t = tip(f);
    var wrap = h("div.chartwrap");
    var hi = 0, lo = 0;
    rows.forEach(function (r) {
      hi = Math.max(hi, r.value || 0, r.value2 || 0);
      lo = Math.min(lo, r.value || 0, r.value2 || 0);
    });
    if (hi === lo) hi = lo + 1;
    var x = function (v) { return f.pad.l + ((v - lo) / (hi - lo)) * f.iw; };
    var rh = (f.ih / rows.length);
    var two = rows.some(function (r) { return r.value2 != null; });
    rows.forEach(function (r, i) {
      var bh = two ? rh * 0.34 : rh * 0.52;
      [["value", r.colour || "var(--measure-1)", 0],
       ["value2", "var(--measure-3)", bh + 2]].forEach(function (pair) {
        var v = r[pair[0]];
        if (v == null) return;
        var y0 = f.pad.t + rh * i + (rh - (two ? bh * 2 + 2 : bh)) / 2 + pair[2];
        var x0 = x(Math.min(0, v)), x1 = x(Math.max(0, v));
        var rect = svg("rect", { x: x0, y: y0, width: Math.max(1, x1 - x0),
                                 height: bh, rx: 3, fill: pair[1], class: "mark" });
        t.bind(rect, "<b>" + r.label + "</b><br>" +
               (pair[0] === "value" ? (opts.label1 || "This period") : (opts.label2 || "Prior")) +
               ": " + F.dollars(v));
        f.node.appendChild(rect);
      });
      f.node.appendChild(svg("text", { x: f.pad.l - 8, y: f.pad.t + rh * i + rh / 2 + 4,
                                       class: "tick", "text-anchor": "end" },
                             r.label.length > 24 ? r.label.slice(0, 23) + "…" : r.label));
      f.node.appendChild(svg("text", { x: x(Math.max(0, r.value)) + 6,
                                       y: f.pad.t + rh * i + rh / 2 + 4,
                                       class: "tick value", "text-anchor": "start" },
                             F.k(r.value)));
    });
    f.node.appendChild(svg("line", { x1: x(0), x2: x(0), y1: f.pad.t,
                                     y2: f.pad.t + f.ih, class: "zero" }));
    wrap.appendChild(f.node);
    wrap.appendChild(t.el);
    return h("div", [
      two ? U.legend([{ label: opts.label1 || "This period", colour: "var(--measure-1)" },
                      { label: opts.label2 || "Prior", colour: "var(--measure-3)" }]) : null,
      wrap,
    ]);
  };

  /* ---- waterfall (the variance bridge) --------------------------------- */
  U.waterfall = function (opts) {
    var steps = opts.steps;                   // [{label,value,kind:'base'|'delta'|'total'}]
    var f = frame({ width: opts.width || 700, height: opts.height || 280,
                    pad: { l: 70, r: 16, t: 16, b: 46 } });
    var t = tip(f);
    var wrap = h("div.chartwrap");
    var run = 0, tops = [], lo = 0, hi = 0;
    steps.forEach(function (s) {
      if (s.kind === "delta") { tops.push([run, run + s.value]); run += s.value; }
      else { tops.push([0, s.value]); run = s.value; }
      lo = Math.min(lo, run, 0); hi = Math.max(hi, run);
    });
    if (hi === lo) hi = lo + 1;
    var y = yAxis(f, lo, hi * 1.05, F.k);
    var bw = f.iw / steps.length;
    steps.forEach(function (s, i) {
      var a = y(tops[i][0]), b = y(tops[i][1]);
      var colour = s.kind !== "delta" ? "var(--ink-strong)"
        : (s.value >= 0 ? "var(--status-good)" : "var(--status-critical)");
      var rect = svg("rect", { x: f.pad.l + bw * i + bw * 0.22,
                               y: Math.min(a, b), width: bw * 0.56,
                               height: Math.max(1, Math.abs(b - a)), rx: 3,
                               fill: colour, class: "mark" });
      t.bind(rect, "<b>" + s.label + "</b><br>" + F.dollars(s.value));
      f.node.appendChild(rect);
      f.node.appendChild(svg("text", { x: f.pad.l + bw * (i + 0.5),
                                       y: Math.min(a, b) - 6, class: "tick value",
                                       "text-anchor": "middle" }, F.k(s.value)));
      f.node.appendChild(svg("text", { x: f.pad.l + bw * (i + 0.5), y: f.h - 26,
                                       class: "tick", "text-anchor": "middle" },
                             s.label.length > 13 ? s.label.slice(0, 12) + "…" : s.label));
      if (i < steps.length - 1 && s.kind === "delta") {
        f.node.appendChild(svg("line", {
          x1: f.pad.l + bw * i + bw * 0.78, x2: f.pad.l + bw * (i + 1) + bw * 0.22,
          y1: y(tops[i][1]), y2: y(tops[i][1]), class: "connector" }));
      }
    });
    wrap.appendChild(f.node);
    wrap.appendChild(t.el);
    return wrap;
  };
})();

/* ===================================================================== *
 * Pages: dashboard, finance and budget.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, h = CTS.h, F = CTS.fmt, E = CTS.engine, U = CTS.ui;
  var P = (CTS.pages = CTS.pages || {});

  /* ---- the period control, shared by most pages ----------------------- */
  var PERIODS = [
    { id: "month", label: "Month" },
    { id: "qtr", label: "Quarter to date" },
    { id: "ytd", label: "Year to date" },
    { id: "fy", label: "Full year" },
  ];
  CTS.period = function () {
    var id = CTS.store.get("period", "ytd");
    var fy = E.currentFY(), rm = E.reportingMonth();
    var keys, label;
    if (id === "month") { keys = [rm]; label = (E.monthIdx[rm] || {}).long || rm; }
    else if (id === "qtr") { keys = E.quarterMonths(); label = "Quarter to " + ((E.monthIdx[rm] || {}).label || rm); }
    else if (id === "fy") { keys = E.monthsOfFY(fy); label = "FY" + fy + " full year, " + E.blendLabel(); }
    else { keys = E.ytd(fy); label = "FY" + fy + " to " + ((E.monthIdx[rm] || {}).label || rm); }
    return { id: id, keys: keys, label: label, fy: fy, rm: rm };
  };
  CTS.periodBar = function (extra) {
    var cur = CTS.store.get("period", "ytd");
    return h("div.toolbar", [
      h("div.seg", PERIODS.map(function (p) {
        return h("button.seg-btn" + (p.id === cur ? ".on" : ""), {
          onclick: function () { CTS.store.set("period", p.id); CTS.router.reload(); },
        }, p.label);
      })),
      deptPicker(),
      h("div.toolbar-right", [
        h("label.inline", "Reporting month"),
        monthSelect(),
      ].concat(extra || [])),
    ]);
  };
  /** The department picker: every page that has a department view follows it. */
  function deptPicker() {
    var scope = E.deptScope();
    if (scope) return h("span.filterchip", { title: "Your role sees this department only" }, [h("span.dot", { style: { background: CTS.ui.colourOf(scope) } }), (E.deptOf[scope] || {}).short || scope]);
    var cur = E.focusDept() || "";
    var opts = [{ code: "", short: "Whole company" }].concat(E.depts.filter(function (d) { return d.isRevenue; }));
    return h("span.row", { style: { marginTop: 0 } }, [
      h("label.inline", "Department"),
      h("select.control", { "aria-label": "Department in view", onchange: function (e) { E.setFocusDept(e.target.value || null); CTS.router.reload(); } },
        opts.map(function (d) { return h("option", { value: d.code, selected: d.code === cur }, d.short); })),
    ]);
  }
  function monthSelect() {
    var sel = h("select.control", {
      onchange: function (e) { E.setReportingMonth(e.target.value); CTS.router.reload(); },
    }, E.months.map(function (m) {
      return h("option", { value: m.key, selected: m.key === E.reportingMonth() }, m.long);
    }));
    return sel;
  }

  function seedBanner() {
    if (E.restored && E.restored.length) {
      return h("div.banner.banner-good", [
        h("strong", "Using data you kept in this browser. "),
        "Restored " + E.restored.join(", ") +
        (E.restoredAt ? ", saved " + F.date(E.restoredAt.slice(0, 10)) : "") +
        ". It is on this machine only. Clear it on ",
        h("a", { href: "#/setup" }, "Setup"), " to go back to the files.",
      ]);
    }
    if (!E.seed) return null;
    return h("div.banner.banner-seed", [
      h("strong", "Seed data. "),
      "The FY26 totals, the July 2025 month and the August 2026 category totals are the real Xero figures and tie to the cent. Everything inside them is modelled from the seasonality and margin norms in the Monthly Reporting manual. Paste the real exports into the templates and run ",
      h("a", { href: "#/build" }, "Build"), " to replace it.",
    ]);
  }
  CTS.seedBanner = seedBanner;

  function allocBanner() {
    var s = E.allocationIsSourced();
    if (s.ok) return null;
    return h("div.banner.banner-warn", [
      h("strong", "Allocated figures rest on placeholder splits. "),
      "The " + s.placeholders.join(" and ") + " percentages live in the budget workbook and are not in this repository. Set them on ",
      h("a", { href: "#/config" }, "Config & Variables"),
      " before any allocated departmental result is relied on.",
    ]);
  }
  CTS.allocBanner = allocBanner;

  /* ===================================================== Dashboard ===== */
  P.home = {
    section: "Dashboard", title: "Dashboard",
    sub: "Where the business is, on one page.",
    render: function () {
      var p = CTS.period();
      var pnl = E.pnlBlend(p.keys);
      var prior = E.pnlActual(E.priorYearMonths(p.keys));
      var bud = E.pnlBudget(p.keys);
      var t = pnl.totals, pt = prior.totals, bt = bud.totals;
      var util = E.utilFor(p.keys);

      function delta(a, b) {
        if (!b) return null;
        return h("span" + (a - b >= 0 ? ".up" : ".down"),
                 (a - b >= 0 ? "+" : "−") + F.money(Math.abs(a - b)) + " vs last year");
      }

      // twelve months of history behind each headline, for the sparklines
      var histKeys = E.months.filter(function (m) { return m.key <= p.rm; }).slice(-12).map(function (m) { return m.key; });
      var hist = E.pnlActual(histKeys);
      var utilHist = histKeys.map(function (k) { var u = E.utilFor([k]); return u.total.util == null ? null : u.total.util * 100; });
      var tiles = h("div.tiles", [
        U.tile({ label: "Revenue", value: F.dollars(t.income), sub: delta(t.income, pt.income),
                 note: "Budget " + F.dollars(bt.income), href: "#/rev-summary", spark: hist.income, sparkColour: "var(--measure-1)" }),
        U.tile({ label: "Gross profit", value: F.dollars(t.grossProfit),
                 sub: F.pct(t.income ? t.grossProfit / t.income : null) + " margin",
                 note: "Budget " + F.pct(bt.income ? bt.grossProfit / bt.income : null), href: "#/pnl", spark: hist.grossProfit, sparkColour: "var(--measure-2)" }),
        (function () {
          var v = E.viewDept(), np = t.netProfit, noteTxt = "Budget " + F.dollars(bt.netProfit);
          if (v) { var ar = E.allocate(p.keys).byCode[v]; if (ar) { np = ar.netProfit; noteTxt = "After the overhead split. Own overhead " + F.money(ar.ownExpenses) + ", share of Admin " + F.money(ar.allocated); } }
          return U.tile({ label: v ? "Net profit after split" : "Net profit", value: F.dollars(np),
                 tone: np >= 0 ? "good" : "critical",
                 sub: F.pct(t.income ? np / t.income : null) + " of revenue",
                 note: noteTxt, href: v ? "#/pnl-dept" : "#/bva", spark: v ? null : hist.netProfit, sparkColour: "var(--measure-3)" });
        })(),
        U.tile({ label: "Utilisation", value: F.pct(util.total.util),
                 sub: util.total.fte ? util.total.fte.toFixed(1) + " full time equivalents" : null,
                 note: "Chargeable over worked hours, leave excluded", href: "#/util", spark: utilHist, sparkColour: "var(--series-5)" }),
      ].concat(E.cashEnabled() && E.forecastEnabled() && E.can("cash") && !E.viewDept() ? [(function () {
        var c = E.cashflow(); if (!c) return null;
        return U.tile({ label: "Cash low point", value: F.dollars(c.low.closing), tone: c.low.closing < 0 ? "critical" : "good",
                        sub: c.low.label + ", from " + F.dollars(c.opening) + " now", note: "Twelve months ahead, indirect method", href: "#/cash",
                        spark: c.rows.map(function (r) { return r.closing; }), sparkColour: "var(--series-2)" });
      })()] : []).concat(E.forecastEnabled() ? [(function () {
        var fyK = E.monthsOfFY(p.fy), land = E.pnlBlend(fyK).totals, fb = E.pnlBudget(fyK).totals;
        return U.tile({ label: "FY" + p.fy + " forecast", value: F.dollars(land.netProfit),
                        tone: land.netProfit >= fb.netProfit ? "good" : "critical",
                        sub: (land.netProfit - fb.netProfit >= 0 ? "+" : "\u2212") + F.money(Math.abs(land.netProfit - fb.netProfit)) + " vs budget",
                        note: "Net profit, actual then forecast. Revenue " + F.dollars(land.income), href: "#/forecast" });
      })()] : []));

      // Revenue, gross profit and net profit by month across the year
      var fyKeys = E.monthsOfFY(p.fy);
      var full = E.pnlBlend(fyKeys);
      var labels = fyKeys.map(function (k) { return E.monthIdx[k].label; });
      var trend = U.columns({
        labels: labels, width: 760, height: 250,
        series: [
          { label: "Revenue", colour: "var(--measure-1)", values: full.income },
          { label: "Gross profit", colour: "var(--measure-2)", values: full.grossProfit },
          { label: "Net profit", colour: "var(--measure-3)", values: full.netProfit },
        ],
        onPick: function (i) { if (E.blendSource(fyKeys[i]) !== "actual") return; E.setReportingMonth(fyKeys[i]); CTS.store.set("period", "month"); CTS.router.reload(); },
      });
      var trendTable = U.table([
        { key: "m", label: "Month", align: "left" },
        { key: "src", label: "", align: "left" },
        { key: "rev", label: "Revenue", fmt: F.money },
        { key: "gp", label: "Gross profit", fmt: F.money },
        { key: "np", label: "Net profit", fmt: F.money },
      ], fyKeys.map(function (k, i) {
        return { m: E.monthIdx[k].label,
                 src: E.blendSource(k) === "actual" ? "" : E.blendSource(k),
                 rev: full.income[i], gp: full.grossProfit[i], np: full.netProfit[i],
                 _cls: E.blendSource(k) === "actual" ? "" : "forecast" };
      }), { dense: true });

      // departmental result after allocation
      var alloc = E.allocate(p.keys);
      var deptRows = alloc.rows.filter(function (r) {
        return r.dept.isRevenue && (r.income || r.expenses) && E.visibleDepts().some(function (d) { return d.code === r.code; });
      });
      var deptChart = U.columns({
        labels: deptRows.map(function (r) { return r.dept.short; }),
        width: 520, height: 250,
        onPick: function (i) { E.setFocusDept(deptRows[i].code); CTS.router.go("pnl-dept"); },
        series: [
          { label: "Revenue", colour: "var(--measure-1)", values: deptRows.map(function (r) { return r.income; }) },
          { label: "Gross profit", colour: "var(--measure-2)", values: deptRows.map(function (r) { return r.grossProfit; }) },
          { label: "Net profit after allocation", colour: "var(--measure-3)", values: deptRows.map(function (r) { return r.netProfit; }) },
        ],
      });
      var deptTable = U.table([
        { key: "d", label: "Department", align: "left" },
        { key: "rev", label: "Revenue", fmt: F.money },
        { key: "gm", label: "GM %", fmt: function (v) { return F.pct(v); } },
        { key: "norm", label: "Norm", align: "left" },
        { key: "np", label: "Net profit", fmt: F.money, cell: U.moneyCell },
      ], deptRows.map(function (r) {
        var norm = r.dept.gmNorm;
        return { d: r.dept.short, rev: r.income, gm: r.gmPct, np: r.netProfit,
                 norm: norm == null ? "not set"
                   : (r.gmPct != null && r.gmPct < norm ? "below " : "at or above ") + F.pct(norm, 0) };
      }));

      return [
        seedBanner(), CTS.periodBar(), U.h1("Dashboard", p.label), tiles,
        U.section("Revenue, gross profit and net profit by month",
          U.figure("FY" + p.fy + ", actual to " + (E.monthIdx[p.rm] || {}).label + " then " + (E.forecastEnabled() ? "forecast" : "budget"),
                   trend, trendTable,
                   E.forecastEnabled()
                     ? "Months up to the reporting month come off the ledger. Months after it come from the forecast, method by method as set on the Forecast page, so the year always reads as twelve and the landing is visible. Budget is the comparison, not the fill."
                     : "Months up to the reporting month come off the ledger. Months after it come from budget, so the year always reads as twelve and the shortfall ahead is visible. That is the Controller Pack's rule.")),
        allocBanner(),
        U.section("Departments, after the overhead split",
          U.figure("Result by department, " + p.label, deptChart, deptTable,
                   "Net profit here is after the Admin overhead pool has been pushed out. The norm column is the margin Monthly Reporting Part 3 says to expect. Click a department to open it; click a month above to make it the reporting month."))
      ];
    },
  };

  /* ================================================ Business Context === */
  P.context = {
    section: "Dashboard", title: "Business Context",
    sub: "What normal looks like, and where the numbers come from.",
    render: function () {
      var cfg = E.cfg;
      var normTable = U.table([
        { key: "d", label: "Department", align: "left" },
        { key: "gm", label: "Gross margin norm" },
        { key: "np", label: "Net profit norm" },
        { key: "util", label: "Utilisation target" },
        { key: "note", label: "What normal looks like", align: "left" },
      ], E.depts.filter(function (d) { return d.code !== "UNALLOCATED"; }).map(function (d) {
        return {
          d: d.short,
          gm: d.gmNorm == null ? h("span.muted", "not set") : F.pct(d.gmNorm, 0),
          np: d.npNorm == null ? h("span.muted", "not set") : F.pct(d.npNorm, 0),
          util: d.utilTarget == null ? h("span.muted", "n/a")
            : h("span", [F.pct(d.utilTarget, 0), " ",
                         d.utilTargetSource === "documented"
                           ? U.flag("good", "documented")
                           : U.flag("warn", "assumed", "Only one utilisation target is written down. The rest carry the same figure until confirmed.")]),
          note: d.normNote + (d.normAssumed ? " (mapping assumed)" : ""),
        };
      }));

      var season = U.table([
        { key: "when", label: "When", align: "left" },
        { key: "what", label: "", align: "left" },
        { key: "note", label: "Why", align: "left" },
      ], cfg.SEASON.map(function (s) {
        return { when: s.months.map(function (m) {
          return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][m - 1];
        }).join(", "), what: s.label, note: s.note };
      }));

      var prov = U.table([
        { key: "what", label: "Figure", align: "left" },
        { key: "src", label: "Where it comes from", align: "left" },
        { key: "state", label: "State", align: "left" },
      ], [
        { what: "Chart of accounts, 190 accounts", src: "Xero, read 3 September 2026", state: U.flag("good", "sourced") },
        { what: "FY26 control totals", src: "Xero profit and loss, accrual, pulled 3 September 2026", state: U.flag("good", "sourced") },
        { what: "August 2026 category totals", src: "The worked example written into the Controller Pack's PL_Check sheet", state: U.flag("good", "sourced") },
        { what: "Department tags", src: "DEPT_TAGS in the Controller Pack build", state: U.flag("good", "sourced") },
        { what: "Risk thresholds", src: "The Controller Pack Setup sheet", state: U.flag("good", "sourced") },
        { what: "3 Way split, 34/33/33", src: "The FY27 budget bridge", state: U.flag("good", "sourced") },
        { what: "Production and integration sub splits", src: "FY23 actuals on the PRD and CONS tabs", state: U.flag("good", "sourced") },
        { what: "Staff and Office Dept splits", src: "Live in the budget workbook, not in this repository", state: U.flag("warn", "placeholder") },
        { what: "Utilisation targets other than consulting", src: "Not written down anywhere", state: U.flag("warn", "assumed") },
        { what: "Which basis each overhead line uses", src: "Labelled per row on the live allocation tab; derived here from the account", state: U.flag("warn", "derived") },
        { what: "Ledger detail, client names, hours", src: "Modelled to fit the real totals", state: E.seed ? U.flag("warn", "seed") : U.flag("good", "loaded") },
      ]);

      return [
        seedBanner(), U.h1("Business Context", "The judgement that does not live in a spreadsheet."),
        U.section("What normal looks like by department", normTable,
          "Monthly Reporting Part 3. This is what tells you whether a number is worth a comment. A consulting loss in most months is normal; production in the 40s is not."),
        U.section("The production calendar", season,
          "Most of what you see in a month is explained by where you are in this calendar."),
        U.section("Where every number comes from", prov,
          "Anything marked placeholder, assumed or derived is a figure this portal had to stand in for. The portal flags them wherever they are used rather than hiding them."),
        U.section("How management want it read", [
          U.note(cfg.COMMENTARY.note),
          h("p", "Top clients carries four headings every time: " + cfg.COMMENTARY.topClientHeadings.join(", ") + ", always against last year."),
        ]),
      ];
    },
  };

  /* ========================================================= P&L ======= */
  function pnlRows(pnl, expanded, toggle) {
    var out = [];
    pnl.cats.forEach(function (cat) {
      if (cat.total === 0 && !cat.liveChildren) return;
      out.push({ level: 0, key: cat.key, label: cat.label, cells: cat.cells,
                 total: cat.total, open: expanded[cat.key], toggle: toggle });
      if (!expanded[cat.key]) return;
      cat.children.forEach(function (sub) {
        if (sub.total === 0) return;
        var skey = cat.key + "|" + sub.sub;
        out.push({ level: 1, key: skey, label: sub.sub, cells: sub.cells,
                   total: sub.total, open: expanded[skey], toggle: toggle });
        if (!expanded[skey]) return;
        sub.children.forEach(function (acc) {
          if (acc.total === 0) return;
          out.push({ level: 2, key: skey + "|" + acc.account, label: acc.account,
                     cells: acc.cells, total: acc.total });
        });
      });
    });
    return out;
  }

  function pnlTable(pnl, opts) {
    opts = opts || {};
    var expanded = CTS.store.get("pnlOpen", { income: true, cos: true, expenses: true });
    function toggle(key) {
      expanded[key] = !expanded[key];
      CTS.store.set("pnlOpen", expanded);
      CTS.router.reload();
    }
    var months = pnl.months;
    var rows = pnlRows(pnl, expanded, toggle);
    var cols = [{ key: "label", label: "Line", align: "left", width: "23rem",
                  value: function (r) {
                    var kids = r.level < 2;
                    return h("span.pnl-label.lvl" + r.level, [
                      kids ? h("button.twisty", {
                        onclick: function (e) { e.stopPropagation(); r.toggle(r.key); },
                        "aria-expanded": !!r.open,
                      }, r.open ? "−" : "+") : h("span.twisty-spacer"),
                      r.label,
                    ]);
                  } }];
    months.forEach(function (m, i) {
      cols.push({ key: "m" + i, label: (E.monthIdx[m] || {}).label || m,
                  cls: E.blendSource(m) === "actual" ? "" : "forecast-col",
                  fmt: F.money,
                  value: function (r) { return signed(r.cells[i], r); } });
    });
    cols.push({ key: "total", label: opts.totalLabel || "Total", cls: "totalcol",
                fmt: F.money, value: function (r) { return signed(r.total, r); } });
    function signed(v, r) { return v; }
    return U.table(cols, rows.map(function (r) {
      r._cls = "pnl-lvl" + r.level;
      return r;
    }), { dense: true });
  }

  P.pnl = {
    section: "Finance", title: "P&L",
    sub: "Category, subcategory, account. Expand any line.",
    render: function () {
      var p = CTS.period();
      var fyKeys = E.monthsOfFY(p.fy);
      var pnl = E.pnlBlend(fyKeys);
      var summary = U.table([
        { key: "l", label: "", align: "left" },
        { key: "v", label: "FY" + p.fy, fmt: F.money },
        { key: "b", label: "Budget", fmt: F.money },
        { key: "d", label: "Variance", fmt: F.money, cell: U.moneyCell },
        { key: "p", label: "%", fmt: function (v) { return F.pct(v); } },
        { key: "r", label: "Risk", align: "left",
          value: function (r) { return h("span.risk.risk-" + U.RISK_CLASS[r.rk], r.rk); } },
      ], (function () {
        var b = E.pnlBudget(fyKeys);
        return [
          ["Income", pnl.totals.income, b.totals.income],
          ["Cost of Sales", pnl.totals.cos, b.totals.cos],
          ["Gross Profit", pnl.totals.grossProfit, b.totals.grossProfit],
          ["Expenses", pnl.totals.expenses, b.totals.expenses],
          ["Other Income", pnl.totals.otherIncome, b.totals.otherIncome],
          ["Net Profit", pnl.totals.netProfit, b.totals.netProfit],
        ].map(function (r) {
          return { l: r[0], v: r[1], b: r[2], d: r[1] - r[2],
                   p: r[2] ? (r[1] - r[2]) / Math.abs(r[2]) : null,
                   rk: E.risk(r[1], r[2]),
                   _cls: /Profit/.test(r[0]) ? "totalrow" : "" };
        });
      })());

      return [
        seedBanner(), CTS.periodBar(), U.h1("Profit & Loss", "FY" + p.fy + ", month by month"),
        U.note("Costs are shown the way the ledger holds them: credit less debit, so a cost is negative and gross profit is income plus cost of sales. Columns after " + (E.monthIdx[p.rm] || {}).label + " are " + (E.forecastEnabled() ? "forecast" : "budget") + ", not actual, and are shaded." + (E.viewDept() ? " This is " + ((E.deptOf[E.viewDept()] || {}).short || E.viewDept()) + " off the ledger by cost centre: overheads are the department's own, and its share of the Admin pool is on P&L by Department." : "")),
        U.section("The year at a glance", summary),
        U.section("Full profit and loss", pnlTable(pnl),
          "Click the plus beside any line to open it. Category opens to subcategory, subcategory opens to the individual Xero accounts."),
      ];
    },
  };

  /* ============================================== P&L by department ==== */
  P["pnl-dept"] = {
    section: "Finance", title: "P&L by Department",
    sub: "Before and after the overhead split.",
    render: function () {
      var p = CTS.period();
      var alloc = E.allocate(p.keys);
      var showAlloc = CTS.store.get("showAlloc", true);
      var scoped = E.visibleDepts();
      var rows = alloc.rows.filter(function (r) {
        return (r.income || r.cos || r.ownExpenses || r.allocated) &&
               scoped.some(function (d) { return d.code === r.code; });
      });

      var cols = [
        { key: "d", label: "Department", align: "left",
          value: function (r) {
            return h("span", [h("span.dot", { style: { background: U.colourOf(r.code) } }), r.dept.short]);
          } },
        { key: "income", label: "Revenue", fmt: F.money },
        { key: "cos", label: "Cost of sales", fmt: F.money, cell: U.moneyCell },
        { key: "grossProfit", label: "Gross profit", fmt: F.money, cell: U.moneyCell },
        { key: "gmPct", label: "GM %", fmt: function (v) { return F.pct(v); },
          cell: function (r) {
            var n = r.dept.gmNorm;
            return n != null && r.gmPct != null && r.gmPct < n ? "neg" : "";
          } },
        { key: "ownExpenses", label: "Own overhead", fmt: F.money, cell: U.moneyCell },
      ];
      if (showAlloc) cols.push({ key: "allocated", label: "Share of Admin", fmt: F.money, cell: U.moneyCell });
      cols.push({ key: "expenses", label: "Total overhead", fmt: F.money, cell: U.moneyCell });
      cols.push({ key: "netProfit", label: "Net profit", fmt: F.money, cell: U.moneyCell });
      cols.push({ key: "npPct", label: "NP %", fmt: function (v) { return F.pct(v); } });

      var table = U.table(cols, rows.map(function (r) {
        if (!showAlloc) {
          return Object.assign({}, r, { expenses: r.ownExpenses,
            netProfit: r.grossProfit + r.otherIncome + r.ownExpenses + r.otherExpenses });
        }
        return r;
      }));

      var totals = rows.reduce(function (a, r) {
        a.income += r.income; a.gp += r.grossProfit; a.np += r.netProfit; return a;
      }, { income: 0, gp: 0, np: 0 });

      var mix = U.columns({
        labels: rows.filter(function (r) { return r.dept.isRevenue; }).map(function (r) { return r.dept.short; }),
        width: 520, height: 230,
        series: [{ label: "Revenue", colour: "var(--measure-1)",
                   values: rows.filter(function (r) { return r.dept.isRevenue; }).map(function (r) { return r.income; }) }],
        tipfmt: F.dollars,
      });

      var only = E.visibleDepts().filter(function (d) { return d.isRevenue; });
      var cover = only.length === 1 ? U.cover(only[0].code) : null;
      return [
        seedBanner(), CTS.periodBar(), cover, allocBanner(),
        U.h1("P&L by Department", p.label),
        h("div.toolbar", [
          h("label.checkbox", [
            h("input", { type: "checkbox", checked: showAlloc,
              onchange: function (e) { CTS.store.set("showAlloc", e.target.checked); CTS.router.reload(); } }),
            " Push the Admin overhead pool out across the departments",
          ]),
        ]),
        U.note(showAlloc
          ? "Own overhead is what is coded straight to the department. Share of Admin is its slice of the pool. They are kept as two numbers on purpose: the documented way this goes wrong is a departmental tab picking up its own overhead twice, and two separate columns make that impossible to do quietly."
          : "Showing raw departmental result, nothing pushed out. Admin still carries the whole overhead pool."),
        U.section("Departmental result, " + p.label, table),
        U.section("Where the revenue sits",
          U.figure("Revenue by department, " + p.label, mix,
            U.table([
              { key: "d", label: "Department", align: "left" },
              { key: "v", label: "Revenue", fmt: F.money },
              { key: "s", label: "Share", fmt: function (v) { return F.pct(v); } },
            ], rows.filter(function (r) { return r.dept.isRevenue; }).map(function (r) {
              return { d: r.dept.short, v: r.income, s: totals.income ? r.income / totals.income : null };
            }), { dense: true }))),
      ];
    },
  };

  /* ================================================ P&L spread ========= */
  P["pnl-spread"] = {
    section: "Finance", title: "P&L Spread",
    sub: "Twelve months across, two years deep.",
    render: function () {
      var fy = E.currentFY();
      var keys = E.monthsOfFY(fy - 1).concat(E.monthsOfFY(fy));
      var pnl = E.pnlBlend(keys);
      var labels = keys.map(function (k) { return E.monthIdx[k].label; });
      var chart = U.lines({
        labels: labels, width: 860, height: 280, zero: true,
        series: [
          { label: "Revenue", colour: "var(--measure-1)", values: pnl.income },
          { label: "Gross profit", colour: "var(--measure-2)", values: pnl.grossProfit },
          { label: "Net profit", colour: "var(--measure-3)", values: pnl.netProfit },
        ],
      });
      var tbl = U.table([
        { key: "m", label: "Month", align: "left" },
        { key: "fy", label: "FY", align: "left" },
        { key: "src", label: "Basis", align: "left" },
        { key: "rev", label: "Revenue", fmt: F.money },
        { key: "cos", label: "Cost of sales", fmt: F.money, cell: U.moneyCell },
        { key: "gp", label: "Gross profit", fmt: F.money },
        { key: "gm", label: "GM %", fmt: function (v) { return F.pct(v); } },
        { key: "exp", label: "Overheads", fmt: F.money, cell: U.moneyCell },
        { key: "np", label: "Net profit", fmt: F.money, cell: U.moneyCell },
      ], keys.map(function (k, i) {
        return { m: E.monthIdx[k].label, fy: "FY" + E.monthIdx[k].fy,
                 src: E.blendSource(k),
                 rev: pnl.income[i], cos: pnl.cos[i], gp: pnl.grossProfit[i],
                 gm: pnl.gmPct[i], exp: pnl.expenses[i], np: pnl.netProfit[i],
                 _cls: E.blendSource(k) === "actual" ? "" : "forecast" };
      }), { dense: true });
      return [
        seedBanner(), CTS.periodBar(), U.h1("P&L Spread", "FY" + (fy - 1) + " and FY" + fy),
        U.section("Two years, month by month", U.figure("Revenue, gross profit and net profit", chart, null)),
        U.section("The workings", tbl,
          "Every figure here is the ledger re-aggregated in the browser. Nothing is copied from another sheet, so a change to the reporting month moves the whole table."),
      ];
    },
  };

  /* ============================================ overhead allocation ==== */
  P.allocation = {
    section: "Finance", title: "Overhead Allocation",
    sub: "The three split bases, line by line.",
    render: function () {
      var p = CTS.period();
      var alloc = E.allocate(p.keys);
      var s = E.subSplits();
      var targets = ["ONSITE", "PRODUCTION", "VIDEO", "INTEGRATION", "CONSULTING"];

      var basisTable = U.table([
        { key: "l", label: "Basis", align: "left" },
        { key: "prdvid", label: "Production + Video", fmt: function (v) { return F.pct(v); } },
        { key: "ons", label: "Onsite", fmt: function (v) { return F.pct(v); } },
        { key: "consint", label: "Consulting + Integration", fmt: function (v) { return F.pct(v); } },
        { key: "src", label: "", align: "left" },
        { key: "note", label: "What it means", align: "left" },
      ], E.cfg.SPLIT_BASES.map(function (b) {
        return { l: b.label, prdvid: b.groups["PRD/VID"], ons: b.groups["ONS"],
                 consint: b.groups["CONS/INT"],
                 src: b.source === "confirmed" ? U.flag("good", "confirmed") : U.flag("warn", "placeholder"),
                 note: b.note };
      }));

      var cols = [
        { key: "sub", label: "Overhead line", align: "left", width: "16rem" },
        { key: "basisLabel", label: "Basis", align: "left",
          value: function (r) {
            return h("select.control.small", {
              onchange: function (e) { E.setBasisFor(r.sub, e.target.value); CTS.router.reload(); },
            }, E.cfg.SPLIT_BASES.map(function (b) {
              return h("option", { value: b.code, selected: b.code === r.basis }, b.label);
            }));
          } },
        { key: "amount", label: "In Admin", fmt: F.money, cell: U.moneyCell },
      ];
      targets.forEach(function (d) {
        cols.push({ key: d, label: E.deptOf[d].short, fmt: F.money, cell: U.moneyCell,
                    value: function (r) { return r.split[d]; } });
      });
      var workTable = U.table(cols, alloc.workings.slice().sort(function (a, b) {
        return a.amount - b.amount;
      }));

      var totalsRow = targets.map(function (d) {
        return alloc.workings.reduce(function (a, w) { return a + w.split[d]; }, 0);
      });

      var chart = U.columns({
        labels: targets.map(function (d) { return E.deptOf[d].short; }),
        width: 520, height: 220,
        series: [{ label: "Share of the Admin pool", colour: "var(--measure-3)",
                   values: totalsRow.map(function (v) { return -v; }) }],
      });

      return [
        seedBanner(), CTS.periodBar(), allocBanner(),
        U.h1("Overhead Allocation", p.label),
        U.note("Anything sitting in Admin is pushed out across the departments. There is not one split, there are three, and each overhead line is labelled with the one that applies to it. Change a basis on the dropdown and every allocated figure in the portal moves with it."),
        U.section("The three bases", basisTable),
        U.section("The two sub splits", [
          h("div.fieldrow", [
            h("label", "Production share of production and video"),
            h("input.control", { type: "number", step: "0.01", min: "0", max: "1", value: s.prd,
              onchange: function (e) { CTS.store.set("subsplit.prd", +e.target.value); CTS.router.reload(); } }),
            h("span.muted", "video takes " + F.pct(1 - s.prd, 0)),
          ]),
          h("div.fieldrow", [
            h("label", "Integration share of consulting and integration"),
            h("input.control", { type: "number", step: "0.01", min: "0", max: "1", value: s.int,
              onchange: function (e) { CTS.store.set("subsplit.int", +e.target.value); CTS.router.reload(); } }),
            h("span.muted", "consulting takes " + F.pct(1 - s.int, 0)),
          ]),
          U.source(E.cfg.SUB_SPLITS.note + " Defaults are the FY23 actuals on the PRD and CONS tabs."),
        ]),
        U.section("Every overhead line and where it lands", workTable,
          "Pool of " + F.dollars(-alloc.pool) + " over " + alloc.workings.length + " lines. The basis against each line is derived from the account here; on the live allocation tab it is labelled per row, so correct any that disagree and it sticks."),
        U.section("Total pushed out", U.figure("Share of the Admin pool by department", chart,
          U.table([
            { key: "d", label: "Department", align: "left" },
            { key: "v", label: "Share received", fmt: F.money },
            { key: "p", label: "% of pool", fmt: function (v) { return F.pct(v); } },
          ], targets.map(function (d, i) {
            return { d: E.deptOf[d].short, v: totalsRow[i],
                     p: alloc.pool ? totalsRow[i] / alloc.pool : null };
          }), { dense: true }),
          "The staff basis follows headcount, and consulting is nearly empty. The day someone is hired into consulting that percentage has to be changed by hand: nothing will prompt you, and until it is changed consulting carries almost no overhead.")),
        U.section("Two standing rules", [
          U.note("Do not change a split mid-year. Changing it in December means the six months already reported were built on a different split and no longer compare. If it is revisited, it takes effect at a new financial year.", "warn"),
          U.note("Whatever you change here has to change in the budget as well, or the budget and the actuals will never reconcile.", "warn"),
        ]),
      ];
    },
  };
})();

/* ===================================================================== *
 * Pages: control, budget, revenue, clients, utilisation, ledger, admin.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, h = CTS.h, F = CTS.fmt, E = CTS.engine, U = CTS.ui;
  var P = CTS.pages, seedBanner = CTS.seedBanner, allocBanner = CTS.allocBanner;

  /* ================================================== P&L control ====== */
  P.control = {
    section: "Finance", title: "P&L Control",
    sub: "Does the ledger agree with the P&L, and is the coding clean.",
    render: function () {
      var p = CTS.period();
      var check = E.plCheck(p.keys);
      var q = E.quality;
      var ok = check.every(function (r) { return r.ok; });

      var checkTable = U.table([
        { key: "label", label: "Category", align: "left" },
        { key: "gl", label: "Per the ledger", fmt: F.money },
        { key: "pl", label: "Per the P&L", fmt: F.money },
        { key: "diff", label: "Difference", fmt: F.money, cell: U.moneyCell },
        { key: "ok", label: "Status", align: "left",
          value: function (r) {
            return h("span.risk.risk-" + (r.ok ? "good" : "critical"), r.ok ? "OK" : "CHECK");
          } },
      ], check.concat([(function () {
        var gl = check.reduce(function (a, r) { return a + r.gl; }, 0);
        var pl = check.reduce(function (a, r) { return a + r.pl; }, 0);
        return { label: "Net profit", gl: gl, pl: pl, diff: gl - pl,
                 ok: Math.abs(gl - pl) < 100, _cls: "totalrow" };
      })()]));

      var unalloc = E.gl.filter(function (r) { return r.dept === "UNALLOCATED"; });
      var noContact = E.gl.filter(function (r) { return r.cat === "income" && !r.contact; });

      var quality = U.table([
        { key: "what", label: "Check", align: "left" },
        { key: "n", label: "Count" },
        { key: "state", label: "", align: "left" },
        { key: "note", label: "What to do", align: "left" },
      ], [
        { what: "Ledger lines loaded", n: q.rows, state: "",
          note: "Earliest " + F.date(q.minDate) + ", latest " + F.date(q.maxDate) },
        { what: "Lines with no department", n: q.noDept,
          state: q.noDept ? U.flag("warn", "fix") : U.flag("good", "clean"),
          note: "Cost Centres blank and no bracket tag on the job number. Code them in Xero, then re-export." },
        { what: "Lines on an account that is not in the chart", n: q.unknownAcct,
          state: q.unknownAcct ? U.flag("warn", "fix") : U.flag("good", "clean"),
          note: "A new Xero account has to be added to the chart or its lines are rejected." },
        { what: "Revenue lines with no contact", n: q.noContact,
          state: q.noContact ? U.flag("warn", "fix") : U.flag("good", "clean"),
          note: "A manual journal carries no contact, so those rows arrive blank. Top clients is built off the contact, so fill them in before you paste." },
      ]);

      return [
        seedBanner(), CTS.periodBar(),
        U.h1("P&L Control", p.label),
        h("div.banner.banner-" + (ok ? "good" : "warn"), [
          h("strong", ok ? "The ledger agrees with the P&L. " : "The ledger and the P&L disagree. "),
          ok ? "Every category is within the one dollar tolerance the Controller Pack uses."
             : "Find the difference before you report anything off this month.",
        ]),
        U.section("Does the ledger agree with the P&L", checkTable,
          "The same control as the Controller Pack's PL_Check sheet. Under a dollar reads OK; the difference is normally a paste that did not cover every month, or a row that slipped out of alignment."),
        U.section("Coding quality", quality),
        unalloc.length ? U.section("Lines with no department (" + unalloc.length + ")",
          U.table([
            { key: "date", label: "Date", align: "left", fmt: F.date },
            { key: "account", label: "Account", align: "left" },
            { key: "contact", label: "Contact", align: "left" },
            { key: "amount", label: "Amount", fmt: F.money, cell: U.moneyCell },
            { key: "jobNo", label: "Job number", align: "left" },
            { key: "costCentre", label: "Cost centre", align: "left" },
          ], unalloc.slice(0, 40)),
          "These fall to UNALLOCATED and sit outside every departmental figure in the portal. They still count in the company P&L, which is why the two can look inconsistent.") : null,
      ];
    },
  };

  /* ================================================ budget vs actual === */
  P.bva = {
    section: "Budget", title: "Budget vs Actual",
    sub: "Account by account, with the risk flag.",
    render: function () {
      var p = CTS.period();
      var act = E.pnlActual(p.keys), bud = E.pnlBudget(p.keys);
      var filter = CTS.store.get("bvaFilter", "all");

      var rows = [];
      act.cats.forEach(function (cat) {
        var bcat = bud.by[cat.key];
        cat.children.forEach(function (sub) {
          var bsub = (bcat.children || []).filter(function (x) { return x.sub === sub.sub; })[0];
          var b = bsub ? bsub.total : 0;
          if (!sub.total && !b) return;
          var risk = E.risk(sub.total, b);
          rows.push({ cat: cat.label, sub: sub.sub, actual: sub.total, budget: b,
                      variance: sub.total - b,
                      pct: b ? (sub.total - b) / Math.abs(b) : null, risk: risk });
        });
      });
      if (filter !== "all") rows = rows.filter(function (r) { return r.risk === filter; });
      rows.sort(function (a, b) {
        return (E.riskRank[b.risk] - E.riskRank[a.risk]) || (Math.abs(b.variance) - Math.abs(a.variance));
      });

      var bars = ["High", "Medium", "Low", "No Activity"].map(function (k) {
        return { k: k, n: rows.filter(function (r) { return r.risk === k; }).length };
      });

      // A net profit bridge, so the steps actually reconcile: every category
      // variance is in the same sign convention as the P&L, and budget net
      // profit plus all of them equals actual net profit.
      var other = (act.totals.otherIncome - bud.totals.otherIncome) +
                  (act.totals.otherExpenses - bud.totals.otherExpenses);
      var waterfall = U.waterfall({
        width: 760,
        steps: [
          { label: "Budget net profit", value: bud.totals.netProfit, kind: "base" },
          { label: "Revenue", value: act.totals.income - bud.totals.income, kind: "delta" },
          { label: "Cost of sales", value: act.totals.cos - bud.totals.cos, kind: "delta" },
          { label: "Overheads", value: act.totals.expenses - bud.totals.expenses, kind: "delta" },
          { label: "Other", value: other, kind: "delta" },
          { label: "Actual net profit", value: act.totals.netProfit, kind: "total" },
        ],
      });

      return [
        seedBanner(), CTS.periodBar(),
        U.h1("Budget vs Actual", p.label),
        h("div.toolbar", [
          h("div.seg", ["all", "High", "Medium", "Low"].map(function (k) {
            return h("button.seg-btn" + (k === filter ? ".on" : ""), {
              onclick: function () { CTS.store.set("bvaFilter", k); CTS.router.reload(); },
            }, k === "all" ? "All" : k);
          })),
          h("div.toolbar-right", bars.map(function (b) {
            return h("span.chip.chip-" + U.RISK_CLASS[b.k], b.k + " " + b.n);
          })),
        ]),
        U.note("The risk flag is the Controller Pack's rule exactly, including the order of the tests. The materiality floor of " + F.dollars(E.cfg.RISK.materialityFloor * 100) + " is checked before the percentage, so a big percentage on a small dollar variance is still Low."),
        U.section("Subcategory variances", U.table([
          { key: "cat", label: "Category", align: "left" },
          { key: "sub", label: "Line", align: "left" },
          { key: "actual", label: "Actual", fmt: F.money, cell: U.moneyCell },
          { key: "budget", label: "Budget", fmt: F.money, cell: U.moneyCell },
          { key: "variance", label: "Variance", fmt: F.money, cell: U.moneyCell },
          { key: "pct", label: "%", fmt: function (v) { return F.pct(v); } },
          { key: "risk", label: "Risk", align: "left",
            value: function (r) { return h("span.risk.risk-" + U.RISK_CLASS[r.risk], r.risk); } },
        ], rows, { dense: true })),
        U.section("Where the variance came from",
          U.figure("Budget net profit bridged to actual net profit", waterfall, null,
            "Each step is that category's variance in the P&L's own sign convention, so a green bar always helped profit and a red one always hurt it. Budget net profit plus the four steps equals actual net profit exactly.")),
      ];
    },
  };

  /* ====================================================== risk actions = */
  P.actions = {
    section: "Budget", title: "Actions",
    sub: "Only the lines worth a comment.",
    render: function () {
      var p = CTS.period();
      var act = E.pnlActual(p.keys), bud = E.pnlBudget(p.keys);
      var alloc = E.allocate(p.keys);
      var items = [];

      act.cats.forEach(function (cat) {
        var bcat = bud.by[cat.key];
        cat.children.forEach(function (sub) {
          var bsub = (bcat.children || []).filter(function (x) { return x.sub === sub.sub; })[0];
          var b = bsub ? bsub.total : 0;
          var risk = E.risk(sub.total, b);
          if (risk !== "High") return;
          items.push({ kind: "Variance", what: cat.label + ": " + sub.sub,
                       detail: F.dollars(sub.total) + " against budget " + F.dollars(b),
                       amount: sub.total - b });
        });
      });

      alloc.rows.forEach(function (r) {
        if (!r.dept.isRevenue || r.dept.gmNorm == null) return;
        var norm = r.dept.gmNorm;
        if (norm != null && r.gmPct != null && r.gmPct < norm) {
          items.push({ kind: "Below norm",
                       what: r.dept.short + " gross margin " + F.pct(r.gmPct),
                       detail: "Part 3 expects " + F.pct(norm, 0) + " or better. " + r.dept.normNote,
                       amount: Math.round((norm - r.gmPct) * r.income) });
        }
      });

      var util = E.utilFor(p.keys);
      util.rows.forEach(function (r) {
        if (r.target == null || r.util == null) return;
        if (r.util < r.target - 0.05) {
          items.push({ kind: "Under target",
                       what: r.dept.short + " utilisation " + F.pct(r.util),
                       detail: "Target " + F.pct(r.target, 0) + ". Read a small negative before flagging it: for a team whose hours are almost all chargeable a small gap usually just means leave. A sustained gap is the one worth raising.",
                       amount: 0 });
        }
      });

      if (E.quality.noDept) {
        items.push({ kind: "Coding", what: E.quality.noDept + " ledger lines carry no department",
                     detail: "They sit outside every departmental figure. Fix the cost centre in Xero and re-export.", amount: 0 });
      }
      var s = E.allocationIsSourced();
      if (!s.ok) {
        items.push({ kind: "Data gap", what: "Overhead split percentages are placeholders",
                     detail: "The " + s.placeholders.join(" and ") + " bases are not in this repository. Every allocated departmental result depends on them.", amount: 0 });
      }

      items.sort(function (a, b) { return Math.abs(b.amount) - Math.abs(a.amount); });

      return [
        seedBanner(), CTS.periodBar(), U.h1("Actions", p.label),
        U.note("Everything the portal thinks is worth a comment this period, biggest first. Commentary explains the variance rather than restating it, and management prefer dollar value against dollar value."),
        items.length ? U.table([
          { key: "kind", label: "Kind", align: "left",
            value: function (r) { return h("span.chip.chip-" + (r.kind === "Variance" ? "warning" : r.kind === "Coding" || r.kind === "Data gap" ? "critical" : "muted"), r.kind); } },
          { key: "what", label: "What", align: "left" },
          { key: "detail", label: "Detail", align: "left" },
          { key: "amount", label: "Effect", fmt: function (v) { return v ? F.money(v) : "-"; }, cell: U.moneyCell },
        ], items) : U.note("Nothing above the thresholds this period."),
      ];
    },
  };

  /* ====================================================== forecast ===== */
  P.forecast = {
    section: "Budget", title: "Forecast",
    sub: "Actual to date, forecast to year end, budget beside it.",
    render: function () {
      var p = CTS.period();
      var fyKeys = E.monthsOfFY(p.fy);
      if (!E.forecastEnabled()) {
        return [
          seedBanner(), CTS.periodBar(), U.h1("Forecast", "FY" + p.fy),
          h("div.banner.banner-warn", [h("strong", "No forecast is loaded. "),
            "Put 09 Forecast.xlsx in the templates folder and run Build. Until then the months after the reporting month are filled with budget, which is how the portal worked before the forecast existed."]),
        ];
      }
      var fc = E.forecast();
      var fcKeys = fyKeys.filter(function (k) { return E.blendSource(k) === "forecast"; });
      var actKeys = fyKeys.filter(function (k) { return E.blendSource(k) === "actual"; });
      var budKeys = fyKeys.filter(function (k) { return E.blendSource(k) === "budget"; });
      var land = E.pnlBlend(fyKeys), act = E.pnlActual(actKeys), fcast = E.pnlForecast(fcKeys), bud = E.pnlBudget(fyKeys);
      var lt = land.totals, at = act.totals, ft = fcast.totals, bt = bud.totals;
      var firstFc = fcKeys.length ? E.monthIdx[fcKeys[0]].label : null, lastFc = fcKeys.length ? E.monthIdx[fcKeys[fcKeys.length - 1]].label : null;
      var a = E.forecastAssumptions();

      var tiles = h("div.tiles", [
        U.tile({ label: "FY" + p.fy + " revenue", value: F.dollars(lt.income),
                 sub: (lt.income - bt.income >= 0 ? "+" : "−") + F.money(Math.abs(lt.income - bt.income)) + " vs budget",
                 note: "Actual " + F.dollars(at.income) + ", forecast " + F.dollars(ft.income) }),
        U.tile({ label: "FY" + p.fy + " gross profit", value: F.dollars(lt.grossProfit),
                 sub: F.pct(lt.income ? lt.grossProfit / lt.income : null) + " margin",
                 note: "Budget " + F.pct(bt.income ? bt.grossProfit / bt.income : null) }),
        U.tile({ label: "FY" + p.fy + " net profit", value: F.dollars(lt.netProfit),
                 tone: lt.netProfit >= bt.netProfit ? "good" : "critical",
                 sub: (lt.netProfit - bt.netProfit >= 0 ? "+" : "−") + F.money(Math.abs(lt.netProfit - bt.netProfit)) + " vs budget",
                 note: "Budget " + F.dollars(bt.netProfit) }),
        U.tile({ label: "Forecast months", value: String(fcKeys.length),
                 sub: firstFc ? firstFc + " to " + lastFc : "none in this year",
                 note: fc.overrides.length + " typed override" + (fc.overrides.length === 1 ? "" : "s") + ", " + fc.history.months + " months of history" }),
      ]);

      var landing = U.table([
        { key: "l", label: "", align: "left" },
        { key: "a", label: "Actual to " + ((E.monthIdx[p.rm] || {}).label || p.rm), fmt: F.money, cell: U.moneyCell },
        { key: "f", label: "Forecast to year end", fmt: F.money, cell: U.moneyCell },
        { key: "t", label: "Full year", fmt: F.money, cell: U.moneyCell },
        { key: "b", label: "Budget", fmt: F.money, cell: U.moneyCell },
        { key: "d", label: "Variance", fmt: F.money, cell: U.moneyCell },
        { key: "p", label: "%", fmt: function (v) { return F.pct(v); } },
        { key: "r", label: "Risk", align: "left", value: function (r) { return h("span.risk.risk-" + U.RISK_CLASS[r.rk], r.rk); } },
      ], [
        ["Income", "income"], ["Cost of Sales", "cos"], ["Gross Profit", "grossProfit"],
        ["Expenses", "expenses"], ["Other Income", "otherIncome"], ["Other Expenses", "otherExpenses"], ["Net Profit", "netProfit"],
      ].map(function (r) {
        var k = r[1];
        return { l: r[0], a: at[k], f: ft[k], t: lt[k], b: bt[k], d: lt[k] - bt[k],
                 p: bt[k] ? (lt[k] - bt[k]) / Math.abs(bt[k]) : null, rk: E.risk(lt[k], bt[k]),
                 _cls: /Profit/.test(r[0]) ? "totalrow" : "" };
      }));

      var labels = fyKeys.map(function (k) { return E.monthIdx[k].label; });
      var trend = U.columns({
        labels: labels, width: 760, height: 250,
        series: [
          { label: "Revenue", colour: "var(--measure-1)", values: land.income },
          { label: "Gross profit", colour: "var(--measure-2)", values: land.grossProfit },
          { label: "Net profit", colour: "var(--measure-3)", values: land.netProfit },
        ],
      });
      var monthTable = U.table([
        { key: "m", label: "Month", align: "left" },
        { key: "src", label: "", align: "left" },
        { key: "rev", label: "Revenue", fmt: F.money },
        { key: "brev", label: "Budget revenue", fmt: F.money },
        { key: "gp", label: "Gross profit", fmt: F.money },
        { key: "np", label: "Net profit", fmt: F.money, cell: U.moneyCell },
        { key: "bnp", label: "Budget net profit", fmt: F.money, cell: U.moneyCell },
      ], fyKeys.map(function (k, i) {
        var src = E.blendSource(k);
        return { m: E.monthIdx[k].label, src: src === "actual" ? "" : src,
                 rev: land.income[i], brev: bud.income[i], gp: land.grossProfit[i], np: land.netProfit[i], bnp: bud.netProfit[i],
                 _cls: src === "actual" ? "" : "forecast" };
      }), { dense: true });

      // departments: actual to date, forecast to year end, budget
      var depts = E.visibleDepts().filter(function (d) { return d.isRevenue; });
      var fInc = E.forecastByDept(fcKeys, "income"), fCos = E.forecastByDept(fcKeys, "cos");
      var bInc = E.budgetByDept(fyKeys, "income"), bCos = E.budgetByDept(fyKeys, "cos");
      var bbInc = E.budgetByDept(budKeys, "income"), bbCos = E.budgetByDept(budKeys, "cos");
      var deptRows = depts.map(function (d) {
        var ai = E.sumMonths((E.idx.deptCatMonth[d.code] || {}).income, actKeys);
        var ac = E.sumMonths((E.idx.deptCatMonth[d.code] || {}).cos, actKeys);
        var ti = ai + (fInc[d.code] || 0) + (bbInc[d.code] || 0), tc = ac + (fCos[d.code] || 0) + (bbCos[d.code] || 0);
        var bi = bInc[d.code] || 0, bc = bCos[d.code] || 0;
        return { d: d.short, code: d.code, ai: ai, fi: fInc[d.code] || 0, ti: ti, bi: bi, vi: ti - bi,
                 gm: ti ? (ti + tc) / ti : null, bgm: bi ? (bi + bc) / bi : null, risk: E.risk(ti, bi) };
      });
      var deptTable = U.table([
        { key: "d", label: "Department", align: "left",
          value: function (r) { return h("span", [h("span.dot", { style: { background: U.colourOf(r.code) } }), r.d]); } },
        { key: "ai", label: "Revenue to date", fmt: F.money },
        { key: "fi", label: "Forecast to year end", fmt: F.money },
        { key: "ti", label: "Full year", fmt: F.money },
        { key: "bi", label: "Budget", fmt: F.money },
        { key: "vi", label: "Variance", fmt: F.money, cell: U.moneyCell },
        { key: "gm", label: "GM % full year", fmt: function (v) { return F.pct(v); } },
        { key: "bgm", label: "GM % budget", fmt: function (v) { return F.pct(v); } },
        { key: "risk", label: "Risk", align: "left", value: function (r) { return h("span.risk.risk-" + U.RISK_CLASS[r.risk], r.risk); } },
      ], deptRows);

      // how each line is made
      var SRC = { account: "account line", subcategory: "subcategory line", category: "category line", "default": "default", mixed: "mixed", pipeline: "Pipeline tab" };
      var lines = E.forecastLines(fcKeys).filter(function (r) { return r.total || r.source !== "default"; });
      var methodTable = U.table([
        { key: "cat", label: "Category", align: "left" },
        { key: "sub", label: "Line", align: "left" },
        { key: "methodLabel", label: "Method", align: "left",
          value: function (r) { return h("span.chip.chip-" + (r.method === "fixed" ? "warning" : r.method === "mixed" ? "muted" : "good"), r.methodLabel || r.method); } },
        { key: "param", label: "Parameter", align: "left",
          value: function (r) { return r.param == null ? h("span.muted", r.method === "revenue_pct" ? "measured" : r.method === "runrate" ? a.runRate + " months" : r.method === "seasonal" ? "growth from Assumptions" : r.method === "fixed" ? "no amount typed, so run rate" : "") : String(r.param); } },
        { key: "source", label: "Set by", align: "left", value: function (r) { return SRC[r.source] || r.source; } },
        { key: "accounts", label: "Accounts" },
        { key: "total", label: "Forecast to year end", fmt: F.money, cell: U.moneyCell },
      ], lines, { dense: true });

      var overrideTable = fc.overrides.length ? U.table([
        { key: "month", label: "Month", align: "left", value: function (r) { return (E.monthIdx[r.month] || {}).label || r.month; } },
        { key: "dept", label: "Department", align: "left", value: function (r) { return r.dept || "Company"; } },
        { key: "account", label: "Account", align: "left" },
        { key: "cents", label: "Typed", fmt: F.money, cell: U.moneyCell },
        { key: "was", label: "Method gave", fmt: F.money, cell: U.moneyCell },
        { key: "note", label: "Why", align: "left" },
      ], fc.overrides, { dense: true }) : U.note("No typed overrides. Every forecast cell is the method's own figure.", "muted");
      var unapplied = fc.unapplied.length ? h("div.banner.banner-warn", [
        h("strong", fc.unapplied.length + " override" + (fc.unapplied.length > 1 ? "s" : "") + " not applied. "),
        fc.unapplied.map(function (u) { return (u.account || "?") + " " + (u.month || "") + ": " + u.why; }).join("; ") + ".",
      ]) : null;

      var growth = Object.keys(a.growth).map(function (k) { return k + " " + F.pct(+a.growth[k], 1); }).join(", ");
      return [
        seedBanner(), CTS.periodBar(), U.h1("Forecast", "FY" + p.fy + ", " + E.blendLabel()),
        U.note("The months after " + ((E.monthIdx[p.rm] || {}).label || p.rm) + " are forecast, line by line, on the method set for each line in 09 Forecast.xlsx. Budget stays as the comparison. History is the ledger by department, so the departmental forecasts add up to the company one. Horizon " + a.horizon + " months, run rate over " + a.runRate + (growth ? ", growth on last year: " + growth : "") + "."),
        tiles,
        U.section("Full year landing", landing,
          "Actual to the reporting month plus forecast to year end. The risk flag is the Controller Pack's rule applied to the full year against budget."),
        U.section("Month by month",
          U.figure("FY" + p.fy + " revenue, gross profit and net profit, " + E.blendLabel(), trend, monthTable,
                   "Forecast months are shaded in the table. Budget revenue and budget net profit sit beside each month for the comparison.")),
        U.section("By department", deptTable,
          "Revenue to date is the ledger. Forecast to year end is the department's own history through the method. Budget by department is derived the way it is everywhere else in the portal, on each account's prior year split."),
        U.section("How each line is forecast", methodTable,
          "Seasonal is the same month last year grown by the assumption. Run rate is the average of recent months. % of revenue follows the department's forecast revenue at the measured or typed rate. Fixed is typed. Change a method on the Methods tab of 09 Forecast.xlsx and rebuild."),
        U.section("Typed overrides", [unapplied, overrideTable],
          "An override replaces the method's figure for that account and month. Typed as the P&L shows it: income positive, a cost positive. A blank department applies to the company and is spread on the method's own split."),
      ];
    },
  };

  /* ========================================================== cash ===== */
  P.cash = {
    section: "Cash", title: "Cash Flow",
    sub: "Twelve months ahead, off the forecast.",
    render: function () {
      var p = CTS.period();
      if (!E.cashEnabled()) {
        return [seedBanner(), CTS.periodBar(), U.h1("Cash Flow", "Twelve months ahead"),
          h("div.banner.banner-warn", [h("strong", "No cash file is loaded. "),
            "Fill 10 Cash and Commitments.xlsx, at least the Bank balances tab for the reporting month, and run Build."])];
      }
      if (!E.forecastEnabled()) {
        return [seedBanner(), CTS.periodBar(), U.h1("Cash Flow", "Twelve months ahead"),
          h("div.banner.banner-warn", [h("strong", "The cash flow needs the forecast. "), "Turn the forecast on in 09 Forecast.xlsx and rebuild."])];
      }
      var c = E.withoutView(function () { return E.cashflow(); });
      if (!c) return [seedBanner(), CTS.periodBar(), U.h1("Cash Flow", "Twelve months ahead"), U.note("No forecast months after the reporting month.")];
      var viewNote = E.viewDept() ? h("div.banner.banner-warn", [h("strong", "Cash is company wide. "), "Bank accounts, tax and commitments belong to CTS, not a department, so this page shows the company even with " + ((E.deptOf[E.viewDept()] || {}).short || E.viewDept()) + " picked."]) : null;
      var S = c.settings, rmLabel = (E.monthIdx[p.rm] || {}).label || p.rm;
      var noOpening = !c.balances.length;

      var tiles = h("div.tiles", [
        U.tile({ label: "Cash at " + rmLabel, value: F.dollars(c.opening), tone: noOpening ? "critical" : null,
                 sub: noOpening ? "no bank balance typed for this month" : c.balances.length + " account" + (c.balances.length > 1 ? "s" : "") + " at " + ((E.monthIdx[c.openingMonth] || {}).label || c.openingMonth),
                 note: "Bank balances tab of 10 Cash and Commitments" }),
        U.tile({ label: "Low point", value: F.dollars(c.low.closing), tone: c.low.closing < 0 ? "critical" : c.low.closing < c.opening * 0.5 ? "warning" : "good",
                 sub: c.low.label, note: S.facility ? "Headroom " + F.dollars(c.low.headroom) + " with the facility" : "No facility set" }),
        U.tile({ label: "Cash in twelve months", value: F.dollars(c.closing), tone: c.closing >= c.opening ? "good" : "warning",
                 sub: (c.closing - c.opening >= 0 ? "+" : "−") + F.money(Math.abs(c.closing - c.opening)) + " over the year",
                 note: E.monthIdx[c.months[c.months.length - 1]].label }),
        U.tile({ label: "Credit cards owing", value: F.dollars(c.cardOwing),
                 sub: c.cards.length ? c.cards.length + " card" + (c.cards.length > 1 ? "s" : "") + ", limit " + F.dollars(c.cardLimit) : "no cards typed",
                 note: c.cardsMonth ? "At " + ((E.monthIdx[c.cardsMonth] || {}).label || c.cardsMonth) + (S.cardsPaidInFull ? ", cleared next month" : "") : "Credit cards tab" }),
      ]);

      var labels = c.rows.map(function (r) { return r.label; });
      var chart = U.lines({
        labels: labels, width: 760, height: 260,
        series: [{ label: "Closing cash", colour: "var(--measure-1)", values: c.rows.map(function (r) { return r.closing; }) }]
          .concat(S.facility ? [{ label: "Facility limit", colour: "var(--measure-3)", values: c.rows.map(function () { return -S.facility; }) }] : []),
      });

      var LINES = [
        ["Opening cash", "opening", true], ["Receipts from customers", "receipts"], ["Payments to suppliers", "payments"],
        ["Wages and super", "wages"], ["GST to the ATO", "gst"], ["Super paid quarterly", "sup"], ["PAYG instalment", "payg"],
        ["Commitment timing", "timing"], ["Loans, capital, tax and distributions", "outside"], ["Credit cards cleared", "cards"],
        ["Net movement", "net", true], ["Closing cash", "closing", true],
      ].filter(function (l) { return l[2] || c.rows.some(function (r) { return r[l[1]]; }); });
      var cols = [{ key: "l", label: "", align: "left" }].concat(c.rows.map(function (r) {
        return { key: r.month, label: r.label + (r.bas ? " BAS" : ""), fmt: F.k, cell: U.moneyCell };
      }));
      var flow = U.table(cols, LINES.map(function (l) {
        var row = { l: l[0], _cls: l[2] ? "totalrow" : "" };
        c.rows.forEach(function (r) { row[r.month] = r[l[1]]; });
        return row;
      }), { dense: true });

      var last = c.rows[c.rows.length - 1];
      function tot(k) { return c.rows.reduce(function (a, r) { return a + r[k]; }, 0); }
      var bridgeRows = [
        { l: "Net profit, forecast", v: tot("netProfit") },
        { l: "Add back what is not cash: depreciation, provisions, write offs", v: tot("noncash") },
        { l: "Receivables: " + (last.ar > c.openingAR ? "more" : "less") + " tied up at " + S.debtorDays + " days to pay", v: c.openingAR - last.ar },
        { l: "Payables: " + (last.ap > c.openingAP ? "more" : "less") + " owed at " + S.creditorDays + " days", v: last.ap - c.openingAP },
        { l: "GST collected less GST paid to the ATO", v: last.gstOwed - c.gstOpening },
        { l: "Super accrued less super paid", v: last.superOwed - c.superOpening },
        { l: "PAYG instalments", v: tot("payg") },
        { l: "Commitment timing", v: tot("timing") },
        { l: "Loans, capital, tax and distributions", v: tot("outside") },
        { l: "Credit cards cleared", v: tot("cards") },
      ].filter(function (r) { return r.v !== 0 || /profit|Receivables|Payables|GST/.test(r.l); });
      var bridgeSum = bridgeRows.reduce(function (a, r) { return a + r.v; }, 0);

      var balTable = U.table([
        { key: "account", label: "Account", align: "left" }, { key: "balance", label: "Balance", fmt: function (v) { return F.money(Math.round(v * 100)); } },
        { key: "note", label: "Note", align: "left" },
      ], c.balances, { dense: true });
      var cardTable = c.cards.length ? U.table([
        { key: "card", label: "Card", align: "left" }, { key: "holder", label: "Holder", align: "left" },
        { key: "balance", label: "Owing", fmt: function (v) { return F.money(Math.round(v * 100)); } },
        { key: "limit", label: "Limit", fmt: function (v) { return F.money(Math.round(v * 100)); } },
        { key: "room", label: "Headroom", value: function (r) { return F.money(Math.round(((+r.limit || 0) - (+r.balance || 0)) * 100)); } },
        { key: "paymentDay", label: "Paid on", align: "left", value: function (r) { return r.paymentDay ? "day " + r.paymentDay : ""; } },
        { key: "note", label: "Note", align: "left" },
      ], c.cards, { dense: true }) : U.note("No credit cards typed. Add each card on the Credit cards tab with its month end balance; Xero does not carry them.", "muted");

      var comTable = c.commitments.length ? U.table([
        { key: "name", label: "Commitment", align: "left" }, { key: "category", label: "Kind", align: "left" },
        { key: "cents", label: "Amount", fmt: F.money }, { key: "freq", label: "How often", align: "left" },
        { key: "next", label: "Next due", align: "left", fmt: function (v) { return v ? F.date(v) : ""; } },
        { key: "inPnl", label: "In the P&L", align: "left", value: function (r) { return r.inPnl ? h("span.chip.chip-muted", "yes, timing only") : h("span.chip.chip-warning", "no, whole payment"); } },
        { key: "dates", label: "Falls in the year", value: function (r) { return String(r.dates.length); } },
        { key: "note", label: "Note", align: "left" },
      ], c.commitments, { dense: true }) : U.note("No commitments typed. Rent, loans, leases, insurance and subscriptions go on the Commitments tab.", "muted");

      var history = (E.cashData.balances || []).reduce(function (m, r) { if (r.month) m[r.month] = (m[r.month] || 0) + Math.round((+r.balance || 0) * 100); return m; }, {});
      var histKeys = Object.keys(history).sort();
      var histTable = histKeys.length > 1 ? U.table([
        { key: "m", label: "Month end", align: "left" }, { key: "v", label: "Bank balances", fmt: F.money },
        { key: "d", label: "Movement", fmt: F.money, cell: U.moneyCell },
      ], histKeys.map(function (k, i) { return { m: (E.monthIdx[k] || {}).label || k, v: history[k], d: i ? history[k] - history[histKeys[i - 1]] : null }; }), { dense: true }) : null;

      return [
        seedBanner(), CTS.periodBar(), U.h1("Cash Flow", "From " + rmLabel + ", twelve months, indirect method"), viewNote,
        noOpening ? h("div.banner.banner-warn", [h("strong", "Opening cash is zero because no bank balance is typed for " + rmLabel + ". "), "Type the month end balances on the Bank balances tab and rebuild; every closing figure below moves by the same amount."]) : null,
        U.note("Starts from the bank balances at " + rmLabel + " and walks forward on the P&L forecast: revenue comes in at " + S.debtorDays + " days (" + S.debtorSource + "), costs go out at " + S.creditorDays + " days (" + S.creditorSource + "), wages in the month, GST " + (S.basFrequency === "monthly" ? "monthly" : "on the quarterly BAS") + ", super " + (S.superTiming === "quarterly" ? "quarterly" : "with each pay") + ". Commitments already in the P&L only move timing; loans, capital, tax and distributions come off in full. Settings are on 10 Cash and Commitments."),
        tiles,
        U.section("Closing cash by month", U.figure("Cash at each month end, " + E.monthIdx[c.months[0]].label + " to " + E.monthIdx[c.months[c.months.length - 1]].label, chart, flow,
          "BAS marks a month the quarterly statement is paid in. Receipts and payments include GST; the GST line is the net handed to the ATO.")),
        U.section("From profit to cash", U.table([
          { key: "l", label: "", align: "left" }, { key: "v", label: "Twelve months", fmt: F.money, cell: U.moneyCell },
        ], bridgeRows.concat([{ l: "Net cash movement", v: c.closing - c.opening, _cls: "totalrow" }])),
          "The indirect method. Profit is the forecast; everything under it is why cash is not profit. The lines add to the movement" + (Math.abs(bridgeSum - (c.closing - c.opening)) > 100 ? ", except that here they do not, by " + F.money(bridgeSum - (c.closing - c.opening)) + ", which is a bug to report" : " exactly") + "."),
        U.section("Bank balances at " + ((E.monthIdx[c.openingMonth] || {}).label || rmLabel), [balTable, histTable],
          "Typed at each month end on the Bank balances tab. The history builds a record of actual cash to check the forecast against."),
        U.section("Credit cards", cardTable,
          "Xero does not carry the cards, so their balances live here. " + (S.cardsPaidInFull ? "The balance at the reporting month is cleared in the first forecast month; spend after that is assumed paid within the month." : "Balances are carried, not cleared; set Credit cards paid in full to Yes on the Settings tab to clear them.")),
        U.section("Commitments", comTable,
          "Rent, loans, leases, insurance, subscriptions. Whether a commitment is already in the P&L decides how it is treated: yes means only the difference between the monthly accrual and the payment dates moves cash; no means the whole payment comes off."),
        U.section("Opening balances the walk starts from", U.table([
          { key: "l", label: "", align: "left" }, { key: "v", label: "", fmt: F.money }, { key: "s", label: "Source", align: "left" },
        ], [
          { l: "Receivables owed to CTS", v: c.openingAR, s: c.arModelled ? "modelled from revenue at " + S.debtorDays + " days; type the Xero figure on Settings" : "typed on Settings, which puts days to pay at " + S.debtorDays },
          { l: "Payables CTS owes", v: c.openingAP, s: c.apModelled ? "modelled from costs at " + S.creditorDays + " days; type the Xero figure on Settings" : "typed on Settings, which puts days to pay at " + S.creditorDays },
          { l: "GST built up this quarter", v: c.gstOpening, s: "from the P&L months of the current quarter" },
          { l: "Facility limit", v: S.facility, s: S.facility ? "typed on Settings" : "none" },
        ], { dense: true })),
      ];
    },
  };

  /* ==================================================== commentary ===== */
  P.commentary = {
    section: "Dashboard", title: "Commentary",
    sub: "What was said about the month, and publishing it.",
    render: function () {
      var month = E.reportingMonth(), label = (E.monthIdx[month] || {}).long || month;
      var items = E.commentaryFor(null, month);
      var byPage = {};
      items.forEach(function (c) { (byPage[c.page] = byPage[c.page] || []).push(c); });
      var pages = Object.keys(byPage).sort(function (a, b) { return Object.keys(CTS.pages).indexOf(a) - Object.keys(CTS.pages).indexOf(b); });
      var pending = E.commentaryDraftCount();
      var status = h("div");
      async function publish() {
        status.innerHTML = "";
        var B = CTS.build;
        if (!B || !window.showDirectoryPicker) { status.appendChild(h("div.banner.banner-warn", "Publishing writes into the portal folder, which needs Edge or Chrome on the synced folder.")); return; }
        try {
          var dir = await CTS.reuseDir();
          if (!dir) { var pk = await B.pickFolder(); if (!pk.ok) throw new Error("that is not the portal folder"); dir = pk.handle; await CTS.saveDir(dir); }
          var data = E.commentaryExport();
          await B.writeCommentary(dir, data);
          E.clearCommentaryDrafts();
          status.appendChild(h("div.banner.banner-good", [h("strong", data.items.length + " comments published. "), "OneDrive carries the file to everyone; reloading…"]));
          setTimeout(function () { location.reload(); }, 1200);
        } catch (e) { status.appendChild(h("div.banner.banner-warn", "Could not publish: " + (e && e.message || e))); }
      }
      var list = pages.length ? pages.map(function (pg) {
        return U.section((CTS.pages[pg] || {}).title || pg, byPage[pg].map(function (c) {
          return h("div.commentary-item", [
            h("div.commentary-meta", [h("strong", c.author), h("span", F.date(String(c.at).slice(0, 10))),
              c.dept ? h("span.chip.chip-muted", (E.deptOf[c.dept] || {}).short || c.dept) : null,
              c.draft ? h("span.chip.chip-warning", "draft") : h("span.chip.chip-good", "published"),
              h("a.small", { href: "#/" + pg }, "open the page")]),
            h("div.commentary-text-view", c.text),
          ]);
        }));
      }) : [U.note("No commentary for " + label + " yet. Every page has an Add commentary box under its heading.")];
      return [
        seedBanner(), CTS.periodBar(), U.h1("Commentary", label),
        U.note("Commentary lives beside the numbers it explains: each page has its own block under the heading. What is typed there is a draft in that browser until finance publishes it, which writes one data file into the portal folder that OneDrive carries to everyone. Published commentary also goes into the monthly emails."),
        E.canPublishCommentary() ? h("div.toolbar", [
          h("button.btn", { type: "button", onclick: publish, disabled: !pending }, pending ? "Publish " + pending + " change" + (pending > 1 ? "s" : "") + " to the portal folder" : "Nothing to publish"),
          h("button.btn.btn-quiet", { type: "button", title: "A first draft for every page that has none yet, written from the numbers", onclick: function () {
            var made = 0;
            ["home", "pnl", "pnl-dept", "rev-summary", "util", "forecast", "cash", "clients"].forEach(function (pg) {
              if (!CTS.pages[pg] || E.commentaryFor(pg, month).length) return;
              if (E.addCommentary({ page: pg, month: month, text: E.proposeCommentary(pg) })) made++;
            });
            CTS.router.reload();
          } }, "Propose commentary for every page"),
          h("span.muted", "Proposals are drafts written from the data; edit each on its page, then publish."),
        ]) : (E.canComment() ? U.note("Your drafts are listed here. Finance publishes them.", "muted") : null),
        status,
      ].concat(list);
    },
  };

  /* ======================================================= the month ===== */
  P.story = {
    section: "Dashboard", title: "The Month",
    sub: "Five things that moved, written from the numbers.",
    points: function () { return P.story._points(); },
    render: function () {
      var p = CTS.period(), rm = E.reportingMonth(), mLabel = (E.monthIdx[rm] || {}).long || rm;
      var points = P.story._points();
      var cards = points.map(function (pt, i) {
        return h("a.storycard.tone-" + pt.tone, { href: pt.href }, [
          h("div.story-n", String(i + 1)), h("div", [h("h3", pt.head), h("p", pt.body)]),
        ]);
      });
      return [
        seedBanner(), CTS.periodBar(), U.h1("The Month", mLabel + " in " + points.length + " points"),
        U.note("Written from the numbers each time the page opens, so it is always the current build. The commentary block above it is where finance says what the numbers do not. Each point opens the page behind it."),
        h("div.story", cards),
      ];
    },
    _points: function () {
      var p = CTS.period(), rm = E.reportingMonth(), mLabel = (E.monthIdx[rm] || {}).long || rm;
      var keys = [rm], act = E.pnlActual(keys).totals, bud = E.pnlBudget(keys).totals, prior = E.pnlActual(E.priorYearMonths(keys)).totals;
      var ytdA = E.pnlActual(E.ytd()).totals, ytdB = E.pnlBudget(E.ytd()).totals;
      var alloc = E.allocate(keys), util = E.utilFor(keys);
      var points = [];
      function money(c) { return F.dollars(Math.abs(c)); }
      function sign(c) { return c >= 0 ? "ahead of" : "behind"; }
      // 1 revenue
      points.push({ head: "Revenue " + money(act.income) + ", " + sign(act.income - bud.income) + " budget by " + money(act.income - bud.income),
                    body: "Against " + money(bud.income) + " budgeted and " + money(prior.income) + " the same month last year. Year to date " + F.dollars(ytdA.income) + " against " + F.dollars(ytdB.income) + ".",
                    tone: act.income >= bud.income ? "good" : "critical", href: "#/rev-summary" });
      // 2 margin
      var gm = act.income ? act.grossProfit / act.income : null, bgm = bud.income ? bud.grossProfit / bud.income : null;
      points.push({ head: "Gross margin " + F.pct(gm) + (bgm != null ? (gm >= bgm ? ", above" : ", below") + " the " + F.pct(bgm) + " budgeted" : ""),
                    body: "Gross profit " + money(act.grossProfit) + " on cost of sales " + money(act.cos) + ". " + (gm != null && bgm != null && gm < bgm ? "Every point of margin on this month's revenue is " + F.dollars(Math.round(act.income / 100)) + "." : "Holding margin at this volume is what turns revenue into profit."),
                    tone: gm != null && bgm != null && gm >= bgm ? "good" : "warning", href: "#/pnl" });
      // 3 biggest department story
      var deptRows = alloc.rows.filter(function (r) { return r.dept.isRevenue && r.dept.code !== "ADMIN"; });
      var budDept = E.budgetByDept(keys, "income");
      var moves = deptRows.map(function (r) { return { r: r, v: r.income - (budDept[r.code] || 0) }; }).sort(function (a, b) { return Math.abs(b.v) - Math.abs(a.v); });
      if (moves.length) {
        var mv = moves[0];
        points.push({ head: mv.r.dept.short + " " + (mv.v >= 0 ? "carried" : "missed") + " the month by " + money(mv.v),
                      body: "Revenue " + money(mv.r.income) + " against " + money(budDept[mv.r.code] || 0) + " budgeted, gross margin " + F.pct(mv.r.gmPct) + (mv.r.dept.gmNorm != null ? " against a norm of " + F.pct(mv.r.dept.gmNorm, 0) : "") + ". Net profit after the overhead split " + F.money(mv.r.netProfit) + ".",
                      tone: mv.v >= 0 ? "good" : "critical", href: "#/pnl-dept" });
      }
      // 4 utilisation
      var low = (util.rows || []).filter(function (r) { return r.target != null && r.util != null && r.util < r.target - 0.05; }).sort(function (a, b) { return (a.util - a.target) - (b.util - b.target); })[0];
      points.push({ head: "Utilisation " + F.pct(util.total.util) + " across " + (util.total.fte ? util.total.fte.toFixed(1) + " full time equivalents" : "the business"),
                    body: low ? low.dept.short + " is the one to look at, at " + F.pct(low.util) + " against a target of " + F.pct(low.target, 0) + ". Read a small gap as leave before flagging it; a sustained gap is the one worth raising." : "No department is more than five points under its target this month.",
                    tone: low ? "warning" : "good", href: "#/util" });
      // 5 cash or profit
      var c = E.cashEnabled() && E.forecastEnabled() ? E.cashflow() : null;
      if (c) points.push({ head: "Cash " + F.dollars(c.opening) + ", low point " + F.dollars(c.low.closing) + " in " + c.low.label,
                           body: "Twelve months ahead on the forecast, " + F.dollars(c.closing) + " at " + E.monthIdx[c.months[c.months.length - 1]].label + ". Credit cards owing " + F.dollars(c.cardOwing) + ".", tone: c.low.closing < 0 ? "critical" : "good", href: "#/cash" });
      else points.push({ head: "Net profit " + F.money(act.netProfit) + " for the month", body: "Against " + F.money(bud.netProfit) + " budgeted. Year to date " + F.money(ytdA.netProfit) + " against " + F.money(ytdB.netProfit) + ".", tone: act.netProfit >= bud.netProfit ? "good" : "critical", href: "#/pnl" });
      if (E.forecastEnabled()) {
        var fy = E.monthsOfFY(p.fy), land = E.pnlBlend(fy).totals, fb = E.pnlBudget(fy).totals;
        points.push({ head: "The year is landing at " + F.money(land.netProfit) + " net profit, " + (land.netProfit >= fb.netProfit ? "ahead of" : "behind") + " budget by " + money(land.netProfit - fb.netProfit),
                      body: "Actual to " + mLabel + " then the forecast, line by line. Revenue " + F.dollars(land.income) + " against " + F.dollars(fb.income) + " budgeted.", tone: land.netProfit >= fb.netProfit ? "good" : "warning", href: "#/forecast" });
      }
      return points;
    },
  };

  /* ==================================================== revenue ======== */
  P["rev-summary"] = {
    section: "Revenue", title: "Revenue Summary",
    sub: "By department, by month, against last year.",
    render: function () {
      var p = CTS.period();
      var fyKeys = E.monthsOfFY(p.fy);
      var labels = fyKeys.map(function (k) { return E.monthIdx[k].label; });
      var depts = E.visibleDepts().filter(function (d) { return d.isRevenue; });
      var series = depts.map(function (d) {
        return { label: d.short, colour: U.colourOf(d.code), key: d.code,
                 values: fyKeys.map(function (k) {
                   var src = E.blendSource(k);
                   if (src === "actual") return E.sumMonths((E.idx.deptCatMonth[d.code] || {}).income, [k]);
                   if (src === "forecast") return E.forecastByDept([k], "income")[d.code] || 0;
                   return E.budgetByDept([k], "income")[d.code] || 0;
                 }) };
      });
      var stacked = U.columns({ labels: labels, series: series, stacked: true,
                                width: 780, height: 280 });

      var prior = E.priorYearMonths(p.keys);
      var rows = depts.map(function (d) {
        var now = E.sumMonths((E.idx.deptCatMonth[d.code] || {}).income, p.keys);
        var was = E.sumMonths((E.idx.deptCatMonth[d.code] || {}).income, prior);
        var bud = E.budgetByDept(p.keys, "income")[d.code] || 0;
        return { d: d.short, code: d.code, now: now, was: was, delta: now - was,
                 bud: bud, bvar: now - bud, risk: E.risk(now, bud) };
      });

      return [
        seedBanner(), CTS.periodBar(), U.h1("Revenue Summary", p.label),
        U.section("Revenue by department across FY" + p.fy,
          U.figure("Stacked by department, " + E.blendLabel(), stacked,
            U.table([{ key: "m", label: "Month", align: "left" }].concat(
              depts.map(function (d) { return { key: d.code, label: d.short, fmt: F.k }; })),
              fyKeys.map(function (k, i) {
                var row = { m: E.monthIdx[k].label,
                            _cls: E.blendSource(k) === "actual" ? "" : "forecast" };
                series.forEach(function (s) { row[s.key] = s.values[i]; });
                return row;
              }), { dense: true }))),
        U.section("This period against last year and against budget", U.table([
          { key: "d", label: "Department", align: "left",
            value: function (r) { return h("span", [h("span.dot", { style: { background: U.colourOf(r.code) } }), r.d]); } },
          { key: "now", label: "This period", fmt: F.money },
          { key: "was", label: "Same period last year", fmt: F.money },
          { key: "delta", label: "Variance", fmt: F.money, cell: U.moneyCell },
          { key: "bud", label: "Budget", fmt: F.money },
          { key: "bvar", label: "vs budget", fmt: F.money, cell: U.moneyCell },
          { key: "risk", label: "Risk", align: "left",
            value: function (r) { return h("span.risk.risk-" + U.RISK_CLASS[r.risk], r.risk); } },
        ], rows),
          "Departmental budget is derived: the budget is held per account, and most accounts carry no department in their name, so each account's budget is spread on that account's own prior year ledger split."),
      ];
    },
  };

  /* ============================================= revenue forecast ====== */
  P["rev-forecast"] = {
    section: "Revenue", title: "Revenue Forecast",
    sub: "Booked, weighted, baseline, typed.",
    render: function () {
      var p = CTS.period();
      var L = E.forecastLayers();
      if (!L) {
        return [seedBanner(), CTS.periodBar(), U.h1("Revenue Forecast", "FY" + p.fy),
          h("div.banner.banner-warn", [h("strong", "No forecast is loaded. "), "Put 09 Forecast.xlsx in the templates folder and run Build."])];
      }
      var PS = E.pipelineSettings();
      var depts = E.visibleDepts().filter(function (d) { return d.isRevenue && d.code !== "ADMIN"; });
      var codes = depts.map(function (d) { return d.code; });
      var pick = CTS.store.get("revFcDept", "ALL");
      if (pick !== "ALL" && codes.indexOf(pick) < 0) pick = "ALL";
      var use = pick === "ALL" ? codes : [pick];
      var months = L.months.filter(function (m) { return E.hasForecast(m); });
      var labels = months.map(function (m) { return E.monthIdx[m].label; });
      function layer(m, k) { var t = 0; use.forEach(function (d) { var x = (L.layers[d] || {})[m]; if (x) t += x[k] || 0; }); return t; }
      function forecastIncome(m) { return use.reduce(function (a, d) { return a + (E.forecastByDept([m], "income")[d] || 0); }, 0); }
      var rows = months.map(function (m, i) {
        var conf = layer(m, "confirmed"), wgt = layer(m, "weighted"), nev = layer(m, "never"), floor = layer(m, "floor"), target = layer(m, "target");
        var fc = forecastIncome(m), typed = fc - target;
        var bud = use.reduce(function (a, d) { return a + (E.budgetByDept([m], "income")[d] || 0); }, 0);
        var pk = (+m.slice(0, 4) - 1) + m.slice(4);
        var ly = use.reduce(function (a, d) { return a + E.sumMonths((E.idx.deptCatMonth[d] || {}).income, [pk]); }, 0);
        var rule = pick === "ALL" ? (i < PS.near ? "near" : "far") : ((L.layers[pick] || {})[m] || {}).rule || "";
        return { m: E.monthIdx[m].label, key: m, rule: rule, conf: conf, wgt: wgt, nev: nev, floor: floor, typed: typed, fc: fc, bud: bud, ly: ly, v: fc - bud };
      });
      var chart = U.columns({
        labels: labels, width: 780, height: 280, stacked: true,
        series: [
          { label: "Confirmed", colour: "var(--series-1)", values: rows.map(function (r) { return r.conf; }) },
          { label: "Weighted deals", colour: "var(--series-2)", values: rows.map(function (r) { return r.wgt; }) },
          { label: "Baseline share", colour: "var(--series-3)", values: rows.map(function (r) { return r.nev; }) },
          { label: "Baseline floor", colour: "var(--series-4)", values: rows.map(function (r) { return r.floor; }) },
          { label: "Typed", colour: "var(--series-5)", values: rows.map(function (r) { return r.typed; }) },
        ],
      });
      var table = U.table([
        { key: "m", label: "Month", align: "left" }, { key: "rule", label: "Rule", align: "left" },
        { key: "conf", label: "Confirmed", fmt: F.k }, { key: "wgt", label: "Weighted deals", fmt: F.k },
        { key: "nev", label: "Baseline share", fmt: F.k }, { key: "floor", label: "Baseline floor", fmt: F.k },
        { key: "typed", label: "Typed", fmt: F.k, cell: U.moneyCell },
        { key: "fc", label: "Forecast", fmt: F.k }, { key: "bud", label: "Budget", fmt: F.k },
        { key: "v", label: "vs budget", fmt: F.k, cell: U.moneyCell }, { key: "ly", label: "Last year", fmt: F.k },
      ], rows.concat([(function () { var t = { m: "Total", rule: "", _cls: "totalrow" }; ["conf", "wgt", "nev", "floor", "typed", "fc", "bud", "v", "ly"].forEach(function (k) { t[k] = rows.reduce(function (a, r) { return a + r[k]; }, 0); }); return t; })()]), { dense: true });

      var items = [], deals = [];
      use.forEach(function (d) { months.forEach(function (m) { var x = (L.layers[d] || {})[m]; if (!x) return; (x.items || []).forEach(function (it) { items.push(Object.assign({ dept: d }, it)); }); (x.deals || []).forEach(function (it) { deals.push(Object.assign({ dept: d }, it)); }); }); });
      items.sort(function (a, b) { return b.value - a.value; }); deals.sort(function (a, b) { return b.weighted - a.weighted; });
      var itemTable = items.length ? U.table([
        { key: "month", label: "Lands", align: "left", value: function (r) { return (E.monthIdx[r.month] || {}).label || r.month; } },
        { key: "dept", label: "Dept", align: "left", value: function (r) { return (E.deptOf[r.dept] || {}).short || r.dept; } },
        { key: "client", label: "Client", align: "left" }, { key: "title", label: "What", align: "left" },
        { key: "kind", label: "From", align: "left", value: function (r) { return h("span.chip.chip-muted", r.kind === "order" ? "OnRent order" : r.kind === "quote" ? "Qwilr quote" : "Zoho won"); } },
        { key: "status", label: "Status", align: "left" }, { key: "value", label: "Value", fmt: F.money },
      ], items.slice(0, 20), { dense: true }) : U.note("Nothing confirmed lands in these months.", "muted");
      var dealTable = deals.length ? U.table([
        { key: "month", label: "Lands", align: "left", value: function (r) { return (E.monthIdx[r.month] || {}).label || r.month; } },
        { key: "dept", label: "Dept", align: "left", value: function (r) { return (E.deptOf[r.dept] || {}).short || r.dept; } },
        { key: "client", label: "Account", align: "left" }, { key: "title", label: "Deal", align: "left" }, { key: "stage", label: "Stage", align: "left" },
        { key: "amount", label: "Amount", fmt: F.money }, { key: "prob", label: "Probability", fmt: function (v) { return F.pct(v, 0); } },
        { key: "weighted", label: "Weighted", fmt: F.money },
      ], deals.slice(0, 20), { dense: true }) : U.note("No open deals land in these months.", "muted");

      var pl = L.pipeline;
      var wonNote = pl && pl.wonUncounted.length ? h("div.banner.banner-warn", [
        h("strong", pl.wonUncounted.length + " won deal" + (pl.wonUncounted.length > 1 ? "s" : "") + " not counted, " + F.dollars(pl.wonUncounted.reduce(function (a, w) { return a + w.value; }, 0)) + ". "),
        "A won deal in Zoho counts only once it has an OnRent order or an accepted Qwilr quote, which is where the date and the final value live. Set Count won deals to Yes on the Pipeline tab to count them anyway.",
      ]) : null;
      var never = codes.map(function (c) { return E.deptOf[c].short + " " + F.pct(PS.never[c] != null ? +PS.never[c] : +PS.never["default"], 0); }).join(", ");
      var meta = (E.pipeline && E.pipeline.meta) || {};

      return [
        seedBanner(), CTS.periodBar(), U.h1("Revenue Forecast", "FY" + p.fy + ", " + (pick === "ALL" ? "all departments" : E.deptOf[pick].short)),
        h("div.toolbar", [h("div.seg", [{ code: "ALL", short: "All" }].concat(depts).map(function (d) {
          return h("button.seg-btn" + (d.code === pick ? ".on" : ""), { onclick: function () { CTS.store.set("revFcDept", d.code); CTS.router.reload(); } }, d.short);
        }))]),
        !PS.enabled ? h("div.banner.banner-warn", [h("strong", E.hasPipeline ? "The pipeline layer is off. " : "No pipeline is loaded. "),
          E.hasPipeline ? "Income is on the baseline method alone. Set Use the pipeline to Yes on the Pipeline tab of 09 Forecast.xlsx." : "Fill the Zoho, OnRent and Qwilr templates and run Build; until then income is the baseline alone."]) : null,
        meta.placeholder && PS.enabled ? h("div.banner.banner-warn", [h("strong", "Placeholder pipeline. "), "The deals, orders and quotes come from the placeholder templates. Once a real export from each system is matched, this page reads the real book."]) : null,
        U.note("Income after " + ((E.monthIdx[p.rm] || {}).label || p.rm) + " is built in layers. Confirmed is OnRent orders and accepted Qwilr quotes by the month they land. Weighted deals is the open Zoho pipeline at each deal's probability. Baseline share is the part of revenue that never goes through a system, " + never + ", as a share of the same month last year. The first " + PS.near + " months are what is booked plus that share; after that the forecast never falls below the baseline. A typed override sits on top."),
        U.section("By month", U.figure("Layers of the income forecast, FY" + p.fy, chart, table,
          "Rule near means booked plus the baseline share; far means the larger of that and the baseline. Typed is the difference an override made.")),
        wonNote,
        U.section("Confirmed work landing in the forecast", itemTable, "Largest first. Value as the system holds it."),
        U.section("Open deals landing in the forecast", dealTable, "Probability is Zoho's unless overridden per stage on the Pipeline tab."),
        pl ? U.section("What was read", U.table([
          { key: "l", label: "", align: "left" }, { key: "v", label: "" },
        ], [
          { l: "Orders counted as confirmed", v: pl.counted.orders }, { l: "Quotes counted as confirmed", v: pl.counted.quotes },
          { l: "Quotes skipped because an order already carries them", v: pl.counted.dupQuotes },
          { l: "Open deals weighted", v: pl.counted.deals }, { l: "Won deals not counted", v: pl.wonUncounted.length },
        ], { dense: true }), "Order statuses counted: " + PS.orderStatuses + ". Quote statuses: " + PS.quoteStatuses + ". Cancellation rate " + F.pct(PS.cancel, 0) + ".") : null,
      ];
    },
  };

  P["rev-schedule"] = {
    section: "Revenue", title: "Revenue Schedule",
    sub: "Per role for onsite, per client elsewhere.",
    render: function () {
      var p = CTS.period();
      var allowed = E.visibleDepts().filter(function (d) {
        return d.isRevenue && d.code !== "ADMIN";
      }).map(function (d) { return d.code; });
      var dept = CTS.store.get("schedDept", "PRODUCTION");
      if (allowed.indexOf(dept) < 0) dept = allowed[0] || "PRODUCTION";
      var s = E.schedule(dept, p.keys);
      var d = E.deptOf[dept];

      return [
        seedBanner(), CTS.periodBar(), U.h1("Revenue Schedule", d.short + ", " + p.label),
        h("div.toolbar", [
          h("div.seg", E.visibleDepts().filter(function (d) {
            return d.isRevenue && d.code !== "ADMIN";
          }).map(function (d) { return d.code; }).map(function (c) {
            return h("button.seg-btn" + (c === dept ? ".on" : ""), {
              onclick: function () { CTS.store.set("schedDept", c); CTS.router.reload(); },
            }, E.deptOf[c].short);
          })),
        ]),
        U.note(dept === "ONSITE"
          ? "Onsite is budgeted per role, not per client. A line such as a team leader position is that person's revenue charged to the client. Where several roles are billed inside a single invoice, open that invoice in Xero and check whether there was a deduction against the role before you record it."
          : dept === "CONSULTING"
            ? "Consulting is usually a short list and you will already know from the revenue meeting which clients are in it. Where revenue did not come in, say whether the deal was lost or simply moved to a later month; that is the one change management asked for."
            : "Production is the hardest and takes the longest, because production and video are separate. Anything with revenue that is not already listed goes into the new business line rather than getting a row of its own."),
        U.section("Schedule", U.table([
          { key: "line", label: "Line", align: "left", width: "22rem" },
          { key: "kind", label: "", align: "left",
            value: function (r) { return h("span.chip.chip-muted", r.kind === "newbiz" ? "new business" : r.kind); } },
          { key: "budget", label: "Budget", fmt: F.money },
          { key: "actual", label: "Actual", fmt: F.money },
          { key: "variance", label: "Variance", fmt: F.money, cell: U.moneyCell },
          { key: "pct", label: "%", fmt: function (v) { return F.pct(v); } },
          { key: "risk", label: "Risk", align: "left",
            value: function (r) { return h("span.risk.risk-" + U.RISK_CLASS[r.risk], r.risk); } },
        ], s.lines.concat([{ line: "Total", kind: "", budget: s.budgetTotal,
                             actual: s.actualTotal, variance: s.variance,
                             pct: s.budgetTotal ? s.variance / Math.abs(s.budgetTotal) : null,
                             risk: E.risk(s.actualTotal, s.budgetTotal), _cls: "totalrow" }]))),
        U.note(s.note),
        U.note("Close is good enough. Work in progress revenue forms part of the P&L but is not invoiced to a client, so an exact tie back is not possible; being close is the standard here.", "muted"),
      ];
    },
  };

  P.clients = {
    section: "Revenue", title: "Top Clients",
    sub: "Against the same period last year.",
    render: function () {
      var p = CTS.period();
      var prior = E.priorYearMonths(p.keys);
      var rows = E.clientRows(p.keys, prior).filter(function (r) { return r.amount || r.prior; });
      var top = rows.slice(0, 10);
      var chart = U.bars({
        rows: top.map(function (r) {
          return { label: r.display, value: r.amount, value2: r.prior,
                   colour: U.colourOf(r.primaryDept) };
        }),
        label1: "This period", label2: "Same period last year", width: 560,
      });

      var head = E.cfg.COMMENTARY.topClientHeadings;
      var increased = rows.filter(function (r) { return r.prior && r.delta > 0; }).slice(0, 5);
      var decreased = rows.filter(function (r) { return r.prior && r.delta < 0; })
        .sort(function (a, b) { return a.delta - b.delta; }).slice(0, 5);
      var isNew = rows.filter(function (r) { return r.amount && !r.prior; });
      var dropped = rows.filter(function (r) { return !r.amount && r.prior; });

      function list(title, items) {
        return h("div.card", [
          h("h3", title),
          items.length ? h("ul.plain", items.map(function (r) {
            return h("li", [h("span.li-name", r.display),
                            h("span.li-val", F.dollars(r.delta || r.amount || -r.prior))]);
          })) : h("p.muted", "None this period."),
        ]);
      }

      return [
        seedBanner(), CTS.periodBar(), U.h1("Top Clients", p.label),
        U.section("Top ten", U.figure("This period against the same period last year", chart,
          U.table([
            { key: "rank", label: "#", align: "left" },
            { key: "display", label: "Client", align: "left" },
            { key: "amount", label: "This period", fmt: F.money },
            { key: "prior", label: "Last year", fmt: F.money },
            { key: "delta", label: "Variance", fmt: F.money, cell: U.moneyCell },
          ], top, { dense: true }),
          "One client has to be one name or the chart shows it several times and understates it in each. The display name is matched by hand, which is why it is a column here and not a lookup.")),
        U.section("The four headings management want every time",
          h("div.cards", [list(head[0], increased), list(head[1], decreased),
                          list(head[2], isNew), list(head[3], dropped)])),
        U.section("Every client", U.table([
          { key: "rank", label: "#", align: "left" },
          { key: "display", label: "Client", align: "left" },
          { key: "industry", label: "Industry", align: "left" },
          { key: "engagement", label: "Engagement", align: "left" },
          { key: "amount", label: "This period", fmt: F.money },
          { key: "prior", label: "Last year", fmt: F.money },
          { key: "delta", label: "Variance", fmt: F.money, cell: U.moneyCell },
          { key: "pct", label: "%", fmt: function (v) { return F.pct(v); } },
        ], rows, { dense: true }),
          "Industry and engagement type are reference data held against the client, not something the ledger knows. In the live process that mapping lives in the previous year's June file and is partly out of date."),
      ];
    },
  };

  P["client-dept"] = {
    section: "Revenue", title: "Clients by Department",
    sub: "Who each department's revenue actually comes from.",
    render: function () {
      var p = CTS.period();
      var rows = E.clientRows(p.keys);
      var depts = E.visibleDepts().filter(function (d) { return d.isRevenue; });
      var cols = [{ key: "display", label: "Client", align: "left" }]
        .concat(depts.map(function (d) {
          return { key: d.code, label: d.short, fmt: F.money,
                   value: function (r) { return r.byDept[d.code] || 0; } };
        }))
        .concat([{ key: "amount", label: "Total", fmt: F.money, cls: "totalcol" }]);
      var top = rows.slice(0, 25);

      var best = depts.map(function (d) {
        var bestRow = null;
        rows.forEach(function (r) {
          var v = r.byDept[d.code] || 0;
          if (!bestRow || v > bestRow.v) bestRow = { name: r.display, v: v };
        });
        return { d: d.short, code: d.code, name: bestRow ? bestRow.name : "-",
                 v: bestRow ? bestRow.v : 0 };
      });

      return [
        seedBanner(), CTS.periodBar(), U.h1("Clients by Department", p.label),
        U.section("Top client in each department", U.figure(
          "Largest client by department",
          U.bars({ rows: best.map(function (b) {
            return { label: b.d + " — " + b.name, value: b.v, colour: U.colourOf(b.code) };
          }), width: 560, labelWidth: 230 }),
          U.table([
            { key: "d", label: "Department", align: "left" },
            { key: "name", label: "Client", align: "left" },
            { key: "v", label: "Revenue", fmt: F.money },
          ], best, { dense: true }))),
        U.section("Revenue by client and department", U.table(cols, top, { dense: true }),
          "Repeat the top clients exercise per department, once for onsite, once for production, once for consulting: that is what the graphs file does by hand."),
      ];
    },
  };

  /* ================================================== utilisation ====== */
  P.util = {
    section: "Utilisation", title: "Utilisation",
    sub: "Chargeable over worked hours, leave excluded.",
    render: function () {
      var p = CTS.period();
      var u = E.utilFor(p.keys);
      var roll = E.utilRolling(p.rm);
      var fyKeys = E.monthsOfFY(p.fy);
      var labels = fyKeys.map(function (k) { return E.monthIdx[k].label; });
      var depts = E.visibleDepts(E.postingDepts);

      var chart = U.lines({
        labels: labels, width: 780, height: 280,
        target: E.deptOf.ONSITE.utilTarget, targetLabel: "target " + F.pct(E.deptOf.ONSITE.utilTarget, 0),
        yfmt: function (v) { return F.pct(v, 0); },
        tipfmt: function (v) { return F.pct(v); },
        series: depts.map(function (d) {
          return { label: d.short, colour: U.colourOf(d.code),
                   values: fyKeys.map(function (k) {
                     var r = (E.utilIdx[d.code] || {})[k];
                     if (!r) return null;
                     var w = r.chargeable + r.nonChargeable;
                     return w ? r.chargeable / w : null;
                   }) };
        }),
      });

      return [
        seedBanner(), CTS.periodBar(), U.h1("Utilisation", p.label),
        h("div.tiles", [
          U.tile({ label: "Utilisation, this period", value: F.pct(u.total.util),
                   sub: F.hours(u.total.chargeable) + " chargeable of " + F.hours(u.total.worked) + " worked",
                   note: "Leave and public holidays are excluded from both sides" }),
          U.tile({ label: "Full time equivalents", value: (u.total.fte || 0).toFixed(1),
                   sub: "worked hours over " + u.hoursPerHead + " available", note: "Pack method" }),
          U.tile({ label: "Rolling twelve months", value: F.pct(roll.total.util),
                   sub: roll.complete ? roll.months[0] + " to " + roll.months[roll.months.length - 1] : "incomplete window",
                   note: (roll.total.fteRolling || 0).toFixed(1) + " FTE, graphs file method" }),
          U.tile({ label: "Annual hours check", value: F.num(roll.annualHours),
                   tone: roll.annualCheck ? "good" : "critical",
                   sub: roll.annualCheck ? "lands on 2,080 or 2,088" : "does not land on 2,080 or 2,088",
                   note: roll.businessDays + " weekdays over the window" }),
        ]),
        U.note("Two different figures are in use and they are not meant to agree. This portal computes both. The pack nets public holidays off the month and counts worked hours only. The graphs file counts every logged hour, leave included, against a plain weekday year, which is why that denominator lands on 2,080 or 2,088. Check it rather than assuming it: the Days tab is also split by state."),
        U.section("Utilisation month by month",
          U.figure("By department across FY" + p.fy, chart,
            U.table([{ key: "m", label: "Month", align: "left" }].concat(
              depts.map(function (d) {
                return { key: d.code, label: d.short, fmt: function (v) { return F.pct(v, 0); } };
              })),
              fyKeys.map(function (k) {
                var row = { m: E.monthIdx[k].label };
                depts.forEach(function (d) {
                  var r = (E.utilIdx[d.code] || {})[k];
                  var w = r ? r.chargeable + r.nonChargeable : 0;
                  row[d.code] = w ? r.chargeable / w : null;
                });
                return row;
              }), { dense: true }))),
        U.section("This period by department", U.table([
          { key: "d", label: "Department", align: "left",
            value: function (r) { return h("span", [h("span.dot", { style: { background: U.colourOf(r.code) } }), r.dept.short]); } },
          { key: "chargeable", label: "Chargeable", fmt: F.hours },
          { key: "nonChargeable", label: "Non-chargeable", fmt: F.hours },
          { key: "leave", label: "Leave", fmt: F.hours },
          { key: "publicHoliday", label: "Public holiday", fmt: F.hours },
          { key: "util", label: "Utilisation", fmt: function (v) { return F.pct(v); } },
          { key: "target", label: "Target", fmt: function (v) { return v == null ? "-" : F.pct(v, 0); } },
          { key: "gap", label: "Gap", fmt: function (v) { return v == null ? "-" : F.pct(v); },
            cell: function (r) { return r.gap != null && r.gap < 0 ? "neg" : ""; } },
          { key: "fte", label: "FTE", fmt: function (v) { return v == null ? "-" : v.toFixed(2); } },
        ], u.rows),
          "A negative gap means the team is under its target. For a team whose hours are almost all chargeable a small negative usually just means leave that month; a sustained gap is the one worth raising."),
        U.section("Working days behind the figures", U.table([
          { key: "m", label: "Month", align: "left" },
          { key: "bd", label: "Weekdays" },
          { key: "wd", label: "Less public holidays (" + u.state + ")" },
          { key: "hrs", label: "Available hours per head" },
        ], fyKeys.map(function (k) {
          var m = E.monthIdx[k];
          return { m: m.label, bd: m.bd, wd: m.wd[u.state],
                   hrs: m.wd[u.state] * E.hoursPerDay() };
        }), { dense: true }),
          "Holidays flagged as proclaimed rather than fixed are amber on Config & Variables. Verify them against your state's official list before relying on these."),
      ];
    },
  };

  P["profit-fte"] = {
    section: "Utilisation", title: "Profitability per FTE",
    sub: "Departmental profit spread over a calculated headcount.",
    render: function () {
      var p = CTS.period();
      var alloc = E.allocate(p.keys);
      var u = E.utilFor(p.keys);
      var byCode = {};
      u.rows.forEach(function (r) { byCode[r.code] = r; });
      var vis = E.visibleDepts();
      var rows = alloc.rows.filter(function (r) {
        return r.dept.isRevenue && vis.some(function (d) { return d.code === r.code; });
      }).map(function (r) {
        var fte = (byCode[r.code] || {}).fte || null;
        return { d: r.dept.short, code: r.code, fte: fte, np: r.netProfit,
                 gp: r.grossProfit, rev: r.income,
                 perFte: fte ? Math.round(r.netProfit / fte) : null,
                 revPerFte: fte ? Math.round(r.income / fte) : null };
      });
      return [
        seedBanner(), CTS.periodBar(), U.h1("Profitability per FTE", p.label),
        h("div.banner.banner-warn", [
          h("strong", "Read this before you present it. "),
          "There is no true profitability per employee in this report. It is departmental profit spread across a calculated headcount. It is a recent idea and the person who built it has said openly that a better method would be welcome. Do not present it as more precise than it is.",
        ]),
        U.section("By department", U.figure("Net profit per full time equivalent",
          U.bars({ rows: rows.map(function (r) {
            return { label: r.d, value: r.perFte || 0, colour: U.colourOf(r.code) };
          }), width: 520, labelWidth: 130 }),
          U.table([
            { key: "d", label: "Department", align: "left" },
            { key: "rev", label: "Revenue", fmt: F.money },
            { key: "np", label: "Net profit", fmt: F.money, cell: U.moneyCell },
            { key: "fte", label: "FTE", fmt: function (v) { return v == null ? "-" : v.toFixed(2); } },
            { key: "revPerFte", label: "Revenue per FTE", fmt: F.money },
            { key: "perFte", label: "Profit per FTE", fmt: F.money, cell: U.moneyCell },
          ], rows),
          "When this figure moves it is almost always the profit that moved, not the headcount: the full time equivalent changes by a fraction of a decimal place month to month. If you see it move materially, something is wrong with the hours rather than with the business.")),
      ];
    },
  };

  /* ====================================================== ledger ======= */
  P.ledger = {
    section: "Ledger", title: "Detail Records",
    sub: "Every line behind every figure.",
    render: function () {
      var p = CTS.period();
      var q = CTS.store.get("ledgerQuery", "");
      var deptF = CTS.store.get("ledgerDept", "all");
      var rows = E.gl.filter(function (r) {
        if (p.keys.indexOf(r.month) < 0) return false;
        if (deptF !== "all" && r.dept !== deptF) return false;
        if (q) {
          var s = (r.account + " " + r.contact + " " + r.jobNo + " " + r.costCentre).toLowerCase();
          if (s.indexOf(q.toLowerCase()) < 0) return false;
        }
        return true;
      });
      var total = rows.reduce(function (a, r) { return a + r.amount; }, 0);
      var shown = rows.slice(0, 300);

      return [
        seedBanner(), CTS.periodBar(), U.h1("Detail Records", p.label),
        h("div.toolbar", [
          h("input.control.grow", { type: "search", value: q,
            placeholder: "Search account, contact, job number or cost centre",
            onchange: function (e) { CTS.store.set("ledgerQuery", e.target.value); CTS.router.reload(); } }),
          h("select.control", {
            onchange: function (e) { CTS.store.set("ledgerDept", e.target.value); CTS.router.reload(); },
          }, [h("option", { value: "all", selected: deptF === "all" }, "All departments")].concat(
            E.depts.map(function (d) {
              return h("option", { value: d.code, selected: d.code === deptF }, d.short);
            }))),
          h("span.chip.chip-muted", rows.length + " lines, " + F.dollars(total)),
        ]),
        U.table([
          { key: "date", label: "Date", align: "left", fmt: F.date },
          { key: "account", label: "Account", align: "left" },
          { key: "sub", label: "Subcategory", align: "left" },
          { key: "dept", label: "Dept", align: "left",
            value: function (r) {
              return h("span", [h("span.dot", { style: { background: U.colourOf(r.dept) } }),
                                (E.deptOf[r.dept] || {}).short || r.dept]); } },
          { key: "contact", label: "Contact", align: "left" },
          { key: "source", label: "Source", align: "left" },
          { key: "jobNo", label: "Job", align: "left" },
          { key: "amount", label: "Amount", fmt: F.money, cell: U.moneyCell },
        ], shown, { dense: true }),
        rows.length > shown.length
          ? U.note("Showing the first " + shown.length + " of " + rows.length + " lines. Narrow the period or the search to see the rest.")
          : null,
      ];
    },
  };
})();

/* ===================================================================== *
 * Pages: admin. Setup, data loaders, config, about.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, h = CTS.h, F = CTS.fmt, E = CTS.engine, U = CTS.ui;
  var P = CTS.pages, seedBanner = CTS.seedBanner;

  /* ---- delimited paste parsing ---------------------------------------- */
  /** Parse a pasted table, tab or comma separated.
   *
   *  Character by character, not line by line, because a real export has
   *  quoted fields with newlines inside them: the Notes column of the
   *  Employment Hero earnings report carries a shift note that wraps. Splitting
   *  on newlines first would tear one row into several and throw the rest of
   *  the file out of alignment.
   */
  function parseTable(text) {
    var src = String(text == null ? "" : text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    if (!src.trim()) return { head: [], rows: [] };

    // Decide the delimiter from the first line outside quotes.
    var delim = "\t", q = false;
    for (var i = 0; i < src.length; i++) {
      var ch = src[i];
      if (ch === '"') q = !q;
      else if (!q && ch === "\n") break;
      else if (!q && ch === ",") { delim = ","; }
      else if (!q && ch === "\t") { delim = "\t"; break; }
    }

    var rows = [], row = [], field = "", inQ = false;
    for (var j = 0; j < src.length; j++) {
      var c = src[j];
      if (inQ) {
        if (c === '"') {
          if (src[j + 1] === '"') { field += '"'; j++; }
          else inQ = false;
        } else field += c;                 // newlines inside quotes are kept
      } else if (c === '"') {
        inQ = true;
      } else if (c === delim) {
        row.push(field); field = "";
      } else if (c === "\n") {
        row.push(field); field = "";
        rows.push(row); row = [];
      } else field += c;
    }
    row.push(field);
    if (row.length > 1 || row[0] !== "") rows.push(row);

    // Drop leading rows that are not the header: an export often opens with a
    // title line such as "Earnings Details" in a single cell.
    while (rows.length > 1 && rows[0].filter(function (x) { return String(x).trim(); }).length < 2) {
      rows.shift();
    }
    if (!rows.length) return { head: [], rows: [] };
    var head = rows[0].map(function (x) { return String(x).trim(); });
    var body = rows.slice(1).filter(function (r) {
      return r.some(function (x) { return String(x).trim(); });
    });
    return { head: head, rows: body };
  }

  function num(s) {
    if (s == null) return 0;
    var t = String(s).replace(/[$,\s]/g, "").trim();
    if (!t || t === "-") return 0;
    var neg = /^\(.*\)$/.test(t);
    if (neg) t = t.slice(1, -1);
    var v = parseFloat(t);
    if (isNaN(v)) return 0;
    return neg ? -v : v;
  }

  function isoDate(s) {
    if (!s) return null;
    s = String(s).trim();
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[1] + "-" + m[2] + "-" + m[3];
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);   // Australian d/m/y
    if (m) {
      var y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
      return y + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
    }
    var d = new Date(s);
    return isNaN(d) ? null : d.toISOString().slice(0, 10);
  }

  function download(name, text) {
    try {
      var blob = new Blob([text], { type: "text/javascript" });
      var a = h("a", { href: URL.createObjectURL(blob), download: name });
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 0);
      return true;
    } catch (e) { return false; }
  }

  /** Three ways to hold on to a load, because only one of them works
   *  everywhere. Keeping it in the browser needs nothing and survives a
   *  reload. Showing the file works even where downloads are blocked, which
   *  they are in a sandboxed frame. Downloading is offered last because it is
   *  the one that silently does nothing when the frame forbids it. */
  function keepOrShow(opts) {
    var wrap = h("div.keeprow");
    var msg = h("div.keepmsg");
    var box = h("div.keepbox", { style: { display: "none" } });

    function show() {
      box.innerHTML = "";
      box.style.display = "block";
      var ta = h("textarea.paste", { rows: 8, readonly: true }, opts.text());
      box.appendChild(U.note("Select all of this and paste it into a file called " +
        opts.name + " in portal/data. That makes it what the portal loads every time, for everyone."));
      box.appendChild(ta);
      ta.focus();
      ta.select();
    }

    wrap.appendChild(h("div.row", [
      opts.keep ? h("button.btn", { onclick: function () {
        var bad = null;
        opts.keep.forEach(function (k) {
          var r = E.snapshot(k.kind, k.payload());
          if (!r.ok && !bad) bad = r.why;
        });
        msg.innerHTML = "";
        msg.appendChild(!bad
          ? h("div.banner.banner-good", "Kept in this browser. It will still be here after a reload, on this machine, for you. To give it to everyone, use Show the file.")
          : h("div.banner.banner-warn", "Could not keep it in the browser: " + bad +
              ". Use Show the file instead."));
      } }, "Keep in this browser") : null,
      h("button.btn.btn-quiet", { onclick: show }, "Show the file"),
      h("button.btn.btn-quiet", { onclick: function () {
        download(opts.name, opts.text());
        msg.innerHTML = "";
        msg.appendChild(h("p.note", "If no download appeared, this page is in a frame that blocks them. Use Show the file."));
      } }, "Download"),
    ]));
    wrap.appendChild(msg);
    wrap.appendChild(box);
    return wrap;
  }

  /* ---- the Employment Hero earnings loader ----------------------------
   *
   * Built for the Earnings Details report, whose columns are:
   *
   *   Employee Id | Employee External Id | Employee Name | Pay Category Id |
   *   Pay Category External Id | Pay Category Name | Units | Unit Type |
   *   Location Id | Location External Id | Location Name | Notes | Rate |
   *   Rate Type | Gross Earnings | Taxable Earnings | SG Super
   *
   * Three things fall out of it without being asked for:
   *
   *   department   the Location Name carries it in brackets, as in
   *                "PWC IT Support - Brisbane [ONS]", the same bracket tag
   *                convention the general ledger uses on job numbers.
   *   employment   the pay category says it: Casual Ordinary Hours against
   *                Permanent Ordinary Hours. So someone moving from casual to
   *                permanent shows up as their categories changing, and the
   *                portal reports that rather than having to be told.
   *   labour cost  Gross Earnings and SG Super, which the timesheet export
   *                never carried.
   *
   * Two things still have to be decided by CTS and are remembered once set:
   * whether a location is client work or internal, and what each pay category
   * counts as.
   */
  var EARN_COLS = [
    { key: "empId", label: "Employee Id", names: ["employeeid"] },
    { key: "empExt", label: "Employee External Id", names: ["employeeexternalid"] },
    { key: "empName", label: "Employee Name", names: ["employeename"] },
    { key: "catId", label: "Pay Category Id", names: ["paycategoryid"] },
    { key: "catExt", label: "Pay Category External Id", names: ["paycategoryexternalid"] },
    { key: "catName", label: "Pay Category Name", names: ["paycategoryname"], need: true },
    { key: "units", label: "Units", names: ["units"], need: true },
    { key: "unitType", label: "Unit Type", names: ["unittype"] },
    { key: "locId", label: "Location Id", names: ["locationid"] },
    { key: "locExt", label: "Location External Id", names: ["locationexternalid"] },
    { key: "locName", label: "Location Name", names: ["locationname"], need: true },
    { key: "notes", label: "Notes", names: ["notes"] },
    { key: "rate", label: "Rate", names: ["rate"] },
    { key: "rateType", label: "Rate Type", names: ["ratetype"] },
    { key: "gross", label: "Gross Earnings", names: ["grossearnings"] },
    { key: "taxable", label: "Taxable Earnings", names: ["taxableearnings"] },
    { key: "super", label: "SG Super", names: ["sgsuper"] },
  ];

  var CAT_KINDS = [
    { v: "work", l: "Worked time" },
    { v: "leave", l: "Leave" },
    { v: "publicHoliday", l: "Public holiday" },
    { v: "exclude", l: "Not hours (pay only)" },
  ];

  function norm(sv) { return String(sv || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }

  function money(v) {
    var t = String(v == null ? "" : v).replace(/[$,\s]/g, "").trim();
    if (!t || t === "-") return 0;
    var neg = /^\(.*\)$/.test(t);
    if (neg) t = t.slice(1, -1);
    var n = parseFloat(t);
    return isNaN(n) ? 0 : (neg ? -n : n);
  }

  /** What a pay category counts as, guessed from its name. Grounded in the
   *  categories the payroll manuals name: Permanent and Casual Ordinary Hours,
   *  the overtime multipliers, leave without pay, time in lieu taken, bonus
   *  leave, per diem, back pay and the termination payouts. */
  function guessCategory(name) {
    var n = String(name || "").toLowerCase();
    if (/leave without pay|lwop/.test(n)) return "exclude";
    if (/public holiday/.test(n)) return "publicHoliday";
    if (/annual leave|personal|carer|compassion|long service|parental|bereave|bonus leave|in lieu taken|til taken|leave loading|unused leave|leave payment/.test(n)) return "leave";
    if (/back ?pay|bonus|per diem|allowance|reimburse|termination|redundan|commission|deduction|salary sacrifice|expense/.test(n)) return "exclude";
    if (/ordinary|overtime|shift|penalty|meal break|hours|time in lieu accru/.test(n)) return "work";
    return "work";
  }

  /** Casual or permanent, straight off the pay category name. */
  function guessEmployment(name) {
    var n = String(name || "").toLowerCase();
    if (/casual/.test(n)) return "Casual";
    if (/permanent|salaried|full ?time|part ?time/.test(n)) return "Permanent";
    return null;
  }

  /** The shift date at the front of a Notes entry, as in
   *  "7/09/2026 - 08:45 to 13:00.". A fortnight that crosses a month boundary
   *  then splits properly instead of landing wholly in the pay run's month. */
  function noteDate(v) {
    var m = String(v || "").match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (!m) return null;
    return m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  }

  function buildEarningsLoader() {
    var st = { parsed: null, col: {}, panel: h("div.ehpanel") };
    var ta = h("textarea.paste", { rows: 7,
      placeholder: "Paste the Employment Hero Earnings Details report here, title row and all" });
    var periodSel = h("select.control", {}, E.months.map(function (m) {
      return h("option", { value: m.key, selected: m.key === E.reportingMonth() }, m.long);
    }));

    function locMap() { return CTS.store.get("ehLocations", {}); }
    function catMap() { return CTS.store.get("ehCategories", {}); }

    function analyse() {
      st.panel.innerHTML = "";
      var t = parseTable(ta.value);
      if (!t.head.length || !t.rows.length) {
        st.panel.appendChild(h("div.banner.banner-warn",
          "Nothing to read. Copy the report including its header row and paste it in."));
        return;
      }
      st.parsed = t;
      var normed = t.head.map(norm);
      EARN_COLS.forEach(function (c) {
        var f = -1;
        for (var i = 0; i < c.names.length && f < 0; i++) f = normed.indexOf(c.names[i]);
        st.col[c.key] = f;
      });
      render();
    }

    function render() {
      var t = st.parsed, c = st.col;
      st.panel.innerHTML = "";
      var missing = EARN_COLS.filter(function (x) { return x.need && c[x.key] < 0; });
      var matched = EARN_COLS.filter(function (x) { return c[x.key] >= 0; });

      st.panel.appendChild(h("h4", "1. Columns"));
      st.panel.appendChild(U.note(t.rows.length + " rows read, " + matched.length +
        " of " + EARN_COLS.length + " expected columns matched by name."));
      EARN_COLS.filter(function (x) { return x.need || c[x.key] < 0; }).forEach(function (x) {
        st.panel.appendChild(h("div.fieldrow", [
          h("label", x.label + (x.need ? "" : " (optional)")),
          h("select.control", {
            onchange: function (e) { c[x.key] = +e.target.value; render(); },
          }, [h("option", { value: -1, selected: c[x.key] < 0 }, "— not present —")]
            .concat(t.head.map(function (hd, i) {
              return h("option", { value: i, selected: c[x.key] === i }, hd || "(col " + (i + 1) + ")");
            }))),
          c[x.key] >= 0 ? U.flag("good", "matched") : U.flag("warn", "needed"),
        ]));
      });
      if (missing.length) {
        st.panel.appendChild(h("div.banner.banner-warn", "Still need: " +
          missing.map(function (x) { return x.label; }).join(", ") + "."));
        return;
      }

      st.panel.appendChild(h("h4", "2. Which period this pay run belongs to"));
      st.panel.appendChild(h("div.fieldrow", [
        h("label", "Pay run period"), periodSel,
        h("span.muted", "Used only where a line's note carries no date"),
      ]));
      var withDate = c.notes >= 0
        ? t.rows.filter(function (r) { return noteDate(r[c.notes]); }).length : 0;
      st.panel.appendChild(U.note(withDate
        ? withDate + " of " + t.rows.length + " lines carry a shift date in the note, so a fortnight crossing a month boundary splits between the two months instead of landing wholly in one. The rest fall to the period above."
        : "No shift dates were found in the notes, so every line falls to the period above. A fortnight crossing a month end will land wholly in that period."));

      renderLocations();
      renderCategories();
      renderAdjustments();
      st.panel.appendChild(h("div.row", [
        h("button.btn", { onclick: apply }, "Load this pay run"),
        h("span.muted", "Adds to whatever is already loaded, month by month."),
      ]));
    }

    function scanLocations() {
      var t = st.parsed, c = st.col, out = {};
      t.rows.forEach(function (r) {
        var name = String(r[c.locName] || "").trim() || "(blank)";
        var o = out[name] || (out[name] = { name: name, rows: 0, hours: 0, gross: 0,
                 ext: c.locExt >= 0 ? String(r[c.locExt] || "") : "" });
        o.rows++;
        if (c.unitType < 0 || /hour/i.test(String(r[c.unitType]))) o.hours += num(r[c.units]);
        if (c.gross >= 0) o.gross += money(r[c.gross]);
      });
      return Object.keys(out).sort().map(function (k) { return out[k]; });
    }

    function renderLocations() {
      var locs = scanLocations(), saved = locMap();
      locs.forEach(function (l) {
        if (!saved[l.name]) {
          var dept = E.resolveDept("", l.name);            // the [ONS] bracket tag
          saved[l.name] = { dept: dept,
            // internal cost centres are not client work; everything else is
            chargeable: dept !== "ADMIN" && dept !== "UNALLOCATED" };
        }
      });
      CTS.store.set("ehLocations", saved);

      st.panel.appendChild(h("h4", "3. Locations"));
      st.panel.appendChild(U.note("The department is read from the bracket tag on the location name, the same way the ledger reads it off a job number. Whether a location is client work or internal is the one thing the export cannot say, so it is set here and remembered."));
      st.panel.appendChild(U.table([
        { key: "name", label: "Location in Employment Hero", align: "left" },
        { key: "ext", label: "External Id", align: "left" },
        { key: "rows", label: "Lines" },
        { key: "hours", label: "Hours", fmt: function (v) { return F.num(v, 2); } },
        { key: "gross", label: "Gross", fmt: function (v) { return "$" + F.num(v, 0); } },
        { key: "dept", label: "Department", align: "left",
          value: function (r) {
            return h("select.control.small", { onchange: function (e) {
              var m = locMap(); m[r.name].dept = e.target.value; CTS.store.set("ehLocations", m);
            } }, E.postingDepts.map(function (d) {
              return h("option", { value: d.code, selected: d.code === saved[r.name].dept }, d.short);
            }).concat([h("option", { value: "UNALLOCATED",
              selected: saved[r.name].dept === "UNALLOCATED" }, "Unallocated")]));
          } },
        { key: "chg", label: "Client work", align: "left",
          value: function (r) {
            return h("select.control.small", { onchange: function (e) {
              var m = locMap(); m[r.name].chargeable = e.target.value === "yes";
              CTS.store.set("ehLocations", m);
            } }, [h("option", { value: "yes", selected: saved[r.name].chargeable }, "Chargeable"),
                  h("option", { value: "no", selected: !saved[r.name].chargeable }, "Non-chargeable")]);
          } },
      ], locs, { dense: true }));
    }

    function scanCategories() {
      var t = st.parsed, c = st.col, out = {};
      t.rows.forEach(function (r) {
        var name = String(r[c.catName] || "").trim() || "(blank)";
        var o = out[name] || (out[name] = { name: name, rows: 0, hours: 0, gross: 0,
                 unit: c.unitType >= 0 ? String(r[c.unitType] || "") : "" });
        o.rows++;
        if (c.unitType < 0 || /hour/i.test(String(r[c.unitType]))) o.hours += num(r[c.units]);
        if (c.gross >= 0) o.gross += money(r[c.gross]);
      });
      return Object.keys(out).sort().map(function (k) { return out[k]; });
    }

    function renderCategories() {
      var cats = scanCategories(), saved = catMap();
      cats.forEach(function (x) { if (!saved[x.name]) saved[x.name] = guessCategory(x.name); });
      CTS.store.set("ehCategories", saved);

      st.panel.appendChild(h("h4", "4. Pay categories"));
      st.panel.appendChild(U.note("Worked time splits into chargeable or non-chargeable by its location. Leave and public holiday are paid but not worked, so they stay out of both sides of the utilisation ratio. Anything set to pay only, a back pay or a per diem for instance, still counts as cost but contributes no hours."));
      st.panel.appendChild(U.table([
        { key: "name", label: "Pay category", align: "left" },
        { key: "unit", label: "Unit type", align: "left" },
        { key: "rows", label: "Lines" },
        { key: "hours", label: "Hours", fmt: function (v) { return F.num(v, 2); } },
        { key: "gross", label: "Gross", fmt: function (v) { return "$" + F.num(v, 0); } },
        { key: "emp", label: "Employment", align: "left",
          value: function (r) {
            var e = guessEmployment(r.name);
            return e ? h("span.chip.chip-muted", e) : h("span.muted", "—");
          } },
        { key: "kind", label: "Counts as", align: "left",
          value: function (r) {
            return h("select.control.small", { onchange: function (e) {
              var m = catMap(); m[r.name] = e.target.value; CTS.store.set("ehCategories", m);
            } }, CAT_KINDS.map(function (k) {
              return h("option", { value: k.v, selected: k.v === saved[r.name] }, k.l);
            }));
          } },
      ], cats, { dense: true }));
    }

    /** A line that carries pay but no hours: a back pay, a bonus, a per diem
     *  in nights. Left alone it moves labour cost without touching utilisation,
     *  which is right for a bonus and wrong for a back pay: a back pay is hours
     *  somebody actually worked, they were just paid for them late. This table
     *  is where those become hours again, against a cost centre and a name. */
    function lineSig(r) {
      var c = st.col;
      return [c.empExt >= 0 ? r[c.empExt] : "", c.catName >= 0 ? r[c.catName] : "",
              c.gross >= 0 ? r[c.gross] : "",
              c.notes >= 0 ? String(r[c.notes] || "").slice(0, 28) : ""].join("|");
    }
    function adjustMap() { return CTS.store.get("ehLineAdjust", {}); }
    function setAdjust(sig, field, value) {
      var m = adjustMap();
      m[sig] = m[sig] || {};
      m[sig][field] = value;
      CTS.store.set("ehLineAdjust", m);
    }

    function payOnlyLines() {
      var t = st.parsed, c = st.col, cats = catMap(), out = [];
      t.rows.forEach(function (r, i) {
        var catName = String(r[c.catName] || "").trim() || "(blank)";
        var kind = cats[catName] || guessCategory(catName);
        var isHours = c.unitType < 0 || /hour/i.test(String(r[c.unitType]));
        var units = isHours ? num(r[c.units]) : 0;
        if (kind !== "exclude" && units > 0) return;
        out.push({ i: i, row: r, sig: lineSig(r), cat: catName,
          who: c.empName >= 0 ? String(r[c.empName] || "").trim() : "",
          loc: String(r[c.locName] || "").trim(),
          note: c.notes >= 0 ? String(r[c.notes] || "").replace(/\s+/g, " ").trim() : "",
          gross: c.gross >= 0 ? money(r[c.gross]) : 0,
          units: num(r[c.units]) });
      });
      return out;
    }

    function renderAdjustments() {
      var lines = payOnlyLines();
      st.panel.appendChild(h("h4", "5. Paid but not worked"));
      if (!lines.length) {
        st.panel.appendChild(U.note("Every line in this pay run carries hours. Nothing to reallocate."));
        return;
      }
      var adj = adjustMap(), locs = locMap();
      var locNames = Object.keys(locs).sort();
      st.panel.appendChild(U.note("These carry pay but no hours, so as they stand they move labour cost without touching utilisation. That is right for a bonus and wrong for a back pay: a back pay is hours somebody worked, paid late. Type the hours in and pick the cost centre they belong to, and they count as worked time against that person and department."));

      st.panel.appendChild(U.table([
        { key: "who", label: "Employee", align: "left",
          value: function (r) { return r.who || h("span.muted", "not named"); } },
        { key: "cat", label: "Pay category", align: "left" },
        { key: "note", label: "Note", align: "left",
          value: function (r) {
            return h("span", { title: r.note }, r.note.length > 44 ? r.note.slice(0, 43) + "\u2026" : r.note);
          } },
        { key: "gross", label: "Gross", fmt: function (v) { return "$" + F.num(v, 2); } },
        { key: "hours", label: "Hours worked", align: "left",
          value: function (r) {
            return h("input.control.small", { type: "number", step: "0.25", min: "0",
              style: { width: "5.5rem" },
              value: (adj[r.sig] && adj[r.sig].hours) || "",
              placeholder: "0",
              onchange: function (e) { setAdjust(r.sig, "hours", e.target.value); } });
          } },
        { key: "loc", label: "Cost centre to post to", align: "left",
          value: function (r) {
            var cur = (adj[r.sig] && adj[r.sig].loc) || r.loc;
            return h("select.control.small", {
              onchange: function (e) { setAdjust(r.sig, "loc", e.target.value); },
            }, locNames.map(function (n) {
              return h("option", { value: n, selected: n === cur },
                n + "  \u2192  " + ((E.deptOf[locs[n].dept] || {}).short || locs[n].dept));
            }));
          } },
        { key: "kind", label: "Counts as", align: "left",
          value: function (r) {
            var typed = parseFloat(adj[r.sig] && adj[r.sig].hours);
            var cur = (adj[r.sig] && adj[r.sig].kind) ||
                      (!isNaN(typed) && typed > 0 ? "work" : "exclude");
            return h("select.control.small", {
              onchange: function (e) { setAdjust(r.sig, "kind", e.target.value); },
            }, [{ v: "exclude", l: "Pay only" }, { v: "work", l: "Worked time" },
                { v: "leave", l: "Leave" }, { v: "publicHoliday", l: "Public holiday" }]
              .map(function (k) {
                return h("option", { value: k.v, selected: k.v === cur }, k.l);
              }));
          } },
      ], lines, { dense: true }));
      st.panel.appendChild(U.note("Enter hours and the line becomes worked time automatically; whether those hours are chargeable then follows the cost centre you pick, the same as every other line. Set it back to pay only if you want the cost without the hours. Entries are remembered, so re-pasting the same pay run keeps them.", "muted"));
    }

    function apply() {
      var t = st.parsed, c = st.col;
      var locs = locMap(), cats = catMap(), fallback = periodSel.value;
      var bucket = {}, labour = {}, people = {}, skipped = 0, monthsSeen = {};
      var usedFallback = 0, payOnly = {}, hourMonths = {};
      var adj = adjustMap(), staff = {}, reallocated = 0;

      t.rows.forEach(function (r) {
        var locName = String(r[c.locName] || "").trim() || "(blank)";
        var loc = locs[locName] || { dept: "UNALLOCATED", chargeable: false };
        var catName = String(r[c.catName] || "").trim() || "(blank)";
        var kind = cats[catName] || guessCategory(catName);
        var emp = guessEmployment(catName);

        var iso = c.notes >= 0 ? noteDate(r[c.notes]) : null;
        var mk = iso ? iso.slice(0, 7) : fallback;
        if (!iso) usedFallback++;
        if (!E.monthIdx[mk]) { skipped++; return; }
        monthsSeen[mk] = 1;

        var isHours = c.unitType < 0 || /hour/i.test(String(r[c.unitType]));
        var units = isHours ? num(r[c.units]) : 0;
        var gross = c.gross >= 0 ? money(r[c.gross]) : 0;
        var sup = c["super"] >= 0 ? money(r[c["super"]]) : 0;

        // a line the reallocation table has given hours and a cost centre
        var a = adj[lineSig(r)];
        if (a) {
          var ah = parseFloat(a.hours);
          if (a.loc && locs[a.loc]) { locName = a.loc; loc = locs[a.loc]; }
          if (a.kind) kind = a.kind;
          if (!isNaN(ah) && ah > 0 && units === 0) {
            units = ah;
            // Typing hours against a line is the whole point of the table, so
            // a line still marked pay only becomes worked time rather than
            // reporting a reallocation that went nowhere.
            if (kind === "exclude") kind = "work";
            reallocated++;
          }
        }

        var key = mk + "|" + loc.dept;
        var b = bucket[key] || (bucket[key] = { month: mk, dept: loc.dept,
                  chargeable: 0, nonChargeable: 0, leave: 0, publicHoliday: 0 });
        if (kind === "work") b[loc.chargeable ? "chargeable" : "nonChargeable"] += units;
        else if (kind === "leave") b.leave += units;
        else if (kind === "publicHoliday") b.publicHoliday += units;

        var lb = labour[key] || (labour[key] = { month: mk, dept: loc.dept, gross: 0, sup: 0 });
        lb.gross += gross; lb.sup += sup;

        // A line carrying hours is what makes a month real for someone. A pay
        // only line, a back pay or a per diem, must not: one of those with no
        // date would otherwise invent a month and make everybody look like a
        // new starter.
        if (units > 0) hourMonths[mk] = 1;
        if (kind === "exclude" || units === 0) {
          var po = payOnly[catName] || (payOnly[catName] = { cat: catName, lines: 0,
                     gross: 0, dated: 0 });
          po.lines++; po.gross += gross; if (iso) po.dated++;
        }

        var pid = (c.empExt >= 0 && String(r[c.empExt]).trim()) ||
                  (c.empId >= 0 && String(r[c.empId]).trim()) || catName;
        var p = people[pid] || (people[pid] = { id: pid,
                  name: c.empName >= 0 ? String(r[c.empName] || "").trim() : pid, months: {} });
        var pm = p.months[mk] || (p.months[mk] = { hours: 0, gross: 0, types: {}, depts: {} });
        pm.hours += units; pm.gross += gross;
        if (units > 0) p.hasHours = true;

        // person level, which is what the profitability by employee page reads
        var skey = mk + "|" + pid + "|" + loc.dept;
        var sr = staff[skey] || (staff[skey] = { month: mk, empId: pid,
          name: (c.empName >= 0 ? String(r[c.empName] || "").trim() : pid) || pid,
          dept: loc.dept, chargeable: 0, nonChargeable: 0, leave: 0,
          publicHoliday: 0, gross: 0, sup: 0, employment: null });
        if (kind === "work") sr[loc.chargeable ? "chargeable" : "nonChargeable"] += units;
        else if (kind === "leave") sr.leave += units;
        else if (kind === "publicHoliday") sr.publicHoliday += units;
        sr.gross += gross; sr.sup += sup;
        if (emp) sr.employment = emp;
        if (emp) pm.types[emp] = (pm.types[emp] || 0) + 1;
        pm.depts[loc.dept] = (pm.depts[loc.dept] || 0) + 1;
      });

      var rows = Object.keys(bucket).sort().map(function (k) {
        var b = bucket[k];
        function r1(v) { return Math.round(v * 100) / 100; }
        return [b.month, b.dept, r1(b.chargeable), r1(b.nonChargeable),
                r1(b.leave), r1(b.publicHoliday), 0];
      });

      st.panel.appendChild(h("h4", "5. What was loaded"));
      if (!rows.length) {
        st.panel.appendChild(h("div.banner.banner-warn", "Nothing could be loaded."));
        return;
      }

      var payload = { meta: { seed: false, built: new Date().toISOString().slice(0, 10),
                              hoursPerDay: E.hoursPerDay(), state: E.utilState(),
                              source: "Employment Hero Earnings Details" },
                      cols: ["month", "dept", "chargeable", "nonChargeable", "leave",
                             "publicHoliday", "fte"],
                      rows: rows };
      E.loadUtil(payload);
      E.labour = labour;

      // people, so Profitability by Employee reads the real pay run
      var staffRows = Object.keys(staff).sort().map(function (k) {
        var x = staff[k];
        function r2(v) { return Math.round(v * 100) / 100; }
        return [x.month, x.empId, x.name, x.dept, r2(x.chargeable), r2(x.nonChargeable),
                r2(x.leave), r2(x.publicHoliday), Math.round(x.gross * 100),
                Math.round(x.sup * 100), x.employment];
      });
      E.loadStaff({ meta: { seed: false, built: payload.meta.built,
                            note: "From the Employment Hero earnings report." },
                    cols: ["month", "empId", "name", "dept", "chargeable", "nonChargeable",
                           "leave", "publicHoliday", "grossCents", "superCents", "employment"],
                    rows: staffRows });

      var keys = Object.keys(monthsSeen).sort();
      var u = E.utilFor(keys);
      st.panel.appendChild(h("div.banner.banner-good",
        rows.length + " department months loaded from " + t.rows.length + " earnings lines, " +
        staffRows.length + " people months" +
        (reallocated ? ", " + reallocated + " reallocated to worked hours" : "") +
        (skipped ? ", " + skipped + " skipped for a period outside the two loaded years" : "") + "."));

      st.panel.appendChild(U.table([
        { key: "d", label: "Department", align: "left" },
        { key: "c", label: "Chargeable", fmt: F.hours },
        { key: "n", label: "Non-chargeable", fmt: F.hours },
        { key: "l", label: "Leave", fmt: F.hours },
        { key: "p", label: "Public holiday", fmt: F.hours },
        { key: "u", label: "Utilisation", fmt: function (v) { return F.pct(v); } },
        { key: "g", label: "Gross earnings", fmt: function (v) { return "$" + F.num(v, 0); } },
        { key: "s", label: "SG super", fmt: function (v) { return "$" + F.num(v, 0); } },
      ], u.rows.filter(function (r) { return r.total; }).map(function (r) {
        var g = 0, sp = 0;
        Object.keys(labour).forEach(function (k) {
          if (labour[k].dept === r.code) { g += labour[k].gross; sp += labour[k].sup; }
        });
        return { d: r.dept.short, c: r.chargeable, n: r.nonChargeable, l: r.leave,
                 p: r.publicHoliday, u: r.util, g: g, s: sp };
      })));

      /* People changes.
       *
       * Only months in which someone was actually paid for hours count, and
       * a first or last month is only worth calling out when there are enough
       * months either side to mean anything. Casuals work some months and not
       * others, so on a short span every casual would look like a starter and
       * a leaver at once. Employment type changes are reported whatever the
       * span, because a category changing from casual to permanent says so on
       * its own. */
      var allMonths = Object.keys(hourMonths).sort();
      var enoughSpan = allMonths.length >= 3;
      var changes = [];
      Object.keys(people).forEach(function (id) {
        var p = people[id];
        if (!p.hasHours) return;
        var pm = Object.keys(p.months).filter(function (m) {
          return p.months[m].hours > 0;
        }).sort();
        if (!pm.length) return;
        var typeOf = function (m) {
          var ts = Object.keys(p.months[m].types);
          return ts.length ? ts.sort(function (a, b) {
            return p.months[m].types[b] - p.months[m].types[a]; })[0] : null;
        };
        for (var i = 1; i < pm.length; i++) {
          var a = typeOf(pm[i - 1]), b2 = typeOf(pm[i]);
          if (a && b2 && a !== b2) {
            changes.push({ who: p.name, what: a + " to " + b2,
              when: (E.monthIdx[pm[i]] || {}).long || pm[i], sort: 0,
              detail: "Pay category changed from " + a.toLowerCase() + " to " + b2.toLowerCase() });
          }
        }
        if (!enoughSpan) return;
        if (pm[0] > allMonths[0]) {
          changes.push({ who: p.name, what: "First month with hours",
            when: (E.monthIdx[pm[0]] || {}).long || pm[0], sort: 1,
            detail: "Nothing paid before this month across the pay runs loaded" });
        }
        if (pm[pm.length - 1] < allMonths[allMonths.length - 1]) {
          changes.push({ who: p.name, what: "Last month with hours",
            when: (E.monthIdx[pm[pm.length - 1]] || {}).long || pm[pm.length - 1], sort: 2,
            detail: "Nothing paid after this month. Normal for a casual between jobs." });
        }
      });
      changes.sort(function (a, b) { return (a.sort - b.sort) || (a.who < b.who ? -1 : 1); });

      st.panel.appendChild(h("h4", "People changes the export shows"));
      if (changes.length) {
        st.panel.appendChild(U.table([
          { key: "who", label: "Employee", align: "left" },
          { key: "what", label: "Change", align: "left" },
          { key: "when", label: "Month", align: "left" },
          { key: "detail", label: "How it was spotted", align: "left" },
        ], changes, { dense: true }));
        st.panel.appendChild(U.note("Read off the pay categories, not from a list anyone maintains: someone moving from casual to permanent changes from Casual Ordinary Hours to Permanent Ordinary Hours, and that is what this is reading. Load several pay runs together and it covers the whole span."));
      } else {
        st.panel.appendChild(U.note(allMonths.length < 2
          ? "Only one month carries hours, so there is nothing to compare against. Paste several pay runs together and casual to permanent moves, starters and leavers are listed here."
          : "Nobody changed employment type across the months loaded."));
      }
      if (!enoughSpan) {
        st.panel.appendChild(U.note("First and last months are not reported on a span of " +
          allMonths.length + (allMonths.length === 1 ? " month" : " months") +
          ". Casuals work some months and not others, so on a short span every casual reads as a starter and a leaver at once. Three months or more and they appear.", "muted"));
      }

      // The adjustment lines: back pay, per diem, anything paid but not worked.
      var po = Object.keys(payOnly).sort().map(function (k) { return payOnly[k]; });
      if (po.length) {
        st.panel.appendChild(h("h4", "Paid but not worked"));
        st.panel.appendChild(U.table([
          { key: "cat", label: "Pay category", align: "left" },
          { key: "lines", label: "Lines" },
          { key: "dated", label: "With a date in the note" },
          { key: "gross", label: "Gross", fmt: function (v) { return "$" + F.num(v, 0); } },
        ], po, { dense: true }));
        st.panel.appendChild(U.note("These carry cost but no hours, so they move labour cost without touching utilisation. A back pay belongs to the period it corrects, not the one it is paid in: where the note carries no date it has fallen to the pay run period above, and " + (po.reduce(function (a, x) { return a + (x.lines - x.dated); }, 0)) + " of these did."));
      }
      if (usedFallback) {
        st.panel.appendChild(h("div.banner.banner-warn", [
          h("strong", usedFallback + " of " + t.rows.length + " lines had no date in their note. "),
          "They were put in " + ((E.monthIdx[fallback] || {}).long || fallback) +
          ". If any of them belong to an earlier period, load that pay run separately with the period set to the month it corrects.",
        ]));
      }

      var staffPayload = { meta: { seed: false, built: payload.meta.built,
                                   note: "From the Employment Hero earnings report." },
                           cols: ["month", "empId", "name", "dept", "chargeable",
                                  "nonChargeable", "leave", "publicHoliday",
                                  "grossCents", "superCents", "employment"],
                           rows: staffRows };
      st.panel.appendChild(h("h4", "Keeping this load"));
      st.panel.appendChild(U.note("Without one of these the load lasts only until you reload the page."));
      st.panel.appendChild(keepOrShow({
        name: "CTS_util_data.js",
        keep: [{ kind: "util", payload: function () { return payload; } },
               { kind: "staff", payload: function () { return staffPayload; } }],
        payload: function () { return payload; },
        text: function () {
          return "// Loaded " + payload.meta.built + " from the Employment Hero earnings report.\n" +
                 "window.CTS_UTIL = " + JSON.stringify(payload) + ";\n";
        },
      }));
      st.panel.appendChild(U.note("Showing the file gives the utilisation half. The people half is CTS_staff_data.js; keeping it in the browser covers both.", "muted"));
    }

    return h("div.card.loader.wide", [
      h("h3", "Utilisation and labour cost, from the Employment Hero earnings report"),
      U.note("Run the Earnings Details report for a pay run, copy it in Excel including the header row, and paste it here. The title row above the headings is fine, and so is a Notes column with line breaks inside it."),
      U.note("The department comes from the bracket tag on the location name, as in PWC IT Support - Brisbane [ONS]. Casual against permanent comes from the pay category. Both are read out of the export rather than maintained anywhere.", "muted"),
      ta,
      h("div.row", [h("button.btn", { onclick: analyse }, "Read the columns")]),
      st.panel,
    ]);
  }

  /* The Build module and the paste loaders share one set of parsing rules. */
  CTS.parse = { isoDate: isoDate, num: num, money: money, noteDate: noteDate,
                guessCategory: guessCategory, guessEmployment: guessEmployment,
                parseTable: parseTable, download: download };

  /* ========================================================== Setup ==== */
  P.setup = {
    section: "Admin", title: "Setup",
    sub: "The controls everything else follows from.",
    render: function () {
      var u = E.utilFor(E.ytd());
      function field(label, control, note) {
        return h("div.fieldrow", [h("label", label), control,
                                  note ? h("span.muted", note) : null]);
      }
      return [
        seedBanner(), U.h1("Setup", "One month, a handful of thresholds, and the rest calculates."),
        U.section("What am I reporting on", [
          field("Reporting month",
            h("select.control", { onchange: function (e) { E.setReportingMonth(e.target.value); CTS.router.reload(); } },
              E.months.map(function (m) {
                return h("option", { value: m.key, selected: m.key === E.reportingMonth() }, m.long);
              })),
            "The one control you change each month. The financial year, the period, every comparative and the actual-then-budget cut all follow from it."),
          field("Financial year", h("span.readout", "FY" + E.currentFY()),
            "1 July to 30 June, named for the year it ends"),
          field("Period in the year",
            h("span.readout", String((E.monthIdx[E.reportingMonth()] || {}).period || "-")), null),
          field("Comparative month last year",
            h("span.readout", (E.monthIdx[E.priorYearMonths([E.reportingMonth()])[0]] || {}).long || "-"), null),
        ]),
        U.section("Utilisation", [
          field("Hours in a working day",
            h("input.control", { type: "number", min: "1", max: "24", value: E.hoursPerDay(),
              onchange: function (e) { CTS.store.set("hoursPerDay", +e.target.value); CTS.router.reload(); } }), null),
          field("State for working days",
            h("select.control", { onchange: function (e) { CTS.store.set("utilState", e.target.value); CTS.router.reload(); } },
              E.cal.states.map(function (s) {
                return h("option", { value: s, selected: s === E.utilState() }, s);
              })),
            "Public holidays differ by state, so the working day count does too"),
          field("Working days, year to date", h("span.readout", String(u.workingDays)), null),
        ]),
        U.section("Risk flag thresholds", [
          U.note("These are the Controller Pack's own thresholds. The materiality floor is tested before the percentage, so a variance below it is always Low however large the percentage."),
          U.table([
            { key: "l", label: "Threshold", align: "left" },
            { key: "v", label: "Value" },
            { key: "n", label: "", align: "left" },
          ], [
            { l: "Materiality floor", v: F.dollars(E.cfg.RISK.materialityFloor * 100), n: "Below this a variance is always Low" },
            { l: "Medium risk", v: F.pct(E.cfg.RISK.mediumPct, 0), n: "Percentage variance" },
            { l: "High risk, percentage", v: F.pct(E.cfg.RISK.highPct, 0), n: "" },
            { l: "High risk, dollars", v: F.dollars(E.cfg.RISK.highDollar * 100), n: "Either test alone makes it High" },
          ]),
        ]),
        U.section("Data status", U.table([
          { key: "l", label: "", align: "left" },
          { key: "v", label: "" },
        ], [
          { l: "Ledger lines loaded", v: F.num(E.quality.rows) },
          { l: "Earliest ledger date", v: F.date(E.quality.minDate) },
          { l: "Latest ledger date", v: F.date(E.quality.maxDate) },
          { l: "Lines with no department", v: F.num(E.quality.noDept) },
          { l: "Lines on an unknown account", v: F.num(E.quality.unknownAcct) },
          { l: "Accounts in the chart", v: F.num(E.accounts.length) },
          { l: "Months with a posted actual", v: F.num(E.actualMonths.length) },
          { l: "Data origin", v: E.seed ? "seed" : (E.restored && E.restored.length ? "kept in this browser" : "loaded this session") },
        ])),
        U.section("Data kept in this browser", (function () {
          var snaps = E.snapshots();
          var keys = Object.keys(snaps);
          var label = { gl: "Ledger", fin: "Profit and loss", util: "Utilisation", staff: "People" };
          return [
            U.note("A load is kept here so it survives a reload, on this machine and for you only. It is not shared with anyone else and it is not what the portal loads for everybody. For that the file has to go in portal/data, which the Show the file button on each loader gives you."),
            keys.length ? U.table([
              { key: "what", label: "Data set", align: "left" },
              { key: "when", label: "Kept", align: "left" },
              { key: "size", label: "Size" },
              { key: "clear", label: "", align: "left", value: function (r) { return r.clear; } },
            ], keys.map(function (k) {
              var raw = JSON.stringify(snaps[k].payload).length;
              return { what: label[k] || k,
                       when: F.date(String(snaps[k].saved).slice(0, 10)),
                       size: (raw / 1024).toFixed(0) + " KB",
                       clear: h("button.btn.btn-quiet.small", { onclick: function () {
                         if (!confirm("Discard the kept " + (label[k] || k) + " and go back to the file?")) return;
                         E.clearSnapshots(k);
                         location.reload();
                       } }, "Discard") };
            })) : U.note("Nothing kept. The portal is reading the files in portal/data.", "muted"),
            keys.length ? h("div.row", [
              h("button.btn.btn-quiet", { onclick: function () {
                if (!confirm("Discard every kept data set and go back to the files?")) return;
                E.clearSnapshots();
                location.reload();
              } }, "Discard everything kept"),
            ]) : null,
          ];
        })()),
      ];
    },
  };

  /* ======================================================== Loaders ==== */
  /* ========================================================= Config ==== */
  P.config = {
    section: "Admin", title: "Config & Variables",
    sub: "Departments, tags, splits and holidays.",
    render: function () {
      var s = E.subSplits();
      var deptTable = U.table([
        { key: "short", label: "Department", align: "left",
          value: function (d) { return h("span", [h("span.dot", { style: { background: U.colourOf(d.code) } }), d.short]); } },
        { key: "long", label: "Full name", align: "left" },
        { key: "order", label: "Order" },
        { key: "isRevenue", label: "Revenue", align: "left",
          value: function (d) { return d.isRevenue ? "yes" : "no"; } },
        { key: "tags", label: "Tags that resolve to it", align: "left",
          value: function (d) { return (d.tags || []).join(", ") || "-"; } },
        { key: "suffix", label: "Account suffix", align: "left",
          value: function (d) { return d.suffix || "-"; } },
        { key: "utilTarget", label: "Util target",
          fmt: function (v) { return v == null ? "-" : F.pct(v, 0); } },
        { key: "gmNorm", label: "GM norm",
          fmt: function (v) { return v == null ? "-" : F.pct(v, 0); } },
      ], E.depts);

      var splitTable = U.table([
        { key: "sub", label: "Overhead line", align: "left" },
        { key: "basis", label: "Basis", align: "left" },
        { key: "src", label: "", align: "left" },
      ], Object.keys(E.idx.deptSubMonth.ADMIN || {}).filter(function (sub) {
        return (E.subsOf.expenses || {})[sub];
      }).sort().map(function (sub) {
        var code = E.basisFor(sub);
        return { sub: sub, basis: E.basis(code).label,
                 src: E.basis(code).source === "confirmed"
                   ? U.flag("good", "confirmed") : U.flag("warn", "placeholder") };
      }), { dense: true });

      var hol = E.cal.holidays.filter(function (x) {
        return x.date >= "2026-07-01" && x.date <= "2027-06-30";
      }).sort(function (a, b) { return a.date < b.date ? -1 : 1; });

      return [
        U.h1("Config & Variables", "What the portal is standing on."),
        U.section("Departments", deptTable,
          "A ledger line resolves to a department off the Cost Centres column first, then the bracket tag on the job number. Longest tag wins, so ONSITE beats ONS."),
        U.section("Overhead split bases", [
          U.table([
            { key: "label", label: "Basis", align: "left" },
            { key: "prd", label: "Production + Video", fmt: function (v) { return F.pct(v); } },
            { key: "ons", label: "Onsite", fmt: function (v) { return F.pct(v); } },
            { key: "ci", label: "Consulting + Integration", fmt: function (v) { return F.pct(v); } },
            { key: "src", label: "", align: "left" },
          ], E.cfg.SPLIT_BASES.map(function (b) {
            return { label: b.label, prd: b.groups["PRD/VID"], ons: b.groups["ONS"],
                     ci: b.groups["CONS/INT"],
                     src: b.source === "confirmed" ? U.flag("good", "confirmed")
                        : U.flag("warn", "placeholder") };
          })),
          U.note("The Staff and Office Dept percentages live in the budget workbook and are not in this repository. Until they are entered, every allocated departmental figure in this portal rests on a placeholder, and the portal says so wherever it uses one.", "warn"),
          h("div.fieldrow", [h("label", "Production share of production and video"),
            h("span.readout", F.pct(s.prd, 0))]),
          h("div.fieldrow", [h("label", "Integration share of consulting and integration"),
            h("span.readout", F.pct(s.int, 0))]),
        ]),
        U.section("Which basis each overhead line uses", splitTable,
          "Change any of these on the Overhead Allocation page. On the live allocation tab the basis is labelled against every row, so correct anything that disagrees."),
        U.section("Public holidays, FY" + E.currentFY(), U.table([
          { key: "date", label: "Date", align: "left", fmt: F.date },
          { key: "name", label: "Holiday", align: "left" },
          { key: "states", label: "States", align: "left",
            value: function (r) { return r.states.length === 8 ? "all" : r.states.join(", "); } },
          { key: "certain", label: "", align: "left",
            value: function (r) {
              return r.certain ? U.flag("good", "fixed or computed")
                               : U.flag("warn", "verify");
            } },
        ], hol, { dense: true }),
          "The fixed-date and Easter-derived holidays are certain. The rest are state proclamations that do move; verify them against your state's official list before the utilisation figures are relied on."),
        U.section("Reset", h("div.row", [
          h("button.btn.btn-quiet", { onclick: function () {
            if (!confirm("Clear every setting saved in this browser and reload?")) return;
            try {
              Object.keys(localStorage).forEach(function (k) {
                if (k.indexOf("cts.bi.") === 0) localStorage.removeItem(k);
              });
            } catch (e) {}
            location.reload();
          } }, "Clear saved settings"),
          h("span.muted", "Reporting month, split overrides, period and filters are saved in this browser only."),
        ])),
      ];
    },
  };

  /* ============================================ Profit by Employee ===== */
  P["staff-profit"] = {
    section: "Utilisation", title: "Profitability by Employee",
    sub: "Who is carrying the work, and what it costs.",
    render: function () {
      var p = CTS.period();
      if (!E.hasStaff) {
        return [CTS.seedBanner(), CTS.periodBar(),
          U.h1("Profitability by Employee", p.label),
          h("div.banner.banner-warn", [
            h("strong", "No people loaded. "),
            "This page reads the Employment Hero earnings export, which is the only source that carries who worked, on what, and what they were paid. Paste it into 03 Employment Hero Earnings.xlsx and run ",
            h("a", { href: "#/build" }, "Build"), ".",
          ])];
      }
      var showAll = CTS.store.get("staffShowAll", false);
      var res = E.staffFor(p.keys);
      var scope = E.deptScope();
      var rows = res.rows.filter(function (r) {
        if (scope && r.dept !== scope) return false;
        return showAll || r.billable;
      });

      var tot = rows.reduce(function (a, r) {
        a.chg += r.chargeable; a.worked += r.worked; a.cost += r.cost;
        a.rev += r.revenue; a.margin += r.margin; return a;
      }, { chg: 0, worked: 0, cost: 0, rev: 0, margin: 0 });

      var top = rows.slice(0, 12);
      // Two series, so colour means margin against cost and nothing else.
      // Department identity sits in the table beside it rather than being
      // painted onto bars whose legend already says something different.
      var chart = U.bars({
        rows: top.map(function (r) {
          return { label: r.name + "  \u00b7  " + r.deptShort,
                   value: r.margin, value2: r.cost };
        }),
        label1: "Margin over labour", label2: "Labour cost",
        width: 620, labelWidth: 210,
      });

      return [
        CTS.seedBanner(), CTS.periodBar(),
        U.h1("Profitability by Employee", p.label),
        h("div.banner.banner-warn", [
          h("strong", "Revenue here is attributed, not earned. "),
          "Nothing bills per person, so each department's revenue is spread across its people on their share of that department's chargeable hours. It is a fair reading of who is carrying the work. It is not what any individual brought in, and it should not be used as one.",
        ]),
        h("div.tiles", [
          U.tile({ label: "Billable people", value: String(rows.filter(function (r) { return r.billable; }).length),
                   sub: showAll ? "showing everyone" : "with chargeable hours this period" }),
          U.tile({ label: "Chargeable hours", value: F.hours(tot.chg),
                   sub: F.pct(tot.worked ? tot.chg / tot.worked : null) + " of hours worked" }),
          U.tile({ label: "Labour cost", value: F.dollars(tot.cost),
                   sub: "gross plus superannuation" }),
          U.tile({ label: "Margin over labour", value: F.dollars(tot.margin),
                   tone: tot.margin >= 0 ? "good" : "critical",
                   sub: tot.rev ? F.pct(tot.margin / tot.rev) + " of attributed revenue" : null }),
        ]),
        h("div.toolbar", [
          h("label.checkbox", [
            h("input", { type: "checkbox", checked: showAll,
              onchange: function (e) { CTS.store.set("staffShowAll", e.target.checked); CTS.router.reload(); } }),
            " Include people with no chargeable hours",
          ]),
          h("div.toolbar-right", h("span.muted",
            "Recovery is attributed revenue over labour cost. Under 1.0 means the person costs more than the work attributed to them.")),
        ]),
        U.section("Margin against cost, top twelve",
          U.figure("Ranked by margin over labour cost", chart, null)),
        U.section("Every person", U.table([
          { key: "name", label: "Employee", align: "left" },
          { key: "deptShort", label: "Department", align: "left",
            value: function (r) {
              return h("span", [h("span.dot", { style: { background: U.colourOf(r.dept) } }), r.deptShort]);
            } },
          { key: "employment", label: "Type", align: "left",
            value: function (r) {
              return r.employment ? h("span.chip.chip-muted", r.employment) : h("span.muted", "\u2014");
            } },
          { key: "chargeable", label: "Chargeable", fmt: F.hours },
          { key: "nonChargeable", label: "Non-charge", fmt: F.hours },
          { key: "util", label: "Util", fmt: function (v) { return F.pct(v, 0); } },
          { key: "chargeRate", label: "Rate per chargeable hour", fmt: function (v) { return v == null ? "-" : F.dollars(v); } },
          { key: "costRate", label: "Cost per hour worked", fmt: function (v) { return v == null ? "-" : F.dollars(v); } },
          { key: "revenue", label: "Revenue attributed", fmt: F.money },
          { key: "cost", label: "Labour cost", fmt: F.money },
          { key: "margin", label: "Margin", fmt: F.money, cell: U.moneyCell },
          { key: "recovery", label: "Recovery", fmt: function (v) { return v == null ? "-" : v.toFixed(2) + "x"; },
            cell: function (r) { return r.recovery != null && r.recovery < 1 ? "neg" : ""; } },
        ], rows.concat([{ name: "Total", deptShort: "", employment: null,
            chargeable: tot.chg, nonChargeable: tot.worked - tot.chg,
            util: tot.worked ? tot.chg / tot.worked : null,
            chargeRate: tot.chg ? Math.round(tot.rev / tot.chg) : null,
            costRate: tot.worked ? Math.round(tot.cost / tot.worked) : null,
            revenue: tot.rev, cost: tot.cost, margin: tot.margin,
            recovery: tot.cost ? tot.rev / tot.cost : null, _cls: "totalrow" }]), { dense: true }),
          "Cost is gross earnings plus superannuation, straight off the pay run. Leave and public holidays are in the cost but not in the hours worked, which is why somebody on leave shows a higher cost per hour."),
      ];
    },
  };

  /* ============================================================ Build == */
  /* The folder handle is kept in IndexedDB so the Build and Distribution
   * pages can reuse it after a reload with one permission prompt, rather
   * than a fresh picker every time. */
  var DIRDB = { name: "cts.bi", store: "handles", key: "portalDir" };
  function dirDb() {
    return new Promise(function (res, rej) {
      try {
        var r = indexedDB.open(DIRDB.name, 1);
        r.onupgradeneeded = function () { r.result.createObjectStore(DIRDB.store); };
        r.onsuccess = function () { res(r.result); };
        r.onerror = function () { rej(r.error); };
      } catch (e) { rej(e); }
    });
  }
  CTS.saveDir = function (handle) {
    return dirDb().then(function (db) {
      return new Promise(function (res) {
        var tx = db.transaction(DIRDB.store, "readwrite");
        tx.objectStore(DIRDB.store).put(handle, DIRDB.key);
        tx.oncomplete = function () { res(true); };
        tx.onerror = function () { res(false); };
      });
    }).catch(function () { return false; });
  };
  CTS.loadDir = function () {
    return dirDb().then(function (db) {
      return new Promise(function (res) {
        var tx = db.transaction(DIRDB.store, "readonly");
        var g = tx.objectStore(DIRDB.store).get(DIRDB.key);
        g.onsuccess = function () { res(g.result || null); };
        g.onerror = function () { res(null); };
      });
    }).catch(function () { return null; });
  };
  /** Reuse a saved handle if the browser still lets us write to it. */
  CTS.reuseDir = async function () {
    var h = await CTS.loadDir();
    if (!h) return null;
    try {
      var q = await h.queryPermission({ mode: "readwrite" });
      if (q === "granted") return h;
      var r = await h.requestPermission({ mode: "readwrite" });
      return r === "granted" ? h : null;
    } catch (e) { return null; }
  };

  function checksTable(checks) {
    var rank = { fail: 0, warn: 1, info: 2, ok: 3 };
    var rows = checks.slice().sort(function (a, b) { return rank[a.level] - rank[b.level]; });
    return U.table([
      { key: "level", label: "", align: "left",
        value: function (r) {
          var cls = { ok: "good", warn: "warning", fail: "critical", info: "muted" }[r.level];
          return h("span.chip.chip-" + cls, r.level === "ok" ? "OK" : r.level === "fail" ? "FAIL" : r.level === "warn" ? "CHECK" : "NOTE");
        } },
      { key: "area", label: "Area", align: "left" },
      { key: "text", label: "", align: "left" },
    ], rows, { dense: true });
  }

  P.build = {
    section: "Admin", title: "Build",
    sub: "Templates in, data files out.",
    render: function () {
      var B = CTS.build;
      var state = P.build._state || (P.build._state = { dir: null, result: null });
      var out = h("div");

      if (!B) {
        return [U.h1("Build", "The build module did not load"),
                U.note("CTS_build.js or vendor/xlsx.full.min.js is missing from the portal folder.", "warn")];
      }
      if (!B.supported()) {
        return [U.h1("Build", "Needs Edge or Chrome"),
          h("div.banner.banner-warn", [h("strong", "This browser cannot write to the folder. "),
            "The Build page reads the templates and writes the data files with the File System Access API, which Edge and Chrome have and Firefox and Safari do not. Open the portal in Edge to build; anyone only viewing it can use any browser."]),
          U.note("If the portal is open from a link rather than the synced folder, there is no folder to write to either. Build from the synced copy.")];
      }

      function stepEl(n, title, kids) {
        return h("div.step", [h("div.step-n", String(n)), h("div", [h("h3", title)].concat(kids))]);
      }
      var folderStatus = h("div");
      var readStatus = h("div");
      var writeStatus = h("div");

      async function chooseFolder(reuse) {
        folderStatus.innerHTML = "";
        try {
          var handle = reuse ? await CTS.reuseDir() : null;
          var ok = !!handle;
          if (!handle) { var picked = await B.pickFolder(); handle = picked.handle; ok = picked.ok; }
          else { try { await handle.getFileHandle("CTS_bi_bundle.js"); await handle.getDirectoryHandle("templates"); ok = true; } catch (e) { ok = false; } }
          if (!ok) {
            folderStatus.appendChild(h("div.banner.banner-warn", "That folder does not look like the portal: it needs CTS_bi_bundle.js and a templates folder in it. Pick the synced CTS Business Portal folder."));
            return;
          }
          state.dir = handle;
          await CTS.saveDir(handle);
          folderStatus.appendChild(h("div.banner.banner-good", "Folder chosen: " + (handle.name || "portal") + ". The templates and data folders are there."));
        } catch (e) {
          if (e && e.name === "AbortError") return;
          folderStatus.appendChild(h("div.banner.banner-warn", "Could not open the folder: " + (e && e.message || e)));
        }
      }

      async function readTemplates() {
        readStatus.innerHTML = "";
        if (!state.dir) { readStatus.appendChild(h("div.banner.banner-warn", "Choose the portal folder first.")); return; }
        readStatus.appendChild(U.note("Reading the templates…"));
        try {
          var r = await B.run(state.dir);
          state.result = r;
          readStatus.innerHTML = "";
          var fails = r.checks.filter(function (c) { return c.level === "fail"; }).length;
          var warns = r.checks.filter(function (c) { return c.level === "warn"; }).length;
          readStatus.appendChild(h("div.banner.banner-" + (fails ? "warn" : "good"), [
            h("strong", fails ? fails + " check" + (fails > 1 ? "s" : "") + " failed. " : "Templates read. "),
            fails ? "Fix the red rows in the template and read again; nothing has been written."
                  : (warns ? warns + " thing" + (warns > 1 ? "s" : "") + " to look at below, none of them stop a write." : "Nothing to look at.")]));
          readStatus.appendChild(checksTable(r.checks));
          if (r.missing && r.missing.length) readStatus.appendChild(U.note("Not found in templates: " + r.missing.join(", ") + ".", "muted"));
          readStatus.appendChild(U.table([
            { key: "f", label: "Data file to write", align: "left" },
            { key: "s", label: "Size", fmt: function (v) { return (v / 1024).toFixed(0) + " KB"; } },
          ], Object.keys(r.outputs).map(function (f) { return { f: f, s: r.outputs[f].length }; }), { dense: true }));
        } catch (e) {
          readStatus.innerHTML = "";
          readStatus.appendChild(h("div.banner.banner-warn", "Could not read the templates: " + (e && e.message || e)));
        }
      }

      async function writeFiles() {
        writeStatus.innerHTML = "";
        if (!state.result) { writeStatus.appendChild(h("div.banner.banner-warn", "Read the templates first.")); return; }
        if (state.result.checks.some(function (c) { return c.level === "fail"; })) {
          writeStatus.appendChild(h("div.banner.banner-warn", "A check failed. Fix it and read again before writing."));
          return;
        }
        var prog = h("p.note", "Writing…");
        writeStatus.appendChild(prog);
        try {
          var month = (state.result.log && state.result.log.reportingMonth) || E.reportingMonth();
          var done = await B.write(state.dir, state.result.outputs, function (name, i, n) {
            prog.textContent = "Writing " + name + " (" + i + " of " + n + ")";
          }, { month: month, buffers: state.result.buffers });
          writeStatus.innerHTML = "";
          writeStatus.appendChild(h("div.banner.banner-good", [
            h("strong", done.length + " data files written. "),
            "The previous set is in data/_previous and this month's data and templates are in archive/" + month + ". Reloading the portal on the new data…"]));
          setTimeout(function () { location.reload(); }, 1500);
        } catch (e) {
          writeStatus.innerHTML = "";
          writeStatus.appendChild(h("div.banner.banner-warn", "Writing failed: " + (e && e.message || e) + ". Nothing may have been written, or some files may have: check data/_previous."));
        }
      }

      var log = window.CTS_BUILD_LOG || null;
      return [
        U.h1("Build", "Templates in, data files out."),
        U.note("Reads the eight Excel templates from the templates folder, checks them, and writes the data files the portal runs on. Last month's files go to data/_previous first, which is the rollback, and a copy of this month's data files and templates goes into archive/YYYY-MM, which is the permanent record. Everyone else's portal picks the new files up on their next open, once OneDrive has synced."),
        h("div.steps", [
          stepEl(1, "Choose the portal folder", [
            U.note("The synced CTS Business Portal folder, the one this page was opened from. Edge asks you to confirm it; that is the one prompt."),
            h("div.row", [
              h("button.btn", { onclick: function () { chooseFolder(true); } }, "Choose portal folder"),
              h("span.muted", "Remembered after the first time; you will only be asked to allow it."),
            ]),
            folderStatus,
          ]),
          stepEl(2, "Read the templates", [
            U.note("Nothing is written at this step. The checks are what you would do by hand: does the ledger tie to the P&L, what is not in the chart, which locations have no department, what rests on a placeholder."),
            h("div.row", [h("button.btn", { onclick: readTemplates }, "Read templates")]),
            readStatus,
          ]),
          stepEl(3, "Write the data files", [
            U.note("Only after every check is green or amber. A red check blocks the write. The write also files a copy of the templates and data under archive, one folder per reporting month, so any month can be looked at again exactly as it was."),
            h("div.row", [h("button.btn", { onclick: writeFiles }, "Write data files")]),
            writeStatus,
          ]),
          stepEl(4, "Then the emails", [
            U.note("Once the portal has reloaded on the new month, go to Distribution to write this month's emails into the outbox. They are rendered from the data that is loaded, so this comes after the write, not before."),
            h("div.row", [h("a.btn.btn-quiet", { href: "#/distribution" }, "Go to Distribution")]),
          ]),
        ]),
        U.section("What the templates are", U.table([
          { key: "file", label: "Template", align: "left" },
          { key: "out", label: "Writes", align: "left" },
          { key: "need", label: "", align: "left",
            value: function (r) { return r.placeholder ? U.flag("warn", "placeholder shape") : r.need ? U.flag("good", "required") : h("span.muted", "optional"); } },
        ], B.TEMPLATES, { dense: true }),
          "Placeholder shape means the columns are what the portal needs until a real export from that system is matched; the paste may need reshaping until then."),
      ];
    },
  };

  /* ===================================================== Distribution == */
  P.distribution = {
    section: "Admin", title: "Distribution",
    sub: "Who gets the month and the fortnight, and how it goes out.",
    render: function () {
      var M = CTS.email, B = CTS.build;
      if (!M) return [U.h1("Distribution", "The email module did not load"),
                      U.note("CTS_email.js is missing from the portal folder.", "warn")];
      var p = CTS.period();
      var recipients = E.recipients();
      var month = E.reportingMonth();
      var canEdit = E.canPublishCommentary();
      var state = P.distribution._state || (P.distribution._state = { editing: null, fortnight: null });
      var previewBox = h("div"), outboxStatus = h("div"), pubStatus = h("div");
      var fnMonth = E.progressMonth();
      var fnX = E.fortnightFor ? E.fortnightFor(null, fnMonth, state.fortnight || undefined) : null;

      function showPreview(msg) {
        previewBox.innerHTML = "";
        previewBox.appendChild(h("div.toolbar", [
          h("strong", msg.subject), h("span.chip.chip-muted", M.TIERS[msg.tier]), msg.kind === "fortnightly" ? h("span.chip.chip-good", "fortnightly") : h("span.chip.chip-muted", "monthly"),
          h("div.toolbar-right", [
            h("a.btn.btn-quiet", { href: M.mailto(msg), target: "_blank" }, "Open in Outlook (text)"),
            h("button.btn.btn-quiet", { onclick: function () {
              try { navigator.clipboard.writeText(msg.html).then(function () { previewBox.insertBefore(h("div.banner.banner-good", "HTML copied. Paste into a new Outlook message."), previewBox.firstChild); }); }
              catch (e) { previewBox.insertBefore(h("div.banner.banner-warn", "Copy failed in this browser. Use the outbox file instead."), previewBox.firstChild); }
            } }, "Copy HTML"),
          ]),
        ]));
        var frame = h("iframe.preview", { sandbox: "", title: "Email preview" });
        previewBox.appendChild(frame);
        frame.srcdoc = msg.html;
        frame.scrollIntoView({ behavior: "smooth", block: "start" });
      }
      async function folder(status) {
        if (!B || !B.supported()) { status.appendChild(h("div.banner.banner-warn", "This needs Edge or Chrome, from the synced folder.")); return null; }
        var dir = await CTS.reuseDir();
        if (!dir) {
          try { var pk = await B.pickFolder(); if (!pk.ok) throw new Error("not the portal folder"); dir = pk.handle; await CTS.saveDir(dir); }
          catch (e) { if (e && e.name !== "AbortError") status.appendChild(h("div.banner.banner-warn", "Could not open the folder: " + (e.message || e))); return null; }
        }
        return dir;
      }
      async function writeOutbox(kind) {
        outboxStatus.innerHTML = "";
        var dir = await folder(outboxStatus); if (!dir) return;
        var msgs = kind === "fortnightly" ? M.renderAllFortnightly({ month: fnMonth, fortnight: state.fortnight || undefined }) : M.renderAll({ keys: p.keys, log: window.CTS_BUILD_LOG || null });
        var fold = kind === "fortnightly" ? fnMonth + "-F" + (fnX ? fnX.current.n : 1) : month;
        var noAddr = msgs.filter(function (m) { return !m.to; });
        try {
          var written = await B.writeOutbox(dir, month, msgs, fold);
          outboxStatus.appendChild(h("div.banner.banner-good", [
            h("strong", written.length + " messages written to outbox/" + fold + ". "),
            "Each is a JSON for the flow and an HTML you can open and copy into Outlook now." +
            (noAddr.length ? " " + noAddr.length + " of them have no address and will be skipped by the flow." : "")]));
        } catch (e) { outboxStatus.appendChild(h("div.banner.banner-warn", "Could not write the outbox: " + (e.message || e))); }
      }

      /* ---- the list, edited here ---- */
      var deptOpts = [{ code: "", short: "All" }].concat(E.depts.filter(function (d) { return d.code !== "UNALLOCATED"; }));
      function form(r, onDone) {
        r = r || { name: "", email: "", tier: 1, dept: "", send: true, cadence: "monthly", note: "" };
        var name = h("input.control", { value: r.name, placeholder: "Name", "aria-label": "Name" });
        var email = h("input.control", { value: r.email || "", placeholder: "name@company.com.au", type: "email", "aria-label": "Email" });
        var tier = h("select.control", [1, 2, 3].map(function (t) { return h("option", { value: t, selected: +r.tier === t }, t + " " + M.TIERS[t]); }));
        var dept = h("select.control", deptOpts.map(function (d) { return h("option", { value: d.code, selected: (r.dept || "") === d.code }, d.short); }));
        var cad = h("select.control", Object.keys(M.CADENCE).map(function (k) { return h("option", { value: k, selected: (r.cadence || "monthly") === k }, M.CADENCE[k]); }));
        var send = h("input", { type: "checkbox", checked: r.send !== false, id: "send-" + (r.name || "new").replace(/\W/g, "") });
        var note = h("input.control.grow", { value: r.note || "", placeholder: "Note", "aria-label": "Note" });
        return h("div.card", [
          h("div.row", [name, email, tier, dept, cad, h("label.checkbox", [send, " send"])]),
          h("div.row", [note]),
          h("div.row", [
            h("button.btn.small", { type: "button", onclick: function () {
              if (!name.value.trim()) { name.focus(); return; }
              onDone({ name: name.value.trim(), email: email.value.trim(), tier: +tier.value, dept: dept.value || null, send: send.checked, cadence: cad.value, note: note.value.trim() });
            } }, "Save"),
            h("button.btn.btn-quiet.small", { type: "button", onclick: function () { state.editing = null; CTS.router.reload(); } }, "Cancel"),
          ]),
        ]);
      }
      function save(list) { E.setRecipients(list); state.editing = null; CTS.router.reload(); }
      var listRows = recipients.map(function (r, i) {
        if (state.editing === i) return form(r, function (nr) { var l = recipients.slice(); l[i] = nr; save(l); });
        return null;
      });
      var table = U.table([
        { key: "name", label: "Name", align: "left" },
        { key: "email", label: "Email", align: "left", value: function (r) { return r.email || h("span.muted", "no address"); } },
        { key: "tier", label: "Tier", align: "left", value: function (r) { return h("span.chip.chip-muted", r.tier + " " + M.TIERS[r.tier]); } },
        { key: "dept", label: "Department", align: "left", value: function (r) { return r.dept ? (E.deptOf[r.dept] || {}).short || r.dept : "all"; } },
        { key: "cadence", label: "Cadence", align: "left", value: function (r) { return M.CADENCE[r.cadence || "monthly"]; } },
        { key: "send", label: "Send", align: "left", value: function (r) { return r.send === false ? U.flag("warn", "no") : U.flag("good", "yes"); } },
        { key: "pv", label: "", align: "left", value: function (r, i) {
          var idx = recipients.indexOf(r);
          return h("span.row", [
            h("button.btn.btn-quiet.small", { onclick: function () { showPreview(M.render(r, { keys: p.keys, log: window.CTS_BUILD_LOG || null })); } }, "Preview month"),
            /fortnight|both/.test(r.cadence || "") && E.fortnightFor ? h("button.btn.btn-quiet.small", { onclick: function () { showPreview(M.renderFortnightly(r, { month: fnMonth, fortnight: state.fortnight || undefined })); } }, "Preview fortnight") : null,
            canEdit ? h("button.btn.btn-quiet.small", { onclick: function () { state.editing = idx; CTS.router.reload(); } }, "Edit") : null,
            canEdit ? h("button.btn.btn-quiet.small", { onclick: function () { var l = recipients.slice(); l.splice(idx, 1); save(l); } }, "Remove") : null,
          ]);
        } },
      ], recipients);
      async function publishList() {
        pubStatus.innerHTML = "";
        var dir = await folder(pubStatus); if (!dir) return;
        try { await B.writeRecipients(dir, recipients); E.clearRecipientsLocal(); pubStatus.appendChild(h("div.banner.banner-good", "Recipients published to the folder. Reloading…")); setTimeout(function () { location.reload(); }, 1000); }
        catch (e) { pubStatus.appendChild(h("div.banner.banner-warn", "Could not publish: " + (e.message || e))); }
      }
      var source = E.recipientsLocal() ? h("span.chip.chip-warning", "edited here, not yet published") : (window.CTS_DISTRIBUTION ? h("span.chip.chip-good", "published list") : h("span.chip.chip-muted", "from 08 Config.xlsx"));

      var fnPick = fnX ? h("div.toolbar", [
        h("strong", "Fortnightly update"), h("span.chip.chip-muted", fnX.monthLabel + ", snapshot " + F.date(fnX.snapshot)),
        h("div.seg", [1, 2].map(function (n) { return h("button.seg-btn" + ((state.fortnight || fnX.current.n) === n ? ".on" : ""), { onclick: function () { state.fortnight = n; CTS.router.reload(); } }, "Fortnight " + n + (fnX.fortnights[n - 1].complete ? " (closed)" : fnX.fortnights[n - 1].started ? " (in progress)" : " (ahead)")); })),
        h("div.toolbar-right", [h("button.btn.btn-quiet", { onclick: function () { showPreview(M.renderFortnightly({ name: "Executive", tier: 1 }, { month: fnMonth, fortnight: state.fortnight || undefined })); } }, "Preview consolidated")]),
      ]) : null;

      return [
        CTS.seedBanner(), CTS.periodBar(),
        U.h1("Distribution", "The monthly email, " + p.label + ", and the fortnightly update"),
        U.note("Three tiers and two cadences per person. The portal renders the messages and writes them to the outbox folder; whatever sends them reads that folder. Nothing goes out until someone has clicked, which is deliberate."),
        U.section("Recipients", [
          h("div.row", [source, canEdit ? h("button.btn.small", { onclick: function () { state.editing = "new"; CTS.router.reload(); } }, "Add a recipient") : null,
                        canEdit && E.recipientsLocal() ? h("button.btn.btn-quiet.small", { onclick: publishList }, "Publish this list to the portal folder") : null,
                        canEdit && E.recipientsLocal() ? h("button.btn.btn-quiet.small", { onclick: function () { E.clearRecipientsLocal(); CTS.router.reload(); } }, "Discard my edits") : null]),
          pubStatus,
          state.editing === "new" ? form(null, function (nr) { save(recipients.concat([nr])); }) : null,
        ].concat(listRows.filter(Boolean)).concat([table]),
          canEdit ? "Edits are kept in this browser until you publish the list, which writes data/CTS_distribution_data.js into the folder and wins over the Config template's Distribution tab from then on." : "The finance head edits this list."),
        U.section("Preview", [previewBox]),
        U.section("Write the monthly emails", [
          U.note("Writes one JSON and one HTML per recipient into outbox/" + month + " in the portal folder. The JSON is what a Power Automate flow sends; the HTML is what you copy into Outlook until the flow is on. See docs/AUTOMATION.md."),
          h("div.row", [h("button.btn", { onclick: function () { writeOutbox("monthly"); } }, "Write monthly outbox for " + ((E.monthIdx[month] || {}).long || month))]),
        ]),
        U.section("Write the fortnightly updates", [
          U.note("Revenue by date off the ledger for the month in progress, hours off the last pay runs, the forecast for the month as the benchmark, and the confirmed work landing next. Department heads get their own department; tier one and finance get the consolidated version. Run Build with the latest GL paste first so the snapshot date is current."),
          fnPick,
          h("div.row", [h("button.btn", { onclick: function () { writeOutbox("fortnightly"); } }, fnX ? "Write fortnightly outbox for " + fnX.monthLabel + " F" + (state.fortnight || fnX.current.n) : "Fortnightly needs the engine")]),
          outboxStatus,
        ]),
        U.section("The month end deck", (function () {
          var D = CTS.deck, deckStatus = h("div");
          if (!D || !D.supported()) return [U.note("The deck module did not load (vendor/pptxgen.bundle.js and CTS_deck.js).", "warn")];
          async function sendDeck() {
            deckStatus.innerHTML = "";
            var dir = await folder(deckStatus); if (!dir) return;
            try {
              deckStatus.appendChild(h("p.note", "Building the deck\u2026"));
              var bytes = await D.bytes(), name = D.fileName();
              await B.writeBinary(dir, month, name, bytes);
              var msgs = D.messages();
              var written = await B.writeOutbox(dir, month, msgs, null, "deck_");
              deckStatus.innerHTML = "";
              deckStatus.appendChild(h("div.banner.banner-good", [h("strong", name + " written to outbox/" + month + " with " + written.length + " covering notes. "), "The flow sends each note with the deck attached; until the flow is on, open the outbox folder and attach it in Outlook."]));
            } catch (e) { deckStatus.innerHTML = ""; deckStatus.appendChild(h("div.banner.banner-warn", "Could not send the deck: " + (e && e.message || e))); }
          }
          return [
            U.note("The month end P&L pack as PowerPoint, in the CTS livery, built from the same engine as the pages: cover, headline, P&L, departments after the split, revenue, top clients, utilisation, the full year landing, cash, and the commentary from each page on its slide. Write and publish the commentary first; the deck carries what is there when it is built."),
            h("div.row", [
              h("button.btn", { type: "button", onclick: function () { try { D.download(); } catch (e) { deckStatus.appendChild(h("div.banner.banner-warn", "Could not build the deck: " + (e && e.message || e))); } } }, "Download " + D.fileName()),
              h("button.btn.btn-quiet", { type: "button", onclick: sendDeck }, "Send to the distribution list (tiers 1 and 3)"),
            ]),
            deckStatus,
          ];
        })()),
        U.section("The three tiers", U.table([
          { key: "t", label: "Tier", align: "left" }, { key: "who", label: "Who", align: "left" }, { key: "gets", label: "Gets", align: "left" },
        ], [
          { t: "1", who: "Executive", gets: "Monthly: company result, departments after the split, top clients, utilisation, commentary, what is worth a comment. Fortnightly: the consolidated update across departments" },
          { t: "2", who: "Department head", gets: "Monthly: their own department in depth, one line on the rest. Fortnightly: their department's update with team, key calls and forward view" },
          { t: "3", who: "Finance", gets: "Monthly: the controls, the company result, every comment. Fortnightly: the consolidated update" },
        ], { dense: true })),
      ];
    },
  };

  /* ========================================================= Pipeline == */
  P.pipeline = {
    section: "Revenue", title: "Pipeline",
    sub: "Deals, quotes and the order book.",
    render: function () {
      var p = CTS.period();
      if (!E.hasPipeline) {
        return [CTS.seedBanner(), CTS.periodBar(), U.h1("Pipeline", p.label),
          h("div.banner.banner-warn", [h("strong", "No pipeline loaded. "),
            "This page reads the Zoho Deals, OnRent Orders and Qwilr Quotes templates. Fill them and run Build."])];
      }
      var fyKeys = E.monthsOfFY(p.fy);
      var pl = E.pipelineFor(fyKeys, p.keys);
      var meta = (E.pipeline && E.pipeline.meta) || {};
      var scope = E.deptScope();
      var labels = fyKeys.map(function (k) { return E.monthIdx[k].label; });
      var bud = E.pnlBudget(fyKeys);

      var expected = fyKeys.map(function (k) { return pl.expectedByMonth[k] || 0; });
      var wonSeries = fyKeys.map(function (k) { return pl.wonByMonth[k] || 0; });
      var chart = U.columns({
        labels: labels, width: 780, height: 260,
        series: [
          { label: "Budget revenue", colour: "var(--measure-3)", values: bud.income },
          { label: "Won, by close month", colour: "var(--measure-2)", values: wonSeries },
          { label: "Open, weighted by probability", colour: "var(--measure-1)", values: expected },
        ],
      });

      var depts = E.visibleDepts().filter(function (d) { return d.isRevenue; });
      var bookSeries = depts.map(function (d) {
        return { label: d.short, colour: U.colourOf(d.code),
                 values: fyKeys.map(function (k) { return (pl.bookByMonthDept[k] || {})[d.code] || 0; }) };
      });
      var book = U.columns({ labels: labels, series: bookSeries, stacked: true, width: 780, height: 240 });

      var stageRows = pl.byStage.slice().sort(function (a, b) { return b.amount - a.amount; });
      var openRows = pl.open.filter(function (d) { return !scope || d.dept === scope; })
        .sort(function (a, b) { return (b.expected || 0) - (a.expected || 0); }).slice(0, 25);

      return [
        CTS.seedBanner(), CTS.periodBar(),
        U.h1("Pipeline", "FY" + p.fy),
        meta.placeholder ? h("div.banner.banner-warn", [h("strong", "Placeholder shapes. "),
          "Deals follow Zoho's standard export; orders and quotes follow the fields the portal needs. Send one real export from each system and the templates will be matched to them exactly."]) : null,
        h("div.tiles", [
          U.tile({ label: "Open pipeline", value: F.dollars(pl.openAmount), sub: pl.open.length + " deals" }),
          U.tile({ label: "Weighted", value: F.dollars(pl.openExpected), sub: "amount times probability" }),
          U.tile({ label: "Won this year", value: F.dollars(pl.wonAmount), sub: pl.won.length + " deals closed won" }),
          U.tile({ label: "Win rate", value: F.pct(pl.winRate, 0), sub: pl.won.length + " won, " + pl.lost.length + " lost",
                   tone: pl.winRate != null && pl.winRate < 0.3 ? "critical" : null }),
        ]),
        U.section("Expected revenue by close month against budget",
          U.figure("Won plus weighted open, by the month the deal closes", chart,
            U.table([{ key: "m", label: "Month", align: "left" }, { key: "b", label: "Budget", fmt: F.money },
                     { key: "w", label: "Won", fmt: F.money }, { key: "e", label: "Weighted open", fmt: F.money },
                     { key: "gap", label: "Gap", fmt: F.money, cell: U.moneyCell }],
              fyKeys.map(function (k, i) { return { m: labels[i], b: bud.income[i], w: wonSeries[i], e: expected[i], gap: wonSeries[i] + expected[i] - bud.income[i] }; }), { dense: true }),
            "The gap is what the pipeline does not yet cover. A department beating budget because its pipeline was understated shows up here as the pipeline being wrong, which is the thing the manual says to raise.")),
        U.section("Open deals by stage", U.table([
          { key: "stage", label: "Stage", align: "left" }, { key: "n", label: "Deals" },
          { key: "amount", label: "Amount", fmt: F.money }, { key: "expected", label: "Weighted", fmt: F.money },
        ], stageRows, { dense: true })),
        U.section("Order book, confirmed and completed by event month",
          U.figure("OnRent orders by department", book, null,
            "Production's forward view. This is the seasonality the monthly report reads, seen ahead rather than behind.")),
        U.section("Quote conversion, " + p.label, U.table([
          { key: "dept", label: "Department", align: "left", value: function (r) { return (E.deptOf[r.dept] || {}).short || r.dept; } },
          { key: "sent", label: "Quotes sent" }, { key: "sentValue", label: "Value sent", fmt: F.money },
          { key: "accepted", label: "Accepted" }, { key: "acceptedValue", label: "Value accepted", fmt: F.money },
          { key: "rate", label: "Conversion", fmt: function (v) { return F.pct(v, 0); },
            value: function (r) { return r.sent ? r.accepted / r.sent : null; } },
        ], pl.quotesByDept.filter(function (r) { return !scope || r.dept === scope; }), { dense: true })),
        U.section("Largest open deals", U.table([
          { key: "name", label: "Deal", align: "left" }, { key: "account", label: "Client", align: "left" },
          { key: "dept", label: "Dept", align: "left", value: function (r) { return (E.deptOf[r.dept] || {}).short || r.dept; } },
          { key: "stage", label: "Stage", align: "left" }, { key: "close", label: "Closes", align: "left", fmt: F.date },
          { key: "amount", label: "Amount", fmt: F.money },
          { key: "prob", label: "Prob", fmt: function (v) { return v == null ? "-" : F.pct(v, 0); } },
          { key: "expected", label: "Weighted", fmt: F.money }, { key: "owner", label: "Owner", align: "left" },
        ], openRows, { dense: true })),
      ];
    },
  };

  /* ================================================= Access Control ==== */
  P.access = {
    section: "Admin", title: "Access Control",
    sub: "Who sees what.",
    render: function () {
      if (!E.isAdmin()) {
        return [U.h1("Access Control", "Not available to you"),
                U.note("Only the finance head role can change who sees what.")];
      }
      var pages = E.users.PAGES || [];
      var roles = E.users.ROLES || [];
      var users = E.userList();
      var sections = [];
      pages.forEach(function (pg) {
        if (sections.indexOf(pg.section) < 0) sections.push(pg.section);
      });

      function toggle(roleId, pageId, on) {
        var cur = E.rolePages(roleId).slice();
        var i = cur.indexOf(pageId);
        if (on && i < 0) cur.push(pageId);
        if (!on && i >= 0) cur.splice(i, 1);
        E.setRolePages(roleId, cur);
        CTS.router.reload();
      }

      // the matrix: one row per page, one column per role
      var matrixRows = [];
      sections.forEach(function (sec) {
        matrixRows.push({ _cls: "totalrow", page: sec, _section: true });
        pages.filter(function (pg) { return pg.section === sec; }).forEach(function (pg) {
          matrixRows.push({ page: pg.title, id: pg.id, sensitive: pg.sensitive });
        });
      });

      var matrixCols = [
        { key: "page", label: "Tab", align: "left", width: "16rem",
          value: function (r) {
            if (r._section) return h("strong", r.page);
            return h("span", [r.page, " ",
              r.sensitive ? U.flag("warn", "sensitive",
                "Carries cost, pay or transaction level detail") : null]);
          } },
      ];
      roles.forEach(function (role) {
        matrixCols.push({ key: role.id, label: role.label, align: "left",
          value: function (r) {
            if (r._section) return h("span", "");
            var on = E.rolePages(role.id).indexOf(r.id) >= 0;
            if (role.admin) {
              return h("span.muted", { title: "The finance head role always sees everything" }, "always");
            }
            return h("input", { type: "checkbox", checked: on,
              onchange: function (e) { toggle(role.id, r.id, e.target.checked); } });
          } });
      });

      // people
      function setUser(i, field, value) {
        var list = E.userList().slice();
        list[i] = Object.assign({}, list[i]);
        list[i][field] = value;
        E.setUserList(list);
        CTS.router.reload();
      }
      var userRows = users.map(function (u, i) {
        var role = E.role(u.role);
        return {
          name: u.name, i: i, note: u.note || "",
          roleCell: h("select.control.small", {
            onchange: function (e) { setUser(i, "role", e.target.value); },
          }, roles.map(function (r) {
            return h("option", { value: r.id, selected: r.id === u.role }, r.label);
          })),
          deptCell: role && role.ownDeptOnly
            ? h("select.control.small", {
                onchange: function (e) { setUser(i, "dept", e.target.value || null); },
              }, [h("option", { value: "", selected: !u.dept }, "\u2014 not set \u2014")]
                .concat(E.postingDepts.map(function (d) {
                  return h("option", { value: d.code, selected: d.code === u.dept }, d.short);
                })))
            : h("span.muted", "all departments"),
          tabs: E.rolePages(u.role).length,
          remove: h("button.btn.btn-quiet.small", {
            onclick: function () {
              if (!confirm("Remove " + u.name + "?")) return;
              var list = E.userList().slice();
              list.splice(i, 1);
              E.setUserList(list);
              CTS.router.reload();
            },
          }, "Remove"),
        };
      });

      var newName = h("input.control", { placeholder: "Name" });
      var newRole = h("select.control", {}, roles.map(function (r) {
        return h("option", { value: r.id }, r.label);
      }));

      function usersText() {
        var out = {
          meta: { version: (E.users.meta.version || 1) + 1,
                  updated: new Date().toISOString().slice(0, 10),
                  note: "Visibility only. Not a security boundary." },
          PAGES: pages,
          ROLES: roles.map(function (r) {
            return Object.assign({}, r, { pages: E.rolePages(r.id) });
          }),
          USERS: E.userList(),
          DEFAULT_ROLE: E.users.DEFAULT_ROLE,
        };
        return "// CTS Business Intelligence Portal - people and access\n" +
          "// Exported " + out.meta.updated + " from the Access Control page.\n" +
          "// Visibility only. This does not protect anything: the portal is a file\n" +
          "// in a browser and every data file beside it is readable regardless.\n" +
          "window.CTS_USERS = " + JSON.stringify(out, null, 2) + ";\n";
      }

      var preview = E.previewRole();

      return [
        U.h1("Access Control", "Which tabs each role is shown"),
        h("div.banner.banner-warn", [
          h("strong", "Read this before you rely on it. "),
          "This decides what people are ", h("em", "shown"), ". It does not stop anyone seeing anything. The portal is a single file running in a browser with no server behind it, so every data file sitting beside it can be opened by anyone who can open the folder, whatever their role says here. It is worth setting up because showing a department head twenty two tabs they do not want is a real problem. It is not worth treating as confidentiality.",
        ]),
        U.section("If something genuinely must not be seen", [
          U.note("Then it must not be in the copy that person opens. Build them their own copy with the data left out: portal/build/build_data.py writes the data files, so a run that omits the ledger, or writes only one department's figures, produces a portal that physically cannot show the rest. That is the only version of this that holds. Ask for it and it can be set up."),
        ]),
        U.section("Preview", [
          U.note("Check what someone else is shown without signing in as them. The tabs on the left change to match while a preview is on."),
          h("div.row", [
            h("div.seg", [{ id: null, label: "Off" }].concat(roles).map(function (r) {
              var id = r.id || null;
              return h("button.seg-btn" + ((preview || null) === id ? ".on" : ""), {
                onclick: function () { E.setPreviewRole(id); location.hash = "#/access"; location.reload(); },
              }, r.label || "Off");
            })),
          ]),
          preview ? h("div.banner.banner-warn",
            "Previewing as " + (E.role(preview) || {}).label + ". Turn it off to get your own tabs back.") : null,
        ]),
        U.section("Which tabs each role sees", [
          U.table(matrixCols, matrixRows, { dense: true }),
          U.note("Sensitive marks the tabs carrying cost, pay or transaction level detail, so you can see at a glance what a tick hands out. Changes take effect immediately and are saved in this browser."),
        ]),
        U.section("Roles", U.table([
          { key: "label", label: "Role", align: "left" },
          { key: "tabs", label: "Tabs", value: function (r) { return E.rolePages(r.id).length; } },
          { key: "scope", label: "Departments", align: "left",
            value: function (r) { return r.ownDeptOnly ? "Own department only" : "All"; } },
          { key: "note", label: "Intended for", align: "left" },
        ], roles)),
        U.section("People", [
          U.table([
            { key: "name", label: "Name", align: "left" },
            { key: "roleCell", label: "Role", align: "left", value: function (r) { return r.roleCell; } },
            { key: "deptCell", label: "Department", align: "left", value: function (r) { return r.deptCell; } },
            { key: "tabs", label: "Tabs shown" },
            { key: "note", label: "Note", align: "left" },
            { key: "remove", label: "", align: "left", value: function (r) { return r.remove; } },
          ], userRows),
          h("div.row", [newName, newRole,
            h("button.btn", { onclick: function () {
              var n = newName.value.trim();
              if (!n) return;
              var list = E.userList().slice();
              list.push({ name: n, role: newRole.value, dept: null, note: "" });
              E.setUserList(list);
              CTS.router.reload();
            } }, "Add person"),
          ]),
          U.note("There is no password. People identify themselves by picking their name in the sidebar, which is an honour system, not a login. CTS has no single sign on wired to this portal."),
        ]),
        U.section("Keeping the changes", [
          U.note("Everything above is saved in this browser only, so it follows you rather than the portal. Export it and put the file in portal/data to make it what everyone gets."),
          keepOrShow({ name: "CTS_users_data.js", text: usersText }),
          h("div.row", [
            h("button.btn.btn-quiet", { onclick: function () {
              if (!confirm("Discard your changes to roles and people and go back to the file?")) return;
              CTS.store.set("accessRoles", {});
              CTS.store.set("accessUsers", null);
              CTS.router.reload();
            } }, "Reset to the file"),
          ]),
        ]),
      ];
    },
  };

  /* ========================================================== About ==== */
  P.about = {
    section: "Admin", title: "About",
    sub: "What this is, and what it is standing on.",
    render: function () {
      return [
        U.h1("CTS Business Intelligence Portal", "Version " + E.cfg.VERSION),
        U.section("What it is", [
          h("p", "One HTML file and a handful of data files. It runs entirely in the browser with no backend and no network calls, so it opens from a Teams folder, a shared drive or a local path. Every figure on screen is re-aggregated from the ledger each time the page loads, the same way the Controller Pack re-aggregates a GL paste."),
          h("p", "It is built around the way CTS actually reports: an Australian July to June year, six departments resolved from the cost centre on the transaction rather than the account name, the Admin overhead pool pushed out on three split bases, and the Controller Pack's own risk thresholds."),
        ]),
        U.section("What is sourced and what is not", U.table([
          { key: "f", label: "Figure", align: "left" },
          { key: "s", label: "State", align: "left" },
          { key: "n", label: "", align: "left" },
        ], [
          { f: "Chart of accounts", s: U.flag("good", "sourced"), n: "190 accounts read from Xero on 3 September 2026" },
          { f: "FY26 P&L control totals", s: U.flag("good", "sourced"), n: "Xero, accrual, pulled 3 September 2026. The portal ties to them to the cent." },
          { f: "August 2026 category totals", s: U.flag("good", "sourced"), n: "The Controller Pack's PL_Check worked example" },
          { f: "Department tags, risk thresholds, 3 Way split, sub splits", s: U.flag("good", "sourced"), n: "The Controller Pack and the budget bridge" },
          { f: "Margin norms and the production calendar", s: U.flag("good", "sourced"), n: "Monthly Reporting Part 3" },
          { f: "Staff and Office Dept split percentages", s: U.flag("warn", "placeholder"), n: "Live in the budget workbook, not in this repository" },
          { f: "Utilisation targets other than consulting", s: U.flag("warn", "assumed"), n: "Only one is written down" },
          { f: "Which basis each overhead line uses", s: U.flag("warn", "derived"), n: "Labelled per row on the live allocation tab" },
          { f: "Ledger detail, client names, hours", s: E.seed ? U.flag("warn", "seed") : U.flag("good", "loaded"),
            n: E.seed ? "Modelled to fit the real totals exactly. Client names are fictional." : "Loaded from a paste in this session" },
        ])),
        U.section("The conventions that will catch you out", [
          h("ul.plain", [
            h("li", "Amount is credit less debit. Income is positive, costs are negative, so gross profit is income plus cost of sales rather than minus."),
            h("li", "On the P&L, months up to the reporting month come from the ledger. Months after it come from budget, so the year always reads as twelve."),
            h("li", "The risk flag tests the materiality floor before the percentage, so a big percentage on a small dollar variance is still Low."),
            h("li", "Department comes from the transaction, not the account: three quarters of the P&L sits in accounts with no department in the name."),
            h("li", "Two utilisation figures exist and are not meant to agree. The pack nets holidays off and counts worked hours; the graphs file counts every logged hour against a plain weekday year, which is why that one lands on 2,080 or 2,088."),
          ]),
        ]),
        U.section("Rebuilding the data", [
          h("p.mono.small", "python3 portal/build/build_data.py"),
          U.note("Regenerates every file in portal/data from reporting/data/accounts.json and reporting/data/validation.json. It is deterministic: the same inputs give byte-identical output."),
        ]),
      ];
    },
  };

  /* ========================================================== Charts === */
  P.charts = {
    section: "Finance", title: "Trends",
    sub: "The pack that goes into the monthly report.",
    render: function () {
      var p = CTS.period();
      var fyKeys = E.monthsOfFY(p.fy);
      var labels = fyKeys.map(function (k) { return E.monthIdx[k].label; });
      var pnl = E.pnlBlend(fyKeys), bud = E.pnlBudget(fyKeys);
      var run = [], runB = [], a = 0, b = 0;
      pnl.netProfit.forEach(function (v, i) { a += v; b += bud.netProfit[i]; run.push(a); runB.push(b); });

      var rows = E.clientRows(p.keys, E.priorYearMonths(p.keys)).slice(0, 10);
      var depts = E.postingDepts;

      return [
        seedBanner(), CTS.periodBar(), U.h1("Charts", "FY" + p.fy),
        U.section("Revenue against profit",
          U.figure("Revenue, gross profit and net profit by month", U.columns({
            labels: labels, width: 780, height: 260,
            series: [
              { label: "Revenue", colour: "var(--measure-1)", values: pnl.income },
              { label: "Gross profit", colour: "var(--measure-2)", values: pnl.grossProfit },
              { label: "Net profit", colour: "var(--measure-3)", values: pnl.netProfit },
            ],
          }), null)),
        U.section("Budget against actual, as a running total",
          U.figure("Net profit, cumulative", U.lines({
            labels: labels, width: 780, height: 250, zero: true,
            series: [
              { label: "Actual, running total", colour: "var(--measure-1)", values: run },
              { label: "Budget, running total", colour: "var(--measure-3)", values: runB },
            ],
          }), null,
            "It has to be a running total. Enter the month on its own and the line goes up and down instead of climbing. It looks odd two months into the year and improves as the months build up.")),
        U.section("Utilisation month on month",
          U.figure("By department", U.lines({
            labels: labels, width: 780, height: 250,
            yfmt: function (v) { return F.pct(v, 0); }, tipfmt: function (v) { return F.pct(v); },
            series: depts.map(function (d) {
              return { label: d.short, colour: U.colourOf(d.code),
                       values: fyKeys.map(function (k) {
                         var r = (E.utilIdx[d.code] || {})[k];
                         if (!r) return null;
                         var w = r.chargeable + r.nonChargeable;
                         return w ? r.chargeable / w : null;
                       }) };
            }),
          }), null)),
        U.section("Top ten clients, year on year",
          U.figure("This period against the same period last year", U.bars({
            rows: rows.map(function (r) {
              return { label: r.display, value: r.amount, value2: r.prior,
                       colour: U.colourOf(r.primaryDept) };
            }), width: 620, label1: "This period", label2: "Last year",
          }), null)),
      ];
    },
  };
})();

/* ===================================================================== *
 * Shell: navigation, routing and boot.
 * ===================================================================== */
(function () {
  "use strict";
  var CTS = window.CTS, h = CTS.h, E = CTS.engine, U = CTS.ui;

  var SECTIONS = ["Dashboard", "Finance", "Budget", "Cash", "Revenue", "Utilisation", "Ledger", "Admin"];
  // one colour per group: brand hues on the dark sidebar, the validated
  // chart steps in the main column where they sit on white
  var NAV_COLOUR = { Dashboard: "#FFFFFF", Finance: "#6D71FF", Budget: "#E0A526", Cash: "#3FD0C9", Revenue: "#F33844", Utilisation: "#9DD6FF", Ledger: "#C9CED6", Admin: "#98A1AD" };
  var SEC_COLOUR = { Dashboard: "var(--accent)", Finance: "var(--series-4)", Budget: "var(--series-3)", Cash: "var(--series-2)", Revenue: "var(--series-1)", Utilisation: "var(--series-5)", Ledger: "var(--series-7)", Admin: "var(--ink-muted)" };
  var NO_COMMENTARY = { commentary: 1, setup: 1, loaders: 1, build: 1, distribution: 1, config: 1, access: 1, about: 1, context: 1 };

  /* ---- read aloud ----------------------------------------------------- */
  var reader = { on: false, btn: null };
  function pageNarrative() {
    var main = document.getElementById("main"), out = [];
    var sel = ".page-head h1, .page-head .page-sub, .banner, .commentary-text-view, .tile, section.block > .block-head h2, section.block .note, .commentary-empty";
    Array.prototype.forEach.call(main.querySelectorAll(sel), function (el) {
      if (el.closest && el.closest('[data-collapsed="1"] .block-body')) return;
      var t;
      if (el.classList.contains("tile")) {
        var l = el.querySelector(".tile-label"), v = el.querySelector(".tile-value"), s = el.querySelector(".tile-sub");
        t = [l && l.textContent, v && v.textContent, s && s.textContent].filter(Boolean).join(", ");
      } else t = el.textContent;
      t = String(t || "").replace(/\s+/g, " ").trim();
      if (t) out.push(t);
    });
    return out;
  }
  function stopReading() {
    reader.on = false;
    try { window.speechSynthesis && window.speechSynthesis.cancel(); } catch (e) {}
    if (reader.btn) { reader.btn.classList.remove("on"); reader.btn.setAttribute("aria-pressed", "false"); reader.btn.replaceChildren(U.icon("speaker"), "Read aloud"); }
  }
  function startReading(btn) {
    if (!window.speechSynthesis) { btn.replaceChildren("No speech in this browser"); return; }
    stopReading();
    reader.on = true; reader.btn = btn;
    btn.classList.add("on"); btn.setAttribute("aria-pressed", "true"); btn.replaceChildren(U.icon("stop"), "Stop reading");
    var parts = pageNarrative();
    var voices = window.speechSynthesis.getVoices ? window.speechSynthesis.getVoices() : [];
    var voice = voices.filter(function (v) { return /en[-_]AU/i.test(v.lang); })[0] || voices.filter(function (v) { return /^en/i.test(v.lang); })[0] || null;
    parts.forEach(function (t, i) {
      var u = new SpeechSynthesisUtterance(t);
      if (voice) u.voice = voice;
      u.lang = (voice && voice.lang) || "en-AU"; u.rate = 1;
      if (i === parts.length - 1) u.onend = stopReading;
      window.speechSynthesis.speak(u);
    });
  }

  /* ---- header tools: read aloud, theme, text size, contrast, sidebar ---- */
  var A11Y = (CTS.a11y = {
    theme: function () {
      var cur = document.documentElement.getAttribute("data-theme");
      var next = cur === "dark" ? "light" : cur === "light" ? "dark"
        : (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "light" : "dark");
      document.documentElement.setAttribute("data-theme", next);
      CTS.store.set("theme", next);
      if (CTS.repaintLogo) CTS.repaintLogo();
    },
    isDark: function () {
      var stamp = document.documentElement.getAttribute("data-theme");
      return stamp === "dark" || (!stamp && !!(window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches));
    },
    sizes: ["normal", "large", "larger"],
    size: function () { return CTS.store.get("textsize", "normal"); },
    cycleSize: function () {
      var cur = A11Y.size(), next = A11Y.sizes[(A11Y.sizes.indexOf(cur) + 1) % A11Y.sizes.length];
      CTS.store.set("textsize", next); A11Y.applySize();
    },
    applySize: function () {
      var v = A11Y.size();
      if (v === "normal") document.documentElement.removeAttribute("data-textsize"); else document.documentElement.setAttribute("data-textsize", v);
    },
    contrast: function () { return !!CTS.store.get("contrast", false); },
    toggleContrast: function () { CTS.store.set("contrast", !A11Y.contrast()); A11Y.applyContrast(); },
    applyContrast: function () {
      if (A11Y.contrast()) document.documentElement.setAttribute("data-contrast", "high"); else document.documentElement.removeAttribute("data-contrast");
    },
    sideHidden: function () { return !!CTS.store.get("sideHidden", false); },
    toggleSide: function () { CTS.store.set("sideHidden", !A11Y.sideHidden()); A11Y.applySide(); },
    applySide: function () {
      var shell = document.querySelector(".shell");
      if (shell) shell.setAttribute("data-side", A11Y.sideHidden() ? "hidden" : "shown");
      Array.prototype.forEach.call(document.querySelectorAll("[data-side-btn]"), function (b) {
        b.setAttribute("aria-pressed", A11Y.sideHidden() ? "true" : "false");
        b.replaceChildren(U.icon("menu"), A11Y.sideHidden() ? "Show menu" : "Hide menu");
      });
    },
    init: function () {
      var saved = CTS.store.get("theme", null);
      if (saved) document.documentElement.setAttribute("data-theme", saved);
      A11Y.applySize(); A11Y.applyContrast(); A11Y.applySide();
    },
  });

  /* ---- a department in focus: set from a chart or a tile, cleared from the header ---- */
  E.focusDept = function () { return E.deptScope() ? null : (CTS.store.get("focusDept", null) || null); };
  E.setFocusDept = function (code) { CTS.store.set("focusDept", code || null); };
  var baseVisible = E.visibleDepts;
  E.visibleDepts = function (list) {
    var out = baseVisible(list);
    var f = E.focusDept();
    if (!f) return out;
    var hit = out.filter(function (d) { return d.code === f; });
    return hit.length ? hit : out;
  };

  function pageTools(id) {
    var tools = [];
    var f = E.focusDept();
    if (f && E.deptOf[f]) {
      tools.push(h("span.filterchip", [
        h("span.dot", { style: { background: U.colourOf(f) } }), "Showing " + E.deptOf[f].short,
        h("button", { type: "button", "aria-label": "Show all departments", title: "Show all departments",
                      onclick: function () { E.setFocusDept(null); CTS.router.reload(); } }, "×"),
      ]));
    }
    var readBtn = h("button.tool", { type: "button", "aria-pressed": "false", title: "Read this page aloud" }, [U.icon("speaker"), "Read aloud"]);
    readBtn.addEventListener("click", function () { reader.on ? stopReading() : startReading(readBtn); });
    tools.push(readBtn);
    var anyOpen = function () { return Array.prototype.some.call(document.querySelectorAll("#main section.block"), function (s) { return s.getAttribute("data-collapsed") !== "1"; }); };
    var colBtn = h("button.tool", { type: "button", title: "Collapse or expand every section on this page" }, [U.icon("collapse"), "Collapse all"]);
    function paintCol() { var open = anyOpen(); colBtn.replaceChildren(U.icon(open ? "collapse" : "expand"), open ? "Collapse all" : "Expand all"); }
    colBtn.addEventListener("click", function () { U.setAllSections(anyOpen()); paintCol(); });
    setTimeout(paintCol, 0);
    tools.push(colBtn);
    var themeBtn = h("button.tool", { type: "button", title: "Switch between light and dark" }, [U.icon("theme"), A11Y.isDark() ? "Light" : "Dark"]);
    themeBtn.addEventListener("click", function () { A11Y.theme(); themeBtn.replaceChildren(U.icon("theme"), A11Y.isDark() ? "Light" : "Dark"); });
    tools.push(themeBtn);
    var sizeBtn = h("button.tool", { type: "button", title: "Text size: normal, large, larger" }, [U.icon("text"), "Text " + A11Y.size()]);
    sizeBtn.addEventListener("click", function () { A11Y.cycleSize(); sizeBtn.replaceChildren(U.icon("text"), "Text " + A11Y.size()); });
    tools.push(sizeBtn);
    var conBtn = h("button.tool" + (A11Y.contrast() ? ".on" : ""), { type: "button", "aria-pressed": A11Y.contrast() ? "true" : "false", title: "High contrast" }, [U.icon("contrast"), "Contrast"]);
    conBtn.addEventListener("click", function () { A11Y.toggleContrast(); conBtn.classList.toggle("on", A11Y.contrast()); conBtn.setAttribute("aria-pressed", A11Y.contrast() ? "true" : "false"); });
    tools.push(conBtn);
    var sideBtn = h("button.tool", { type: "button", "data-side-btn": "1", title: "Show or hide the menu (S)" }, [U.icon("menu"), "Hide menu"]);
    sideBtn.addEventListener("click", A11Y.toggleSide);
    tools.push(sideBtn);
    setTimeout(A11Y.applySide, 0);
    var helpBtn = h("button.tool", { type: "button", title: "Keyboard shortcuts (?)", "aria-label": "Keyboard shortcuts" }, "?");
    helpBtn.addEventListener("click", showHelp);
    tools.push(helpBtn);
    return tools;
  }

  /* ---- keyboard shortcuts ---------------------------------------------- */
  var SHORTCUTS = [
    ["[", "Previous reporting month"], ["]", "Next reporting month"],
    ["1 2 3 4", "Month, quarter, year to date, full year"],
    ["S", "Show or hide the menu"], ["R", "Read aloud, or stop"], ["C", "Collapse or expand every section"],
    ["D", "Light or dark"], ["G then a key", "Go: H dashboard, P profit and loss, F forecast, $ cash, U utilisation, M commentary"],
    ["Esc", "Clear the department filter, close this"], ["?", "This list"],
  ];
  var helpOpen = null, goMode = false;
  function showHelp() {
    if (helpOpen) { helpOpen.remove(); helpOpen = null; return; }
    var box = h("div.help", { role: "dialog", "aria-label": "Keyboard shortcuts" }, [
      h("div.help-head", [h("h2", "Keyboard shortcuts"), h("button.tool", { type: "button", onclick: showHelp }, "Close")]),
      h("dl", SHORTCUTS.map(function (s) { return [h("dt", h("kbd", s[0])), h("dd", s[1])]; }).reduce(function (a, b) { return a.concat(b); }, [])),
    ]);
    document.body.appendChild(box); helpOpen = box; box.querySelector("button").focus();
  }
  function shiftMonth(n) {
    var m = E.monthIdx[E.reportingMonth()]; if (!m) return;
    var next = E.months[m.i + n]; if (!next) return;
    E.setReportingMonth(next.key); CTS.router.reload();
  }
  document.addEventListener("keydown", function (ev) {
    var t = ev.target, tag = (t && t.tagName || "").toLowerCase();
    if (tag === "input" || tag === "textarea" || tag === "select" || (t && t.isContentEditable)) return;
    if (ev.ctrlKey || ev.metaKey || ev.altKey) return;
    var k = ev.key;
    if (goMode) {
      goMode = false;
      var map = { h: "home", p: "pnl", f: "forecast", "$": "cash", "4": "cash", u: "util", m: "commentary", r: "rev-summary", b: "bva" };
      if (map[k.toLowerCase()] && E.can(map[k.toLowerCase()])) CTS.router.go(map[k.toLowerCase()]);
      return;
    }
    if (k === "?") { ev.preventDefault(); showHelp(); }
    else if (k === "Escape") { if (helpOpen) showHelp(); else if (E.focusDept()) { E.setFocusDept(null); CTS.router.reload(); } }
    else if (k === "[") shiftMonth(-1);
    else if (k === "]") shiftMonth(1);
    else if (k === "1" || k === "2" || k === "3" || k === "4") { CTS.store.set("period", { "1": "month", "2": "qtr", "3": "ytd", "4": "fy" }[k]); CTS.router.reload(); }
    else if (k === "s" || k === "S") A11Y.toggleSide();
    else if (k === "d" || k === "D") { A11Y.theme(); CTS.router.reload(); }
    else if (k === "c" || k === "C") { var open = Array.prototype.some.call(document.querySelectorAll("#main section.block"), function (s) { return s.getAttribute("data-collapsed") !== "1"; }); U.setAllSections(open); CTS.router.reload(); }
    else if (k === "r" || k === "R") { var b = document.querySelector('.page-tools .tool[aria-pressed]'); if (b) b.click(); }
    else if (k === "g" || k === "G") goMode = true;
  });

  var router = (CTS.router = {
    current: null,
    go: function (id) { location.hash = "#/" + id; },
    reload: function () { router.render(router.current || "home"); },
    render: function (id) {
      stopReading();
      var page = CTS.pages[id];
      if (!page) { id = "home"; page = CTS.pages.home; }
      if (!E.can(id)) {
        // Land somewhere the person can actually see rather than an error.
        var first = Object.keys(CTS.pages).filter(function (k) { return E.can(k); })[0];
        if (id !== first && first) {
          router.current = first;
          return router.render(first);
        }
        page = { title: "Not available", section: "", render: function () {
          var r = E.currentRole();
          return [CTS.ui.h1("Not available to you",
            "This tab is not part of the " + ((r && r.label) || "current") + " role."),
            CTS.ui.note("Whoever holds the finance head role sets which tabs each role sees, on the Access Control page. Ask them if you need this one."),
          ];
        } };
      }
      router.current = id;
      var main = document.getElementById("main");
      main.innerHTML = "";
      main.style.setProperty("--sec", SEC_COLOUR[page.section] || "var(--accent)");
      var frag = document.createDocumentFragment();
      var nodes;
      try {
        nodes = page.render();
      } catch (err) {
        nodes = [h("div.banner.banner-warn", [
          h("strong", "This page could not be built. "), String(err && err.message || err)]),
          h("pre.small", String(err && err.stack || ""))];
      }
      (Array.isArray(nodes) ? nodes : [nodes]).forEach(function (n) {
        if (n) frag.appendChild(n);
      });
      main.appendChild(frag);
      var head = main.querySelector(".page-head");
      if (head) {
        var slot = head.querySelector("[data-tools]");
        if (slot) pageTools(id).forEach(function (b) { slot.appendChild(b); });
        if (!NO_COMMENTARY[id] && E.commentaryData) {
          try { head.parentNode.insertBefore(U.commentary(id), head.nextSibling); } catch (e) { console.warn("commentary", e); }
        }
      }
      main.scrollTop = 0;
      window.scrollTo(0, 0);
      Array.prototype.forEach.call(document.querySelectorAll(".nav-link"), function (a) {
        a.classList.toggle("on", a.getAttribute("data-page") === id);
        if (a.getAttribute("data-page") === id) a.setAttribute("aria-current", "page"); else a.removeAttribute("aria-current");
      });
      document.title = page.title + " · CTS Business Intelligence";
    },
  });

  function buildNav() {
    var nav = document.getElementById("nav");
    nav.innerHTML = "";
    var groups = [];
    function setAll(collapsed) {
      groups.forEach(function (g) { g.set(collapsed); });
    }
    nav.appendChild(h("div.nav-tools", [
      h("button", { type: "button", onclick: function () { setAll(true); } }, "Close groups"),
      h("button", { type: "button", onclick: function () { setAll(false); } }, "Open all"),
    ]));
    SECTIONS.forEach(function (sec) {
      var ids = Object.keys(CTS.pages).filter(function (id) {
        return CTS.pages[id].section === sec && E.can(id);
      });
      if (!ids.length) return;
      var key = "nav." + U.slug(sec), collapsed = !!CTS.store.get(key, false);
      var listId = "navlist-" + U.slug(sec);
      var group = h("div.nav-group", { "data-collapsed": collapsed ? "1" : "0", style: { "--sec": SEC_COLOUR[sec] || "var(--accent)" } });
      var head = h("button.nav-head", { type: "button", "aria-expanded": collapsed ? "false" : "true", "aria-controls": listId }, [
        h("span.sec-dot", { "aria-hidden": "true" }), h("span", sec), U.chev("chev"),
      ]);
      function set(c) { group.setAttribute("data-collapsed", c ? "1" : "0"); head.setAttribute("aria-expanded", c ? "false" : "true"); CTS.store.set(key, c); }
      head.addEventListener("click", function () { set(group.getAttribute("data-collapsed") !== "1"); });
      groups.push({ set: set });
      group.appendChild(head);
      group.appendChild(h("ul.nav-list", { id: listId }, ids.map(function (id) {
        return h("li", h("a.nav-link", {
          href: "#/" + id, "data-page": id, title: CTS.pages[id].sub || "",
        }, CTS.pages[id].title));
      })));
      nav.appendChild(group);
    });
  }

  /** Identification, not authentication: people pick their name. There is no
   *  password because there is nothing behind this to protect. */
  /** The logo travels inside the brand data file, so it cannot go missing
   *  when the folder is moved, and it swaps with the theme. */
  function buildLogo() {
    var img = document.getElementById("brandlogo");
    var b = window.CTS_BRAND;
    if (!img || !b || !b.logo) return;
    function pick() {
      var stamp = document.documentElement.getAttribute("data-theme");
      var dark = stamp === "dark" || (!stamp && window.matchMedia &&
                 window.matchMedia("(prefers-color-scheme: dark)").matches);
      img.src = dark ? b.logo.dark : b.logo.light;
      img.alt = b.logo.alt || "CTS";
      img.hidden = false;
    }
    pick();
    CTS.repaintLogo = pick;
    if (window.matchMedia) {
      try {
        window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", pick);
      } catch (e) {}
    }
  }

  function buildWhoAmI() {
    var slot = document.getElementById("whoami");
    if (!slot) return;
    slot.innerHTML = "";
    var users = E.userList();
    if (!users.length) return;
    var me = E.identity();
    var role = E.currentRole();
    var sel = h("select.control.small", {
      onchange: function (e) { E.setIdentity(e.target.value || null); location.reload(); },
    }, [h("option", { value: "", selected: !me }, "Not signed in")]
      .concat(users.map(function (u) {
        return h("option", { value: u.name, selected: u.name === me }, u.name);
      })));
    slot.appendChild(h("div.whoami-label", "Viewing as"));
    slot.appendChild(sel);
    if (role) {
      slot.appendChild(h("div.whoami-role", [
        role.label,
        E.previewRole() ? h("span.chip.chip-warning", "preview") : null,
        E.deptScope() ? h("span.chip.chip-muted",
          (E.deptOf[E.deptScope()] || {}).short || E.deptScope()) : null,
      ]));
    }
  }

  CTS.boot = function () {
    try {
      E.init();
    } catch (err) {
      document.getElementById("splash").innerHTML =
        '<div class="fatal"><h1>Data load error</h1><p>' +
        String(err && err.message || err) +
        '</p><p class="small">The portal could not build its index from the data files. ' +
        'Check that every file in portal/data loaded; the console lists any that did not.</p></div>';
      return;
    }
    buildNav();
    CTS.a11y.init();
    var badge = document.getElementById("databadge");
    if (badge) {
      badge.textContent = E.seed ? "seed data" : "loaded data";
      badge.className = "badge " + (E.seed ? "badge-warn" : "badge-good");
    }
    var rm = document.getElementById("rmlabel");
    if (rm) rm.textContent = (E.monthIdx[E.reportingMonth()] || {}).long || "";
    buildLogo();
    buildWhoAmI();

    window.addEventListener("hashchange", function () {
      router.render((location.hash || "#/home").replace(/^#\/?/, "") || "home");
    });
    router.render((location.hash || "#/home").replace(/^#\/?/, "") || "home");

    var splash = document.getElementById("splash");
    if (splash) {
      splash.style.transition = "opacity .35s"; splash.style.opacity = "0";
      setTimeout(function () { splash.remove(); }, 380);
    }
    document.body.classList.add("ready");
  };
})();
