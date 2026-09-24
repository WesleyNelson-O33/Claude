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
  E.setReportingMonth = function (k) { store.set("reportingMonth", k); };
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
    return E.hasActual(monthKey)
      ? (E.finActual[name] || {})[monthKey] || 0
      : (E.finBudget[name] || {})[monthKey] || 0;
  };
  E.blendSource = function (monthKey) { return E.hasActual(monthKey) ? "actual" : "budget"; };

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

  E.pnlActual = function (keys) {
    return E.buildPnl(keys, function (n, m) { return (E.finActual[n] || {})[m] || 0; });
  };
  E.pnlBudget = function (keys) {
    return E.buildPnl(keys, function (n, m) { return (E.finBudget[n] || {})[m] || 0; });
  };
  E.pnlBlend = function (keys) { return E.buildPnl(keys, E.acctBlend); };
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
  U.tile = function (opts) {
    return h("div.tile" + (opts.tone ? ".tone-" + opts.tone : ""), [
      h("div.tile-label", opts.label),
      h("div.tile-value", opts.value),
      opts.sub ? h("div.tile-sub", opts.sub) : null,
      opts.note ? h("div.tile-note", opts.note) : null,
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

  U.h1 = function (title, sub) {
    return h("header.page-head", [
      h("h1", title),
      sub ? h("p.page-sub", sub) : null,
    ]);
  };

  U.section = function (title, kids, note) {
    return h("section.block", [
      h("h2", title),
      note ? U.note(note) : null,
    ].concat(Array.isArray(kids) ? kids : [kids]));
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
    else if (id === "fy") { keys = E.monthsOfFY(fy); label = "FY" + fy + " full year, actual then budget"; }
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
      h("div.toolbar-right", [
        h("label.inline", "Reporting month"),
        monthSelect(),
      ].concat(extra || [])),
    ]);
  };
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
      "The FY26 totals, the July 2025 month and the August 2026 category totals are the real Xero figures and tie to the cent. Everything inside them is modelled from the seasonality and margin norms in the Monthly Reporting manual. Load the real Xero export on ",
      h("a", { href: "#/loaders" }, "Data Loaders"), " to replace it.",
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

      var tiles = h("div.tiles", [
        U.tile({ label: "Revenue", value: F.dollars(t.income), sub: delta(t.income, pt.income),
                 note: "Budget " + F.dollars(bt.income) }),
        U.tile({ label: "Gross profit", value: F.dollars(t.grossProfit),
                 sub: F.pct(t.income ? t.grossProfit / t.income : null) + " margin",
                 note: "Budget " + F.pct(bt.income ? bt.grossProfit / bt.income : null) }),
        U.tile({ label: "Net profit", value: F.dollars(t.netProfit),
                 tone: t.netProfit >= 0 ? "good" : "critical",
                 sub: F.pct(t.income ? t.netProfit / t.income : null) + " of revenue",
                 note: "Budget " + F.dollars(bt.netProfit) }),
        U.tile({ label: "Utilisation", value: F.pct(util.total.util),
                 sub: util.total.fte ? util.total.fte.toFixed(1) + " full time equivalents" : null,
                 note: "Chargeable over worked hours, leave excluded" }),
      ]);

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
      });
      var trendTable = U.table([
        { key: "m", label: "Month", align: "left" },
        { key: "src", label: "", align: "left" },
        { key: "rev", label: "Revenue", fmt: F.money },
        { key: "gp", label: "Gross profit", fmt: F.money },
        { key: "np", label: "Net profit", fmt: F.money },
      ], fyKeys.map(function (k, i) {
        return { m: E.monthIdx[k].label,
                 src: E.blendSource(k) === "actual" ? "" : "budget",
                 rev: full.income[i], gp: full.grossProfit[i], np: full.netProfit[i],
                 _cls: E.blendSource(k) === "actual" ? "" : "forecast" };
      }), { dense: true });

      // departmental result after allocation
      var alloc = E.allocate(p.keys);
      var deptRows = alloc.rows.filter(function (r) {
        return r.dept.isRevenue && (r.income || r.expenses);
      });
      var deptChart = U.columns({
        labels: deptRows.map(function (r) { return r.dept.short; }),
        width: 520, height: 250,
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
          U.figure("FY" + p.fy + ", actual to " + (E.monthIdx[p.rm] || {}).label + " then budget",
                   trend, trendTable,
                   "Months up to the reporting month come off the ledger. Months after it come from budget, so the year always reads as twelve and the shortfall ahead is visible. That is the Controller Pack's rule.")),
        allocBanner(),
        U.section("Departments, after the overhead split",
          U.figure("Result by department, " + p.label, deptChart, deptTable,
                   "Net profit here is after the Admin overhead pool has been pushed out. The norm column is the margin Monthly Reporting Part 3 says to expect."))
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
        U.note("Costs are shown the way the ledger holds them: credit less debit, so a cost is negative and gross profit is income plus cost of sales. Columns after " + (E.monthIdx[p.rm] || {}).label + " are budget, not actual, and are shaded."),
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

      return [
        seedBanner(), CTS.periodBar(), allocBanner(),
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
                 src: E.blendSource(k) === "actual" ? "actual" : "budget",
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
                   return E.sumMonths((E.idx.deptCatMonth[d.code] || {}).income, [k]);
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
          U.figure("Stacked by department, actual then budget", stacked,
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
  P.loaders = {
    section: "Admin", title: "Data Loaders",
    sub: "Paste straight out of Xero. No add-in, no upload.",
    render: function () {
      var out = h("div.loadout");

      function loader(opts) {
        var ta = h("textarea.paste", { rows: 6, placeholder: opts.placeholder });
        var btn = h("button.btn", { onclick: function () {
          out.innerHTML = "";
          try { opts.run(parseTable(ta.value), out); }
          catch (err) { out.appendChild(h("div.banner.banner-warn", String(err && err.message || err))); }
        } }, opts.action);
        return h("div.card.loader", [
          h("h3", opts.title),
          U.note(opts.note),
          opts.columns ? h("p.mono.small", "Expected columns: " + opts.columns) : null,
          ta, h("div.row", [btn]),
        ]);
      }

      function ok(node, msg) { node.appendChild(h("div.banner.banner-good", msg)); }

      return [
        U.h1("Data Loaders", "Copy the cells in Excel, click in the box and paste."),
        U.note("Everything here is parsed in the browser. Nothing is uploaded anywhere. Copying a range out of Excel puts it on the clipboard tab separated, which is what these boxes read; a comma separated file works too."),
        h("div.cards", [
          loader({
            title: "Ledger, from Xero",
            note: "Run the saved custom report Account transactions for P&L analysis, year to date from 1 July, sorted by account code. Convert the values to number before you copy, or the totals will not add up. Re-paste every prior month, not just the new one: earlier months do move.",
            columns: "Account Name, Date, Contact, Debit, Credit, Job Numbers, Cost Centres. Others are ignored.",
            placeholder: "Paste the report here, including its header row",
            action: "Load ledger",
            run: function (t, node) {
              var ci = {};
              t.head.forEach(function (c, i) { ci[c.toLowerCase().replace(/[^a-z]/g, "")] = i; });
              function col() {
                for (var i = 0; i < arguments.length; i++) {
                  if (ci[arguments[i]] != null) return ci[arguments[i]];
                }
                return -1;
              }
              var cAcc = col("accountname", "account"), cDate = col("date"),
                  cContact = col("contact"), cDebit = col("debit"), cCredit = col("credit"),
                  cJob = col("jobnumbers", "jobnumber", "job"),
                  cCC = col("costcentres", "costcentre", "department");
              if (cAcc < 0 || cDate < 0) throw new Error("Could not find an Account Name and a Date column in the header row.");
              var rows = [], bad = 0;
              t.rows.forEach(function (r) {
                var d = isoDate(r[cDate]);
                if (!d || !r[cAcc]) { bad++; return; }
                rows.push([d, String(r[cAcc]).trim(),
                           cContact >= 0 ? String(r[cContact] || "").trim() : "",
                           num(r[cDebit]), num(r[cCredit]),
                           "", cJob >= 0 ? String(r[cJob] || "") : "", "",
                           cCC >= 0 ? String(r[cCC] || "") : "", ""]);
              });
              if (!rows.length) throw new Error("No usable rows found.");
              var payload = { meta: { seed: false, built: new Date().toISOString().slice(0, 10),
                                      rows: rows.length, note: "Loaded from a Xero paste." },
                              cols: ["date", "account", "contact", "debit", "credit",
                                     "source", "jobNo", "invoiceNo", "costCentre", "description"],
                              rows: rows };
              E.loadGL(payload);
              E._deptShare = null;
              E.seed = false;
              ok(node, rows.length + " lines loaded" + (bad ? ", " + bad + " skipped for a missing date or account" : "") + ". Every page is now reading them.");
              node.appendChild(U.table([
                { key: "l", label: "", align: "left" }, { key: "v", label: "" },
              ], [
                { l: "Lines with no department", v: F.num(E.quality.noDept) },
                { l: "Lines on an account not in the chart", v: F.num(E.quality.unknownAcct) },
                { l: "Revenue lines with no contact", v: F.num(E.quality.noContact) },
                { l: "Date range", v: F.date(E.quality.minDate) + " to " + F.date(E.quality.maxDate) },
              ]));
              node.appendChild(keepOrShow({
                name: "CTS_gl_data.js",
                keep: [{ kind: "gl", payload: function () { return payload; } }],
                text: function () {
                  return "// Loaded " + payload.meta.built + " from a Xero paste.\n" +
                         "window.CTS_GL = " + JSON.stringify(payload) + ";\n";
                },
              }));
            },
          }),
          loader({
            title: "Profit and loss, from Xero",
            note: "Run the profit and loss report for the year to date and copy it with the account names down the side and the months across the top. This is what the P&L Control page checks the ledger against.",
            columns: "First column the account name, then one column per month headed Jul-25, Aug-25 and so on.",
            placeholder: "Account\tJul-26\tAug-26\nContract Support Staff\t120,450\t131,200",
            action: "Load P&L",
            run: function (t, node) {
              var months = t.head.slice(1).map(function (lbl) {
                var m = E.months.filter(function (x) {
                  return x.label.toLowerCase() === lbl.toLowerCase().trim() ||
                         x.long.toLowerCase() === lbl.toLowerCase().trim() ||
                         x.key === lbl.trim();
                })[0];
                return m ? m.key : null;
              });
              if (!months.filter(Boolean).length) throw new Error("No month columns recognised. Head them Jul-26, Aug-26 or 2026-07.");
              var actual = {}, unknown = [];
              t.rows.forEach(function (r) {
                var name = String(r[0] || "").trim();
                if (!name) return;
                if (!E.acct[name]) { unknown.push(name); return; }
                months.forEach(function (mk, i) {
                  if (!mk) return;
                  var v = Math.round(num(r[i + 1]) * 100);
                  if (v) (actual[name] = actual[name] || {})[mk] = v;
                });
              });
              E.fin.actual = actual;
              E.loadFin(E.fin);
              ok(node, Object.keys(actual).length + " accounts loaded across " +
                 months.filter(Boolean).length + " months.");
              if (unknown.length) {
                node.appendChild(h("div.banner.banner-warn", [
                  h("strong", unknown.length + " account names were not recognised and were skipped. "),
                  "They have to match the Xero chart exactly: " + unknown.slice(0, 6).join(", ") +
                  (unknown.length > 6 ? " and others." : ""),
                ]));
              }
              var finPayload = { meta: { seed: false, basis: "accrual", currency: "AUD" },
                                 actual: actual, budget: E.finBudget,
                                 months: E.months.map(function (m) { return m.key; }) };
              node.appendChild(keepOrShow({
                name: "CTS_fin_data.js",
                keep: [{ kind: "fin", payload: function () { return finPayload; } }],
                text: function () {
                  return "// Loaded " + new Date().toISOString().slice(0, 10) + " from a Xero paste.\n" +
                         "window.CTS_FIN = " + JSON.stringify(finPayload) + ";\n";
                },
              }));
              node.appendChild(U.note("Keeping it also keeps the budget already loaded.", "muted"));
            },
          }),
          buildEarningsLoader(),
        ]),
        out,
        U.section("Rolling back", [
          U.note("A load lives in this browser tab only. Reload the page and the portal goes back to whatever is in portal/data. Nothing you paste here is written to disk unless you click a save button, and nothing leaves the machine."),
        ]),
      ];
    },
  };

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
            "This page reads the Employment Hero earnings export, which is the only source that carries who worked, on what, and what they were paid. Load it on ",
            h("a", { href: "#/loaders" }, "Data Loaders"), ".",
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
    section: "Ledger", title: "Charts",
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

  var SECTIONS = ["Dashboard", "Finance", "Budget", "Revenue", "Utilisation", "Ledger", "Admin"];

  var router = (CTS.router = {
    current: null,
    go: function (id) { location.hash = "#/" + id; },
    reload: function () { router.render(router.current || "home"); },
    render: function (id) {
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
      main.scrollTop = 0;
      window.scrollTo(0, 0);
      Array.prototype.forEach.call(document.querySelectorAll(".nav-link"), function (a) {
        a.classList.toggle("on", a.getAttribute("data-page") === id);
      });
      document.title = page.title + " · CTS Business Intelligence";
    },
  });

  function buildNav() {
    var nav = document.getElementById("nav");
    nav.innerHTML = "";
    SECTIONS.forEach(function (sec) {
      var ids = Object.keys(CTS.pages).filter(function (id) {
        return CTS.pages[id].section === sec && E.can(id);
      });
      if (!ids.length) return;
      nav.appendChild(h("div.nav-group", [
        h("div.nav-head", sec),
        h("ul.nav-list", ids.map(function (id) {
          return h("li", h("a.nav-link", {
            href: "#/" + id, "data-page": id, title: CTS.pages[id].sub || "",
          }, CTS.pages[id].title));
        })),
      ]));
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

  function themeToggle() {
    var btn = document.getElementById("theme");
    if (!btn) return;
    var saved = CTS.store.get("theme", null);
    if (saved) document.documentElement.setAttribute("data-theme", saved);
    btn.addEventListener("click", function () {
      var cur = document.documentElement.getAttribute("data-theme");
      var next = cur === "dark" ? "light" : cur === "light" ? "dark"
        : (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "light" : "dark");
      document.documentElement.setAttribute("data-theme", next);
      CTS.store.set("theme", next);
      if (CTS.repaintLogo) CTS.repaintLogo();
    });
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
    themeToggle();
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
    if (splash) splash.remove();
    document.body.classList.add("ready");
  };
})();
