/* FGS Recipe Bundle Simulator - UI.  All logic lives in engine.js; this file only reads files and draws. */
(function () {
  "use strict";
  var E = window.FGSEngine;

  /* ---------- tiny DOM helper ---------- */
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === "class") el.className = v;
      else if (k === "text") el.textContent = v;
      else if (k === "on") Object.keys(v).forEach(function (ev) { el.addEventListener(ev, v[ev]); });
      else if (k === "style") el.setAttribute("style", v);
      else if (v === true) el.setAttribute(k, "");
      else el.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(el, arguments[i]);
    return el;
  }
  function add(el, c) {
    if (c === null || c === undefined || c === false) return;
    if (Array.isArray(c)) c.forEach(function (x) { add(el, x); });
    else el.appendChild(typeof c === "object" ? c : document.createTextNode(String(c)));
  }
  function $(s, r) { return (r || document).querySelector(s); }
  function clear(el) { while (el.firstChild) el.removeChild(el.firstChild); return el; }
  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
  function pLabel(k) { var p = E.PROTEINS.filter(function (x) { return x.key === k; })[0]; return p ? p.label : k; }

  /* ---------- state ---------- */
  var CFG = null, PROJ = {}, RAW = {};      /* PROJ[id] = built project; RAW[id] = repo version (for "reset") */
  var S = {
    proj: "fgss", tab: "sim", mode: "target", seed: 7,
    sel: { proteins: [], diet: "none", restrict: [], allergies: [], cook: "" },
    manual: null, manualSig: "", allShowEffect: false, allFilter: "all", allSearch: "", allProtein: ""
  };

  function cfgOf(id) { return CFG.projects.filter(function (p) { return p.id === id; })[0]; }
  function cur() { return PROJ[S.proj]; }

  /* ---------- loading ---------- */
  function fetchOK(url, kind) {
    return fetch(url, { cache: "no-cache" }).then(function (r) {
      if (!r.ok) throw new Error(url + " -> HTTP " + r.status);
      return kind === "json" ? r.json() : r.arrayBuffer();
    });
  }
  function sheetRows(wb, name) { return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null, raw: true }); }
  function looksLikeRecipeSheet(rows) {
    var tags = false, bundle = false, i, j;
    for (i = 0; i < Math.min(rows.length, 8); i++) for (j = 0; j < (rows[i] || []).length; j++)
      if (/^\s*tags?\s*$/i.test(String(rows[i][j] == null ? "" : rows[i][j]))) tags = true;
    for (i = 0; i < Math.min(rows.length, 12); i++) if (/^Bundle\s*\d+/i.test(String((rows[i] || [])[0] == null ? "" : rows[i][0]))) bundle = true;
    return tags && bundle;
  }
  function findSheet(wb, wanted) {
    var names = wb.SheetNames, i;
    if (wanted) for (i = 0; i < names.length; i++) if (names[i].toLowerCase() === wanted.toLowerCase()) return names[i];
    var ok = names.filter(function (n) { return looksLikeRecipeSheet(sheetRows(wb, n)); });
    var pref = ok.filter(function (n) { return /canva|final/i.test(n); });
    return (pref[0] || ok[0] || names[0]);
  }
  function labelMapFromWorkbook(wb, labelSheet) {
    var L = {}, names = labelSheet && wb.Sheets[labelSheet] ? [labelSheet] : wb.SheetNames;
    names.forEach(function (n) {
      sheetRows(wb, n).forEach(function (r) {
        var m = /^Bundle\s*(\d+)\s*[\(\-:]\s*(.+?)\)?\s*$/i.exec(String((r && r[0]) == null ? "" : r[0]).trim());
        if (m && !L[parseInt(m[1], 10)]) L[parseInt(m[1], 10)] = m[2].trim();
      });
    });
    return L;
  }
  function build(cfg, wb, patches, src) {
    var sheet = findSheet(wb, cfg.sheet);
    var labels = Object.assign({}, cfg.labels || {}, labelMapFromWorkbook(wb, cfg.labelSheet));
    var bundles = E.parseRows(sheetRows(wb, sheet), { labels: labels });
    var used = E.applyPatches(bundles, patches);
    return { cfg: cfg, bundles: bundles, patches: patches, used: used, src: src, sheet: sheet, sheetsInFile: wb.SheetNames.slice(), wb: wb };
  }
  function loadProject(cfg) {
    return Promise.all([fetchOK(cfg.xlsx, "buf"), cfg.patches ? fetchOK(cfg.patches, "json").catch(function () { return []; }) : Promise.resolve([])])
      .then(function (r) {
        var wb = XLSX.read(r[0], { type: "array" });
        var p = build(cfg, wb, r[1], { kind: "repo", name: cfg.xlsx });
        RAW[cfg.id] = p; PROJ[cfg.id] = p;
      });
  }
  function loadUpload(file) {
    var fr = new FileReader();
    fr.onload = function () {
      try {
        var wb = XLSX.read(fr.result, { type: "array" });
        var p = cur();
        PROJ[S.proj] = build(p.cfg, wb, p.patches, { kind: "upload", name: file.name });
        S.manual = null; renderAll();
      } catch (e) { alert("Could not read that file: " + e.message); }
    };
    fr.readAsArrayBuffer(file);
  }

  /* ---------- URL hash ---------- */
  function writeHash() {
    var s = S.sel, q = [];
    q.push("p=" + S.proj);
    if (S.tab !== "sim") q.push("t=" + S.tab);
    if (s.proteins.length) q.push("pr=" + s.proteins.join(","));
    if (s.diet !== "none") q.push("d=" + s.diet);
    if (s.restrict.length) q.push("r=" + s.restrict.join(","));
    if (s.allergies.length) q.push("a=" + encodeURIComponent(s.allergies.join(",")));
    if (s.cook) q.push("c=" + s.cook);
    if (S.mode !== "target") q.push("m=" + S.mode);
    q.push("s=" + S.seed);
    try { history.replaceState(null, "", "#" + q.join("&")); } catch (e) {}
  }
  function readHash() {
    var o = {};
    location.hash.replace(/^#/, "").split("&").forEach(function (kv) { var i = kv.indexOf("="); if (i > 0) o[kv.slice(0, i)] = decodeURIComponent(kv.slice(i + 1)); });
    function list(x) { return x ? x.split(",").filter(Boolean) : []; }
    if (o.p) S.proj = o.p;
    if (o.t) S.tab = o.t;
    if (o.pr) S.sel.proteins = list(o.pr);
    if (o.d) S.sel.diet = o.d;
    if (o.r) S.sel.restrict = list(o.r);
    if (o.a) S.sel.allergies = list(o.a);
    if (o.c) S.sel.cook = o.c;
    if (o.m) S.mode = o.m;
    if (o.s) S.seed = parseInt(o.s, 10) || 7;
  }

  /* ---------- simulation plumbing ---------- */
  function rng(seed) { var a = seed >>> 0; return function () { a |= 0; a = (a + 0x6D2B79F5) | 0; var t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }
  function jitterFor(bundles, seed) { var r = rng(seed), j = {}; bundles.slice().sort(function (a, b) { return a.id - b.id; }).forEach(function (b) { j[b.id] = r(); }); return j; }
  function opts() { var p = cur(); return { mode: (p.cfg.legacy ? S.mode : "target"), mainCount: p.cfg.mainCount }; }
  function sanitizeSel() {
    var meat = S.sel.proteins.some(function (p) { return p !== "vegetarian"; });
    if (meat && (S.sel.diet === "vegetarian" || S.sel.diet === "vegan")) S.sel.diet = "none";
    return meat;
  }
  function run() {
    var p = cur(); sanitizeSel();
    var r = E.simulate(p.bundles, S.sel, jitterFor(p.bundles, S.seed), opts());
    var sig = JSON.stringify([S.sel, S.mode, S.seed, S.proj, p.src.name]);
    if (sig !== S.manualSig) { S.manual = null; S.manualSig = sig; }
    if (S.manual) {
      var byId = {}; r.ranked.forEach(function (x) { byId[x.key.id] = x; });
      var order = S.manual.filter(function (id) { return byId[id]; });
      if (order.length === r.ranked.length) {
        r.ranked = order.map(function (id, i) { var x = byId[id]; x.rank = i + 1; return x; });
        r.fields = []; for (var i = 0; i < 6; i++) r.fields.push(r.ranked[i] ? r.ranked[i].saved : "");
      }
    }
    return r;
  }

  /* ====================================================================== */
  /*  Header (project tabs + data source bar)                               */
  /* ====================================================================== */
  function renderHeader() {
    var pt = clear($("#projTabs"));
    CFG.projects.forEach(function (c) {
      pt.appendChild(h("button", { role: "tab", class: c.id === S.proj ? "on" : "", text: c.label, title: c.long,
        on: { click: function () { S.proj = c.id; S.manual = null; if (!cur().cfg.legacy) S.mode = "target"; renderAll(); } } }));
    });
    var p = cur(), nRec = p.bundles.reduce(function (n, b) { return n + b.recipes.length; }, 0);
    var sb = clear($("#sourceBar"));
    var up = p.src.kind !== "repo";
    sb.appendChild(h("span", { class: "pill" + (up ? " warn" : ""), text: p.src.kind === "repo" ? "Repo file" : p.src.kind === "upload" ? "Uploaded file (not saved)" : "Data from a REDCap script" }));
    sb.appendChild(h("span", { text: p.src.name + (p.src.kind !== "script" ? "  ·  sheet \"" + p.sheet + "\"" : "") + "  ·  " + plural(p.bundles.length, "bundle") + ", " + plural(nRec, "recipe slot") + "  ·  " + p.cfg.mainCount + " main meals per bundle" }));
    sb.appendChild(h("button", { text: "Try a different .xlsx…", title: "Preview an updated spreadsheet without committing it", on: { click: function () { $("#xlsxPick").click(); } } }));
    if (up) sb.appendChild(h("button", { text: "Back to repo file", on: { click: function () { PROJ[S.proj] = RAW[S.proj]; S.manual = null; renderAll(); } } }));
    Array.prototype.forEach.call(document.querySelectorAll("#tabs button"), function (b) { b.classList.toggle("on", b.getAttribute("data-tab") === S.tab); });
  }

  /* ====================================================================== */
  /*  TAB 1 - participant simulator                                          */
  /* ====================================================================== */
  var PRESETS = [
    { t: "Reset", s: { proteins: [], diet: "none", restrict: [], allergies: [], cook: "" } },
    { t: "No preferences", s: { proteins: [], diet: "none", restrict: [], allergies: [], cook: "c" } },
    { t: "Chicken + turkey + beef", s: { proteins: ["chicken", "turkey", "beef"], diet: "none", restrict: [], allergies: [], cook: "" } },
    { t: "Vegan, 30 min or less", s: { proteins: [], diet: "vegan", restrict: [], allergies: [], cook: "a" } },
    { t: "Vegetarian + tofu", s: { proteins: ["vegetarian"], diet: "vegetarian", restrict: [], allergies: [], cook: "" } },
    { t: "Pescatarian", s: { proteins: [], diet: "pescatarian", restrict: [], allergies: [], cook: "" } },
    { t: "Halal + kosher", s: { proteins: [], diet: "none", restrict: ["halal", "kosher"], allergies: [], cook: "" } },
    { t: "Gluten-free + milk allergy", s: { proteins: [], diet: "none", restrict: ["gluten_free"], allergies: ["Milk"], cook: "" } },
    { t: "Many allergies", s: { proteins: ["chicken"], diet: "none", restrict: [], allergies: ["Milk", "Eggs", "Wheat", "Soy"], cook: "" } }
  ];
  var COOK_UI = [{ k: "", t: "(not answered)" }, { k: "a", t: "30 minutes or less" }, { k: "b", t: "45 minutes to 1 hour" }, { k: "c", t: "More than 1 hour" }];
  var PROT_UI = { chicken: "Chicken", turkey: "Turkey", beef: "Beef", fish_shellfish: "Fish or shellfish", vegetarian: "Tofu or other plant-based protein" };
  var DIET_UI = { none: "No specific preference (we eat meat, fish, and plant-based meals)", vegetarian: "Vegetarian", vegan: "Vegan", pescatarian: "Pescatarian (fish but no other meat)" };
  var REST_UI = { halal: "Halal", kosher: "Kosher", gluten_free: "Gluten-free" };

  function toggle(arr, v, on) { var i = arr.indexOf(v); if (on && i < 0) arr.push(v); if (!on && i > -1) arr.splice(i, 1); }

  function optRow(type, name, value, label, checked, disabled, onChange) {
    var inp = h("input", { type: type, name: name, value: value });
    inp.checked = !!checked; inp.disabled = !!disabled;
    inp.addEventListener("change", onChange);
    return h("label", { class: "opt" + (disabled ? " off" : "") }, inp, h("span", { text: label }));
  }

  function renderSim() {
    var root = clear($("#tab-sim")), p = cur(), meat = sanitizeSel();
    var left = h("div", { class: "card" });
    left.appendChild(h("h3", { text: "The family's answers" }));
    left.appendChild(h("div", { class: "presets" }, PRESETS.map(function (ps) {
      return h("button", { class: "chip-btn", text: ps.t, on: { click: function () { S.sel = JSON.parse(JSON.stringify(ps.s)); renderSim(); writeHash(); } } });
    })));

    var qc = h("div", { class: "q" }, h("div", { class: "q-title", text: "How long do you typically spend preparing a meal when you do cook at home?" }));
    COOK_UI.forEach(function (c) { qc.appendChild(optRow("radio", "cook", c.k, c.t, S.sel.cook === c.k, false, function () { S.sel.cook = c.k; refreshSim(); })); });
    var legacy = p.cfg.legacy && S.mode === "legacy";
    qc.appendChild(h("div", { class: "q-note", text: legacy ? "Legacy script: recipes longer than this are removed." : "Not a filter: this answer does not remove or reorder anything. Cook time is only shown beside each recipe and used as the last sort tie-break." }));
    left.appendChild(qc);

    var qp = h("div", { class: "q" }, h("div", { class: "q-title", text: "Which proteins does your family eat? (leave blank = no preference)" }));
    E.PROTEINS.forEach(function (pr) { qp.appendChild(optRow("checkbox", "protein", pr.key, PROT_UI[pr.key], S.sel.proteins.indexOf(pr.key) > -1, false, function (e) { toggle(S.sel.proteins, pr.key, e.target.checked); renderSim(); writeHash(); })); });
    left.appendChild(qp);

    var qd = h("div", { class: "q" }, h("div", { class: "q-title", text: "Which best describes your family's eating pattern?" }));
    E.DIETS.forEach(function (d) {
      var off = meat && (d.key === "vegetarian" || d.key === "vegan");
      qd.appendChild(optRow("radio", "diet", d.key, DIET_UI[d.key], S.sel.diet === d.key, off, function () { S.sel.diet = d.key; refreshSim(); }));
    });
    if (meat) qd.appendChild(h("div", { class: "q-note", text: "Vegetarian and Vegan are not available while a meat or fish protein is ticked above (same rule as the survey page)." }));
    left.appendChild(qd);

    var qr = h("div", { class: "q" }, h("div", { class: "q-title", text: "Dietary restrictions (select all that apply)" }));
    E.RESTRICTIONS.forEach(function (r) { qr.appendChild(optRow("checkbox", "rest", r.key, REST_UI[r.key], S.sel.restrict.indexOf(r.key) > -1, false, function (e) { toggle(S.sel.restrict, r.key, e.target.checked); refreshSim(); })); });
    left.appendChild(qr);

    var qa = h("div", { class: "q" }, h("div", { class: "q-title", text: "Food allergy or intolerance (select all that apply)" }));
    E.ALLERGENS.forEach(function (a) { qa.appendChild(optRow("checkbox", "allergy", a, a, S.sel.allergies.indexOf(a) > -1, false, function (e) { toggle(S.sel.allergies, a, e.target.checked); refreshSim(); })); });
    left.appendChild(qa);

    var right = h("div", { class: "stack", id: "simOut" });
    root.appendChild(h("div", { class: "grid-sim" }, left, right));
    refreshSim(true);
  }

  function refreshSim(skipRebuildForm) {
    sanitizeSel();
    var p = cur(), r = run(), o = opts(), out = clear($("#simOut"));
    writeHash();

    /* controls */
    var ctl = h("div", { class: "card" });
    var row = h("div", { class: "ctl-row" });
    if (p.cfg.legacy) {
      row.appendChild(h("span", { class: "small muted", text: "Rules:" }));
      row.appendChild(h("div", { class: "seg" },
        h("button", { class: S.mode === "target" ? "on" : "", text: "Target (FGS-S rules)", on: { click: function () { S.mode = "target"; renderSim(); } } }),
        h("button", { class: S.mode === "legacy" ? "on" : "", text: "Current live 2.0 script", on: { click: function () { S.mode = "legacy"; renderSim(); } } })));
    } else row.appendChild(h("span", { class: "small muted", text: "Rules: FGS-S (cook time display-only, mains only)" }));
    var seed = h("input", { type: "number", value: S.seed, style: "width:80px;border:1px solid var(--line);border-radius:7px;padding:4px 6px", "aria-label": "Tie-break seed" });
    seed.addEventListener("change", function () { S.seed = parseInt(seed.value, 10) || 0; refreshSim(); });
    row.appendChild(h("span", { class: "small muted", text: "Tie-break shuffle seed:" })); row.appendChild(seed);
    row.appendChild(h("button", { class: "btn", text: "Re-shuffle ties", title: "REDCap draws a new random tie-break on every page load", on: { click: function () { S.seed = Math.floor(Math.random() * 9000) + 1; refreshSim(); } } }));
    row.appendChild(h("button", { class: "btn", text: "Copy link to this scenario", on: { click: function (e) { copyText(location.href, e.target, "Link copied"); } } }));
    ctl.appendChild(row);
    if (p.cfg.legacy) ctl.appendChild(h("div", { class: "small muted", style: "margin-top:8px", text: S.mode === "target"
      ? "Target: what FGS 2.0 will do once it gets the same change as FGS-S - cook time stops filtering recipes, ranking counts main meals only, quicker bundles sort last."
      : "Current: the FGS 2.0 script as it runs in REDCap today - cook time removes recipes, ranking counts main + bonus recipes. Compare it with Target to see what the change does." }));
    out.appendChild(ctl);

    /* participant view */
    var shown = r.ranked.length;
    var frame = h("div", { class: "survey-frame" });
    frame.appendChild(h("div", { class: "survey-bar" }, h("span", { text: "Family Grocery Study · baseline survey" }), h("span", { text: "what the participant sees" })));
    var body = h("div", { class: "survey-body", id: "recipe-bundle-app" });
    body.appendChild(h("div", { class: "survey-q", text: "Please rank the recipes below based on what you and your family members would most like to receive in the next few weeks (the first recipe being the one you'd most like to try). Please build your ranked top 6 below — the ones you’d most like to receive first." }));
    if (!shown) body.appendChild(h("div", { class: "fgs-empty", text: "No recipes fit these selections yet — our team will follow up with options directly." }));
    else {
      body.appendChild(h("div", { class: "fgs-count", text: shown + " bundle" + (shown !== 1 ? "s" : "") + " selected, ranked below. Drag a bundle by its ⋮⋮ handle, or use the arrows, to change the order." }));
      var list = h("div", { class: "fgs-list" });
      r.ranked.forEach(function (x, i) { list.appendChild(rankedItem(x.view, i, shown)); });
      body.appendChild(list);
    }
    frame.appendChild(body);
    out.appendChild(frame);

    /* summary line */
    var cls = shown === 0 ? "bad" : (r.eligibleCount < 6 ? "warn" : "ok");
    var msg = r.eligibleCount + " of " + r.totalBundles + " bundles keep at least one main meal for this family; " + shown + " shown.";
    if (shown && shown < 6) msg += " Fewer than 6 bundles are available, so the ranking box is not full.";
    var prots = Object.keys(r.state.protein);
    if (prots.length) msg += " Protein coverage in the list: " + prots.map(function (k) { return pLabel(k) + " " + (r.coverage[k] || 0) + " of " + shown; }).join(", ") + ".";
    out.appendChild(h("div", { class: "banner " + cls, text: msg }));

    /* researcher export */
    var ex = h("div", { class: "card" });
    ex.appendChild(h("h3", { text: "What REDCap saves for the researchers" }));
    ex.appendChild(h("div", { class: "small muted", style: "margin-bottom:6px", text: "bundle_rank_1 to bundle_rank_6 hold the family's whole filtered bundle: the main meals the participant saw plus the bonus recipes they did not see." }));
    for (var i = 0; i < 6; i++) {
      var x = r.ranked[i], c = h("div", { class: "mono" });
      if (x) {
        c.appendChild(document.createTextNode("Bundle " + x.view.bundle.id + " (" + x.view.bundle.subtitle + "): "));
        x.view.survivors.forEach(function (rec, ri) {
          var isMain = x.view.bundle.recipes.indexOf(rec) < o.mainCount;
          c.appendChild(h("span", { class: isMain ? "" : "bonus", title: isMain ? "main meal (shown to participant)" : "bonus recipe (hidden from participant)", text: rec.name + (isMain ? "" : " [bonus]") }));
          if (ri < x.view.survivors.length - 1) c.appendChild(document.createTextNode("; "));
        });
      } else c.appendChild(h("span", { class: "muted", text: "(empty)" }));
      ex.appendChild(h("div", { class: "field-row" }, h("div", { class: "fn mono", text: "bundle_rank_" + (i + 1) }), c));
    }
    out.appendChild(ex);

    /* why these six */
    out.appendChild(whyTable(r));
    if (!skipRebuildForm) renderSimFormState();
  }
  function renderSimFormState() { /* form inputs are rebuilt by renderSim; refreshSim only redraws the output */ }

  function rankedItem(v, idx, total) {
    var item = h("div", { class: "fgs-ranked-item" });
    item.appendChild(h("div", { class: "fgs-handle", text: "⋮⋮" }));
    item.appendChild(h("div", { class: "fgs-rank", text: String(idx + 1) }));
    var body = h("div", { class: "fgs-body" }, h("h4", { text: "Bundle " + v.bundle.id }));
    var rw = h("div", { class: "fgs-recipes" });
    v.mainSurvivors.forEach(function (rec) {
      var wrap = h("div", { class: "fgs-recipe-wrap" });
      var row = h("div", { class: "fgs-recipe" }, h("span", { class: "fgs-dot fgs-p-" + rec.protein }),
        h("span", { text: rec.name + (rec.timeMin ? " (" + E.fmtTime(rec.timeMin) + ")" : "") }));
      if (rec.allergens.length) row.appendChild(h("span", { class: "fgs-allerg", text: " (" + rec.allergens.join(", ") + ")" }));
      wrap.appendChild(row);
      if (v.state.restrict.gluten_free && !rec.gfStrict && rec.gfNote) wrap.appendChild(h("div", { class: "fgs-swap", text: "GF swap: " + rec.gfNote }));
      rw.appendChild(wrap);
    });
    body.appendChild(rw); item.appendChild(body);
    var ctrl = h("div", { class: "fgs-ctrl" },
      h("button", { type: "button", text: "▲ Move up", disabled: idx === 0, on: { click: function () { move(idx, -1); } } }),
      h("button", { type: "button", text: "▼ Move down", disabled: idx === total - 1, on: { click: function () { move(idx, 1); } } }));
    item.appendChild(ctrl);
    return item;
  }
  function move(idx, d) {
    var r = run(), ids = r.ranked.map(function (x) { return x.key.id; });
    var j = idx + d; if (j < 0 || j >= ids.length) return;
    var t = ids[idx]; ids[idx] = ids[j]; ids[j] = t;
    S.manual = ids; refreshSim();
  }

  function whyTable(r) {
    var o = opts(), p = cur(), st = r.state, prots = Object.keys(st.protein);
    var shownIds = {}; r.ranked.forEach(function (x) { shownIds[x.key.id] = x; });
    var keyById = {}; r.allKeys.forEach(function (k) { keyById[k.id] = k; });
    var rows = [];
    r.ranked.forEach(function (x) { rows.push({ b: x.view.bundle, v: x.view, k: x.key, status: "Shown #" + x.rank, why: x.key.why, cls: "picked" }); });
    r.views.filter(function (v) { return !shownIds[v.bundle.id]; })
      .sort(function (a, b) { return (keyById[b.bundle.id].n - keyById[a.bundle.id].n) || (a.bundle.id - b.bundle.id); })
      .forEach(function (v) { rows.push({ b: v.bundle, v: v, k: keyById[v.bundle.id], status: "Eligible, not shown", why: "outranked by the selection rules (fewer main meals kept" + (prots.length ? " or protein already covered)" : ")"), cls: "" }); });
    p.bundles.forEach(function (b) {
      if (r.views.some(function (v) { return v.bundle.id === b.id; })) return;
      var agg = {};
      b.recipes.slice(0, o.mainCount).forEach(function (rec) { E.whyOut(rec, st, o).forEach(function (w) { agg[w] = (agg[w] || 0) + 1; }); });
      rows.push({ b: b, v: null, k: null, status: "Removed", why: "no main meal left: " + Object.keys(agg).map(function (w) { return w + " ×" + agg[w]; }).join(", "), cls: "gone" });
    });
    var tbl = h("table", { class: "t" });
    var head = h("tr", null, h("th", { text: "Bundle" }), h("th", { text: "Label" }), h("th", { text: "Main meals kept" }));
    prots.forEach(function (k) { head.appendChild(h("th", { text: pLabel(k) + " mains" })); });
    head.appendChild(h("th", { text: "Avg main time" })); head.appendChild(h("th", { text: "Result" })); head.appendChild(h("th", { text: "Why" }));
    tbl.appendChild(h("thead", null, head));
    var tb = h("tbody");
    rows.forEach(function (rw) {
      var tr = h("tr", { class: rw.cls });
      var kept = rw.v ? rw.v.mainSurvivors.length : 0, total = Math.min(o.mainCount, rw.b.recipes.length);
      tr.appendChild(h("td", { text: "Bundle " + rw.b.id }));
      tr.appendChild(h("td", { text: rw.b.subtitle }));
      tr.appendChild(h("td", { text: kept + " of " + total + (rw.v && o.mode === "legacy" ? "  (ranked on " + rw.k.n + " incl. bonus)" : "") }));
      prots.forEach(function (k) { tr.appendChild(h("td", { text: rw.k ? String(rw.k.pc[k]) : "–" })); });
      var sp = rw.v ? E.bundleSpeed(rw.v) : null;
      tr.appendChild(h("td", { text: sp !== null && isFinite(sp) ? Math.round(sp) + " min" : "–" }));
      tr.appendChild(h("td", { text: rw.status }));
      tr.appendChild(h("td", { class: "small", text: rw.why || "" }));
      tb.appendChild(tr);
    });
    tbl.appendChild(tb);
    var d = h("details", { class: "d", open: true }, h("summary", null, h("span", { text: "Why these 6? Every bundle and what happened to it" })), h("div", { class: "inner" },
      h("div", { class: "small muted", style: "margin-bottom:8px", text: o.mode === "legacy"
        ? "Current live script order: most recipes kept (main + bonus) first, then the protein round-robin order. No speed sort."
        : "Order: most MAIN meals kept first; then protein round-robin round; then quicker bundles (lower average main-meal time); then selection order. Ties in selection are broken at random (seed above)." }),
      h("div", { class: "tbl-wrap" }, tbl)));
    return d;
  }

  /* ====================================================================== */
  /*  TAB 2 - all bundles                                                    */
  /* ====================================================================== */
  function flagTags(r) {
    var t = [];
    if (r.vegan) t.push(h("span", { class: "tag g", text: "Vegan" }));
    else if (r.vegetarian) t.push(h("span", { class: "tag g", text: "Veg" }));
    if (r.halal) t.push(h("span", { class: "tag b", text: "Halal" }));
    if (r.kosher) t.push(h("span", { class: "tag b", text: "Kosher" }));
    if (r.gfStrict) t.push(h("span", { class: "tag a", text: "GF" }));
    else if (r.gfOk) t.push(h("span", { class: "tag a", title: r.gfNote || "", text: "GF swap" }));
    return t;
  }
  function renderAll() {
    renderHeader();
    ["sim", "all", "qc", "data", "help"].forEach(function (t) { $("#tab-" + t).hidden = (t !== S.tab); });
    $("#loading").hidden = true;
    if (S.tab === "sim") renderSim();
    if (S.tab === "all") renderAllTab();
    if (S.tab === "qc") renderQC();
    if (S.tab === "data") renderData();
    if (S.tab === "help") renderHelp();
    writeHash();
  }

  function renderAllTab() {
    var root = clear($("#tab-all")), p = cur(), o = opts(), mc = o.mainCount;
    var sums = p.bundles.map(function (b) { return { b: b, s: E.summarizeBundle(b, mc) }; });
    var nVegan = sums.filter(function (x) { return x.s.allVegan; }).length, nVeg = sums.filter(function (x) { return x.s.allVegetarian && !x.s.allVegan; }).length;
    var stats = h("div", { class: "stat-grid" });
    function stat(n, l) { stats.appendChild(h("div", { class: "stat" }, h("div", { class: "n", text: String(n) }), h("div", { class: "l", text: l }))); }
    stat(p.bundles.length, "bundles"); stat(p.bundles.reduce(function (n, b) { return n + b.recipes.length; }, 0), "recipe slots");
    stat(nVegan, "bundles with all-vegan mains"); stat(nVeg, "all-vegetarian (not all vegan)");
    stat(sums.filter(function (x) { return x.s.mainCount && x.s.halal === x.s.mainCount; }).length, "all mains halal");
    stat(sums.filter(function (x) { return x.s.mainCount && x.s.kosher === x.s.mainCount; }).length, "all mains kosher");
    stat(sums.filter(function (x) { return x.s.mainCount && x.s.gfOk === x.s.mainCount; }).length, "all mains GF or GF-swap");
    E.PROTEINS.filter(function (pr) { return pr.key !== "vegetarian"; }).forEach(function (pr) {
      stat(sums.filter(function (x) { return x.s.proteins[pr.key]; }).length, "bundles with " + pr.label.toLowerCase() + " in the mains");
    });
    root.appendChild(stats);

    var tb = h("div", { class: "toolbar" });
    var search = h("input", { type: "search", placeholder: "Search recipe or label…", value: S.allSearch });
    search.addEventListener("input", function () { S.allSearch = search.value; drawBundles(); });
    var fsel = h("select", { "aria-label": "Filter bundles" });
    [["all", "All bundles"], ["vegan", "All-vegan mains"], ["veg", "All-vegetarian mains"], ["halal", "All mains halal"], ["kosher", "All mains kosher"], ["gf", "All mains GF / GF swap"]].forEach(function (x) { fsel.appendChild(h("option", { value: x[0], text: x[1], selected: S.allFilter === x[0] })); });
    fsel.addEventListener("change", function () { S.allFilter = fsel.value; drawBundles(); });
    var psel = h("select", { "aria-label": "Protein in mains" });
    psel.appendChild(h("option", { value: "", text: "Any protein" }));
    E.PROTEINS.forEach(function (pr) { psel.appendChild(h("option", { value: pr.key, text: "Has " + pr.label + " main", selected: S.allProtein === pr.key })); });
    psel.addEventListener("change", function () { S.allProtein = psel.value; drawBundles(); });
    var eff = h("input", { type: "checkbox" }); eff.checked = S.allShowEffect;
    eff.addEventListener("change", function () { S.allShowEffect = eff.checked; drawBundles(); });
    tb.appendChild(search); tb.appendChild(fsel); tb.appendChild(psel);
    tb.appendChild(h("label", { class: "opt", style: "padding:0" }, eff, h("span", { text: "Show what the simulator's current answers would remove" })));
    root.appendChild(tb);
    root.appendChild(h("div", { class: "legend" },
      h("span", null, h("span", { class: "badge main", text: "Main" }), "shown to participants"),
      h("span", null, h("span", { class: "badge bonus", text: "Bonus" }), "saved for researchers only"),
      E.PROTEINS.map(function (pr) { return h("span", null, h("span", { class: "fgs-dot fgs-p-" + pr.key }), pr.label); }),
      h("span", null, h("span", { class: "tag r", text: "Milk" }), "allergen (from the T9 tag or the ingredient scan)")));
    root.appendChild(h("div", { id: "bundleList" }));
    drawBundles();
  }

  function drawBundles() {
    var host = clear($("#bundleList")), p = cur(), o = opts(), mc = o.mainCount;
    var res = S.allShowEffect ? run() : null, st = S.allShowEffect ? E.makeState(S.sel) : null;
    var shownIds = {}; if (res) res.ranked.forEach(function (x) { shownIds[x.key.id] = x.rank; });
    var q = S.allSearch.trim().toLowerCase(), n = 0;
    p.bundles.forEach(function (b) {
      var s = E.summarizeBundle(b, mc);
      if (S.allFilter === "vegan" && !s.allVegan) return;
      if (S.allFilter === "veg" && !s.allVegetarian) return;
      if (S.allFilter === "halal" && !(s.mainCount && s.halal === s.mainCount)) return;
      if (S.allFilter === "kosher" && !(s.mainCount && s.kosher === s.mainCount)) return;
      if (S.allFilter === "gf" && !(s.mainCount && s.gfOk === s.mainCount)) return;
      if (S.allProtein && !s.proteins[S.allProtein]) return;
      if (q && !(b.subtitle.toLowerCase().indexOf(q) > -1 || ("bundle " + b.id) === q || b.recipes.some(function (r) { return r.name.toLowerCase().indexOf(q) > -1; }))) return;
      n++;
      var eligible = res ? res.views.some(function (v) { return v.bundle.id === b.id; }) : true;
      var card = h("div", { class: "bundle-card" + (res && !eligible ? " off" : "") });
      var head = h("div", { class: "bundle-head" }, h("h3", { text: "Bundle " + b.id + " · " + b.subtitle }),
        h("span", { class: "meta", text: s.avgMainMin ? "avg main time " + Math.round(s.avgMainMin) + " min" : "" }),
        s.allVegan ? h("span", { class: "tag g", text: "all-vegan mains" }) : (s.allVegetarian ? h("span", { class: "tag g", text: "all-vegetarian mains" }) : null),
        res ? (shownIds[b.id] ? h("span", { class: "tag g", text: "shown at #" + shownIds[b.id] }) : h("span", { class: "tag " + (eligible ? "n" : "r"), text: eligible ? "eligible, not in top 6" : "removed: no main meal left" })) : null);
      card.appendChild(head);
      var tbl = h("table", { class: "t" }, h("thead", null, h("tr", null, h("th", { text: "#" }), h("th", { text: "Role" }), h("th", { text: "Recipe" }), h("th", { text: "Protein" }), h("th", { text: "Diet / faith / GF" }), h("th", { text: "Allergens" }), h("th", { text: "Time" }))));
      var tbd = h("tbody");
      b.recipes.forEach(function (r, i) {
        var main = i < mc, why = st ? E.whyOut(r, st, o) : [];
        var tr = h("tr", { class: (main ? "" : "bonus-row ") + (why.length ? "dropped" : "") });
        tr.appendChild(h("td", { text: String(i + 1) }));
        tr.appendChild(h("td", null, h("span", { class: "badge " + (main ? "main" : "bonus"), text: main ? "Main" : "Bonus" })));
        var nm = h("td", { class: "nm" }, h("span", { title: "Sheet text: " + (r.rawName || r.name), text: r.name }));
        if (why.length) nm.appendChild(h("div", { class: "why", text: why.join("; ") }));
        if (r.patched) nm.appendChild(h("div", { class: "small muted", title: r.patched.join(" | "), text: "patched" }));
        tr.appendChild(nm);
        tr.appendChild(h("td", null, h("span", { class: "fgs-dot fgs-p-" + r.protein }), " " + pLabel(r.protein)));
        tr.appendChild(h("td", null, flagTags(r)));
        tr.appendChild(h("td", null, r.allergens.map(function (a) { var sc = r.scannedAllergens && r.scannedAllergens.indexOf(a) > -1; return h("span", { class: "tag r", title: sc ? "found in ingredients, not in the T9 tag" : "T9 tag", text: a + (sc ? "*" : "") }); })));
        tr.appendChild(h("td", { text: r.timeMin ? E.fmtTime(r.timeMin) : "–", title: r.rawTime }));
        tbd.appendChild(tr);
      });
      tbl.appendChild(tbd);
      card.appendChild(h("div", { class: "tbl-wrap" }, tbl));
      host.appendChild(card);
    });
    if (!n) host.appendChild(h("div", { class: "muted", text: "No bundles match." }));
  }

  /* ====================================================================== */
  /*  TAB 3 - checks, compare, coverage                                      */
  /* ====================================================================== */
  function renderQC() {
    var root = clear($("#tab-qc")), p = cur(), o = opts();
    /* sheet checks */
    var checks = p.src.kind === "script" ? [] : E.qcChecks(p.bundles, o.mainCount);
    var cnt = { error: 0, warn: 0, info: 0 }; checks.forEach(function (c) { cnt[c.level]++; });
    var d1 = h("details", { class: "d", open: true }, h("summary", null, h("span", { text: "Spreadsheet checks" }),
      h("span", { class: "small muted", text: cnt.error + " errors · " + cnt.warn + " warnings · " + cnt.info + " notes" })));
    var in1 = h("div", { class: "inner" });
    if (p.src.kind === "script") in1.appendChild(h("div", { class: "banner info", text: "The simulator is currently using data embedded in a REDCap script, which has no tags or raw times to check. Switch back to the repo file to see the spreadsheet checks." }));
    else if (!checks.length) in1.appendChild(h("div", { class: "banner ok", text: "Nothing to flag." }));
    var lv = h("select"); [["", "All levels"], ["error", "Errors"], ["warn", "Warnings"], ["info", "Notes"]].forEach(function (x) { lv.appendChild(h("option", { value: x[0], text: x[1] })); });
    var ctab = h("div", { class: "tbl-wrap" });
    function drawChecks() {
      var rows = checks.filter(function (c) { return !lv.value || c.level === lv.value; });
      var t = h("table", { class: "t" }, h("thead", null, h("tr", null, h("th", { text: "Level" }), h("th", { text: "Where" }), h("th", { text: "What" }))));
      var tb = h("tbody"); rows.slice(0, 400).forEach(function (c) { tb.appendChild(h("tr", null, h("td", null, h("span", { class: "tag " + (c.level === "error" ? "r" : c.level === "warn" ? "a" : "n"), text: c.level })), h("td", { text: c.where }), h("td", { text: c.msg }))); });
      t.appendChild(tb); clear(ctab).appendChild(t);
      if (rows.length > 400) ctab.appendChild(h("div", { class: "small muted", text: "Showing the first 400 of " + rows.length + "." }));
    }
    lv.addEventListener("change", drawChecks);
    in1.appendChild(h("div", { class: "toolbar" }, lv)); in1.appendChild(ctab); d1.appendChild(in1); root.appendChild(d1); drawChecks();

    /* patches */
    var d2 = h("details", { class: "d" }, h("summary", null, h("span", { text: "Reviewed corrections applied on top of the sheet (patches)" }), h("span", { class: "small muted", text: p.patches.length + " entries" })));
    var in2 = h("div", { class: "inner" });
    in2.appendChild(h("div", { class: "small muted", style: "margin-bottom:8px", text: "These live in " + (p.cfg.patches || "data/patches/*.json") + " and work like the REDCap script's DATA_PATCHES: they can only add an allergen, upgrade a recipe to vegetarian/vegan, or set the protein bucket. 'Used on' = how many recipe slots in the current sheet matched; 0 means the recipe was renamed or removed and the patch does nothing." }));
    var pt = h("table", { class: "t" }, h("thead", null, h("tr", null, h("th", { text: "Recipe" }), h("th", { text: "Change" }), h("th", { text: "Used on" }), h("th", { text: "Why" }))));
    var pb = h("tbody");
    p.patches.forEach(function (pa, i) {
      var ch = []; if (pa.allergensAdd) ch.push("+ allergens: " + pa.allergensAdd.join(", ")); if (pa.vegetarian) ch.push("vegetarian"); if (pa.vegan) ch.push("vegan"); if (pa.protein) ch.push("protein: " + pa.protein);
      pb.appendChild(h("tr", { class: p.used[i] ? "" : "gone" }, h("td", { text: pa.name }), h("td", { text: ch.join("; ") }), h("td", { text: String(p.used[i] || 0) }), h("td", { class: "small", text: pa.why || "" })));
    });
    pt.appendChild(pb); in2.appendChild(h("div", { class: "tbl-wrap" }, pt)); d2.appendChild(in2); root.appendChild(d2);

    /* compare with script */
    var d3 = h("details", { class: "d", open: true }, h("summary", null, h("span", { text: "Compare the spreadsheet with a REDCap JS Injector file" }), h("span", { class: "small muted", text: "finds drift between the sheet and what is live" })));
    var in3 = h("div", { class: "inner" });
    in3.appendChild(h("div", { class: "small muted", style: "margin-bottom:8px", text: "Pick the injector .js file for this project (" + (p.cfg.script || "recipe_bundle_injector_….js") + "). The page reads its embedded BUNDLES data and its patches and lists every difference from the spreadsheet. Nothing is uploaded anywhere." }));
    in3.appendChild(h("button", { class: "btn primary", text: "Choose injector .js file…", on: { click: function () { $("#jsPick").click(); } } }));
    in3.appendChild(h("div", { id: "cmpOut", style: "margin-top:12px" }));
    d3.appendChild(in3); root.appendChild(d3);

    /* coverage */
    var d4 = h("details", { class: "d" }, h("summary", null, h("span", { text: "Coverage sweep: can every family get 6 bundles?" }), h("span", { class: "small muted", text: "all diet × restriction × allergy combinations" })));
    var in4 = h("div", { class: "inner" });
    in4.appendChild(h("div", { class: "small muted", style: "margin-bottom:8px", text: "Runs every combination of eating pattern (4) × restrictions (8) × allergies (512) = 16,384 families through the same filter and counts how many bundles keep at least one main meal. Protein is left blank. Cook time is ignored under the target rules." }));
    in4.appendChild(h("button", { class: "btn primary", text: "Run the sweep", on: { click: function () { sweep(in4); } } }));
    in4.appendChild(h("div", { id: "sweepOut", style: "margin-top:12px" }));
    d4.appendChild(in4); root.appendChild(d4);
    if (S.lastCmp && S.lastCmp.proj === S.proj) drawCompare(S.lastCmp);
  }

  function parseScriptFile(text) {
    var m = /var BUNDLES\s*=\s*(\[[\s\S]*?\]);\s*\n/.exec(text);
    if (!m) throw new Error("No `var BUNDLES = [...]` line found in that file.");
    var bundles = JSON.parse(m[1]), patches = [], patchNote = "";
    var pm = /var DATA_PATCHES\s*=\s*(\[[\s\S]*?\n\s*\]);\s*\n/.exec(text);
    if (pm) {
      try { patches = JSON.parse(pm[1]); }
      catch (e) {
        try { patches = JSON.parse(pm[1].replace(/([{,]\s*)([A-Za-z_]\w*)\s*:/g, '$1"$2":').replace(/,(\s*[}\]])/g, "$1")); }
        catch (e2) { patchNote = "The script's DATA_PATCHES could not be read; compared without them."; }
      }
    }
    var tag = /var FGS_BUILD\s*=\s*"([^"]+)"/.exec(text);
    return { bundles: bundles, patches: patches, patchNote: patchNote, build: tag ? tag[1] : "" };
  }
  var CMP_FIELDS = ["protein", "vegetarian", "vegan", "halal", "kosher", "gfStrict", "gfOk", "allergens", "timeMin", "gfNote"];
  function sameVal(a, b, f) {
    if (f === "allergens") return JSON.stringify((a || []).slice().sort()) === JSON.stringify((b || []).slice().sort());
    if (f === "gfNote") return String(a || "").toLowerCase().replace(/\s+/g, " ").trim() === String(b || "").toLowerCase().replace(/\s+/g, " ").trim();
    if (f === "timeMin") return (a || 0) === (b || 0);
    return a === b;
  }
  function compareData(sheetB, scriptB) {
    var diffs = [], notes = [];
    scriptB.forEach(function (sb) {
      var mine = sheetB.filter(function (b) { return b.id === sb.id; })[0];
      if (!mine) { notes.push("Bundle " + sb.id + " (" + sb.subtitle + ") is in the script but not in the sheet."); return; }
      if (mine.subtitle !== sb.subtitle) notes.push("Bundle " + sb.id + " label: script \"" + sb.subtitle + "\" vs sheet \"" + mine.subtitle + "\".");
      if (mine.recipes.length !== sb.recipes.length) notes.push("Bundle " + sb.id + ": script has " + sb.recipes.length + " recipes, sheet has " + mine.recipes.length + ".");
      sb.recipes.forEach(function (r, i) {
        var m = mine.recipes[i]; if (!m) return;
        if (E.normName(m.name) !== E.normName(r.name) && E.normName(m.name).replace(/&/g, "and") !== E.normName(r.name).replace(/&/g, "and"))
          diffs.push({ id: sb.id, slot: i + 1, name: m.name, field: "name", script: r.name, sheet: m.name });
        CMP_FIELDS.forEach(function (f) {
          if (!sameVal(r[f], m[f], f)) diffs.push({ id: sb.id, slot: i + 1, name: m.name, field: f, script: Array.isArray(r[f]) ? r[f].join(", ") : String(r[f]), sheet: Array.isArray(m[f]) ? m[f].join(", ") : String(m[f]) });
        });
      });
    });
    sheetB.forEach(function (b) { if (!scriptB.some(function (s) { return s.id === b.id; })) notes.push("Bundle " + b.id + " (" + b.subtitle + ") is in the sheet but not in the script."); });
    return { diffs: diffs, notes: notes };
  }
  function onScriptFile(file) {
    var fr = new FileReader();
    fr.onload = function () {
      try {
        var sc = parseScriptFile(String(fr.result)), p = cur();
        var patched = JSON.parse(JSON.stringify(sc.bundles)); E.applyPatches(patched, sc.patches);
        var cmp = compareData(p.bundles, patched);
        S.lastCmp = { proj: S.proj, file: file.name, script: sc, patched: patched, cmp: cmp };
        S.tab = "qc"; renderAll();
      } catch (e) { alert(e.message); }
    };
    fr.readAsText(file);
  }
  function drawCompare(c) {
    var out = clear($("#cmpOut")); if (!out) return;
    var byField = {}; c.cmp.diffs.forEach(function (d) { byField[d.field] = (byField[d.field] || 0) + 1; });
    out.appendChild(h("div", { class: "banner " + (c.cmp.diffs.length || c.cmp.notes.length ? "warn" : "ok"), text:
      c.file + (c.script.build ? " (build " + c.script.build + ")" : "") + ": " + (c.cmp.diffs.length ? plural(c.cmp.diffs.length, "difference") + " between the script's data (patches applied) and the sheet - " + Object.keys(byField).map(function (k) { return k + " " + byField[k]; }).join(", ") + "." : "no differences in the recipe data.") + (c.script.patchNote ? " " + c.script.patchNote : "") }));
    if (c.cmp.notes.length) out.appendChild(h("ul", null, c.cmp.notes.map(function (n) { return h("li", { text: n }); })));
    if (c.cmp.diffs.length) {
      var t = h("table", { class: "t" }, h("thead", null, h("tr", null, h("th", { text: "Bundle" }), h("th", { text: "Slot" }), h("th", { text: "Recipe (sheet)" }), h("th", { text: "Field" }), h("th", { text: "REDCap script" }), h("th", { text: "Spreadsheet" }))));
      var tb = h("tbody"); c.cmp.diffs.slice(0, 500).forEach(function (d) { tb.appendChild(h("tr", null, h("td", { text: String(d.id) }), h("td", { text: String(d.slot) }), h("td", { text: d.name }), h("td", { text: d.field }), h("td", { text: d.script }), h("td", { text: d.sheet }))); });
      t.appendChild(tb); out.appendChild(h("div", { class: "tbl-wrap" }, t));
      if (c.cmp.diffs.length > 500) out.appendChild(h("div", { class: "small muted", text: "Showing the first 500." }));
    }
    out.appendChild(h("div", { class: "ctl-row", style: "margin-top:10px" },
      h("button", { class: "btn", text: "Simulate with the script's data instead", on: { click: function () {
        PROJ[S.proj] = Object.assign({}, cur(), { bundles: c.patched, used: {}, src: { kind: "script", name: c.file } }); S.tab = "sim"; S.manual = null; renderAll(); } } }),
      h("span", { class: "small muted", text: "Use this to see exactly what participants get today." })));
  }

  function sweep(host) {
    var p = cur(), o = { mode: "target", mainCount: p.cfg.mainCount };
    var out = clear($("#sweepOut", host)); var bar = h("div", { class: "progress" }, h("i")); out.appendChild(bar);
    var diets = ["none", "vegetarian", "vegan", "pescatarian"], rests = [], i, j;
    for (i = 0; i < 8; i++) rests.push(["halal", "kosher", "gluten_free"].filter(function (x, k) { return i & (1 << k); }));
    var res = [], total = 4 * 8 * 512, done = 0, di = 0;
    function step() {
      var t0 = Date.now();
      while (di < 4 && Date.now() - t0 < 40) {
        for (var ri = 0; ri < 8; ri++) for (var m = 0; m < 512; m++) {
          var al = E.ALLERGENS.filter(function (x, k) { return m & (1 << k); });
          var st = E.makeState({ diet: diets[di], restrict: rests[ri], allergies: al });
          res.push({ diet: diets[di], restrict: rests[ri], allergies: al, n: E.computeViews(p.bundles, st, o).length });
        }
        di++;
      }
      $("i", bar).style.width = Math.round(100 * di / 4) + "%";
      if (di < 4) setTimeout(step, 0); else finish();
    }
    function finish() {
      clear(out);
      var real = res.filter(function (x) { return x.allergies.length <= 3 && x.restrict.length <= 2; });
      function cnt(a, f) { return a.filter(f).length; }
      out.appendChild(h("div", { class: "stat-grid" },
        h("div", { class: "stat" }, h("div", { class: "n", text: String(cnt(res, function (x) { return x.n >= 6; })) }), h("div", { class: "l", text: "of " + res.length + " combinations still fill all 6 slots" })),
        h("div", { class: "stat" }, h("div", { class: "n", text: String(cnt(real, function (x) { return x.n < 6; })) }), h("div", { class: "l", text: "plausible families (≤3 allergies, ≤2 restrictions) get fewer than 6" })),
        h("div", { class: "stat" }, h("div", { class: "n", text: String(cnt(real, function (x) { return x.n === 0; })) }), h("div", { class: "l", text: "plausible families get no bundle at all" }))));
      var bad = real.filter(function (x) { return x.n < 6; }).sort(function (a, b) { return (a.n - b.n) || ((a.allergies.length + a.restrict.length) - (b.allergies.length + b.restrict.length)); });
      out.appendChild(h("div", { class: "small muted", style: "margin:8px 0", text: "Plausible families with fewer than 6 bundles, fewest bundles first (smallest selections first among ties). Showing up to 60 of " + bad.length + "." }));
      var t = h("table", { class: "t" }, h("thead", null, h("tr", null, h("th", { text: "Eating pattern" }), h("th", { text: "Restrictions" }), h("th", { text: "Allergies" }), h("th", { text: "Bundles available" }))));
      var tb = h("tbody"); bad.slice(0, 60).forEach(function (x) { tb.appendChild(h("tr", null, h("td", { text: x.diet }), h("td", { text: x.restrict.join(", ") || "–" }), h("td", { text: x.allergies.join(", ") || "–" }), h("td", { text: String(x.n) }))); });
      t.appendChild(tb); out.appendChild(h("div", { class: "tbl-wrap" }, t));
    }
    setTimeout(step, 0);
  }

  /* ====================================================================== */
  /*  TAB 4 - data & export                                                  */
  /* ====================================================================== */
  function cleanForScript(bundles) {
    return bundles.map(function (b) {
      return { id: b.id, subtitle: b.subtitle, recipes: b.recipes.map(function (r) {
        return { name: r.name, protein: r.protein, vegetarian: r.vegetarian, vegan: r.vegan, halal: r.halal, kosher: r.kosher, gfStrict: r.gfStrict, gfOk: r.gfOk, gfNote: r.gfNote || null, allergens: r.allergens.slice(), timeMin: r.timeMin === undefined ? null : r.timeMin };
      }) };
    });
  }
  function csvCell(v) { v = v === null || v === undefined ? "" : String(v); return /[",\n]/.test(v) ? "\"" + v.replace(/"/g, "\"\"") + "\"" : v; }
  function download(name, text, type) {
    var a = h("a", { href: URL.createObjectURL(new Blob([text], { type: type || "text/plain" })), download: name });
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function copyText(text, btn, done) {
    function ok() { var t = btn.textContent; btn.textContent = done || "Copied"; setTimeout(function () { btn.textContent = t; }, 1400); }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(ok, function () { fallback(); });
    else fallback();
    function fallback() { var ta = h("textarea"); ta.value = text; document.body.appendChild(ta); ta.select(); try { document.execCommand("copy"); ok(); } catch (e) {} ta.remove(); }
  }
  function renderData() {
    var root = clear($("#tab-data")), p = cur(), mc = p.cfg.mainCount;
    var line = "var BUNDLES = " + JSON.stringify(cleanForScript(p.bundles)) + ";";
    root.appendChild(h("div", { class: "card", style: "margin-bottom:14px" },
      h("h3", { text: "Refresh the REDCap script's data from this sheet" }),
      h("p", { class: "small muted", text: "Copy this line and paste it over the existing `var BUNDLES = [...];` line in " + (p.cfg.script || "the injector script") + ". Patches are already applied, so the script's own DATA_PATCHES become a no-op. Labels, main meals (first " + mc + " recipes) and every flag come from this page, so what you see here is what REDCap will do." }),
      h("div", { class: "ctl-row" },
        h("button", { class: "btn primary", text: "Copy var BUNDLES line", on: { click: function (e) { copyText(line, e.target); } } }),
        h("button", { class: "btn", text: "Download as .json", on: { click: function () { download(S.proj + "_bundles.json", JSON.stringify(cleanForScript(p.bundles), null, 1), "application/json"); } } }),
        h("button", { class: "btn", text: "Download all recipes as .csv", on: { click: function () { download(S.proj + "_recipes.csv", recipesCSV(p), "text/csv"); } } })),
      h("div", { class: "small muted", style: "margin-top:8px", text: line.length.toLocaleString() + " characters · " + plural(p.bundles.length, "bundle") })));
    var rows = [];
    p.bundles.forEach(function (b) { b.recipes.forEach(function (r, i) { rows.push({ b: b, r: r, i: i }); }); });
    var t = h("table", { class: "t" }, h("thead", null, h("tr", null, ["Bundle", "Label", "#", "Role", "Recipe (cleaned)", "Sheet text", "Protein", "Vegan", "Veg", "Halal", "Kosher", "GF", "GF swap note", "Allergens", "Min", "Raw tags"].map(function (x) { return h("th", { text: x }); }))));
    var tb = h("tbody");
    rows.forEach(function (x) {
      var r = x.r;
      tb.appendChild(h("tr", { class: x.i < mc ? "" : "bonus-row" }, h("td", { text: String(x.b.id) }), h("td", { text: x.b.subtitle }), h("td", { text: String(x.i + 1) }), h("td", { text: x.i < mc ? "Main" : "Bonus" }),
        h("td", { text: r.name }), h("td", { class: "small", text: r.rawName }), h("td", { text: r.protein }), h("td", { text: r.vegan ? "Y" : "" }), h("td", { text: r.vegetarian ? "Y" : "" }), h("td", { text: r.halal ? "Y" : "" }), h("td", { text: r.kosher ? "Y" : "" }),
        h("td", { text: r.gfStrict ? "GF" : (r.gfOk ? "swap" : "") }), h("td", { class: "small", text: r.gfNote || "" }), h("td", { text: r.allergens.join(", ") }), h("td", { text: r.timeMin || "" }), h("td", { class: "small mono", text: (r.rawTags || "").replace(/\s+/g, " ") })));
    });
    t.appendChild(tb);
    root.appendChild(h("div", { class: "card" }, h("h3", { text: "Every recipe as the simulator reads it" }), h("div", { class: "tbl-wrap", style: "max-height:70vh;overflow:auto" }, t)));
  }
  function recipesCSV(p) {
    var mc = p.cfg.mainCount, head = ["project", "bundle", "label", "slot", "role", "recipe", "sheet_text", "protein", "vegan", "vegetarian", "halal", "kosher", "gf_as_written", "gf_ok", "gf_swap_note", "allergens", "total_minutes", "raw_tags", "raw_time", "patched"];
    var out = [head.join(",")];
    p.bundles.forEach(function (b) { b.recipes.forEach(function (r, i) {
      out.push([S.proj, b.id, b.subtitle, i + 1, i < mc ? "main" : "bonus", r.name, r.rawName, r.protein, r.vegan, r.vegetarian, r.halal, r.kosher, r.gfStrict, r.gfOk, r.gfNote, r.allergens.join("; "), r.timeMin, r.rawTags, r.rawTime, (r.patched || []).join(" | ")].map(csvCell).join(","));
    }); });
    return out.join("\n");
  }

  /* ====================================================================== */
  /*  TAB 5 - how it works                                                   */
  /* ====================================================================== */
  function renderHelp() {
    var root = clear($("#tab-help")), p = cur();
    var d = h("div", { class: "prose card" });
    d.innerHTML = [
      "<h3 style='margin-top:0'>What this page is</h3>",
      "<p>A copy of the REDCap recipe-ranking logic that runs in a browser, so anyone can try a family's answers and see the six bundles that family would be offered, without opening REDCap. It reads the same bundle spreadsheets the team maintains.</p>",
      "<h3>Updating the bundles</h3>",
      "<ol><li>Replace the spreadsheet in the repo's <code>data/</code> folder with the new version (keep the file name, keep the same sheet name and columns: recipe name in column A, <code>Ingredients</code>, <code>Tags</code>, <code>Prep/Cook Time</code>).</li><li>Commit. GitHub Pages republishes in about a minute; reload this page.</li><li>To preview a sheet first, use <b>Try a different .xlsx…</b> at the top. It changes only your browser tab.</li></ol>",
      "<h3>How the sheet is read</h3>",
      "<ul><li>A row whose first cell says <code>Bundle N</code> starts bundle N. The label comes from <code>Bundle N (Label)</code> in the same cell, otherwise from another sheet in the workbook, otherwise from <code>labels</code> in <code>config.json</code>.</li><li>The first <b>" + p.cfg.mainCount + "</b> recipes of a bundle are the <b>main meals</b>. The rest are <b>bonus</b>.</li><li><b>Tags</b>: <code>V</code> = vegan, <code>Veg</code> = vegetarian, <code>H</code> = halal, <code>K</code> = kosher, <code>GF</code> = gluten-free as written, <code>GF: …</code> or <code>GF- …</code> = gluten-free with that swap, <code>T9 (Milk, Wheat)</code> = allergens. A bare <code>V</code> on a recipe that lists milk or eggs is not treated as vegan.</li><li><b>Allergens</b> are the T9 tag <i>plus</i> an automatic ingredient scan (milk, egg, wheat, soy, sesame, nuts, fish, shellfish). Items marked optional, and vegan substitutes, are ignored. Scan-only allergens carry an asterisk.</li><li><b>Time</b> = every minutes figure plus hours in the Prep/Cook cell, added together.</li><li><b>Protein</b> = chicken, turkey, beef or fish named in the recipe title; otherwise vegetarian-tagged recipes are plant-based; otherwise the most-mentioned meat in the ingredients; otherwise plant-based.</li></ul>",
      "<h3>What a family's answers do</h3>",
      "<ul><li><b>Eating pattern</b>: Vegetarian keeps vegetarian or vegan recipes; Vegan keeps vegan recipes; Pescatarian keeps fish and plant-based recipes.</li><li><b>Restrictions</b>: Halal needs <code>H</code>, Kosher needs <code>K</code>, Gluten-free needs <code>GF</code> or a GF swap (the swap text is shown under the recipe).</li><li><b>Allergies</b>: any recipe containing a ticked allergen is removed.</li><li><b>Cook time</b>: under the target rules it removes nothing. Time is printed next to each recipe.</li><li>A bundle is offered only if at least one <b>main</b> meal survives. Participants see only the surviving main meals; the saved <code>bundle_rank_*</code> text holds the surviving main <i>and</i> bonus recipes.</li></ul>",
      "<h3>Which six bundles, and in what order</h3>",
      "<ul><li><b>No protein ticked</b>: the bundles keeping the most main meals, ties broken at random.</li><li><b>Proteins ticked</b>: bundles are picked in rounds across the ticked proteins (so beef and fish are not crowded out by chicken); each protein offers its richest bundles first; leftover slots go to the richest remaining bundles.</li><li><b>Display order</b>: most main meals kept first, then the protein round, then <b>quicker bundles first</b> (lowest average main-meal time) as the last tie-break.</li><li>Random ties are re-drawn on every REDCap page load; use the seed or <i>Re-shuffle ties</i> to see other draws. The participant can move bundles up or down; this page lets you do the same.</li></ul>",
      "<h3>Target vs current FGS 2.0 rules</h3>",
      "<p>FGS-S already works the target way. FGS 2.0 (3 and 4 meal) is scheduled to get the same change, so the page defaults to the target rules for every project. For FGS 2.0 you can flip to <b>Current live 2.0 script</b> to see what REDCap does today (cook time removes recipes; ranking counts main plus bonus recipes; no speed sort).</p>",
      "<h3>Checks</h3>",
      "<p><b>Checks &amp; compare</b> lists suspicious tags, shows the reviewed patches, compares the sheet with a REDCap injector file, and sweeps every diet × restriction × allergy combination to find families who would get fewer than six bundles.</p>",
      "<h3>What this page does not do</h3>",
      "<p>It does not read REDCap, so it never shows participant data. It does not simulate dragging or the survey's page flow, and the cook-time answer is carried from an earlier page in REDCap but is a plain choice here. The row order shown is what REDCap's first render would show.</p>"
    ].join("");
    root.appendChild(d);
  }

  /* ====================================================================== */
  /*  boot                                                                   */
  /* ====================================================================== */
  function boot() {
    readHash();
    document.getElementById("tabs").addEventListener("click", function (e) {
      var b = e.target.closest("button[data-tab]"); if (!b) return; S.tab = b.getAttribute("data-tab"); renderAll();
    });
    document.getElementById("xlsxPick").addEventListener("change", function (e) { if (e.target.files[0]) loadUpload(e.target.files[0]); e.target.value = ""; });
    document.getElementById("jsPick").addEventListener("change", function (e) { if (e.target.files[0]) onScriptFile(e.target.files[0]); e.target.value = ""; });
    fetchOK("config.json", "json").then(function (c) {
      CFG = c; if (!cfgOf(S.proj)) S.proj = c.projects[0].id;
      return Promise.all(c.projects.map(loadProject));
    }).then(function () {
      if (!cur().cfg.legacy) S.mode = "target";
      renderAll();
    }).catch(function (e) {
      $("#loading").hidden = false;
      clear($("#loading")).appendChild(h("div", { class: "err" },
        h("b", { text: "Could not load the bundle spreadsheets. " }),
        document.createTextNode(String(e.message || e) + " — If you opened this file directly from your computer, the browser blocks reading the data folder. Open it from GitHub Pages or run a local web server (python -m http.server in this folder).")));
    });
  }
  boot();
})();
