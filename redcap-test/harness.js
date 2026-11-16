/* REDCap script tester: a fake baseline survey page (FGS-S, FGS 2.0 3-meal or 4-meal, detected from the script you load) + your real injector file + automated checks. */
(function () {
  "use strict";
  var $ = window.jQuery, E = window.FGSEngine;
  var SFX = "", MAIN = 3, COOKKEY = "fgss_cook_time", GLOB = "FGSS", KIND = "FGS-S";
  function readMeta(t) {
    var m;
    if ((m = /var SFX\s*=\s*"([^"]*)"/.exec(t))) SFX = m[1];
    if ((m = /var MAIN_COUNT\s*=\s*(\d+)/.exec(t))) MAIN = parseInt(m[1], 10);
    if ((m = /var COOK_KEY\s*=\s*"([^"]*)"/.exec(t))) COOKKEY = m[1];
    if ((m = /window\.(FGS\w*)\s*=\s*\{\s*\n?\s*build/.exec(t))) GLOB = m[1];
    if ((m = /var SCRIPT_ID\s*=\s*"([^"]*)"/.exec(t))) KIND = m[1] === "3MEAL" ? "FGS 2.0 3-meal" : m[1] === "4MEAL" ? "FGS 2.0 4-meal" : "FGS-S";
  }

  /* ---- field specs: codes and labels copied from FGS-S_Baseline_Survey_instrument.csv ---- */
  var FIELDS = {
    baseline_consent: { type: "radio", label: "Consent (page 1)", choices: [["1", "Yes, I agree"], ["2", "No"]] },
    cook_time: { type: "radio", label: "How long do you typically spend preparing a meal when you do cook at home?", choices: [["a", "30 minutes or less"], ["b", "45 minutes to 1 hour"], ["c", "More than 1 hour"]] },
    protein_preference: { type: "checkbox", label: "Which proteins does your family eat? (select all that apply — leave blank if no particular preference)", choices: [["1", "Chicken"], ["2", "Turkey"], ["3", "Beef"], ["4", "Fish or shellfish"], ["5", "Tofu or other plant-based protein"]] },
    dietary_preference: { type: "radio", label: "Which best describes your family's eating pattern?", choices: [["1", "No specific preference (we eat meat, fish, and plant-based meals)"], ["2", "Vegetarian"], ["3", "Vegan"], ["4", "Pescatarian (fish but no other meat)"]] },
    dietary_restrictions: { type: "checkbox", label: "My family follows the following dietary restrictions", choices: [["1", "Halal"], ["2", "Kosher"], ["3", "Gluten-free"], ["4", "None of the above"]] },
    food_allergies: { type: "checkbox", label: "Does anyone in your family have a food allergy or intolerance to any of the following?", choices: [["1", "Milk"], ["2", "Eggs"], ["3", "Wheat"], ["4", "Soy"], ["5", "Peanuts"], ["6", "Tree nuts"], ["7", "Fish"], ["8", "Shellfish"], ["9", "Sesame"], ["10", "Other (specify)"], ["11", "No known food allergies"]] }
  };
  var PCODE = { chicken: "1", turkey: "2", beef: "3", fish_shellfish: "4", vegetarian: "5" };
  var DCODE = { none: "1", vegetarian: "2", vegan: "3", pescatarian: "4" };
  var RCODE = { halal: "1", kosher: "2", gluten_free: "3" };
  var ACODE = { Milk: "1", Eggs: "2", Wheat: "3", Soy: "4", Peanuts: "5", "Tree Nuts": "6", Fish: "7", Shellfish: "8", Sesame: "9" };

  /* ---- console capture (script output is shown on the right) ---- */
  var logLines = [], errCount = 0;
  ["log", "warn", "error"].forEach(function (k) {
    var orig = console[k];
    console[k] = function () {
      var t = Array.prototype.slice.call(arguments).map(function (a) { try { return typeof a === "string" ? a : JSON.stringify(a); } catch (e) { return String(a); } }).join(" ");
      if (/^\[FGS/.test(t)) { logLines.push((k === "log" ? "" : k.toUpperCase() + " ") + t); if (k === "error") errCount++; if (logLines.length > 400) logLines.shift(); paintLog(); }
      return orig.apply(console, arguments);
    };
  });
  var logTimer = null;
  function paintLog() { if (logTimer) return; logTimer = setTimeout(function () { logTimer = null; var p = document.getElementById("log"); p.textContent = logLines.join("\n") || "(no output yet)"; p.scrollTop = p.scrollHeight; }, 100); }

  /* ---- mock REDCap markup (same selectors REDCap uses and the script looks for) ---- */
  function radioRow(name) {
    var f = FIELDS[name], h = '<tr id="' + name + SFX + '-tr"><td class="labelrc">' + f.label + '</td><td class="data">';
    f.choices.forEach(function (c) { h += '<div class="choicevert"><input type="radio" name="' + name + SFX + '___radio" id="' + name + SFX + '-' + c[0] + '" value="' + c[0] + '"> <label for="' + name + SFX + '-' + c[0] + '">' + c[1] + '</label></div>'; });
    return h + '<input type="hidden" name="' + name + SFX + '" value=""><a class="smalllink" href="#" data-reset="' + name + SFX + '">reset</a></td></tr>';
  }
  function checkRow(name) {
    var f = FIELDS[name], h = '<tr id="' + name + SFX + '-tr"><td class="labelrc">' + f.label + '</td><td class="data">';
    f.choices.forEach(function (c) {
      h += '<div class="choicevert"><input type="checkbox" name="__chkn__' + name + SFX + '" code="' + c[0] + '" id="id-__chk__' + name + SFX + '_RC_' + c[0] + '"> <label for="id-__chk__' + name + SFX + '_RC_' + c[0] + '">' + c[1] + '</label>' +
           '<input type="hidden" name="' + name + SFX + '___' + c[0] + '" value=""></div>';
    });
    return h + '</td></tr>';
  }
  function buildPage(n) {
    var h = "<table><tbody>";
    if (n === 1) h += radioRow("baseline_consent") + radioRow("cook_time");
    else {
      h += checkRow("protein_preference") + radioRow("dietary_preference") + checkRow("dietary_restrictions") + checkRow("food_allergies");
      h += '<tr id="recipe_ranks-tr"><td colspan="2"><div class="labelrc">Please rank the recipes below…</div><div id="recipe-bundle-app' + (SFX ? "-v2" : "") + '"></div></td></tr>';
      for (var i = 1; i <= 6; i++) h += '<tr id="bundle_rank_' + i + SFX + '-tr"><td class="labelrc">bundle_rank_' + i + SFX + '</td><td><input type="text" id="bundle_rank_' + i + SFX + '" name="bundle_rank_' + i + SFX + '"></td></tr>';
    }
    document.getElementById("rc").innerHTML = h + "</tbody></table>";
    /* REDCap behaviours the script relies on: hidden "shadow" inputs follow the visible choices; reset link clears a radio */
    $("#rc").on("change", 'input[type=checkbox][name^="__chkn__"]', function () {
      var f = this.name.replace("__chkn__", ""), code = this.getAttribute("code");
      $('input[type=hidden][name="' + f + "___" + code + '"]').val(this.checked ? code : "");
    });
    $("#rc").on("change", 'input[type=radio][name$="___radio"]', function () {
      var f = this.name.replace("___radio", ""); $('input[type=hidden][name="' + f + '"]').val(this.value);
    });
    $("#rc").on("click", "a[data-reset]", function (e) {
      e.preventDefault(); var f = this.getAttribute("data-reset");
      $('input[type=radio][name="' + f + '___radio"]').prop("checked", false); $('input[type=hidden][name="' + f + '"]').val("").trigger("change");
    });
    $("#rc").on("input change", 'input[id^="bundle_rank_"]', paintSaved);
  }
  function savedValues() { var v = []; for (var i = 1; i <= 6; i++) v.push($("#bundle_rank_" + i + SFX).val() || ""); return v; }
  function paintSaved() {
    var box = document.getElementById("saved"), v = savedValues();
    if (!document.getElementById("bundle_rank_1" + SFX)) { box.textContent = "Open page 2."; return; }
    box.innerHTML = ""; v.forEach(function (x, i) {
      var d = document.createElement("div"); d.style.cssText = "font:12px ui-monospace,monospace;padding:3px 0;border-bottom:1px solid #eee";
      d.textContent = "bundle_rank_" + (i + 1) + SFX + ": " + (x || "(empty)"); box.appendChild(d);
    });
  }

  /* ---- loading the script ---- */
  var scriptText = "";
  function inject(text) {
    var s = document.createElement("script"); s.textContent = text; document.body.appendChild(s);
    var m = /var FGS_BUILD\s*=\s*"([^"]+)"/.exec(text), pill = document.getElementById("build");
    pill.textContent = (m ? "loaded: " + KIND + " build " + m[1] : "loaded (no build tag found)") + " | field suffix \"" + SFX + "\" | main meals " + MAIN; pill.className = "pill";
  }
  function page() { try { return parseInt(sessionStorage.getItem("rcmock_page"), 10) || 2; } catch (e) { return 2; } }
  function goto(n) { sessionStorage.setItem("rcmock_page", String(n)); location.reload(); }

  document.getElementById("pick").onclick = function () { document.getElementById("file").click(); };
  document.getElementById("file").onchange = function (e) {
    var f = e.target.files[0]; if (!f) return; var r = new FileReader();
    r.onload = function () { sessionStorage.setItem("rcmock_js", String(r.result)); sessionStorage.setItem("rcmock_name", f.name); location.reload(); }; r.readAsText(f);
  };
  document.getElementById("pg1").onclick = function () { goto(1); };
  document.getElementById("pg2").onclick = function () { goto(2); };
  document.getElementById("clr").onclick = function () { sessionStorage.clear(); location.reload(); };

  /* ---- automated checks ---- */
  function rng(seed) { var a = seed >>> 0; return function () { a |= 0; a = (a + 0x6D2B79F5) | 0; var t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }
  function parseBundles(text) {
    var m = /var BUNDLES\s*=\s*(\[[\s\S]*?\]);\s*\n/.exec(text); if (!m) throw new Error("no BUNDLES line in the script");
    return JSON.parse(m[1]);
  }
  function applyScenario(sc) {
    $('#rc input[type=checkbox]').prop("checked", false); $('#rc input[type=hidden]').val(""); $('#rc input[type=radio]').prop("checked", false);
    (sc.proteins || []).forEach(function (p) { $("#id-__chk__protein_preference" + SFX + "_RC_" + PCODE[p]).prop("checked", true).trigger("change-silent"); $('input[name="protein_preference' + SFX + '___' + PCODE[p] + '"]').val(PCODE[p]); });
    if (sc.diet) { $("#dietary_preference" + SFX + "-" + DCODE[sc.diet]).prop("checked", true); $('input[name="dietary_preference' + SFX + '"]').val(DCODE[sc.diet]); }
    (sc.restrict || []).forEach(function (r) { $("#id-__chk__dietary_restrictions" + SFX + "_RC_" + RCODE[r]).prop("checked", true); $('input[name="dietary_restrictions' + SFX + '___' + RCODE[r] + '"]').val(RCODE[r]); });
    (sc.allergies || []).forEach(function (a) { $("#id-__chk__food_allergies" + SFX + "_RC_" + ACODE[a]).prop("checked", true); $('input[name="food_allergies' + SFX + '___' + ACODE[a] + '"]').val(ACODE[a]); });
    try { if (sc.cook) sessionStorage.setItem(COOKKEY, sc.cook); else sessionStorage.removeItem(COOKKEY); } catch (e) {}
    /* one real change event; the script re-reads every field */
    $("#id-__chk__food_allergies" + SFX + "_RC_1").trigger("change");
  }
  function shown() {
    return $(".fgs-list > .fgs-ranked-item").map(function () {
      var id = parseInt($(this).find("h4").first().text().replace(/\D/g, ""), 10);
      var names = $(this).find(".fgs-recipe").map(function () { return $(this).children("span").eq(1).text().replace(/\s*\((?:\d+ hr(?: \d+ min)?|\d+ min)\)$/, ""); }).get();
      var text = $(this).text(), swaps = $(this).find(".fgs-swap").length;
      return { id: id, names: names, text: text, swaps: swaps };
    }).get();
  }
  var COOKMAX = { a: 30, b: 60, c: Infinity };
  function check(sc, bundles, mode) {
    var problems = [], diet = sc.diet || "none";
    var meat = (sc.proteins || []).some(function (p) { return p !== "vegetarian"; });
    var effDiet = (meat && (diet === "vegetarian" || diet === "vegan")) ? "none" : diet;
    var sel = { proteins: sc.proteins || [], diet: effDiet, restrict: sc.restrict || [], allergies: sc.allergies || [], cook: sc.cook || "" };
    var o = { mode: mode, mainCount: MAIN }, st = E.makeState(sel), views = E.computeViews(bundles, st, o);
    var byId = {}; views.forEach(function (v) { byId[v.bundle.id] = v; });
    var got = shown(), saved = savedValues(), ids = got.map(function (g) { return g.id; });
    if (meat && diet !== "none" && effDiet === "none" && $('input[name="dietary_preference' + SFX + '___radio"]:checked').length) problems.push("meat protein ticked but the " + diet + " radio is still selected");
    var want = Math.min(6, views.length);
    if (views.length === 0) { if (!$(".fgs-empty").length) problems.push("no bundle fits, but the 'team will follow up' message is not shown"); }
    if (ids.length !== want) problems.push("shows " + ids.length + " bundles, expected " + want + " (" + views.length + " eligible)");
    ids.forEach(function (id, i) {
      var v = byId[id]; if (!v) { problems.push("Bundle " + id + " is shown but does not fit this family"); return; }
      var exp = v.mainSurvivors.map(function (r) { return r.name; });
      if (mode === "bundle" && exp.length !== MAIN) problems.push("Bundle " + id + " is partial in the expected data (internal)");
      if (JSON.stringify(got[i].names) !== JSON.stringify(exp)) problems.push("Bundle " + id + " shows [" + got[i].names.join(" | ") + "] expected [" + exp.join(" | ") + "]");
      if (saved[i] !== E.rankFieldValue(v)) problems.push("bundle_rank_" + (i + 1) + " saved '" + saved[i].slice(0, 80) + "…' expected '" + E.rankFieldValue(v).slice(0, 80) + "…'");
      if (/not shown|of \d+ recipes/i.test(got[i].text)) problems.push("Bundle " + id + " shows a 'not shown' / 'X of Y' note");
      if (sel.restrict.indexOf("gluten_free") > -1) { var exSw = v.mainSurvivors.filter(function (r) { return !r.gfStrict && r.gfNote; }).length; if (got[i].swaps !== exSw) problems.push("Bundle " + id + " shows " + got[i].swaps + " GF swap notes, expected " + exSw); }
    });
    for (var i = ids.length; i < 6; i++) if (saved[i] !== "") problems.push("bundle_rank_" + (i + 1) + " should be empty but holds '" + saved[i].slice(0, 60) + "'");
    if (window[GLOB] && window[GLOB].top6 && JSON.stringify(window[GLOB].top6()) !== JSON.stringify(ids)) problems.push(GLOB + ".top6() " + JSON.stringify(window.FGSS.top6()) + " differs from the screen " + JSON.stringify(ids));
    /* each ticked protein that any eligible bundle offers should appear in the list */
    var keys = views.map(function (v) { return E.keyOf(v, st, o); });
    if (ids.length && sel.proteins.length && sel.proteins.length <= 6) sel.proteins.forEach(function (p) {
      var avail = keys.filter(function (k) { return k.pc[p] > 0; }).length;
      var inList = keys.filter(function (k) { return k.pc[p] > 0 && ids.indexOf(k.id) > -1; }).length;
      if (avail && !inList) problems.push("protein '" + p + "' is available in " + avail + " eligible bundles but none is in the list");
    });
    return problems;
  }
  function tick() { return new Promise(function (r) { setTimeout(r, 0); }); }
  async function runChecks() {
    var out = document.getElementById("res");
    if (!window[GLOB]) { out.innerHTML = '<span class="fail">Load an injector file and open page 2 first.</span>'; return; }
    if (!document.getElementById("recipe-bundle-app" + (SFX ? "-v2" : ""))) { out.innerHTML = '<span class="fail">Open page 2 (the ranking page) first.</span>'; return; }
    var bundles, patched;
    try { bundles = parseBundles(scriptText); } catch (e) { out.innerHTML = '<span class="fail">' + e.message + "</span>"; return; }
    E.applyPatches(bundles, window[GLOB].patches || []);
    var mode = document.getElementById("rules").value, n = Math.max(0, parseInt(document.getElementById("nrand").value, 10) || 0), r = rng(20261111);
    var scen = [
      {}, { diet: "vegan" }, { diet: "vegetarian" }, { diet: "pescatarian" }, { proteins: ["chicken", "turkey", "beef"] },
      { proteins: ["chicken", "beef", "fish_shellfish", "vegetarian"], restrict: ["halal"] }, { restrict: ["kosher"] }, { restrict: ["gluten_free"], allergies: ["Milk"] },
      { diet: "vegan", cook: "a" }, { diet: "vegan", proteins: ["chicken"] }, { restrict: ["kosher"], allergies: ["Wheat", "Soy"] }, { allergies: ["Milk", "Eggs", "Wheat", "Soy", "Peanuts", "Tree Nuts", "Fish", "Shellfish", "Sesame"] },
      { proteins: ["vegetarian"], diet: "vegetarian", cook: "b" }, { cook: "a" }, { cook: "c" }
    ];
    function pick(a, p) { return a.filter(function () { return r() < p; }); }
    for (var i = 0; i < n; i++) scen.push({ proteins: r() < 0.4 ? [] : pick(Object.keys(PCODE), 0.4), diet: ["none", "none", "vegetarian", "vegan", "pescatarian"][Math.floor(r() * 5)], restrict: pick(Object.keys(RCODE), 0.25), allergies: pick(Object.keys(ACODE), 0.12), cook: ["", "a", "b", "c"][Math.floor(r() * 4)] });
    var fails = [], t0 = Date.now(), errBefore = errCount, empty = 0;
    for (i = 0; i < scen.length; i++) {
      applyScenario(scen[i]); await tick();
      var p = check(scen[i], bundles, mode);
      if (!shown().length) empty++;
      if (p.length) fails.push({ sc: scen[i], p: p });
    }
    /* reorder test: Move down on #1 must swap the saved fields too */
    var moveMsg = "";
    applyScenario({}); await tick();
    var before = savedValues(), idsBefore = shown().map(function (x) { return x.id; });
    $('.fgs-ranked-item [data-action="down"]').first().trigger("click"); await tick();
    var after = savedValues(), idsAfter = shown().map(function (x) { return x.id; });
    var okMove = idsAfter[0] === idsBefore[1] && idsAfter[1] === idsBefore[0] && after[0] === before[1] && after[1] === before[0];
    moveMsg = okMove ? "Move down swaps screen order and the saved fields: OK" : "Move down did NOT swap consistently";
    if (!okMove) fails.push({ sc: "move down test", p: [moveMsg] });
    await new Promise(function (r2) { setTimeout(r2, 500); });
    var html = '<div><b class="' + (fails.length ? "fail" : "ok") + '">' + (fails.length ? fails.length + " of " + scen.length + " scenarios failed" : "All " + scen.length + " scenarios passed") + "</b> (" + (Date.now() - t0) + " ms; " + empty + " families got no bundle)</div>";
    html += '<div class="' + (okMove ? "ok" : "fail") + '">' + moveMsg + "</div>";
    html += '<div class="' + (errCount > errBefore ? "fail" : "ok") + '">' + (errCount > errBefore ? "The script logged errors - see the console output." : "No script errors logged.") + "</div>";
    if (fails.length) {
      html += '<table class="r"><tr><th>Family</th><th>Problem</th></tr>';
      fails.slice(0, 25).forEach(function (f) { html += "<tr><td>" + (typeof f.sc === "string" ? f.sc : JSON.stringify(f.sc)) + "</td><td>" + f.p.map(function (x) { return x.replace(/</g, "&lt;"); }).join("<br>") + "</td></tr>"; });
      html += "</table>";
      if (fails.length > 25) html += "<div class='sub'>Showing 25 of " + fails.length + ".</div>";
    }
    out.innerHTML = html; window.__RC_RESULT = { total: scen.length, failed: fails.length, empty: empty, move: okMove, errors: errCount - errBefore, fails: fails.slice(0, 10) };
  }
  document.getElementById("run").onclick = runChecks;

  /* ---- boot ---- */
  try { scriptText = sessionStorage.getItem("rcmock_js") || ""; } catch (e) {}
  if (scriptText) readMeta(scriptText);
  var n = page();
  document.getElementById("pg" + n).classList.add("on");
  buildPage(n); paintSaved();
  if (scriptText) inject(scriptText); else document.getElementById("log").textContent = "Choose your injector .js file to begin.";
})();