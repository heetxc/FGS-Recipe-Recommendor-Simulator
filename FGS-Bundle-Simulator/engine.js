/* ============================================================================
   FGS recipe-bundle ENGINE  (used by the simulator page; no DOM, no SheetJS here)

   1) parseRows()   : spreadsheet rows  ->  bundles  (tags -> diet / halal / kosher / GF / allergens, time -> minutes)
   2) applyPatches(): reviewed corrections (data/patches/*.json), same add-only semantics as the REDCap scripts' DATA_PATCHES
   3) recipeOk / computeViews / selectTop6 : a line-for-line port of the REDCap JS Injector logic
        mode "target"  = FGS-S rules  (cook time is display-only, bonus meals hidden, ranking counts MAIN meals,
                         quicker bundles first as the LAST sort priority). Used for FGS-S and the planned FGS 2.0 update.
        mode "legacy"  = what the current FGS 2.0 3-meal / 4-meal scripts still do (cook time filters recipes,
                         ranking counts main + bonus survivors, no speed tie-break).
   ========================================================================== */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api; else root.FGSEngine = api;
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ---------- field <-> code maps (identical to the instruments / scripts) ---------- */
  var PROTEINS = [
    { key: "chicken", label: "Chicken" }, { key: "turkey", label: "Turkey" }, { key: "beef", label: "Beef" },
    { key: "fish_shellfish", label: "Fish / Shellfish" }, { key: "vegetarian", label: "Vegetarian (no meat)" }
  ];
  var DIETS = [
    { key: "none", label: "No dietary preference" }, { key: "vegetarian", label: "Vegetarian" },
    { key: "vegan", label: "Vegan" }, { key: "pescatarian", label: "Pescatarian" }
  ];
  var RESTRICTIONS = [{ key: "halal", label: "Halal" }, { key: "kosher", label: "Kosher" }, { key: "gluten_free", label: "Gluten free" }];
  var ALLERGENS = ["Milk", "Eggs", "Wheat", "Soy", "Peanuts", "Tree Nuts", "Fish", "Shellfish", "Sesame"];
  var COOK_CODES = [
    { key: "", label: "(not answered)", max: Infinity },
    { key: "a", label: "30 minutes or less", max: 30 },
    { key: "b", label: "More than 30 min, under 1 hour", max: 60 },
    { key: "c", label: "1 hour or more", max: Infinity }
  ];

  /* ====================================================================== */
  /*  1. SPREADSHEET PARSING                                                */
  /* ====================================================================== */
  function str(x) { return x === null || x === undefined ? "" : String(x); }
  function normName(x) { return String(x).replace(/\s+/g, "").toLowerCase(); }

  /* "Blackened Chicken Mac and Cheese (HOAH) -p" -> "Blackened Chicken Mac and Cheese" */
  function cleanName(raw) {
    var s = str(raw).replace(/\s+/g, " ").trim();
    s = s.replace(/\((B|S|T)\)/g, function (m, c) { return "(" + { B: "Beef", S: "Shrimp", T: "Turkey" }[c] + ")"; });
    s = s.replace(/\s*\((?:HOAH|IHF|FGS[^)]*|TLD[^)]*|[CFP])\)/gi, "");           /* source / testing codes */
    s = s.replace(/\s*[-–]\s*[cfp]\s*$/i, "").replace(/\s+[cfp]\s*$/i, "");      /* trailing -c / -f / -p marker */
    s = s.replace(/\s*[-–]\s*$/, "").replace(/\s+/g, " ").trim();
    return s;
  }

  function mapAllergen(tok) {
    var t = tok.toLowerCase().trim();
    if (!t) return null;
    if (/^(milk|dairy|lactose)/.test(t)) return "Milk";
    if (/^egg/.test(t)) return "Eggs";
    if (/^(wheat|gluten)/.test(t)) return "Wheat";
    if (/^soy/.test(t)) return "Soy";
    if (/^peanut/.test(t)) return "Peanuts";
    if (/^(tree\s*nut|nuts?$|almond|walnut|cashew|pecan)/.test(t)) return "Tree Nuts";
    if (/^shellfish|^shrimp/.test(t)) return "Shellfish";
    if (/^fish/.test(t)) return "Fish";
    if (/^sesame/.test(t)) return "Sesame";
    return null;
  }

  /* total minutes = every "N min..." + every "N hour(s)" in the Prep/Cook cell */
  function timeMinutes(t) {
    t = str(t);
    var n = 0, m, re = /(\d+)\s*min/gi;
    while ((m = re.exec(t))) n += parseInt(m[1], 10);
    re = /(\d+(?:\.\d+)?)\s*hour/gi;
    while ((m = re.exec(t))) n += Math.round(parseFloat(m[1]) * 60);
    return n || null;
  }

  var PROT_KW = [
    ["chicken", /chicken/g], ["turkey", /turkey/g],
    ["beef", /\bbeef\b|steak|sirloin|brisket|ground beef/g],
    ["fish_shellfish", /salmon|tuna|\bcod\b|tilapia|shrimp|\bfish\b|haddock|mahi|trout|crab|scallop/g]
  ];
  /* protein = what the recipe is built around. Name keywords win; otherwise the most-mentioned meat in the ingredients;
     a recipe with no meat is "vegetarian" (this is the protein bucket, independent of the Veg tag). */
  function inferProtein(name, ingredients, vegTag) {
    var n = name.toLowerCase(), i = str(ingredients).toLowerCase(), k;
    for (k = 0; k < PROT_KW.length; k++) if (new RegExp(PROT_KW[k][1].source).test(n)) return PROT_KW[k][0];
    if (vegTag) return "vegetarian";
    var best = null, bestN = 0;
    PROT_KW.forEach(function (p) {
      var c = (i.match(p[1]) || []).length;
      if (p[0] === "chicken") c = (i.replace(/chicken (?:broth|stock|bouillon)/g, "").match(p[1]) || []).length;   /* broth is not the protein */
      if (c > bestN) { best = p[0]; bestN = c; }
    });
    return best || "vegetarian";
  }

  /* Ingredient scan: allergens that are in the ingredient list but missing from the T9 tag. Optional items and "vegan ..." items are skipped. */
  var ALLERGEN_RULES = [
    ["Milk", /\b(milk|cheese|mozzarella|parmesan|feta|cotija|yogurt|butter|cream|ricotta|cheddar|half[- ]and[- ]half)\b/],
    ["Eggs", /\beggs?\b|liquid egg/],
    ["Wheat", /\b(wheat|flour|tortillas?|bread|breadcrumbs?|bread crumbs|panko|pasta|penne|rotini|noodles?|rolls?|buns?|pita|couscous|orzo|spaghetti|macaroni|bagel|crackers?|ramen|farro|barley|bulgur|corn flakes?)\b/],
    ["Soy", /\b(soy|tofu|edamame|tempeh|miso)\b/],
    ["Sesame", /\bsesame\b|tahini/],
    ["Peanuts", /\bpeanuts?\b/],
    ["Tree Nuts", /\b(almonds?|walnuts?|cashews?|pecans?|pistachios?|hazelnuts?)\b/],
    ["Fish", /\b(fish|salmon|tuna|cod|haddock|tilapia|trout|anchov\w+)\b/],
    ["Shellfish", /\b(shrimp|crab|lobster|scallops?|clams?)\b/]
  ];
  function scanAllergens(ingredients) {
    var found = {};
    str(ingredients).split(/\r?\n/).forEach(function (line) {
      var l = line.toLowerCase();
      if (/optional|\bvegan\b/.test(l)) return;
      l = l.replace(/gluten[- ]free/g, "");
      ALLERGEN_RULES.forEach(function (rule) { if (rule[1].test(l)) found[rule[0]] = true; });
    });
    return Object.keys(found);
  }
  var OTHER_MEAT = /\b(pork|bacon|ham|sausage|lamb|pepperoni|salami|veal|bologna|prosciutto|meatballs?)\b/;

  function parseTags(tags) {
    tags = str(tags);
    var warnings = [];
    /* GF swap note: text after "GF:" / "GF-" */
    var gfNote = null, head = tags, mm = /\bGF\s*[:\-]\s*/i.exec(tags);
    if (mm) { gfNote = tags.slice(mm.index + mm[0].length).replace(/\s+/g, " ").trim() || null; head = tags.slice(0, mm.index); }
    /* allergens: T9 (a, b, c)  |  T9(milk)  |  T9 soy */
    var allergens = [], t9re = /T9\s*\(([^)]*)\)/gi, m;
    while ((m = t9re.exec(head))) {
      m[1].split(/[,;&]|\band\b/i).forEach(function (a) {
        var c = mapAllergen(a);
        if (c) { if (allergens.indexOf(c) < 0) allergens.push(c); }
        else if (a.trim()) warnings.push("unrecognised allergen \"" + a.trim() + "\" in T9 tag");
      });
    }
    var rest = head.replace(/T9\s*\([^)]*\)/gi, " ").replace(/T9\s*\w+/gi, function (x) {
      var c = mapAllergen(x.replace(/T9\s*/i, ""));
      if (c && allergens.indexOf(c) < 0) allergens.push(c);
      return " ";
    });
    var tok = {};
    rest.split(/[,\s.;]+/).forEach(function (x) { if (x) tok[x.toLowerCase()] = true; });
    allergens.sort();
    var isVeganTag = !!(tok.v || tok.vegan);
    var gfStrict = !!tok.gf;
    return {
      vegan: isVeganTag, vegetarian: !!(tok.veg || tok.vegetarian) || isVeganTag,
      halal: !!tok.h, kosher: !!tok.k, gfStrict: gfStrict, gfOk: gfStrict || !!gfNote, gfNote: gfNote,
      allergens: allergens, unknown: !!tok.q, warnings: warnings, tokens: Object.keys(tok)
    };
  }

  /* rows = array of arrays (sheet_to_json header:1). opts.labels = {id:"label"} fallback labels. */
  function parseRows(rows, opts) {
    opts = opts || {};
    var hdr = -1, cTags = 4, cTime = 5, cIng = 2, i, j;
    for (i = 0; i < Math.min(rows.length, 6); i++) {
      var hit = false;
      for (j = 0; j < (rows[i] || []).length; j++) if (/^\s*tags?\s*$/i.test(str(rows[i][j]))) { hit = true; cTags = j; }
      if (hit) {
        hdr = i;
        rows[i].forEach(function (c, ci) {
          var s = str(c);
          if (/prep|cook|time/i.test(s) && !/kcal|price/i.test(s)) cTime = ci;
          if (/^\s*ingredients\s*$/i.test(s)) cIng = ci;
        });
        break;
      }
    }
    var bundles = [], cur = null, problems = [];
    rows.forEach(function (r, ri) {
      if (!r) return;
      var a = str(r[0]).trim();
      if (!a) return;
      var m = /^Bundle\s*(\d+)\s*(.*)$/i.exec(a);
      var tagCell = str(r[cTags]).trim();
      if (m && (!tagCell || /^tags?$/i.test(tagCell))) {
        cur = { id: parseInt(m[1], 10), subtitle: m[2].replace(/^[\s(]+|[\s)]+$/g, ""), recipes: [], sheetRow: ri + 1 };
        bundles.push(cur); return;
      }
      if (!cur) return;
      if (ri === hdr) return;
      var f = parseTags(tagCell);
      var name = cleanName(a);
      var scanned = opts.scanIngredients === false ? [] : scanAllergens(r[cIng]);
      var al = f.allergens.slice(); scanned.forEach(function (x) { if (al.indexOf(x) < 0) al.push(x); }); al.sort();
      /* "V" without "Veg" is read as vegan, unless the allergens show milk or eggs */
      var vegan = f.vegan && !(f.tokens.indexOf("veg") < 0 && f.tokens.indexOf("vegan") < 0 && (al.indexOf("Milk") > -1 || al.indexOf("Eggs") > -1));
      f.vegan = vegan; f.vegetarian = f.vegetarian || vegan;
      var rec = {
        name: name, rawName: a, protein: inferProtein(name, r[cIng], f.vegetarian), meatInIngredients: (inferProtein(name, r[cIng], false) !== "vegetarian") || OTHER_MEAT.test(str(r[cIng]).toLowerCase()),
        scannedAllergens: scanned.filter(function (x) { return f.allergens.indexOf(x) < 0; }),
        vegetarian: f.vegetarian, vegan: f.vegan, halal: f.halal, kosher: f.kosher,
        gfStrict: f.gfStrict, gfOk: f.gfOk, gfNote: f.gfNote, allergens: al,
        timeMin: timeMinutes(r[cTime]), rawTags: tagCell, rawTime: str(r[cTime]).trim(), sheetRow: ri + 1
      };
      if (f.warnings.length) rec.warnings = f.warnings;
      cur.recipes.push(rec);
    });
    bundles.forEach(function (b) {
      if (!b.subtitle && opts.labels && opts.labels[b.id]) b.subtitle = opts.labels[b.id];
      if (!b.subtitle) b.subtitle = deriveLabel(b);
    });
    return bundles;
  }

  function deriveLabel(b) {
    var c = {};
    b.recipes.slice(0, 3).forEach(function (r) { c[r.protein] = (c[r.protein] || 0) + 1; });
    var top = Object.keys(c).sort(function (x, y) { return c[y] - c[x]; })[0] || "mixed";
    return { chicken: "Chicken", turkey: "Turkey", beef: "Beef", fish_shellfish: "Fish", vegetarian: "Vegetarian" }[top] || "Mixed";
  }

  /* ====================================================================== */
  /*  2. PATCHES  (same semantics as the REDCap scripts' DATA_PATCHES)       */
  /* ====================================================================== */
  /* entry: {name, vegetarian?:true, vegan?:true, allergensAdd?:[], protein?:"beef", why:""}
     add-only: can add allergens, upgrade to vegetarian/vegan, or set the protein bucket. Never removes a restriction. */
  function applyPatches(bundles, patches) {
    var used = {};
    (patches || []).forEach(function (p, i) { used[i] = 0; });
    bundles.forEach(function (b) {
      b.recipes.forEach(function (r) {
        (patches || []).forEach(function (p, i) {
          if (normName(p.name) !== normName(r.name)) return;
          if (p.vegan === true) { r.vegan = true; r.vegetarian = true; }
          if (p.vegetarian === true) r.vegetarian = true;
          if (p.protein) r.protein = p.protein;
          (p.allergensAdd || []).forEach(function (a) { if (r.allergens.indexOf(a) < 0) r.allergens.push(a); });
          r.patched = (r.patched || []).concat([p.why || "patch"]);
          used[i]++;
        });
      });
    });
    return used;   /* patch index -> number of recipe slots it touched (0 = no longer matches anything) */
  }

  /* ====================================================================== */
  /*  3. FILTERS + RANKING  (port of the REDCap script)                     */
  /* ====================================================================== */
  function emptyState() {
    return { protein: {}, anyProtein: true, diet: "none", restrict: {}, allergy: {}, cookCode: "", cookMaxMin: Infinity, cookTimeSet: false };
  }
  function makeState(sel) {
    var st = emptyState();
    (sel.proteins || []).forEach(function (p) { st.protein[p] = true; });
    st.anyProtein = Object.keys(st.protein).length === 0;
    st.diet = sel.diet || "none";
    (sel.restrict || []).forEach(function (r) { st.restrict[r] = true; });
    (sel.allergies || []).forEach(function (a) { st.allergy[a] = true; });
    COOK_CODES.forEach(function (c) { if (c.key && c.key === sel.cook) { st.cookCode = c.key; st.cookMaxMin = c.max; st.cookTimeSet = true; } });
    return st;
  }

  /* reasons a recipe is dropped for this family (empty = it stays). opts.mode "legacy" adds the cook-time filter. */
  function whyOut(r, st, opts) {
    var why = [];
    if (st.diet === "vegetarian" && !r.vegetarian) why.push("not vegetarian");
    if (st.diet === "vegan" && !r.vegan) why.push("not vegan");
    if (st.diet === "pescatarian" && !(r.protein === "vegetarian" || r.protein === "fish_shellfish")) why.push("not pescatarian (meat)");
    if (st.restrict.halal && !r.halal) why.push("not halal");
    if (st.restrict.kosher && !r.kosher) why.push("not kosher");
    if (st.restrict.gluten_free && !r.gfOk) why.push("not gluten-free / no GF swap");
    for (var i = 0; i < r.allergens.length; i++) if (st.allergy[r.allergens[i]]) why.push("contains " + r.allergens[i]);
    if (opts && opts.mode === "legacy" && st.cookTimeSet && r.timeMin !== null && r.timeMin !== undefined && r.timeMin > st.cookMaxMin)
      why.push("cook time " + r.timeMin + " min > " + st.cookMaxMin + " (legacy filter)");
    return why;
  }
  function recipeOk(r, st, opts) { return whyOut(r, st, opts).length === 0; }

  function computeViews(bundles, st, opts) {
    var mainCount = opts.mainCount;
    return bundles.map(function (b) {
      var survivors = b.recipes.filter(function (r) { return recipeOk(r, st, opts); });
      var mainSurvivors = survivors.filter(function (r) { var i = b.recipes.indexOf(r); return i > -1 && i < mainCount; });
      var mainTotal = Math.min(mainCount, b.recipes.length);
      return { bundle: b, survivors: survivors, mainSurvivors: mainSurvivors, mainTotal: mainTotal, hidden: mainTotal - mainSurvivors.length, state: st };
    }).filter(function (v) { return v.mainSurvivors.length > 0; });   /* a bundle with no surviving MAIN meal is not offered */
  }

  function bundleSpeed(v) {
    var t = v.mainSurvivors.filter(function (r) { return r.timeMin !== null && r.timeMin !== undefined && r.timeMin > 0; })
                           .map(function (r) { return r.timeMin; });
    if (!t.length) return Infinity;
    return t.reduce(function (a, b) { return a + b; }, 0) / t.length;
  }

  function keyOf(v, st, opts) {
    var legacy = opts.mode === "legacy", pool = legacy ? v.survivors : v.mainSurvivors, pc = {}, mc = 0;
    Object.keys(st.protein).forEach(function (p) {
      pc[p] = pool.filter(function (r) { return r.protein === p; }).length; mc += pc[p];
    });
    var k = { id: v.bundle.id, n: pool.length, mc: mc, pc: pc, isMatch: (st.anyProtein || mc > 0) ? 1 : 0 };
    if (!legacy) k.speed = bundleSpeed(v);
    return k;
  }

  /* jitter = {bundleId: random 0..1}  (REDCap draws this once per page load). Returns chosen keys in DISPLAY order. */
  function selectTop6(views, st, jitter, opts) {
    var legacy = opts.mode === "legacy";
    var keys = views.map(function (v) { return keyOf(v, st, opts); });
    var picked = [], seen = {};
    var plistAll = Object.keys(st.protein), perRound = Math.max(1, plistAll.length);
    function take(k, seq, why) {
      if (picked.length < 6 && !seen[k.id]) {
        seen[k.id] = true; k.seq = seq; k.why = why;
        if (!legacy) k.grp = (seq >= 100) ? 1000 : Math.floor(seq / perRound);
        picked.push(k); return true;
      }
      return false;
    }
    if (!st.anyProtein) {
      var plist = Object.keys(st.protein);
      var queues = plist.map(function (p) {
        return keys.filter(function (k) { return k.pc[p] > 0; })
                   .sort(function (a, b) { return (b.n - a.n) || (b.pc[p] - a.pc[p]) || (jitter[b.id] - jitter[a.id]); });
      });
      var pos = plist.map(function () { return 0; }), seq = 0, progress = true;
      while (picked.length < 6 && progress) {
        progress = false;
        for (var qi = 0; qi < queues.length && picked.length < 6; qi++) {
          while (pos[qi] < queues[qi].length && seen[queues[qi][pos[qi]].id]) pos[qi]++;
          if (pos[qi] < queues[qi].length) {
            var round = Math.floor(seq / perRound) + 1;
            take(queues[qi][pos[qi]], seq++, "protein round " + round + " (" + plist[qi] + ")"); progress = true;
          }
        }
      }
    }
    var rest = keys.filter(function (k) { return !seen[k.id]; })
                   .sort(function (a, b) { return (b.isMatch - a.isMatch) || (b.n - a.n) || (b.mc - a.mc) || (jitter[b.id] - jitter[a.id]); });
    var seq2 = 100;
    for (var i = 0; i < rest.length && picked.length < 6; i++) take(rest[i], seq2++, st.anyProtein ? "richest first (no protein preference)" : "fill (richest remaining)");

    var speedOf = function (k) { return k.speed === Infinity ? 1e9 : k.speed; };
    if (legacy) return picked.sort(function (a, b) { return (b.n - a.n) || (a.seq - b.seq); });
    return picked.sort(function (a, b) { return (b.n - a.n) || (a.grp - b.grp) || (speedOf(a) - speedOf(b)) || (a.seq - b.seq); });
  }

  function rankFieldValue(v) {
    return "Bundle " + v.bundle.id + " (" + v.bundle.subtitle + "): " + v.survivors.map(function (r) { return r.name; }).join("; ");
  }

  /* heading must not contradict what is shown (script's displayLabel; kept for the all-bundles view) */
  function displayLabel(v) {
    var d = v.state && v.state.diet;
    if (d === "vegan" || d === "vegetarian") return v.mainSurvivors.every(function (r) { return r.vegan; }) ? "Vegan" : "Vegetarian";
    return v.bundle.subtitle;
  }

  function fmtTime(m) {
    if (m < 60) return m + " min";
    var h = Math.floor(m / 60), r = m % 60;
    return h + " hr" + (r ? " " + r + " min" : "");
  }

  /* One call = everything the participant (and the researcher export) would get. */
  function simulate(bundles, sel, jitter, opts) {
    var st = makeState(sel);
    var views = computeViews(bundles, st, opts);
    var keys = selectTop6(views, st, jitter, opts);
    var byId = {}; views.forEach(function (v) { byId[v.bundle.id] = v; });
    var all = views.map(function (v) { return keyOf(v, st, opts); });
    var ranked = keys.map(function (k, i) {
      var v = byId[k.id];
      return { rank: i + 1, key: k, view: v, saved: rankFieldValue(v), speed: opts.mode === "legacy" ? bundleSpeed(v) : k.speed };
    });
    var fields = []; for (var i = 0; i < 6; i++) fields.push(ranked[i] ? ranked[i].saved : "");
    var coverage = {};
    Object.keys(st.protein).forEach(function (p) { coverage[p] = keys.filter(function (k) { return k.pc[p] > 0; }).length; });
    return { state: st, views: views, ranked: ranked, eligibleCount: views.length, totalBundles: bundles.length, fields: fields, coverage: coverage, allKeys: all };
  }

  /* ---------- per-bundle audit used by the "All bundles" tab ---------- */
  function summarizeBundle(b, mainCount) {
    var mains = b.recipes.slice(0, mainCount), s = { proteins: {}, vegan: 0, vegetarian: 0, halal: 0, kosher: 0, gfOk: 0, allergens: {} };
    mains.forEach(function (r) {
      s.proteins[r.protein] = (s.proteins[r.protein] || 0) + 1;
      if (r.vegan) s.vegan++; if (r.vegetarian) s.vegetarian++; if (r.halal) s.halal++; if (r.kosher) s.kosher++; if (r.gfOk) s.gfOk++;
      r.allergens.forEach(function (a) { s.allergens[a] = (s.allergens[a] || 0) + 1; });
    });
    s.mainCount = mains.length;
    s.allVegan = mains.length > 0 && s.vegan === mains.length;
    s.allVegetarian = mains.length > 0 && s.vegetarian === mains.length;
    var t = mains.filter(function (r) { return r.timeMin; }).map(function (r) { return r.timeMin; });
    s.avgMainMin = t.length ? t.reduce(function (a, c) { return a + c; }, 0) / t.length : null;
    return s;
  }

  /* sanity checks on the parsed sheet, shown in the QC tab */
  function qcChecks(bundles, mainCount) {
    var out = [];
    bundles.forEach(function (b) {
      b.recipes.forEach(function (r, i) {
        var where = "Bundle " + b.id + " slot " + (i + 1) + " (" + r.name + ")";
        if (!r.rawTags) out.push({ level: "warn", where: where, msg: "no tags in the sheet (treated as: not vegetarian, not halal, not kosher, not GF, no allergens)" });
        if (!r.timeMin) out.push({ level: "info", where: where, msg: "no readable time (\"" + r.rawTime + "\"): shown without a time and ignored in the speed sort" });
        if (r.vegan && !r.vegetarian) out.push({ level: "warn", where: where, msg: "vegan but not vegetarian" });
        if (r.vegetarian && r.meatInIngredients) out.push({ level: "warn", where: where, msg: "tagged vegetarian/vegan but a meat word appears in the name or ingredients - check" });
        if (r.scannedAllergens && r.scannedAllergens.length) out.push({ level: "info", where: where, msg: "allergen(s) " + r.scannedAllergens.join(", ") + " found in the ingredients but not in the T9 tag (added automatically)" });
        if (r.protein === "vegetarian" && !r.vegetarian && !/^\s*$/.test(r.name)) out.push({ level: "info", where: where, msg: "no meat found but not tagged Veg (so it is hidden from Vegetarian/Vegan families)" });
        if (r.gfStrict === false && r.gfOk && !r.gfNote) out.push({ level: "warn", where: where, msg: "GF flag inconsistent" });
        (r.warnings || []).forEach(function (w) { out.push({ level: "warn", where: where, msg: w }); });
        if (r.timeMin && r.timeMin >= 180 && i < mainCount) out.push({ level: "info", where: where, msg: "very long time (" + fmtTime(r.timeMin) + ") - displayed to participants; counts toward the bundle's speed sort" });
      });
      if (b.recipes.length < mainCount) out.push({ level: "error", where: "Bundle " + b.id, msg: "only " + b.recipes.length + " recipes but " + mainCount + " main meals are expected" });
    });
    var ids = {}; bundles.forEach(function (b) { if (ids[b.id]) out.push({ level: "error", where: "Bundle " + b.id, msg: "appears twice in the sheet" }); ids[b.id] = 1; });
    return out;
  }

  return {
    PROTEINS: PROTEINS, DIETS: DIETS, RESTRICTIONS: RESTRICTIONS, ALLERGENS: ALLERGENS, COOK_CODES: COOK_CODES,
    parseRows: parseRows, scanAllergens: scanAllergens, parseTags: parseTags, cleanName: cleanName, timeMinutes: timeMinutes, inferProtein: inferProtein,
    applyPatches: applyPatches, makeState: makeState, whyOut: whyOut, recipeOk: recipeOk, computeViews: computeViews,
    bundleSpeed: bundleSpeed, keyOf: keyOf, selectTop6: selectTop6, rankFieldValue: rankFieldValue, displayLabel: displayLabel,
    fmtTime: fmtTime, simulate: simulate, summarizeBundle: summarizeBundle, qcChecks: qcChecks, normName: normName
  };
});
