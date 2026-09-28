/* assets/graphify-analytics.js — B-Step 3: fault + analytics log.
   Vanilla, no dependency. Ring cap 200, localStorage persist, console mirror.
   Codes: NOT_WORKSHEET | NO_NUMERIC_DATA | NO_NUMERIC_BLOCK | RENDER_FAILED | RENDER_OK.
   Payload counts only — never workbook values. */
(function () {
  var KEY = "graphify.faults.v1", CAP = 200;
  function load() { try { return JSON.parse(localStorage.getItem(KEY) || "[]"); } catch (e) { return []; } }
  function save(ring) { try { localStorage.setItem(KEY, JSON.stringify(ring.slice(-CAP))); } catch (e) {} }
  function paint() {
    var body = document.getElementById("gs-fault-body");
    if (!body) return;
    var ring = load();
    if (!ring.length) { body.innerHTML = '<tr><td colspan="5">No faults recorded.</td></tr>'; return; }
    var debug = /[?&]debug=1/.test(location.search);
    body.innerHTML = ring.slice(-20).reverse().map(function (e) {
      return "<tr><td>" + e.t + "</td><td>" + e.page + "</td><td><code>" + e.code + "</code></td>" +
        "<td>" + (debug ? e.detail : String(e.detail).slice(0, 80)) + "</td><td>" + (e.source || "") + "</td></tr>";
    }).join("");
  }
  function push(entry) {
    var ring = load();
    entry.t = new Date().toISOString();
    ring.push(entry);
    save(ring); paint();
    if (window.console) console.log("[graphify]", entry.code, entry.page, entry.detail || "");
  }
  window.GraphifyAnalytics = {
    logPageView: function (page) { push({ page: page, code: "RENDER_OK", detail: "page_view", source: "" }); paint(); },
    logChartRender: function (o) { push({ page: o.page, code: o.code || "RENDER_OK", detail: "nPts=" + o.nPts, source: o.source }); },
    logFault: function (o) { push({ page: o.page, code: o.code, detail: o.detail || "", source: o.source || "" }); },
    logAction: function (action, meta, page) { push({ page: page, code: "RENDER_OK", detail: "action=" + action + " " + JSON.stringify(meta || {}), source: "" }); },
    showEmpty: function (msg) { var el = document.getElementById("gs-empty"); if (el) { el.hidden = false; el.textContent = msg; } },
    hideEmpty: function () { var el = document.getElementById("gs-empty"); if (el) { el.hidden = true; el.textContent = ""; } }
  };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", paint);
  else paint();
})();
