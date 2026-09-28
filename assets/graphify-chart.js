/* assets/graphify-chart.js — B-Step 2: render/reuse mirroring Graphify.bas.
   Resolution: selection (dataset switch) -> column I default -> auto -> empty state.
   Reuses window.GRAPHIFY_CHART by id (myChart), categories R1..Rn. */
(function () {
  var chart = null;
  function valuesFor(key) {
    var cfg = window.GRAPHIFY_CONFIG.datasets[key] || window.GRAPHIFY_CONFIG.datasets.sample;
    if (cfg.values) return { values: cfg.values.slice(), source: cfg.source };
    var n = cfg.count || 0, out = [];
    for (var i = 0; i < n; i++) out.push((i % 12) + 1); // shape/count demo only
    return { values: out, source: cfg.source };
  }
  function render(key) {
    var page = "graphify.html";
    try {
      var picked = valuesFor(key);
      var nums = picked.values.filter(function (v) { return typeof v === "number" && isFinite(v); });
      if (nums.length < window.GRAPHIFY_CONFIG.minNumeric) {
        GraphifyAnalytics.logFault({ page: page, code: "NO_NUMERIC_DATA", detail: "dataset=" + key, source: picked.source });
        GraphifyAnalytics.showEmpty("No numeric data found. Select a column and run Graphify again.");
        return;
      }
      var ctx = document.getElementById("myChart").getContext("2d");
      var labels = nums.map(function (_, i) { return "R" + (i + 1); });
      GraphifyAnalytics.hideEmpty();
      if (chart) { chart.data.labels = labels; chart.data.datasets[0].data = nums; chart.update(); }
      else {
        chart = new Chart(ctx, {
          type: "bar",
          data: { labels: labels, datasets: [{ label: "Values", data: nums,
            backgroundColor: "rgba(54, 162, 235, 0.6)", borderColor: "rgba(54, 162, 235, 1)", borderWidth: 1 }] },
          options: { responsive: true,
            plugins: { title: { display: true, text: "Graphify - Column Chart" }, legend: { display: false } },
            scales: { y: { beginAtZero: true, title: { display: true, text: "Value" } },
                      x: { title: { display: true, text: "Row" } } } }
        });
        window.GRAPHIFY_CHART = chart;
      }
      document.getElementById("gs-source").textContent = "Source: " + picked.source + " (" + nums.length + " pts)";
      GraphifyAnalytics.logChartRender({ page: page, source: picked.source, nPts: nums.length, code: "RENDER_OK" });
    } catch (e) {
      GraphifyAnalytics.logFault({ page: page, code: "RENDER_FAILED", detail: String((e && e.message) || e), source: key });
    }
  }
  function boot() {
    var sel = document.getElementById("gs-dataset");
    GraphifyAnalytics.logPageView("graphify.html");
    render(sel ? sel.value : "sample");
    if (sel) sel.addEventListener("change", function () {
      GraphifyAnalytics.logAction("dataset-switch", { dataset: sel.value }, "graphify.html");
      render(sel.value);
    });
  }
  window.GraphifyChart = { render: render, boot: boot };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})();
