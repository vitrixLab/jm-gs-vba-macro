/* assets/graphify-config.js — B-Step 2: dataset presets. Charting only, no accounting math.
   CHART_COLUMN default mirrors Graphify.bas ("I"). GL-shape demo is point COUNT only (552). */
window.GRAPHIFY_CONFIG = {
  chartColumn: "I",
  minNumeric: 2,
  datasets: {
    sample: { label: "Values", source: "column I", values: [23, 45, 12, 78, 34, 56, 29, 80, 15, 62] },
    glshape: { label: "GL 46x12 shape (count demo)", source: "auto (552 pts)", count: 552 }
  }
};
