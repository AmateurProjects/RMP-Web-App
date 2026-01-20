/* global require */

require([
  "esri/Map",
  "esri/views/MapView",
  "esri/layers/FeatureLayer",
  "esri/layers/GraphicsLayer",
  "esri/widgets/Sketch",
  "esri/Graphic"
], function (Map, MapView, FeatureLayer, GraphicsLayer, Sketch, Graphic) {

  // ---------- DOM ----------
  const modeSelect = document.getElementById("modeSelect");
  const selectionLayerSelect = document.getElementById("selectionLayerSelect");
  const selectModeControls = document.getElementById("selectModeControls");
  const drawModeControls = document.getElementById("drawModeControls");

  const drawBtn = document.getElementById("drawBtn");
  const stopDrawBtn = document.getElementById("stopDrawBtn");
  const runBtn = document.getElementById("runBtn");
  const clearBtn = document.getElementById("clearBtn");
  const exportAllBtn = document.getElementById("exportAllBtn");

  const statusEl = document.getElementById("status");
  const resultsEl = document.getElementById("results");
  const layerListEl = document.getElementById("layerList");

  function setStatus(msg) { statusEl.textContent = "Status: " + msg; }

  // ---------- State ----------
  let config = null;

  let view = null;
  let selectionGeom = null;

  let selectionLayers = []; // { cfg, layer }
  let activeSelectionLayer = null; // FeatureLayer
  let activeSelectionLayerView = null;
  let activeHighlightHandle = null;

  let drawLayer = null;
  let sketch = null;

  let lastReportRowsByLayer = []; // for export-all

  // ---------- Helpers ----------
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;"
    }[c]));
  }

  function isFeatureServerRoot(url) {
    // ends with /FeatureServer (no trailing /0 etc.)
    return /\/FeatureServer\/?$/.test(url);
  }

  async function fetchJson(url) {
    const res = await fetch(url, { credentials: "omit" });
    if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText} for ${url}`);
    return res.json();
  }

  async function expandServiceToSublayers(serviceUrl) {
    // Returns array of { title, url } for each sublayer
    const pjsonUrl = serviceUrl.replace(/\/$/, "") + "?f=pjson";
    const info = await fetchJson(pjsonUrl);
    const layers = (info && info.layers) ? info.layers : [];
    return layers.map(l => ({
      title: `${info.serviceDescription ? info.serviceDescription + " - " : ""}${l.name}`.trim() || l.name || `Layer ${l.id}`,
      url: serviceUrl.replace(/\/$/, "") + "/" + l.id
    }));
  }

  function clearHighlight() {
    if (activeHighlightHandle) {
      activeHighlightHandle.remove();
      activeHighlightHandle = null;
    }
  }

  function clearAll() {
    selectionGeom = null;
    clearHighlight();
    resultsEl.innerHTML = "";
    exportAllBtn.disabled = true;
    lastReportRowsByLayer = [];

    if (drawLayer) drawLayer.removeAll();

    runBtn.disabled = true;
    setStatus("cleared");
  }

  function setGeometryFromSelection(geom) {
    selectionGeom = geom || null;
    runBtn.disabled = !selectionGeom;
  }

  function setMode(mode) {
    if (mode === "select") {
      selectModeControls.classList.remove("hidden");
      drawModeControls.classList.add("hidden");
      // stop sketch if running
      if (sketch) sketch.cancel();
      setStatus("select mode: click a polygon");
    } else {
      selectModeControls.classList.add("hidden");
      drawModeControls.classList.remove("hidden");
      clearHighlight();
      setStatus("draw mode: draw a polygon");
    }
    // keep current selectionGeom if user switches modes intentionally
  }

  function renderConfiguredLayerList() {
    const lines = [];

    lines.push("Selection layers:");
    (config.selectionLayers || []).forEach(l => lines.push("  - " + l.title));

    lines.push("");
    lines.push("Report layers:");
    (config.reportLayers || []).forEach(l => lines.push("  - " + l.title));

    layerListEl.textContent = lines.join("\n");
  }

  // ---------- Report rendering ----------
  function renderResults(cardsHtml) {
    resultsEl.innerHTML = cardsHtml || `<div class="small">No results yet.</div>`;
  }

  function makeTable(features, maxFields) {
    if (!features || !features.length) return `<div class="small">No sample features fetched.</div>`;

    const attrs0 = features[0].attributes || {};
    const keys = Object.keys(attrs0).slice(0, maxFields);

    const th = keys.map(k => `<th>${escapeHtml(k)}</th>`).join("");
    const rows = features.map(f => {
      const a = f.attributes || {};
      const tds = keys.map(k => `<td>${escapeHtml(a[k] == null ? "" : a[k])}</td>`).join("");
      return `<tr>${tds}</tr>`;
    }).join("");

    return `
      <div class="table-wrap">
        <table>
          <thead><tr>${th}</tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    `;
  }

  function downloadText(filename, text) {
    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function toCsv(rows) {
    if (!rows || !rows.length) return "";
    const cols = Object.keys(rows[0]);
    const escape = (v) => {
      const s = (v == null) ? "" : String(v);
      if (/[,"\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
      return s;
    };
    const header = cols.map(escape).join(",");
    const body = rows.map(r => cols.map(c => escape(r[c])).join(",")).join("\n");
    return header + "\n" + body;
  }

  function flattenAttributes(features) {
    return (features || []).map(f => (f && f.attributes) ? f.attributes : {});
  }

  // ---------- Query logic ----------
  async function querySingleLayer(layerUrl, layerTitle, geom) {
    const layer = new FeatureLayer({ url: layerUrl, outFields: ["*"] });
    const q = layer.createQuery();
    q.geometry = geom;
    q.spatialRelationship = "intersects";
    q.returnGeometry = false;
    q.outFields = ["*"];

    const count = await layer.queryFeatureCount(q);

    const maxSamples = (config.report && config.report.maxSampleFeaturesPerLayer) ? config.report.maxSampleFeaturesPerLayer : 25;
    let features = [];

    if (count > 0 && maxSamples > 0) {
      const q2 = q.clone();
      q2.num = Math.min(maxSamples, 2000);
      const fs = await layer.queryFeatures(q2);
      features = (fs && fs.features) ? fs.features : [];
    }

    return { title: layerTitle, url: layerUrl, count, features };
  }

  async function runReport() {
    if (!selectionGeom) return;

    setStatus("running report…");
    resultsEl.innerHTML = "";
    exportAllBtn.disabled = true;
    lastReportRowsByLayer = [];

    const reportCfgs = config.reportLayers || [];
    const expandedTargets = [];

    // Expand FeatureServer roots into sublayers
    for (const cfg of reportCfgs) {
      if (isFeatureServerRoot(cfg.url)) {
        try {
          const sublayers = await expandServiceToSublayers(cfg.url);
          sublayers.forEach(sl => expandedTargets.push({
            title: `${cfg.title}: ${sl.title}`,
            url: sl.url
          }));
        } catch (e) {
          expandedTargets.push({
            title: `${cfg.title} (FAILED to expand)`,
            url: cfg.url,
            error: e
          });
        }
      } else {
        expandedTargets.push({ title: cfg.title, url: cfg.url });
      }
    }

    const cards = [];
    for (let i = 0; i < expandedTargets.length; i++) {
      const t = expandedTargets[i];

      if (t.error) {
        cards.push(`
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(t.title)}</div>
              <div class="badge">error</div>
            </div>
            <div class="small mono">${escapeHtml(String(t.error))}</div>
          </div>
        `);
        continue;
      }

      try {
        const r = await querySingleLayer(t.url, t.title, selectionGeom);

        const rows = flattenAttributes(r.features);
        lastReportRowsByLayer.push({ title: r.title, url: r.url, rows });

        const maxFields = (config.report && config.report.maxFieldsInTable) ? config.report.maxFieldsInTable : 8;
        const tableHtml = (r.features && r.features.length) ? makeTable(r.features, maxFields) : `<div class="small">No sample rows.</div>`;

        cards.push(`
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(r.title)}</div>
              <div class="badge">count: <b>${r.count}</b></div>
            </div>
            <div class="small mono">${escapeHtml(r.url)}</div>
            <div style="margin-top:8px;">
              ${tableHtml}
              <div class="row" style="margin-top:8px;">
                <button class="btn subtle" data-export="${escapeHtml(r.title)}">Export CSV (sample)</button>
              </div>
            </div>
          </div>
        `);
      } catch (e) {
        cards.push(`
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(t.title)}</div>
              <div class="badge">error</div>
            </div>
            <div class="small mono">${escapeHtml(String(e))}</div>
          </div>
        `);
      }

      setStatus(`running report… (${i + 1}/${expandedTargets.length})`);
    }

    renderResults(cards.join(""));
    wireExportButtons();
    exportAllBtn.disabled = (lastReportRowsByLayer.length === 0);
    setStatus("done");
  }

  function wireExportButtons() {
    resultsEl.querySelectorAll("button[data-export]").forEach(btn => {
      btn.addEventListener("click", () => {
        const title = btn.getAttribute("data-export");
        const item = lastReportRowsByLayer.find(x => x.title === title);
        if (!item) return;
        const csv = toCsv(item.rows);
        downloadText(safeFilename(title) + ".csv", csv || "");
      });
    });
  }

  function safeFilename(name) {
    return String(name).replace(/[^\w\-]+/g, "_").replace(/_+/g, "_").replace(/^_+|_+$/g, "").slice(0, 120) || "export";
  }

  // ---------- Selection layer setup ----------
  async function setActiveSelectionLayerByIndex(idx) {
    const entry = selectionLayers[idx];
    if (!entry) return;

    activeSelectionLayer = entry.layer;
    activeSelectionLayerView = await view.whenLayerView(activeSelectionLayer);

    clearHighlight();
    setGeometryFromSelection(null);
    setStatus("select mode: click a polygon");
  }

  function attachClickToSelect() {
    view.on("click", async (event) => {
      if (modeSelect.value !== "select") return;
      if (!activeSelectionLayerView) return;

      try {
        const hit = await view.hitTest(event);
        const results = (hit && hit.results) ? hit.results : [];
        const match = results.find(r => r.graphic && r.graphic.layer && activeSelectionLayer && r.graphic.layer === activeSelectionLayer);

        if (!match) return;

        const graphic = match.graphic;
        if (!graphic || !graphic.geometry) return;

        clearHighlight();
        activeHighlightHandle = activeSelectionLayerView.highlight(graphic);

        setGeometryFromSelection(graphic.geometry);
        setStatus("polygon selected (ready to run)");
      } catch (e) {
        console.error(e);
        setStatus("select failed (see console)");
      }
    });
  }

  // ---------- Init ----------
  async function init() {
    setStatus("loading config…");

    config = await fetchJson("./config.json");

    const map = new Map({ basemap: config.map?.basemap || "gray-vector" });

    view = new MapView({
      container: "viewDiv",
      map,
      center: config.map?.center || [-98.5795, 39.8283],
      zoom: config.map?.zoom || 4
    });

    // Draw layer + sketch
    drawLayer = new GraphicsLayer({ title: "Drawn polygon" });
    map.add(drawLayer);

    sketch = new Sketch({
      view,
      layer: drawLayer,
      availableCreateTools: ["polygon"],
      creationMode: "single"
    });

    sketch.on("create", (evt) => {
      if (evt.state === "complete") {
        // use the completed sketch geometry
        setGeometryFromSelection(evt.graphic.geometry);
        setStatus("drawn polygon ready (run report)");
      }
    });

    // Selection layers (visible)
    const selCfgs = config.selectionLayers || [];
    selectionLayers = selCfgs.map(cfg => ({
      cfg,
      layer: new FeatureLayer({
        url: cfg.url,
        title: cfg.title,
        outFields: ["*"]
      })
    }));
    selectionLayers.forEach(e => map.add(e.layer));

    // Populate selection layer dropdown
    selectionLayerSelect.innerHTML = selectionLayers.map((e, i) =>
      `<option value="${i}">${escapeHtml(e.cfg.title)}</option>`
    ).join("");

    await view.when();
    attachClickToSelect();

    await setActiveSelectionLayerByIndex(0);

    // UI wiring
    modeSelect.addEventListener("change", () => setMode(modeSelect.value));
    selectionLayerSelect.addEventListener("change", () => setActiveSelectionLayerByIndex(Number(selectionLayerSelect.value)));

    drawBtn.addEventListener("click", () => {
      view.ui.add(sketch, "top-left");
      sketch.create("polygon");
      setStatus("drawing polygon…");
    });

    stopDrawBtn.addEventListener("click", () => {
      sketch.cancel();
      setStatus("draw stopped");
    });

    runBtn.addEventListener("click", runReport);
    clearBtn.addEventListener("click", clearAll);

    exportAllBtn.addEventListener("click", () => {
      // bundle all samples into one CSV with an added __layer column
      const allRows = [];
      for (const item of lastReportRowsByLayer) {
        for (const r of (item.rows || [])) {
          allRows.push({ __layer: item.title, ...r });
        }
      }
      const csv = toCsv(allRows);
      downloadText("intersect_report_all_samples.csv", csv || "");
    });

    setMode("select");
    renderConfiguredLayerList();
    setStatus("ready");
  }

  init().catch((e) => {
    console.error(e);
    setStatus("failed to initialize (see console)");
  });

});
