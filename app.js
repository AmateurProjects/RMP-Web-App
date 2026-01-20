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
    const selectionLayerTogglesEl = document.getElementById("selectionLayerToggles");
    const reportLayerTogglesEl = document.getElementById("reportLayerToggles");


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
    let reportLayerViews = new Map(); // url -> FeatureLayer (only if user toggles it on for map visibility)


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

    function renderLayerToggles(map) {
        // Guard: if the HTML containers don't exist, do nothing
        if (!selectionLayerTogglesEl || !reportLayerTogglesEl) return;

        // ---- Selection layers (already on map): toggle visibility
        selectionLayerTogglesEl.innerHTML = (selectionLayers || []).map((e, i) => {
            const checked = e.layer.visible ? "checked" : "";
            return `
            <div class="toggle-row">
                <input type="checkbox" id="sellayer_${i}" ${checked} />
                <label class="toggle-name" for="sellayer_${i}">${escapeHtml(e.cfg.title)}</label>
            </div>
            `;
        }).join("");

        (selectionLayers || []).forEach((e, i) => {
            const cb = document.getElementById(`sellayer_${i}`);
            if (!cb) return;
            cb.addEventListener("change", () => {
                e.layer.visible = cb.checked;
            });
        });

        // ---- Report layers (ALWAYS included in report): toggle ONLY map visibility
        // If a report URL is a FeatureServer ROOT (no /0 etc.), it cannot be drawn directly.
        // We will show it in the list but disable the checkbox to avoid confusion.
        reportLayerTogglesEl.innerHTML = (config.reportLayers || []).map((l, i) => {
            const isRoot = isFeatureServerRoot(l.url);
            const existing = reportLayerViews.get(l.url);
            const checked = existing ? (existing.visible ? "checked" : "") : "";
            const disabled = isRoot ? "disabled" : "";
            const note = isRoot ? ` <span class="small">(service root; not drawable)</span>` : "";

            return `
            <div class="toggle-row">
                <input type="checkbox" id="rptlayer_${i}" ${checked} ${disabled} />
                <label class="toggle-name" for="rptlayer_${i}">${escapeHtml(l.title)}${note}</label>
            </div>
            `;
        }).join("");

        (config.reportLayers || []).forEach((l, i) => {
            const cb = document.getElementById(`rptlayer_${i}`);
            if (!cb) return;

            // If disabled (FeatureServer root), no handler
            if (cb.disabled) return;

            cb.addEventListener("change", () => {
                const wantVisible = cb.checked;

                if (wantVisible) {
                    // Lazily create & add to map if needed
                    let lyr = reportLayerViews.get(l.url);
                    if (!lyr) {
                        lyr = new FeatureLayer({
                            url: l.url,
                            title: l.title,
                            outFields: ["*"],
                            visible: true
                        });
                        map.add(lyr);
                        reportLayerViews.set(l.url, lyr);
                    } else {
                        lyr.visible = true;
                    }
                } else {
                    const lyr = reportLayerViews.get(l.url);
                    if (lyr) lyr.visible = false;
                }
            });
        });
    }

    async function queryAllFeaturesPaged(layer, baseQuery, pageSize, maxExportFeatures) {
        const all = [];
        let offset = 0;

        while (true) {
            const q = baseQuery.clone();
            q.num = pageSize;
            q.start = offset;               // ArcGIS JS uses start for resultOffset
            q.returnGeometry = false;

            const fs = await layer.queryFeatures(q);
            const feats = (fs && fs.features) ? fs.features : [];

            all.push(...feats);

            if (feats.length < pageSize) break;
            offset += pageSize;

            if (maxExportFeatures && all.length >= maxExportFeatures) break;
        }

        return all;
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
        q.outFields = ["*"];
        q.returnGeometry = false;

        const count = await layer.queryFeatureCount(q);

        const maxSamples = config.report?.maxSampleFeaturesPerLayer ?? 25;
        let features = [];

        if (count > 0 && maxSamples > 0) {
            const q2 = q.clone();
            q2.num = Math.min(maxSamples, 2000);
            const fs = await layer.queryFeatures(q2);
            features = fs?.features ?? [];
        }

        return { title: layerTitle, url: layerUrl, count, features, layer, exportQuery: q };
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

            // Store sample rows PLUS the objects we need for FULL export paging
            lastReportRowsByLayer.push({
                title: r.title,
                url: r.url,
                rows,                 // sample rows shown in the UI table
                _layer: r.layer,      // FeatureLayer instance used for querying
                _exportQuery: r.exportQuery, // Query object (intersects geometry etc.)
                fullRows: null        // will be filled on-demand when user exports FULL
            });

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
                <button class="btn subtle" data-export="${escapeHtml(r.title)}">
                Export FULL CSV
                </button>
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
                btn.addEventListener("click", async () => {
                    const title = btn.getAttribute("data-export");
                    const item = lastReportRowsByLayer.find(x => x.title === title);
                    if (!item) return;

                    // If we already fetched full rows once, just export again
                    if (item.fullRows && item.fullRows.length) {
                        const csvCached = toCsv(item.fullRows);
                        downloadText(safeFilename(title) + "_FULL.csv", csvCached || "");
                        return;
                    }

                    // Defensive: make sure we have what we need
                    if (!item._layer || !item._exportQuery) {
                        // fallback to sample if something is missing
                        const csvSample = toCsv(item.rows);
                        downloadText(safeFilename(title) + "_SAMPLE.csv", csvSample || "");
                        return;
                    }

                    btn.disabled = true;
                    const oldStatus = statusEl.textContent;

                    try {
                        setStatus("exporting FULL CSV…");

                        const pageSize = config.report?.pageSize ?? 1000;
                        const maxExport = config.report?.maxExportFeatures ?? 50000;

                        // Page through all intersecting features
                        const fullFeatures = await queryAllFeaturesPaged(
                            item._layer,
                            item._exportQuery,
                            pageSize,
                            maxExport
                        );

                        // Convert to rows + cache
                        item.fullRows = flattenAttributes(fullFeatures);

                        const csvFull = toCsv(item.fullRows);
                        downloadText(safeFilename(title) + "_FULL.csv", csvFull || "");

                        // Optional: tell user if capped
                        if (maxExport && fullFeatures.length >= maxExport) {
                            setStatus(`exported FULL (capped at ${maxExport})`);
                        } else {
                            setStatus("exported FULL");
                        }
                    } catch (e) {
                        console.error(e);
                        setStatus("export failed (see console)");
                    } finally {
                        btn.disabled = false;
                        // If you prefer to restore prior status:
                        // statusEl.textContent = oldStatus;
                    }
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
                outFields: ["*"],
                visible: cfg.visible !== false
            })

        }));
        selectionLayers.forEach(e => map.add(e.layer));
        renderLayerToggles(map);


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

        exportAllBtn.addEventListener("click", async () => {
            if (!lastReportRowsByLayer.length) return;

            exportAllBtn.disabled = true;

            try {
                setStatus("exporting ALL (FULL)…");

                const pageSize = config.report?.pageSize ?? 1000;
                const maxExport = config.report?.maxExportFeatures ?? 50000;

                const allRows = [];

                for (let i = 0; i < lastReportRowsByLayer.length; i++) {
                    const item = lastReportRowsByLayer[i];

                    // Skip if we somehow don't have the query objects
                    if (!item._layer || !item._exportQuery) continue;

                    setStatus(`exporting ALL (FULL)… (${i + 1}/${lastReportRowsByLayer.length})`);

                    // Use cached full results if available
                    if (!item.fullRows) {
                        const fullFeatures = await queryAllFeaturesPaged(
                            item._layer,
                            item._exportQuery,
                            pageSize,
                            maxExport
                        );
                        item.fullRows = flattenAttributes(fullFeatures);
                    }

                    for (const r of (item.fullRows || [])) {
                        allRows.push({ __layer: item.title, ...r });
                    }
                }

                const csv = toCsv(allRows);
                downloadText("intersect_report_ALL_FULL.csv", csv || "");
                setStatus("exported ALL (FULL)");
            } catch (e) {
                console.error(e);
                setStatus("export ALL failed (see console)");
            } finally {
                exportAllBtn.disabled = false;
            }
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
