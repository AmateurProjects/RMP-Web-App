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

    // Tabs + Services DOM
    const tabReportBtn = document.getElementById("tabReportBtn");
    const tabServicesBtn = document.getElementById("tabServicesBtn");
    const tabReportPanel = document.getElementById("tabReportPanel");
    const tabServicesPanel = document.getElementById("tabServicesPanel");

    const servicesListEl = document.getElementById("servicesList");
    const refreshServicesBtn = document.getElementById("refreshServicesBtn");

    function setStatus(msg) { statusEl.textContent = "Status: " + msg; }

    // ---------- State ----------
    let config = null;

    let view = null;
    let selectionGeom = null;

    // AOI overlay (always on top)
    let aoiLayer = null;      // GraphicsLayer
    let aoiGraphic = null;    // Graphic (single AOI graphic)

    // Renderer lookup helpers
    let layerCfgByUrl = new Map(); // url -> {kind, cfg}

    let selectionLayers = []; // { cfg, layer }
    let activeSelectionLayer = null; // FeatureLayer
    let activeSelectionLayerView = null;
    let activeHighlightHandle = null;

    let drawLayer = null;
    let sketch = null;

    let lastReportRowsByLayer = []; // for export-all
    let reportLayerViews = new Map(); // url -> FeatureLayer (only if user toggles it on for map visibility)


    // ---------- Helpers ----------
    // Tabs
    function setActiveTab(tabName) {
    const isReport = (tabName === "report");

        if (tabReportPanel) tabReportPanel.classList.toggle("active", isReport);
        if (tabServicesPanel) tabServicesPanel.classList.toggle("active", !isReport);

        if (tabReportBtn) tabReportBtn.classList.toggle("active", isReport);
        if (tabServicesBtn) tabServicesBtn.classList.toggle("active", !isReport);
    }

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

    // timed JSON fetch for "UP/DOWN" checks
    async function fetchJsonWithTimeout(url, timeoutMs = 8000) {
    const controller = new AbortController();
    const t = window.setTimeout(() => controller.abort(), timeoutMs);

    try {
        const res = await fetch(url, { credentials: "omit", signal: controller.signal });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return await res.json();
    } finally {
        window.clearTimeout(t);
    }
    }

    // read basic description from service/layer pjson
    function pickServiceDescription(pjson) {
    // Different services expose different fields; we pick the first useful one.
    const candidates = [
        pjson?.serviceDescription,
        pjson?.description,
        pjson?.documentInfo?.Title,
        pjson?.name
    ].filter(Boolean);

    return candidates.length ? String(candidates[0]) : "";
    }

    function normalizePjsonUrl(u) {
    return u.replace(/\/$/, "") + "?f=pjson";
    }

    function buildLayerCfgIndex(cfg) {
        const m = new Map();

        const addList = (kind, arr) => {
            (arr || []).forEach(l => {
                if (!l || !l.url) return;
                m.set(String(l.url), { kind, cfg: l });
            });
        };

        addList("selection", cfg?.selectionLayers);
        addList("report", cfg?.reportLayers);

        return m;
    }

    function getPresetRenderer(kind, cfgObj) {
        const sym = config?.symbology || {};
        const defaults = sym.defaults || {};
        const presets = sym.presets || {};

        // Allow per-layer override later (optional)
        const presetId =
            (cfgObj && cfgObj.symbologyPreset) ||
            (kind === "selection" ? defaults.selectionPreset :
            kind === "report" ? defaults.reportPreset :
            defaults.aoiPreset);

        const r = presetId ? presets[presetId] : null;
        return r || null;
    }

    function ensureAoiOnTop(map) {
        if (!map || !aoiLayer) return;
        // Put AOI layer at top draw order
        map.reorder(aoiLayer, map.layers.length - 1);
    }

    function setAoiGeometry(geom) {
        // Clears and redraws AOI graphic so it’s always visible (and exportable later)
        if (!aoiLayer) return;

        aoiLayer.removeAll();
        aoiGraphic = null;

        if (!geom) return;

        const aoiRenderer = getPresetRenderer("aoi", null);
        const aoiSymbol = aoiRenderer?.symbol; // simple renderer expected

        aoiGraphic = new Graphic({
            geometry: geom,
            symbol: aoiSymbol || undefined
        });

        aoiLayer.add(aoiGraphic);
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

        if (aoiLayer) aoiLayer.removeAll();
        aoiGraphic = null;

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
                        const cfgMatch = layerCfgByUrl.get(l.url)?.cfg;

                        lyr = new FeatureLayer({
                            url: l.url,
                            title: l.title,
                            outFields: ["*"],
                            visible: true,
                            renderer: getPresetRenderer("report", cfgMatch) || undefined
                        });
                        map.add(lyr);
                        reportLayerViews.set(l.url, lyr);

                        // Keep AOI above everything
                        ensureAoiOnTop(map);
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

    // ---------- Services tab ----------
    function getConfiguredServices() {
    // Show the “services used by the app itself” from config
    const seen = new Set();
    const out = [];

    const add = (kind, title, url) => {
        const key = `${kind}||${url}`;
        if (seen.has(key)) return;
        seen.add(key);
        out.push({ kind, title, url });
    };

    (config.selectionLayers || []).forEach(l => add("Selection", l.title, l.url));
    (config.reportLayers || []).forEach(l => add("Report", l.title, l.url));

    return out;
    }

    async function refreshServicesTab() {
    if (!servicesListEl) return;

    const items = getConfiguredServices();
    if (!items.length) {
        servicesListEl.innerHTML = `<div class="small">No services configured.</div>`;
        return;
    }

    servicesListEl.innerHTML = `<div class="small">Checking services…</div>`;

    const timeoutMs = config?.services?.timeoutMs ?? 8000;

    // Run checks sequentially (simple + predictable). We can add concurrency later if needed.
    const cards = [];
    for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const pjsonUrl = normalizePjsonUrl(it.url);

        let status = "DOWN";
        let desc = "";
        let errText = "";

        try {
        const pjson = await fetchJsonWithTimeout(pjsonUrl, timeoutMs);
        status = "UP";
        desc = pickServiceDescription(pjson);
        } catch (e) {
        status = "DOWN";
        errText = String(e?.message || e);
        }

        const pillClass = (status === "UP") ? "pill pill-up" : "pill pill-down";
        const descHtml = desc
        ? `<div class="small" style="margin-top:6px;">${escapeHtml(desc)}</div>`
        : `<div class="small" style="margin-top:6px; opacity:.8;">(No description found in pjson)</div>`;

        const errHtml = (status === "DOWN")
        ? `<div class="small mono" style="margin-top:6px;">${escapeHtml(errText)}</div>`
        : "";

        cards.push(`
        <div class="service-card">
            <div class="service-head">
            <div>
                <div class="result-title">${escapeHtml(it.title)}</div>
                <div class="small">${escapeHtml(it.kind)}</div>
            </div>
            <div class="${pillClass}">${status}</div>
            </div>
            <div class="small mono service-url">${escapeHtml(it.url)}</div>
            ${descHtml}
            ${errHtml}
        </div>
        `);
    }

    servicesListEl.innerHTML = cards.join("");
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

    function toCsv(rows, preferredFirstCols = []) {
        if (!rows || !rows.length) return "";

        // Union of all keys across all rows
        const colSet = new Set();
        for (const r of rows) {
            if (!r) continue;
            Object.keys(r).forEach(k => colSet.add(k));
        }

        // Put preferred columns first (if present), then the rest alphabetically
        const preferred = (preferredFirstCols || []).filter(c => colSet.has(c));
        preferred.forEach(c => colSet.delete(c));

        const rest = Array.from(colSet).sort((a, b) => a.localeCompare(b));
        const cols = [...preferred, ...rest];

        const escape = (v) => {
            const s = (v == null) ? "" : String(v);
            if (/[,"\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
            return s;
        };

        const header = cols.map(escape).join(",");
        const body = rows.map(r => cols.map(c => escape(r ? r[c] : "")).join(",")).join("\n");
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

                clearHighlight(); // optional: keep highlight off to avoid double-outline clutter

                setAoiGeometry(graphic.geometry);      // AOI boundary on top (exportable later)
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
        layerCfgByUrl = buildLayerCfgIndex(config);

        const map = new Map({ basemap: config.map?.basemap || "gray-vector" });

        view = new MapView({
            container: "viewDiv",
            map,
            center: config.map?.center || [-98.5795, 39.8283],
            zoom: config.map?.zoom || 4
        });

        // AOI layer + sketch (AOI must always be visible and on top)
        aoiLayer = new GraphicsLayer({ title: "AOI" });
        map.add(aoiLayer);

        // Sketch draws directly into AOI layer
        sketch = new Sketch({
            view,
            layer: aoiLayer,
            availableCreateTools: ["polygon"],
            creationMode: "single"
        });

        // Apply AOI symbol to Sketch (uses the AOI preset renderer symbol)
        const aoiRenderer = getPresetRenderer("aoi", null);
        if (aoiRenderer && aoiRenderer.symbol) {
            sketch.polygonSymbol = aoiRenderer.symbol;
        }

        sketch.on("create", (evt) => {
            if (evt.state === "complete") {
                const geom = evt.graphic?.geometry || null;
                setAoiGeometry(geom);          // ensure AOI is a single clean graphic
                setGeometryFromSelection(geom);
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
                visible: cfg.visible !== false,
                renderer: getPresetRenderer("selection", cfg) || undefined
            })
        }));
        selectionLayers.forEach(e => map.add(e.layer));
        renderLayerToggles(map);
        ensureAoiOnTop(map);

        // Populate selection layer dropdown
        selectionLayerSelect.innerHTML = selectionLayers.map((e, i) =>
            `<option value="${i}">${escapeHtml(e.cfg.title)}</option>`
        ).join("");

        await view.when();
        attachClickToSelect();

        await setActiveSelectionLayerByIndex(0);

        // Tab wiring
        if (tabReportBtn) tabReportBtn.addEventListener("click", () => setActiveTab("report"));
        if (tabServicesBtn) tabServicesBtn.addEventListener("click", async () => {
        setActiveTab("services");
        await refreshServicesTab();
        });
        if (refreshServicesBtn) refreshServicesBtn.addEventListener("click", refreshServicesTab);

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

                const csv = toCsv(allRows, ["__layer"]);
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

        // Preload service status once (optional). Keeps Services tab fast.
        if (servicesListEl) {
        refreshServicesTab().catch(() => {});
        }
    }

    init().catch((e) => {
        console.error(e);
        setStatus("failed to initialize (see console)");
    });

});
