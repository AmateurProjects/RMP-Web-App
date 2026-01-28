/* global require */

require([
    "esri/Map",
    "esri/views/MapView",
    "esri/layers/FeatureLayer",
    "esri/layers/GraphicsLayer",
    "esri/widgets/Sketch",
    "esri/widgets/BasemapToggle",
    "esri/Graphic"
], function (EsriMap, MapView, FeatureLayer, GraphicsLayer, Sketch, BasemapToggle, Graphic) {

    // ---------- DOM ----------
    const modeSelect = document.getElementById("modeSelect");
    // PLSS selection tools (Township / Section / Intersected)
    const plssTownshipBtn = document.getElementById("plssTownshipBtn");
    const plssSectionBtn = document.getElementById("plssSectionBtn");
    const plssIntersectedBtn = document.getElementById("plssIntersectedBtn");
    const selectModeControls = document.getElementById("selectModeControls");
    const drawModeControls = document.getElementById("drawModeControls");

    const drawBtn = document.getElementById("drawBtn");
    const stopDrawBtn = document.getElementById("stopDrawBtn");
    const runBtn = document.getElementById("runBtn");
    const clearBtn = document.getElementById("clearBtn");
    const exportAllBtn = document.getElementById("exportAllBtn");

    const statusEl = document.getElementById("status");
    const statusTextEl = document.getElementById("statusText");
    const busyIndicatorEl = document.getElementById("busyIndicator");

    const resultsEl = document.getElementById("results");
    const selectionLayerTogglesEl = document.getElementById("selectionLayerToggles");
    const reportLayerTogglesEl = document.getElementById("reportLayerToggles");

    // Tabs + Services + Visual DOM
    const tabReportBtn = document.getElementById("tabReportBtn");
    const tabVisualBtn = document.getElementById("tabVisualBtn");
    const tabServicesBtn = document.getElementById("tabServicesBtn");

    const tabReportPanel = document.getElementById("tabReportPanel");
    const tabVisualPanel = document.getElementById("tabVisualPanel");
    const tabServicesPanel = document.getElementById("tabServicesPanel");

    // Visual report DOM
    const generateVisualBtn = document.getElementById("generateVisualBtn");
    const visualReportStatusEl = document.getElementById("visualReportStatus");
    const visualReportMapWrapEl = document.getElementById("visualReportMapWrap");
    const visualReportImgEl = document.getElementById("visualReportImg");
    const visualReportSummaryEl = document.getElementById("visualReportSummary");
    const downloadMapBtn = document.getElementById("downloadMapBtn");
    const printVisualBtn = document.getElementById("printVisualBtn");

    const servicesListEl = document.getElementById("servicesList");
    const refreshServicesBtn = document.getElementById("refreshServicesBtn");

    function setStatus(msg) {
        const text = "Status: " + msg;
        if (statusTextEl) statusTextEl.textContent = text;
        else if (statusEl) statusEl.textContent = text;
    }

    function setBusy(isBusy) {
        if (!busyIndicatorEl) return;
        busyIndicatorEl.classList.toggle("hidden", !isBusy);
    }

    // ---------- State ----------
    let config = null;

    let view = null;
    let selectionGeom = null;
    let map = null; // <-- add (so PLSS buttons can add/remove selection layers)

    // AOI overlay (always on top)
    let aoiLayer = null;      // GraphicsLayer
    let aoiGraphic = null;    // Graphic (single AOI graphic)

    // Renderer lookup helpers
    let layerCfgByUrl = new Map(); // url -> {kind, cfg}

    let selectionLayers = []; // { cfg, layer }
    let activeSelectionLayer = null; // FeatureLayer
    let activeSelectionLayerView = null;

    let sketch = null;

    let lastReportRowsByLayer = []; // for export-all
    let reportLayerViews = new Map(); // url -> FeatureLayer (only if user toggles it on for map visibility)


    // ---------- Helpers ----------

    // Tabs
    function setActiveTab(tabName) {
        const isReport = (tabName === "report");
        const isVisual = (tabName === "visual");
        const isServices = (tabName === "services");

        if (tabReportPanel) tabReportPanel.classList.toggle("active", isReport);
        if (tabVisualPanel) tabVisualPanel.classList.toggle("active", isVisual);
        if (tabServicesPanel) tabServicesPanel.classList.toggle("active", isServices);

        if (tabReportBtn) tabReportBtn.classList.toggle("active", isReport);
        if (tabVisualBtn) tabVisualBtn.classList.toggle("active", isVisual);
        if (tabServicesBtn) tabServicesBtn.classList.toggle("active", isServices);
    }

    function escapeHtml(s) {
        return String(s).replace(/[&<>"']/g, (c) => ({
            "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;"
        }[c]));
    }

    function normalize(s){ return String(s || "").toLowerCase(); }

    function findSelectionLayerIndexByNameIncludes(needle) {
        const n = normalize(needle);
        return (selectionLayers || []).findIndex(e => normalize(e?.cfg?.title).includes(n));
    }

    function setPlssToolActive(which) {
        const set = (btn, on) => {
            if (!btn) return;
            btn.setAttribute("aria-pressed", on ? "true" : "false");
        };
        set(plssTownshipBtn, which === "township");
        set(plssSectionBtn, which === "section");
        set(plssIntersectedBtn, which === "intersected");
    }

    function updateSelectionToggleCheckbox(idx, checked) {
        const cb = document.getElementById(`sellayer_${idx}`);
        if (!cb) return;
        cb.checked = !!checked;
    }

    function isLayerOnMap(layer) {
        if (!map || !layer) return false;
        return map.layers.includes(layer);
    }

    function enableSelectionLayer(idx) {
        const entry = selectionLayers[idx];
        if (!entry) return;
        if (!isLayerOnMap(entry.layer)) map.add(entry.layer);
        entry.layer.visible = true;
        updateSelectionToggleCheckbox(idx, true);
        ensureAoiOnTop(map);
    }

    function disableSelectionLayer(idx) {
        const entry = selectionLayers[idx];
        if (!entry) return;

        // Remove from map (your desired behavior vs hide)
        if (isLayerOnMap(entry.layer)) map.remove(entry.layer);

        // Also mark it invisible for safety (even though removed)
        entry.layer.visible = false;

        updateSelectionToggleCheckbox(idx, false);

        // If we just disabled the active selection layer, clear active pointers
        if (activeSelectionLayer === entry.layer) {
            activeSelectionLayer = null;
            activeSelectionLayerView = null;
        }
    }

async function autoZoomToLayerMinVisible(layer) {
    if (!view || !layer) return;

    const minScale = Number(layer.minScale || 0);
    if (!minScale || !isFinite(minScale) || minScale <= 0) return;

    // Nudge a bit more zoomed-in than minScale so the layer reliably renders.
    // Smaller scale number = more zoomed in.
    const nudgeFactor = 0.50; // 50% more zoomed in than minScale (tweak 0.90–0.95 if desired)
    const targetScale = Math.max(1, Math.floor(minScale * nudgeFactor));

    if (view.scale > targetScale) {
        await view.goTo({ scale: targetScale }, { animate: true, duration: 450 });
    }
}


    function isFeatureServerRoot(url) {
        // ends with /FeatureServer (no trailing /0 etc.)
        return /\/FeatureServer\/?$/.test(url);
    }

    function isMapServerRoot(url) {
        return /\/MapServer\/?$/.test(url);
    }

    // Expand a MapServer root into sublayers that can be used by FeatureLayer.
    // Optionally filters to polygon layers only (best for “click a polygon to select”).
    async function expandMapServerToSublayers(serviceUrl, { polygonOnly = true } = {}) {
        const pjsonUrl = serviceUrl.replace(/\/$/, "") + "?f=pjson";
        const info = await fetchJson(pjsonUrl);
        const layers = Array.isArray(info?.layers) ? info.layers : [];

        // If polygonOnly: we need each layer’s pjson to know geometryType (MapServer root doesn’t always include it).
        const out = [];

        for (const l of layers) {
            const layerUrl = serviceUrl.replace(/\/$/, "") + "/" + l.id;

            if (polygonOnly) {
                try {
                    const lpjson = await fetchJson(layerUrl + "?f=pjson");
                    const g = String(lpjson?.geometryType || "");
                    // ArcGIS geometry types are like "esriGeometryPolygon"
                    if (!g.toLowerCase().includes("polygon")) continue;
                } catch (e) {
                    // If a sublayer doesn’t return pjson, skip it (safer)
                    continue;
                }
            }

            out.push({
                title: `${l.name}`,
                url: layerUrl
            });
        }

        return out;
    }

    function setBasemapBaseLayerOpacity(basemap, opacity) {
        try {
            const baseLayers = basemap?.baseLayers?.toArray ? basemap.baseLayers.toArray() : [];
            baseLayers.forEach(l => { l.opacity = opacity; });
        } catch (e) {
            // ignore
        }
    }

    function isImageryBasemap(basemap) {
        // For ArcGIS JS basemap IDs like "satellite", "hybrid"
        const id = (basemap && (basemap.id || basemap.portalItem?.id || basemap.title)) ? String(basemap.id || basemap.title || "") : "";
        const title = basemap?.title ? String(basemap.title).toLowerCase() : "";
        return title.includes("satellite") || title.includes("imagery") || title.includes("hybrid") || id.toLowerCase().includes("satellite") || id.toLowerCase().includes("hybrid");
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

    async function wireLayerUpdatingSpinner(layer, spinnerEl) {
        if (!layer || !spinnerEl || !view) return;

        try {
            await layer.when();
            const lv = await view.whenLayerView(layer);

            spinnerEl.classList.toggle("hidden", !lv.updating);

            lv.watch("updating", (isUpdating) => {
                spinnerEl.classList.toggle("hidden", !isUpdating);
            });
        } catch (e) {
            spinnerEl.classList.add("hidden");
        }
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
            title: (l && l.name) ? String(l.name) : `Layer ${l.id}`,
            url: serviceUrl.replace(/\/$/, "") + "/" + l.id
        }));
    }


    function clearAll() {
        selectionGeom = null;
        resultsEl.innerHTML = "";
        exportAllBtn.disabled = true;
        lastReportRowsByLayer = [];

        if (aoiLayer) aoiLayer.removeAll();
        aoiGraphic = null;

        runBtn.disabled = true;
        setStatus("cleared");
        setBusy(false);
    }

    function setGeometryFromSelection(geom) {
        selectionGeom = geom || null;
        runBtn.disabled = !selectionGeom;
    }

    function setMode(mode) {
        function startDrawingNow() {
            if (!sketch) return;
            // Cancel any prior sketch session and start a new polygon immediately
            sketch.cancel();
            sketch.create("polygon");
            setStatus("drawing polygon…");
        }

        if (mode === "select") {
            selectModeControls.classList.remove("hidden");
            drawModeControls.classList.add("hidden");
            // stop sketch if running
            if (sketch) sketch.cancel();
            setStatus("select mode: click a polygon");
        } else {
            selectModeControls.classList.add("hidden");
            drawModeControls.classList.remove("hidden");
            startDrawingNow(); // <-- auto start drawing immediately
        }
        // keep current selectionGeom if user switches modes intentionally
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
                <span id="sellayer_spin_${i}" class="layer-spinner hidden" aria-label="loading"></span>
            </div>
            `;
        }).join("");

    (selectionLayers || []).forEach((e, i) => {
        const cb = document.getElementById(`sellayer_${i}`);
        if (!cb) return;

        cb.addEventListener("change", async () => {
            const isOnMap = map.layers.includes(e.layer);

            if (cb.checked) {
                // turning ON — add to map if not present
                if (!isOnMap) map.add(e.layer);
                e.layer.visible = true;
                ensureAoiOnTop(map);

                // If nothing is active, make this the active selection layer
                if (!activeSelectionLayer) {
                    await setActiveSelectionLayerByIndex(i);
                }
            } else {
                // turning OFF — remove from map so it *actually disappears*
                if (isOnMap) map.remove(e.layer);

                // If the user just removed the active selection layer,
                // switch active selection to the next available ON-map layer (or null).
                if (activeSelectionLayer === e.layer) {
                    activeSelectionLayer = null;
                    activeSelectionLayerView = null;

                    // find first selection layer currently ON the map
                    const nextIdx = (selectionLayers || []).findIndex(x => map.layers.includes(x.layer));
                    if (nextIdx >= 0) {
                        await setActiveSelectionLayerByIndex(nextIdx);
                    } else {
                        setGeometryFromSelection(null);
                        setStatus("no selection layers visible (turn one on)");
                    }
                }
            }
        });

        // spinner wiring (shows while layer view is updating)
        const spin = document.getElementById(`sellayer_spin_${i}`);
        wireLayerUpdatingSpinner(e.layer, spin);
    });

        // ---- Report layers (ALWAYS included in report): toggle ONLY map visibility
        // If a report URL is a FeatureServer ROOT (no /0 etc.), it cannot be drawn directly.
        // We will show it in the list but disable the checkbox to avoid confusion.
        reportLayerTogglesEl.innerHTML = (config.reportLayers || []).map((l, i) => {
            const isRoot = isFeatureServerRoot(l.url) || isMapServerRoot(l.url);
            const key = String(l.url || "").replace(/\/+$/, "");
            const existing = reportLayerViews.get(key);
            const checked = existing ? (existing.visible ? "checked" : "") : "";
            const disabled = isRoot ? "disabled" : "";
            const note = isRoot ? ` <span class="small">(service root; not drawable)</span>` : "";

            return `
            <div class="toggle-row">
                <input type="checkbox" id="rptlayer_${i}" ${checked} ${disabled} />
                <label class="toggle-name" for="rptlayer_${i}">${escapeHtml(l.title)}${note}</label>
                <span id="rptlayer_spin_${i}" class="layer-spinner hidden" aria-label="loading"></span>
            </div>
            `;
        }).join("");

        (config.reportLayers || []).forEach((l, i) => {
            const cb = document.getElementById(`rptlayer_${i}`);
            if (!cb) return;

            // If disabled (FeatureServer root), no handler
            if (cb.disabled) return;

            // Normalize URL key so get/set/delete always match (trailing slash is the usual culprit)
            const key = String(l.url || "").replace(/\/+$/, "");

            cb.addEventListener("change", () => {
                let lyr = reportLayerViews.get(key);

                if (cb.checked) {
                    // turning ON
                    if (!lyr) {
                        const cfgMatch = layerCfgByUrl.get(key)?.cfg || layerCfgByUrl.get(l.url)?.cfg;

                        lyr = new FeatureLayer({
                            url: l.url,
                            title: l.title,
                            outFields: ["*"],
                            visible: true,
                            renderer: getPresetRenderer("report", cfgMatch) || undefined
                        });

                        map.add(lyr);
                        reportLayerViews.set(key, lyr);

                        const spin = document.getElementById(`rptlayer_spin_${i}`);
                        wireLayerUpdatingSpinner(lyr, spin);

                        ensureAoiOnTop(map);
                    } else {
                        lyr.visible = true;
                    }
                } else {
                    // turning OFF — remove from map so it *actually disappears*
                    if (lyr) {
                        map.remove(lyr);
                        reportLayerViews.delete(key);
                    } else {
                        // defensive cleanup: if for some reason we missed the reference, try removing by URL match
                        const toRemove = map.layers
                            .toArray()
                            .find(x => x?.type === "feature" && String(x?.url || "").replace(/\/+$/, "") === key);

                        if (toRemove) map.remove(toRemove);
                        reportLayerViews.delete(key);
                    }
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
        ? `
        <div class="small service-desc" id="svc_desc_${i}">${escapeHtml(desc)}</div>
        <button class="service-desc-toggle" type="button" data-desc-toggle="${i}">Show more</button>
        `
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
            <div class="small mono service-url">
            <a href="${escapeHtml(it.url)}" target="_blank" rel="noopener">Service URL</a>
            </div>
            ${descHtml}
            ${errHtml}
        </div>
        `);
    }

    servicesListEl.innerHTML = cards.join("");

    // Wire description expand/collapse toggles
    servicesListEl.querySelectorAll("button[data-desc-toggle]").forEach(btn => {
        btn.addEventListener("click", () => {
            const idx = btn.getAttribute("data-desc-toggle");
            const card = btn.closest(".service-card");
            if (!card) return;
            const isExpanded = card.classList.toggle("expanded");
            btn.textContent = isExpanded ? "Show less" : "Show more";
        });
    });

    }

    // ---------- Report rendering ----------
    function renderResults(cardsHtml) {
        resultsEl.innerHTML = cardsHtml || `<div class="small">No results yet.</div>`;
    }


    function sampleWithoutReplacement(arr, n) {
    const a = (arr || []).slice();
    if (a.length <= n) return a;
    // Fisher–Yates shuffle partial
    for (let i = a.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [a[i], a[j]] = [a[j], a[i]];
    }
    return a.slice(0, n);
    }

    function makeTable(features, maxFields) {
        if (!features || !features.length) return `<div class="small">No sample features fetched.</div>`;

        const picked = sampleWithoutReplacement(features, 4);

        const attrs0 = picked[0].attributes || {};
        const keys = Object.keys(attrs0).slice(0, maxFields);

        const th = keys.map(k => `<th title="${escapeHtml(k)}">${escapeHtml(k)}</th>`).join("");

        const rows = picked.map(f => {
            const a = f.attributes || {};
            const tds = keys.map(k => {
                const raw = (a[k] == null) ? "" : String(a[k]);

                // truncate values longer than the *column name*
                const maxLen = Math.max(4, String(k).length); // keep sane minimum
                let shown = raw;

                if (raw.length > maxLen) {
                    shown = raw.slice(0, Math.max(1, maxLen - 1)) + "…";
                }

                const safeFull = escapeHtml(raw);
                const safeShown = escapeHtml(shown);

                return `<td title="${safeFull}">${safeShown}</td>`;
            }).join("");
            return `<tr>${tds}</tr>`;
        }).join("");

        return `
        <div class="table-wrap">
            <table class="result-table">
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

        setBusy(true);
        setStatus("running report…");
        resultsEl.innerHTML = "";
        exportAllBtn.disabled = true;
        lastReportRowsByLayer = [];

    const combinedCfgs = [
        ...(config.reportLayers || []),
        ...(config.selectionLayers || [])
    ];

    // De-duplicate by URL (same service could appear in both lists)
    const seenUrls = new Set();
    const reportCfgs = combinedCfgs.filter(l => {
        const url = String(l?.url || "");
        if (!url) return false;
        if (seenUrls.has(url)) return false;
        seenUrls.add(url);
        return true;
    });

    const expandedTargets = [];

    // Expand service roots into sublayers
    for (const cfg of reportCfgs) {
        const url = String(cfg.url || "");

        if (isFeatureServerRoot(url)) {
            try {
                const sublayers = await expandServiceToSublayers(url);
                sublayers.forEach(sl => expandedTargets.push({
                    title: `${cfg.title}: ${sl.title}`,
                    url: sl.url
                }));
            } catch (e) {
                expandedTargets.push({
                    title: `${cfg.title} (FAILED to expand)`,
                    url,
                    error: e
                });
            }
            continue;
        }

        if (isMapServerRoot(url)) {
            try {
                // For reporting, include ALL sublayers (not polygonOnly)
                const subs = await expandMapServerToSublayers(url, { polygonOnly: false });
                subs.forEach(sl => expandedTargets.push({
                    title: `${cfg.title}: ${sl.title}`,
                    url: sl.url
                }));
            } catch (e) {
                expandedTargets.push({
                    title: `${cfg.title} (FAILED to expand)`,
                    url,
                    error: e
                });
            }
            continue;
        }

        expandedTargets.push({ title: cfg.title, url });
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
                count: r.count,       // <-- store count for summary stats
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
                <div class="small mono">
                <a href="${escapeHtml(r.url)}" target="_blank" rel="noopener">Service URL</a>
                </div>
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
        renderVisualSummary();
        setBusy(false);
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


    function getVisualSummaryLines() {
        // Uses the same stats as renderVisualSummary(), but returns plain text lines for PNG.
        if (!selectionGeom) return ["No AOI selected."];

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            return ["Run the report to populate layer counts."];
        }

        const totalLayers = lastReportRowsByLayer.length;
        const layersWithHits = lastReportRowsByLayer.filter(x => (x.count || 0) > 0);
        const totalHits = lastReportRowsByLayer.reduce((sum, x) => sum + (x.count || 0), 0);

        const top = layersWithHits
            .slice()
            .sort((a, b) => (b.count || 0) - (a.count || 0))
            .slice(0, 10);

        const lines = [
            `Layers queried: ${totalLayers}`,
            `Layers with hits: ${layersWithHits.length}`,
            `Total intersecting features (sum of counts): ${totalHits}`,
            ""
        ];

        if (top.length) {
            lines.push("Top layers:");
            top.forEach(x => lines.push(`• ${x.title} (${x.count || 0})`));
        } else {
            lines.push("(No intersect hits.)");
        }

        return lines;
    }

    function wrapText(ctx, text, maxWidth) {
        const words = String(text || "").split(/\s+/).filter(Boolean);
        if (!words.length) return [""];

        const lines = [];
        let line = words[0];

        for (let i = 1; i < words.length; i++) {
            const test = line + " " + words[i];
            if (ctx.measureText(test).width <= maxWidth) line = test;
            else { lines.push(line); line = words[i]; }
        }
        lines.push(line);
        return lines;
    }

    async function buildVisualPngWithSummary(mapDataUrl) {
        const img = new Image();
        img.crossOrigin = "anonymous";

        await new Promise((resolve, reject) => {
            img.onload = () => resolve();
            img.onerror = (e) => reject(e);
            img.src = mapDataUrl;
        });

        const padding = 18;
        const lineH = 18;
        const titleH = 22;

        // Create a canvas the same width as the screenshot
        const w = img.naturalWidth || img.width;
        const summaryLines = getVisualSummaryLines();

        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");

        // Set fonts for measuring/wrapping
        ctx.font = "14px system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif";

        // Wrap lines to fit
        const maxTextWidth = w - padding * 2;
        const wrapped = [];
        for (const line of summaryLines) {
            if (!line) { wrapped.push(""); continue; }
            wrapText(ctx, line, maxTextWidth).forEach(x => wrapped.push(x));
        }

        const summaryBlockH = padding + titleH + (wrapped.length * lineH) + padding;

        canvas.width = w;
        canvas.height = img.height + summaryBlockH;

        // Background
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);

        // Draw map screenshot
        ctx.drawImage(img, 0, 0);

        // Draw summary panel background
        const y0 = img.height;
        ctx.fillStyle = "rgba(255,255,255,0.96)";
        ctx.fillRect(0, y0, canvas.width, summaryBlockH);

        // Summary title
        ctx.fillStyle = "#111111";
        ctx.font = "700 16px system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif";
        ctx.fillText("Visual Report Summary", padding, y0 + padding + 16);

        // Summary lines
        ctx.font = "14px system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif";
        let y = y0 + padding + titleH;

        for (const line of wrapped) {
            if (!line) { y += lineH; continue; }
            ctx.fillText(line, padding, y);
            y += lineH;
        }

        return canvas.toDataURL("image/png");
    }


    function setVisualStatus(msg) {
        if (visualReportStatusEl) visualReportStatusEl.textContent = msg || "";
    }

    function renderVisualSummary() {
        if (!visualReportSummaryEl) return;

        if (!selectionGeom) {
            visualReportSummaryEl.innerHTML = `<div class="small">(No AOI selected.)</div>`;
            return;
        }

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            visualReportSummaryEl.innerHTML = `<div class="small">(Run the report to populate layer counts.)</div>`;
            return;
        }

        const totalLayers = lastReportRowsByLayer.length;
        const layersWithHits = lastReportRowsByLayer.filter(x => (x.count || 0) > 0);
        const totalHits = lastReportRowsByLayer.reduce((sum, x) => sum + (x.count || 0), 0);

        const top = layersWithHits
            .slice()
            .sort((a, b) => (b.count || 0) - (a.count || 0))
            .slice(0, 12);

        const listHtml = top.length
            ? `<div style="margin-top:8px;">
                ${top.map(x => `<div class="small">• ${escapeHtml(x.title)} <span class="mono">(${x.count})</span></div>`).join("")}
            </div>`
            : `<div class="small" style="margin-top:8px;">(No intersect hits.)</div>`;

        visualReportSummaryEl.innerHTML = `
        <div class="small">Layers queried: <b>${totalLayers}</b></div>
        <div class="small">Layers with hits: <b>${layersWithHits.length}</b></div>
        <div class="small">Total intersecting features (sum of counts): <b>${totalHits}</b></div>
        ${listHtml}
        `;
    }

    async function generateVisualReport() {
        if (!view) return;

        if (!selectionGeom) {
            setVisualStatus("Select or draw an AOI first.");
            return;
        }

        setBusy(true);
        setVisualStatus("Generating map…");
        if (visualReportMapWrapEl) visualReportMapWrapEl.classList.add("hidden");
        if (downloadMapBtn) downloadMapBtn.disabled = true;
        if (printVisualBtn) printVisualBtn.disabled = true;

        try {
            // Zoom to AOI with padding
            const paddingFactor = config?.visualReport?.paddingFactor ?? 1.25;
            const width = config?.visualReport?.screenshotWidth ?? 1400;

            // Use AOI extent if available; fallback to goTo geometry directly
            const ext = selectionGeom?.extent;
            if (ext && ext.expand) {
                await view.goTo(ext.expand(paddingFactor), { animate: true, duration: 450 });
            } else {
                await view.goTo(selectionGeom, { animate: true, duration: 450 });
            }

            // Ensure AOI draws on top (if you implemented AOI reorder helper)
            // If you have ensureAoiOnTop(map) in your codebase, keep this:
            try { ensureAoiOnTop(view.map); } catch (e) {}

            // Screenshot the current view
            const ss = await view.takeScreenshot({ format: "png", quality: 100, width });
            const dataUrl = ss?.dataUrl;

            if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

            // ✅ Item 6: bake summary stats into the PNG
            const combinedUrl = await buildVisualPngWithSummary(dataUrl);

            if (visualReportImgEl) visualReportImgEl.src = combinedUrl;
            if (visualReportMapWrapEl) visualReportMapWrapEl.classList.remove("hidden");

            // Enable download
            if (downloadMapBtn) {
                downloadMapBtn.disabled = false;
                downloadMapBtn.onclick = () => {
                    const a = document.createElement("a");
                    a.href = combinedUrl;
                    a.download = "AOI_map_with_summary.png";
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                };
            }

            // Enable print (simple print window with image + summary)
            if (printVisualBtn) {
                printVisualBtn.disabled = false;
                printVisualBtn.onclick = () => {
                    const summaryHtml = visualReportSummaryEl ? visualReportSummaryEl.innerHTML : "";
                    const w = window.open("", "_blank");
                    if (!w) return;

                    w.document.write(`
                    <!doctype html>
                    <html>
                    <head>
                        <meta charset="utf-8" />
                        <meta name="viewport" content="width=device-width,initial-scale=1" />
                        <title>Visual Report</title>
                        <style>
                        body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; }
                        h1 { font-size: 18px; margin: 0 0 10px 0; }
                        img { width: 100%; height: auto; border: 1px solid #ddd; border-radius: 12px; }
                        .small { font-size: 12px; color: #444; }
                        .mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
                        .section { margin-top: 14px; }
                        </style>
                    </head>
                    <body>
                        <h1>Visual Report</h1>
                        <div class="section">
                        <img src="${combinedUrl}" alt="AOI map" />
                        </div>
                        <div class="section">${summaryHtml}</div>
                        <script>window.onload = () => window.print();</script>
                    </body>
                    </html>
                    `);
                    w.document.close();
                };
            }

            setVisualStatus("Done.");
        } catch (e) {
            console.error(e);
            setVisualStatus("Failed to generate map (see console).");
        } finally {
            renderVisualSummary();
            setBusy(false);
        }
    }


    // ---------- Selection layer setup ----------
    async function setActiveSelectionLayerByIndex(idx) {
        const entry = selectionLayers[idx];
        if (!entry) return;

        activeSelectionLayer = entry.layer;
        activeSelectionLayerView = await view.whenLayerView(activeSelectionLayer);

        setGeometryFromSelection(null);
        setStatus("select mode: click a polygon");
    }

    function attachClickToSelect() {
        view.on("click", async (event) => {

            // If drawing AOI, let Sketch own the click experience
            if (modeSelect.value === "draw") {
                return;
            }

            if (!activeSelectionLayerView) return;

            try {
                const hit = await view.hitTest(event);
                const results = (hit && hit.results) ? hit.results : [];

                // First: try to select from active selection layer
                const match = results.find(r =>
                    r.graphic && r.graphic.layer && activeSelectionLayer && r.graphic.layer === activeSelectionLayer
                );

                if (match) {
                    const graphic = match.graphic;
                    if (!graphic || !graphic.geometry) return;
                    setAoiGeometry(graphic.geometry);
                    setGeometryFromSelection(graphic.geometry);
                    setStatus("polygon selected (ready to run)");
                    return;
                }


            } catch (e) {
                console.error(e);
                setStatus("click inspect failed (see console)");
            }
        });
    }

    // ---------- Init ----------
    async function init() {
        setStatus("loading config…");

        config = await fetchJson("./config.json");
        layerCfgByUrl = buildLayerCfgIndex(config);

        map = new EsriMap({ basemap: config.map?.basemap || "gray-vector" });

        view = new MapView({
            container: "viewDiv",
            map,
            center: config.map?.center || [-98.5795, 39.8283],
            zoom: config.map?.zoom || 4
        });

        // Disable Esri popup UI; we’ll use our own minimal popup
        view.popup.autoOpenEnabled = false;


        // Basemap toggle (near zoom controls)
        const imageryBasemapId = config?.map?.imageryBasemap || "satellite"; // "satellite" is Esri World Imagery
        const imageryOpacity = config?.map?.imageryOpacity ?? 0.75;

        const basemapToggle = new BasemapToggle({
            view,
            nextBasemap: imageryBasemapId
        });
        view.ui.add(basemapToggle, "top-left");

        // Enforce imagery opacity when imagery is active (and restore for non-imagery)
        view.watch("map.basemap", (bm) => {
            if (!bm) return;
            if (isImageryBasemap(bm)) setBasemapBaseLayerOpacity(bm, imageryOpacity);
            else setBasemapBaseLayerOpacity(bm, 1);
        });

        // Apply once on load
        if (isImageryBasemap(view.map.basemap)) setBasemapBaseLayerOpacity(view.map.basemap, imageryOpacity);

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

        // Selection layers (may include MapServer roots that expand into many sublayers)
        const selCfgs = config.selectionLayers || [];
        const expandedSelectionCfgs = [];

        // Track PLSS State Boundaries so it can be report-only (not selectable)
        let plssStateBoundary = null; // { title, url }

        for (const cfg of selCfgs) {
            const url = String(cfg?.url || "");
            if (isMapServerRoot(url)) {
                // Expand MapServer into polygon sublayers for selection
                const subs = await expandMapServerToSublayers(url, { polygonOnly: true });

                subs.forEach(sl => {
                    const subTitle = String(sl.title || "");

                    // ✅ Item 4: remove "State Boundaries" from Selection (but keep for Report)
                    if (subTitle.toLowerCase() === "state boundaries") {
                        plssStateBoundary = {
                            title: `${cfg.title}: ${subTitle}`,
                            url: sl.url
                        };
                        return; // skip adding to selection
                    }

                    expandedSelectionCfgs.push({
                        title: `${cfg.title}: ${subTitle}`,
                        url: sl.url,
                        visible: true
                    });
                });
            } else {
                expandedSelectionCfgs.push(cfg);
            }
        }

        // ✅ Ensure State Boundaries still appears in REPORT layers
        if (plssStateBoundary) {
            const alreadyInReport = (config.reportLayers || []).some(r => {
                return String(r?.url || "").replace(/\/+$/, "") === String(plssStateBoundary.url).replace(/\/+$/, "");
            });

            if (!alreadyInReport) {
                config.reportLayers = config.reportLayers || [];
                config.reportLayers.push({
                    title: plssStateBoundary.title,
                    url: plssStateBoundary.url
                });
            }
        }

        selectionLayers = expandedSelectionCfgs.map(cfg => ({
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


        await view.when();
        attachClickToSelect();

    // ---------- PLSS tool wiring (Township / Section / Intersected) ----------
    const townshipIdx = findSelectionLayerIndexByNameIncludes("township");
    const sectionIdx = findSelectionLayerIndexByNameIncludes("section");
    const intersectedIdx = findSelectionLayerIndexByNameIncludes("intersected"); // "PLSS Intersected"

    // Helper: make ONE PLSS layer active, disable the other two, and auto-zoom if needed
    async function activatePlss(which, idxToEnable) {
        // Force select mode (PLSS tools are select-only)
        if (modeSelect && modeSelect.value !== "select") {
            modeSelect.value = "select";
            setMode("select");
        }

        // Enable chosen layer even if user unchecked it earlier
        const trio = [townshipIdx, sectionIdx, intersectedIdx].filter(i => i >= 0);

        // Disable the other two first
        for (const idx of trio) {
            if (idx !== idxToEnable) disableSelectionLayer(idx);
        }

        // Enable the chosen one
        if (idxToEnable >= 0) enableSelectionLayer(idxToEnable);

        // Set as active selection layer
        if (idxToEnable >= 0) {
            await setActiveSelectionLayerByIndex(idxToEnable);
            setPlssToolActive(which);

            // Auto-zoom to minimum visible zoom level (using layer.minScale)
            const lyr = selectionLayers[idxToEnable]?.layer;
            await autoZoomToLayerMinVisible(lyr);

            setStatus(`PLSS select: ${which} (click a polygon)`);
        } else {
            setPlssToolActive(which);
            setStatus("PLSS select: layer not found in selection layers");
        }
    }

    if (plssTownshipBtn) plssTownshipBtn.addEventListener("click", () => activatePlss("township", townshipIdx));
    if (plssSectionBtn) plssSectionBtn.addEventListener("click", () => activatePlss("section", sectionIdx));
    if (plssIntersectedBtn) plssIntersectedBtn.addEventListener("click", () => activatePlss("intersected", intersectedIdx));

    // Default to Township if present, otherwise Section, otherwise Intersected, otherwise first selection layer
    if (townshipIdx >= 0) {
        await activatePlss("township", townshipIdx);
    } else if (sectionIdx >= 0) {
        await activatePlss("section", sectionIdx);
    } else if (intersectedIdx >= 0) {
        await activatePlss("intersected", intersectedIdx);
    } else if (selectionLayers.length) {
        enableSelectionLayer(0);
        await setActiveSelectionLayerByIndex(0);
        setPlssToolActive("township"); // best-effort UI state
    }


        // Tab wiring
        if (tabReportBtn) tabReportBtn.addEventListener("click", () => setActiveTab("report"));

        if (tabVisualBtn) tabVisualBtn.addEventListener("click", () => {
            setActiveTab("visual");
            renderVisualSummary();
            setVisualStatus(selectionGeom ? "Ready." : "Select or draw an AOI first.");
        });

        if (tabServicesBtn) tabServicesBtn.addEventListener("click", async () => {
            setActiveTab("services");
            await refreshServicesTab();
        });

        if (refreshServicesBtn) refreshServicesBtn.addEventListener("click", refreshServicesTab);
        if (generateVisualBtn) generateVisualBtn.addEventListener("click", generateVisualReport);


        // UI wiring
        if (modeSelect) {
            modeSelect.addEventListener("change", () => setMode(modeSelect.value));
        }

        if (drawBtn) {
            drawBtn.addEventListener("click", () => {
                // No sketch toolbar UI; just start drawing immediately
                if (modeSelect && modeSelect.value !== "draw") modeSelect.value = "draw";
                setMode("draw"); // will start drawing automatically
            });
        }

        if (stopDrawBtn) {
            stopDrawBtn.addEventListener("click", () => {
                if (sketch) sketch.cancel();
                setStatus("draw stopped");
            });
        }

        if (runBtn) runBtn.addEventListener("click", runReport);
        if (clearBtn) clearBtn.addEventListener("click", clearAll);


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
