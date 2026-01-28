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
    let reportLayerViews = new Map(); 
    // key -> FeatureLayer OR FeatureLayer[] (for FeatureServer/MapServer roots that expand into multiple drawable layers)


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

    // Expand a FeatureServer root into polygon sublayers (drawable FeatureLayer URLs).
    async function expandFeatureServerToPolygonSublayers(serviceUrl) {
        const pjsonUrl = serviceUrl.replace(/\/$/, "") + "?f=pjson";
        const info = await fetchJson(pjsonUrl);
        const layers = Array.isArray(info?.layers) ? info.layers : [];

        const out = [];
        for (const l of layers) {
            const layerUrl = serviceUrl.replace(/\/$/, "") + "/" + l.id;

            try {
                const lpjson = await fetchJson(layerUrl + "?f=pjson");
                const g = String(lpjson?.geometryType || "").toLowerCase();
                if (!g.includes("polygon")) continue;
            } catch (e) {
                continue;
            }

            out.push({
                title: l?.name ? String(l.name) : `Layer ${l.id}`,
                url: layerUrl
            });
        }

        return out;
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

    // ---- Selection layers
    selectionLayerTogglesEl.innerHTML = (selectionLayers || []).map((e, i) => {
        const checked = e.layer.visible ? "checked" : "";
        return `
            <div class="toggle-row">
                <input type="checkbox" id="sellayer_${i}" ${checked} />
                <span class="layer-swatch layer-swatch-selection" aria-hidden="true" title="Selection layer"></span>
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
                if (!isOnMap) map.add(e.layer);
                e.layer.visible = true;
                ensureAoiOnTop(map);

                if (!activeSelectionLayer) {
                    await setActiveSelectionLayerByIndex(i);
                }
            } else {
                if (isOnMap) map.remove(e.layer);

                if (activeSelectionLayer === e.layer) {
                    activeSelectionLayer = null;
                    activeSelectionLayerView = null;

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

        const spin = document.getElementById(`sellayer_spin_${i}`);
        wireLayerUpdatingSpinner(e.layer, spin);
    });

    // ---- Report layers (map visibility only; report still always runs)
    reportLayerTogglesEl.innerHTML = (config.reportLayers || []).map((l, i) => {
        const isRoot = isFeatureServerRoot(l.url) || isMapServerRoot(l.url);
        const key = String(l.url || "").replace(/\/+$/, "");
        const existing = reportLayerViews.get(key);

        const isChecked =
            Array.isArray(existing) ? (existing.length > 0) :
            existing ? !!existing.visible :
            false;

        const checked = isChecked ? "checked" : "";
        const note = isRoot ? ` <span class="small">(expands to polygon sublayers)</span>` : "";

        return `
            <div class="toggle-row">
                <input type="checkbox" id="rptlayer_${i}" ${checked} />
                <span class="layer-swatch layer-swatch-report" aria-hidden="true" title="Report layer"></span>
                <label class="toggle-name" for="rptlayer_${i}">${escapeHtml(l.title)}${note}</label>
                <span id="rptlayer_spin_${i}" class="layer-spinner hidden" aria-label="loading"></span>
            </div>
        `;
    }).join("");

    (config.reportLayers || []).forEach((l, i) => {
        const cb = document.getElementById(`rptlayer_${i}`);
        if (!cb) return;

        const key = String(l.url || "").replace(/\/+$/, "");

        cb.addEventListener("change", async () => {
            const spin = document.getElementById(`rptlayer_spin_${i}`);

            if (cb.checked) {
                try {
                    if (spin) spin.classList.remove("hidden");

                    // FeatureServer root -> expand polygon sublayers
                    if (isFeatureServerRoot(l.url)) {
                        const subs = await expandFeatureServerToPolygonSublayers(l.url);
                        const cfgMatch = layerCfgByUrl.get(key)?.cfg || layerCfgByUrl.get(l.url)?.cfg;

                        const created = subs.map(sl => new FeatureLayer({
                            url: sl.url,
                            title: `${l.title}: ${sl.title}`,
                            outFields: ["*"],
                            visible: true,
                            renderer: getPresetRenderer("report", cfgMatch) || undefined
                        }));

                        created.forEach(lyr => map.add(lyr));
                        reportLayerViews.set(key, created);

                        if (created.length && spin) wireLayerUpdatingSpinner(created[0], spin);
                        ensureAoiOnTop(map);
                        return;
                    }

                    // Normal single layer
                    let lyr = reportLayerViews.get(key);

                    if (Array.isArray(lyr)) {
                        // Defensive: shouldn't happen here, but don't crash
                        lyr.forEach(x => map.add(x));
                        return;
                    }

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

                        if (spin) wireLayerUpdatingSpinner(lyr, spin);
                        ensureAoiOnTop(map);
                    } else {
                        lyr.visible = true;
                    }
                } catch (e) {
                    console.error(e);
                    setStatus("failed to enable report layer (see console)");
                    cb.checked = false;
                } finally {
                    if (spin) spin.classList.add("hidden");
                }
            } else {
                // turning OFF — remove from map
                const lyr = reportLayerViews.get(key);

                if (Array.isArray(lyr)) {
                    lyr.forEach(x => { try { map.remove(x); } catch (e) {} });
                    reportLayerViews.delete(key);
                    return;
                }

                if (lyr) {
                    map.remove(lyr);
                    reportLayerViews.delete(key);
                } else {
                    // defensive cleanup by URL match
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
});
