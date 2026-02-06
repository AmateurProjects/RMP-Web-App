/* global require */

require([
    "esri/Map",
    "esri/views/MapView",
    "esri/layers/FeatureLayer",
    "esri/layers/GraphicsLayer",
    "esri/widgets/Sketch",
    "esri/widgets/BasemapToggle",
    "esri/Graphic",
    "esri/geometry/geometryEngine",
    "esri/layers/TileLayer"
], function (EsriMap, MapView, FeatureLayer, GraphicsLayer, Sketch, BasemapToggle, Graphic, geometryEngine, TileLayer) {


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

    const cancelRunBtn = document.getElementById("cancelRunBtn");
    const viewBlockerEl = document.getElementById("viewBlocker");

    const statusEl = document.getElementById("status");
    const statusTextEl = document.getElementById("statusText");
    const busyIndicatorEl = document.getElementById("busyIndicator");

    const resultsCardEl = document.getElementById("resultsCard");
    const resultsEl = document.getElementById("results");
    const selectionLayerTogglesEl = document.getElementById("selectionLayerToggles");
    const reportLayerTogglesEl = document.getElementById("reportLayerToggles");

    // Tabs + Panels (NEW STRUCTURE)
    const tabLayersBtn = document.getElementById("tabLayersBtn");
    const tabServicesBtn = document.getElementById("tabServicesBtn");
    const tabReportBtn = document.getElementById("tabReportBtn");
    const tabVisualBtn = document.getElementById("tabVisualBtn");
    const tabFinalReportBtn = document.getElementById("tabFinalReportBtn");

    const tabLayersPanel = document.getElementById("tabLayersPanel");
    const tabServicesPanel = document.getElementById("tabServicesPanel");  
    const tabReportPanel = document.getElementById("tabReportPanel");      
    const tabVisualPanel = document.getElementById("tabVisualPanel");      
    const tabReportFinalPanel = document.getElementById("tabReportFinalPanel");

    // Visual report DOM
    const visualReportStatusEl = document.getElementById("visualReportStatus");
    const visualReportMapWrapEl = document.getElementById("visualReportMapWrap");
    // const visualReportImgEl = document.getElementById("visualReportImg"); // no longer used
    const visualReportOutputsEl = document.getElementById("visualReportOutputs");
    const visualReportSummaryEl = document.getElementById("visualReportSummary");

    // Final report DOM
    const viewReportBtn = document.getElementById("viewReportBtn");
    const finalReportStatus = document.getElementById("finalReportStatus");

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
    let aoiSource = null;            // "draw" | "select"
    let aoiSourceLayerTitle = null;  // optional: which selection layer was clicked
    let map = null; // <-- add (so PLSS buttons can add/remove selection layers)

    // AOI overlay (always on top)
    let aoiLayer = null;      // GraphicsLayer
    let aoiGraphic = null;    // Graphic (single AOI graphic)

    // Renderer lookup helpers
    let layerCfgByUrl = new Map(); // url -> {kind, cfg}

    let selectionLayers = []; // { cfg, layer }
    let activeSelectionLayer = null; // FeatureLayer
    let activeSelectionLayerView = null;
    let aoiSourcePlssTool = null; // "township" | "section" | "intersected" | null
    let aoiSourceLayerUrl = null;      // URL of the selection layer used to pick AOI (select mode)
    let plssParcelLayerUrl = null;     // URL of PLSS Intersected (UI will call "Parcel")
    let plssStateLayerUrl = null;      // URL of PLSS State Boundaries (single canonical)
    let aoiSourceObjectId = null;      // ObjectID of the clicked AOI polygon (select mode)
    let aoiSourceObjectIdField = null; // ObjectID field name for that layer
    let aoiSourceFeature = null;       // ✅ cached clicked feature (attributes for AOI Source card)


    let sketch = null;

    let lastReportRowsByLayer = []; // for export-all
    let reportLayerViews = new Map();

    // Cached final report HTML
    let cachedFinalReportHtml = null;

    // Track layerView "updating" watch handles so we can remove them (prevents leaks)
    const spinnerWatchByLayerUid = new Map(); // layer.uid -> watchHandle

    function setSpinnerWatch(layer, handle) {
        if (!layer) return;
        const uid = layer.uid;
        if (!uid) return;

        // Remove old watch if present
        const old = spinnerWatchByLayerUid.get(uid);
        if (old && old.remove) old.remove();

        if (handle) spinnerWatchByLayerUid.set(uid, handle);
        else spinnerWatchByLayerUid.delete(uid);
    }

    function clearSpinnerWatch(layer) {
        if (!layer) return;
        const uid = layer.uid;
        const h = uid ? spinnerWatchByLayerUid.get(uid) : null;
        if (h && h.remove) h.remove();
        if (uid) spinnerWatchByLayerUid.delete(uid);
    }

    // key -> FeatureLayer OR FeatureLayer[] (for FeatureServer/MapServer roots that expand into multiple drawable layers)

    // ---------- report-layer toggle cancellation tokens ----------
    const reportToggleToken = new Map(); // key -> number

    function bumpReportToken(key) {
        const next = (reportToggleToken.get(key) || 0) + 1;
        reportToggleToken.set(key, next);
        return next;
    }

    function getReportToken(key) {
        return reportToggleToken.get(key) || 0;
    }

    function isTokenCurrent(key, token) {
        return getReportToken(key) === token;
    }

    // ---------- Operation lock + cancel (Report / Visual) ----------
    let reportOpToken = 0;

    const navDefaults = { captured: false, values: {} };
    const navProps = [
        "mouseWheelEnabled",
        "dragPanEnabled",
        "browserTouchPanEnabled",
        "keyboardEnabled",
        "doubleClickZoomEnabled"
    ];

    function lockMapInteraction(isLocked) {
        // UI overlay that blocks pointer events
        if (viewBlockerEl) viewBlockerEl.classList.toggle("hidden", !isLocked);

        // Also disable navigation toggles (belt + suspenders)
        try {
            if (!view?.navigation) return;

            const nav = view.navigation;

            if (!navDefaults.captured) {
                navDefaults.captured = true;
                navProps.forEach(p => { if (p in nav) navDefaults.values[p] = nav[p]; });
            }

            if (isLocked) {
                navProps.forEach(p => { if (p in nav) nav[p] = false; });
            } else {
                navProps.forEach(p => {
                    if (p in nav && p in navDefaults.values) nav[p] = navDefaults.values[p];
                });
            }
        } catch (e) {
            // ignore
        }
    }

    function startReportOp() {
        const my = ++reportOpToken;
        lockMapInteraction(true);
        if (cancelRunBtn) cancelRunBtn.classList.remove("hidden");
        return my;
    }

    function endReportOp(myToken) {
        // Only unlock if this is the most recent op (prevents weird edge cases)
        if (myToken === reportOpToken) {
            lockMapInteraction(false);
            if (cancelRunBtn) cancelRunBtn.classList.add("hidden");
        }
    }

    function isReportCanceled(myToken) { return myToken !== reportOpToken; }


    // ---------- Helpers ----------

    // Tabs
function setActiveTab(tabName) {
    const isLayers = (tabName === "layers");
    const isServices = (tabName === "services");
    const isReport = (tabName === "report");
    const isVisual = (tabName === "visual");
    const isFinalReport = (tabName === "finalReport");

    // Panels
    if (tabLayersPanel) tabLayersPanel.classList.toggle("active", isLayers);
    if (tabServicesPanel) tabServicesPanel.classList.toggle("active", isServices);
    if (tabReportPanel) tabReportPanel.classList.toggle("active", isReport);
    if (tabVisualPanel) tabVisualPanel.classList.toggle("active", isVisual);
    if (tabReportFinalPanel) tabReportFinalPanel.classList.toggle("active", isFinalReport);

    // Buttons
    if (tabLayersBtn) tabLayersBtn.classList.toggle("active", isLayers);
    if (tabServicesBtn) tabServicesBtn.classList.toggle("active", isServices);
    if (tabReportBtn) tabReportBtn.classList.toggle("active", isReport);
    if (tabVisualBtn) tabVisualBtn.classList.toggle("active", isVisual);
    if (tabFinalReportBtn) tabFinalReportBtn.classList.toggle("active", isFinalReport);

    // Force MapView to re-measure + redraw after layout changes
    if (view) {
        requestAnimationFrame(() => {
            try { view.resize(); } catch (e) { }
            try { view.requestRender(); } catch (e) { }
        });
    }
}

    function plssToolLabel(which) {
        return (which === "intersected") ? "Parcel" :
            (which === "township") ? "Township" :
                (which === "section") ? "Section" :
                    "PLSS";
    }


    function escapeHtml(s) {
        return String(s).replace(/[&<>"']/g, (c) => ({
            "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;"
        }[c]));
    }

    function normalize(s) { return String(s || "").toLowerCase(); }

    function isPlssLayerTitleOrUrl(title, url) {
        const t = normalize(title);
        const u = normalize(url);
        // Tune this if you want it stricter; this keeps it PLSS-focused.
        return (
            t.includes("plss") ||
            t.includes("township") ||
            t.includes("section") ||
            t.includes("intersected") ||
            u.includes("/plss") ||
            u.includes("plss")
        );
    }

    function isPlssIntersectedLayerTitle(title) {
        const t = normalize(title);
        // Match your naming: "PLSS: Intersected" etc.
        return t.includes("intersected");
    }


    function filterTouchingOnly(features, aoiGeom) {
        // Drops polygon features that only touch AOI at an edge/vertex (intersection area == 0)
        if (!features?.length || !aoiGeom) return features || [];
        const EPS = 0.000001; // sq meters

        return features.filter(f => {
            const g = f?.geometry;
            if (!g) return false;
            try {
                const inter = geometryEngine.intersect(aoiGeom, g);
                if (!inter) return false;
                const area = geometryEngine.geodesicArea(inter, "square-meters");
                return area > EPS;
            } catch (e) {
                // If geometry ops fail, keep it (don’t accidentally drop legit hits)
                return true;
            }
        });
    }

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

    function setSelectionSpinner(idx, isOn) {
        const spin = document.getElementById(`sellayer_spin_${idx}`);
        if (!spin) return;
        spin.classList.toggle("hidden", !isOn);
    }


    function isLayerOnMap(layer) {
        if (!map || !layer) return false;
        return map.layers.includes(layer);
    }

    async function enableSelectionLayer(idx) {
        const entry = selectionLayers[idx];
        if (!entry) return;

        clearSpinnerWatch(entry.layer);

        // show spinner immediately
        setSelectionSpinner(idx, true);

        // ✅ DO NOT add/remove; selection layers are already added once in init()
        entry.layer.visible = true;
        updateSelectionToggleCheckbox(idx, true);
        ensureAoiOnTop(map);

        // ✅ Make sure we're zoomed to a scale where this layer can draw
        await ensureLayerVisibleAtScale(entry.layer);
        await waitForViewStationary(1500);

        // ✅ Wire updating -> spinner truth
        const spin = document.getElementById(`sellayer_spin_${idx}`);
        clearSpinnerWatch(entry.layer);
        wireLayerUpdatingSpinner(entry.layer, spin).then((h) => setSpinnerWatch(entry.layer, h));

        // ✅ Stronger refresh (includes double refresh + micro scale nudge)
        await hardRefreshLayer(entry.layer);
    }

    function disableSelectionLayer(idx) {
        const entry = selectionLayers[idx];
        if (!entry) return;

        clearSpinnerWatch(entry.layer);

        setSelectionSpinner(idx, false);

        // ✅ never remove; just hide
        entry.layer.visible = false;

        updateSelectionToggleCheckbox(idx, false);

        if (activeSelectionLayer === entry.layer) {
            activeSelectionLayer = null;
            activeSelectionLayerView = null;
        }
    }

async function autoZoomToLayerMinVisible(layer) {
    if (!view || !layer) return;

    const minScale = Number(layer.minScale || 0);
    if (!minScale || !isFinite(minScale) || minScale <= 0) return;

    // ✅ VERY AGGRESSIVE ZOOM: 90% more zoomed in than minScale
    // Smaller scale number = more zoomed in
    // 0.10 = zoom in 10x closer than minimum required
    const nudgeFactor = 0.10;
    const targetScale = Math.max(1, Math.floor(minScale * nudgeFactor));

    if (view.scale > targetScale) {
        await view.goTo({ scale: targetScale }, { animate: true, duration: 450 });
    }
}

    async function ensureLayerVisibleAtScale(layer) {
        if (!view || !layer) return;

        const minScale = Number(layer.minScale || 0);
        const maxScale = Number(layer.maxScale || 0);

        // ArcGIS scale logic:
        // - If view.scale is GREATER than minScale (zoomed out too far), layer may not draw.
        // - If view.scale is LESS than maxScale (zoomed in too far), layer may not draw.
        let targetScale = null;

        if (minScale > 0 && isFinite(minScale) && view.scale > minScale) {
            // zoom IN a bit past minScale
            targetScale = Math.max(1, Math.floor(minScale * 0.90));
        } else if (maxScale > 0 && isFinite(maxScale) && view.scale < maxScale) {
            // zoom OUT a bit past maxScale
            targetScale = Math.ceil(maxScale * 1.10);
        }

        if (targetScale && isFinite(targetScale) && targetScale > 0) {
            // Keep center fixed so extent stays "basically" locked (scale-only nudge)
            await view.goTo(
                { center: view.center, scale: targetScale },
                { animate: true, duration: 250 }
            );
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

            let title = String(l.name || "");

            // Normalize PLSS naming (handles "Intersected" and "PLSS Intersected", etc.)
            title = title.replace(/intersected/ig, "Parcel");

            out.push({
                title,
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
            // ✅ avoid stale cached pjson making DOWN services look UP
            const res = await fetch(url, {
                credentials: "omit",
                signal: controller.signal,
                cache: "no-store"
            });

            if (!res.ok) throw new Error(`HTTP ${res.status}`);

            // Read as text first so we can give a better error if it isn't JSON
            const txt = await res.text();

            let json = null;
            try {
                json = JSON.parse(txt);
            } catch (e) {
                // Common if server returns an HTML error page with HTTP 200
                throw new Error("Non-JSON response (possible HTML error page)");
            }

            // ✅ ArcGIS often returns HTTP 200 with an error payload
            if (json && json.error) {
                const code = json.error.code != null ? json.error.code : "";
                const msg = json.error.message || "ArcGIS error";
                throw new Error(`ArcGIS error ${code}: ${msg}`);
            }

            return json;
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

    function normalizeUrlKey(u) {
        return String(u || "").replace(/\/+$/, "");
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
        if (!layer || !spinnerEl || !view) return null;

        try {
            await layer.when();
            const lv = await view.whenLayerView(layer);

            spinnerEl.classList.toggle("hidden", !lv.updating);

            const handle = lv.watch("updating", (isUpdating) => {
                spinnerEl.classList.toggle("hidden", !isUpdating);
            });

            return handle;
        } catch (e) {
            spinnerEl.classList.add("hidden");
            return null;
        }
    }

    function waitForViewStationary(timeoutMs = 1200) {
        if (!view) return Promise.resolve();
        if (view.stationary) return Promise.resolve();

        return new Promise(resolve => {
            const t = window.setTimeout(() => { try { h.remove(); } catch (e) { } resolve(); }, timeoutMs);
            const h = view.watch("stationary", (s) => {
                if (s) {
                    window.clearTimeout(t);
                    try { h.remove(); } catch (e) { }
                    resolve();
                }
            });
        });
    }


async function hardRefreshLayer(layer, { timeoutMs = 5000 } = {}) {
    if (!view || !layer) return;

    try { await layer.when(); } catch (e) { }

    let lv = null;
    try { lv = await view.whenLayerView(layer); } catch (e) { return; }
    if (!lv) return;

    // Wait for view to stop moving
    await waitForViewStationary(1500);

    // If suspended, wait for resume
    if (lv.suspended) {
        await new Promise(resolve => {
            const t = window.setTimeout(() => { try { h.remove(); } catch (e) { } resolve(); }, 2000);
            const h = lv.watch("suspended", (s) => {
                if (!s) {
                    window.clearTimeout(t);
                    try { h.remove(); } catch (e) { }
                    resolve();
                }
            });
        });
    }

    // Single refresh
    if (typeof lv.refresh === "function") lv.refresh();

    // Wait for updating to finish
    await new Promise(resolve => {
        const t = window.setTimeout(() => { try { h.remove(); } catch (e) { } resolve(); }, timeoutMs);
        const h = lv.watch("updating", (u) => {
            if (!u) {
                window.clearTimeout(t);
                try { h.remove(); } catch (e) { }
                resolve();
            }
        });
        if (!lv.updating) {
            window.clearTimeout(t);
            try { h.remove(); } catch (e) { }
            resolve();
        }
    });

    try { view.requestRender(); } catch (e) { }
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

            let title = l?.name ? String(l.name) : `Layer ${l.id}`;

            // Normalize PLSS naming (handles "Intersected" and "PLSS Intersected", etc.)
            title = title.replace(/intersected/ig, "Parcel");

            out.push({
                title,
                url: layerUrl
            });
        }

        return out;
    }

async function buildReportDisplayLayers() {
    if (!map) return;

    // Clear any previous (defensive)
    reportLayerViews.clear();

    for (const cfg of (config.reportLayers || [])) {
        const key = normalizeUrlKey(cfg.url);
        if (!key) continue;

        // Skip if already built
        if (reportLayerViews.has(key)) continue;

        // FeatureServer root: expand to polygon sublayers
        if (isFeatureServerRoot(key)) {
            const subs = await expandFeatureServerToPolygonSublayers(key);

            const layers = subs.map(sl => new FeatureLayer({
                url: sl.url,
                title: `${cfg.title}: ${sl.title}`,
                outFields: ["*"],
                visible: false,
                renderer: getPresetRenderer("report", cfg) || undefined
            }));

            layers.forEach(l => map.add(l));
            reportLayerViews.set(key, layers);
            continue;
        }

        // MapServer root: expand to polygon sublayers
        if (isMapServerRoot(key)) {
            const subs = await expandMapServerToSublayers(key, { polygonOnly: true });

            const layers = subs.map(sl => new FeatureLayer({
                url: sl.url,
                title: `${cfg.title}: ${sl.title}`,
                outFields: ["*"],
                visible: false,
                renderer: getPresetRenderer("report", cfg) || undefined
            }));

            layers.forEach(l => map.add(l));
            reportLayerViews.set(key, layers);
            continue;
        }

        // Normal single layer
        const lyr = new FeatureLayer({
            url: key,
            title: cfg.title,
            outFields: ["*"],
            visible: false,
            renderer: getPresetRenderer("report", cfg) || undefined
        });

        map.add(lyr);
        reportLayerViews.set(key, lyr);
    }

    ensureAoiOnTop(map);
}




function clearAll() {
    selectionGeom = null;
    aoiSourceObjectId = null;
    aoiSourceObjectIdField = null;
    aoiSource = null;
    aoiSourceLayerTitle = null;
    aoiSourceLayerUrl = null;
    aoiSourceFeature = null;
    
    // Clear results
    resultsEl.innerHTML = "";
    exportAllBtn.disabled = true;
    lastReportRowsByLayer = [];
    
    // Clear map outputs
    if (visualReportOutputsEl) visualReportOutputsEl.innerHTML = "";
    if (visualReportMapWrapEl) visualReportMapWrapEl.classList.add("hidden");
    
    // Clear final report
    cachedFinalReportHtml = null;
    if (viewReportBtn) viewReportBtn.disabled = true;

    if (aoiLayer) aoiLayer.removeAll();
    aoiGraphic = null;

    runBtn.disabled = true;
    setStatus("cleared");
    coverageCache.clear();
    coverageAoiKey = "";
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

        // ✅ #3: clean up any old layerView.updating watchers before rebuilding DOM
        try {
            // Selection layers
            (selectionLayers || []).forEach(e => {
                if (e?.layer) clearSpinnerWatch(e.layer);
            });

            // Report layers (single layer or expanded root -> array)
            for (const v of reportLayerViews.values()) {
                if (Array.isArray(v)) {
                    v.forEach(x => { if (x) clearSpinnerWatch(x); });
                } else if (v) {
                    clearSpinnerWatch(v);
                }
            }
        } catch (e) {
            // best-effort cleanup
        }

        // ---- Selection layers (already on map): toggle visibility
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
        const spin = document.getElementById(`sellayer_spin_${i}`);  // ✅ CORRECT

        if (cb.checked) {
            if (spin) spin.classList.remove("hidden");

            const isOnMapNow = map.layers.includes(e.layer);
            if (!isOnMapNow) map.add(e.layer);

            e.layer.visible = true;
            ensureAoiOnTop(map);
            await hardRefreshLayer(e.layer);

            clearSpinnerWatch(e.layer);
            wireLayerUpdatingSpinner(e.layer, spin).then((h) => setSpinnerWatch(e.layer, h));

            if (!activeSelectionLayer) {
                await setActiveSelectionLayerByIndex(i);
            }

        } else {
            if (spin) spin.classList.add("hidden");

            clearSpinnerWatch(e.layer);

            // ✅ Don't remove; just hide
            e.layer.visible = false;
            updateSelectionToggleCheckbox(i, false);

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
});

        // ---- Report layers (ALWAYS included in report): toggle ONLY map visibility
        // If a report URL is a FeatureServer ROOT (no /0 etc.), it cannot be drawn directly.
        // We will show it in the list but disable the checkbox to avoid confusion.
        reportLayerTogglesEl.innerHTML = (config.reportLayers || []).map((l, i) => {
            const isRoot = isFeatureServerRoot(l.url) || isMapServerRoot(l.url);
            const key = normalizeUrlKey(l.url);
            const existing = reportLayerViews.get(key);

            // If existing is an array (expanded root), consider it checked if any layer exists
            const isChecked = false;
                Array.isArray(existing) ? (existing.length > 0) :
                    existing ? !!existing.visible :
                        false;

            const checked = isChecked ? "checked" : "";

            // ✅ Do NOT disable FeatureServer roots anymore (we will expand them to drawable polygon sublayers)
            const disabled = "";

            // Update note text
            const note = isRoot ? ` <span class="small">(expands to polygon sublayers)</span>` : "";

            return `
                <div class="toggle-row">
                    <input type="checkbox" id="rptlayer_${i}" ${checked} ${disabled} />
                    <span class="layer-swatch layer-swatch-report" aria-hidden="true" title="Report layer"></span>
                    <label class="toggle-name" for="rptlayer_${i}">${escapeHtml(l.title)}${note}</label>
                    <span id="rptlayer_spin_${i}" class="layer-spinner hidden" aria-label="loading"></span>
                </div>
            `;
        }).join("");

        (config.reportLayers || []).forEach((l, i) => {
            const cb = document.getElementById(`rptlayer_${i}`);
            if (!cb) return;

            const key = normalizeUrlKey(l.url);

cb.addEventListener("change", async () => {
    const spin = document.getElementById(`rptlayer_spin_${i}`);
    const key = normalizeUrlKey(l.url);

    // Get existing drawn layers for this report entry (single or array)
    const existing = reportLayerViews.get(key);

    // Helper to set visibility on single/array
    const setVisible = (obj, vis) => {
        if (!obj) return;
        if (Array.isArray(obj)) obj.forEach(x => { try { x.visible = vis; } catch (e) { } });
        else { try { obj.visible = vis; } catch (e) { } }
    };

    if (cb.checked) {
        if (spin) spin.classList.remove("hidden");  // Show spinner

        // Show existing layers
        setVisible(existing, true);

        // ✅ Wire spinner watches for report layers (THIS WAS MISSING!)
        if (Array.isArray(existing)) {
            for (const lyr of existing) {
                clearSpinnerWatch(lyr);
                wireLayerUpdatingSpinner(lyr, spin).then((h) => setSpinnerWatch(lyr, h));
            }
        } else if (existing) {
            clearSpinnerWatch(existing);
            wireLayerUpdatingSpinner(existing, spin).then((h) => setSpinnerWatch(existing, h));
        }

        // Don't hide the spinner immediately - let the watch handle it
    } else {
        if (spin) spin.classList.add("hidden");
        
        // ✅ Clear spinner watches when hiding
        if (Array.isArray(existing)) {
            existing.forEach(lyr => clearSpinnerWatch(lyr));
        } else if (existing) {
            clearSpinnerWatch(existing);
        }
        
        setVisible(existing, false);
    }

    ensureAoiOnTop(map);
    try { view.requestRender(); } catch (e) { }
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

                // ✅ basic “looks like ArcGIS REST” sanity
                // (many valid pjson payloads include currentVersion)
                if (pjson == null || (pjson.currentVersion == null && pjson.layers == null && pjson.type == null)) {
                    throw new Error("Unexpected JSON (missing expected ArcGIS REST fields)");
                }

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

    function makeTable(features, maxFields, totalCount) {
        if (!features || !features.length) return `<div class="small">No sample features fetched.</div>`;

        const picked = sampleWithoutReplacement(features, 4);

        const attrs0 = picked[0].attributes || {};
        const keysAll = Object.keys(attrs0);

        // ✅ Show ALL columns so horizontal scrolling actually reveals more fields
        const keys = keysAll;

        // ✅ “Default view” target: ~5 columns visible in the panel before scrolling
        const defaultVisibleCols = 5;
        const colPx = 140; // keep aligned with your CSS td max-width
        const minTableWidth = Math.max(520, keys.length * colPx);

        const th = keys.map(k => `<th title="${escapeHtml(k)}">${escapeHtml(k)}</th>`).join("");

        const rows = picked.map(f => {
            const a = f.attributes || {};
            const tds = keys.map(k => {
                const raw = (a[k] == null) ? "" : String(a[k]);

                // truncate values longer than the *column name*
                const maxLen = Math.max(4, String(k).length);
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

        // “more rows” message stays the same
        let moreRowHtml = "";
        const total = Number(totalCount || 0);
        const shown = picked.length;

        if (total > shown) {
            const more = total - shown;
            const msg = `… ${more} more row${more === 1 ? "" : "s"} (see FULL export)`;
            moreRowHtml = `<tr><td colspan="${keys.length}" class="small" style="opacity:.8;">${escapeHtml(msg)}</td></tr>`;
        }

        // ✅ Hint about horizontal scrolling + “first 5 columns by default”
        const colHint = (keys.length > defaultVisibleCols)
            ? `<div class="small table-hint">Table has ${keys.length} columns — scroll → for more.</div>`
            : "";

        return `
        <div class="table-wrap">
            <table class="result-table" style="min-width:${minTableWidth}px">
            <thead><tr>${th}</tr></thead>
            <tbody>${rows}${moreRowHtml}</tbody>
            </table>
        </div>
        ${colHint}
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

    function getReportGeometry() {
        if (!selectionGeom) return null;

        // Only shrink when AOI was selected from PLSS (boundary-touch neighbors)
        if (aoiSource !== "select") return selectionGeom;

        // Only shrink polygons
        if (selectionGeom.type !== "polygon") return selectionGeom;

        try {
            // Shrink inward ~1 meter so touching neighbors don't count as intersects
            const shrunk = geometryEngine.geodesicBuffer(selectionGeom, -1, "meters");
            return shrunk || selectionGeom;
        } catch (e) {
            console.warn("AOI shrink failed; using original geometry", e);
            return selectionGeom;
        }
    }


    // ---------- Query logic ----------
    async function querySingleLayer(layerUrl, layerTitle, geom, spatialRel = "intersects", options = {}) {
        const applyTouchFilter = !!options.applyTouchFilter;
        const objectId = options.objectId ?? null;
        const objectIdField = options.objectIdField || "OBJECTID";

        const layer = new FeatureLayer({ url: layerUrl, outFields: ["*"] });

        const q = layer.createQuery();
        q.outFields = ["*"];

        // ✅ Special case: AOI-source layer should return the exact clicked feature (1 row)
        if (objectId != null) {
            // Ensure layer is loaded so objectIdField is correct
            await layer.load();

            const trueOidField = layer.objectIdField || objectIdField || "OBJECTID";

            // Coerce OID to number if it looks numeric (ArcGIS OIDs are typically numeric)
            const oidNum = Number(objectId);
            const oidIsNumeric = Number.isFinite(oidNum);

            // Use WHERE instead of objectIds (more robust across services)
            q.where = oidIsNumeric
                ? `${trueOidField} = ${oidNum}`
                : `${trueOidField} = '${String(objectId).replace(/'/g, "''")}'`;

            q.returnGeometry = false;
            q.outFields = ["*"];

            const fs = await layer.queryFeatures(q);
            const feats = fs?.features ?? [];

            // Optional: debug if it ever happens again
            // console.log("AOI-source query", { layerUrl, trueOidField, objectId, oidNum, featsLen: feats.length });

            return {
                title: layerTitle,
                url: layerUrl,
                count: feats.length,
                features: feats,
                layer,
                exportQuery: q
            };
        }


        // Default: geometry intersect behavior for normal report layers
        q.geometry = geom;
        q.spatialRelationship = spatialRel;
        q.returnGeometry = applyTouchFilter; // only return geometry when filtering touching-only

        const count = await layer.queryFeatureCount(q);

        const maxSamples = config.report?.maxSampleFeaturesPerLayer ?? 25;
        let features = [];

        if (count > 0 && maxSamples > 0) {
            const q2 = q.clone();
            q2.num = Math.min(maxSamples, 2000);
            const fs = await layer.queryFeatures(q2);
            const raw = fs?.features ?? [];
            features = applyTouchFilter ? filterTouchingOnly(raw, geom) : raw;
        }

        return { title: layerTitle, url: layerUrl, count, features, layer, exportQuery: q };
    }

// ========================================
// REFACTORED ANALYSIS FUNCTIONS
// ========================================

// Main orchestrator - runs ALL analysis steps
async function runAnalysis() {
    const myOp = startReportOp();

    const reportGeom = getReportGeometry();
    if (!reportGeom) { endReportOp(myOp); return; }

    setBusy(true);

    try {
        // Step 1: Data Check
        setStatus("Running analysis... (checking services)");
        await refreshServicesTab();

        if (isReportCanceled(myOp)) {
            setStatus("canceled");
            return;
        }

        // Step 2: Query all layers
        setStatus("Running analysis... (querying layers)");
        await queryAllLayers(reportGeom, myOp);

        if (isReportCanceled(myOp)) {
            setStatus("canceled");
            return;
        }

        // Step 3: Generate map screenshots
        setStatus("Running analysis... (generating maps)");
        await generateVisualReportData(myOp);

        if (isReportCanceled(myOp)) {
            setStatus("canceled");
            return;
        }

        // Step 4: Build final report HTML
        setStatus("Running analysis... (building report)");
        await buildFinalReportHtml();

        // Enable "View Report" button
        if (viewReportBtn) viewReportBtn.disabled = false;

        setStatus("Analysis complete!");
    } catch (e) {
        console.error(e);
        setStatus("Analysis failed (see console)");
    } finally {
        setBusy(false);
        endReportOp(myOp);
    }
}

// Extracted query logic (was: runReport)
async function queryAllLayers(reportGeom, myOp) {
    const toolLabel = plssToolLabel(aoiSourcePlssTool);

    resultsEl.innerHTML = "";
    exportAllBtn.disabled = true;
    lastReportRowsByLayer = [];

    const combinedCfgs = [
        ...(config.reportLayers || [])
    ];

    if (plssStateLayerUrl) {
        combinedCfgs.push({
            title: "PLSS: State Boundaries",
            url: plssStateLayerUrl
        });
    }

    if (aoiSource === "draw" && plssParcelLayerUrl) {
        combinedCfgs.push({
            title: "PLSS: Parcel",
            url: String(plssParcelLayerUrl).replace(/\/+$/, "")
        });
    }

    const byUrl = new Map();

    for (const l of combinedCfgs) {
        const urlKey = String(l?.url || "").replace(/\/+$/, "");
        if (!urlKey) continue;

        byUrl.set(urlKey, { title: l.title, url: urlKey });
    }

    const reportCfgs = Array.from(byUrl.values());
    const expandedTargets = [];

    if (aoiSource === "select" && aoiSourceLayerUrl) {
        if (aoiSourceObjectId == null) {
            console.warn("AOI source objectId is null; AOI source table will not be 1-row exact.");
        }

        const toolLabel =
            (aoiSourcePlssTool === "township") ? "Township" :
                (aoiSourcePlssTool === "section") ? "Section" :
                    (aoiSourcePlssTool === "intersected") ? "Parcel" :
                        "PLSS";

        expandedTargets.push({
            title: `AOI Source (${toolLabel})`,
            url: String(aoiSourceLayerUrl).replace(/\/+$/, ""),
            __pinnedAoiFeature: aoiSourceFeature || null
        });
    }

    for (const cfg of reportCfgs) {
        const url = String(cfg.url || "");

        if (isMapServerRoot(url) && isPlssLayerTitleOrUrl(cfg.title, url)) {
            continue;
        }

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
        if (isReportCanceled(myOp)) {
            setStatus("canceled");
            break;
        }

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
            const plss = isPlssLayerTitleOrUrl(t.title, t.url);
            const targetIsPlssIntersected = isPlssIntersectedLayerTitle(t.title);

            const spatialRel =
                (targetIsPlssIntersected && (aoiSourcePlssTool === "township" || aoiSourcePlssTool === "section"))
                    ? "within"
                    : "intersects";

            if (t.__pinnedAoiFeature) {
                const f = t.__pinnedAoiFeature;
                const feats = f ? [f] : [];
                const r = {
                    title: t.title,
                    url: t.url,
                    count: feats.length,
                    features: feats,
                    layer: null,
                    exportQuery: null
                };

                const rows = flattenAttributes(r.features);

                lastReportRowsByLayer.push({
                    title: r.title,
                    url: r.url,
                    count: r.count,
                    rows,
                    _layer: null,
                    _exportQuery: null,
                    fullRows: rows
                });

                const maxFields = (config.report && config.report.maxFieldsInTable) ? config.report.maxFieldsInTable : 8;
                const tableHtml = feats.length
                    ? makeTable(feats, maxFields, r.count)
                    : `<div class="small">No sample rows.</div>`;

                cards.push(`
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(r.title)}</div>
                <div class="badge">
                count: <b>${r.count}</b>
                </div>
            </div>
                <div class="small mono">
                <a href="${escapeHtml(r.url)}" target="_blank" rel="noopener">Service URL</a>
                </div>
            <div style="margin-top:8px;">
              ${tableHtml}
                ${(r.count > 0) ? `
                <div class="row" style="margin-top:8px;">
                    <button class="btn subtle" data-export="${escapeHtml(r.title)}">
                    Export FULL CSV
                    </button>
                </div>
                ` : ``}
            </div>
          </div>
        `);

                setStatus(`Running analysis... (querying ${i + 1}/${expandedTargets.length})`);
                continue;
            }


            const r = await querySingleLayer(
                t.url,
                t.title,
                reportGeom,
                spatialRel,
            );
            const rows = flattenAttributes(r.features);

            lastReportRowsByLayer.push({
                title: r.title,
                url: r.url,
                count: r.count,
                rows,
                _layer: r.layer,
                _exportQuery: r.exportQuery,
                fullRows: null
            });


            const maxFields = (config.report && config.report.maxFieldsInTable) ? config.report.maxFieldsInTable : 8;
            const tableHtml = (r.features && r.features.length)
                ? makeTable(r.features, maxFields, r.count)
                : `<div class="small">No sample rows.</div>`;

            cards.push(`
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(r.title)}</div>
                <div class="badge">
                count: <b>${r.count}</b>
                ${(config.report?.maxExportFeatures && r.count > config.report.maxExportFeatures)
                    ? `<span class="small" style="margin-left:8px; opacity:.85;">(FULL export capped at ${config.report.maxExportFeatures})</span>`
                    : ``
                }
                </div>
            </div>
                <div class="small mono">
                    <a href="${escapeHtml(r.url)}" target="_blank" rel="noopener">Service URL</a>
                </div>
            
            <div style="margin-top:8px;">
            ${tableHtml}
            ${(r.count > 0) ? `
                <div class="row" style="margin-top:8px;">
                <button class="btn subtle" data-export="${escapeHtml(r.title)}">
                Export FULL CSV
                </button>
                </div>
            ` : ``}
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

        setStatus(`Running analysis... (querying ${i + 1}/${expandedTargets.length})`);
    }

    renderResults(cards.join(""));
    wireExportButtons();
    exportAllBtn.disabled = (lastReportRowsByLayer.length === 0);
    renderVisualSummary();
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

    // ---------- Coverage stats (AOI acres + % covered by layer) ----------
    const SQM_PER_ACRE = 4046.8564224;
    const coverageCache = new Map(); // key: `${aoiKey}||${layerUrl}` -> { acresCovered, pctAoiCovered }
    let coverageAoiKey = "";         // changes whenever AOI changes

    function getAoiKey(geom) {
        // stable-enough signature: extent + rounded area
        try {
            const ex = geom?.extent;
            const area = geometryEngine.geodesicArea(geom, "square-meters");
            return [
                ex?.xmin, ex?.ymin, ex?.xmax, ex?.ymax,
                Math.round(area)
            ].join("|");
        } catch (e) {
            return String(Date.now());
        }
    }

    function resetCoverageCacheForAoi(geom) {
        coverageCache.clear();
        coverageAoiKey = getAoiKey(geom);
    }

    function formatNumber(n, digits = 2) {
        const x = Number(n);
        if (!isFinite(x)) return "";
        return x.toLocaleString(undefined, { maximumFractionDigits: digits, minimumFractionDigits: digits });
    }

    async function queryAllFeaturesPagedWithGeometry(layer, baseQuery, pageSize, maxExportFeatures) {
        const all = [];
        let offset = 0;

        while (true) {
            const q = baseQuery.clone();
            q.num = pageSize;
            q.start = offset;
            q.returnGeometry = true;              // ✅ geometry required for area
            q.outFields = [];                    // we only need geometry
            q.outSpatialReference = view?.spatialReference;

            const fs = await layer.queryFeatures(q);
            const feats = (fs && fs.features) ? fs.features : [];

            all.push(...feats);

            if (feats.length < pageSize) break;
            offset += pageSize;

            if (maxExportFeatures && all.length >= maxExportFeatures) break;
        }

        return all;
    }

    function unionGeomsChunked(geoms) {
        // geometryEngine.union can choke on huge arrays; do it in chunks.
        const CHUNK = 25;
        let acc = null;

        for (let i = 0; i < geoms.length; i += CHUNK) {
            const chunk = geoms.slice(i, i + CHUNK).filter(Boolean);
            if (!chunk.length) continue;

            const u = geometryEngine.union(chunk);
            if (!acc) acc = u;
            else acc = geometryEngine.union([acc, u]);
        }

        return acc;
    }

    async function computeLayerCoverageStats(item, aoiGeom) {
        // Returns: { acresCovered, pctAoiCovered }
        if (!item || !item._layer || !item._exportQuery || !aoiGeom) return null;

        const layerUrlKey = String(item.url || "").replace(/\/+$/, "");
        const aoiKey = coverageAoiKey || getAoiKey(aoiGeom);
        const cacheKey = `${aoiKey}||${layerUrlKey}`;

        if (coverageCache.has(cacheKey)) {
            return coverageCache.get(cacheKey);
        }


        // AOI area (sqm)
        let aoiAreaSqm = 0;
        try {
            aoiAreaSqm = Math.max(0, geometryEngine.geodesicArea(aoiGeom, "square-meters"));
        } catch (e) {
            aoiAreaSqm = 0;
        }
        if (!aoiAreaSqm) return { acresCovered: 0, pctAoiCovered: 0 };

        const pageSize = config.report?.pageSize ?? 1000;
        const maxExport = config.report?.maxExportFeatures ?? 50000;

        // Page through intersecting features WITH geometry
        const feats = await queryAllFeaturesPagedWithGeometry(item._layer, item._exportQuery, pageSize, maxExport);

        // Intersect each feature with AOI and collect intersection geometries
        const interGeoms = [];
        for (const f of feats) {
            const g = f?.geometry;
            if (!g) continue;

            try {
                const inter = geometryEngine.intersect(aoiGeom, g);
                if (!inter) continue;

                // Drop pure edge-touch (0 area) intersections
                const area = geometryEngine.geodesicArea(inter, "square-meters");
                if (area <= 0) continue;

                interGeoms.push(inter);
            } catch (e) {
                // ignore bad geometries
            }
        }

        if (!interGeoms.length) return { acresCovered: 0, pctAoiCovered: 0 };

        // Union intersections to avoid double-counting overlap
        let unionGeom = null;
        try {
            unionGeom = unionGeomsChunked(interGeoms);
        } catch (e) {
            unionGeom = null;
        }

        let coveredSqm = 0;
        try {
            coveredSqm = unionGeom
                ? Math.max(0, geometryEngine.geodesicArea(unionGeom, "square-meters"))
                : 0;
        } catch (e) {
            coveredSqm = 0;
        }

        const acresCovered = coveredSqm / SQM_PER_ACRE;
        const pctAoiCovered = Math.min(100, Math.max(0, (coveredSqm / aoiAreaSqm) * 100));

        const out = { acresCovered, pctAoiCovered };
        coverageCache.set(cacheKey, out);
        return out;
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

async function generateVisualReportData(myOp) {

        if (!view) return;

        if (!selectionGeom) {
            setVisualStatus("No AOI selected.");
            return;
        }

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            setVisualStatus("No query results available.");
            return;
        }

        setVisualStatus("Generating maps for intersecting layers…");

        if (visualReportMapWrapEl) visualReportMapWrapEl.classList.add("hidden");
        if (visualReportOutputsEl) visualReportOutputsEl.innerHTML = "";

        try {
            // Only layers with real intersect hits AND usable query objects
            const targets = lastReportRowsByLayer
                .filter(x => (x?.count || 0) > 0)
                .filter(x => x?._layer && x?._exportQuery); // excludes pinned AOI source etc.

            if (!targets.length) {
                setVisualStatus("No intersecting layers to map (all counts are 0).");
                if (visualReportMapWrapEl) visualReportMapWrapEl.classList.remove("hidden");
                return;
            }

            // Zoom to AOI with padding once (we’ll keep the view there)
            const paddingFactor = config?.visualReport?.paddingFactor ?? 1.25;
            const width = config?.visualReport?.screenshotWidth ?? 1400;

            // 🔒 Compute and lock a single extent for ALL screenshots
            let fixedExtent = null;
            const ext = selectionGeom?.extent;

            if (ext && ext.expand) {
                fixedExtent = ext.expand(paddingFactor);
                await view.goTo(fixedExtent, { animate: true, duration: 450 });
            } else {
                await view.goTo(selectionGeom, { animate: true, duration: 450 });
            }

            // Snapshot current layer visibility so we can restore after each screenshot
            const allLayers = view.map.layers.toArray();

            const visSnapshot = allLayers.map(l => ({ layer: l, visible: l.visible }));

            // Helper to hide everything except AOI + basemap overlay + a temp target layer
            function setVisibilityForScreenshot(tempLayer) {
                for (const l of allLayers) {
                    // Keep AOI layer visible
                    if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }

                    // Keep SMA overlay visible (your TileLayer at bottom) if present
                    // (We don’t have the variable here; keep TileLayers visible by default.)
                    if (l?.type === "tile") { l.visible = true; continue; }

                    // Hide everything else (selection layers, other report layers, etc.)
                    l.visible = false;
                }

                if (tempLayer) tempLayer.visible = true;
                ensureAoiOnTop(view.map);
            }

            function restoreVisibility() {
                visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
                ensureAoiOnTop(view.map);
            }

            // AOI area in acres (used for context)
            let aoiAcres = 0;
            try {
                const aoiSqm = Math.max(0, geometryEngine.geodesicArea(selectionGeom, "square-meters"));
                aoiAcres = aoiSqm / SQM_PER_ACRE;
            } catch (e) {
                aoiAcres = 0;
            }

            const outCards = [];

            for (let i = 0; i < targets.length; i++) {
                if (isReportCanceled(myOp)) { 
                    setVisualStatus("canceled");
                    break;
                }

                const item = targets[i];

                setVisualStatus(`Generating map ${i + 1} / ${targets.length}…`);

                // Create a temporary layer for this URL, regardless of toggle state
                const temp = new FeatureLayer({
                    url: item.url,
                    title: item.title,
                    outFields: ["*"],
                    visible: true,
                    renderer: getPresetRenderer("report", layerCfgByUrl.get(item.url)?.cfg || null) || undefined
                });

                // 🔒 Prevent scale-based rendering rules from forcing view scale changes
                temp.minScale = 0;
                temp.maxScale = 0;

                // Add temp, hide everything else, screenshot, then remove temp
                view.map.add(temp);
                try {
                    setVisibilityForScreenshot(temp);

                    // Wait for layer to load
                    try { await temp.when(); } catch (e) { }

                    // ✅ Wait until layerView is not suspended AND not updating (best effort)
                    try {
                        const lv = await view.whenLayerView(temp);

                        // Wait for suspended -> false OR timeout
                        if (lv?.suspended) {
                            await new Promise(resolve => {
                                const h = lv.watch("suspended", (s) => {
                                    if (!s) { h.remove(); resolve(); }
                                });
                                window.setTimeout(() => { try { h.remove(); } catch (e) { } resolve(); }, 4000);
                            });
                        }

                        // Wait for updating -> false OR timeout
                        if (lv?.updating) {
                            await new Promise(resolve => {
                                const h = lv.watch("updating", (u) => {
                                    if (!u) { h.remove(); resolve(); }
                                });
                                window.setTimeout(() => { try { h.remove(); } catch (e) { } resolve(); }, 6000);
                            });
                        }
                    } catch (e) { }

                    // 🔒 Re-apply locked extent to guarantee identical framing
                    if (fixedExtent) {
                        await view.goTo(fixedExtent, { animate: false });
                    }

                    const ss = await view.takeScreenshot({ format: "png", quality: 100, width });
                    const dataUrl = ss?.dataUrl;
                    if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

                    // Compute coverage stats (acres + % AOI covered)
                    const cov = await computeLayerCoverageStats(item, selectionGeom);

                    // Render a card for this layer
                    const acresCovered = cov ? cov.acresCovered : 0;
                    const pctCovered = cov ? cov.pctAoiCovered : 0;

                    outCards.push(`
                  <div class="visual-output-card">
                    <div class="visual-output-title">${escapeHtml(item.title)}</div>
                    <img class="visual-output-img" src="${dataUrl}" alt="AOI + ${escapeHtml(item.title)}" />
                    <div class="visual-output-meta">
                      <table>
                        <tr><td>AOI area</td><td><span class="mono">${formatNumber(aoiAcres, 2)}</span> acres</td></tr>
                        <tr><td>Intersecting features</td><td><span class="mono">${escapeHtml(String(item.count || 0))}</span></td></tr>
                        <tr><td>AOI covered by layer</td><td><span class="mono">${formatNumber(acresCovered, 2)}</span> acres</td></tr>
                        <tr><td>% AOI covered</td><td><span class="mono">${formatNumber(pctCovered, 2)}</span>%</td></tr>
                      </table>
                    </div>
                  </div>
                `);

                } finally {
                    // Remove temp layer and restore visibility
                    try { view.map.remove(temp); } catch (e) { }
                    restoreVisibility();
                }
            }

            if (visualReportOutputsEl) visualReportOutputsEl.innerHTML = outCards.join("");
            if (visualReportMapWrapEl) visualReportMapWrapEl.classList.remove("hidden");

            // Keep your existing summary panel behavior
            renderVisualSummary();

            setVisualStatus("Maps generated.");
        } catch (e) {
            console.error(e);
            setVisualStatus("Failed to generate maps (see console).");
        }    
    }


    // ---------- Final Report (printable HTML in new tab) ----------
    function openHtmlInNewTab(htmlString) {
        const blob = new Blob([htmlString], { type: "text/html;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const win = window.open(url, "_blank", "noopener,noreferrer");
        // Cleanup the blob URL later (give the new tab time to load)
        window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
        return win;
    }

    function formatDateTimeForReport(d = new Date()) {
        try {
            return d.toLocaleString();
        } catch (e) {
            return d.toString();
        }
    }

    function getAoiSummaryForReport(aoiAcres) {
        const src = aoiSource === "draw" ? "Drawn AOI" : "Selected AOI";
        const tool = aoiSource === "select" ? plssToolLabel(aoiSourcePlssTool) : "";
        const srcDetail = (aoiSource === "select" && tool) ? ` (${tool})` : "";
        const layer = aoiSourceLayerTitle ? ` • Source layer: ${aoiSourceLayerTitle}` : "";
        return `${src}${srcDetail} • AOI area: ${formatNumber(aoiAcres, 2)} acres${layer}`;
    }

    function buildFinalReportHtmlDoc({ title, createdAt, aoiSummary, totalsHtml, sectionsHtml }) {
        const safeTitle = escapeHtml(title || "Final Report");

        // Minimal, clean print CSS (no external deps)
        return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>${safeTitle}</title>
  <style>
    :root{
      --fg:#111;
      --muted:#666;
      --border:#e6e6e6;
      --bg:#fff;
    }
    html,body{ margin:0; padding:0; background:var(--bg); color:var(--fg); font-family: system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif; }
    .wrap{ max-width: 980px; margin: 0 auto; padding: 28px 22px 60px; }
    h1{ font-size: 22px; margin: 0 0 6px; }
    .meta{ font-size: 12px; color: var(--muted); margin-bottom: 14px; }
    .aoi{ font-size: 13px; padding: 10px 12px; border:1px solid var(--border); border-radius: 10px; background:#fafafa; }
    .totals{ margin-top: 12px; }
    .totals .row{ display:flex; gap:12px; flex-wrap:wrap; margin-top:10px; }
    .pill{ border:1px solid var(--border); border-radius:999px; padding:6px 10px; font-size: 12px; background:#fff; }
    .section{ margin-top: 18px; padding-top: 14px; border-top: 1px solid var(--border); }
    .section h2{ margin:0 0 8px; font-size: 16px; }
    .section .sub{ font-size: 12px; color: var(--muted); margin-bottom: 10px; }
    .map{ width:100%; border:1px solid var(--border); border-radius: 12px; overflow:hidden; background:#fff; }
    .map img{ display:block; width:100%; height:auto; }
    table.metaTbl{ width:100%; border-collapse: collapse; margin-top:10px; font-size: 12px; }
    table.metaTbl td{ padding: 6px 8px; border-bottom: 1px solid var(--border); }
    table.metaTbl td:first-child{ color: var(--muted); width: 220px; }
    .actions{ margin-top: 14px; display:flex; gap:10px; flex-wrap: wrap; }
    .btn{
      display:inline-block; border:1px solid var(--border); background:#fff; border-radius: 10px;
      padding:8px 10px; font-size:12px; text-decoration:none; color:var(--fg);
    }
    .btn:hover{ background:#f4f4f4; }
    .hint{ font-size: 12px; color: var(--muted); margin-top: 8px; }

    /* Print */
    @media print{
      .actions, .hint{ display:none !important; }
      .wrap{ max-width: none; padding: 0.5in; }
      .section{ break-inside: avoid; }
      .pagebreak{ break-after: page; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>${safeTitle}</h1>
    <div class="meta">Created: ${escapeHtml(createdAt || "")}</div>

    <div class="aoi">${escapeHtml(aoiSummary || "")}</div>

    <div class="actions">
      <a class="btn" href="javascript:window.print()">Print / Save as PDF</a>
    </div>
    <div class="hint">Tip: Use your browser print dialog → “Save as PDF” later, when you’re ready.</div>

    <div class="totals">
      ${totalsHtml || ""}
    </div>

    ${sectionsHtml || ""}
  </div>
</body>
</html>`;
    }

    async function buildFinalReportHtml() {
        if (!view) return;

        if (!selectionGeom) {
            setStatus("Select or draw an AOI first.");
            return;
        }

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            setStatus("Run the report first (Tables tab) so we know which layers intersect.");
            return;
        }

        if (finalReportStatus) finalReportStatus.textContent = "Building report...";
        try {
            // Only layers with real hits and usable query objects (skip pinned AOI source card)
            const targets = lastReportRowsByLayer
                .filter(x => (x?.count || 0) > 0)
                .filter(x => x?._layer && x?._exportQuery);

            // AOI area in acres
            let aoiAcres = 0;
            try {
                const aoiSqm = Math.max(0, geometryEngine.geodesicArea(selectionGeom, "square-meters"));
                aoiAcres = aoiSqm / SQM_PER_ACRE;
            } catch (e) {
                aoiAcres = 0;
            }

            // Totals summary (counts)
            const totalLayers = lastReportRowsByLayer.length;
            const layersWithHits = lastReportRowsByLayer.filter(x => (x.count || 0) > 0).length;
            const totalHits = lastReportRowsByLayer.reduce((sum, x) => sum + (x.count || 0), 0);

            const totalsHtml = `
          <div class="row">
            <div class="pill">Layers queried: <b>${escapeHtml(String(totalLayers))}</b></div>
            <div class="pill">Layers with hits: <b>${escapeHtml(String(layersWithHits))}</b></div>
            <div class="pill">Total intersecting features: <b>${escapeHtml(String(totalHits))}</b></div>
          </div>
        `;

            // Zoom to AOI once
            const paddingFactor = config?.visualReport?.paddingFactor ?? 1.25;
            const width = config?.visualReport?.screenshotWidth ?? 1400;

            // 🔒 Compute and lock a single extent for ALL screenshots
            let fixedExtent = null;
            const ext = selectionGeom?.extent;

            if (ext && ext.expand) {
                fixedExtent = ext.expand(paddingFactor);
                await view.goTo(fixedExtent, { animate: true, duration: 450 });
            } else {
                await view.goTo(selectionGeom, { animate: true, duration: 450 });
            }

            // Snapshot vis so we can restore after each screenshot
            const allLayers = view.map.layers.toArray();
            const visSnapshot = allLayers.map(l => ({ layer: l, visible: l.visible }));

            function setVisibilityForScreenshot(tempLayer) {
                for (const l of allLayers) {
                    // Keep AOI visible
                    if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }
                    // Keep tile overlays (SMA) visible
                    if (l?.type === "tile") { l.visible = true; continue; }
                    // Hide everything else
                    l.visible = false;
                }
                if (tempLayer) tempLayer.visible = true;
                ensureAoiOnTop(view.map);
            }

            function restoreVisibility() {
                visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
                ensureAoiOnTop(view.map);
            }

            // Build per-layer sections
            let sectionsHtml = "";

            if (!targets.length) {
                sectionsHtml = `
              <div class="section">
                <h2>No intersecting layers</h2>
                <div class="sub">(All layer counts are 0.)</div>
              </div>
            `;
            } else {
                for (let i = 0; i < targets.length; i++) {

                    const item = targets[i];
                    setStatus(`building final report… (${i + 1}/${targets.length})`);

                    // Temp layer (so it renders regardless of user toggles)
                    const temp = new FeatureLayer({
                        url: item.url,
                        title: item.title,
                        outFields: ["*"],
                        visible: true,
                        renderer: getPresetRenderer("report", layerCfgByUrl.get(item.url)?.cfg || null) || undefined
                    });

                    // 🔒 Prevent scale-based rendering rules from forcing view scale changes
                    temp.minScale = 0;
                    temp.maxScale = 0;

                    // Add temp, hide everything else, screenshot, then remove temp
                    view.map.add(temp);
                    try {
                        setVisibilityForScreenshot(temp);

                        // Wait for layer to load
                        try { await temp.when(); } catch (e) { }

                        // ✅ Wait until layerView is not suspended AND not updating (best effort)
                        try {
                            const lv = await view.whenLayerView(temp);

                            // Wait for suspended -> false OR timeout
                            if (lv?.suspended) {
                                await new Promise(resolve => {
                                    const h = lv.watch("suspended", (s) => {
                                        if (!s) { h.remove(); resolve(); }
                                    });
                                    window.setTimeout(() => { try { h.remove(); } catch (e) { } resolve(); }, 4000);
                                });
                            }

                            // Wait for updating -> false OR timeout
                            if (lv?.updating) {
                                await new Promise(resolve => {
                                    const h = lv.watch("updating", (u) => {
                                        if (!u) { h.remove(); resolve(); }
                                    });
                                    window.setTimeout(() => { try { h.remove(); } catch (e) { } resolve(); }, 6000);
                                });
                            }
                        } catch (e) { }


                        // 🔒 Re-apply locked extent to guarantee identical framing
                        if (fixedExtent) {
                            await view.goTo(fixedExtent, { animate: false });
                        }

                        const ss = await view.takeScreenshot({ format: "png", quality: 100, width });
                        const dataUrl = ss?.dataUrl;
                        if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

                        // Coverage stats
                        const cov = await computeLayerCoverageStats(item, selectionGeom);
                        const acresCovered = cov ? cov.acresCovered : 0;
                        const pctCovered = cov ? cov.pctAoiCovered : 0;

                        const layerUrlLink = escapeHtml(item.url);

                        sectionsHtml += `
                    <div class="section">
                        <h2>${escapeHtml(item.title)}</h2>
                        <div class="sub">
                        Service: <a href="${layerUrlLink}" target="_blank" rel="noopener">${layerUrlLink}</a>
                        </div>

                        <div class="map">
                        <img src="${dataUrl}" alt="AOI + ${escapeHtml(item.title)}"/>
                        </div>

                        <table class="metaTbl">
                        <tr><td>AOI area</td><td><b>${formatNumber(aoiAcres, 2)}</b> acres</td></tr>
                        <tr><td>Intersecting features</td><td><b>${escapeHtml(String(item.count || 0))}</b></td></tr>
                        <tr><td>AOI covered by layer</td><td><b>${formatNumber(acresCovered, 2)}</b> acres</td></tr>
                        <tr><td>% AOI covered</td><td><b>${formatNumber(pctCovered, 2)}</b>%</td></tr>
                        </table>
                    </div>
                    <div class="pagebreak"></div>
                    `;
                    } finally {
                        try { view.map.remove(temp); } catch (e) { }
                        restoreVisibility();
                    }
                }
            }

            const htmlDoc = buildFinalReportHtmlDoc({
                title: "RMP Viewer — Final Report",
                createdAt: formatDateTimeForReport(new Date()),
                aoiSummary: getAoiSummaryForReport(aoiAcres),
                totalsHtml,
                sectionsHtml
            });
        
        cachedFinalReportHtml = htmlDoc;

        if (finalReportStatus) finalReportStatus.textContent = "Report ready.";

        } catch (e) {
            console.error(e);
            if (finalReportStatus) finalReportStatus.textContent = "Failed to build report (see console).";
        }        
      }
    


    // Simple wrapper - opens cached HTML
    function viewFinalReport() {
        if (!cachedFinalReportHtml) {
            alert("Run analysis first to generate the report.");
            return;
        }
        openHtmlInNewTab(cachedFinalReportHtml);
    }

    async function getFullFeatureGeometryFromLayer(layer, graphic) {
        if (!layer || !graphic) {
            return { geometry: graphic?.geometry || null, objectId: null, objectIdField: null, feature: graphic || null };
        }

        await layer.load();

        const oidField = layer.objectIdField || "OBJECTID";
        const oid = graphic?.attributes?.[oidField];

        // If we can’t determine OID, fall back to whatever we have
        if (oid == null) {
            return { geometry: graphic?.geometry || null, objectId: null, objectIdField: oidField };
        }

        try {
            const q = layer.createQuery();
            q.objectIds = [oid];
            q.returnGeometry = true;
            q.outFields = ["*"];                // ✅ fetch all attributes (not just OID)
            q.outSpatialReference = view?.spatialReference;
            q.maxAllowableOffset = 0;

            const fs = await layer.queryFeatures(q);
            const feat = fs?.features?.[0];

            return {
                geometry: feat?.geometry || graphic?.geometry || null,
                objectId: feat?.attributes?.[oidField] ?? oid,
                objectIdField: oidField,
                feature: feat || graphic || null
            };

        } catch (e) {
            console.warn("Full-geometry query failed; falling back to hit geometry", e);
            return { geometry: graphic?.geometry || null, objectId: oid, objectIdField: oidField, feature: graphic || null };
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
                    if (!graphic) return;

                    // ✅ Fetch the “true” polygon geometry (not the generalized hitTest geometry)
                    const full = await getFullFeatureGeometryFromLayer(activeSelectionLayer, graphic);
                    aoiSourceFeature = full?.feature || graphic || null; // ✅ cache clicked feature for AOI Source card
                    const fullGeom = full?.geometry || null;
                    if (!fullGeom) return;

                    setAoiGeometry(fullGeom);
                    setGeometryFromSelection(fullGeom);
                    resetCoverageCacheForAoi(fullGeom);

                    aoiSource = "select";
                    aoiSourceLayerTitle = activeSelectionLayer?.title || null;
                    aoiSourceLayerUrl = activeSelectionLayer?.url || null;

                    aoiSourceObjectIdField = full?.objectIdField || activeSelectionLayer?.objectIdField || "OBJECTID";
                    aoiSourceObjectId = (full?.objectId != null)
                        ? full.objectId
                        : (graphic?.attributes?.[aoiSourceObjectIdField] ?? null);

                    console.log("AOI source captured:", {
                        layerTitle: aoiSourceLayerTitle,
                        layerUrl: aoiSourceLayerUrl,
                        objectIdField: aoiSourceObjectIdField,
                        objectId: aoiSourceObjectId
                    });


                    // Keep PLSS tool context in-sync even if user didn’t click the toolbar button
                    if (aoiSourceLayerTitle) {
                        const t = normalize(aoiSourceLayerTitle);
                        if (t.includes("township")) aoiSourcePlssTool = "township";
                        else if (t.includes("section")) aoiSourcePlssTool = "section";
                        else if (t.includes("intersected") || t.includes("parcel")) aoiSourcePlssTool = "intersected";
                    }
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

        // --- Always-on basemap overlay: BLM SMA (BLM Only) ---
        const smaBlmOnly = new TileLayer({
            url: "https://gis.blm.gov/arcgis/rest/services/lands/BLM_Natl_SMA_Cached_BLM_Only/MapServer",
            title: "SMA (BLM only)",
            opacity: 0.8,      // tweak to taste
            visible: true
        });

        // Add as an operational layer at the bottom so it behaves like a basemap overlay
        map.add(smaBlmOnly, 0);


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
            creationMode: "single",
            updateOnGraphicClick: false
        });

        // Hard-disable editing UI for existing AOI graphics
        sketch.viewModel.updateOnGraphicClick = false;

        // Apply AOI symbol to Sketch (uses the AOI preset renderer symbol)
        const aoiRenderer = getPresetRenderer("aoi", null);
        if (aoiRenderer && aoiRenderer.symbol) {
            sketch.polygonSymbol = aoiRenderer.symbol;
        }

        sketch.on("create", (evt) => {
            if (evt.state === "complete") {
                const geom = evt.graphic?.geometry || null;
                setAoiGeometry(geom);          // ensure AOI is a single clean graphic
                resetCoverageCacheForAoi(geom);
                aoiSource = "draw";
                aoiSourceLayerTitle = null;
                aoiSourceLayerUrl = null;
                aoiSourceObjectId = null;
                aoiSourceObjectIdField = null;
                aoiSourceFeature = null; // ✅ drawn AOI has no source feature
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
                    let subTitle = String(sl.title || "");
                    subTitle = subTitle.replace(/intersected/ig, "Parcel");

                    // ✅ Item 4: remove "State Boundaries" from Selection (but keep for Report)
                    if (subTitle.toLowerCase() === "state boundaries") {
                        plssStateBoundary = {
                            title: `${cfg.title}: ${subTitle}`,
                            url: sl.url
                        };
                        plssStateLayerUrl = sl.url;
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

        // ✅ NEW: build report layers (for map display toggles)
        await buildReportDisplayLayers();

        renderLayerToggles(map);
        ensureAoiOnTop(map);


        await view.when();
        attachClickToSelect();

        // ---------- PLSS tool wiring (Township / Section / Intersected) ----------
        const townshipIdx = findSelectionLayerIndexByNameIncludes("township");
        const sectionIdx = findSelectionLayerIndexByNameIncludes("section");
        const intersectedIdx =
            (findSelectionLayerIndexByNameIncludes("parcel") >= 0)
                ? findSelectionLayerIndexByNameIncludes("parcel")
                : findSelectionLayerIndexByNameIncludes("intersected");

        plssParcelLayerUrl = (intersectedIdx >= 0) ? (selectionLayers[intersectedIdx]?.cfg?.url || null) : null;



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
                aoiSourcePlssTool = which; // <-- ADD: remember which PLSS tool is driving AOI selection
                setPlssToolActive(which);

                // Auto-zoom to minimum visible zoom level (using layer.minScale)
                const lyr = selectionLayers[idxToEnable]?.layer;
                await autoZoomToLayerMinVisible(lyr);
                await waitForViewStationary(1500);

                const whichLabel = (which === "intersected") ? "parcel" : which;
                setStatus(`PLSS select: ${whichLabel} (click a polygon)`);
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
        if (tabLayersBtn) tabLayersBtn.addEventListener("click", () => setActiveTab("layers"));

        if (tabServicesBtn) tabServicesBtn.addEventListener("click", () => {
            setActiveTab("services");
        });

        if (tabReportBtn) tabReportBtn.addEventListener("click", () => setActiveTab("report"));

        if (tabVisualBtn) tabVisualBtn.addEventListener("click", () => {
            setActiveTab("visual");
            renderVisualSummary();
        });

        if (tabFinalReportBtn) tabFinalReportBtn.addEventListener("click", () => {
            setActiveTab("finalReport");
        });

        if (refreshServicesBtn) refreshServicesBtn.addEventListener("click", refreshServicesTab);
        if (viewReportBtn) viewReportBtn.addEventListener("click", viewFinalReport);


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

        if (runBtn) runBtn.addEventListener("click", runAnalysis);
        
        if (cancelRunBtn) {
            cancelRunBtn.addEventListener("click", () => {
                // bump token to cancel; unlock immediately
                reportOpToken++;
                lockMapInteraction(false);
                cancelRunBtn.classList.add("hidden");
                setStatus("cancel requested…");
            });
        }

        if (clearBtn) clearBtn.addEventListener("click", clearAll);


        if (exportAllBtn) exportAllBtn.addEventListener("click", async () => {
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
        setActiveTab("layers");
        setStatus("ready");

        // Preload service status once (optional). Keeps Services tab fast.
        if (servicesListEl) {
            refreshServicesTab().catch(() => { });
        }
    }

    init().catch((e) => {
        console.error(e);
        setStatus("failed to initialize (see console)");
    });

});