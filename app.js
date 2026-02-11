/* global require */

require([
    "esri/Map",
    "esri/views/MapView",
    "esri/layers/FeatureLayer",
    "esri/layers/GraphicsLayer",
    "esri/widgets/Sketch",
    "esri/Graphic",
    "esri/geometry/geometryEngine",
    "esri/layers/TileLayer"
], function (EsriMap, MapView, FeatureLayer, GraphicsLayer, Sketch, Graphic, geometryEngine, TileLayer) {


    // ---------- DOM ----------
    const modeSelect = document.getElementById("modeSelect");
    // PLSS selection tools (Township / Section / Intersected)
    const plssTownshipBtn = document.getElementById("plssTownshipBtn");
    const plssSectionBtn = document.getElementById("plssSectionBtn");
    const plssIntersectedBtn = document.getElementById("plssIntersectedBtn");
    // Selection layer group selector
    const selectionGroupSelect = document.getElementById("selectionGroupSelect");
    const plssSelectGroup = document.getElementById("plssSelectGroup");
    const permitSelectGroup = document.getElementById("permitSelectGroup");
    // Permit layer selection buttons
    const grazingAllotmentBtn = document.getElementById("grazingAllotmentBtn");
    const grazingPastureBtn = document.getElementById("grazingPastureBtn");
    const oilGasLeaseBtn = document.getElementById("oilGasLeaseBtn");

    const selectModeControls = document.getElementById("selectModeControls");
    const drawModeControls = document.getElementById("drawModeControls");

    const drawBtn = document.getElementById("drawBtn");
    const stopDrawBtn = document.getElementById("stopDrawBtn");
    const runBtn = document.getElementById("runBtn");
    const clearBtn = document.getElementById("clearBtn");
    const exportAllBtn = document.getElementById("exportAllBtn");

    const viewBlockerEl = document.getElementById("viewBlocker");

    const statusEl = document.getElementById("status");
    const statusTextEl = document.getElementById("statusText");
    const busyIndicatorEl = document.getElementById("busyIndicator");

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

    // ---------- Analysis Modal Helpers ----------
    const analysisModal = {
        el: null,
        progressFill: null,
        currentStep: null,
        log: null,
        stats: {
            layersQueried: null,
            featuresFound: null,
            mapsGenerated: null
        },
        success: null,
        
        init() {
            this.el = document.getElementById("analysisModal");
            this.progressFill = document.getElementById("analysisProgressFill");
            this.currentStep = document.getElementById("analysisCurrentStep");
            this.log = document.getElementById("analysisLog");
            this.stats.layersQueried = document.getElementById("analysisLayersQueried");
            this.stats.featuresFound = document.getElementById("analysisFeaturesFound");
            this.stats.mapsGenerated = document.getElementById("analysisMapsGenerated");
            this.success = document.getElementById("analysisSuccess");
            
            // Wire cancel button
            const cancelBtn = document.getElementById("analysisModalCancelBtn");
            if (cancelBtn) {
                cancelBtn.addEventListener("click", () => {
                    reportOpToken++; // Cancel analysis
                    lockMapInteraction(false);
                    setBusy(false);
                    setStatus("canceled");
                    this.hide();
                });
            }
            
            // Wire success close button
            const successCloseBtn = document.getElementById("analysisSuccessCloseBtn");
            if (successCloseBtn) {
                successCloseBtn.addEventListener("click", () => {
                    this.hide();
                    setActiveTab("report"); // Jump to Tables tab
                });
            }
        },
        
        show() {
            if (!this.el) return;
            this.el.classList.remove("hidden");
            this.reset();
        },
        
        hide() {
            if (!this.el) return;
            this.el.classList.add("hidden");
        },
        
        reset() {
            this.setProgress(0);
            this.setStep("Initializing...");
            this.clearLog();
            this.updateStats(0, 0, 0);
            if (this.success) this.success.classList.add("hidden");
        },
        
        setProgress(percent) {
            if (this.progressFill) {
                this.progressFill.style.width = `${Math.min(100, Math.max(0, percent))}%`;
            }
        },
        
        setStep(text) {
            if (this.currentStep) {
                this.currentStep.textContent = text;
            }
        },
        
        addLog(message, type = "info") {
            if (!this.log) return;
            const entry = document.createElement("div");
            entry.className = `analysis-log-entry ${type}`;
            entry.textContent = `${new Date().toLocaleTimeString()} - ${message}`;
            this.log.appendChild(entry);
            this.log.scrollTop = this.log.scrollHeight; // Auto-scroll
        },
        
        clearLog() {
            if (this.log) this.log.innerHTML = "";
        },
        
        updateStats(layersQueried, featuresFound, mapsGenerated) {
            if (this.stats.layersQueried) this.stats.layersQueried.textContent = layersQueried;
            if (this.stats.featuresFound) this.stats.featuresFound.textContent = featuresFound;
            if (this.stats.mapsGenerated) this.stats.mapsGenerated.textContent = mapsGenerated;
        },
        
        showSuccess(layersQueried, featuresFound, mapsGenerated) {
            if (!this.success) return;
            
            // Update success summary
            const successLayersQueried = document.getElementById("successLayersQueried");
            const successFeaturesFound = document.getElementById("successFeaturesFound");
            const successMapsGenerated = document.getElementById("successMapsGenerated");
            
            if (successLayersQueried) successLayersQueried.textContent = layersQueried;
            if (successFeaturesFound) successFeaturesFound.textContent = featuresFound;
            if (successMapsGenerated) successMapsGenerated.textContent = mapsGenerated;
            
            this.success.classList.remove("hidden");
            
        }
    };


    // ---------- State ----------
    let config = null;

    let view = null;
    let selectionGeom = null;
    let aoiSource = null;            // "draw" | "select"
    let aoiSourceLayerTitle = null;  // optional: which selection layer was clicked
    let map = null; // <-- add (so PLSS buttons can add/remove selection layers)

    // Track service status (url -> "UP" | "DOWN")
    const serviceStatus = new Map();

    // AOI overlay (always on top)
    let aoiLayer = null;      // GraphicsLayer
    let aoiGraphic = null;    // Graphic (single AOI graphic)
    let aoiMaskLayer = null;  // GraphicsLayer for transparent mask outside AOI

    // Renderer lookup helpers
    let layerCfgByUrl = new Map(); // url -> {kind, cfg}

    // Always-visible layers (e.g. BLM Admin Units) — kept visible on main map + all report maps
    const alwaysVisibleLayers = [];

    let selectionLayers = []; // { cfg, layer }
    let activeSelectionLayer = null; // FeatureLayer
    let activeSelectionLayerView = null;
    let aoiSourcePlssTool = null; // "township" | "section" | "intersected" | null
    let aoiSourceLayerUrl = null;      // URL of the selection layer used to pick AOI (select mode)
    let plssParcelLayerUrl = null;     // URL of PLSS Intersected (UI will call "Parcel")
    let plssStateLayerUrl = "https://services2.arcgis.com/FiaPA4ga0iQKduv3/arcgis/rest/services/Public_Land_Survey_System_view/FeatureServer/0"; // Living Atlas State Boundary
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
        return my;
    }

    function endReportOp(myToken) {
        // Only unlock if this is the most recent op (prevents weird edge cases)
        if (myToken === reportOpToken) {
            lockMapInteraction(false);
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

    function setPermitToolActive(which) {
        const set = (btn, on) => {
            if (!btn) return;
            btn.setAttribute("aria-pressed", on ? "true" : "false");
        };
        set(grazingAllotmentBtn, which === "allotment");
        set(grazingPastureBtn, which === "pasture");
        set(oilGasLeaseBtn, which === "oilgas");
    }

    function clearAllSelectionToolButtons() {
        setPlssToolActive(null);
        setPermitToolActive(null);
    }

    function switchSelectionGroup(group) {
        if (plssSelectGroup) plssSelectGroup.classList.toggle("hidden", group !== "plss");
        if (permitSelectGroup) permitSelectGroup.classList.toggle("hidden", group !== "permit");
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

    function getPresetRenderer(kind, cfgObj, geometryType) {
        const sym = config?.symbology || {};
        const defaults = sym.defaults || {};
        const presets = sym.presets || {};

        // Allow per-layer override later (optional)
        let presetId =
            (cfgObj && cfgObj.symbologyPreset) ||
            (kind === "selection" ? defaults.selectionPreset :
                kind === "report" ? defaults.reportPreset :
                    defaults.aoiPreset);

        // For report layers, select point/line/polygon preset based on geometry type
        if (kind === "report" && geometryType) {
            const gt = String(geometryType).toLowerCase();
            if (gt.includes("point")) {
                presetId = "reportPoint";
            } else if (gt.includes("line") || gt.includes("polyline")) {
                presetId = "reportLine";
            }
            // polygons use default "report" preset
        }

        const r = presetId ? presets[presetId] : null;
        return r || null;
    }

    // Helper to get geometry type from a layer URL via pjson
    async function getLayerGeometryType(layerUrl) {
        try {
            const pjsonUrl = layerUrl.replace(/\/$/, "") + "?f=pjson";
            const pjson = await fetchJsonWithTimeout(pjsonUrl, 5000);
            return pjson?.geometryType || null;
        } catch (e) {
            return null;
        }
    }

    function ensureAoiOnTop(map) {
        if (!map || !aoiLayer) return;
        // Put AOI layer at top draw order
        map.reorder(aoiLayer, map.layers.length - 1);
        // Keep mask layer just below AOI
        if (aoiMaskLayer) {
            map.reorder(aoiMaskLayer, map.layers.length - 2);
        }
    }

    /**
     * Update the AOI mask layer with a "donut" polygon that lightens areas outside the AOI.
     * @param {boolean} show - Whether to show the mask
     */
    function updateAoiMask(show = true) {
        if (!aoiMaskLayer || !view) return;
        
        aoiMaskLayer.removeAll();
        
        if (!show || !selectionGeom) {
            aoiMaskLayer.visible = false;
            return;
        }

        // Get a large extent that covers beyond the current view
        const viewExt = view.extent;
        if (!viewExt) {
            aoiMaskLayer.visible = false;
            return;
        }

        // Expand extent significantly to ensure full coverage when panning/zooming
        const expandedExt = viewExt.expand(5);
        
        // Create outer ring (clockwise for exterior)
        const outerRing = [
            [expandedExt.xmin, expandedExt.ymin],
            [expandedExt.xmin, expandedExt.ymax],
            [expandedExt.xmax, expandedExt.ymax],
            [expandedExt.xmax, expandedExt.ymin],
            [expandedExt.xmin, expandedExt.ymin]
        ];

        // Get AOI ring(s) - need to reverse for hole (counter-clockwise)
        let aoiRings = [];
        if (selectionGeom.rings && selectionGeom.rings.length > 0) {
            // For polygon geometry, get all rings and reverse them for holes
            aoiRings = selectionGeom.rings.map(ring => [...ring].reverse());
        } else if (selectionGeom.type === "polygon") {
            // Fallback for different geometry structures
            const ext = selectionGeom.extent;
            aoiRings = [[
                [ext.xmin, ext.ymin],
                [ext.xmax, ext.ymin],
                [ext.xmax, ext.ymax],
                [ext.xmin, ext.ymax],
                [ext.xmin, ext.ymin]
            ]];
        }

        if (aoiRings.length === 0) {
            aoiMaskLayer.visible = false;
            return;
        }

        // Create donut polygon: outer ring + hole ring(s)
        const allRings = [outerRing, ...aoiRings];
        
        const maskGraphic = new Graphic({
            geometry: {
                type: "polygon",
                rings: allRings,
                spatialReference: view.spatialReference
            },
            symbol: {
                type: "simple-fill",
                color: [255, 255, 255, 0.45], // Semi-transparent white
                outline: null
            }
        });

        aoiMaskLayer.add(maskGraphic);
        aoiMaskLayer.visible = true;
    }

    /**
     * Hide the AOI mask layer
     */
    function hideAoiMask() {
        if (aoiMaskLayer) {
            aoiMaskLayer.visible = false;
            aoiMaskLayer.removeAll();
        }
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

    // ✅ NEW: Comprehensive layer view ready check - waits for layer to be fully rendered
    async function waitForLayerReadyToCapture(layer, view, { timeoutMs = 8000 } = {}) {
        if (!view || !layer) return;

        try { await layer.when(); } catch (e) { console.warn("Layer.when() failed:", e); }

        let lv = null;
        try {
            lv = await view.whenLayerView(layer);
        } catch (e) {
            console.warn("whenLayerView failed:", e);
            return;
        }

        if (!lv) return;

        // Wait for suspended state to resolve
        if (lv.suspended) {
            await new Promise(resolve => {
                const t = window.setTimeout(() => { h?.remove?.(); resolve(); }, 3000);
                const h = lv.watch("suspended", (s) => {
                    if (!s) {
                        clearTimeout(t);
                        h.remove();
                        resolve();
                    }
                });
            });
        }

        // Wait for initial updating to complete
        if (lv.updating) {
            await new Promise(resolve => {
                const t = window.setTimeout(() => { h?.remove?.(); resolve(); }, timeoutMs);
                const h = lv.watch("updating", (u) => {
                    if (!u) {
                        clearTimeout(t);
                        h.remove();
                        resolve();
                    }
                });
            });
        }

        // Force a final render by clearing and reapplying visibility
        try {
            await new Promise(r => setTimeout(r, 300));
        } catch (e) { }
    }

    // ✅ NEW: Improved screenshot capture with proper wait for tiles and basemap changes
    async function captureScreenshotWithWait(screenConfig = {}) {
        if (!view) return null;

        const width = screenConfig.width || (config?.visualReport?.screenshotWidth ?? 1400);
        
        // ✅ Wait for view to be completely stationary and rendered
        await waitForViewStationary(3000);

        // ✅ Wait for tile loading in cycles with longer delays
        for (let i = 0; i < 4; i++) {
            await new Promise(r => setTimeout(r, 300)); // Longer delay for tiles to load
        }

        // ✅ Wait one more time after renders
        await waitForViewStationary(2000);

        // ✅ Capture with improved settings
        try {
            const ss = await view.takeScreenshot({
                format: "png",
                quality: 100,
                width: width,
                height: Math.round(width * 0.5625) // 16:9 aspect ratio
            });
            return ss?.dataUrl || null;
        } catch (e) {
            console.error("Screenshot capture failed:", e);
            return null;
        }
    }


async function hardRefreshLayer(layer, { timeoutMs = 5000 } = {}) {
    if (!view || !layer) return;

    try { await layer.when(); } catch (e) { console.warn("Layer.when() error:", e); }

    let lv = null;
    try { lv = await view.whenLayerView(layer); } catch (e) {
        console.warn("whenLayerView error:", e);
        return;
    }
    if (!lv) return;

    // Wait for view to stop moving
    await waitForViewStationary(1500);

    // If suspended, wait for resume
    if (lv.suspended) {
        await new Promise(resolve => {
            const t = window.setTimeout(() => { h?.remove?.(); resolve(); }, 2000);
            const h = lv.watch("suspended", (s) => {
                if (!s) {
                    clearTimeout(t);
                    h.remove();
                    resolve();
                }
            });
        });
    }

    // Single refresh
    if (typeof lv.refresh === "function") {
        try {
            lv.refresh();
        } catch (e) {
            console.warn("Layer refresh failed:", e);
        }
    }

    // Wait for updating to finish
    await new Promise(resolve => {
        const t = window.setTimeout(() => { h?.remove?.(); resolve(); }, timeoutMs);
        const h = lv.watch("updating", (u) => {
            if (!u) {
                clearTimeout(t);
                h.remove();
                resolve();
            }
        });
        if (!lv.updating) {
            clearTimeout(t);
            h.remove();
            resolve();
        }
    });

    // Automatic render happens after state changes
        try {
            await new Promise(r => setTimeout(r, 200));
        } catch (e) { }
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

    // Expand a FeatureServer root into ALL sublayers (points, lines, polygons).
    async function expandFeatureServerToAllSublayers(serviceUrl) {
        const pjsonUrl = serviceUrl.replace(/\/$/, "") + "?f=pjson";
        const info = await fetchJson(pjsonUrl);
        const layers = Array.isArray(info?.layers) ? info.layers : [];

        const out = [];
        for (const l of layers) {
            const layerUrl = serviceUrl.replace(/\/$/, "") + "/" + l.id;
            let title = l?.name ? String(l.name) : `Layer ${l.id}`;
            out.push({ title, url: layerUrl });
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

        const useServiceRenderer = cfg.useServiceRenderer === true;
        const isAlwaysVisible = cfg.alwaysVisible === true;

        try {

        // FeatureServer root: expand sublayers
        if (isFeatureServerRoot(key)) {
            // If useServiceRenderer, expand to ALL sublayers (including points/lines);
            // otherwise expand to polygon-only as before.
            const subs = useServiceRenderer
                ? await expandFeatureServerToAllSublayers(key)
                : await expandFeatureServerToPolygonSublayers(key);

            const layers = [];
            for (const sl of subs) {
                const geomType = await getLayerGeometryType(sl.url);
                const layerOpts = {
                    url: sl.url,
                    title: `${cfg.title}: ${sl.title}`,
                    outFields: ["*"],
                    visible: isAlwaysVisible
                };
                // Only override renderer if NOT using service renderer
                if (!useServiceRenderer) {
                    layerOpts.renderer = getPresetRenderer("report", cfg, geomType) || undefined;
                }
                const lyr = new FeatureLayer(layerOpts);
                layers.push(lyr);
                if (isAlwaysVisible) alwaysVisibleLayers.push(lyr);
                // Register sublayer URL in config index (inherits parent config flags)
                layerCfgByUrl.set(sl.url, { kind: "report", cfg });
            }

            layers.forEach(l => map.add(l));
            reportLayerViews.set(key, layers);
            continue;
        }

        // MapServer root: expand to polygon sublayers
        if (isMapServerRoot(key)) {
            const subs = await expandMapServerToSublayers(key, { polygonOnly: !useServiceRenderer });

            const layers = [];
            for (const sl of subs) {
                const geomType = await getLayerGeometryType(sl.url);
                const layerOpts = {
                    url: sl.url,
                    title: `${cfg.title}: ${sl.title}`,
                    outFields: ["*"],
                    visible: isAlwaysVisible
                };
                if (!useServiceRenderer) {
                    layerOpts.renderer = getPresetRenderer("report", cfg, geomType) || undefined;
                }
                const lyr = new FeatureLayer(layerOpts);
                layers.push(lyr);
                if (isAlwaysVisible) alwaysVisibleLayers.push(lyr);
                // Register sublayer URL in config index (inherits parent config flags)
                layerCfgByUrl.set(sl.url, { kind: "report", cfg });
            }

            layers.forEach(l => map.add(l));
            reportLayerViews.set(key, layers);
            continue;
        }

        // Normal single layer - get geometry type
        const geomType = await getLayerGeometryType(key);
        const layerOpts = {
            url: key,
            title: cfg.title,
            outFields: ["*"],
            visible: isAlwaysVisible
        };
        if (!useServiceRenderer) {
            layerOpts.renderer = getPresetRenderer("report", cfg, geomType) || undefined;
        }
        const lyr = new FeatureLayer(layerOpts);
        if (isAlwaysVisible) alwaysVisibleLayers.push(lyr);

        map.add(lyr);
        reportLayerViews.set(key, lyr);

        } catch (e) {
            console.warn(`[buildReportDisplayLayers] Failed to load "${cfg.title}" (${key}):`, e);
            // Continue loading remaining layers
        }
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
            const status = serviceStatus.get(e.cfg.url);
            const statusIcon = (status === "DOWN") 
                ? `<span class="status-warning" title="Service is DOWN">⚠️</span>` 
                : "";
            
            return `
            <div class="toggle-row">
                <input type="checkbox" id="sellayer_${i}" ${checked} />
                <span class="layer-swatch layer-swatch-selection" aria-hidden="true" title="Selection layer"></span>
                <label class="toggle-name" for="sellayer_${i}">${statusIcon}${escapeHtml(e.cfg.title)}</label>
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
            const isAlwaysOn = l.alwaysVisible === true;
            const isChecked = isAlwaysOn;
            const status = serviceStatus.get(l.url);
            const statusIcon = (status === "DOWN") 
                ? `<span class="status-warning" title="Service is DOWN">⚠️</span>` 
                : "";

            const checked = isChecked ? "checked" : "";

            // ✅ Do NOT disable FeatureServer roots anymore (we will expand them to drawable polygon sublayers)
            // alwaysVisible layers default ON but user can toggle them off on the main map
            const disabled = "";

            return `
                <div class="toggle-row">
                    <input type="checkbox" id="rptlayer_${i}" ${checked} ${disabled} />
                    <span class="layer-swatch layer-swatch-report" aria-hidden="true" title="Report layer"></span>
                    <label class="toggle-name" for="rptlayer_${i}">${statusIcon}${escapeHtml(l.title)}</label>
                    <span id="rptlayer_spin_${i}" class="layer-spinner hidden" aria-label="loading"></span>
                </div>
            `;
        }).join("");

        (config.reportLayers || []).forEach((l, i) => {
            const cb = document.getElementById(`rptlayer_${i}`);
            if (!cb) return;

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


async function checkServiceStatusBackground() {
    const items = getConfiguredServices();
    const timeoutMs = config?.services?.timeoutMs ?? 8000;

    for (const it of items) {
        const pjsonUrl = normalizePjsonUrl(it.url);
        
        try {
            await fetchJsonWithTimeout(pjsonUrl, timeoutMs);
            serviceStatus.set(it.url, "UP");
        } catch (e) {
            serviceStatus.set(it.url, "DOWN");
        }
    }

    // Re-render layer toggles to show status icons
    renderLayerToggles(map);
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

            serviceStatus.set(it.url, status);

            if (desc) {
                serviceStatus.set(it.url + "::desc", desc);
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

    // ✅ Show analysis modal
    analysisModal.show();
    analysisModal.setProgress(0);
    analysisModal.setStep("Starting analysis...");
    analysisModal.addLog("Analysis started");

    setBusy(true);

    let layersQueried = 0;
    let featuresFound = 0;
    let mapsGenerated = 0;

    try {
        // Step 1: Data Check (10% progress)
        analysisModal.setStep("Step 1/4: Checking services...");
        analysisModal.setProgress(10);
        analysisModal.addLog("Checking service availability");
        
        await refreshServicesTab();

        if (isReportCanceled(myOp)) {
            analysisModal.addLog("Analysis canceled by user", "error");
            analysisModal.hide();
            setStatus("canceled");
            return;
        }

        analysisModal.addLog("Service check complete", "success");
        analysisModal.setProgress(25);

        // Step 2: Query all layers (25% → 60% progress)
        analysisModal.setStep("Step 2/4: Querying layers...");
        analysisModal.addLog("Starting layer queries");
        
        // Pass modal reference for live updates
        await queryAllLayers(reportGeom, myOp, analysisModal);

        if (isReportCanceled(myOp)) {
            analysisModal.addLog("Analysis canceled by user", "error");
            analysisModal.hide();
            setStatus("canceled");
            return;
        }

        // Update stats from query results
        layersQueried = lastReportRowsByLayer.length;
        featuresFound = lastReportRowsByLayer.reduce((sum, x) => sum + (x.count || 0), 0);
        analysisModal.updateStats(layersQueried, featuresFound, 0);
        analysisModal.addLog(`Found ${featuresFound} features across ${layersQueried} layers`, "success");
        analysisModal.setProgress(60);

        // Step 3: Generate map screenshots (60% → 85% progress)
        analysisModal.setStep("Step 3/4: Generating maps...");
        analysisModal.addLog("Starting map generation");
        
        await generateVisualReportData(myOp, analysisModal);

        if (isReportCanceled(myOp)) {
            analysisModal.addLog("Analysis canceled by user", "error");
            analysisModal.hide();
            setStatus("canceled");
            return;
        }

        // Count maps generated
        mapsGenerated = lastReportRowsByLayer.filter(x => (x?.count || 0) > 0 && x?._layer && x?._exportQuery).length;
        analysisModal.updateStats(layersQueried, featuresFound, mapsGenerated);
        analysisModal.addLog(`Generated ${mapsGenerated} maps`, "success");
        analysisModal.setProgress(85);

        // Step 4: Build final report HTML (85% → 100% progress)
        analysisModal.setStep("Step 4/4: Building final report...");
        analysisModal.addLog("Compiling final report");
        
        await buildFinalReportHtml();

        analysisModal.addLog("Final report ready", "success");
        analysisModal.setProgress(100);

        // Enable "View Report" button
        if (viewReportBtn) viewReportBtn.disabled = false;

        setStatus("Analysis complete!");
        
        // ✅ Show success animation
        analysisModal.showSuccess(layersQueried, featuresFound, mapsGenerated);

    } catch (e) {
        console.error(e);
        analysisModal.addLog(`Error: ${e.message}`, "error");
        analysisModal.setStep("Analysis failed");
        setStatus("Analysis failed (see console)");
        
        setTimeout(() => analysisModal.hide(), 3000);
    } finally {
        setBusy(false);
        endReportOp(myOp);
    }
}


// Extracted query logic (was: runReport)
async function queryAllLayers(reportGeom, myOp, modal = null) {
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

        if (modal) {
            const progress = 25 + (35 * (i / expandedTargets.length)); // 25% → 60%
            modal.setProgress(progress);
            modal.setStep(`Step 2/4: Querying ${t.title}...`);
            modal.addLog(`Querying: ${t.title}`);
        }

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
                    Export CSV
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
                fullRows: null  // ✅ Will be populated on-demand for State/Parcel
            });

            // ✅ Pre-fetch full rows for State Boundaries & Parcel (needed for Final Report)
            const isStateBoundaries = r.title && r.title.toLowerCase().includes("state boundaries");
            const isParcel = r.title && (r.title.toLowerCase().includes("parcel") || r.title.toLowerCase().includes("intersected"));

            if ((isStateBoundaries || isParcel) && r.count > 0 && r.layer && r.exportQuery) {
                try {
                    const pageSize = config.report?.pageSize ?? 1000;
                    const maxExport = config.report?.maxExportFeatures ?? 50000;

                    const fullFeatures = await queryAllFeaturesPaged(
                        r.layer,
                        r.exportQuery,
                        pageSize,
                        Math.min(maxExport, 100)  // Cap at 100 for State/Parcel (they should be small)
                    );

                    // Store in the item we just pushed
                    lastReportRowsByLayer[lastReportRowsByLayer.length - 1].fullRows = flattenAttributes(fullFeatures);
                } catch (e) {
                    console.warn(`Failed to pre-fetch full rows for ${r.title}:`, e);
                }
            }



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
                Export CSV
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

        // ✅ Update modal stats after each successful query
        if (modal) {
            const totalFeatures = lastReportRowsByLayer.reduce((sum, x) => sum + (x.count || 0), 0);
            modal.updateStats(lastReportRowsByLayer.length, totalFeatures, 0);
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

    /**
     * Build a per-feature breakdown table for layers with multiple intersecting features.
     * Each row = one feature, with relevant attribute columns + Acres + % AOI.
     */
    async function buildPerFeatureTable(item, aoiGeom) {
        if (!item || !item._layer || !item._exportQuery || !aoiGeom) return "";
        if ((item.count || 0) < 2) return ""; // only for multiple features

        // AOI area
        let aoiAreaSqm = 0;
        try {
            aoiAreaSqm = Math.max(0, geometryEngine.geodesicArea(aoiGeom, "square-meters"));
        } catch (e) { aoiAreaSqm = 0; }
        if (!aoiAreaSqm) return "";

        const pageSize = config.report?.pageSize ?? 1000;
        const maxExport = config.report?.maxExportFeatures ?? 50000;

        // Query all features WITH geometry + attributes
        let feats = [];
        try {
            const q = item._exportQuery.clone();
            q.returnGeometry = true;
            q.outFields = ["*"];
            q.outSpatialReference = view?.spatialReference;

            const all = [];
            let offset = 0;
            while (true) {
                const pq = q.clone();
                pq.num = pageSize;
                pq.start = offset;
                const fs = await item._layer.queryFeatures(pq);
                const batch = fs?.features ?? [];
                all.push(...batch);
                if (batch.length < pageSize) break;
                offset += pageSize;
                if (all.length >= maxExport) break;
            }
            feats = all;
        } catch (e) {
            console.warn("buildPerFeatureTable: query failed", e);
            return "";
        }

        if (feats.length < 2) return "";

        // Determine relevant attribute columns (skip OID-like, shape, and GlobalID fields)
        const skipPatterns = /^(objectid|oid|fid|shape|shape_area|shape_length|shape\.area|shape\.len|globalid|st_area|st_length|st_perimeter)$/i;
        const allKeys = new Set();
        for (const f of feats) {
            if (!f.attributes) continue;
            for (const k of Object.keys(f.attributes)) {
                if (!skipPatterns.test(k)) allKeys.add(k);
            }
        }

        // Pick up to 6 attribute columns to keep the table readable
        const maxAttrCols = 6;
        const attrKeys = Array.from(allKeys).slice(0, maxAttrCols);

        // Build rows: compute per-feature intersection area
        const tableRows = [];
        for (const f of feats) {
            const attrs = f.attributes || {};
            const geom = f.geometry;

            let acresCovered = 0;
            let pctAoi = 0;

            if (geom) {
                try {
                    const inter = geometryEngine.intersect(aoiGeom, geom);
                    if (inter) {
                        const areaSqm = Math.max(0, geometryEngine.geodesicArea(inter, "square-meters"));
                        acresCovered = areaSqm / SQM_PER_ACRE;
                        pctAoi = Math.min(100, Math.max(0, (areaSqm / aoiAreaSqm) * 100));
                    }
                } catch (e) { /* skip bad geometry */ }
            }

            tableRows.push({ attrs, acresCovered, pctAoi });
        }

        // Sort by acres descending
        tableRows.sort((a, b) => b.acresCovered - a.acresCovered);

        // Build HTML
        const thCells = attrKeys.map(k => `<th>${escapeHtml(k)}</th>`).join("");
        const headerHtml = `<tr>${thCells}<th>Acres</th><th>% of AOI</th></tr>`;

        const bodyHtml = tableRows.map(row => {
            const tdCells = attrKeys.map(k => {
                const raw = (row.attrs[k] == null) ? "" : String(row.attrs[k]);
                const truncated = raw.length > 40 ? raw.slice(0, 39) + "…" : raw;
                return `<td title="${escapeHtml(raw)}">${escapeHtml(truncated)}</td>`;
            }).join("");
            // Flag rows with <3% coverage as potential sliver/anomaly
            const isSliverWarning = row.pctAoi < 3;
            const rowStyle = isSliverWarning ? ' style="background-color:#fff3cd;"' : '';
            const warningIcon = isSliverWarning ? '<span title="Low coverage (&lt;3%) — possible sliver or boundary artifact" style="color:#856404; cursor:help;">⚠️</span> ' : '';
            return `<tr${rowStyle}>${tdCells}<td style="text-align:right;">${formatNumber(row.acresCovered, 2)}</td><td style="text-align:right;">${warningIcon}${formatNumber(row.pctAoi, 2)}%</td></tr>`;
        }).join("");

        return `
            <div style="margin-top:16px;">
                <b>Per-Feature Breakdown</b>
                <table class="metaTbl" style="margin-top:8px; width:100%; font-size:11px;">
                    <thead>${headerHtml}</thead>
                    <tbody>${bodyHtml}</tbody>
                </table>
            </div>
        `;
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

async function generateVisualReportData(myOp, modal = null) {

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
                .filter(x => x?._layer && x?._exportQuery) // excludes pinned AOI source etc.
                .filter(x => !(x.title && x.title.toLowerCase().includes("state boundaries")))
                .filter(x => !(x.title && x.title.toLowerCase().includes("administrative unit")));

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
                    // Keep AOI mask layer visible
                    if (aoiMaskLayer && l === aoiMaskLayer) { l.visible = true; continue; }

                    // Keep SMA overlay visible (your TileLayer at bottom) if present
                    // (We don’t have the variable here; keep TileLayers visible by default.)
                    if (l?.type === "tile") { l.visible = true; continue; }

                    // Keep always-visible layers (e.g. BLM Admin Units) visible
                    if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }

                    // Hide everything else (selection layers, other report layers, etc.)
                    l.visible = false;
                }

                if (tempLayer) tempLayer.visible = true;
                // ✅ Show AOI mask to lighten areas outside AOI
                updateAoiMask(true);
                ensureAoiOnTop(view.map);
            }

            function restoreVisibility() {
                visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
                hideAoiMask();
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

                // ✅ Update modal progress
                if (modal) {
                    const progress = 60 + (25 * (i / targets.length)); // 60% → 85%
                    modal.setProgress(progress);
                    modal.setStep(`Step 3/4: Generating map ${i + 1}/${targets.length}...`);
                    modal.addLog(`Generating map for: ${item.title}`);
                }

                setVisualStatus(`Generating map ${i + 1} / ${targets.length}…`);

                // Create a temporary layer for this URL, regardless of toggle state
                // Get geometry type for appropriate renderer — fully opaque for visual report
                const tempGeomType = await getLayerGeometryType(item.url);
                const vrItemCfg = layerCfgByUrl.get(item.url)?.cfg || null;
                const vrUseNative = vrItemCfg?.useServiceRenderer === true;
                const opaqueVRRenderer = vrUseNative
                    ? undefined
                    : makeRendererOpaque(getPresetRenderer("report", vrItemCfg, tempGeomType));
                const tempOpts = {
                    url: item.url,
                    title: item.title,
                    outFields: ["*"],
                    visible: true
                };
                if (!vrUseNative) {
                    tempOpts.renderer = opaqueVRRenderer || undefined;
                }
                const temp = new FeatureLayer(tempOpts);

                // 🔒 Only force scale override for non-native-renderer layers
                if (!vrUseNative) {
                    temp.minScale = 0;
                    temp.maxScale = 0;
                }

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


    // ---------- Helper: Format PLSS Legal Description ----------
    function formatLegalDescription(row) {
        if (!row) return "";
        
        // Required fields from PLSS: Parcel table
        const twnshpNo = row.TWNSHPNO || "";
        const twnshpDir = row.TWNSHPDIR || "";
        const rangeNo = row.RANGENO || "";
        const rangeDir = row.RANGEDIR || "";
        const frstDivTxt = row.FRSTDIVTXT || "";
        const frstDivNo = row.FRSTDIVNO || "";
        
        // Optional fields - only include if not NULL
        const secDivTyp = row.SECDIVTYP || "";
        const secDivNo = row.SECDIVNO || "";
        
        // Build the Legal Land Description
        let desc = "";
        
        // Township and Range are required
        if (twnshpNo && twnshpDir && rangeNo && rangeDir) {
            desc = `Township ${twnshpNo} ${twnshpDir} Range: ${rangeNo} ${rangeDir}`;
            
            // Add first division info
            if (frstDivTxt || frstDivNo) {
                desc += ` ${frstDivTxt} ${frstDivNo}`.trim();
            }
            
            // Add second division info only if present
            if (secDivTyp && secDivNo) {
                desc += ` ${secDivTyp} ${secDivNo}`;
            } else if (secDivTyp) {
                desc += ` ${secDivTyp}`;
            } else if (secDivNo) {
                desc += ` ${secDivNo}`;
            }
        }
        
        return desc.trim();
    }

    // ---------- Helper: Generate layer-specific attribute summary for Final Report ----------
    function generateLayerAttributeSummary(item) {
        if (!item) return "";
        
        const title = (item.title || "").toLowerCase();
        const rows = item.fullRows || item.rows || [];
        
        if (!rows.length) return "";
        
        let summaryHtml = "";
        
        // BLM National Visual Resource Inventory Classes
        if (title.includes("visual resource inventory") || title.includes("vri")) {
            const vriClassCounts = new Map();
            const scenicRatingCounts = new Map();
            
            for (const row of rows) {
                // VRI Class Code
                const vriClass = row.VRI_CLASS_CODE || row.VRI_CLASS || row.CLASS_CODE || "";
                if (vriClass) {
                    vriClassCounts.set(vriClass, (vriClassCounts.get(vriClass) || 0) + 1);
                }
                
                // Scenic Quality Rating
                const scenicRating = row.SL_OVRL_RT || row.SCENIC_QUALITY || row.SQ_RATING || "";
                if (scenicRating) {
                    scenicRatingCounts.set(scenicRating, (scenicRatingCounts.get(scenicRating) || 0) + 1);
                }
            }
            
            if (vriClassCounts.size > 0 || scenicRatingCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>VRI Attributes</b></td></tr>`;
                
                if (vriClassCounts.size > 0) {
                    const classItems = Array.from(vriClassCounts.entries())
                        .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
                        .map(([cls, count]) => `${escapeHtml(cls)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>VRI Class Codes</td><td>${classItems}</td></tr>`;
                }
                
                if (scenicRatingCounts.size > 0) {
                    const ratingItems = Array.from(scenicRatingCounts.entries())
                        .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
                        .map(([rating, count]) => `${escapeHtml(rating)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Scenic Quality Ratings</td><td>${ratingItems}</td></tr>`;
                }
            }
        }
        
        // USFWS Critical Habitat
        if (title.includes("critical habitat")) {
            const speciesCounts = new Map();
            const statusCounts = new Map();
            
            for (const row of rows) {
                const species = row.COMNAME || row.SCINAME || row.SPECIES || "";
                if (species) {
                    speciesCounts.set(species, (speciesCounts.get(species) || 0) + 1);
                }
                
                const status = row.STATUS || row.LISTING_STATUS || "";
                if (status) {
                    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                }
            }
            
            if (speciesCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Critical Habitat Details</b></td></tr>`;
                const speciesItems = Array.from(speciesCounts.entries())
                    .sort((a, b) => b[1] - a[1]) // Sort by count descending
                    .slice(0, 10) // Limit to top 10
                    .map(([sp, count]) => `${escapeHtml(sp)} (${count})`)
                    .join(", ");
                summaryHtml += `<tr><td>Species</td><td>${speciesItems}${speciesCounts.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }
        
        // BLM Grazing Allotments
        if (title.includes("grazing allotment")) {
            const allotmentNames = new Set();
            
            for (const row of rows) {
                const name = row.ALLOT_NAME || row.ALLOTMENT_NAME || row.NAME || "";
                if (name) allotmentNames.add(name);
            }
            
            if (allotmentNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Grazing Allotment Details</b></td></tr>`;
                const names = Array.from(allotmentNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>Allotment Names</td><td>${names}${allotmentNames.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }
        
        // BLM Wilderness Areas / WSA
        if (title.includes("wilderness")) {
            const areaNames = new Set();
            
            for (const row of rows) {
                const name = row.NLCS_NAME || row.WSA_NAME || row.NAME || row.UNIT_NAME || "";
                if (name) areaNames.add(name);
            }
            
            if (areaNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Wilderness Area Details</b></td></tr>`;
                const names = Array.from(areaNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>Area Names</td><td>${names}${areaNames.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }
        
        // BLM ACEC
        if (title.includes("acec") || title.includes("critical environmental concern")) {
            const acecNames = new Set();
            
            for (const row of rows) {
                const name = row.ACEC_NAME || row.NAME || row.UNIT_NAME || "";
                if (name) acecNames.add(name);
            }
            
            if (acecNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>ACEC Details</b></td></tr>`;
                const names = Array.from(acecNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>ACEC Names</td><td>${names}${acecNames.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }
        
        // Wild Horse and Burro Areas
        if (title.includes("wild horse") || title.includes("burro")) {
            const herdNames = new Set();
            
            for (const row of rows) {
                const name = row.HA_NAME || row.HMA_NAME || row.NAME || "";
                if (name) herdNames.add(name);
            }
            
            if (herdNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Herd Area Details</b></td></tr>`;
                const names = Array.from(herdNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>Herd Area Names</td><td>${names}${herdNames.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }
        
        // Ungulate Migration
        if (title.includes("ungulate") || title.includes("migration")) {
            const speciesCounts = new Map();
            const useCounts = new Map();
            
            for (const row of rows) {
                const species = row.SPECIES || row.COMMON_NAME || "";
                if (species) {
                    speciesCounts.set(species, (speciesCounts.get(species) || 0) + 1);
                }
                
                const use = row.USE_TYPE || row.SEASON || row.MOVEMENT_TYPE || "";
                if (use) {
                    useCounts.set(use, (useCounts.get(use) || 0) + 1);
                }
            }
            
            if (speciesCounts.size > 0 || useCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Migration Corridor Details</b></td></tr>`;
                
                if (speciesCounts.size > 0) {
                    const speciesItems = Array.from(speciesCounts.entries())
                        .map(([sp, count]) => `${escapeHtml(sp)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Species</td><td>${speciesItems}</td></tr>`;
                }
                
                if (useCounts.size > 0) {
                    const useItems = Array.from(useCounts.entries())
                        .map(([u, count]) => `${escapeHtml(u)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Use Type / Season</td><td>${useItems}</td></tr>`;
                }
            }
        }
        
        // BLM MLRS LUA ROW (Rights of Way)
        if (title.includes("mlrs") && (title.includes("row") || title.includes("lua"))) {
            const authTypes = new Map();
            const statusCounts = new Map();
            const caseNumbers = new Set();
            
            for (const row of rows) {
                const authType = row.AUTH_TYPE || row.AUTHORIZATION_TYPE || row.TYPE || "";
                if (authType) {
                    authTypes.set(authType, (authTypes.get(authType) || 0) + 1);
                }
                
                const status = row.CASE_STATUS || row.STATUS || row.AUTH_STATUS || "";
                if (status) {
                    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                }
                
                const caseNo = row.CASE_NR || row.SERIAL_NR || row.CASE_NUMBER || "";
                if (caseNo) caseNumbers.add(caseNo);
            }
            
            if (authTypes.size > 0 || statusCounts.size > 0 || caseNumbers.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>ROW Authorization Details</b></td></tr>`;
                
                if (authTypes.size > 0) {
                    const items = Array.from(authTypes.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Authorization Types</td><td>${items}</td></tr>`;
                }
                
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
                
                if (caseNumbers.size > 0 && caseNumbers.size <= 10) {
                    const items = Array.from(caseNumbers).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Case Numbers</td><td>${items}</td></tr>`;
                } else if (caseNumbers.size > 10) {
                    summaryHtml += `<tr><td>Case Numbers</td><td>${caseNumbers.size} cases</td></tr>`;
                }
            }
        }
        
        // BLM National Land Use Plans (including Revision/Development)
        if (title.includes("land use plan") || title.includes("revision") && title.includes("development")) {
            const planNames = new Set();
            const statusCounts = new Map();
            const epLinks = new Set();
            const nepaNumbers = new Set();
            const rodYears = new Set();
            
            for (const row of rows) {
                // LUPName - plan name field
                const name = row.LUPName || row.LUPNAME || row.PLAN_NAME || row.PLAN_NM || 
                             row.RMP_NAME || row.NAME || row.LUP_NAME || "";
                if (name) planNames.add(name);
                
                // Status field
                const status = row.Status || row.STATUS || row.PLAN_STATUS || 
                               row.APPROVAL_STATUS || row.LUP_STATUS || "";
                if (status) {
                    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                }
                
                // ePLink - ePlanning link
                const epLink = row.ePLink || row.EPLINK || row.EP_LINK || row.EPLANNING_LINK || "";
                if (epLink) epLinks.add(epLink);
                
                // NEPAnum - NEPA number
                const nepaNum = row.NEPAnum || row.NEPANUM || row.NEPA_NUM || row.NEPA_NUMBER || "";
                if (nepaNum) nepaNumbers.add(nepaNum);
                
                // RODyear - Record of Decision year
                const rodYear = row.RODyear || row.RODYEAR || row.ROD_YEAR || row.ROD_YR || "";
                if (rodYear) rodYears.add(String(rodYear));
            }
            
            if (planNames.size > 0 || statusCounts.size > 0 || epLinks.size > 0 || nepaNumbers.size > 0 || rodYears.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Land Use Plan Details</b></td></tr>`;
                
                if (planNames.size > 0) {
                    const names = Array.from(planNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>LUP Name</td><td>${names}${planNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
                
                if (epLinks.size > 0) {
                    const links = Array.from(epLinks).slice(0, 5).map(link => {
                        const escaped = escapeHtml(link);
                        return `<a href="${escaped}" target="_blank">${escaped.length > 50 ? escaped.substring(0, 50) + "..." : escaped}</a>`;
                    }).join("<br>");
                    summaryHtml += `<tr><td>ePlanning Link</td><td>${links}${epLinks.size > 5 ? "<br>..." : ""}</td></tr>`;
                }
                
                if (nepaNumbers.size > 0) {
                    const nums = Array.from(nepaNumbers).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>NEPA Number</td><td>${nums}${nepaNumbers.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (rodYears.size > 0) {
                    const years = Array.from(rodYears).sort().map(y => escapeHtml(y)).join(", ");
                    summaryHtml += `<tr><td>ROD Year</td><td>${years}</td></tr>`;
                }
            }
        }
        
        // BLM National Conservation Areas (NLCS) - NM/NCA
        if (title.includes("nlcs") || title.includes("conservation area") || title.includes("national monument")) {
            const areaNames = new Set();
            const designations = new Map();
            
            for (const row of rows) {
                const name = row.NLCS_NAME || row.NCA_NAME || row.NM_NAME || row.NAME || row.UNIT_NAME || "";
                if (name) areaNames.add(name);
                
                const desig = row.DESIGNATION || row.NLCS_TYPE || row.TYPE || "";
                if (desig) {
                    designations.set(desig, (designations.get(desig) || 0) + 1);
                }
            }
            
            if (areaNames.size > 0 || designations.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Conservation Area Details</b></td></tr>`;
                
                if (areaNames.size > 0) {
                    const names = Array.from(areaNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Area Names</td><td>${names}${areaNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (designations.size > 0) {
                    const items = Array.from(designations.entries())
                        .map(([d, count]) => `${escapeHtml(d)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Designation Type</td><td>${items}</td></tr>`;
                }
            }
        }
        
        // BLM Locatable Mineral Allocations
        if (title.includes("locatable") || title.includes("mineral allocation")) {
            const allocations = new Map();
            const statusCounts = new Map();
            
            for (const row of rows) {
                const alloc = row.LOC_ALLOC || row.ALLOCATION || row.ALLOC_TYPE || row.MINERAL_ALLOCATION || "";
                if (alloc) {
                    allocations.set(alloc, (allocations.get(alloc) || 0) + 1);
                }
                
                const status = row.STATUS || row.ALLOC_STATUS || "";
                if (status) {
                    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                }
            }
            
            if (allocations.size > 0 || statusCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Mineral Allocation Details</b></td></tr>`;
                
                if (allocations.size > 0) {
                    const items = Array.from(allocations.entries())
                        .map(([a, count]) => `${escapeHtml(a)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Allocation Types</td><td>${items}</td></tr>`;
                }
                
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
            }
        }
        
        // BLM Timber Allocations
        if (title.includes("timber")) {
            const allocations = new Map();
            const statusCounts = new Map();
            
            for (const row of rows) {
                const alloc = row.TIMBER_ALLOC || row.ALLOCATION || row.ALLOC_TYPE || row.HARVEST_TYPE || "";
                if (alloc) {
                    allocations.set(alloc, (allocations.get(alloc) || 0) + 1);
                }
                
                const status = row.STATUS || row.ALLOC_STATUS || "";
                if (status) {
                    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                }
            }
            
            if (allocations.size > 0 || statusCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Timber Allocation Details</b></td></tr>`;
                
                if (allocations.size > 0) {
                    const items = Array.from(allocations.entries())
                        .map(([a, count]) => `${escapeHtml(a)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Allocation Types</td><td>${items}</td></tr>`;
                }
                
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
            }
        }
        
        // USFWS Regions
        if (title.includes("usfws") && title.includes("region")) {
            const regionNames = new Set();
            
            for (const row of rows) {
                const region = row.REGNAME || row.REGION_NAME || row.REGION || row.NAME || "";
                if (region) regionNames.add(region);
            }
            
            if (regionNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>USFWS Region Details</b></td></tr>`;
                const names = Array.from(regionNames).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>Regions</td><td>${names}</td></tr>`;
            }
        }
        
        // BLM Grazing Pasture
        if (title.includes("grazing pasture") || title.includes("pasture polygon")) {
            const pastureNames = new Set();
            const allotmentNames = new Set();
            
            for (const row of rows) {
                const pasture = row.PASTURE_NAME || row.PAST_NAME || row.NAME || "";
                if (pasture) pastureNames.add(pasture);
                
                const allotment = row.ALLOT_NAME || row.ALLOTMENT_NAME || "";
                if (allotment) allotmentNames.add(allotment);
            }
            
            if (pastureNames.size > 0 || allotmentNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Grazing Pasture Details</b></td></tr>`;
                
                if (pastureNames.size > 0) {
                    const names = Array.from(pastureNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Pasture Names</td><td>${names}${pastureNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (allotmentNames.size > 0) {
                    const names = Array.from(allotmentNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Allotment Names</td><td>${names}${allotmentNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
            }
        }
        
        // BLM Oil and Gas Leases (MLRS)
        if (title.includes("oil") && title.includes("gas")) {
            const statusCounts = new Map();
            const typeCounts = new Map();
            const caseNumbers = new Set();
            const lesseeNames = new Set();
            const commodityCounts = new Map();
            
            for (const row of rows) {
                // Status variations
                const status = row.CASE_STATUS || row.LEASE_STATUS || row.STATUS || 
                               row.CASE_STAT || row.STAT || row.DISP_STATUS || 
                               row.CASE_TYP || row.AUTH_STAT || "";
                if (status) {
                    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                }
                
                // Type variations
                const type = row.LEASE_TYPE || row.AUTH_TYPE || row.TYPE || 
                             row.CASE_TYPE || row.TYP || row.DISP_TYPE ||
                             row.AUTH_TYP || row.CASETYPE || "";
                if (type) {
                    typeCounts.set(type, (typeCounts.get(type) || 0) + 1);
                }
                
                // Case/Serial number variations
                const caseNo = row.CASE_NR || row.SERIAL_NR || row.LEASE_NUMBER || 
                               row.CASE_NO || row.SERIAL_NO || row.CASENR || 
                               row.SERIALNR || row.CASE_ID || row.AUTH_NR || "";
                if (caseNo) caseNumbers.add(caseNo);
                
                // Lessee/Holder name
                const lessee = row.HOLDER_NAME || row.LESSEE || row.LESSEE_NAME ||
                               row.HOLDER || row.COMPANY || row.OPERATOR ||
                               row.CUSTOMER_NAME || row.CUST_NAME || "";
                if (lessee) lesseeNames.add(lessee);
                
                // Commodity
                const commodity = row.COMMODITY || row.CMDTY || row.RESOURCE ||
                                  row.MINERAL || row.PRODUCT || "";
                if (commodity) {
                    commodityCounts.set(commodity, (commodityCounts.get(commodity) || 0) + 1);
                }
            }
            
            if (statusCounts.size > 0 || typeCounts.size > 0 || caseNumbers.size > 0 || 
                lesseeNames.size > 0 || commodityCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Oil & Gas Lease Details</b></td></tr>`;
                
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Lease Status</td><td>${items}</td></tr>`;
                }
                
                if (typeCounts.size > 0) {
                    const items = Array.from(typeCounts.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Lease Type</td><td>${items}</td></tr>`;
                }
                
                if (commodityCounts.size > 0) {
                    const items = Array.from(commodityCounts.entries())
                        .map(([c, count]) => `${escapeHtml(c)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Commodity</td><td>${items}</td></tr>`;
                }
                
                if (lesseeNames.size > 0) {
                    const names = Array.from(lesseeNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Lessee/Holder</td><td>${names}${lesseeNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (caseNumbers.size > 0 && caseNumbers.size <= 10) {
                    const items = Array.from(caseNumbers).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Case/Serial Numbers</td><td>${items}</td></tr>`;
                } else if (caseNumbers.size > 10) {
                    summaryHtml += `<tr><td>Case/Serial Numbers</td><td>${caseNumbers.size} leases</td></tr>`;
                }
            }
        }
        
        // BLM National Recreation Sites
        if (title.includes("recreation site") || title.includes("recreation_site")) {
            const siteNames = new Set();
            const siteTypes = new Map();
            const feeCounts = new Map();
            const activityTypes = new Set();
            
            for (const row of rows) {
                // Site name variations
                const name = row.SITE_NAME || row.REC_SITE_NAME || row.NAME || 
                             row.SITENAME || row.SITE_NM || row.REC_NAME || "";
                if (name) siteNames.add(name);
                
                // Site type variations
                const type = row.SITE_TYPE || row.REC_SITE_TYPE || row.TYPE || 
                             row.SITETYPE || row.REC_TYPE || row.FACILITY_TYPE || "";
                if (type) {
                    siteTypes.set(type, (siteTypes.get(type) || 0) + 1);
                }
                
                // Fee status
                const fee = row.FEE_YN || row.FEE || row.FEE_STATUS || 
                            row.USER_FEE || row.FEES || "";
                if (fee) {
                    feeCounts.set(fee, (feeCounts.get(fee) || 0) + 1);
                }
                
                // Activities
                const activity = row.ACTIVITIES || row.ACTIVITY || row.REC_ACTIVITY ||
                                 row.PRIMARY_ACTIVITY || row.USE_TYPE || "";
                if (activity) activityTypes.add(activity);
            }
            
            if (siteNames.size > 0 || siteTypes.size > 0 || feeCounts.size > 0 || activityTypes.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Recreation Site Details</b></td></tr>`;
                
                if (siteNames.size > 0) {
                    const names = Array.from(siteNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Site Names</td><td>${names}${siteNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (siteTypes.size > 0) {
                    const items = Array.from(siteTypes.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Site Types</td><td>${items}</td></tr>`;
                }
                
                if (feeCounts.size > 0) {
                    const items = Array.from(feeCounts.entries())
                        .map(([f, count]) => `${escapeHtml(f)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Fee Status</td><td>${items}</td></tr>`;
                }
                
                if (activityTypes.size > 0) {
                    const activities = Array.from(activityTypes).slice(0, 10).map(a => escapeHtml(a)).join(", ");
                    summaryHtml += `<tr><td>Activities</td><td>${activities}${activityTypes.size > 10 ? " ..." : ""}</td></tr>`;
                }
            }
        }
        
        // BLM Land and Water Conservation Fund (LWCF)
        if (title.includes("lwcf") || title.includes("land and water conservation") || title.includes("conservation fund")) {
            const projectNames = new Set();
            const statusCounts = new Map();
            const purposeCounts = new Map();
            const fiscalYears = new Set();
            
            for (const row of rows) {
                // Project/tract name variations
                const name = row.PROJECT_NAME || row.PROJ_NAME || row.NAME || 
                             row.TRACT_NAME || row.LWCF_NAME || row.UNIT_NAME || "";
                if (name) projectNames.add(name);
                
                // Status variations
                const status = row.STATUS || row.PROJ_STATUS || row.PROJECT_STATUS ||
                               row.LWCF_STATUS || row.ACQ_STATUS || "";
                if (status) {
                    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                }
                
                // Purpose/use
                const purpose = row.PURPOSE || row.LWCF_PURPOSE || row.USE || 
                                row.PROJECT_TYPE || row.PROJ_TYPE || row.ACQ_TYPE || "";
                if (purpose) {
                    purposeCounts.set(purpose, (purposeCounts.get(purpose) || 0) + 1);
                }
                
                // Fiscal year
                const fy = row.FISCAL_YEAR || row.FY || row.YEAR || row.ACQ_YEAR || "";
                if (fy) fiscalYears.add(fy);
            }
            
            if (projectNames.size > 0 || statusCounts.size > 0 || purposeCounts.size > 0 || fiscalYears.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>LWCF Project Details</b></td></tr>`;
                
                if (projectNames.size > 0) {
                    const names = Array.from(projectNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Project/Tract Names</td><td>${names}${projectNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (purposeCounts.size > 0) {
                    const items = Array.from(purposeCounts.entries())
                        .map(([p, count]) => `${escapeHtml(p)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Purpose/Type</td><td>${items}</td></tr>`;
                }
                
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
                
                if (fiscalYears.size > 0) {
                    const years = Array.from(fiscalYears).sort().slice(0, 10).map(y => escapeHtml(String(y))).join(", ");
                    summaryHtml += `<tr><td>Fiscal Years</td><td>${years}${fiscalYears.size > 10 ? " ..." : ""}</td></tr>`;
                }
            }
        }
        
        // BLM ePlanning Projects
        if (title.includes("eplanning") || title.includes("epl_comment") || title.includes("nepa")) {
            const projectNames = new Set();
            const statusCounts = new Map();
            const typeCounts = new Map();
            const nepaNumbers = new Set();
            
            for (const row of rows) {
                const name = row.PROJECT_NAME || row.PROJ_NAME || row.NAME || "";
                if (name) projectNames.add(name);
                
                const status = row.NEPA_STATUS || row.PROJECT_STATUS || row.STATUS || "";
                if (status) {
                    statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                }
                
                const type = row.PROJECT_TYPE || row.NEPA_TYPE || row.TYPE || "";
                if (type) {
                    typeCounts.set(type, (typeCounts.get(type) || 0) + 1);
                }
                
                const nepaNo = row.NEPA_NUMBER || row.NEPA_NO || row.DOI_NUMBER || "";
                if (nepaNo) nepaNumbers.add(nepaNo);
            }
            
            if (projectNames.size > 0 || statusCounts.size > 0 || typeCounts.size > 0 || nepaNumbers.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>ePlanning Project Details</b></td></tr>`;
                
                if (projectNames.size > 0) {
                    const names = Array.from(projectNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Project Names</td><td>${names}${projectNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (typeCounts.size > 0) {
                    const items = Array.from(typeCounts.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Project Type</td><td>${items}</td></tr>`;
                }
                
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>NEPA Status</td><td>${items}</td></tr>`;
                }
                
                if (nepaNumbers.size > 0 && nepaNumbers.size <= 10) {
                    const items = Array.from(nepaNumbers).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>NEPA Numbers</td><td>${items}</td></tr>`;
                } else if (nepaNumbers.size > 10) {
                    summaryHtml += `<tr><td>NEPA Numbers</td><td>${nepaNumbers.size} projects</td></tr>`;
                }
            }
        }
        
        // BLM Fire Perimeters
        if (title.includes("fire perimeter") || title.includes("fireperimeter")) {
            const fireNames = new Set();
            const causeCounts = new Map();
            const discoveryYears = new Map();
            const adminStates = new Map();
            const totalAcres = [];
            const complexNames = new Set();
            
            for (const row of rows) {
                const name = row.INCDNT_NM || "";
                if (name) fireNames.add(name);
                
                const cause = row.FIRE_CAUSE_NM || "";
                if (cause) causeCounts.set(cause, (causeCounts.get(cause) || 0) + 1);
                
                const yr = row.FIRE_DSCVR_CY || "";
                if (yr) discoveryYears.set(String(yr), (discoveryYears.get(String(yr)) || 0) + 1);
                
                const adminSt = row.ADMIN_ST || "";
                if (adminSt) adminStates.set(adminSt, (adminStates.get(adminSt) || 0) + 1);
                
                const acres = parseFloat(row.GIS_ACRES || row.TOTAL_RPT_ACRES_NR || 0);
                if (acres > 0) totalAcres.push(acres);
                
                const cmplx = row.CMPLX_NM || "";
                if (cmplx) complexNames.add(cmplx);
            }
            
            if (fireNames.size > 0 || causeCounts.size > 0 || discoveryYears.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Fire Perimeter Details</b></td></tr>`;
                
                if (fireNames.size > 0) {
                    const names = Array.from(fireNames).slice(0, 15).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Fire Names</td><td>${names}${fireNames.size > 15 ? " ..." : ""}</td></tr>`;
                }
                
                if (causeCounts.size > 0) {
                    const items = Array.from(causeCounts.entries())
                        .map(([c, count]) => `${escapeHtml(c)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Fire Cause</td><td>${items}</td></tr>`;
                }
                
                if (discoveryYears.size > 0) {
                    const sorted = Array.from(discoveryYears.entries()).sort((a, b) => b[0].localeCompare(a[0]));
                    const items = sorted.slice(0, 15)
                        .map(([y, count]) => `${escapeHtml(y)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Discovery Years</td><td>${items}${sorted.length > 15 ? " ..." : ""}</td></tr>`;
                }
                
                if (adminStates.size > 0) {
                    const items = Array.from(adminStates.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Admin State</td><td>${items}</td></tr>`;
                }
                
                if (totalAcres.length > 0) {
                    const sum = totalAcres.reduce((a, b) => a + b, 0);
                    summaryHtml += `<tr><td>Total Burned Acres</td><td>${formatNumber(sum, 1)} (${totalAcres.length} fires)</td></tr>`;
                }
                
                if (complexNames.size > 0) {
                    const names = Array.from(complexNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Fire Complexes</td><td>${names}${complexNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
            }
        }
        
        // BLM Administrative Units (Admin Unit boundaries, office points)
        if (title.includes("administrative unit") || title.includes("admin unit") || title.includes("adminunit")) {
            const unitNames = new Set();
            const orgTypes = new Map();
            const adminStates = new Map();
            const parentNames = new Set();
            const stateUrls = new Set();
            const cityLabels = new Set();
            
            for (const row of rows) {
                const name = row.ADMU_NAME || row.Label_Full_Name || row.Label || "";
                if (name) unitNames.add(name);
                
                const orgType = row.BLM_ORG_TYPE || "";
                if (orgType) orgTypes.set(orgType, (orgTypes.get(orgType) || 0) + 1);
                
                const adminSt = row.ADMIN_ST || "";
                if (adminSt) adminStates.set(adminSt, (adminStates.get(adminSt) || 0) + 1);
                
                const parent = row.PARENT_NAME || "";
                if (parent) parentNames.add(parent);
                
                const stUrl = row.ADMU_ST_URL || "";
                if (stUrl) stateUrls.add(stUrl);
                
                const city = row.City_Label || "";
                if (city) cityLabels.add(city);
            }
            
            if (unitNames.size > 0 || orgTypes.size > 0 || adminStates.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>BLM Administrative Unit Details</b></td></tr>`;
                
                if (unitNames.size > 0) {
                    const names = Array.from(unitNames).slice(0, 15).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Unit Names</td><td>${names}${unitNames.size > 15 ? " ..." : ""}</td></tr>`;
                }
                
                if (orgTypes.size > 0) {
                    const items = Array.from(orgTypes.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Organization Type</td><td>${items}</td></tr>`;
                }
                
                if (adminStates.size > 0) {
                    const items = Array.from(adminStates.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Admin State</td><td>${items}</td></tr>`;
                }
                
                if (parentNames.size > 0) {
                    const names = Array.from(parentNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Parent Office</td><td>${names}${parentNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (cityLabels.size > 0) {
                    const cities = Array.from(cityLabels).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Office Location</td><td>${cities}${cityLabels.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (stateUrls.size > 0) {
                    const urls = Array.from(stateUrls).slice(0, 5).map(u => `<a href="${escapeHtml(u)}" target="_blank" rel="noopener">${escapeHtml(u)}</a>`).join("<br/>");
                    summaryHtml += `<tr><td>State Office URL</td><td>${urls}</td></tr>`;
                }
            }
        }
        
        // Generic fallback: If no specific handler generated content, scan for common name/status fields
        if (!summaryHtml && rows.length > 0) {
            const genericNames = new Set();
            const genericStatus = new Map();
            const genericTypes = new Map();
            
            // Common field name patterns that typically contain useful values
            const namePatterns = /^(.*_)?(NAME|NM|TITLE|LABEL|DESCRIPTION|DESC)(_.*)?$/i;
            const statusPatterns = /^(.*_)?(STATUS|STAT|STATE|CONDITION)(_.*)?$/i;
            const typePatterns = /^(.*_)?(TYPE|TYP|CLASS|CATEGORY|CAT|KIND)(_.*)?$/i;
            
            for (const row of rows) {
                for (const [key, val] of Object.entries(row)) {
                    if (val == null || val === "" || typeof val === "number") continue;
                    const strVal = String(val).trim();
                    if (!strVal || strVal.length > 200) continue; // Skip very long values
                    
                    if (namePatterns.test(key)) {
                        genericNames.add(strVal);
                    } else if (statusPatterns.test(key)) {
                        genericStatus.set(strVal, (genericStatus.get(strVal) || 0) + 1);
                    } else if (typePatterns.test(key)) {
                        genericTypes.set(strVal, (genericTypes.get(strVal) || 0) + 1);
                    }
                }
            }
            
            if (genericNames.size > 0 || genericStatus.size > 0 || genericTypes.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Feature Details</b></td></tr>`;
                
                if (genericNames.size > 0) {
                    const names = Array.from(genericNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Names</td><td>${names}${genericNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                
                if (genericTypes.size > 0) {
                    const items = Array.from(genericTypes.entries())
                        .slice(0, 10)
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Types</td><td>${items}</td></tr>`;
                }
                
                if (genericStatus.size > 0) {
                    const items = Array.from(genericStatus.entries())
                        .slice(0, 10)
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
            }
        }
        
        return summaryHtml;
    }

    // Helper: clone a renderer and force all color alphas to fully opaque
    function makeRendererOpaque(renderer) {
        if (!renderer) return renderer;
        const r = JSON.parse(JSON.stringify(renderer));
        const forceOpaque = (c) => { if (Array.isArray(c) && c.length >= 4) c[3] = (c[3] <= 1) ? 1 : 255; };
        if (r.symbol) {
            if (r.symbol.color) forceOpaque(r.symbol.color);
            if (r.symbol.outline && r.symbol.outline.color) forceOpaque(r.symbol.outline.color);
        }
        return r;
    }


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


function buildFinalReportHtmlDoc({ title, createdAt, totalsHtml, aoiSectionHtml, sectionsHtml, dataSourcesHtml }) {
    const safeTitle = escapeHtml(title || "Final Report");

    return `<!doctype html>
            <html lang="en">
            <head>
            <meta charset="utf-8"/>
            <meta name="viewport" content="width=device-width,initial-scale=1"/>
            <title>${safeTitle}</title>
            <link rel="preconnect" href="https://fonts.googleapis.com">
            <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
            <link href="https://fonts.googleapis.com/css2?family=Merriweather:wght@400;700&family=Source+Sans+Pro:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                :root{
                    --blm-green: #1a472a;
                    --blm-green-light: #2d5a3d;
                    --blm-tan: #f5f0e6;
                    --blm-gold: #c5a43e;
                    --blm-brown: #5c4827;
                    --fg: #2c2c2c;
                    --muted: #5a5a5a;
                    --border: #d4cfc4;
                    --bg: #fdfcfa;
                    --white: #ffffff;
                }
                html,body{ 
                    margin:0; 
                    padding:0; 
                    background: var(--blm-tan); 
                    color:var(--fg); 
                    font-family: 'Source Sans Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    font-size: 14px;
                    line-height: 1.5;
                }
                
                /* Header Banner */
                .report-header{
                    background: linear-gradient(135deg, var(--blm-green) 0%, var(--blm-green-light) 100%);
                    color: var(--white);
                    padding: 28px 32px;
                    margin-bottom: 0;
                }
                .report-header .agency-name{
                    font-family: 'Merriweather', Georgia, serif;
                    font-size: 13px;
                    font-weight: 400;
                    letter-spacing: 1.5px;
                    text-transform: uppercase;
                    opacity: 0.9;
                    margin-bottom: 8px;
                }
                .report-header h1{
                    font-family: 'Merriweather', Georgia, serif;
                    font-size: 28px;
                    font-weight: 700;
                    margin: 0 0 6px 0;
                    letter-spacing: 0.5px;
                }
                .report-header .meta{
                    font-size: 13px;
                    opacity: 0.85;
                    margin: 0;
                }
                
                /* Main Content Wrapper */
                .wrap{ 
                    max-width: 900px; 
                    margin: 0 auto; 
                    padding: 32px 40px 60px; 
                    background: var(--bg);
                    box-shadow: 0 0 40px rgba(0,0,0,0.08);
                    min-height: 100vh;
                }
                
                /* Section Headers */
                h2{ 
                    font-family: 'Merriweather', Georgia, serif;
                    font-size: 20px; 
                    font-weight: 700;
                    color: var(--blm-green);
                    margin: 36px 0 18px; 
                    padding-bottom: 10px;
                    border-bottom: 3px solid var(--blm-gold); 
                }
                h3{
                    font-family: 'Source Sans Pro', sans-serif;
                    font-size: 16px;
                    font-weight: 700;
                    color: var(--blm-brown);
                    margin: 24px 0 12px;
                }
                
                /* Summary Totals */
                .totals{ margin-top: 24px; }
                .totals .row{ display:flex; gap:14px; flex-wrap:wrap; margin-top:12px; }
                .pill{ 
                    border: 1px solid var(--border); 
                    border-radius: 6px; 
                    padding: 10px 16px; 
                    font-size: 13px; 
                    font-weight: 600;
                    background: var(--white);
                    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
                }
                
                /* AOI Section */
                .aoi-map{ 
                    margin: 18px 0; 
                    border: 2px solid var(--border); 
                    border-radius: 8px; 
                    overflow:hidden; 
                    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
                }
                .aoi-map img{ display:block; width:100%; height:auto; }
                .aoi-details{ margin-top: 20px; }
                .aoi-field{ margin: 10px 0; font-size: 14px; }
                .aoi-label{ 
                    font-weight: 600; 
                    color: var(--blm-green);
                    display: inline-block;
                    min-width: 160px;
                }
                .legal-list{ margin: 6px 0 0 20px; padding: 0; }
                .legal-list li{ margin: 5px 0; }

                /* Admin Unit Table in AOI Section */
                table.admin-unit-tbl{
                    width:100%;
                    border-collapse: collapse;
                    margin-top: 12px;
                    font-size: 13px;
                    background: var(--white);
                    border: 1px solid var(--border);
                    border-radius: 6px;
                    overflow: hidden;
                }
                table.admin-unit-tbl th{
                    background: var(--blm-green);
                    color: var(--white);
                    padding: 10px 12px;
                    text-align: left;
                    font-weight: 600;
                    font-size: 12px;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                table.admin-unit-tbl td{
                    padding: 8px 12px;
                    border-bottom: 1px solid var(--border);
                    vertical-align: top;
                }
                table.admin-unit-tbl tr:nth-child(even){
                    background: var(--blm-tan);
                }
                table.admin-unit-tbl tr:last-child td{
                    border-bottom: none;
                }
                
                /* Per-Layer Map Sections */
                .section{ 
                    margin-top: 32px; 
                    padding: 24px;
                    background: var(--white);
                    border: 1px solid var(--border);
                    border-radius: 8px;
                    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
                }
                .section h3{
                    margin-top: 0;
                    padding-bottom: 8px;
                    border-bottom: 1px solid var(--border);
                }
                .section .sub{ 
                    font-size: 12px; 
                    color: var(--muted); 
                    margin-bottom: 14px;
                    font-style: italic;
                }
                .map{ 
                    width:100%; 
                    border: 1px solid var(--border); 
                    border-radius: 6px; 
                    overflow:hidden; 
                    background: var(--white); 
                    margin: 16px 0;
                    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
                }
                .map img{ display:block; width:100%; height:auto; }
                
                /* Metadata Tables */
                table.metaTbl{ 
                    width:100%; 
                    border-collapse: collapse; 
                    margin-top: 16px; 
                    font-size: 13px;
                    background: var(--blm-tan);
                    border-radius: 6px;
                    overflow: hidden;
                }
                table.metaTbl td{ 
                    padding: 10px 14px; 
                    border-bottom: 1px solid var(--border); 
                }
                table.metaTbl tr:last-child td{
                    border-bottom: none;
                }
                table.metaTbl td:first-child{ 
                    color: var(--blm-green); 
                    font-weight: 600;
                    width: 200px;
                    background: rgba(26,71,42,0.05);
                }

                /* Data Sources Table */
                table.data-sources-table{ 
                    width:100%; 
                    border-collapse: collapse; 
                    margin-top: 16px; 
                    font-size: 12px; 
                    table-layout: fixed;
                    background: var(--white);
                    border: 1px solid var(--border);
                    border-radius: 6px;
                    overflow: hidden;
                }
                table.data-sources-table th{ 
                    background: var(--blm-green); 
                    color: var(--white);
                    padding: 12px 14px; 
                    text-align: left; 
                    font-weight: 600;
                    font-size: 12px;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                table.data-sources-table td{ 
                    padding: 10px 14px; 
                    border-bottom: 1px solid var(--border); 
                    vertical-align: top; 
                    word-wrap: break-word; 
                }
                table.data-sources-table tr:nth-child(even){
                    background: var(--blm-tan);
                }
                table.data-sources-table tr:last-child td{
                    border-bottom: none;
                }
                table.data-sources-table .service-desc-col{ white-space: normal; line-height: 1.5; }
                table.data-sources-table .service-url-col{ 
                    font-family: 'Consolas', 'Monaco', monospace; 
                    font-size: 10px; 
                    word-break: break-all;
                    color: var(--muted);
                }    

                .mono{ font-family: 'Consolas', 'Monaco', monospace; font-size: 11px; }
                .status-up{ color: #2e7d32; font-weight: 600; }
                .status-down{ color: #c62828; font-weight: 600; }
                
                /* Action Buttons */
                .actions{ 
                    margin-top: 20px; 
                    display:flex; 
                    gap:12px; 
                    flex-wrap: wrap; 
                }
                .btn{
                    display:inline-block; 
                    background: var(--blm-green); 
                    color: var(--white);
                    border: none;
                    border-radius: 6px;
                    padding: 12px 20px; 
                    font-size: 14px; 
                    font-weight: 600;
                    text-decoration:none;
                    cursor: pointer;
                    transition: background 0.2s ease;
                }
                .btn:hover{ background: var(--blm-green-light); }
                .hint{ 
                    font-size: 12px; 
                    color: var(--muted); 
                    margin-top: 10px;
                    font-style: italic;
                }
                
                /* Footer */
                .report-footer{
                    margin-top: 48px;
                    padding-top: 24px;
                    border-top: 2px solid var(--border);
                    font-size: 11px;
                    color: var(--muted);
                    text-align: center;
                }
                .report-footer .dept-name{
                    font-weight: 600;
                    color: var(--blm-green);
                    margin-bottom: 4px;
                }

                /* Print Styles */
                @media print{
                    html, body{ background: white; }
                    .actions, .hint{ display:none !important; }
                    .wrap{ 
                        max-width: none; 
                        padding: 0; 
                        box-shadow: none;
                        background: white;
                    }
                    .report-header{
                        background: var(--blm-green) !important;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                    .section{ 
                        break-inside: avoid; 
                        box-shadow: none;
                        border: 1px solid #ccc;
                    }
                    .pagebreak{ break-after: page; }
                    table.data-sources-table th{
                        background: var(--blm-green) !important;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                }
            </style>
            </head>
            <body>
            <div class="report-header">
                <div class="agency-name">U.S. Department of the Interior • Bureau of Land Management</div>
                <h1>${safeTitle}</h1>
                <div class="meta">Report Generated: ${escapeHtml(createdAt || "")}</div>
            </div>
            <div class="wrap">
                <div class="actions">
                    <a class="btn" href="javascript:window.print()">🖨️ Print / Save as PDF</a>
                </div>
                <div class="hint">Use your browser's print dialog and select "Save as PDF" to create a permanent copy of this report.</div>

                <h2>Report Summary</h2>
                <div class="totals">
                ${totalsHtml || ""}
                </div>

                ${aoiSectionHtml || ""}

                <h2>Layer Analysis Maps</h2>
                ${sectionsHtml || ""}

                ${dataSourcesHtml || ""}
                
                <div class="report-footer">
                    <div class="dept-name">Bureau of Land Management</div>
                    <div>U.S. Department of the Interior</div>
                    <div style="margin-top:8px;">This report was generated using geospatial data from BLM and partner agency web services.</div>
                </div>
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
        // ========================================
        // STEP 1: Compute AOI area
        // ========================================
        let aoiAcres = 0;
        try {
            const aoiSqm = Math.max(0, geometryEngine.geodesicArea(selectionGeom, "square-meters"));
            aoiAcres = aoiSqm / SQM_PER_ACRE;
        } catch (e) {
            aoiAcres = 0;
        }

// ========================================
        // STEP 2: Generate AOI Section Data
        // ========================================
        
        // 2a. Primary State (from PLSS: State Boundaries)
        let primaryState = "";
        let additionalStates = "";
        
        const stateItem = lastReportRowsByLayer.find(x => 
            x.title && x.title.toLowerCase().includes("state boundaries")
        );
        
        // ✅ Fetch full rows if not already cached
        if (stateItem && !stateItem.fullRows && stateItem._layer && stateItem._exportQuery) {
            try {
                const pageSize = config.report?.pageSize ?? 1000;
                const fullFeatures = await queryAllFeaturesPaged(
                    stateItem._layer,
                    stateItem._exportQuery,
                    pageSize,
                    100  // Cap at 100 for states
                );
                stateItem.fullRows = flattenAttributes(fullFeatures);
            } catch (e) {
                console.warn("Failed to fetch state data:", e);
            }
        }
        
        if (stateItem && stateItem.fullRows && stateItem.fullRows.length > 0) {
            const stateNames = stateItem.fullRows
                .map(r => r.NAME || r.STATE_NAME || r.STATE || r.STUSPS || "")
                .filter(Boolean);
            
            if (stateNames.length > 0) {
                primaryState = stateNames[0];
                if (stateNames.length > 1) {
                    additionalStates = stateNames.slice(1).join(", ");
                }
            }
        }

        // 2b. Legal Land Description (from PLSS: Parcel)
        const legalDescriptions = [];
        const parcelItem = lastReportRowsByLayer.find(x => 
            x.title && (x.title.toLowerCase().includes("parcel") || x.title.toLowerCase().includes("intersected"))
        );
        
        // ✅ Fetch full rows if not already cached
        if (parcelItem && !parcelItem.fullRows && parcelItem._layer && parcelItem._exportQuery) {
            try {
                const pageSize = config.report?.pageSize ?? 1000;
                const fullFeatures = await queryAllFeaturesPaged(
                    parcelItem._layer,
                    parcelItem._exportQuery,
                    pageSize,
                    100  // Cap at 100 for parcels
                );
                parcelItem.fullRows = flattenAttributes(fullFeatures);
            } catch (e) {
                console.warn("Failed to fetch parcel data:", e);
            }
        }
        
        if (parcelItem && parcelItem.fullRows && parcelItem.fullRows.length > 0) {
            for (const row of parcelItem.fullRows) {
                const legalDesc = formatLegalDescription(row);
                if (legalDesc) {
                    legalDescriptions.push(legalDesc);
                } else {
                    // Debug: if format returned empty, log the row to see what fields exist
                    console.warn("Empty legal description for parcel row:", row);
                }
            }
        }

        // 2c. AOI Method
        let aoiMethod = "Manually Drawn";
        if (aoiSource === "select") {
            const tool = plssToolLabel(aoiSourcePlssTool);
            aoiMethod = `Selected ${tool}`;
        }

        // 2d. Generate AOI Maps (zoomed out + county-level with red circle)
        setStatus("building final report… (generating AOI maps)");
        
        const aoiMapsHtml = await generateAoiMapsWithCircles();

// ========================================
        // STEP 3: Generate Map Section (existing layers)
        // ========================================
        setStatus("building final report… (generating layer maps)");
        
        // ✅ Use EXACT same extent as Visual Report (where maps were already generated)
        const paddingFactor = config?.visualReport?.paddingFactor ?? 1.25;
        const width = config?.visualReport?.screenshotWidth ?? 1400;

        let fixedExtent = null;
        const ext = selectionGeom?.extent;
        if (ext && ext.expand) {
            fixedExtent = ext.expand(paddingFactor);
        }

        // Filter out PLSS: State Boundaries from map section
        const targets = lastReportRowsByLayer
            .filter(x => (x?.count || 0) > 0)
            .filter(x => x?._layer && x?._exportQuery)
            .filter(x => !(x.title && x.title.toLowerCase().includes("state boundaries")))
            .filter(x => !(x.title && x.title.toLowerCase().includes("administrative unit")));

        // Build per-layer sections (reuse screenshots from Visual Report if possible)
        let sectionsHtml = "";

        if (!targets.length) {
            sectionsHtml = `
              <div class="section">
                <h3>No Intersecting Layers Found</h3>
                <p style="color: var(--muted); font-style: italic;">The analysis found no layers with features intersecting the selected Area of Interest.</p>
              </div>
            `;
        } else {
            // ✅ For Final Report, use imagery basemap and tight zoom on AOI only
            
            const allLayers = view.map.layers.toArray();
            const visSnapshot = allLayers.map(l => ({ layer: l, visible: l.visible }));
            const originalBasemap = view.map.basemap;
            const imageryBasemapId = config?.map?.imageryBasemap || "satellite";

            // Find PLSS Section layer to show on all per-layer maps
            let plssSectionLayer = null;
            for (const l of allLayers) {
                if (l.title && l.title.toLowerCase().includes("section")) {
                    plssSectionLayer = l;
                    break;
                }
            }

            function setVisibilityForScreenshot(tempLayer) {
                for (const l of allLayers) {
                    if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }
                    if (aoiMaskLayer && l === aoiMaskLayer) { l.visible = true; continue; }
                    if (l?.type === "tile") { l.visible = true; continue; }
                    // ✅ Show PLSS Section layer on all per-layer maps
                    if (plssSectionLayer && l === plssSectionLayer) { l.visible = true; continue; }
                    // ✅ Keep always-visible layers (e.g. BLM Admin Units) visible
                    if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }
                    l.visible = false;
                }
                if (tempLayer) tempLayer.visible = true;
                // ✅ Show AOI mask to lighten areas outside AOI
                updateAoiMask(true);
                ensureAoiOnTop(view.map);
            }

            function restoreVisibility() {
                visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
                hideAoiMask();
                ensureAoiOnTop(view.map);
            }

            // ✅ Calculate consistent extent for all per-layer maps (AOI with small buffer)
            const consistentExtent = selectionGeom.extent.expand(1.2);

            // ✅ Switch to imagery basemap ONCE before all per-layer maps
            try {
                view.map.basemap = imageryBasemapId;
                await new Promise(r => setTimeout(r, 2000));
                await waitForViewStationary(3500);
            } catch (e) {
                console.warn("Failed to switch to imagery basemap:", e);
            }

            for (let i = 0; i < targets.length; i++) {
                const item = targets[i];
                setStatus(`building final report… (${i + 1}/${targets.length})`);

                // Get geometry type for appropriate renderer — fully opaque for final report
                const tempGeomType = await getLayerGeometryType(item.url);
                const itemCfg = layerCfgByUrl.get(item.url)?.cfg || null;
                const useNativeRenderer = itemCfg?.useServiceRenderer === true;
                const opaqueRenderer = useNativeRenderer
                    ? undefined
                    : makeRendererOpaque(getPresetRenderer("report", itemCfg, tempGeomType));
                const tempOpts = {
                    url: item.url,
                    title: item.title,
                    outFields: ["*"],
                    visible: true
                };
                if (!useNativeRenderer) {
                    tempOpts.renderer = opaqueRenderer || undefined;
                }
                const temp = new FeatureLayer(tempOpts);

                // Only force scale override for non-native-renderer layers
                if (!useNativeRenderer) {
                    temp.minScale = 0;
                    temp.maxScale = 0;
                }

                view.map.add(temp);
                try {
                    setVisibilityForScreenshot(temp);
                    
                    // ✅ IMPROVED: Use comprehensive layer ready check
                    await waitForLayerReadyToCapture(temp, view, { timeoutMs: 8000 });

                    // ✅ Zoom in to consistent extent (AOI with small buffer) for all per-layer maps
                    await view.goTo(consistentExtent, { animate: false });
                    await waitForViewStationary(2000);

                    // ✅ Use improved screenshot capture with tile wait logic
                    const dataUrl = await captureScreenshotWithWait({ width });
                    if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

                    const cov = await computeLayerCoverageStats(item, selectionGeom);
                    const acresCovered = cov ? cov.acresCovered : 0;
                    const pctCovered = cov ? cov.pctAoiCovered : 0;

                    // Generate layer-specific attribute summary
                    const layerAttrSummary = generateLayerAttributeSummary(item);

                    // Generate per-feature breakdown table (only if multiple features)
                    const perFeatureTableHtml = (item.count > 1)
                        ? await buildPerFeatureTable(item, selectionGeom)
                        : "";

                    sectionsHtml += `
                    <div class="section">
                        <h3>${escapeHtml(item.title)}</h3>

                        <div class="map">
                            <img src="${dataUrl}" alt="AOI + ${escapeHtml(item.title)}"/>
                        </div>

                        <table class="metaTbl">
                            <tr><td>AOI Area</td><td><b>${formatNumber(aoiAcres, 2)}</b> acres</td></tr>
                            <tr><td>Intersecting Features</td><td><b>${escapeHtml(String(item.count || 0))}</b></td></tr>
                            <tr><td>Layer Coverage</td><td><b>${formatNumber(acresCovered, 2)}</b> acres</td></tr>
                            <tr><td>Percent of AOI Covered</td><td><b>${formatNumber(pctCovered, 2)}%</b></td></tr>
                            ${layerAttrSummary}
                        </table>
                        ${perFeatureTableHtml}
                    </div>
                    <div class="pagebreak"></div>
                    `;
                } finally {
                    try { view.map.remove(temp); } catch (e) { }
                    restoreVisibility();
                }
            }

            // ✅ Restore original basemap ONCE after all per-layer maps are done
            try {
                view.map.basemap = originalBasemap;
                await new Promise(r => setTimeout(r, 1000));
            } catch (e) {
                console.warn("Failed to restore original basemap:", e);
            }
        }



        // ========================================
        // STEP 4: Generate Data Sources Appendix
        // ========================================
        const dataSourcesHtml = buildDataSourcesSection();

        // ========================================
        // STEP 5: Build Final HTML Document
        // ========================================
        
        // Totals summary
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

        // AOI Section HTML
        const stateHtml = primaryState 
            ? `<div class="aoi-field"><span class="aoi-label">Primary State:</span> ${escapeHtml(primaryState)}${additionalStates ? ` (Additional: ${escapeHtml(additionalStates)})` : ""}</div>`
            : "";

        const legalHtml = legalDescriptions.length > 0
            ? `<div class="aoi-field">
                <span class="aoi-label">Legal Land Description:</span>
                <ul class="legal-list">
                    ${legalDescriptions.map(ld => `<li>${escapeHtml(ld)}</li>`).join("")}
                </ul>
               </div>`
            : "";

        // BLM Administrative Units info table for AOI section
        let adminUnitTableHtml = "";
        if (lastReportRowsByLayer && lastReportRowsByLayer.length) {
            const adminItems = lastReportRowsByLayer.filter(x =>
                x.title && x.title.toLowerCase().includes("administrative unit") && (x.count || 0) > 0
            );
            if (adminItems.length > 0) {
                const allAdminRows = [];
                for (const item of adminItems) {
                    for (const row of (item.rows || [])) {
                        allAdminRows.push({
                            name: row.ADMU_NAME || row.Label_Full_Name || row.Label || "",
                            orgType: row.BLM_ORG_TYPE || "",
                            adminSt: row.ADMIN_ST || "",
                            parent: row.PARENT_NAME || "",
                            city: row.City_Label || "",
                            url: row.ADMU_ST_URL || ""
                        });
                    }
                }
                // Deduplicate by name + orgType
                const seen = new Set();
                const unique = allAdminRows.filter(r => {
                    const key = `${r.name}|${r.orgType}`;
                    if (!r.name || seen.has(key)) return false;
                    seen.add(key);
                    return true;
                });
                if (unique.length > 0) {
                    const trs = unique.map(r => `
                        <tr>
                            <td>${escapeHtml(r.name)}</td>
                            <td>${escapeHtml(r.orgType)}</td>
                            <td>${escapeHtml(r.adminSt)}</td>
                            <td>${escapeHtml(r.parent)}</td>
                            <td>${escapeHtml(r.city)}</td>
                            ${r.url ? `<td><a href="${escapeHtml(r.url)}" target="_blank" rel="noopener">Link</a></td>` : `<td></td>`}
                        </tr>`).join("");
                    adminUnitTableHtml = `
                        <div style="margin-top:20px;">
                            <span class="aoi-label" style="display:block; margin-bottom:8px;">BLM Administrative Units Intersecting AOI:</span>
                            <table class="admin-unit-tbl">
                                <thead>
                                    <tr>
                                        <th>Unit Name</th>
                                        <th>Type</th>
                                        <th>State</th>
                                        <th>Parent Office</th>
                                        <th>Office Location</th>
                                        <th>URL</th>
                                    </tr>
                                </thead>
                                <tbody>${trs}</tbody>
                            </table>
                        </div>`;
                }
            }
        }

        const aoiSectionHtml = `
            <h2>Area of Interest</h2>
            ${aoiMapsHtml}
            <div class="aoi-details">
                ${stateHtml}
                ${legalHtml}
                <div class="aoi-field"><span class="aoi-label">Area:</span> ${formatNumber(aoiAcres, 2)} acres</div>
                <div class="aoi-field"><span class="aoi-label">Method:</span> ${escapeHtml(aoiMethod)}</div>
                ${adminUnitTableHtml}
            </div>
        `;

        const htmlDoc = buildFinalReportHtmlDoc({
            title: "Land & Resource Intersection Analysis Report",
            createdAt: formatDateTimeForReport(new Date()),
            totalsHtml,
            aoiSectionHtml,
            sectionsHtml,
            dataSourcesHtml
        });
    
        cachedFinalReportHtml = htmlDoc;

        if (finalReportStatus) finalReportStatus.textContent = "Report ready.";

    } catch (e) {
        console.error(e);
        if (finalReportStatus) finalReportStatus.textContent = "Failed to build report (see console).";
    }        
}

// ========================================
// HELPER: Generate AOI maps at different zoom levels
// ========================================
async function generateAoiMapsWithCircles() {
    if (!view || !selectionGeom) return "";

    const width = config?.visualReport?.screenshotWidth ?? 1400;
    const height = Math.round(width * 0.5625); // 16:9
    const maps = [];

    // Overview inset dimensions (fraction of main image)
    const insetFrac = 0.22; // 22% of width
    const insetW = Math.round(width * insetFrac);
    const insetH = Math.round(height * insetFrac);
    const insetMargin = 12;
    const overviewZoomFactor = 8; // how much wider the overview is vs the main view

    // Snapshot visibility
    const allLayers = view.map.layers.toArray();
    const visSnapshot = allLayers.map(l => ({ layer: l, visible: l.visible }));

    // Find PLSS Township layer to show on AOI maps
    let plssTownshipLayer = null;
    for (const l of allLayers) {
        if (l.title && l.title.toLowerCase().includes("township")) {
            plssTownshipLayer = l;
            break;
        }
    }

    function setVisibilityForAoi() {
        for (const l of allLayers) {
            if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }
            if (aoiMaskLayer && l === aoiMaskLayer) { l.visible = true; continue; }
            if (l?.type === "tile") { l.visible = true; continue; }
            // ✅ Show PLSS Township layer on AOI maps
            if (plssTownshipLayer && l === plssTownshipLayer) { l.visible = true; continue; }
            // ✅ Keep always-visible layers (e.g. BLM Admin Units) visible
            if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }
            l.visible = false;
        }
        // ✅ Show AOI mask to lighten areas outside AOI
        updateAoiMask(true);
        ensureAoiOnTop(view.map);
    }

    function restoreVisibility() {
        visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
        hideAoiMask();
        ensureAoiOnTop(view.map);
    }

    /**
     * Composite an overview inset onto a main screenshot.
     * @param {string} mainDataUrl - data URL of the main screenshot
     * @param {object} mainExtent  - the view.extent when main was captured
     * @param {number} scale       - the scale used for the main view
     * @returns {Promise<string>}  - composited data URL
     */
    async function compositeWithOverview(mainDataUrl, mainExtent, scale) {
        // 1. Zoom out for overview capture
        const overviewScale = scale * overviewZoomFactor;
        await view.goTo({ target: selectionGeom.extent, scale: overviewScale }, { animate: false });
        const ovSs = await captureScreenshotWithWait({ width });
        if (!ovSs) return mainDataUrl;

        // 2. Compute where the main extent falls within the overview extent
        const ovExtent = view.extent;

        // 3. Load both images
        const [mainImg, ovImg] = await Promise.all([
            loadImageFromDataUrl(mainDataUrl),
            loadImageFromDataUrl(ovSs)
        ]);
        if (!mainImg || !ovImg) return mainDataUrl;

        // 4. Draw onto canvas
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");

        // Draw main image
        ctx.drawImage(mainImg, 0, 0, width, height);

        // Inset position: lower-left with margin
        const ix = insetMargin;
        const iy = height - insetH - insetMargin;

        // Draw inset background (white border + shadow effect)
        ctx.save();
        ctx.shadowColor = "rgba(0,0,0,0.35)";
        ctx.shadowBlur = 8;
        ctx.shadowOffsetX = 2;
        ctx.shadowOffsetY = 2;
        ctx.fillStyle = "#fff";
        ctx.fillRect(ix - 2, iy - 2, insetW + 4, insetH + 4);
        ctx.restore();

        // Draw overview image clipped to inset area
        ctx.drawImage(ovImg, ix, iy, insetW, insetH);

        // 5. Draw red extent rectangle on overview
        // Map mainExtent coords → pixel coords within the inset
        if (ovExtent && mainExtent) {
            const ovXMin = ovExtent.xmin;
            const ovYMin = ovExtent.ymin;
            const ovW = ovExtent.xmax - ovExtent.xmin;
            const ovH = ovExtent.ymax - ovExtent.ymin;

            if (ovW > 0 && ovH > 0) {
                const rxPx = ix + ((mainExtent.xmin - ovXMin) / ovW) * insetW;
                const ryPx = iy + ((ovExtent.ymax - mainExtent.ymax) / ovH) * insetH;
                const rwPx = ((mainExtent.xmax - mainExtent.xmin) / ovW) * insetW;
                const rhPx = ((mainExtent.ymax - mainExtent.ymin) / ovH) * insetH;

                ctx.strokeStyle = "#e63946";
                ctx.lineWidth = 2.5;
                ctx.strokeRect(rxPx, ryPx, rwPx, rhPx);
            }
        }

        // 6. Thin border around inset
        ctx.strokeStyle = "#333";
        ctx.lineWidth = 1.5;
        ctx.strokeRect(ix - 1, iy - 1, insetW + 2, insetH + 2);

        return canvas.toDataURL("image/png");
    }

    /** Load an Image from a data URL */
    function loadImageFromDataUrl(dataUrl) {
        return new Promise(resolve => {
            const img = new Image();
            img.onload = () => resolve(img);
            img.onerror = () => resolve(null);
            img.src = dataUrl;
        });
    }

    try {
        setVisibilityForAoi();

        // ✅ Map 1: 1:900,000 scale (showing regional context)
        const ext1 = selectionGeom.extent;
        await view.goTo({ target: ext1, scale: 900000 }, { animate: false });
        const ss1 = await captureScreenshotWithWait({ width });
        const mainExtent1 = view.extent.clone();

        if (ss1) {
            const composited1 = await compositeWithOverview(ss1, mainExtent1, 900000);
            maps.push(`<div class="aoi-map"><img src="${composited1}" alt="AOI Context (Regional 1:900,000)" /></div>`);
        }

        // ✅ Map 2: 1:200,000 scale (county-level zoom)
        const ext2 = selectionGeom.extent;
        await view.goTo({ target: ext2, scale: 200000 }, { animate: false });
        const ss2 = await captureScreenshotWithWait({ width });
        const mainExtent2 = view.extent.clone();

        if (ss2) {
            const composited2 = await compositeWithOverview(ss2, mainExtent2, 200000);
            maps.push(`<div class="aoi-map"><img src="${composited2}" alt="AOI Context (County 1:200,000)" /></div>`);
        }

    } finally {
        restoreVisibility();
    }

    return maps.join("");
}




// ========================================
// HELPER: Build data sources appendix (IMPROVED TABLE)
// ========================================
function buildDataSourcesSection() {
    const services = getConfiguredServices();
    
    const rows = services.map(svc => {
        const status = serviceStatus.get(svc.url) || "UNKNOWN";
        const statusClass = status === "UP" ? "status-up" : "status-down";
        const desc = serviceStatus.get(svc.url + "::desc") || "(Description not available)";
        
        return `
            <tr>
                <td class="service-name-col">${escapeHtml(svc.title)}</td>
                <td class="service-url-col"><a href="${escapeHtml(svc.url)}" target="_blank" rel="noopener">${escapeHtml(svc.url)}</a></td>
                <td class="service-desc-col">${escapeHtml(desc)}</td>
                <td class="service-status-col"><span class="${statusClass}">${status}</span></td>
            </tr>
        `;
    }).join("");

    return `
        <div class="section" style="background: transparent; border: none; box-shadow: none; padding: 0;">
            <h2>Data Sources</h2>
            <p style="font-size: 13px; color: var(--muted); margin-bottom: 16px;">
                The following geospatial web services were used to generate this report. Service availability was verified at the time of report generation.
            </p>
            <table class="data-sources-table">
                <thead>
                    <tr>
                        <th style="width: 20%;">Service Name</th>
                        <th style="width: 25%;">Service URL</th>
                        <th style="width: 45%;">Description</th>
                        <th style="width: 10%;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows}
                </tbody>
            </table>
        </div>
    `;
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

    // ---------- Feature Picker Modal (for overlapping polygons) ----------
    const featurePickerModal = document.getElementById("featurePickerModal");
    const featurePickerList = document.getElementById("featurePickerList");
    const featurePickerCancelBtn = document.getElementById("featurePickerCancelBtn");
    const featurePickerConfirmBtn = document.getElementById("featurePickerConfirmBtn");

    // State for feature picker
    let pickerFeatures = [];
    let pickerSelectedIdx = -1;
    let pickerOnSelect = null;

    function showFeaturePicker(features, onSelect) {
        if (!featurePickerModal || !featurePickerList) return;

        // Store state
        pickerFeatures = features;
        pickerSelectedIdx = -1;
        pickerOnSelect = onSelect;

        // Reset confirm button
        if (featurePickerConfirmBtn) {
            featurePickerConfirmBtn.disabled = true;
            featurePickerConfirmBtn.textContent = "Select This Polygon";
        }

        // Build the list of features
        featurePickerList.innerHTML = features.map((f, idx) => {
            const attrs = f.graphic?.attributes || {};
            const name = getFeaturePickerDisplayName(attrs);
            const details = getFeaturePickerDetails(attrs);
            
            return `
                <div class="feature-picker-item" data-idx="${idx}">
                    <div class="feature-picker-index">${idx + 1}</div>
                    <div class="feature-picker-info">
                        <div class="feature-picker-name">${escapeHtml(name)}</div>
                        ${details ? `<div class="feature-picker-details">${escapeHtml(details)}</div>` : ""}
                    </div>
                </div>
            `;
        }).join("");

        // Add click handlers for preview (not immediate selection)
        const items = featurePickerList.querySelectorAll(".feature-picker-item");
        items.forEach((item) => {
            item.addEventListener("click", () => {
                const idx = parseInt(item.getAttribute("data-idx"), 10);
                selectPickerRow(idx);
            });
        });

        featurePickerModal.classList.remove("hidden");
    }

    function selectPickerRow(idx) {
        if (idx < 0 || idx >= pickerFeatures.length) return;

        pickerSelectedIdx = idx;

        // Update visual selection state
        const items = featurePickerList.querySelectorAll(".feature-picker-item");
        items.forEach((item, i) => {
            item.classList.toggle("selected", i === idx);
        });

        // Highlight the selected polygon on the map
        highlightPickerFeature(pickerFeatures[idx]?.graphic);

        // Enable confirm button and update text
        if (featurePickerConfirmBtn) {
            featurePickerConfirmBtn.disabled = false;
            const name = getFeaturePickerDisplayName(pickerFeatures[idx]?.graphic?.attributes || {});
            const shortName = name.length > 30 ? name.substring(0, 27) + "..." : name;
            featurePickerConfirmBtn.textContent = `Select "${shortName}"`;
        }
    }

    function confirmPickerSelection() {
        if (pickerSelectedIdx < 0 || !pickerFeatures[pickerSelectedIdx]) return;

        const selectedFeature = pickerFeatures[pickerSelectedIdx];
        hideFeaturePicker();

        if (pickerOnSelect) {
            pickerOnSelect(selectedFeature);
        }
    }

    function hideFeaturePicker() {
        if (featurePickerModal) {
            featurePickerModal.classList.add("hidden");
        }
        clearPickerHighlight();
        pickerFeatures = [];
        pickerSelectedIdx = -1;
        pickerOnSelect = null;
    }

    // Highlight layer for picker hover
    let pickerHighlightLayer = null;

    function highlightPickerFeature(graphic) {
        if (!graphic || !map || !view) return;
        
        clearPickerHighlight();
        
        pickerHighlightLayer = new GraphicsLayer({ title: "Picker Highlight" });
        map.add(pickerHighlightLayer);
        
        const highlightGraphic = new Graphic({
            geometry: graphic.geometry,
            symbol: {
                type: "simple-fill",
                color: [0, 200, 100, 0.35],
                outline: { color: [0, 150, 50], width: 3 }
            }
        });
        pickerHighlightLayer.add(highlightGraphic);
    }

    function clearPickerHighlight() {
        if (pickerHighlightLayer && map) {
            try { map.remove(pickerHighlightLayer); } catch (e) {}
            pickerHighlightLayer = null;
        }
    }

    // Get display name for feature picker
    function getFeaturePickerDisplayName(attrs) {
        const nameFields = [
            "NAME", "Name", "name",
            "ALLOT_NAME", "ALLOTMENT_NAME", "Allotment",
            "PASTURE_NAME", "PASTURE",
            "LEASE_NAME", "LEASE_NUM", "CASE_FILE_N",
            "PLAN_NAME", "UNIT_NAME", "AREA_NAME",
            "LABEL", "TITLE", "DESCRIPTION"
        ];
        
        for (const field of nameFields) {
            if (attrs[field] && String(attrs[field]).trim()) {
                return String(attrs[field]).trim();
            }
        }
        
        // Fallback
        for (const [key, val] of Object.entries(attrs)) {
            if (typeof val === "string" && val.trim() &&
                !key.toLowerCase().includes("objectid") &&
                !key.toLowerCase().includes("globalid") &&
                !key.toLowerCase().includes("shape")) {
                return val.trim().substring(0, 60);
            }
        }
        
        return "Unnamed Feature";
    }

    // Get details for feature picker
    function getFeaturePickerDetails(attrs) {
        const detailParts = [];
        const skipFields = ["OBJECTID", "GLOBALID", "SHAPE", "SHAPE_LENGTH", "SHAPE_AREA"];
        
        let count = 0;
        for (const [key, val] of Object.entries(attrs)) {
            if (count >= 2) break;
            if (skipFields.some(s => key.toUpperCase().includes(s))) continue;
            if (val && String(val).trim() && typeof val !== "object") {
                const displayVal = String(val).trim();
                if (displayVal.length <= 50) {
                    detailParts.push(`${key}: ${displayVal}`);
                    count++;
                }
            }
        }
        
        return detailParts.join(" • ");
    }

    // Cancel button handler
    if (featurePickerCancelBtn) {
        featurePickerCancelBtn.addEventListener("click", hideFeaturePicker);
    }

    // Confirm button handler
    if (featurePickerConfirmBtn) {
        featurePickerConfirmBtn.addEventListener("click", confirmPickerSelection);
    }

    // Close on click outside
    if (featurePickerModal) {
        featurePickerModal.addEventListener("click", (e) => {
            if (e.target === featurePickerModal) {
                hideFeaturePicker();
            }
        });
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

                // Get ALL matching features from the active selection layer
                const matches = results.filter(r =>
                    r.graphic && r.graphic.layer && activeSelectionLayer && r.graphic.layer === activeSelectionLayer
                );

                if (matches.length === 0) {
                    return;

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
        const loadingOverlay = document.getElementById("appLoadingOverlay");
        const loadingStatus = document.getElementById("appLoadingStatus");
        const loadingBar = document.getElementById("appLoadingBarFill");

        function setLoadingState(text, pct) {
            if (loadingStatus) loadingStatus.textContent = text;
            if (loadingBar) loadingBar.style.width = pct + "%";
        }

        function hideLoadingOverlay() {
            if (loadingOverlay) {
                loadingOverlay.classList.add("fade-out");
                setTimeout(() => { loadingOverlay.remove(); }, 700);
            }
        }

        setLoadingState("Loading configuration...", 5);
        setStatus("loading config…");

        config = await fetchJson("./config.json");
        layerCfgByUrl = buildLayerCfgIndex(config);

        setLoadingState("Initializing map...", 20);

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


        // Basemap config
        const imageryBasemapId = config?.map?.imageryBasemap || "satellite";
        const imageryOpacity = config?.map?.imageryOpacity ?? 0.75;
        const defaultBasemapId = config.map?.basemap || "gray-vector";

        // Enforce imagery opacity when imagery is active (and restore for non-imagery)
        view.watch("map.basemap", (bm) => {
            if (!bm) return;
            if (isImageryBasemap(bm)) setBasemapBaseLayerOpacity(bm, imageryOpacity);
            else setBasemapBaseLayerOpacity(bm, 1);
        });

        // Apply once on load
        if (isImageryBasemap(view.map.basemap)) setBasemapBaseLayerOpacity(view.map.basemap, imageryOpacity);

        // ---------- Esri Overview Map ----------
        // Second MapView placed into the main view's UI at top-left.
        // Shows the opposite basemap of the main map.
        // A GraphicsLayer draws the main view's extent as a red rectangle.
        // Click to swap basemaps between main and overview.


        // Overview map: fixed extent/zoom (contiguous US), red rectangle shows main map extent
        const overviewExtentLayer = new GraphicsLayer();
        let overviewMap = new EsriMap({
            basemap: imageryBasemapId,
            layers: [overviewExtentLayer]
        });
        const overviewView = new MapView({
            container: "overviewMapView",
            map: overviewMap,
            ui: { components: [] },
            constraints: { snapToZoom: false, rotationEnabled: false },
            center: [-98.5795, 39.8283], // Center of contiguous US
            zoom: 2 // Zoomed out to show full US with margin
        });

        // Disable all interaction on the overview
        overviewView.when(() => {
            const nav = overviewView.navigation;
            if (nav) {
                nav.mouseWheelEnabled = false;
                nav.browserTouchPanEnabled = false;
                nav.dragPanEnabled = false;
                nav.keyboardEnabled = false;
                nav.doubleClickZoomEnabled = false;
            }
        });

        // Place the overview div into the main view's UI
        const overviewDiv = document.getElementById("overviewDiv");
        if (overviewDiv) {
            view.ui.add(overviewDiv, "bottom-left");
        }


        // --- Draw red rectangle for main map extent on minimap ---
        function updateOverviewExtentGraphic() {
            if (!view.extent) return;
            overviewExtentLayer.removeAll();
            const extentGraphic = new Graphic({
                geometry: view.extent.clone(),
                symbol: {
                    type: "simple-line",
                    color: [255, 0, 0, 1],
                    width: 2
                }
            });
            overviewExtentLayer.add(extentGraphic);
        }

        // Update rectangle when main view moves
        view.watch("extent", updateOverviewExtentGraphic);
        // Initial draw
        view.when(() => overviewView.when(updateOverviewExtentGraphic));

        // --- Click to swap basemaps ---
        if (overviewDiv) {
            overviewDiv.addEventListener("click", async () => {
                const mainIsImagery = isImageryBasemap(view.map.basemap);
                if (mainIsImagery) {
                    view.map.basemap = defaultBasemapId;
                    overviewMap.basemap = imageryBasemapId;
                } else {
                    view.map.basemap = imageryBasemapId;
                    overviewMap.basemap = defaultBasemapId;
                }

                // Wait for main view to settle after basemap change
                await waitForViewStationary(2000);

                // Refresh all visible operational layers so they redraw on the new basemap
                const allLayers = view.map.layers.toArray();
                for (const lyr of allLayers) {
                    if (!lyr.visible) continue;
                    try {
                        const lv = await view.whenLayerView(lyr);
                        if (typeof lv.refresh === "function") lv.refresh();
                    } catch (e) { /* ignore layers that can't refresh */ }
                }

                // Ensure AOI stays on top and extent indicator is current
                ensureAoiOnTop(map);
                updateOverviewExtentGraphic();
            });
        }

        // AOI layer + sketch (AOI must always be visible and on top)
        aoiLayer = new GraphicsLayer({ title: "AOI" });
        map.add(aoiLayer);

        // AOI Mask layer (transparent overlay outside AOI for reports)
        aoiMaskLayer = new GraphicsLayer({ title: "AOI Mask", visible: false });
        map.add(aoiMaskLayer);

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

        // Selection layers (Living Atlas PLSS FeatureServer — listed individually in config)
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

        setLoadingState("Loading report layers...", 45);

        // ✅ NEW: build report layers (for map display toggles)
        await buildReportDisplayLayers();

        renderLayerToggles(map);
        ensureAoiOnTop(map);

        setLoadingState("Waiting for map view...", 65);

        await view.when();
        setLoadingState("Setting up tools...", 80);
        attachClickToSelect();

        // ---------- PLSS tool wiring (Township / Section / Intersected) ----------
        const townshipIdx = findSelectionLayerIndexByNameIncludes("township");
        const sectionIdx = findSelectionLayerIndexByNameIncludes("section");
        const intersectedIdx =
            (findSelectionLayerIndexByNameIncludes("parcel") >= 0)
                ? findSelectionLayerIndexByNameIncludes("parcel")
                : findSelectionLayerIndexByNameIncludes("intersected");

        plssParcelLayerUrl = (intersectedIdx >= 0) ? (selectionLayers[intersectedIdx]?.cfg?.url || null) : null;

        // ---------- Permit layer indexes (Grazing Allotment / Pasture / Oil & Gas) ----------
        const allotmentIdx = findSelectionLayerIndexByNameIncludes("grazing allotment");
        const pastureIdx = findSelectionLayerIndexByNameIncludes("grazing pasture");
        const oilGasIdx = findSelectionLayerIndexByNameIncludes("oil and gas");


        // Helper: make ONE PLSS layer active, disable the other two, and auto-zoom if needed
        async function activatePlss(which, idxToEnable) {
            // Force select mode (PLSS tools are select-only)
            if (modeSelect && modeSelect.value !== "select") {
                modeSelect.value = "select";
                setMode("select");
            }

            // Disable all permit layers first
            const permitTrio = [allotmentIdx, pastureIdx, oilGasIdx].filter(i => i >= 0);
            for (const idx of permitTrio) disableSelectionLayer(idx);

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
                setPermitToolActive(null); // Clear permit tool selection

                // Auto-zoom to minimum visible zoom level (using layer.minScale)
                const lyr = selectionLayers[idxToEnable]?.layer;
                await ensureLayerVisibleAtScale(lyr);
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

        // ---------- Permit layer wiring (Grazing Allotment / Pasture / Oil & Gas) ----------
        // Helper: make ONE permit layer active, disable the other permit layers, and auto-zoom if needed
        async function activatePermitLayer(which, idxToEnable) {
            // Force select mode
            if (modeSelect && modeSelect.value !== "select") {
                modeSelect.value = "select";
                setMode("select");
            }

            // Disable all PLSS layers first
            const plssTrio = [townshipIdx, sectionIdx, intersectedIdx].filter(i => i >= 0);
            for (const idx of plssTrio) disableSelectionLayer(idx);

            // Disable other permit layers
            const permitTrio = [allotmentIdx, pastureIdx, oilGasIdx].filter(i => i >= 0);
            for (const idx of permitTrio) {
                if (idx !== idxToEnable) disableSelectionLayer(idx);
            }

            // Enable the chosen one
            if (idxToEnable >= 0) enableSelectionLayer(idxToEnable);

            // Set as active selection layer
            if (idxToEnable >= 0) {
                await setActiveSelectionLayerByIndex(idxToEnable);
                aoiSourcePlssTool = null; // Not a PLSS tool
                setPlssToolActive(null);
                setPermitToolActive(which);

                // Auto-zoom to minimum visible zoom level
                const lyr = selectionLayers[idxToEnable]?.layer;
                await ensureLayerVisibleAtScale(lyr);
                await waitForViewStationary(1500);

                const labels = {
                    allotment: "Grazing Allotments",
                    pasture: "Grazing Pastures",
                    oilgas: "Oil & Gas Leases"
                };
                setStatus(`Permit select: ${labels[which] || which} (click a polygon)`);
            } else {
                setPermitToolActive(which);
                setStatus("Permit select: layer not found");
            }
        }

        if (grazingAllotmentBtn) grazingAllotmentBtn.addEventListener("click", () => activatePermitLayer("allotment", allotmentIdx));
        if (grazingPastureBtn) grazingPastureBtn.addEventListener("click", () => activatePermitLayer("pasture", pastureIdx));
        if (oilGasLeaseBtn) oilGasLeaseBtn.addEventListener("click", () => activatePermitLayer("oilgas", oilGasIdx));

        // Selection group dropdown handler
        if (selectionGroupSelect) {
            selectionGroupSelect.addEventListener("change", async () => {
                const group = selectionGroupSelect.value;
                switchSelectionGroup(group);

                if (group === "plss") {
                    // Activate township by default when switching to PLSS
                    if (townshipIdx >= 0) {
                        await activatePlss("township", townshipIdx);
                    } else if (sectionIdx >= 0) {
                        await activatePlss("section", sectionIdx);
                    } else if (intersectedIdx >= 0) {
                        await activatePlss("intersected", intersectedIdx);
                    }
                } else if (group === "permit") {
                    // Activate grazing allotment by default when switching to permit layers
                    if (allotmentIdx >= 0) {
                        await activatePermitLayer("allotment", allotmentIdx);
                    } else if (pastureIdx >= 0) {
                        await activatePermitLayer("pasture", pastureIdx);
                    } else if (oilGasIdx >= 0) {
                        await activatePermitLayer("oilgas", oilGasIdx);
                    }
                }
            });
        }

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


        // ========================================
        // FEATURE SEARCH WIDGET
        // ========================================
        const searchInput = document.getElementById("featureSearchInput");
        const searchResults = document.getElementById("featureSearchResults");
        const searchClear = document.getElementById("featureSearchClear");
        const searchIcon = document.getElementById("featureSearchIcon");
        const searchSpinner = document.getElementById("featureSearchSpinner");

        let searchDebounceTimer = null;
        let searchAbortController = null;
        let searchGeneration = 0; // tracks which search is "current" to avoid race conditions
        const fieldMetadataCache = new Map(); // url -> string field names (avoids re-fetching)

        // Get all searchable layer URLs from config
        function getSearchableLayers() {
            const layers = [];
            
            // Add report layers
            (config.reportLayers || []).forEach(cfg => {
                if (cfg?.url) {
                    layers.push({ title: cfg.title || "Unknown Layer", url: cfg.url, type: "report" });
                }
            });
            
            // Add selection layers (expanded version already in selectionLayers array)
            (selectionLayers || []).forEach(entry => {
                if (entry?.cfg?.url) {
                    // Avoid duplicates
                    const exists = layers.some(l => l.url === entry.cfg.url);
                    if (!exists) {
                        layers.push({ title: entry.cfg.title || "Unknown Layer", url: entry.cfg.url, type: "selection" });
                    }
                }
            });
            
            return layers;
        }

        // Patterns for fields that should be prioritized in search (name-like fields)
        const NAME_FIELD_PATTERNS = [
            /^name$/i, /name$/i, /_name$/i, /name_/i,
            /^title$/i, /^label$/i, /^description$/i, /^desc$/i,
            /comname/i, /sciname/i, /common.*name/i, /scientific.*name/i,
            /plan.*name/i, /proj.*name/i, /unit.*name/i, /area.*name/i,
            /site.*name/i, /allot.*name/i, /lup.*name/i, /permit/i
        ];

        // Fields that should be excluded from search entirely (IDs, codes, internal fields)
        const EXCLUDED_FIELD_PATTERNS = [
            /objectid/i, /globalid/i, /^oid$/i, /^fid$/i, /^id$/i,
            /shape/i, /geometry/i, /^guid$/i, /uuid/i,
            /_id$/i, /^.*id$/i, /code$/i, /_code$/i, /^code/i,
            /serial/i, /row_?num/i, /unique/i, /key$/i,
            /created/i, /modified/i, /edit.*date/i, /update/i,
            /^gis_/i, /^sys_/i, /^db_/i, /^meta/i
        ];

        // Categorize fields into name fields (high priority) vs other searchable fields
        function categorizeSearchFields(fields) {
            const nameFields = [];
            const otherFields = [];

            for (const field of fields) {
                const fname = field.name || "";
                
                // Skip excluded fields
                if (EXCLUDED_FIELD_PATTERNS.some(p => p.test(fname))) continue;
                
                // Check if it's a name-like field
                if (NAME_FIELD_PATTERNS.some(p => p.test(fname))) {
                    nameFields.push(fname);
                } else {
                    otherFields.push(fname);
                }
            }

            return { nameFields, otherFields };
        }

        // Get string field names from a layer's field metadata (cached)
        async function getStringFieldsForLayer(url) {
            const cacheKey = url.replace(/\/$/, "");
            if (fieldMetadataCache.has(cacheKey)) return fieldMetadataCache.get(cacheKey);
            try {
                const pjsonUrl = cacheKey + "?f=pjson";
                const info = await fetchJson(pjsonUrl);
                const fields = info?.fields || [];
                
                // Get string fields and categorize them
                const stringFields = fields.filter(f => f.type === "esriFieldTypeString");
                const categorized = categorizeSearchFields(stringFields);
                
                fieldMetadataCache.set(cacheKey, categorized);
                return categorized;
            } catch (e) {
                console.warn("Failed to get fields for", url, e);
                return { nameFields: [], otherFields: [] };
            }
        }

        // Check if a result has a meaningful name match (not just ID match)
        function hasNameFieldMatch(attributes, searchTerm, nameFields) {
            const termLower = searchTerm.toLowerCase();
            for (const field of nameFields) {
                const val = attributes[field];
                if (val && String(val).toLowerCase().includes(termLower)) {
                    return true;
                }
            }
            return false;
        }

        // Calculate relevance score for a search result
        function calculateRelevance(attributes, searchTerm, nameFields) {
            const termLower = searchTerm.toLowerCase();
            let score = 0;
            
            // Check name fields for matches (high value)
            for (const field of nameFields) {
                const val = String(attributes[field] || "").toLowerCase();
                if (val) {
                    if (val === termLower) score += 100; // Exact match
                    else if (val.startsWith(termLower)) score += 50; // Starts with
                    else if (val.includes(termLower)) score += 25; // Contains
                }
            }
            
            // Check if display name would show the match (important for UX)
            const displayName = getFeatureDisplayName(attributes).toLowerCase();
            if (displayName.includes(termLower)) {
                score += 30;
            }
            
            return score;
        }

        // Search a single layer for matching features
        async function searchLayer(layerInfo, searchTerm, signal, maxResults = 5) {
            try {
                const { nameFields, otherFields } = await getStringFieldsForLayer(layerInfo.url);
                
                // If no searchable fields, skip this layer
                if (!nameFields.length && !otherFields.length) return [];

                const escapedTerm = searchTerm.replace(/'/g, "''");
                
                // Build WHERE clause - prioritize name fields, but include other fields too
                // We search both but will score/filter results later
                const allSearchFields = [...nameFields, ...otherFields];
                const whereClauses = allSearchFields.map(f => `UPPER(${f}) LIKE '%${escapedTerm.toUpperCase()}%'`);
                const where = whereClauses.join(" OR ");

                const queryUrl = layerInfo.url.replace(/\/$/, "") + "/query";
                const params = new URLSearchParams({
                    where,
                    outFields: "*",
                    returnGeometry: "true",
                    outSR: String(view?.spatialReference?.wkid || 102100),
                    resultRecordCount: String(maxResults * 2), // Fetch extra to filter
                    f: "json"
                });

                const response = await fetch(`${queryUrl}?${params.toString()}`, { signal, credentials: "omit" });
                if (!response.ok) return [];
                
                const data = await response.json();
                const features = data?.features || [];

                // Map features and calculate relevance
                const results = features.map(f => {
                    const attrs = f.attributes || {};
                    const relevance = calculateRelevance(attrs, searchTerm, nameFields);
                    const hasNameMatch = hasNameFieldMatch(attrs, searchTerm, nameFields);
                    
                    return {
                        layerTitle: layerInfo.title,
                        layerUrl: layerInfo.url,
                        attributes: attrs,
                        geometry: f.geometry,
                        relevance,
                        hasNameMatch
                    };
                });

                // Filter: prefer results with name field matches or visible display name matches
                // Only include low-relevance results if they have something meaningful to show
                const filtered = results.filter(r => {
                    // Always keep if there's a name field match
                    if (r.hasNameMatch) return true;
                    
                    // Keep if the display name contains the search term
                    const displayName = getFeatureDisplayName(r.attributes).toLowerCase();
                    if (displayName.includes(searchTerm.toLowerCase())) return true;
                    
                    // Filter out results where the match is only in non-name fields
                    // and the display name doesn't show the match (confusing for users)
                    return false;
                });

                // Sort by relevance and limit results
                return filtered
                    .sort((a, b) => b.relevance - a.relevance)
                    .slice(0, maxResults);

            } catch (e) {
                // Return empty for all errors - the performSearch wrapper handles logging
                if (e.name !== "AbortError") {
                    console.warn("Search failed for layer:", layerInfo.title, e);
                }
                return [];
            }
        }

        // Get a display name for a feature from its attributes
        function getFeatureDisplayName(attributes) {
            // Common name fields in priority order
            const nameFields = [
                "NAME", "Name", "name",
                "PLAN_NAME", "LUPName", "LUPNAME",
                "ALLOT_NAME", "ALLOTMENT_NAME",
                "COMNAME", "SCINAME", "comname", "sciname",
                "UNIT_NAME", "AREA_NAME", "SITE_NAME",
                "PROJ_NAME", "PROJECT_NAME",
                "CASEFILE_N", "CASE_FILE",
                "LABEL", "Label", "TITLE", "Title",
                "DESCRIPTION", "DESC"
            ];

            for (const field of nameFields) {
                if (attributes[field] && String(attributes[field]).trim()) {
                    return String(attributes[field]).trim();
                }
            }

            // Fallback: use first non-ID string attribute
            for (const [key, val] of Object.entries(attributes)) {
                if (typeof val === "string" && val.trim() && 
                    !key.toLowerCase().includes("objectid") &&
                    !key.toLowerCase().includes("globalid") &&
                    !key.toLowerCase().includes("shape")) {
                    return val.trim().substring(0, 80);
                }
            }

            return "Unnamed Feature";
        }

        // Get additional details for display
        function getFeatureDetails(attributes) {
            const details = [];
            const skipFields = ["OBJECTID", "GLOBALID", "SHAPE", "SHAPE_LENGTH", "SHAPE_AREA", "SHAPE.LEN", "SHAPE.AREA"];
            
            let count = 0;
            for (const [key, val] of Object.entries(attributes)) {
                if (count >= 2) break;
                if (skipFields.some(s => key.toUpperCase().includes(s))) continue;
                if (val && String(val).trim()) {
                    details.push(`${key}: ${String(val).trim().substring(0, 40)}`);
                    count++;
                }
            }
            
            return details.join(" | ");
        }

        // Store all search results for click handling
        let allSearchResults = [];

        // Perform the search across all layers
        async function performSearch(searchTerm) {
            if (!searchTerm || searchTerm.length < 2) {
                searchResults.innerHTML = '<div class="search-hint">Type at least 2 characters to search</div>';
                searchResults.classList.add("visible");
                // Reset spinner in case a previous search was in progress
                if (searchIcon) searchIcon.style.display = "block";
                if (searchSpinner) searchSpinner.style.display = "none";
                return;
            }

            // Cancel previous search
            if (searchAbortController) {
                try { searchAbortController.abort(); } catch (e) { }
                searchAbortController = null;
            }
            searchAbortController = new AbortController();
            const signal = searchAbortController.signal;
            const myGen = ++searchGeneration; // snapshot our generation

            // Show loading state
            if (searchIcon) searchIcon.style.display = "none";
            if (searchSpinner) searchSpinner.style.display = "block";

            try {
                const layers = getSearchableLayers();
                
                // Search all layers in parallel (limit to first 15 to avoid too many requests)
                // Use Promise.allSettled to avoid one failure killing all searches
                const searchPromises = layers.slice(0, 15).map(layerInfo => 
                    searchLayer(layerInfo, searchTerm, signal, 5).catch(e => {
                        // Swallow errors (including AbortError) and return empty results
                        if (e.name !== "AbortError") {
                            console.warn("Search failed for layer:", layerInfo.title, e);
                        }
                        return [];
                    })
                );

                const results = await Promise.all(searchPromises);
                
                // Check if this search was superseded
                if (signal.aborted || myGen !== searchGeneration) {
                    return;
                }

                // Flatten all results and sort by global relevance
                allSearchResults = [];
                results.forEach((layerResults) => {
                    layerResults.forEach(f => allSearchResults.push(f));
                });
                
                // Sort all results by relevance (best matches first across all layers)
                allSearchResults.sort((a, b) => (b.relevance || 0) - (a.relevance || 0));
                
                // Limit total results to top 25
                allSearchResults = allSearchResults.slice(0, 25);
                
                // Re-group by layer for display, preserving relevance order within groups
                const groupedResults = new Map();
                allSearchResults.forEach(f => {
                    const layerTitle = f.layerTitle;
                    if (!groupedResults.has(layerTitle)) {
                        groupedResults.set(layerTitle, []);
                    }
                    groupedResults.get(layerTitle).push(f);
                });

                // Build results HTML
                if (groupedResults.size === 0) {
                    searchResults.innerHTML = '<div class="search-no-results">No matching features found</div>';
                } else {
                    let html = "";
                    let globalIdx = 0;
                    for (const [layerTitle, features] of groupedResults) {
                        html += `<div class="search-result-group">`;
                        html += `<div class="search-result-group-title">${escapeHtml(layerTitle)}</div>`;
                        
                        features.forEach((feature) => {
                            const name = getFeatureDisplayName(feature.attributes);
                            const details = getFeatureDetails(feature.attributes);
                            
                            html += `<div class="search-result-item" data-result-idx="${globalIdx}">`;
                            html += `<div class="search-result-name">${escapeHtml(name)}</div>`;
                            if (details) {
                                html += `<div class="search-result-details">${escapeHtml(details)}</div>`;
                            }
                            html += `</div>`;
                            globalIdx++;
                        });
                        
                        html += `</div>`;
                    }
                    searchResults.innerHTML = html;

                    // Attach click handlers to results
                    const resultItems = searchResults.querySelectorAll(".search-result-item");
                    resultItems.forEach((item) => {
                        item.addEventListener("click", async () => {
                            const idx = parseInt(item.getAttribute("data-result-idx"), 10);
                            const feature = allSearchResults[idx];
                            if (feature) {
                                // Enable the layer if it's not visible
                                await enableLayerByUrl(feature.layerUrl);
                                // Zoom to the feature
                                await zoomToFeature(feature);
                            }
                            searchResults.classList.remove("visible");
                        });
                    });
                }

                searchResults.classList.add("visible");

            } catch (e) {
                // If this search was superseded, don't show error state
                if (myGen !== searchGeneration) return;
                // AbortError means a new search started - also don't show error
                if (e.name === "AbortError") return;
                
                console.error("Search error:", e);
                searchResults.innerHTML = '<div class="search-no-results">Search failed</div>';
                searchResults.classList.add("visible");
            } finally {
                // Always reset spinner for the current/active search
                // This check ensures we don't interfere with a newer search's spinner state
                if (myGen === searchGeneration) {
                    if (searchIcon) searchIcon.style.display = "block";
                    if (searchSpinner) searchSpinner.style.display = "none";
                }
            }
        }

        // Enable a layer by its URL (for search results)
        async function enableLayerByUrl(layerUrl) {
            if (!layerUrl) return;
            
            const normalizedUrl = String(layerUrl).replace(/\/+$/, "").toLowerCase();
            const normalizedKey = normalizeUrlKey(layerUrl);
            
            // Check selection layers first
            for (let i = 0; i < (selectionLayers || []).length; i++) {
                const entry = selectionLayers[i];
                const entryUrl = String(entry?.cfg?.url || "").replace(/\/+$/, "").toLowerCase();
                if (entryUrl === normalizedUrl) {
                    if (!entry.layer.visible) {
                        await enableSelectionLayer(i);
                    }
                    return;
                }
            }
            
            // Check report layers - match by config URL key or actual layer URL
            for (const [key, layerOrArray] of reportLayerViews.entries()) {
                const layers = Array.isArray(layerOrArray) ? layerOrArray : [layerOrArray];
                
                // Check if the config key matches the search URL
                const keyMatch = key === normalizedKey;
                
                // Also check if any sublayer URL matches
                const urlMatch = layers.some(lyr => {
                    const lyrUrl = String(lyr?.url || "").replace(/\/+$/, "").toLowerCase();
                    return lyrUrl === normalizedUrl;
                });
                
                if (keyMatch || urlMatch) {
                    // Enable all layers in this group
                    layers.forEach(lyr => { lyr.visible = true; });
                    
                    // Find the config index to update checkbox
                    const cfgIdx = (config.reportLayers || []).findIndex(cfg => 
                        normalizeUrlKey(cfg.url) === key
                    );
                    if (cfgIdx >= 0) {
                        const checkbox = document.getElementById(`rptlayer_${cfgIdx}`);
                        if (checkbox && !checkbox.checked) {
                            checkbox.checked = true;
                        }
                    }
                    return;
                }
            }
            
            console.warn("Could not find layer to enable:", layerUrl);
        }

        // Zoom to a feature on the map
        async function zoomToFeature(feature) {
            if (!view || !feature.geometry) {
                console.warn("Cannot zoom: no geometry", feature);
                return;
            }

            try {
                const geomJson = feature.geometry;
                const sr = geomJson.spatialReference || view.spatialReference || { wkid: 102100 };
                
                // Determine geometry type and create a proper graphic
                let graphic = null;
                let geomType = null;
                
                if (geomJson.rings && geomJson.rings.length > 0) {
                    geomType = "polygon";
                    graphic = new Graphic({
                        geometry: { type: "polygon", rings: geomJson.rings, spatialReference: sr },
                        symbol: { type: "simple-fill", color: [255, 255, 0, 0.4], outline: { color: [255, 100, 0], width: 3 } }
                    });
                } else if (geomJson.paths && geomJson.paths.length > 0) {
                    geomType = "polyline";
                    graphic = new Graphic({
                        geometry: { type: "polyline", paths: geomJson.paths, spatialReference: sr },
                        symbol: { type: "simple-line", color: [255, 255, 0], width: 6 }
                    });
                } else if (geomJson.x !== undefined && geomJson.y !== undefined) {
                    geomType = "point";
                    graphic = new Graphic({
                        geometry: { type: "point", x: geomJson.x, y: geomJson.y, spatialReference: sr },
                        symbol: { type: "simple-marker", color: [255, 255, 0, 0.8], size: 16, outline: { color: [255, 100, 0], width: 3 } }
                    });
                }

                if (!graphic) {
                    console.warn("Could not create graphic from geometry", geomJson);
                    return;
                }

                // Use the geometry for zoom - view.goTo expects geometry, not graphic
                const goToOptions = {
                    animate: true,
                    duration: 800
                };

                if (geomType === "point") {
                    // For points, zoom to a specific level centered on the point
                    await view.goTo({
                        target: graphic.geometry,
                        zoom: 14
                    }, goToOptions);
                } else {
                    // For lines and polygons, zoom to the geometry's extent
                    await view.goTo(graphic.geometry, goToOptions);
                }

                // Briefly highlight the feature
                const highlightLayer = new GraphicsLayer({ title: "Search Highlight" });
                map.add(highlightLayer);
                highlightLayer.add(graphic);

                // Remove highlight after 4 seconds
                setTimeout(() => {
                    try { map.remove(highlightLayer); } catch (e) { }
                }, 4000);

            } catch (e) {
                console.error("Zoom to feature failed:", e, feature);
            }
        }

        // Event handlers
        if (searchInput) {
            searchInput.addEventListener("input", (e) => {
                const val = e.target.value.trim();
                
                // Show/hide clear button
                if (searchClear) {
                    searchClear.style.display = val ? "flex" : "none";
                }

                // Debounce search
                clearTimeout(searchDebounceTimer);
                searchDebounceTimer = setTimeout(() => {
                    performSearch(val);
                }, 400);
            });

            searchInput.addEventListener("focus", () => {
                const val = searchInput.value.trim();
                if (val.length >= 2) {
                    searchResults.classList.add("visible");
                } else if (val.length > 0) {
                    searchResults.innerHTML = '<div class="search-hint">Type at least 2 characters to search</div>';
                    searchResults.classList.add("visible");
                }
            });

            // Close results when clicking outside
            document.addEventListener("click", (e) => {
                const widget = document.getElementById("featureSearchWidget");
                if (widget && !widget.contains(e.target)) {
                    searchResults.classList.remove("visible");
                }
            });

            // Clear button
            if (searchClear) {
                searchClear.addEventListener("click", () => {
                    searchInput.value = "";
                    searchClear.style.display = "none";
                    searchResults.classList.remove("visible");
                    searchResults.innerHTML = "";
                    // Clear any pending debounce
                    clearTimeout(searchDebounceTimer);
                    // Abort any in-progress search
                    if (searchAbortController) {
                        try { searchAbortController.abort(); } catch (e) { }
                        searchAbortController = null;
                    }
                    // Reset spinner state
                    if (searchIcon) searchIcon.style.display = "block";
                    if (searchSpinner) searchSpinner.style.display = "none";
                    // Increment generation to invalidate any pending results
                    searchGeneration++;
                });
            }

            // Escape key to close
            searchInput.addEventListener("keydown", (e) => {
                if (e.key === "Escape") {
                    searchResults.classList.remove("visible");
                    searchInput.blur();
                }
            });
        }


        setMode("select");
        setActiveTab("layers");
        setLoadingState("Ready", 100);
        setStatus("ready");

        // Brief delay so the bar visually reaches 100% before fading
        await new Promise(r => setTimeout(r, 350));
        hideLoadingOverlay();

        // Initialize analysis modal
        analysisModal.init();

        // Preload service status once (optional). Keeps Services tab fast.
        if (servicesListEl) {
            refreshServicesTab().catch(() => { });
        }

        // Background service check on startup
        checkServiceStatusBackground().catch(() => {});
    }

    init().catch((e) => {
        console.error(e);
        setStatus("failed to initialize (see console)");
    });

});