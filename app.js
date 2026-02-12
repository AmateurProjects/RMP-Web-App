/* global require */

require([
    "app/config-helpers",
    "app/map-utils",
    "app/query-engine",
    "app/final-report",
    "esri/Map",
    "esri/views/MapView",
    "esri/layers/FeatureLayer",
    "esri/layers/GraphicsLayer",
    "esri/widgets/Sketch",
    "esri/Graphic",
    "esri/geometry/geometryEngine",
    "esri/layers/TileLayer",
    "esri/layers/ImageryLayer"
], function (configHelpers, mapUtilsModule, queryEngineModule, finalReportModule, EsriMap, MapView, FeatureLayer, GraphicsLayer, Sketch, Graphic, geometryEngine, TileLayer, ImageryLayer) {

    // ── Destructure config-helpers for functions already extracted ──
    const {
        escapeHtml, normalize, plssToolLabel,
        isPlssLayerTitleOrUrl, isPlssIntersectedLayerTitle,
        isFeatureServerRoot, isMapServerRoot,
        safeFilename, formatNumber,
        fetchJson, fetchJsonWithTimeout,
        normalizePjsonUrl, normalizeUrlKey,
        pickServiceDescription, buildLayerCfgIndex, getConfiguredServices,
        setBasemapBaseLayerOpacity, isImageryBasemap,
        expandMapServerToSublayers, expandServiceToSublayers,
        expandFeatureServerToPolygonSublayers, expandFeatureServerToAllSublayers,
        flattenAttributes, toCsv, downloadText
    } = configHelpers;


    // ---------- DOM ----------
    const modeSelect = document.getElementById("modeSelect");
    // Panel minimize toggle
    const panelEl = document.getElementById("panel");
    const panelToggleBtn = document.getElementById("panelToggleBtn");
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

    // ── Shared state object for modules (properties updated by app.js) ──
    const state = {
        get config() { return config; },
        get view() { return view; },
        get map() { return map; },
        get selectionGeom() { return selectionGeom; },
        get aoiLayer() { return aoiLayer; },
        get aoiMaskLayer() { return aoiMaskLayer; },
        get aoiGraphic() { return aoiGraphic; },
        set aoiGraphic(v) { aoiGraphic = v; },
        get aoiSource() { return aoiSource; },
        get aoiSourcePlssTool() { return aoiSourcePlssTool; },
        get aoiSourceLayerTitle() { return aoiSourceLayerTitle; },
        get reportLayerViews() { return reportLayerViews; },
        get layerCfgByUrl() { return layerCfgByUrl; },
        get alwaysVisibleLayers() { return alwaysVisibleLayers; },
        get serviceStatus() { return serviceStatus; },
        get selectionLayers() { return selectionLayers; },
        get lastReportRowsByLayer() { return lastReportRowsByLayer; }
    };

    // ── Initialize map-utils module with shared state ──
    const mapUtils = mapUtilsModule.init(state);
    const {
        getPresetRenderer, getLayerGeometryType, makeRendererOpaque,
        ensureAoiOnTop, updateAoiMask, hideAoiMask, setAoiGeometry,
        ensureLayerVisibleAtScale, wireLayerUpdatingSpinner,
        waitForViewStationary, waitForLayerReadyToCapture,
        captureScreenshotWithWait, hardRefreshLayer,
        buildReportDisplayLayers
    } = mapUtils;

    // ── Initialize query-engine module with shared state ──
    const queryEngine = queryEngineModule.init(state);
    const {
        queryAllFeaturesPaged, queryAllFeaturesPagedWithGeometry,
        filterTouchingOnly, getReportGeometry, unionGeomsChunked,
        querySingleLayer, computeElevationStats,
        computeLayerCoverageStats, buildPerFeatureTable,
        getAoiKey, resetCoverageCacheForAoi, SQM_PER_ACRE,
        sampleWithoutReplacement, makeTable
    } = queryEngine;

    // ── Initialize final-report module with shared state + deps ──
    const finalReport = finalReportModule.init(state, {
        mapUtils, queryEngine, ImageryLayer, FeatureLayer, geometryEngine,
        setStatus, finalReportStatus
    });
    const {
        formatLegalDescription, generateLayerAttributeSummary,
        openHtmlInNewTab, formatDateTimeForReport, buildFinalReportHtmlDoc,
        getAoiSummaryForReport, buildDataSourcesSection,
        generateAoiMapsWithCircles, buildFinalReportHtml, viewFinalReport,
        getCachedFinalReportHtml, setCachedFinalReportHtml
    } = finalReport;

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
        ensureAoiOnTop();

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
    setCachedFinalReportHtml(null);
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
            ensureAoiOnTop();
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

    ensureAoiOnTop();
    });

});
}

    // ---------- Services tab ----------


async function checkServiceStatusBackground() {
    const items = getConfiguredServices(config);
    const timeoutMs = config?.services?.timeoutMs ?? 8000;

    // ✅ Parallel: all service pings are independent
    const results = await Promise.allSettled(
        items.map(async (it) => {
            const pjsonUrl = normalizePjsonUrl(it.url);
            try {
                await fetchJsonWithTimeout(pjsonUrl, timeoutMs);
                return { url: it.url, status: "UP" };
            } catch (e) {
                return { url: it.url, status: "DOWN" };
            }
        })
    );

    for (const r of results) {
        if (r.status === "fulfilled") {
            serviceStatus.set(r.value.url, r.value.status);
        }
    }

    // Re-render layer toggles to show status icons
    renderLayerToggles(map);
}




    async function refreshServicesTab() {
        if (!servicesListEl) return;

        const items = getConfiguredServices(config);
        if (!items.length) {
            servicesListEl.innerHTML = `<div class="small">No services configured.</div>`;
            return;
        }

        servicesListEl.innerHTML = `<div class="small">Checking services…</div>`;

        const timeoutMs = config?.services?.timeoutMs ?? 8000;

        // ✅ Parallel: all service pings are independent — runs in ~timeoutMs instead of N×timeoutMs
        const checkResults = await Promise.allSettled(
            items.map(async (it) => {
                const pjsonUrl = normalizePjsonUrl(it.url);
                let status = "DOWN";
                let desc = "";
                let errText = "";

                try {
                    const pjson = await fetchJsonWithTimeout(pjsonUrl, timeoutMs);

                    if (pjson == null || (pjson.currentVersion == null && pjson.layers == null && pjson.type == null)) {
                        throw new Error("Unexpected JSON (missing expected ArcGIS REST fields)");
                    }

                    status = "UP";
                    desc = pickServiceDescription(pjson);
                } catch (e) {
                    status = "DOWN";
                    errText = String(e?.message || e);
                }

                return { it, status, desc, errText };
            })
        );

        const cards = [];
        checkResults.forEach((result, i) => {
            if (result.status !== "fulfilled") return;
            const { it, status, desc, errText } = result.value;

                // ✅ basic “looks like ArcGIS REST” sanity

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
        });

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

        // Preserve all config properties (including imageService, renderingRule, etc.)
        byUrl.set(urlKey, { ...l, url: urlKey });
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

        // Handle ImageServer layers specially
        if (cfg.imageService === true) {
            expandedTargets.push({ 
                title: cfg.title, 
                url,
                __isImageService: true,
                __renderingRule: cfg.renderingRule || null
            });
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

    // ── Helper: process a single expanded target (returns { card, reportEntry }) ──
    async function processOneTarget(t) {
        const maxFields = (config.report && config.report.maxFieldsInTable)
            ? config.report.maxFieldsInTable : 8;

        // 1. Pre-existing expansion error
        if (t.error) {
            return {
                card: `
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(t.title)}</div>
              <div class="badge">error</div>
            </div>
            <div class="small mono">${escapeHtml(String(t.error))}</div>
          </div>`,
                reportEntry: null
            };
        }

        // 2. ImageServer layers — fetch metadata
        if (t.__isImageService) {
            const metaUrl = `${t.url}?f=json`;
            const resp = await fetch(metaUrl);
            const meta = await resp.json();

            const serviceDesc = meta.description || meta.serviceDescription || "No description available.";
            const copyright = meta.copyrightText || "";
            const serviceName = meta.name || t.title;

            return {
                card: `
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(t.title)}</div>
              <div class="badge">Image Service</div>
            </div>
            <div class="small mono">
              <a href="${escapeHtml(t.url)}" target="_blank" rel="noopener">Service URL</a>
            </div>
            <div style="margin-top:8px;">
              <div class="small"><b>Service:</b> ${escapeHtml(serviceName)}</div>
              <div class="small" style="margin-top:4px;">${escapeHtml(serviceDesc.slice(0, 200))}${serviceDesc.length > 200 ? '…' : ''}</div>
              ${copyright ? `<div class="small" style="margin-top:4px; color: var(--muted);">${escapeHtml(copyright)}</div>` : ''}
            </div>
          </div>`,
                reportEntry: {
                    title: t.title,
                    url: t.url,
                    count: 1,
                    rows: [],
                    _layer: null,
                    _exportQuery: null,
                    fullRows: null,
                    __isImageService: true,
                    __renderingRule: t.__renderingRule,
                    __serviceMeta: { name: serviceName, description: serviceDesc, copyright }
                }
            };
        }

        // 3. Pinned AOI feature (synchronous — no network call)
        if (t.__pinnedAoiFeature) {
            const f = t.__pinnedAoiFeature;
            const feats = f ? [f] : [];
            const rows = flattenAttributes(feats);
            const tableHtml = feats.length
                ? makeTable(feats, maxFields, feats.length)
                : `<div class="small">No sample rows.</div>`;

            return {
                card: `
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(t.title)}</div>
              <div class="badge">count: <b>${feats.length}</b></div>
            </div>
            <div class="small mono">
              <a href="${escapeHtml(t.url)}" target="_blank" rel="noopener">Service URL</a>
            </div>
            <div style="margin-top:8px;">
              ${tableHtml}
              ${(feats.length > 0) ? `
              <div class="row" style="margin-top:8px;">
                <button class="btn subtle" data-export="${escapeHtml(t.title)}">Export CSV</button>
              </div>` : ``}
            </div>
          </div>`,
                reportEntry: {
                    title: t.title,
                    url: t.url,
                    count: feats.length,
                    rows,
                    _layer: null,
                    _exportQuery: null,
                    fullRows: rows
                }
            };
        }

        // 4. Regular feature layer query
        const plss = isPlssLayerTitleOrUrl(t.title, t.url);
        const targetIsPlssIntersected = isPlssIntersectedLayerTitle(t.title);
        const spatialRel =
            (targetIsPlssIntersected && (aoiSourcePlssTool === "township" || aoiSourcePlssTool === "section"))
                ? "within" : "intersects";

        const r = await querySingleLayer(t.url, t.title, reportGeom, spatialRel);
        const rows = flattenAttributes(r.features);

        const reportEntry = {
            title: r.title,
            url: r.url,
            count: r.count,
            rows,
            _layer: r.layer,
            _exportQuery: r.exportQuery,
            fullRows: null
        };

        // Pre-fetch full rows for State Boundaries & Parcel (needed for Final Report)
        const isStateBoundaries = r.title && r.title.toLowerCase().includes("state boundaries");
        const isParcel = r.title && (r.title.toLowerCase().includes("parcel") || r.title.toLowerCase().includes("intersected"));

        if ((isStateBoundaries || isParcel) && r.count > 0 && r.layer && r.exportQuery) {
            try {
                const pageSize = config.report?.pageSize ?? 1000;
                const maxExport = config.report?.maxExportFeatures ?? 50000;
                const fullFeatures = await queryAllFeaturesPaged(
                    r.layer, r.exportQuery, pageSize, Math.min(maxExport, 100)
                );
                reportEntry.fullRows = flattenAttributes(fullFeatures);
            } catch (e) {
                console.warn(`Failed to pre-fetch full rows for ${r.title}:`, e);
            }
        }

        const tableHtml = (r.features && r.features.length)
            ? makeTable(r.features, maxFields, r.count)
            : `<div class="small">No sample rows.</div>`;

        return {
            card: `
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
                <button class="btn subtle" data-export="${escapeHtml(r.title)}">Export CSV</button>
              </div>` : ``}
            </div>
          </div>`,
            reportEntry
        };
    }
    // ── End of processOneTarget ──

    // ── Batched parallel queries (batches of 5) ──
    const BATCH_SIZE = 5;
    const cards = [];

    for (let bStart = 0; bStart < expandedTargets.length; bStart += BATCH_SIZE) {
        if (isReportCanceled(myOp)) {
            setStatus("canceled");
            break;
        }

        const bEnd = Math.min(bStart + BATCH_SIZE, expandedTargets.length);

        if (modal) {
            const progress = 25 + (35 * (bStart / expandedTargets.length));
            modal.setProgress(progress);
            modal.setStep(`Step 2/4: Querying layers ${bStart + 1}-${bEnd} of ${expandedTargets.length}...`);
            for (let k = bStart; k < bEnd; k++) {
                modal.addLog(`Querying: ${expandedTargets[k].title}`);
            }
        }

        const batchResults = await Promise.allSettled(
            expandedTargets.slice(bStart, bEnd).map(t => processOneTarget(t))
        );

        // Collect results in order
        for (const r of batchResults) {
            if (r.status === "fulfilled" && r.value) {
                cards.push(r.value.card);
                if (r.value.reportEntry) {
                    lastReportRowsByLayer.push(r.value.reportEntry);
                }
            } else if (r.status === "rejected") {
                cards.push(`
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">(query failed)</div>
              <div class="badge">error</div>
            </div>
            <div class="small mono">${escapeHtml(String(r.reason))}</div>
          </div>`);
            }
        }

        // Update modal stats after each batch
        if (modal) {
            const totalFeatures = lastReportRowsByLayer.reduce((sum, x) => sum + (x.count || 0), 0);
            modal.updateStats(lastReportRowsByLayer.length, totalFeatures, 0);
        }

        setStatus(`Running analysis... (queried ${bEnd}/${expandedTargets.length})`);
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
            // Include ImageServer layers even though they don't have _layer/_exportQuery
            const targets = lastReportRowsByLayer
                .filter(x => (x?.count || 0) > 0)
                .filter(x => (x?._layer && x?._exportQuery) || x?.__isImageService) // excludes pinned AOI source etc.
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
                    // Keep AOI mask layer visible for per-layer maps
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
                // Show AOI mask to lighten areas outside AOI
                updateAoiMask(true);
                ensureAoiOnTop();
            }

            function restoreVisibility() {
                visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
                hideAoiMask();
                ensureAoiOnTop();
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

                // Handle ImageServer layers differently
                if (item.__isImageService) {
                    const imgLayerOpts = {
                        url: item.url,
                        title: item.title,
                        visible: true
                    };
                    if (item.__renderingRule) {
                        imgLayerOpts.renderingRule = { functionName: item.__renderingRule };
                    }
                    const temp = new ImageryLayer(imgLayerOpts);
                    view.map.add(temp);
                    
                    try {
                        setVisibilityForScreenshot(temp);
                        await waitForLayerReadyToCapture(temp, view, { timeoutMs: 10000 });
                        await view.goTo(fixedExtent, { animate: false });
                        await waitForViewStationary(2500);
                        
                        const ss = await view.takeScreenshot({ format: "png", quality: 100, width });
                        if (!ss?.dataUrl) throw new Error("Screenshot failed");
                        const dataUrl = ss.dataUrl;

                        const meta = item.__serviceMeta || {};
                        
                        // Compute elevation statistics for the AOI
                        const elevStats = await computeElevationStats(item.url, selectionGeom);
                        
                        let elevStatsHtml = '';
                        if (elevStats) {
                            elevStatsHtml = `
                                <tr><td colspan="2" style="font-weight:600; padding-top:8px;">Elevation (AOI)</td></tr>
                                <tr><td>Min</td><td>${formatNumber(elevStats.minFt, 0)} ft</td></tr>
                                <tr><td>Max</td><td>${formatNumber(elevStats.maxFt, 0)} ft</td></tr>
                                <tr><td>Change</td><td>${formatNumber(elevStats.elevationChangeFt, 0)} ft</td></tr>
                            `;
                        }
                        
                        outCards.push(`
                          <div class="visual-output-card">
                            <div class="visual-output-title">${escapeHtml(item.title)}</div>
                            <img class="visual-output-img" src="${dataUrl}" alt="${escapeHtml(item.title)}" />
                            <div class="visual-output-meta">
                              <table>
                                <tr><td>Type</td><td>Image Service</td></tr>
                                <tr><td>Service</td><td>${escapeHtml(meta.name || item.title)}</td></tr>
                                ${elevStatsHtml}
                                ${meta.copyright ? `<tr><td>Source</td><td>${escapeHtml(meta.copyright)}</td></tr>` : ''}
                              </table>
                            </div>
                          </div>
                        `);
                    } finally {
                        try { view.map.remove(temp); } catch (e) { }
                        restoreVisibility();
                    }
                    continue;
                }

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

                    // Check for low coverage warning (single feature with <3% coverage)
                    const isSingleFeatureLowCoverage = (item.count === 1 && pctCovered < 3);
                    const lowCoverageWarningHtml = isSingleFeatureLowCoverage
                        ? `<div style="margin-top:8px; padding:6px; background-color:#fff3cd; border:1px solid #ffc107; border-radius:4px; font-size:11px;">
                            <span style="color:#856404;">⚠️ Low coverage (&lt;3%) — possible sliver or boundary artifact</span>
                           </div>`
                        : "";

                    outCards.push(`
                  <div class="visual-output-card">
                    <div class="visual-output-title">${escapeHtml(item.title)}</div>
                    <img class="visual-output-img" src="${dataUrl}" alt="AOI + ${escapeHtml(item.title)}" />
                    <div class="visual-output-meta">
                      <table>
                        <tr><td>AOI area</td><td><span class="mono">${formatNumber(aoiAcres, 2)}</span> acres</td></tr>
                        <tr><td>Intersecting features</td><td><span class="mono">${escapeHtml(String(item.count || 0))}</span></td></tr>
                        <tr><td>AOI covered by layer</td><td><span class="mono">${formatNumber(acresCovered, 2)}</span> acres</td></tr>
                        <tr><td>% AOI covered</td><td><span class="mono">${formatNumber(pctCovered, 2)}</span>%${isSingleFeatureLowCoverage ? ' <span style="color:#856404;" title="Low coverage — possible sliver">⚠️</span>' : ''}</td></tr>
                      </table>
                      ${lowCoverageWarningHtml}
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

    // buildFinalReportHtml moved to final-report.js module

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
    const featurePickerContent = document.querySelector(".feature-picker-content");
    const featurePickerHeader = document.querySelector(".feature-picker-header");

    // Drag functionality for feature picker
    let pickerDragState = { isDragging: false, startX: 0, startY: 0, offsetX: 0, offsetY: 0 };

    if (featurePickerHeader && featurePickerContent) {
        featurePickerHeader.addEventListener("mousedown", (e) => {
            if (e.target.tagName === "BUTTON") return; // Don't drag if clicking buttons
            pickerDragState.isDragging = true;
            pickerDragState.startX = e.clientX;
            pickerDragState.startY = e.clientY;
            const rect = featurePickerContent.getBoundingClientRect();
            pickerDragState.offsetX = rect.left;
            pickerDragState.offsetY = rect.top;
            featurePickerContent.style.transition = "none";
        });

        document.addEventListener("mousemove", (e) => {
            if (!pickerDragState.isDragging) return;
            const dx = e.clientX - pickerDragState.startX;
            const dy = e.clientY - pickerDragState.startY;
            const newLeft = pickerDragState.offsetX + dx;
            const newTop = pickerDragState.offsetY + dy;
            featurePickerContent.style.position = "fixed";
            featurePickerContent.style.left = newLeft + "px";
            featurePickerContent.style.top = newTop + "px";
            featurePickerContent.style.right = "auto";
            featurePickerContent.style.margin = "0";
        });

        document.addEventListener("mouseup", () => {
            pickerDragState.isDragging = false;
            featurePickerContent.style.transition = "";
        });
    }

    // State for feature picker
    let pickerFeatures = [];
    let pickerSelectedIdx = -1;
    let pickerOnSelect = null;

    function showFeaturePicker(features, onSelect) {
        if (!featurePickerModal || !featurePickerList) return;

        // Reset position for new selection
        if (featurePickerContent) {
            featurePickerContent.style.position = "";
            featurePickerContent.style.left = "";
            featurePickerContent.style.top = "";
            featurePickerContent.style.right = "";
            featurePickerContent.style.margin = "";
        }

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
        const callback = pickerOnSelect; // Save callback before hiding clears it
        hideFeaturePicker();

        if (callback) {
            callback(selectedFeature);
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
                }

                    // ✅ Fetch the “true” // Handler for when a feature is selected (single or from picker)
                async function handleFeatureSelection(graphic) {
                    const full = await getFullFeatureGeometryFromLayer(activeSelectionLayer, graphic);
                    aoiSourceFeature = full?.feature || graphic || null;
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
                }

                // If multiple features, show picker; otherwise select directly
                if (matches.length > 1) {
                    showFeaturePicker(matches, (selected) => {
                        handleFeatureSelection(selected.graphic);
                    });
                } else {
                    await handleFeatureSelection(matches[0].graphic);
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

        // Panel minimize toggle wiring
        if (panelToggleBtn && panelEl) {
            panelToggleBtn.addEventListener("click", () => {
                const isMinimized = panelEl.classList.toggle("minimized");
                panelToggleBtn.title = isMinimized ? "Expand panel" : "Minimize panel";
                panelToggleBtn.setAttribute("aria-label", isMinimized ? "Expand panel" : "Minimize panel");
            });
        }

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
                ensureAoiOnTop();
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
        ensureAoiOnTop();

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
        async function activatePlss(which, idxToEnable, { skipAutoZoom = false } = {}) {
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
                if (!skipAutoZoom) {
                    const lyr = selectionLayers[idxToEnable]?.layer;
                    await ensureLayerVisibleAtScale(lyr);
                    await waitForViewStationary(1500);
                }

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

        // No PLSS layer selected by default - user must choose
        // Clear any default button states
        setPlssToolActive(null);
        setPermitToolActive(null);
        setStatus("Select a layer or draw a polygon to define your AOI");


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