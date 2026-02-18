/* global require */

require([
    "app/config-helpers",
    "app/map-utils",
    "app/query-engine",
    "app/final-report",
    "app/visual-report",
    "app/feature-picker",
    "app/search",
    "app/upload-aoi",
    "esri/Map",
    "esri/views/MapView",
    "esri/layers/FeatureLayer",
    "esri/layers/GraphicsLayer",
    "esri/widgets/Sketch",
    "esri/Graphic",
    "esri/geometry/geometryEngine",
    "esri/layers/TileLayer",
    "esri/layers/ImageryLayer",
    "esri/identity/IdentityManager"
], function (configHelpers, mapUtilsModule, queryEngineModule, finalReportModule, visualReportModule, featurePickerModule, searchModule, uploadAoiModule, EsriMap, MapView, FeatureLayer, GraphicsLayer, Sketch, Graphic, geometryEngine, TileLayer, ImageryLayer, esriId) {

    // ── Suppress ArcGIS Online sign-in popup ──
    // All services used by this app are publicly shared; prevent the
    // IdentityManager from showing a login dialog on 401/403 responses.
    esriId.useSignInPage = false;
    esriId.on("credential-create", function (evt) {
        if (evt.credential) evt.credential.destroy();
    });

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
    // PERF-TEST: Advanced-mode DOM elements — HTML is commented out, these will be null
    const modeSelect = document.getElementById("modeSelect");             // null (Advanced panel commented out)
    // Panel minimize toggle
    const panelEl = document.getElementById("panel");
    const panelToggleBtn = document.getElementById("panelToggleBtn");
    // PLSS selection tools (Advanced panel — now null)
    const plssTownshipBtn = document.getElementById("plssTownshipBtn");   // null
    const plssSectionBtn = document.getElementById("plssSectionBtn");     // null
    const plssIntersectedBtn = document.getElementById("plssIntersectedBtn"); // null
    // Selection layer group selector (Advanced panel — now null)
    const selectionGroupSelect = document.getElementById("selectionGroupSelect"); // null
    const plssSelectGroup = document.getElementById("plssSelectGroup");   // null
    const permitSelectGroup = document.getElementById("permitSelectGroup"); // null
    // Permit layer selection buttons (Advanced panel — now null)
    const grazingAllotmentBtn = document.getElementById("grazingAllotmentBtn"); // null
    const grazingPastureBtn = document.getElementById("grazingPastureBtn");     // null
    const oilGasLeaseBtn = document.getElementById("oilGasLeaseBtn");           // null

    const selectModeControls = document.getElementById("selectModeControls"); // null
    const drawModeControls = document.getElementById("drawModeControls");     // null

    const drawBtn = document.getElementById("drawBtn");       // null
    const stopDrawBtn = document.getElementById("stopDrawBtn"); // null
    const runBtn = document.getElementById("runBtn");           // null
    const clearBtn = document.getElementById("clearBtn");       // null
    const exportAllBtn = document.getElementById("exportAllBtn"); // null

    const viewBlockerEl = document.getElementById("viewBlocker");

    // Layer Manager floating window
    const advLayersBtn = document.getElementById("advLayersBtn");
    const layerManagerWindow = document.getElementById("layerManagerWindow");
    const layerMgrCloseBtn = document.getElementById("layerMgrCloseBtn");
    const layerMgrHeader = document.getElementById("layerMgrHeader");
    const layerMgrSelectionList = document.getElementById("layerMgrSelectionList");
    const layerMgrReportList = document.getElementById("layerMgrReportList");

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

    // Visual report DOM managed by visual-report.js module

    // Final report DOM
    const viewReportBtn = document.getElementById("viewReportBtn");
    const copyReportLinkBtn = document.getElementById("copyReportLinkBtn");
    const finalReportStatus = document.getElementById("finalReportStatus");

    const servicesListEl = document.getElementById("servicesList");
    const refreshServicesBtn = document.getElementById("refreshServicesBtn");

    // ── Permitting Mode DOM ──
    const permitModePanel = document.getElementById("permitModePanel");
    const advancedModePanel = document.getElementById("advancedModePanel"); // null (commented out)
    const permitModeBtn = document.getElementById("permitModeBtn");
    const advancedModeBtn = null; // document.getElementById("advancedModeBtn"); // PERF-TEST: commented out
    const wizardStep1 = document.getElementById("wizardStep1");
    const wizardStep2 = document.getElementById("wizardStep2");
    const wizardStep3 = document.getElementById("wizardStep3");
    const wizScreenBtn = document.getElementById("wizScreenBtn");
    const wizBackToStep1 = document.getElementById("wizBackToStep1");
    const wizNewScreening = document.getElementById("wizNewScreening");
    const wizFullReport = document.getElementById("wizFullReport");
    const wizCopyReportLink = document.getElementById("wizCopyReportLink");
    const wizExportAll = document.getElementById("wizExportAll");
    const wizStopDrawBtn = document.getElementById("wizStopDrawBtn");
    const wizTownshipBtn = document.getElementById("wizTownshipBtn");
    const wizSectionBtn = document.getElementById("wizSectionBtn");
    const wizParcelBtn = document.getElementById("wizParcelBtn");
    const wizPermitList = document.getElementById("wizPermitList");
    const wizLocationInput = document.getElementById("wizLocationInput");
    const wizLocationResults = document.getElementById("wizLocationResults");
    const tierLayerCountEl = document.getElementById("tierLayerCount");

    /* ── Tier selection helper ── */
    function getSelectedTier() {
      const sel = document.querySelector('input[name="analysisTier"]:checked');
      return sel ? parseInt(sel.value, 10) : 1;
    }
    function updateTierLayerCount() {
      const tier = getSelectedTier();
      const count = (config.reportLayers || []).filter(l => (l.tier || 1) <= tier).length;
      if (tierLayerCountEl) tierLayerCountEl.textContent = count + " layer" + (count !== 1 ? "s" : "") + " will be screened";
    }
    // wire up radio change events
    document.querySelectorAll('input[name="analysisTier"]').forEach(r => {
      r.addEventListener("change", updateTierLayerCount);
    });

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
                    if (currentAppMode === "permit") {
                        goToWizardStep(3);
                    } else {
                        setActiveTab("report"); // Jump to Tables tab
                    }
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
        
        showSuccess(layersQueried, featuresFound, mapsGenerated, elapsedMs) {
            if (!this.success) return;
            
            // Update success summary
            const successLayersQueried = document.getElementById("successLayersQueried");
            const successFeaturesFound = document.getElementById("successFeaturesFound");
            const successMapsGenerated = document.getElementById("successMapsGenerated");
            const successDuration = document.getElementById("successDuration");
            
            if (successLayersQueried) successLayersQueried.textContent = layersQueried;
            if (successFeaturesFound) successFeaturesFound.textContent = featuresFound;
            if (successMapsGenerated) successMapsGenerated.textContent = mapsGenerated;
            
            if (successDuration && elapsedMs != null) {
                const totalSec = Math.round(elapsedMs / 1000);
                const min = Math.floor(totalSec / 60);
                const sec = totalSec % 60;
                successDuration.textContent = min > 0
                    ? `Completed in ${min}m ${sec}s`
                    : `Completed in ${sec}s`;
            }
            
            this.success.classList.remove("hidden");
            
        }
    };


    // ---------- State ----------
    let config = null;

    let view = null;
    let selectionGeom = null;
    let aoiSource = null;            // "draw" | "select" | "upload"
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
    let currentDrawToolType = "polygon";

    let lastReportRowsByLayer = []; // for export-all
    let reportLayerViews = new Map();

    // ── Permitting Mode State ──
    let currentAppMode = "permit"; // "permit" | "advanced"
    let currentWizardStep = 1;
    let currentAoiMethod = null; // "search" | "permit" | "select" | "draw" | "upload"
    let currentInteractionMode = "select"; // PERF-TEST: tracks draw/select without modeSelect DOM

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
        querySingleLayer, querySingleLayerChunked, computeElevationStats,
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
        getCachedFinalReportHtml, setCachedFinalReportHtml,
        getLastReportId, getReportShareUrl, loadReportFromDb, cleanupExpiredReports
    } = finalReport;

    // ── Initialize visual-report module with shared state + deps ──
    const visualReport = visualReportModule.init(state, {
        mapUtils, queryEngine, ImageryLayer, FeatureLayer, geometryEngine,
        isReportCanceled
    });
    const {
        setVisualStatus, clearVisualReport, renderVisualSummary, generateVisualReportData
    } = visualReport;

    // ── Initialize feature-picker module with shared state + deps ──
    featurePickerModule.init(state, { GraphicsLayer, Graphic });
    const { showFeaturePicker, hideFeaturePicker } = featurePickerModule;

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

    // ──────────────────────────────────────────────────────────────
    // PERMITTING MODE — Mode switching, wizard, AOI methods, buckets
    // ──────────────────────────────────────────────────────────────

    const PERMIT_BUCKETS = {
        "land-status": {
            label: "Land Status & Authority", icon: "🏛️",
            description: "Federal land ownership, administrative boundaries, tribal lands, and jurisdictional authority over BLM-managed lands.",
            patterns: [/federal lands/i, /admin.*unit/i, /state boundar/i, /usfws.*region/i, /aoi source/i,
                        /bia.*aian/i, /indian/i, /alaska.*native/i, /tribal/i, /surface.*ownership/i,
                        /land use planning bound/i]
        },
        "land-use": {
            label: "Land Use Plans & Allocations", icon: "📑",
            description: "Resource Management Plans, timber and mineral allocations that may govern permitted activities in this area.",
            patterns: [/land use plan/i, /revision.*development/i, /timber/i, /locatable.*mineral/i,
                        /taylor grazing/i, /tga/i]
        },
        "special": {
            label: "Special Designations", icon: "⭐",
            description: "ACECs, wilderness, conservation lands, wild & scenic rivers, roadless areas, and other designations that may restrict or condition activities.",
            patterns: [/acec/i, /critical environmental/i, /nlcs/i, /conservation area/i, /national monument/i,
                        /wilderness/i, /wsa/i, /recreation site/i, /lwcf/i, /conservation fund/i, /visual resource/i,
                        /wild.*scenic.*river/i, /roadless/i, /national forest bound/i, /national wildlife refuge/i, /nwr/i]
        },
        "environmental": {
            label: "Environmental & ESA", icon: "🌿",
            description: "Threatened and endangered species habitat, wetlands, hydrology, wildlife corridors, flood hazards, and fire history.",
            patterns: [/critical habitat/i, /ungulate/i, /migration/i, /wild horse/i, /burro/i, /elevation/i, /fire perim/i,
                        /wetland/i, /nwi/i, /riparian/i, /nhd/i, /hydrography/i, /watershed/i, /wbd/i,
                        /flood/i, /nfhl/i, /fema/i, /sagebrush/i, /fiat/i, /danl/i, /disturbance/i,
                        /at.risk.*species/i, /t\&e/i, /threatened/i]
        },
        "authorizations": {
            label: "Existing Authorizations", icon: "📝",
            description: "Active permits, leases, rights-of-way, mining claims, and other authorizations that currently overlap your project area.",
            patterns: [/grazing allot/i, /grazing pasture/i, /oil.*gas/i, /mlrs.*row/i, /lua.*row/i, /eplanning/i, /plss.*parcel/i,
                        /mining claim/i, /lua.*lease/i, /lua.*permit/i, /lua.*easem/i, /geothermal/i, /coal case/i,
                        /oil shale/i, /non.energy/i, /mineral material/i, /locatable notice/i, /locatable plan/i,
                        /participating area/i, /agreement/i, /gtlf/i, /road.*trail/i]
        }
    };

    function categorizeIntoBuckets(reportItems) {
        const buckets = {};
        for (const key of Object.keys(PERMIT_BUCKETS)) buckets[key] = [];
        buckets["uncategorized"] = [];
        for (const item of (reportItems || [])) {
            const title = (item.title || "");
            let placed = false;
            for (const [bk, bd] of Object.entries(PERMIT_BUCKETS)) {
                if (bd.patterns.some(p => p.test(title))) { buckets[bk].push(item); placed = true; break; }
            }
            if (!placed) buckets["uncategorized"].push(item);
        }
        return buckets;
    }

    // PERF-TEST: setAppMode simplified — always "permit", Advanced mode commented out
    function setAppMode(mode) {
        currentAppMode = "permit"; // force permit mode
        if (permitModePanel) permitModePanel.classList.remove("hidden");
        // Advanced panel is commented out in HTML, no-op:
        // if (advancedModePanel) advancedModePanel.classList.toggle("hidden", mode !== "advanced");
        if (permitModeBtn) permitModeBtn.classList.add("active");
        // if (advancedModeBtn) advancedModeBtn.classList.toggle("active", mode === "advanced");
        if (view) requestAnimationFrame(() => { try { view.resize(); } catch (e) {} });
    }

    function goToWizardStep(step) {
        currentWizardStep = step;
        document.querySelectorAll("#wizardSteps .wizard-step").forEach(el => {
            const s = parseInt(el.dataset.step, 10);
            el.classList.toggle("active", s === step);
            el.classList.toggle("completed", s < step);
        });
        document.querySelectorAll("#wizardSteps .wizard-connector").forEach((conn, idx) => {
            conn.style.background = (idx < step - 1) ? "var(--blm-green)" : "var(--border-light)";
        });
        [wizardStep1, wizardStep2, wizardStep3].forEach((panel, idx) => {
            if (panel) panel.classList.toggle("active", idx + 1 === step);
        });
    }

    function showAoiMethod(method) {
        // Reset any existing AOI process before switching methods
        if (sketch) sketch.cancel();
        currentInteractionMode = "select";
        selectionGeom = null;
        aoiSource = null;
        aoiSourceLayerTitle = null;
        aoiSourceLayerUrl = null;
        aoiSourceObjectId = null;
        aoiSourceObjectIdField = null;
        aoiSourceFeature = null;
        if (aoiLayer) aoiLayer.removeAll();
        aoiGraphic = null;
        if (runBtn) runBtn.disabled = true;
        if (wizFullReport) wizFullReport.disabled = true;
        if (wizExportAll) wizExportAll.disabled = true;
        const dbp = document.getElementById("drawBufferPanel");
        if (dbp) dbp.classList.add("hidden");
        setStatus("");

        currentAoiMethod = method;
        const methodsEl = document.getElementById("aoiMethods");
        if (methodsEl) methodsEl.classList.add("hidden");
        const panels = { search: "aoiMethodSearch", permit: "aoiMethodPermit", select: "aoiMethodSelect", draw: "aoiMethodDraw", upload: "aoiMethodUpload" };
        for (const [key, id] of Object.entries(panels)) {
            const p = document.getElementById(id);
            if (p) p.classList.toggle("hidden", key !== method);
        }
        if (method === "draw") {
            if (modeSelect) modeSelect.value = "draw";
            setMode("draw");
        } else if (method === "select" || method === "permit") {
            if (modeSelect) modeSelect.value = "select";
            setMode("select");
        }

        // Scroll the opened panel into view so the user sees the options
        const targetPanel = document.getElementById(panels[method]);
        if (targetPanel) {
            requestAnimationFrame(() => {
                targetPanel.scrollIntoView({ behavior: "smooth", block: "start" });
            });
        }
    }

    function hideAoiMethodPanels() {
        ["aoiMethodSearch", "aoiMethodPermit", "aoiMethodSelect", "aoiMethodDraw", "aoiMethodUpload"].forEach(id => {
            const p = document.getElementById(id);
            if (p) p.classList.add("hidden");
        });
        const methodsEl = document.getElementById("aoiMethods");
        if (methodsEl) methodsEl.classList.remove("hidden");
        currentAoiMethod = null;

        // Stop any active sketch and restore click-to-select
        if (sketch) sketch.cancel();
        currentInteractionMode = "select";
    }

    function setWizPlssActive(which) {
        const set = (btn, on) => { if (btn) btn.setAttribute("aria-pressed", on ? "true" : "false"); };
        set(wizTownshipBtn, which === "township");
        set(wizSectionBtn, which === "section");
        set(wizParcelBtn, which === "intersected");
    }

    // Track whether the current AOI exceeds the large-AOI threshold
    let aoiIsLarge = false;
    let aoiCurrentAcres = 0;

    function populateAoiConfirmation() {
        aoiIsLarge = false;
        aoiCurrentAcres = 0;
        const warningEl = document.getElementById("aoiSizeWarning");

        if (selectionGeom) {
            try {
                const areaSqm = geometryEngine.geodesicArea(selectionGeom, "square-meters");
                const acres = Math.abs(areaSqm) / 4046.8564224;
                aoiCurrentAcres = acres;
                const el = document.getElementById("wizAoiAcres");
                if (el) el.textContent = formatNumber(acres, 2) + " acres";

                // Check against warning threshold
                const warningThreshold = config.report?.aoiWarningAcres ?? 250000;
                if (acres > warningThreshold) {
                    aoiIsLarge = true;
                    const sqMiles = (acres / 640).toFixed(0);
                    const estMinutes = Math.max(2, Math.round(acres / 50000));
                    if (warningEl) {
                        warningEl.innerHTML = `
                            <span class="warn-icon">\u26A0\uFE0F</span>
                            <div class="warn-body">
                                <strong>Large Area of Interest Detected</strong>
                                Your AOI is approximately <b>${formatNumber(acres, 0)} acres</b> (~${sqMiles} sq mi).
                                Analysis over such large areas may take <b>${estMinutes}+ minutes</b> and some service queries may time out.
                                The analysis will automatically use spatial chunking to improve reliability.
                                <div class="warn-detail">For faster results, consider selecting a smaller area or using Tier 1 (Essential) analysis depth.</div>
                            </div>`;
                        warningEl.classList.remove("hidden");
                    }
                } else {
                    if (warningEl) warningEl.classList.add("hidden");
                }
            } catch (e) {
                const el = document.getElementById("wizAoiAcres");
                if (el) el.textContent = "(unable to compute)";
                if (warningEl) warningEl.classList.add("hidden");
            }
        } else {
            if (warningEl) warningEl.classList.add("hidden");
        }
        const sourceEl = document.getElementById("wizAoiSource");
        if (sourceEl) {
            if (aoiSource === "draw") sourceEl.textContent = "Custom drawn polygon";
            else if (aoiSource === "upload") sourceEl.textContent = "Uploaded file: " + (aoiSourceLayerTitle || "unknown");
            else if (aoiSourceLayerTitle) sourceEl.textContent = aoiSourceLayerTitle;
            else sourceEl.textContent = "Map selection";
        }
        // Update tier layer count when step 2 is shown
        updateTierLayerCount();
    }

    let wizLocationDebounce = null;
    function performLocationSearch(query) {
        if (!query || query.length < 3) {
            if (wizLocationResults) { wizLocationResults.innerHTML = ""; wizLocationResults.classList.add("hidden"); }
            return;
        }
        clearTimeout(wizLocationDebounce);
        wizLocationDebounce = setTimeout(async () => {
            try {
                const url = "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/suggest?text=" + encodeURIComponent(query) + "&maxSuggestions=5&f=json";
                const data = await fetchJson(url);
                const suggestions = (data && data.suggestions) ? data.suggestions : [];
                if (!suggestions.length) {
                    if (wizLocationResults) {
                        wizLocationResults.innerHTML = '<div class="wiz-location-item" style="color:var(--text-muted);">No results found</div>';
                        wizLocationResults.classList.remove("hidden");
                    }
                    return;
                }
                if (wizLocationResults) {
                    wizLocationResults.innerHTML = suggestions.map(s =>
                        '<div class="wiz-location-item" data-magic-key="' + escapeHtml(s.magicKey || "") + '" data-text="' + escapeHtml(s.text || "") + '">' + escapeHtml(s.text) + '</div>'
                    ).join("");
                    wizLocationResults.classList.remove("hidden");
                    wizLocationResults.querySelectorAll(".wiz-location-item").forEach(item => {
                        item.addEventListener("click", async () => {
                            const mk = item.getAttribute("data-magic-key");
                            const txt = item.getAttribute("data-text");
                            if (wizLocationInput) wizLocationInput.value = txt;
                            wizLocationResults.classList.add("hidden");
                            try {
                                const gUrl = "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates?SingleLine=" + encodeURIComponent(txt) + "&magicKey=" + encodeURIComponent(mk) + "&outSR=4326&f=json";
                                const gData = await fetchJson(gUrl);
                                const c = gData && gData.candidates && gData.candidates[0];
                                if (c && c.location) {
                                    await view.goTo({ center: [c.location.x, c.location.y], zoom: 12 }, { animate: true, duration: 800 });
                                    setStatus("Zoomed to location \u2014 now select a boundary or draw your area");
                                    showAoiMethod("select");
                                }
                            } catch (e) {
                                console.warn("Geocode failed:", e);
                                setStatus("Location search failed \u2014 try another method");
                            }
                        });
                    });
                }
            } catch (e) { console.warn("Location suggest failed:", e); }
        }, 350);
    }

    function populatePermitBuckets() {
        const buckets = categorizeIntoBuckets(lastReportRowsByLayer);

        // Update tab badges
        document.querySelectorAll("#permitBucketTabs .permit-tab").forEach(tab => {
            const bKey = tab.dataset.bucket;
            if (!bKey || bKey === "overview" || bKey === "all-data") return;
            const items = buckets[bKey] || [];
            const hitCount = items.reduce((s, it) => s + (it.count || 0), 0);
            const existing = tab.querySelector(".bucket-badge");
            if (existing) existing.remove();
            const badge = document.createElement("span");
            badge.className = "bucket-badge" + (hitCount === 0 ? " zero" : "");
            badge.textContent = String(hitCount);
            tab.appendChild(badge);
        });

        // Overview bucket — compact dashboard
        const overviewEl = document.getElementById("bucketOverview");
        if (overviewEl) {
            const tl = lastReportRowsByLayer.length;
            const th = lastReportRowsByLayer.reduce((s, x) => s + (x.count || 0), 0);
            const lwh = lastReportRowsByLayer.filter(x => (x.count || 0) > 0).length;
            let oh = '<div class="overview-dash">';

            // ── Summary stats row ──
            oh += '<div class="permit-summary-grid">';
            oh += '<div class="permit-summary-stat"><div class="permit-summary-stat-value">' + tl + '</div><div class="permit-summary-stat-label">Datasets Checked</div></div>';
            oh += '<div class="permit-summary-stat"><div class="permit-summary-stat-value">' + lwh + '</div><div class="permit-summary-stat-label">With Findings</div></div>';
            oh += '<div class="permit-summary-stat"><div class="permit-summary-stat-value">' + th + '</div><div class="permit-summary-stat-label">Total Features</div></div>';
            oh += '</div>';

            // ── Category status rows (traffic-light dashboard) ──
            oh += '<div class="overview-category-grid">';
            for (const [bk, bd] of Object.entries(PERMIT_BUCKETS)) {
                const items = buckets[bk] || [];
                const hc = items.reduce((s, it) => s + (it.count || 0), 0);
                const lhc = items.filter(it => (it.count || 0) > 0).length;
                const statusClass = hc > 0 ? 'findings' : 'clear';
                const statusLabel = hc > 0 ? (hc + ' feature' + (hc !== 1 ? 's' : '') + ' in ' + lhc + ' layer' + (lhc !== 1 ? 's' : '')) : 'Clear';
                oh += '<button class="overview-cat-row ' + statusClass + '" type="button" data-goto-bucket="' + bk + '">';
                oh += '<span class="overview-cat-indicator"></span>';
                oh += '<span class="overview-cat-icon">' + bd.icon + '</span>';
                oh += '<span class="overview-cat-label">' + escapeHtml(bd.label) + '</span>';
                oh += '<span class="overview-cat-status">' + statusLabel + '</span>';
                oh += '<span class="overview-cat-arrow">›</span>';
                oh += '</button>';
            }
            oh += '</div>';

            // ── Disclaimer ──
            oh += '<div class="overview-disclaimer"><strong>Important:</strong> These results are based on available GIS data and may not reflect all conditions on the ground. Additional site-specific review may be required during the permitting process. Contact your local BLM field office for authoritative guidance.</div>';
            oh += '</div>';
            overviewEl.innerHTML = oh;

            // Wire up category row clicks to switch to that bucket tab
            overviewEl.querySelectorAll('.overview-cat-row[data-goto-bucket]').forEach(btn => {
                btn.addEventListener('click', () => setActiveBucket(btn.dataset.gotoBucket));
            });
        }

        // Summary in header card
        const summaryEl = document.getElementById("permitResultsSummary");
        if (summaryEl) {
            const t2 = lastReportRowsByLayer.reduce((s, x) => s + (x.count || 0), 0);
            const l2 = lastReportRowsByLayer.filter(x => (x.count || 0) > 0).length;
            summaryEl.innerHTML = '<div class="small"><strong>' + l2 + '</strong> of <strong>' + lastReportRowsByLayer.length + '</strong> datasets have findings &mdash; <strong>' + t2 + '</strong> total features intersect your project area.</div>';
        }

        // Individual bucket panels
        for (const [bk, bd] of Object.entries(PERMIT_BUCKETS)) {
            const pid = "bucket" + bk.split("-").map(w => w.charAt(0).toUpperCase() + w.slice(1)).join("");
            const pe = document.getElementById(pid);
            if (!pe) continue;
            const items = buckets[bk] || [];
            let h = '<div class="bucket-card"><div class="bucket-card-head"><div class="bucket-card-title">' + bd.icon + ' ' + escapeHtml(bd.label) + '</div>';
            const tc = items.reduce((s, it) => s + (it.count || 0), 0);
            h += '<div class="bucket-card-count' + (tc === 0 ? ' zero' : '') + '">' + tc + ' feature' + (tc !== 1 ? 's' : '') + '</div></div>';
            h += '<div class="bucket-card-desc">' + escapeHtml(bd.description) + '</div><ul class="bucket-layer-list">';
            for (const it of items) {
                const c = it.count || 0;
                h += '<li class="bucket-layer-item"><span class="bucket-layer-name">' + escapeHtml(it.title) + '</span>';
                h += '<span class="bucket-layer-count' + (c > 0 ? ' has-hits' : '') + '">' + c + ' feature' + (c !== 1 ? 's' : '') + '</span></li>';
            }
            if (!items.length) h += '<li class="bucket-layer-item" style="color:var(--text-muted);font-style:italic;">No layers in this category</li>';
            h += '</ul>';
            if (tc === 0) h += '<div class="hint" style="margin-top:8px;">No intersecting features were found. This does not guarantee the absence of relevant considerations not captured in available GIS data.</div>';
            h += '</div>';
            pe.innerHTML = h;
        }

        // All Data bucket
        const adEl = document.getElementById("bucketAllData");
        if (adEl) {
            const adT = lastReportRowsByLayer.reduce((s, x) => s + (x.count || 0), 0);
            let ad = '<div class="bucket-card"><div class="bucket-card-head"><div class="bucket-card-title">📊 All Queried Layers</div>';
            ad += '<div class="bucket-card-count">' + adT + ' total</div></div><ul class="bucket-layer-list">';
            const sorted = lastReportRowsByLayer.slice().sort((a, b) => (b.count || 0) - (a.count || 0));
            for (const it of sorted) {
                const c = it.count || 0;
                ad += '<li class="bucket-layer-item"><span class="bucket-layer-name">' + escapeHtml(it.title) + '</span>';
                ad += '<span class="bucket-layer-count' + (c > 0 ? ' has-hits' : ' zero') + '">' + c + '</span></li>';
            }
            ad += '</ul></div>';
            adEl.innerHTML = ad;
        }
    }

    function setActiveBucket(bucketKey) {
        document.querySelectorAll("#permitBucketTabs .permit-tab").forEach(tab => {
            tab.classList.toggle("active", tab.dataset.bucket === bucketKey);
        });
        const pm = { "overview": "bucketOverview", "land-status": "bucketLandStatus", "land-use": "bucketLandUse", "special": "bucketSpecial", "environmental": "bucketEnvironmental", "authorizations": "bucketAuthorizations", "all-data": "bucketAllData" };
        for (const [key, id] of Object.entries(pm)) {
            const p = document.getElementById(id);
            if (p) p.classList.toggle("active", key === bucketKey);
        }
    }
    // ── END PERMITTING MODE FUNCTIONS ──

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
        // Update aria-pressed on all permit-item buttons in the custom list
        document.querySelectorAll("#wizPermitList .permit-item").forEach(btn => {
            const val = btn.dataset.permit;
            btn.setAttribute("aria-pressed", val === which ? "true" : "false");
        });
        // Also update legacy advanced-mode buttons (may be null)
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
    // Cancel any active sketch drawing session
    if (sketch) sketch.cancel();

    selectionGeom = null;
    aoiSourceObjectId = null;
    aoiSourceObjectIdField = null;
    aoiSource = null;
    aoiSourceLayerTitle = null;
    aoiSourceLayerUrl = null;
    aoiSourceFeature = null;
    aoiSourcePlssTool = null;
    
    // Clear results
    if (resultsEl) resultsEl.innerHTML = "";
    if (exportAllBtn) exportAllBtn.disabled = true;
    lastReportRowsByLayer = [];
    
    // Clear map outputs
    clearVisualReport();
    
    // Clear final report
    setCachedFinalReportHtml(null);
    if (viewReportBtn) viewReportBtn.disabled = true;
    if (copyReportLinkBtn) copyReportLinkBtn.disabled = true;

    if (aoiLayer) aoiLayer.removeAll();
    aoiGraphic = null;

    if (runBtn) runBtn.disabled = true;
    setStatus("cleared");
    resetCoverageCacheForAoi(null);
    setBusy(false);

    // Reset wizard-specific UI state
    if (wizFullReport) wizFullReport.disabled = true;
    if (wizCopyReportLink) wizCopyReportLink.disabled = true;
    if (wizExportAll) wizExportAll.disabled = true;
    setWizPlssActive(null);

    // Clear wizard location search
    clearTimeout(wizLocationDebounce);
    wizLocationDebounce = null;
    if (wizLocationInput) wizLocationInput.value = "";
    if (wizLocationResults) { wizLocationResults.innerHTML = ""; wizLocationResults.classList.add("hidden"); }

    // Clear permit bucket DOM (step 3 results)
    const bucketPanelIds = ["bucketOverview", "bucketLandStatus", "bucketLandUse", "bucketSpecial", "bucketEnvironmental", "bucketAuthorizations", "bucketAllData"];
    bucketPanelIds.forEach(id => { const el = document.getElementById(id); if (el) el.innerHTML = ""; });
    const summaryEl = document.getElementById("permitResultsSummary");
    if (summaryEl) summaryEl.innerHTML = "";
    document.querySelectorAll("#permitBucketTabs .permit-tab .bucket-badge").forEach(b => b.remove());

    // Clear upload status
    const uploadStatusEl = document.getElementById("uploadStatus");
    if (uploadStatusEl) uploadStatusEl.classList.add("hidden");
}

    function setGeometryFromSelection(geom) {
        selectionGeom = geom || null;
        if (runBtn) runBtn.disabled = !selectionGeom;

        // Permitting mode: advance to Step 2 when AOI is defined
        if (selectionGeom && currentAppMode === "permit" && currentWizardStep === 1) {
            populateAoiConfirmation();
            goToWizardStep(2);
        }
    }

    function setMode(mode) {
        currentInteractionMode = mode; // PERF-TEST: track mode locally
        function startDrawingNow() {
            if (!sketch) return;
            // Cancel any prior sketch session and start drawing
            sketch.cancel();
            // currentDrawToolType is set by the draw tool selector buttons (default: "polygon")
            const toolType = (typeof currentDrawToolType !== "undefined" && currentDrawToolType) || "polygon";
            sketch.create(toolType);
            setStatus("drawing " + toolType + "…");
        }

        if (mode === "select") {
            if (selectModeControls) selectModeControls.classList.remove("hidden");
            if (drawModeControls) drawModeControls.classList.add("hidden");
            // stop sketch if running
            if (sketch) sketch.cancel();
            setStatus("select mode: click a polygon");
        } else {
            if (selectModeControls) selectModeControls.classList.add("hidden");
            if (drawModeControls) drawModeControls.classList.remove("hidden");
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

    // ---------- Layer Manager (floating window) ----------

    function renderLayerManagerToggles() {
        if (!layerMgrSelectionList || !layerMgrReportList) return;

        // Selection layers
        layerMgrSelectionList.innerHTML = (selectionLayers || []).map((e, i) => {
            const checked = e.layer.visible ? "checked" : "";
            const status = serviceStatus.get(e.cfg.url);
            const statusIcon = (status === "DOWN")
                ? `<span class="status-warning" title="Service is DOWN">⚠️</span>`
                : "";
            return `
                <div class="toggle-row">
                    <input type="checkbox" id="lm_sellayer_${i}" ${checked} />
                    <span class="layer-swatch layer-swatch-selection" aria-hidden="true"></span>
                    <label class="toggle-name" for="lm_sellayer_${i}">${statusIcon}${escapeHtml(e.cfg.title)}</label>
                    <span id="lm_sellayer_spin_${i}" class="layer-spinner hidden" aria-label="loading"></span>
                </div>`;
        }).join("");

        // Wire selection layer checkboxes
        (selectionLayers || []).forEach((e, i) => {
            const cb = document.getElementById(`lm_sellayer_${i}`);
            if (!cb) return;
            cb.addEventListener("change", async () => {
                const spin = document.getElementById(`lm_sellayer_spin_${i}`);
                if (cb.checked) {
                    if (spin) spin.classList.remove("hidden");
                    const isOnMapNow = map.layers.includes(e.layer);
                    if (!isOnMapNow) map.add(e.layer);
                    e.layer.visible = true;
                    ensureAoiOnTop();
                    await hardRefreshLayer(e.layer);
                    clearSpinnerWatch(e.layer);
                    wireLayerUpdatingSpinner(e.layer, spin).then(h => setSpinnerWatch(e.layer, h));
                    if (!activeSelectionLayer) await setActiveSelectionLayerByIndex(i);
                } else {
                    if (spin) spin.classList.add("hidden");
                    clearSpinnerWatch(e.layer);
                    e.layer.visible = false;
                    if (activeSelectionLayer === e.layer) {
                        activeSelectionLayer = null;
                        activeSelectionLayerView = null;
                    }
                }
                // Sync the old panel checkbox if it exists
                const oldCb = document.getElementById(`sellayer_${i}`);
                if (oldCb) oldCb.checked = cb.checked;
            });
        });

        // Report layers
        layerMgrReportList.innerHTML = (config.reportLayers || []).map((l, i) => {
            const key = normalizeUrlKey(l.url);
            const existing = reportLayerViews.get(key);
            const isAlwaysOn = l.alwaysVisible === true;
            const isVisible = isAlwaysOn || (existing
                ? (Array.isArray(existing) ? existing.some(x => x.visible) : existing.visible)
                : false);
            const checked = isVisible ? "checked" : "";
            const status = serviceStatus.get(l.url);
            const statusIcon = (status === "DOWN")
                ? `<span class="status-warning" title="Service is DOWN">⚠️</span>`
                : "";
            const tierBadge = l.tier ? `<span style="font-size:10px;opacity:.55;margin-left:4px;">T${l.tier}</span>` : "";
            return `
                <div class="toggle-row">
                    <input type="checkbox" id="lm_rptlayer_${i}" ${checked} />
                    <span class="layer-swatch layer-swatch-report" aria-hidden="true"></span>
                    <label class="toggle-name" for="lm_rptlayer_${i}">${statusIcon}${escapeHtml(l.title)}${tierBadge}</label>
                    <span id="lm_rptlayer_spin_${i}" class="layer-spinner hidden" aria-label="loading"></span>
                </div>`;
        }).join("");

        // Wire report layer checkboxes
        (config.reportLayers || []).forEach((l, i) => {
            const cb = document.getElementById(`lm_rptlayer_${i}`);
            if (!cb) return;
            cb.addEventListener("change", async () => {
                const spin = document.getElementById(`lm_rptlayer_spin_${i}`);
                const key = normalizeUrlKey(l.url);
                const existing = reportLayerViews.get(key);
                const setVisible = (obj, vis) => {
                    if (!obj) return;
                    if (Array.isArray(obj)) obj.forEach(x => { try { x.visible = vis; } catch (e) {} });
                    else { try { obj.visible = vis; } catch (e) {} }
                };
                if (cb.checked) {
                    if (spin) spin.classList.remove("hidden");
                    setVisible(existing, true);
                    if (Array.isArray(existing)) {
                        for (const lyr of existing) {
                            clearSpinnerWatch(lyr);
                            wireLayerUpdatingSpinner(lyr, spin).then(h => setSpinnerWatch(lyr, h));
                        }
                    } else if (existing) {
                        clearSpinnerWatch(existing);
                        wireLayerUpdatingSpinner(existing, spin).then(h => setSpinnerWatch(existing, h));
                    }
                } else {
                    if (spin) spin.classList.add("hidden");
                    if (Array.isArray(existing)) existing.forEach(lyr => clearSpinnerWatch(lyr));
                    else if (existing) clearSpinnerWatch(existing);
                    setVisible(existing, false);
                }
                ensureAoiOnTop();
                // Sync the old panel checkbox if it exists
                const oldCb = document.getElementById(`rptlayer_${i}`);
                if (oldCb) oldCb.checked = cb.checked;
            });
        });
    }

    // Layer Manager drag logic
    (function wireLayerMgrDrag() {
        if (!layerMgrHeader || !layerManagerWindow) return;
        let isDragging = false, startX = 0, startY = 0, origLeft = 0, origTop = 0;

        layerMgrHeader.addEventListener("mousedown", function (e) {
            if (e.target.tagName === "BUTTON") return;
            isDragging = true;
            const rect = layerManagerWindow.getBoundingClientRect();
            startX = e.clientX;
            startY = e.clientY;
            origLeft = rect.left;
            origTop = rect.top;
            layerManagerWindow.style.transition = "none";
        });

        document.addEventListener("mousemove", function (e) {
            if (!isDragging) return;
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;
            layerManagerWindow.style.left = (origLeft + dx) + "px";
            layerManagerWindow.style.top = (origTop + dy) + "px";
            layerManagerWindow.style.right = "auto";
        });

        document.addEventListener("mouseup", function () {
            isDragging = false;
            layerManagerWindow.style.transition = "";
        });

        // Touch support
        layerMgrHeader.addEventListener("touchstart", function (e) {
            if (e.target.tagName === "BUTTON") return;
            const touch = e.touches[0];
            isDragging = true;
            const rect = layerManagerWindow.getBoundingClientRect();
            startX = touch.clientX;
            startY = touch.clientY;
            origLeft = rect.left;
            origTop = rect.top;
            layerManagerWindow.style.transition = "none";
        }, { passive: true });

        document.addEventListener("touchmove", function (e) {
            if (!isDragging) return;
            const touch = e.touches[0];
            const dx = touch.clientX - startX;
            const dy = touch.clientY - startY;
            layerManagerWindow.style.left = (origLeft + dx) + "px";
            layerManagerWindow.style.top = (origTop + dy) + "px";
            layerManagerWindow.style.right = "auto";
        }, { passive: true });

        document.addEventListener("touchend", function () {
            isDragging = false;
            layerManagerWindow.style.transition = "";
        });
    })();

    // Layer Manager open/close
    if (advLayersBtn) {
        advLayersBtn.addEventListener("click", () => {
            if (!layerManagerWindow) return;
            const isOpen = !layerManagerWindow.classList.contains("hidden");
            if (isOpen) {
                layerManagerWindow.classList.add("hidden");
            } else {
                renderLayerManagerToggles();
                layerManagerWindow.classList.remove("hidden");
            }
        });
    }
    if (layerMgrCloseBtn) {
        layerMgrCloseBtn.addEventListener("click", () => {
            if (layerManagerWindow) layerManagerWindow.classList.add("hidden");
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
                await fetchJsonWithTimeout(pjsonUrl, timeoutMs, { noCache: true });
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
                    const pjson = await fetchJsonWithTimeout(pjsonUrl, timeoutMs, { noCache: true });

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
        if (resultsEl) resultsEl.innerHTML = cardsHtml || `<div class="small">No results yet.</div>`;
    }



// ========================================
// REFACTORED ANALYSIS FUNCTIONS
// ========================================

// Main orchestrator - runs ALL analysis steps

async function runAnalysis() {
    const myOp = startReportOp();

    const reportGeom = getReportGeometry();
    if (!reportGeom) { endReportOp(myOp); return; }

    const analysisStartTime = Date.now();

    // ── Background-tab warning system ──
    const bgBanner = document.getElementById("bgTabWarning");
    let bgNotifSent = false;

    function onVisibilityChange() {
        if (document.hidden) {
            // Tab went to background during analysis — show warning
            if (bgBanner) bgBanner.classList.remove("hidden");
            analysisModal.addLog("Tab hidden — screenshots will pause until you return", "warning");

            // Send a browser Notification so the user knows to return
            if (!bgNotifSent && "Notification" in window && Notification.permission === "granted") {
                new Notification("RMP Screening Tool", {
                    body: "Analysis is paused — return to this tab to continue map generation.",
                    icon: "https://www.blm.gov/themes/usasearch/img/favicon.ico"
                });
                bgNotifSent = true;
            } else if (!bgNotifSent && "Notification" in window && Notification.permission === "default") {
                Notification.requestPermission(); // prompt for next time
            }
        } else {
            // Tab visible again — hide warning
            if (bgBanner) bgBanner.classList.add("hidden");
            bgNotifSent = false;
        }
    }
    document.addEventListener("visibilitychange", onVisibilityChange);

    // ✅ Show analysis modal
    analysisModal.show();
    analysisModal.setProgress(0);
    analysisModal.setStep("Starting analysis...");
    const tier = getSelectedTier();
    const tierLabel = tier === 1 ? "Essential" : tier === 2 ? "Comprehensive" : "Complete";
    analysisModal.addLog("Analysis started — Tier " + tier + " (" + tierLabel + ")");

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

        if (aoiIsLarge) {
            const gridSize = config.report?.aoiChunkGridSize ?? 4;
            analysisModal.addLog(`Large AOI detected (${formatNumber(aoiCurrentAcres, 0)} acres) — using ${gridSize}×${gridSize} spatial chunking`, "warning");
        }
        
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
        if (copyReportLinkBtn) copyReportLinkBtn.disabled = false;

        setStatus("Analysis complete!");
        
        // ✅ Show success animation
        analysisModal.showSuccess(layersQueried, featuresFound, mapsGenerated, Date.now() - analysisStartTime);

        // Permitting mode: populate bucket results
        if (currentAppMode === "permit") {
            populatePermitBuckets();
            if (wizFullReport) wizFullReport.disabled = false;
            if (wizCopyReportLink) wizCopyReportLink.disabled = false;
            if (wizExportAll) wizExportAll.disabled = false;
        }

    } catch (e) {
        console.error(e);
        analysisModal.addLog(`Error: ${e.message}`, "error");
        analysisModal.setStep("Analysis failed");
        setStatus("Analysis failed (see console)");
        
        setTimeout(() => analysisModal.hide(), 3000);
    } finally {
        setBusy(false);
        endReportOp(myOp);
        document.removeEventListener("visibilitychange", onVisibilityChange);
        if (bgBanner) bgBanner.classList.add("hidden");
    }
}


// Extracted query logic (was: runReport)
async function queryAllLayers(reportGeom, myOp, modal = null) {
    if (resultsEl) resultsEl.innerHTML = "";
    if (exportAllBtn) exportAllBtn.disabled = true;
    lastReportRowsByLayer = [];

    const selectedTier = getSelectedTier();

    const combinedCfgs = [
        ...(config.reportLayers || []).filter(l => (l.tier || 1) <= selectedTier)
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

    // Separate configs into categories for parallel expansion
    const directTargets = [];
    const featureServerRoots = [];
    const mapServerRoots = [];

    for (const cfg of reportCfgs) {
        const url = String(cfg.url || "");

        if (isMapServerRoot(url) && isPlssLayerTitleOrUrl(cfg.title, url)) {
            continue;
        }

        if (cfg.imageService === true) {
            directTargets.push({ 
                title: cfg.title, 
                url,
                __isImageService: true,
                __renderingRule: cfg.renderingRule || null
            });
            continue;
        }

        if (isFeatureServerRoot(url)) {
            featureServerRoots.push(cfg);
            continue;
        }

        if (isMapServerRoot(url)) {
            mapServerRoots.push(cfg);
            continue;
        }

        directTargets.push({ title: cfg.title, url });
    }

    // Expand all service roots in parallel
    const [fsResults, msResults] = await Promise.all([
        Promise.allSettled(featureServerRoots.map(async cfg => {
            try {
                const sublayers = await expandServiceToSublayers(cfg.url);
                return sublayers.map(sl => ({
                    title: `${cfg.title}: ${sl.title}`,
                    url: sl.url
                }));
            } catch (e) {
                return [{ title: `${cfg.title} (FAILED to expand)`, url: cfg.url, error: e }];
            }
        })),
        Promise.allSettled(mapServerRoots.map(async cfg => {
            try {
                const subs = await expandMapServerToSublayers(cfg.url, { polygonOnly: false });
                return subs.map(sl => ({
                    title: `${cfg.title}: ${sl.title}`,
                    url: sl.url
                }));
            } catch (e) {
                return [{ title: `${cfg.title} (FAILED to expand)`, url: cfg.url, error: e }];
            }
        }))
    ]);

    // Collect all expanded targets
    expandedTargets.push(...directTargets);
    for (const r of fsResults) {
        if (r.status === "fulfilled") expandedTargets.push(...r.value);
    }
    for (const r of msResults) {
        if (r.status === "fulfilled") expandedTargets.push(...r.value);
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

        // Use chunked query for large AOIs to improve reliability
        const chunkThreshold = config.report?.aoiChunkThresholdAcres ?? 250000;
        const useChunking = aoiIsLarge && aoiCurrentAcres > chunkThreshold;

        const r = useChunking
            ? await querySingleLayerChunked(t.url, t.title, reportGeom, spatialRel)
            : await querySingleLayer(t.url, t.title, reportGeom, spatialRel);
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

    // ── Batched parallel queries (batches of 8) ──
    const BATCH_SIZE = 8;
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
    if (exportAllBtn) exportAllBtn.disabled = (lastReportRowsByLayer.length === 0);
    renderVisualSummary();
}




    // ── Shared export-all helper (used by both exportAllBtn and wizExportAll) ──
    async function doExportAll(callerBtn) {
        if (!lastReportRowsByLayer.length) return;
        if (callerBtn) callerBtn.disabled = true;
        try {
            setStatus("exporting ALL (FULL)…");
            const pageSize = config.report?.pageSize ?? 1000;
            const maxExport = config.report?.maxExportFeatures ?? 50000;

            // Build per-layer CSV blocks separated by a header row
            const blocks = [];
            for (let i = 0; i < lastReportRowsByLayer.length; i++) {
                const item = lastReportRowsByLayer[i];
                if (!item._layer || !item._exportQuery) continue;
                setStatus(`exporting ALL (FULL)… (${i + 1}/${lastReportRowsByLayer.length})`);
                if (!item.fullRows) {
                    const fullFeatures = await queryAllFeaturesPaged(
                        item._layer, item._exportQuery, pageSize, maxExport
                    );
                    item.fullRows = flattenAttributes(fullFeatures);
                }
                if (item.fullRows && item.fullRows.length) {
                    // Layer header row
                    blocks.push(`\n"=== ${item.title.replace(/"/g, '""')} ==="`);
                    // CSV for this layer's data
                    blocks.push(toCsv(item.fullRows));
                }
            }
            const csv = blocks.join("\n");
            downloadText("intersect_report_ALL_FULL.csv", csv || "");
            setStatus("exported ALL (FULL)");
        } catch (e) {
            console.error(e);
            setStatus("export ALL failed (see console)");
        } finally {
            if (callerBtn) callerBtn.disabled = false;
        }
    }

    function wireExportButtons() {
        if (!resultsEl) return;
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






    // setVisualStatus, renderVisualSummary, generateVisualReportData moved to visual-report.js
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

    // Feature picker modal managed by feature-picker.js module

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
            if (currentInteractionMode === "draw") {
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
            availableCreateTools: ["polygon", "polyline", "point"],
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

        // Point and line symbols for sketch preview
        sketch.pointSymbol = {
            type: "simple-marker",
            color: [40, 100, 60, 0.7],
            size: 10,
            outline: { color: [40, 100, 60, 1], width: 2 }
        };
        sketch.polylineSymbol = {
            type: "simple-line",
            color: [40, 100, 60, 0.85],
            width: 3,
            style: "solid"
        };

        sketch.on("create", (evt) => {
            if (evt.state === "complete") {
                currentInteractionMode = "select"; // restore click-to-select after drawing
                const geom = evt.graphic?.geometry || null;
                if (!geom) return;

                const gType = geom.type; // "polygon", "polyline", "point"
                const needsBuffer = gType !== "polygon";

                // Show the raw drawn geometry as a preview
                setAoiGeometry(geom);

                // Show the draw buffer panel
                const drawBufPanel = document.getElementById("drawBufferPanel");
                const drawBufMsg   = document.getElementById("drawBufferMsg");
                const drawBufInput = document.getElementById("drawBufferInput");
                const drawBufSkip  = document.getElementById("drawBufferSkip");

                if (drawBufPanel) {
                    const typeLabel = gType === "polyline" ? "line" : gType;
                    if (needsBuffer) {
                        if (drawBufMsg) drawBufMsg.innerHTML = "Optionally add a buffer around your drawn <strong>" + typeLabel + "</strong>.";
                        if (drawBufInput) drawBufInput.value = "";
                    } else {
                        if (drawBufMsg) drawBufMsg.innerHTML = "Optionally add a buffer around your drawn <strong>polygon</strong>.";
                        if (drawBufInput) drawBufInput.value = "";
                    }
                    if (drawBufSkip) drawBufSkip.classList.remove("hidden");

                    // Store drawn geometry for buffer apply
                    drawBufPanel._drawnGeom = geom;
                    drawBufPanel.classList.remove("hidden");

                    // Scroll the buffer options into view
                    requestAnimationFrame(() => {
                        drawBufPanel.scrollIntoView({ behavior: "smooth", block: "nearest" });
                    });

                    // Clear active preset
                    document.querySelectorAll(".draw-buf-preset").forEach(b => b.classList.remove("active"));
                }
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

        // ---------- Permit layer indexes ----------
        const allotmentIdx = findSelectionLayerIndexByNameIncludes("grazing allotment");
        const pastureIdx = findSelectionLayerIndexByNameIncludes("grazing pasture");
        const oilGasIdx = findSelectionLayerIndexByNameIncludes("oil and gas");
        const rowIdx = findSelectionLayerIndexByNameIncludes("rights-of-way");
        const miningIdx = findSelectionLayerIndexByNameIncludes("mining claims");
        const luaIdx = findSelectionLayerIndexByNameIncludes("lua leases");
        const geothermalIdx = findSelectionLayerIndexByNameIncludes("geothermal");
        const coalIdx = findSelectionLayerIndexByNameIncludes("coal");

        // All permit layer indices (used for mutual exclusion)
        const allPermitIndices = [allotmentIdx, pastureIdx, oilGasIdx, rowIdx, miningIdx, luaIdx, geothermalIdx, coalIdx].filter(i => i >= 0);


        // Helper: make ONE PLSS layer active, disable the other two, and auto-zoom if needed
        async function activatePlss(which, idxToEnable, { skipAutoZoom = false } = {}) {
            // Force select mode (PLSS tools are select-only)
            if (modeSelect && modeSelect.value !== "select") {
                modeSelect.value = "select";
                setMode("select");
            }

            // Disable all permit layers first
            for (const idx of allPermitIndices) disableSelectionLayer(idx);

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
            for (const idx of allPermitIndices) {
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
                    oilgas: "Oil & Gas Leases",
                    row: "Rights-of-Way",
                    mining: "Mining Claims (Active)",
                    lua: "Leases, Permits & Easements",
                    geothermal: "Geothermal Leases",
                    coal: "Coal Leases / Cases"
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

        // Copy Report Link buttons
        async function copyReportLink(btn) {
            const reportId = getLastReportId();
            if (!reportId) {
                alert("Run analysis first to generate the report.");
                return;
            }
            const url = getReportShareUrl(reportId);
            try {
                await navigator.clipboard.writeText(url);
                const origText = btn.textContent;
                btn.textContent = "✅ Copied!";
                setTimeout(() => { btn.textContent = origText; }, 2000);
            } catch (e) {
                // Fallback: select from a prompt
                window.prompt("Copy this link:", url);
            }
        }
        if (copyReportLinkBtn) copyReportLinkBtn.addEventListener("click", () => copyReportLink(copyReportLinkBtn));
        if (wizCopyReportLink) wizCopyReportLink.addEventListener("click", () => copyReportLink(wizCopyReportLink));


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

        if (runBtn) runBtn.addEventListener("click", () => {
            if (aoiIsLarge) {
                const sqMiles = (aoiCurrentAcres / 640).toFixed(0);
                const confirmed = window.confirm(
                    `\u26A0\uFE0F Large Area of Interest\n\n` +
                    `Your AOI is approximately ${formatNumber(aoiCurrentAcres, 0)} acres (~${sqMiles} sq mi).\n\n` +
                    `Analysis over this area will take significantly longer and some ` +
                    `service queries may time out. The analysis will use spatial chunking ` +
                    `to improve reliability.\n\n` +
                    `Do you want to proceed?`
                );
                if (!confirmed) return;
            }
            runAnalysis();
        });

        if (clearBtn) clearBtn.addEventListener("click", clearAll);


        if (exportAllBtn) exportAllBtn.addEventListener("click", () => doExportAll(exportAllBtn));


        // ========================================
        // FEATURE SEARCH WIDGET
        // ========================================
        searchModule.init(state, {
            Graphic, GraphicsLayer, enableSelectionLayer
        });

        // ========================================
        // PERMITTING MODE — Wizard button wiring
        // ========================================

        // Mode toggle
        if (permitModeBtn) permitModeBtn.addEventListener("click", () => setAppMode("permit"));
        if (advancedModeBtn) advancedModeBtn.addEventListener("click", () => setAppMode("advanced"));

        // AOI method card clicks
        document.querySelectorAll("#aoiMethods .aoi-method-card").forEach(card => {
            card.addEventListener("click", () => {
                const method = card.dataset.method;
                if (method) showAoiMethod(method);
            });
        });

        // ── File Upload AOI wiring ──
        (function wireUploadAoi() {
            const dropZone     = document.getElementById("uploadDropZone");
            const fileInput    = document.getElementById("aoiFileInput");
            const uploadStatus = document.getElementById("uploadStatus");
            if (!dropZone || !fileInput) return;

            function showUploadStatus(msg, cls) {
                if (!uploadStatus) return;
                uploadStatus.className = "upload-status " + cls;
                uploadStatus.innerHTML = msg;
                uploadStatus.classList.remove("hidden");
            }
            function hideUploadStatus() {
                if (uploadStatus) uploadStatus.classList.add("hidden");
            }

            async function handleFile(file) {
                if (!file) return;
                hideUploadStatus();
                showUploadStatus('<span class="spinner" style="width:14px;height:14px;"></span> Processing <b>' + configHelpers.escapeHtml(file.name) + '</b>…', "loading");

                // Hide any previous buffer panel
                const bufferPanel = document.getElementById("uploadBufferPanel");
                if (bufferPanel) bufferPanel.classList.add("hidden");

                try {
                    const viewSR = state.view ? state.view.spatialReference : null;
                    const result = await uploadAoiModule.processFile(file, viewSR);

                    const needsBuffer = !result.hasPolygons; // points/lines MUST be buffered
                    const geomLabel = result.geometryType === "point" ? "point"
                                    : result.geometryType === "polyline" ? "line" : "polygon";
                    const geomLabelPlural = result.featureCount === 1 ? geomLabel : geomLabel + "s";

                    // Show the file info
                    const countLabel = result.featureCount + " " + geomLabelPlural;
                    showUploadStatus("📂 <b>" + configHelpers.escapeHtml(result.fileName) + "</b> — " + countLabel, "success");

                    // Zoom to uploaded geometry preview
                    if (view && result.geometry && result.geometry.extent) {
                        // Show the raw geometry on the map as a preview
                        setAoiGeometry(result.geometry);
                        await view.goTo(result.geometry.extent.expand(1.5), { animate: true, duration: 600 });
                    }

                    // Show buffer panel
                    if (bufferPanel) {
                        const bufferMsg = document.getElementById("uploadBufferMsg");
                        const bufferInput = document.getElementById("uploadBufferInput");
                        const bufferApplyBtn = document.getElementById("uploadBufferApply");
                        const bufferSkipBtn = document.getElementById("uploadBufferSkip");

                        if (needsBuffer) {
                            // Points/lines require a buffer
                            if (bufferMsg) bufferMsg.innerHTML = "Your file contains <strong>" + geomLabelPlural + "</strong>. A buffer is required to create a polygon AOI.";
                            if (bufferInput) bufferInput.value = geomLabel === "point" ? "1" : "0.5";
                            if (bufferSkipBtn) bufferSkipBtn.classList.add("hidden");
                        } else {
                            // Polygons — buffer is optional
                            if (bufferMsg) bufferMsg.innerHTML = "Optionally add a buffer around your <strong>" + geomLabelPlural + "</strong>.";
                            if (bufferInput) bufferInput.value = "";
                            if (bufferSkipBtn) bufferSkipBtn.classList.remove("hidden");
                        }

                        bufferPanel.classList.remove("hidden");

                        // Store result for the Apply/Skip handlers
                        bufferPanel._uploadResult = result;
                    }
                } catch (err) {
                    console.error("Upload AOI error:", err);
                    showUploadStatus("❌ " + configHelpers.escapeHtml(err.message || "Failed to process file"), "error");
                }
            }

            /**
             * Finalize the upload AOI — apply buffer if needed, set geometry.
             */
            function finalizeUpload(result, bufferMiles) {
                let geometry;
                if (bufferMiles && bufferMiles > 0) {
                    geometry = uploadAoiModule.applyBuffer(result.allGeometries, bufferMiles);
                    if (!geometry) {
                        showUploadStatus("❌ Buffer operation failed. Try a different value.", "error");
                        return;
                    }
                } else if (result.hasPolygons) {
                    // Use polygon geometries without buffer
                    geometry = uploadAoiModule.applyBuffer(result.allGeometries, 0);
                    if (!geometry) {
                        showUploadStatus("❌ No polygon geometry available.", "error");
                        return;
                    }
                } else {
                    showUploadStatus("❌ A buffer is required for point/line features.", "error");
                    return;
                }

                selectionGeom = geometry;
                aoiSource = "upload";
                aoiSourceLayerTitle = result.fileName;
                setAoiGeometry(geometry);

                if (runBtn) runBtn.disabled = false;

                // Zoom to buffered extent
                if (view && geometry && geometry.extent) {
                    view.goTo(geometry.extent.expand(1.3), { animate: true, duration: 600 });
                }

                // Update mask
                if (typeof updateAoiMask === "function") updateAoiMask();

                const bufLabel = bufferMiles > 0 ? " with " + bufferMiles + " mi buffer" : "";
                showUploadStatus("✅ <b>" + configHelpers.escapeHtml(result.fileName) + "</b> loaded" + bufLabel, "success");

                // Hide buffer panel
                const bp = document.getElementById("uploadBufferPanel");
                if (bp) bp.classList.add("hidden");

                // Advance wizard to Step 2
                if (currentAppMode === "permit" && currentWizardStep === 1) {
                    populateAoiConfirmation();
                    goToWizardStep(2);
                }
            }

            // Wire buffer panel buttons
            const bufferApplyBtn = document.getElementById("uploadBufferApply");
            const bufferSkipBtn  = document.getElementById("uploadBufferSkip");
            const bufferPanel    = document.getElementById("uploadBufferPanel");

            if (bufferApplyBtn) {
                bufferApplyBtn.addEventListener("click", function () {
                    const input = document.getElementById("uploadBufferInput");
                    const val = parseFloat(input ? input.value : "");
                    const result = bufferPanel ? bufferPanel._uploadResult : null;
                    if (!result) return;

                    if (isNaN(val) || val <= 0) {
                        if (!result.hasPolygons) {
                            showUploadStatus("❌ Please enter a buffer distance greater than 0.", "error");
                            return;
                        }
                        // Polygon with no buffer — treat as skip
                        finalizeUpload(result, 0);
                        return;
                    }
                    if (val > 50) {
                        showUploadStatus("❌ Buffer cannot exceed 50 miles.", "error");
                        return;
                    }
                    finalizeUpload(result, val);
                });
            }

            if (bufferSkipBtn) {
                bufferSkipBtn.addEventListener("click", function () {
                    const result = bufferPanel ? bufferPanel._uploadResult : null;
                    if (!result) return;
                    finalizeUpload(result, 0);
                });
            }

            // Wire buffer preset buttons
            document.querySelectorAll(".upload-buffer-preset").forEach(function (btn) {
                btn.addEventListener("click", function () {
                    const input = document.getElementById("uploadBufferInput");
                    if (input) input.value = btn.dataset.val;
                    // Highlight active preset
                    document.querySelectorAll(".upload-buffer-preset").forEach(b => b.classList.remove("active"));
                    btn.classList.add("active");
                });
            });

            // File input change
            fileInput.addEventListener("change", function () {
                if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
                fileInput.value = ""; // allow re-uploading same file
            });

            // Click on drop zone triggers file browse
            dropZone.addEventListener("click", function (e) {
                // Don't trigger if clicking the browse button's label (it already triggers input)
                if (e.target.closest(".upload-browse-btn")) return;
                fileInput.click();
            });

            // Drag-and-drop events
            ["dragenter", "dragover"].forEach(function (evtName) {
                dropZone.addEventListener(evtName, function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    dropZone.classList.add("drag-over");
                });
            });
            ["dragleave", "drop"].forEach(function (evtName) {
                dropZone.addEventListener(evtName, function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    dropZone.classList.remove("drag-over");
                });
            });
            dropZone.addEventListener("drop", function (e) {
                const files = e.dataTransfer && e.dataTransfer.files;
                if (files && files.length > 0) handleFile(files[0]);
            });
        })();

        // Wizard PLSS buttons
        if (wizTownshipBtn) wizTownshipBtn.addEventListener("click", () => { activatePlss("township", townshipIdx); setWizPlssActive("township"); });
        if (wizSectionBtn) wizSectionBtn.addEventListener("click", () => { activatePlss("section", sectionIdx); setWizPlssActive("section"); });
        if (wizParcelBtn) wizParcelBtn.addEventListener("click", () => { activatePlss("intersected", intersectedIdx); setWizPlssActive("intersected"); });

        // Wizard permit type list — click handlers on custom buttons
        const permitValueToIdx = {
            allotment: allotmentIdx,
            pasture: pastureIdx,
            oilgas: oilGasIdx,
            row: rowIdx,
            mining: miningIdx,
            lua: luaIdx,
            geothermal: geothermalIdx,
            coal: coalIdx
        };

        if (wizPermitList) {
            wizPermitList.querySelectorAll(".permit-item").forEach(btn => {
                btn.addEventListener("click", () => {
                    const val = btn.dataset.permit;
                    const idx = permitValueToIdx[val];
                    if (val && idx !== undefined) {
                        activatePermitLayer(val, idx);
                    }
                });
            });
        }

        // ── Permit layer feature-in-view indicators ──
        // Query each permit layer to see if features exist in the current map extent.
        // Debounced on extent change, runs when the permit panel is visible.
        let _permitExtentTimer = null;

        function updatePermitIndicators() {
            if (!wizPermitList) return;
            // Only run if the permit panel is visible
            const panel = document.getElementById("aoiMethodPermit");
            if (!panel || panel.classList.contains("hidden")) return;

            const extent = view?.extent;
            if (!extent) return;

            // For each permit item, query feature count within extent
            wizPermitList.querySelectorAll(".permit-item").forEach(btn => {
                const val = btn.dataset.permit;
                const idx = permitValueToIdx[val];
                if (idx === undefined || idx < 0) return;

                const layer = selectionLayers[idx]?.layer;
                if (!layer) return;

                const dot = btn.querySelector(".permit-feat-dot");
                const spinner = btn.querySelector(".permit-spinner");
                if (!dot) return;

                // Show spinner, set dot to checking
                if (spinner) spinner.classList.remove("hidden");
                dot.dataset.status = "checking";
                dot.title = "Checking…";

                const query = layer.createQuery();
                query.geometry = extent;
                query.spatialRelationship = "intersects";
                query.returnGeometry = false;

                layer.queryFeatureCount(query).then(count => {
                    if (spinner) spinner.classList.add("hidden");
                    if (count > 0) {
                        dot.dataset.status = "found";
                        dot.title = `${count.toLocaleString()} feature${count !== 1 ? "s" : ""} in view`;
                    } else {
                        dot.dataset.status = "none";
                        dot.title = "No features in current view";
                    }
                }).catch(() => {
                    if (spinner) spinner.classList.add("hidden");
                    dot.dataset.status = "error";
                    dot.title = "Could not query layer";
                });
            });
        }

        function schedulePermitIndicatorUpdate() {
            clearTimeout(_permitExtentTimer);
            _permitExtentTimer = setTimeout(updatePermitIndicators, 600);
        }

        // Watch for extent changes (debounced)
        if (view) {
            view.watch("stationary", (stationary) => {
                if (stationary) schedulePermitIndicatorUpdate();
            });
        }

        // Also update when the permit panel becomes visible
        const _origShowAoiMethod = typeof showAoiMethod === "function" ? showAoiMethod : null;
        if (_origShowAoiMethod) {
            // Patch showAoiMethod to trigger indicator update when "permit" panel opens
            const _aoiMethodPermitPanel = document.getElementById("aoiMethodPermit");
            if (_aoiMethodPermitPanel) {
                const observer = new MutationObserver(() => {
                    if (!_aoiMethodPermitPanel.classList.contains("hidden")) {
                        schedulePermitIndicatorUpdate();
                    }
                });
                observer.observe(_aoiMethodPermitPanel, { attributes: true, attributeFilter: ["class"] });
            }
        }

        // Wizard draw buttons

        // Draw tool type selector
        const drawToolBtns = document.querySelectorAll(".draw-tool-btn");
        const drawHintEl   = document.getElementById("drawHint");
        const DRAW_HINTS = {
            polygon: "Click on the map to place vertices. Double-click to finish the polygon.",
            polyline: "Click on the map to place vertices. Double-click to finish the line.",
            point: "Click on the map to place a point."
        };

        drawToolBtns.forEach(btn => {
            btn.addEventListener("click", () => {
                drawToolBtns.forEach(b => b.classList.remove("active"));
                btn.classList.add("active");
                if (btn.id === "wizDrawPolygon")  currentDrawToolType = "polygon";
                else if (btn.id === "wizDrawLine") currentDrawToolType = "polyline";
                else if (btn.id === "wizDrawPoint") currentDrawToolType = "point";
                if (drawHintEl) drawHintEl.textContent = DRAW_HINTS[currentDrawToolType];

                // Immediately start drawing with the selected tool type
                if (!sketch) return;
                const dbp = document.getElementById("drawBufferPanel");
                if (dbp) dbp.classList.add("hidden");
                sketch.cancel();
                currentInteractionMode = "draw";
                sketch.create(currentDrawToolType);
                setStatus("drawing " + currentDrawToolType + "\u2026");
            });
        });

        if (wizStopDrawBtn) {
            wizStopDrawBtn.addEventListener("click", () => {
                if (sketch) sketch.cancel();
                currentInteractionMode = "select";

                // Clear any existing drawn AOI
                selectionGeom = null;
                aoiSource = null;
                aoiSourceLayerTitle = null;
                aoiSourceLayerUrl = null;
                aoiSourceObjectId = null;
                aoiSourceObjectIdField = null;
                aoiSourceFeature = null;
                if (aoiLayer) aoiLayer.removeAll();
                aoiGraphic = null;
                if (runBtn) runBtn.disabled = true;

                // Hide draw buffer panel
                const dbp = document.getElementById("drawBufferPanel");
                if (dbp) dbp.classList.add("hidden");

                // Reset wizard UI
                if (wizFullReport) wizFullReport.disabled = true;
                if (wizExportAll) wizExportAll.disabled = true;

                setStatus("draw canceled");
            });
        }

        // Draw buffer panel handlers
        function finalizeDrawAoi(geom, bufferMiles) {
            let finalGeom;
            if (bufferMiles && bufferMiles > 0) {
                finalGeom = uploadAoiModule.applyBuffer([geom], bufferMiles);
                if (!finalGeom) {
                    setStatus("Buffer operation failed.");
                    return;
                }
            } else {
                finalGeom = geom;
            }

            setAoiGeometry(finalGeom);
            resetCoverageCacheForAoi(finalGeom);
            aoiSource = "draw";
            aoiSourceLayerTitle = null;
            aoiSourceLayerUrl = null;
            aoiSourceObjectId = null;
            aoiSourceObjectIdField = null;
            aoiSourceFeature = null;
            setGeometryFromSelection(finalGeom);

            const bufLabel = bufferMiles > 0 ? " with " + bufferMiles + " mi buffer" : "";
            setStatus("Drawn AOI ready" + bufLabel);

            // Hide buffer panel
            const dbp = document.getElementById("drawBufferPanel");
            if (dbp) dbp.classList.add("hidden");
        }

        const drawBufApply = document.getElementById("drawBufferApply");
        const drawBufSkip  = document.getElementById("drawBufferSkip");
        const drawBufPanel = document.getElementById("drawBufferPanel");

        if (drawBufApply) {
            drawBufApply.addEventListener("click", () => {
                const input = document.getElementById("drawBufferInput");
                const val = parseFloat(input ? input.value : "");
                const geom = drawBufPanel ? drawBufPanel._drawnGeom : null;
                if (!geom) return;

                if (isNaN(val) || val <= 0) {
                    finalizeDrawAoi(geom, 0);
                    return;
                }
                if (val > 50) {
                    setStatus("Buffer cannot exceed 50 miles.");
                    return;
                }
                finalizeDrawAoi(geom, val);
            });
        }

        if (drawBufSkip) {
            drawBufSkip.addEventListener("click", () => {
                const geom = drawBufPanel ? drawBufPanel._drawnGeom : null;
                if (!geom) return;
                finalizeDrawAoi(geom, 0);
            });
        }

        // Draw buffer preset buttons
        document.querySelectorAll(".draw-buf-preset").forEach(btn => {
            btn.addEventListener("click", () => {
                const input = document.getElementById("drawBufferInput");
                if (input) input.value = btn.dataset.val;
                document.querySelectorAll(".draw-buf-preset").forEach(b => b.classList.remove("active"));
                btn.classList.add("active");
            });
        });

        // Location search
        if (wizLocationInput) {
            wizLocationInput.addEventListener("input", () => {
                performLocationSearch(wizLocationInput.value.trim());
            });
        }

        // Wizard navigation
        if (wizBackToStep1) {
            wizBackToStep1.addEventListener("click", () => {
                // Cancel any active sketch drawing
                if (sketch) sketch.cancel();
                // Clear location search state
                clearTimeout(wizLocationDebounce);
                wizLocationDebounce = null;
                if (wizLocationInput) wizLocationInput.value = "";
                if (wizLocationResults) { wizLocationResults.innerHTML = ""; wizLocationResults.classList.add("hidden"); }
                // Reset PLSS button states
                setWizPlssActive(null);
                goToWizardStep(1);
                hideAoiMethodPanels();
            });
        }

        if (wizScreenBtn) wizScreenBtn.addEventListener("click", () => {
            if (aoiIsLarge) {
                const sqMiles = (aoiCurrentAcres / 640).toFixed(0);
                const confirmed = window.confirm(
                    `\u26A0\uFE0F Large Area of Interest\n\n` +
                    `Your AOI is approximately ${formatNumber(aoiCurrentAcres, 0)} acres (~${sqMiles} sq mi).\n\n` +
                    `Analysis over this area will take significantly longer and some ` +
                    `service queries may time out. The analysis will use spatial chunking ` +
                    `to improve reliability.\n\n` +
                    `Do you want to proceed?`
                );
                if (!confirmed) return;
            }
            runAnalysis();
        });

        if (wizNewScreening) {
            wizNewScreening.addEventListener("click", () => {
                clearAll();
                goToWizardStep(1);
                hideAoiMethodPanels();
            });
        }

        if (wizFullReport) wizFullReport.addEventListener("click", viewFinalReport);

        if (wizExportAll) {
            wizExportAll.addEventListener("click", () => doExportAll(wizExportAll));
        }

        // Bucket tabs
        document.querySelectorAll("#permitBucketTabs .permit-tab").forEach(tab => {
            tab.addEventListener("click", () => {
                const bk = tab.dataset.bucket;
                if (bk) setActiveBucket(bk);
            });
        });

        // Initialize in permit mode
        setAppMode("permit");
        goToWizardStep(1);

        setMode("select");
        setActiveTab("layers");
        setLoadingState("Ready", 100);
        setStatus("ready");

        // Brief delay so the bar visually reaches 100% before fading
        await new Promise(r => setTimeout(r, 350));
        hideLoadingOverlay();

        // Check for ?report=<id> in URL and open the saved report
        (async function checkSharedReport() {
            try {
                const params = new URLSearchParams(window.location.search);
                const reportId = params.get("report");
                if (!reportId) return;
                const html = await loadReportFromDb(reportId);
                if (html) {
                    openHtmlInNewTab(html);
                } else {
                    alert("This report link has expired or is not available on this device.\n\nReport links are valid for 7 days and can only be opened on the device where the analysis was run.");
                }
                // Clean the URL so a refresh doesn't re-open
                const cleanUrl = window.location.origin + window.location.pathname;
                window.history.replaceState(null, "", cleanUrl);
            } catch (e) {
                console.warn("Failed to load shared report:", e);
            }
        })();

        // Cleanup expired reports from IndexedDB
        cleanupExpiredReports();

        // Initialize analysis modal
        analysisModal.init();

        // Background service check — delay 5s so it doesn't compete with
        // critical layer loading for browser connection slots
        setTimeout(() => checkServiceStatusBackground().catch(() => {}), 5000);
    }

    init().catch((e) => {
        console.error(e);
        setStatus("failed to initialize (see console)");
    });

});