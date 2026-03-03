/* global require */

require([
    "app/config-helpers",
    "app/map-utils",
    "app/query-engine",
    "app/final-report",
    "app/feature-picker",
    "app/search",
    "app/upload-aoi",
    "app/permit-types",
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
], function (configHelpers, mapUtilsModule, queryEngineModule, finalReportModule, featurePickerModule, searchModule, uploadAoiModule, permitTypesModule, EsriMap, MapView, FeatureLayer, GraphicsLayer, Sketch, Graphic, geometryEngine, TileLayer, ImageryLayer, esriId) {

    // ── Suppress ArcGIS Online sign-in popup ──
    // All services used by this app are publicly shared; prevent the
    // IdentityManager from showing a login dialog on 401/403 responses.
    esriId.useSignInPage = false;
    esriId.on("credential-create", function (evt) {
        if (evt.credential) evt.credential.destroy();
    });

    // ── Destructure config-helpers for functions already extracted ──
    const {
        escapeHtml, normalize,
        isPlssLayerTitleOrUrl,
        isFeatureServerRoot, isMapServerRoot,
        safeFilename, formatNumber,
        fetchJson, fetchJsonWithTimeout,
        normalizePjsonUrl, normalizeUrlKey,
        pickServiceDescription, buildLayerCfgIndex, getConfiguredServices,
        setBasemapBaseLayerOpacity, isImageryBasemap,
        expandMapServerToSublayers, expandServiceToSublayers,
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

    // Feedback floating window
    const feedbackBtn = document.getElementById("feedbackBtn");
    const feedbackWindow = document.getElementById("feedbackWindow");
    const feedbackCloseBtn = document.getElementById("feedbackCloseBtn");
    const feedbackForm = document.getElementById("feedbackForm");
    const feedbackStatus = document.getElementById("feedbackStatus");

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
    const tabFinalReportBtn = document.getElementById("tabFinalReportBtn");

    const tabLayersPanel = document.getElementById("tabLayersPanel");
    const tabServicesPanel = document.getElementById("tabServicesPanel");  
    const tabReportPanel = document.getElementById("tabReportPanel");      
    const tabReportFinalPanel = document.getElementById("tabReportFinalPanel");

    // Final report DOM
    const viewReportBtn = document.getElementById("viewReportBtn");
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
    const wizNewScreening = document.getElementById("wizNewScreening");
    const wizFullReport = document.getElementById("wizFullReport");
    const wizStopDrawBtn = document.getElementById("wizStopDrawBtn");
    const wizTownshipBtn = document.getElementById("wizTownshipBtn");
    const wizSectionBtn = document.getElementById("wizSectionBtn");
    const wizParcelBtn = document.getElementById("wizParcelBtn");
    const wizPermitList = document.getElementById("wizPermitList");
    const wizLocationInput = document.getElementById("wizLocationInput");
    const wizLocationResults = document.getElementById("wizLocationResults");

    // Service-down warning modal
    const serviceDownModalEl = document.getElementById("serviceDownModal");
    const serviceDownSummaryEl = document.getElementById("serviceDownSummary");
    const serviceDownListEl = document.getElementById("serviceDownList");
    const serviceDownCloseBtn = document.getElementById("serviceDownCloseBtn");

    /* ── Layer count helper (all layers, no tier filtering) ── */
    function getTotalLayerCount() {
      return (config.reportLayers || []).length;
    }

    /**
     * Smoothly scroll the side-panel so the latest content at the bottom is visible.
     * Optionally pass a target element to scroll that specific element into view,
     * otherwise scrolls the panel all the way to the bottom.
     */
    function scrollPanelToBottom(targetEl) {
        const panel = document.getElementById("panel");
        if (!panel) return;
        // Use a small timeout so the DOM fully renders before measuring scroll height.
        // requestAnimationFrame alone can fire before layout completes on mobile Safari.
        setTimeout(() => {
            requestAnimationFrame(() => {
                if (targetEl) {
                    targetEl.scrollIntoView({ behavior: "smooth", block: "end" });
                } else {
                    panel.scrollTo({ top: panel.scrollHeight, behavior: "smooth" });
                }
            });
        }, 80);
    }

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
                    if (reportAbortController) {
                        reportAbortController.abort(); // Cancel in-flight network requests
                        reportAbortController = null;
                    }
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
            if (!this.el) {
                // Fallback: try to find it again
                this.el = document.getElementById("analysisModal");
            }
            if (!this.el) {
                console.error("[analysisModal] Cannot find #analysisModal element!");
                return;
            }
            this.el.classList.remove("hidden");
            this.el.style.display = "flex"; // belt-and-suspenders
            this.reset();
            console.log("[analysisModal] Modal shown", this.el);
        },
        
        hide() {
            if (!this.el) return;
            this.el.classList.add("hidden");
            this.el.style.display = ""; // clear inline override
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

    // ---------- Report Building Modal Helpers ----------
    let reportBuildCanceled = false;

    const reportModal = {
        el: null,
        progressFill: null,
        currentStep: null,
        progressDetail: null,
        success: null,
        _startTime: null,
        _onViewReport: null,
        stats: {
            mapsGenerated: null
        },
        
        init() {
            this.el = document.getElementById("reportModal");
            this.progressFill = document.getElementById("reportProgressFill");
            this.currentStep = document.getElementById("reportCurrentStep");
            this.progressDetail = document.getElementById("reportProgressDetail");
            this.stats.mapsGenerated = document.getElementById("reportMapsGenerated");
            this.success = document.getElementById("reportSuccess");
            
            // Wire cancel button
            const cancelBtn = document.getElementById("reportModalCancelBtn");
            if (cancelBtn) {
                cancelBtn.addEventListener("click", () => {
                    reportBuildCanceled = true;
                    this.hide();
                    lockMapInteraction(false);
                    setStatus("Report canceled");
                });
            }

            // Wire success "View Report" button
            const successCloseBtn = document.getElementById("reportSuccessCloseBtn");
            if (successCloseBtn) {
                successCloseBtn.addEventListener("click", async () => {
                    this.hide();
                    try {
                        if (this._onViewReport) await this._onViewReport();
                    } catch (e) {
                        console.warn('[report] Error opening report:', e);
                    }
                });
            }
        },
        
        show() {
            if (!this.el) {
                this.el = document.getElementById("reportModal");
            }
            if (!this.el) {
                console.error("[reportModal] Cannot find #reportModal element!");
                return;
            }
            reportBuildCanceled = false;
            this._startTime = Date.now();
            this.el.classList.remove("hidden");
            this.reset();
            lockMapInteraction(true);
        },
        
        hide() {
            if (!this.el) return;
            this.el.classList.add("hidden");
            lockMapInteraction(false);
        },
        
        reset() {
            this.setProgress(0);
            this.setStep("Initializing...");
            this.setProgressDetail("Preparing report...");
            this.updateStats(0);
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
        
        setProgressDetail(text) {
            if (this.progressDetail) {
                this.progressDetail.textContent = text;
            }
        },
        
        updateStats(mapsGenerated) {
            if (this.stats.mapsGenerated) this.stats.mapsGenerated.textContent = mapsGenerated;
        },
        
        showSuccess({ mapsGenerated, layersQueried, featuresFound } = {}) {
            if (!this.success) return;

            const elMaps = document.getElementById("reportSuccessMaps");
            const elLayers = document.getElementById("reportSuccessLayers");
            const elFeatures = document.getElementById("reportSuccessFeatures");
            const elDuration = document.getElementById("reportSuccessDuration");

            if (elMaps) elMaps.textContent = mapsGenerated ?? 0;
            if (elLayers) elLayers.textContent = layersQueried ?? 0;
            if (elFeatures) elFeatures.textContent = featuresFound ?? 0;

            if (elDuration && this._startTime) {
                const totalSec = Math.round((Date.now() - this._startTime) / 1000);
                const min = Math.floor(totalSec / 60);
                const sec = totalSec % 60;
                elDuration.textContent = min > 0
                    ? `Completed in ${min}m ${sec}s`
                    : `Completed in ${sec}s`;
            }

            this.success.classList.remove("hidden");
        },

        isCanceled() {
            return reportBuildCanceled;
        }
    };

    // ---------- Service Down Warning Modal Helpers ----------
    let serviceDownNoticeShown = false;

    const serviceDownModal = {
        show(downItems, warnItems, sourceLabel, refreshedAt) {
            const allItems = (downItems || []).concat(warnItems || []);
            if (!serviceDownModalEl || !serviceDownListEl || !allItems.length) return;

            const sourceText = sourceLabel === "r2"
                ? "from cached service checks"
                : "from live fallback checks";
            const refreshedText = refreshedAt ? ` (last refresh ${refreshedAt})` : "";

            const parts = [];
            if (downItems.length) parts.push(`${downItems.length} DOWN`);
            if (warnItems.length) parts.push(`${warnItems.length} degraded`);

            if (serviceDownSummaryEl) {
                serviceDownSummaryEl.textContent = `${parts.join(", ")} service${allItems.length !== 1 ? "s" : ""} detected ${sourceText}${refreshedText}.`;
            }

            serviceDownListEl.innerHTML = allItems.map(it => {
                const statusClass = it._warnLevel === "WARN" ? "service-warn-item" : "service-down-item";
                const badge = it._warnLevel === "WARN"
                    ? '<span class="service-badge warn">DEGRADED</span>'
                    : '<span class="service-badge down">DOWN</span>';
                return `
                    <li class="${statusClass}">
                        <div class="service-down-item-title">${badge} ${escapeHtml(it.title || "Unnamed Service")}</div>
                        <div class="service-down-item-meta">${escapeHtml(it.kind || "Service")}</div>
                        <div class="service-down-item-url"><a href="${escapeHtml(it.url)}" target="_blank" rel="noopener">${escapeHtml(it.url)}</a></div>
                    </li>`;
            }).join("");

            serviceDownModalEl.classList.remove("hidden");
            serviceDownNoticeShown = true;
        },
        hide() {
            if (!serviceDownModalEl) return;
            serviceDownModalEl.classList.add("hidden");
            serviceDownNoticeShown = false;
        },
        wire() {
            if (serviceDownCloseBtn) {
                serviceDownCloseBtn.addEventListener("click", () => this.hide());
            }
            if (serviceDownModalEl) {
                serviceDownModalEl.addEventListener("click", (e) => {
                    if (e.target && e.target.classList && e.target.classList.contains("service-down-modal-backdrop")) {
                        this.hide();
                    }
                });
            }
        }
    };

    function maybeShowDownServiceWarning(items, sourceLabel, refreshedAt) {
        if (serviceDownNoticeShown || !items || !items.length) return;
        const downItems = items.filter(it => {
            if (serviceStatus.get(it.url) !== "DOWN") return false;
            if (sourceLabel !== "r2") return true;
            const hasHistory = serviceStatus.get(it.url + "::normallyHasFeatures");
            if (hasHistory === false) return false;
            return true; // true or unknown/null => still show as potentially impactful
        });
        // Also collect WARN items (only from R2 source, not fallback pings)
        const warnItems = (sourceLabel === "r2")
            ? items.filter(it => serviceStatus.get(it.url) === "WARN").map(it => ({ ...it, _warnLevel: "WARN" }))
            : [];
        if (!downItems.length && !warnItems.length) return;
        serviceDownModal.show(downItems, warnItems, sourceLabel, refreshedAt);
    }


    // ---------- State ----------
    let config = null;

    let view = null;
    let selectionGeom = null;
    let aoiOriginalGeom = null;      // pre-buffer geometry for reset
    let aoiSource = null;            // "draw" | "select" | "upload"
    let aoiSourceLayerTitle = null;  // optional: which selection layer was clicked
    let map = null; // <-- add (so PLSS buttons can add/remove selection layers)

    // Track service status (url -> "UP" | "DOWN")
    const serviceStatus = new Map();

    // ── Sublayer expansion cache ──
    // Pre-expanded sublayer lists keyed by service root URL → Promise<sublayer[]>
    const _sublayerCache = new Map();

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
    let plssStateLayerUrl = null; // Set from config.referenceLayers.plssStateBoundary at init
    let aoiSourceObjectId = null;      // ObjectID of the clicked AOI polygon (select mode)
    let aoiSourceObjectIdField = null; // ObjectID field name for that layer
    let aoiSourceFeature = null;       // ✅ cached clicked feature (attributes for AOI Source card)
    let aoiLocationLabel = null;       // Reverse-geocoded location label for reports
    let aoiBufferMiles = 0;            // Buffer distance applied to AOI (miles), 0 = none


    let sketch = null;
    let currentDrawToolType = "polygon";
    let _clickSelectGen = 0; // generation counter to discard stale click-to-select results

    let lastReportRowsByLayer = []; // for export-all
    let reportLayerViews = new Map();

    // ── Permit Types Module ──
    const {
        PERMIT_TYPES, PERMIT_TYPE_ORDER, CATEGORY_DEFS, CATEGORY_ORDER,
        resolveCategory, filterLayersByPermitType, categorizeByPermitType,
        categorizeIntoBuckets: _categorizeIntoBucketsShared
    } = permitTypesModule;

    // ── Permitting Mode State ──
    let currentAppMode = "permit"; // "permit" | "advanced"
    let currentWizardStep = 1;
    let selectedPermitType = null; // e.g. "oil-gas", "grazing", etc.
    let currentAoiMethod = null; // "search" | "permit" | "select" | "draw" | "upload"
    let currentInteractionMode = "select"; // PERF-TEST: tracks draw/select without modeSelect DOM

    // ── Shared state object for modules (properties updated by app.js) ──
    const state = {
        get config() { return config; },
        get view() { return view; },
        get map() { return map; },
        get selectionGeom() { return selectionGeom; },
        get aoiOriginalGeom() { return aoiOriginalGeom; },
        get aoiLayer() { return aoiLayer; },
        get aoiMaskLayer() { return aoiMaskLayer; },
        get aoiGraphic() { return aoiGraphic; },
        set aoiGraphic(v) { aoiGraphic = v; },
        get aoiSource() { return aoiSource; },
        get aoiSourcePlssTool() { return aoiSourcePlssTool; },
        get aoiSourceLayerTitle() { return aoiSourceLayerTitle; },
        get aoiLocationLabel() { return aoiLocationLabel; },
        get aoiBufferMiles() { return aoiBufferMiles; },
        get currentAoiMethod() { return currentAoiMethod; },
        get reportLayerViews() { return reportLayerViews; },
        get layerCfgByUrl() { return layerCfgByUrl; },
        get alwaysVisibleLayers() { return alwaysVisibleLayers; },
        get serviceStatus() { return serviceStatus; },
        get selectionLayers() { return selectionLayers; },
        get lastReportRowsByLayer() { return lastReportRowsByLayer; },
        get selectedPermitType() { return selectedPermitType; }
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
        getClipEnvelope, clipFeaturesToEnvelope,
        filterTouchingOnly, getReportGeometry, unionGeomsChunked,
        querySingleLayer, querySingleLayerChunked, computeElevationStats,
        computeLayerCoverageStats, buildPerFeatureTable,
        getAoiKey, resetCoverageCacheForAoi, SQM_PER_ACRE,
        preWarmReportLayers, preFireCoverageStats, clearPreWarmCache,
        sampleWithoutReplacement, makeTable
    } = queryEngine;

    // ── Initialize final-report module with shared state + deps ──
    const finalReport = finalReportModule.init(state, {
        mapUtils, queryEngine, ImageryLayer, FeatureLayer, geometryEngine,
        setStatus, finalReportStatus
    });
    const {
        openHtmlInNewTab, setCachedFinalReportHtml,
        saveReportToDb, loadReportFromDb, cleanupExpiredReports,
        buildReportInBackground, openCompletedReport, uploadAndOpenReport
    } = finalReport;

    const { isMobileBrowser } = mapUtils;

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
    let reportAbortController = null; // AbortController for in-flight network requests

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
    const isFinalReport = (tabName === "finalReport");

    // Panels
    if (tabLayersPanel) tabLayersPanel.classList.toggle("active", isLayers);
    if (tabServicesPanel) tabServicesPanel.classList.toggle("active", isServices);
    if (tabReportPanel) tabReportPanel.classList.toggle("active", isReport);
    if (tabReportFinalPanel) tabReportFinalPanel.classList.toggle("active", isFinalReport);

    // Buttons
    if (tabLayersBtn) tabLayersBtn.classList.toggle("active", isLayers);
    if (tabServicesBtn) tabServicesBtn.classList.toggle("active", isServices);
    if (tabReportBtn) tabReportBtn.classList.toggle("active", isReport);
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

    // ── Permit-type relevance for "Find Existing Permit or Lease" buttons ──
    // Maps each permitType key to the data-permit values that are relevant.
    // If a user selected "grazing", they only see allotment/pasture (not oil-gas, mining, etc.).
    const PERMIT_ITEM_RELEVANCE = {
        "oil-gas":  ["oilgas", "geothermal", "row", "lua"],
        "grazing":  ["allotment", "pasture", "lua"],
        "mining":   ["mining", "lua"],
        "row":      ["row", "oilgas", "lua", "allotment", "mining", "geothermal", "coal"],
        "realty":   ["oilgas", "row", "lua", "allotment", "mining", "geothermal", "coal"]
    };

    /**
     * Filter the permit-item buttons in the "Find Existing Permit" slide
     * to show only those relevant to the selected permit type.
     * Also hides group labels that have no visible items.
     */
    function filterPermitItemsBySelectedType() {
        const list = document.getElementById("wizPermitList");
        if (!list) return;
        const relevant = selectedPermitType ? PERMIT_ITEM_RELEVANCE[selectedPermitType] : null;

        // Show/hide individual permit-item buttons
        list.querySelectorAll(".permit-item").forEach(btn => {
            if (!relevant) {
                btn.style.display = "";
            } else {
                btn.style.display = relevant.includes(btn.dataset.permit) ? "" : "none";
            }
        });

        // Show/hide group labels — hide if all their following siblings are hidden
        list.querySelectorAll(".permit-group-label").forEach(label => {
            let hasVisible = false;
            let sibling = label.nextElementSibling;
            while (sibling && !sibling.classList.contains("permit-group-label")) {
                if (sibling.classList.contains("permit-item") && sibling.style.display !== "none") {
                    hasVisible = true;
                    break;
                }
                sibling = sibling.nextElementSibling;
            }
            label.style.display = hasVisible ? "" : "none";
        });
    }

    // ──────────────────────────────────────────────────────────────
    // PERMITTING MODE — Mode switching, wizard, AOI methods, permit-type-driven buckets
    // ──────────────────────────────────────────────────────────────

    // PERMIT_BUCKETS replaced — categorization logic now comes from permit-types.js.
    // The legacy categorizeIntoBuckets is still available as _categorizeIntoBucketsShared
    // for backward-compatible full-report (no permit type selected).

    // setAppMode — only "permit" mode is supported (advanced mode removed).
    function setAppMode() {
        currentAppMode = "permit";
        if (permitModePanel) permitModePanel.classList.remove("hidden");
        if (permitModeBtn) permitModeBtn.classList.add("active");
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
        // Auto-scroll: for step 3 (results), scroll to top of results.
        // For other steps we no longer auto-scroll — the panel already shows
        // the new step content at its natural position.
        if (step === 3) {
            const step3El = document.getElementById("wizardStep3");
            if (step3El) {
                setTimeout(() => {
                    step3El.scrollIntoView({ behavior: "smooth", block: "start" });
                }, 100);
            }
        }
    }

    // Swipe slide index mapping
    const aoiSlideMap = { methods: 0, search: 1, permit: 2, select: 3, draw: 4, upload: 5, namesearch: 6, confirm: 7 };

    function swipeToAoiSlide(slideIndex) {
        const track = document.getElementById("aoiSwipeTrack");
        if (track) {
            track.style.transform = `translateX(-${slideIndex * 100}%)`;
            // Mark the active slide for CSS height calculation
            const slides = track.querySelectorAll(".aoi-swipe-slide");
            slides.forEach((slide, i) => {
                slide.classList.toggle("active", i === slideIndex);
            });
        }
    }

    function showAoiMethod(method) {
        // Reset any existing AOI process before switching methods
        if (sketch) sketch.cancel();
        currentInteractionMode = "select";
        selectionGeom = null;
        aoiOriginalGeom = null;
        aoiSource = null;
        aoiSourceLayerTitle = null;
        aoiSourceLayerUrl = null;
        aoiSourceObjectId = null;
        aoiSourceObjectIdField = null;
        aoiSourceFeature = null;
        aoiLocationLabel = null;
        aoiBufferMiles = 0;
        if (aoiLayer) aoiLayer.removeAll();
        aoiGraphic = null;
        if (runBtn) runBtn.disabled = true;
        if (wizFullReport) wizFullReport.disabled = true;
        setStatus("");

        // Filter permit items when navigating to the permit slide
        if (method === "permit") filterPermitItemsBySelectedType();

        // Hide all selection layers (permit & PLSS) so they don't linger when switching methods
        for (let i = 0; i < selectionLayers.length; i++) disableSelectionLayer(i);
        clearAllSelectionToolButtons();

        currentAoiMethod = method;

        // Swipe to the method slide
        const slideIndex = aoiSlideMap[method] ?? 0;
        swipeToAoiSlide(slideIndex);

        if (method === "draw") {
            if (modeSelect) modeSelect.value = "draw";
            // Reset draw tool button selection state
            const drawToolBtns = document.querySelectorAll(".draw-tool-btn");
            drawToolBtns.forEach(b => b.classList.remove("active"));
            const drawHintEl = document.getElementById("drawHint");
            if (drawHintEl) drawHintEl.textContent = "Select a draw type above to begin.";
            setMode("draw");
        } else if (method === "select" || method === "permit") {
            if (modeSelect) modeSelect.value = "select";
            setMode("select");
        } else if (method === "namesearch") {
            // Focus the search input when slide opens
            const searchInput = document.getElementById("featureSearchInput");
            if (searchInput) setTimeout(() => searchInput.focus(), 400);
        }
    }

    function hideAoiMethodPanels() {
        // Swipe back to the methods selection slide
        swipeToAoiSlide(0);
        currentAoiMethod = null;

        // Stop any active sketch and restore click-to-select
        if (sketch) sketch.cancel();
        currentInteractionMode = "select";

        // Hide all selection layers
        for (let i = 0; i < selectionLayers.length; i++) disableSelectionLayer(i);
        clearAllSelectionToolButtons();

        // Hide buffer panels (upload)
        const ubp = document.getElementById("uploadBufferPanel");
        if (ubp) ubp.classList.add("hidden");
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
                if (el) el.textContent = formatNumber(acres, 0) + " acres";

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
                                <div class="warn-detail">For faster results, consider selecting a smaller area.</div>
                            </div>`;
                        warningEl.classList.remove("hidden");
                        scrollPanelToBottom(warningEl);
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
        const sourceEl = document.getElementById("wizAoiLocation");
        if (sourceEl) {
            sourceEl.textContent = "…";
            // Reverse-geocode the AOI centroid for a human-readable location
            try {
                const center = selectionGeom.extent
                    ? selectionGeom.extent.center
                    : selectionGeom;
                // Project to lon/lat if needed
                let lon, lat;
                const sr = center.spatialReference;
                if (sr && (sr.isWebMercator || sr.wkid === 102100 || sr.wkid === 3857)) {
                    // Web Mercator → WGS84
                    const DEG = 180 / Math.PI;
                    lon = (center.x / 6378137) * DEG;
                    lat = (Math.atan(Math.exp(center.y / 6378137)) * 2 - Math.PI / 2) * DEG;
                } else if (sr && sr.wkid === 4326) {
                    lon = center.x;
                    lat = center.y;
                } else {
                    lon = center.x;
                    lat = center.y;
                }
                const revUrl = `https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/reverseGeocode?location=${lon},${lat}&outSR=4326&f=json`;
                fetch(revUrl).then(r => r.json()).then(data => {
                    if (!data || !data.address) { sourceEl.textContent = `${lat.toFixed(4)}°, ${lon.toFixed(4)}°`; return; }
                    const a = data.address;
                    const parts = [a.Subregion || a.City || "", a.Region || ""].filter(Boolean);
                    const locLabel = parts.length ? parts.join(", ") : `${lat.toFixed(4)}°, ${lon.toFixed(4)}°`;
                    aoiLocationLabel = locLabel;
                    sourceEl.textContent = locLabel;
                }).catch(() => {
                    const fallback = `${lat.toFixed(4)}°, ${lon.toFixed(4)}°`;
                    aoiLocationLabel = fallback;
                    sourceEl.textContent = fallback;
                });
            } catch (e) {
                sourceEl.textContent = "—";
            }
        }

        // Reset the confirm-section buffer panel
        const confirmBufInput = document.getElementById("aoiConfirmBufferInput");
        const confirmBufReset = document.getElementById("aoiConfirmBufferReset");
        if (confirmBufInput) confirmBufInput.value = "";
        if (confirmBufReset) confirmBufReset.classList.add("hidden");
        document.querySelectorAll(".confirm-buf-preset").forEach(b => b.classList.remove("active"));
    }

    let wizLocationDebounce = null;
    let _searchAbort = null;
    function performLocationSearch(query) {
        if (!query || query.length < 3) {
            if (wizLocationResults) { wizLocationResults.innerHTML = ""; wizLocationResults.classList.add("hidden"); }
            if (_searchAbort) { _searchAbort.abort(); _searchAbort = null; }
            return;
        }
        clearTimeout(wizLocationDebounce);
        wizLocationDebounce = setTimeout(async () => {
            if (_searchAbort) _searchAbort.abort();
            _searchAbort = new AbortController();
            const signal = _searchAbort.signal;
            try {
                const geocodeBase = config.referenceLayers?.geocodeService || "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer";
                const url = geocodeBase + "/suggest?text=" + encodeURIComponent(query) + "&maxSuggestions=5&f=json";
                const resp = await fetch(url, { signal });
                const data = await resp.json();
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
                                const gUrl = (config.referenceLayers?.geocodeService || "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer") + "/findAddressCandidates?SingleLine=" + encodeURIComponent(txt) + "&magicKey=" + encodeURIComponent(mk) + "&outSR=4326&f=json";
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



    /**
     * Populate screening results — dynamically builds accordion sections
     * based on the currently-selected permit type's groups.
     * Each category is a collapsible section the user can tap to expand/collapse.
     */
    function populatePermitResults() {
        const ptKey = selectedPermitType;
        const ptDef = ptKey ? PERMIT_TYPES[ptKey] : null;
        const groups = ptDef ? ptDef.groups : [];

        // Categorize using permit-type-aware function
        const categorized = ptKey
            ? categorizeByPermitType(lastReportRowsByLayer, ptKey, layerCfgByUrl)
            : null;

        // ── Build dynamic accordion DOM ──
        const container = document.getElementById("bucketAccordion");
        if (!container) return;

        let html = '';

        // Disclaimer at top
        html += '<div class="overview-disclaimer"><strong>Important:</strong> These results show which layers have features intersecting your project area. Generate a report to see detailed analysis. Contact your local BLM field office for authoritative guidance.</div>';

        // One accordion section per group
        for (const g of groups) {
            const items = (categorized && categorized.groups[g.key]) || [];
            const layersWithFeatures = items.filter(it => it.hasCoverage).length;
            const totalLayers = items.length;
            const statusClass = layersWithFeatures > 0 ? 'findings' : 'clear';
            const statusLabel = layersWithFeatures > 0
                ? (layersWithFeatures + ' of ' + totalLayers + ' layer' + (totalLayers !== 1 ? 's' : ''))
                : 'No features';

            html += '<div class="bucket-accordion-section" data-bucket="' + g.key + '">';

            // Header row (reuses overview-cat-row styling)
            html += '<button class="overview-cat-row ' + statusClass + '" type="button" aria-expanded="false">';
            html += '<span class="overview-cat-indicator"></span>';
            html += '<span class="overview-cat-icon">' + g.icon + '</span>';
            html += '<span class="overview-cat-label">' + escapeHtml(g.label) + '</span>';
            html += '<span class="overview-cat-status">' + statusLabel + '</span>';
            html += '<span class="overview-cat-arrow">›</span>';
            html += '</button>';

            // Collapsible body
            html += '<div class="bucket-accordion-body">';
            html += '<div class="bucket-card">';
            if (g.description) html += '<div class="bucket-card-desc">' + escapeHtml(g.description) + '</div>';
            html += '<ul class="bucket-layer-list">';
            for (const it of items) {
                const hasCov = it.hasCoverage;
                html += '<li class="bucket-layer-item"><span class="bucket-layer-name">' + escapeHtml(it.title) + '</span>';
                html += '<span class="bucket-layer-count' + (hasCov ? ' has-hits' : '') + '">' + (hasCov ? '✓' : '—') + '</span></li>';
            }
            if (!items.length) html += '<li class="bucket-layer-item" style="color:var(--text-muted);font-style:italic;">No layers in this group</li>';
            html += '</ul>';
            if (layersWithFeatures === 0) html += '<div class="hint" style="margin-top:8px;">No features found in this group for your project area.</div>';
            html += '</div></div>'; // close bucket-card + accordion-body

            html += '</div>'; // close accordion-section
        }

        // All Data accordion section
        {
            const tl = lastReportRowsByLayer.length;
            const lwh = lastReportRowsByLayer.filter(x => x.hasCoverage).length;
            html += '<div class="bucket-accordion-section" data-bucket="all-data">';
            html += '<button class="overview-cat-row all-data" type="button" aria-expanded="false">';
            html += '<span class="overview-cat-indicator"></span>';
            html += '<span class="overview-cat-icon">📊</span>';
            html += '<span class="overview-cat-label">All Layers</span>';
            html += '<span class="overview-cat-status">' + lwh + ' of ' + tl + ' layer' + (tl !== 1 ? 's' : '') + '</span>';
            html += '<span class="overview-cat-arrow">›</span>';
            html += '</button>';

            html += '<div class="bucket-accordion-body">';
            html += '<div class="bucket-card"><ul class="bucket-layer-list">';
            const sorted = lastReportRowsByLayer.slice().sort((a, b) => (b.hasCoverage ? 1 : 0) - (a.hasCoverage ? 1 : 0));
            for (const it of sorted) {
                const hasCov = it.hasCoverage;
                html += '<li class="bucket-layer-item"><span class="bucket-layer-name">' + escapeHtml(it.title) + '</span>';
                html += '<span class="bucket-layer-count' + (hasCov ? ' has-hits' : ' zero') + '">' + (hasCov ? '✓' : '—') + '</span></li>';
            }
            html += '</ul></div></div>'; // close bucket-card + accordion-body
            html += '</div>'; // close accordion-section
        }

        container.innerHTML = html;

        // Wire up accordion header clicks
        container.querySelectorAll('.bucket-accordion-section > .overview-cat-row').forEach(btn => {
            btn.addEventListener('click', () => {
                const section = btn.closest('.bucket-accordion-section');
                section.classList.toggle('open');
                btn.setAttribute('aria-expanded', section.classList.contains('open'));
            });
        });

        // Summary in header card — also show count of additional layers queried during report
        const summaryEl = document.getElementById("permitResultsSummary");
        if (summaryEl) {
            const lwc = lastReportRowsByLayer.filter(x => x.hasCoverage).length;
            const typeLabel = ptDef ? ptDef.label : 'All';
            // Count additional (non-core) layers tagged for this permit type
            const allLayers = config.reportLayers || [];
            const screenedUrls = new Set(lastReportRowsByLayer.map(r => String(r.url || '').replace(/\/+$/, '')));
            const additionalCount = ptKey
                ? allLayers.filter(l => {
                    const pts = l.permitTypes || [];
                    return pts.includes(ptKey) && !pts.includes("core")
                        && !screenedUrls.has(String(l.url || '').replace(/\/+$/, ''));
                }).length
                : 0;
            let summaryHtml = '<div class="small"><strong>' + typeLabel + '</strong> screening: <strong>' + lwc + '</strong> of <strong>' + lastReportRowsByLayer.length + '</strong> core layers have features in your project area.';
            if (additionalCount > 0) {
                summaryHtml += '<br><span style="color:var(--accent);">+' + additionalCount + ' additional ' + escapeHtml(typeLabel) + '-specific layer' + (additionalCount !== 1 ? 's' : '') + ' will be analyzed in the report.</span>';
            }
            summaryHtml += '</div>';
            summaryEl.innerHTML = summaryHtml;
        }
    }

    /** Toggle an accordion section open/closed by bucket key */
    function toggleAccordionSection(bucketKey) {
        const container = document.getElementById("bucketAccordion");
        if (!container) return;
        const section = container.querySelector('.bucket-accordion-section[data-bucket="' + bucketKey + '"]');
        if (section) section.classList.toggle('open');
    }

    // Cache for generated permit-type reports.
    // On mobile: stores IndexedDB report IDs (strings) to avoid memory pressure.
    // On desktop: stores full HTML strings for instant re-open.
    const cachedPermitReports = {};

    /**
     * Open a report from cache — loads from IndexedDB if the cache holds
     * just a report ID (mobile), or opens directly from HTML (desktop).
     */
    async function openCachedReport(cacheEntry) {
        if (!cacheEntry) return false;
        // If it's an R2 URL, just open it — zero local memory needed
        if (typeof cacheEntry === 'string' && cacheEntry.startsWith('http')) {
            window.open(cacheEntry, '_blank', 'noopener');
            return true;
        }
        // If it's a short string, it's an IDB report ID
        if (typeof cacheEntry === 'string' && cacheEntry.length < 30) {
            try {
                const html = await loadReportFromDb(cacheEntry);
                if (html) return openCompletedReport(html);
            } catch (e) {
                console.warn('[report] Failed to load report from IndexedDB:', e);
            }
            return false;
        }
        // Desktop path: it's the full HTML string
        return openCompletedReport(cacheEntry);
    }

    /**
     * Generate the report for the active permit type.
     * Only layers tagged for the selected permit type are included.
     * Shows modal during generation, then button changes to "View" state.
     */
    async function generateFullProgressiveReport() {
        console.log("[report] generateFullProgressiveReport called, selectionGeom:", !!selectionGeom,
            "rows:", lastReportRowsByLayer?.length, "permitType:", selectedPermitType);
        if (!selectionGeom) {
            setStatus("No AOI selected — cannot generate report");
            return;
        }
        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            setStatus("Run analysis first before generating reports");
            return;
        }

        const ptKey = selectedPermitType || '__all__';
        const ptDef = selectedPermitType ? PERMIT_TYPES[selectedPermitType] : null;
        const ptLabel = ptDef ? ptDef.label : 'Full';
        
        // Check if report is already ready to view
        if (wizFullReport && wizFullReport.dataset.reportReady === 'true' && cachedPermitReports[ptKey]) {
            const opened = await openCachedReport(cachedPermitReports[ptKey]);
            if (!opened) {
                console.warn('[report] openCachedReport returned false — popup may have been blocked');
            }
            return;
        }

        // Show modal
        reportModal.show();
        reportModal.setStep(`Building ${ptLabel} Report`);
        console.log(`[report] Starting ${ptLabel} report generation…`);
        
        const _isMobile = isMobileBrowser();
        let reportFinalStats = {};
        try {
            let htmlContent = await buildReportInBackground({
                bucketKey: null,
                permitTypeKey: selectedPermitType || null,
                onProgress: (pct, maps, sections, stats) => {
                    reportModal.setProgress(pct);
                    reportModal.updateStats(maps);
                    if (stats) reportFinalStats = stats;
                },
                onStep: (stepText) => {
                    reportModal.setProgressDetail(stepText);
                },
                isCanceled: () => reportModal.isCanceled()
            });
            
            if (reportModal.isCanceled()) {
                setStatus("Report canceled");
                return;
            }

            // On mobile, save to IndexedDB as a fallback, then try
            // uploading to R2 so we can open by URL with zero local
            // memory.  Null out htmlContent ASAP to reduce pressure.
            let reportRef = htmlContent; // desktop: keep HTML in memory
            if (_isMobile) {
                let idbId = null;
                try {
                    reportModal.setProgressDetail('Saving report…');
                    idbId = await saveReportToDb(htmlContent);
                    console.log(`[report] Saved to IndexedDB as "${idbId}"`);
                } catch (idbErr) {
                    console.warn('[report] IndexedDB save failed:', idbErr);
                }

                // Try uploading to R2 — opens via remote URL, no local Blob
                try {
                    reportModal.setProgressDetail('Preparing report link…');
                    const r2Url = await uploadAndOpenReport(htmlContent, false);
                    if (r2Url) {
                        reportRef = r2Url;
                        console.log('[report] Uploaded to R2 — will open via URL');
                    } else if (idbId) {
                        reportRef = idbId;
                    }
                } catch (uploadErr) {
                    console.warn('[report] R2 upload failed, falling back to IDB:', uploadErr);
                    if (idbId) reportRef = idbId;
                }

                // Release the large HTML string from memory
                htmlContent = null;
            }
            
            // Show success overlay with stats
            reportModal._onViewReport = async () => {
                const opened = await openCachedReport(reportRef);
                if (!opened) {
                    console.warn('[report] openCachedReport returned false — popup may have been blocked');
                }
            };
            reportModal.showSuccess({
                mapsGenerated: reportFinalStats.mapsGenerated ?? 0,
                layersQueried: reportFinalStats.layersQueried ?? 0,
                featuresFound: reportFinalStats.totalFeatures ?? 0
            });
            
            console.log(`[report] ${ptLabel} report generated successfully`);
            
            // Cache the report reference and update button to "View" state
            cachedPermitReports[ptKey] = reportRef;
            
            if (wizFullReport) {
                wizFullReport.dataset.reportReady = 'true';
                wizFullReport.innerHTML = '📄 View ' + escapeHtml(ptLabel) + ' Report';
                wizFullReport.classList.add('ready-to-view');
            }
            
            setStatus(`${ptLabel} report ready`);
            
        } catch (e) {
            console.error("[report] Report generation failed:", e);
            reportModal.hide();
            if (e.message === "Canceled") {
                setStatus("Report canceled");
            } else {
                setStatus("Report generation failed — see console");
                alert("Report generation failed: " + e.message);
            }
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
    aoiOriginalGeom = null;
    aoiSourceObjectId = null;
    aoiSourceObjectIdField = null;
    aoiSource = null;
    aoiSourceLayerTitle = null;
    aoiSourceLayerUrl = null;
    aoiSourceFeature = null;
    aoiSourcePlssTool = null;

    // Reset permit type selection
    selectedPermitType = null;
    document.querySelectorAll('.permit-type-card').forEach(c => c.classList.remove('selected'));
    const badgeEl = document.getElementById("wizStep2PermitBadge");
    if (badgeEl) badgeEl.textContent = "";

    // Hide AOI confirm section (swipe handles this now)
    
    // Clear results
    if (resultsEl) resultsEl.innerHTML = "";
    if (exportAllBtn) exportAllBtn.disabled = true;
    lastReportRowsByLayer = [];
    
    // Clear final report
    setCachedFinalReportHtml(null);
    if (viewReportBtn) viewReportBtn.disabled = true;

    // Clear cached permit reports (prevents stale HTML from accumulating in memory)
    Object.keys(cachedPermitReports).forEach(k => delete cachedPermitReports[k]);

    if (aoiLayer) aoiLayer.removeAll();
    aoiGraphic = null;

    if (runBtn) runBtn.disabled = true;
    setStatus("cleared");
    resetCoverageCacheForAoi(null);
    clearPreWarmCache();
    setBusy(false);

    // Reset wizard-specific UI state
    if (wizFullReport) {
        wizFullReport.disabled = true;
        delete wizFullReport.dataset.reportReady;
        wizFullReport.innerHTML = '📋 Generate Report';
        wizFullReport.classList.remove('ready-to-view');
    }
    setWizPlssActive(null);

    // Clear wizard location search
    clearTimeout(wizLocationDebounce);
    wizLocationDebounce = null;
    if (wizLocationInput) wizLocationInput.value = "";
    if (wizLocationResults) { wizLocationResults.innerHTML = ""; wizLocationResults.classList.add("hidden"); }

    // Clear dynamic accordion sections (step 3) — rebuilt per screening
    const bucketAccordion = document.getElementById("bucketAccordion");
    if (bucketAccordion) {
        bucketAccordion.innerHTML = "";
    }

    const summaryEl = document.getElementById("permitResultsSummary");
    if (summaryEl) summaryEl.innerHTML = "";

    // Clear upload status
    const uploadStatusEl = document.getElementById("uploadStatus");
    if (uploadStatusEl) uploadStatusEl.classList.add("hidden");

    // Hide buffer panels (upload + confirm)
    const uploadBufPanelEl = document.getElementById("uploadBufferPanel");
    if (uploadBufPanelEl) uploadBufPanelEl.classList.add("hidden");

    // Hide AOI mask
    if (typeof hideAoiMask === "function") hideAoiMask();
}

    /**
     * Count total vertices across all rings/paths of a geometry.
     */
    function countVertices(geom) {
        if (!geom) return 0;
        const rings = geom.rings || geom.paths || [];
        let count = 0;
        for (let i = 0; i < rings.length; i++) count += rings[i].length;
        return count;
    }

    /**
     * If the AOI polygon exceeds the configured vertex threshold, simplify it
     * using geometryEngine.generalize to keep query payloads small.
     * Returns the (possibly simplified) geometry.
     */
    function generalizeAoiIfNeeded(geom) {
        if (!geom || geom.type !== "polygon") return geom;
        const maxVerts = config?.report?.aoiMaxVertices ?? 500;
        const maxDev   = config?.report?.aoiGeneralizeMaxDeviation ?? 10; // meters (Web Mercator)
        const before = countVertices(geom);
        if (before <= maxVerts) return geom;

        try {
            const simplified = geometryEngine.generalize(geom, maxDev, true, "meters");
            const after = countVertices(simplified);
            console.log(`[AOI] Generalized polygon: ${before} → ${after} vertices (maxDev=${maxDev}m)`);
            // If generalize collapsed the geometry, keep original
            if (!simplified || countVertices(simplified) < 3) {
                console.warn("[AOI] Generalize produced degenerate geometry — keeping original");
                return geom;
            }
            return simplified;
        } catch (e) {
            console.warn("[AOI] Generalize failed — keeping original", e);
            return geom;
        }
    }

    function setGeometryFromSelection(geom) {
        selectionGeom = generalizeAoiIfNeeded(geom) || null;
        if (runBtn) runBtn.disabled = !selectionGeom;

        // Permitting mode: swipe to the AOI confirm slide when AOI is defined
        if (selectionGeom && currentAppMode === "permit" && currentWizardStep === 2) {
            populateAoiConfirmation();
            swipeToAoiSlide(aoiSlideMap.confirm);
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
            // Don't auto-start drawing - wait for user to select a draw tool
            setStatus("Select a draw type to begin");
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
                : (status === "WARN")
                ? `<span class="status-degraded" title="Service degraded">⚡</span>`
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
                : (status === "WARN")
                ? `<span class="status-degraded" title="Service degraded">⚡</span>`
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
                : (status === "WARN")
                ? `<span class="status-degraded" title="Service degraded">⚡</span>`
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
                : (status === "WARN")
                ? `<span class="status-degraded" title="Service degraded">⚡</span>`
                : "";
            return `
                <div class="toggle-row">
                    <input type="checkbox" id="lm_rptlayer_${i}" ${checked} />
                    <span class="layer-swatch layer-swatch-report" aria-hidden="true"></span>
                    <label class="toggle-name" for="lm_rptlayer_${i}">${statusIcon}${escapeHtml(l.title)}</label>
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
            const rect = layerManagerWindow.getBoundingClientRect();
            let newLeft = origLeft + dx;
            let newTop  = origTop + dy;
            newLeft = Math.max(0, Math.min(newLeft, window.innerWidth - rect.width));
            newTop  = Math.max(0, Math.min(newTop,  window.innerHeight - rect.height));
            layerManagerWindow.style.left = newLeft + "px";
            layerManagerWindow.style.top  = newTop + "px";
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
            e.preventDefault(); // prevent page scroll while dragging layer manager
            const touch = e.touches[0];
            const dx = touch.clientX - startX;
            const dy = touch.clientY - startY;
            const rect = layerManagerWindow.getBoundingClientRect();
            let newLeft = origLeft + dx;
            let newTop  = origTop + dy;
            newLeft = Math.max(0, Math.min(newLeft, window.innerWidth - rect.width));
            newTop  = Math.max(0, Math.min(newTop,  window.innerHeight - rect.height));
            layerManagerWindow.style.left = newLeft + "px";
            layerManagerWindow.style.top  = newTop + "px";
            layerManagerWindow.style.right = "auto";
        }, { passive: false });

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

    // Feedback window open/close & submit
    if (feedbackBtn) {
        feedbackBtn.addEventListener("click", () => {
            if (!feedbackWindow) return;
            feedbackWindow.classList.toggle("hidden");
            if (!feedbackWindow.classList.contains("hidden")) {
                feedbackStatus.textContent = "";
                feedbackStatus.className = "feedback-status";
            }
        });
    }
    if (feedbackCloseBtn) {
        feedbackCloseBtn.addEventListener("click", () => {
            if (feedbackWindow) feedbackWindow.classList.add("hidden");
        });
    }
    if (feedbackForm) {
        feedbackForm.addEventListener("submit", async (e) => {
            e.preventDefault();
            const typeEl = document.getElementById("feedbackType");
            const descEl = document.getElementById("feedbackDesc");
            const submitBtn = document.getElementById("feedbackSubmitBtn");
            const desc = (descEl.value || "").trim();
            if (!desc) return;

            feedbackStatus.textContent = "Submitting…";
            feedbackStatus.className = "feedback-status info";
            submitBtn.disabled = true;

            try {
                const workerUrl = config?.metadataWorkerUrl;
                if (!workerUrl) throw new Error("Worker URL not configured");
                const res = await fetch(workerUrl + "/feedback", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ type: typeEl.value, description: desc }),
                });
                const data = await res.json();
                if (res.ok && data.ok) {
                    feedbackStatus.textContent = "Thank you! Your feedback has been submitted.";
                    feedbackStatus.className = "feedback-status success";
                    descEl.value = "";
                } else {
                    feedbackStatus.textContent = data.error || "Submission failed. Please try again.";
                    feedbackStatus.className = "feedback-status error";
                }
            } catch (err) {
                console.error("Feedback submit error:", err);
                feedbackStatus.textContent = "Network error. Please try again.";
                feedbackStatus.className = "feedback-status error";
            } finally {
                submitBtn.disabled = false;
            }
        });
    }

    // ---------- Services tab ----------


async function checkServiceStatusBackground() {
    const items = getConfiguredServices(config);
    const timeoutMs = config?.services?.timeoutMs ?? 8000;

    // ── Try R2 metadata cache first (single fetch vs N pings) ──
    const workerUrl = config?.metadataWorkerUrl;
    if (workerUrl) {
        try {
            const cached = await fetchJsonWithTimeout(workerUrl + "/metadata", 6000, { noCache: true });
            if (cached && cached.layers) {
                for (const [url, meta] of Object.entries(cached.layers)) {
                    serviceStatus.set(url, meta.status || "DOWN");
                    if (meta.serviceDescription) {
                        serviceStatus.set(url + "::desc", meta.serviceDescription);
                    }
                    const hasFeatureHistory = meta && typeof meta.normallyHasFeatures === "boolean"
                        ? meta.normallyHasFeatures
                        : null;
                    serviceStatus.set(url + "::normallyHasFeatures", hasFeatureHistory);
                }

                // Recheck cached DOWN statuses in-browser to avoid worker-only false downs
                const cachedDownItems = items.filter(it => serviceStatus.get(it.url) === "DOWN");
                if (cachedDownItems.length) {
                    const rechecks = await Promise.allSettled(
                        cachedDownItems.map(async (it) => {
                            const pjsonUrl = normalizePjsonUrl(it.url);
                            try {
                                await fetchJsonWithTimeout(pjsonUrl, timeoutMs, { noCache: true });
                                return { url: it.url, status: "UP" };
                            } catch (e) {
                                return { url: it.url, status: "DOWN" };
                            }
                        })
                    );

                    let recovered = 0;
                    for (const r of rechecks) {
                        if (r.status === "fulfilled") {
                            const prev = serviceStatus.get(r.value.url);
                            serviceStatus.set(r.value.url, r.value.status);
                            if (prev === "DOWN" && r.value.status === "UP") recovered++;
                        }
                    }
                    if (recovered > 0) {
                        console.log(`[metadata-cache] Recovered ${recovered} cached DOWN service(s) via live browser recheck`);
                    }
                }

                renderLayerToggles(map);
                maybeShowDownServiceWarning(items, "r2", cached.lastRefresh || null);
                console.log(`[metadata-cache] Loaded ${Object.keys(cached.layers).length} layers from R2 (refreshed ${cached.lastRefresh})`);
                return; // cache hit — skip direct pings
            }
        } catch (e) {
            console.warn("[metadata-cache] R2 cache unavailable, falling back to direct pings:", e.message);
        }
    }

    // ── Fallback: parallel direct pings ──
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
    maybeShowDownServiceWarning(items, "fallback", null);
}

/**
 * Pre-expand FeatureServer/MapServer roots in the background so
 * the sublayer list is cached before the user runs analysis.
 */
function preExpandServiceRoots() {
    const reportLayers = config.reportLayers || [];
    for (const cfg of reportLayers) {
        const url = String(cfg.url || "");
        if (isFeatureServerRoot(url)) {
            const key = "fs:" + url;
            if (!_sublayerCache.has(key)) {
                _sublayerCache.set(key, expandServiceToSublayers(url).catch(() => []));
            }
        } else if (isMapServerRoot(url)) {
            const key = "ms:" + url;
            if (!_sublayerCache.has(key)) {
                _sublayerCache.set(key, expandMapServerToSublayers(url, { polygonOnly: false }).catch(() => []));
            }
        }
    }
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

    // Create a new AbortController for this analysis run
    if (reportAbortController) reportAbortController.abort();
    reportAbortController = new AbortController();

    const reportGeom = getReportGeometry();
    if (!reportGeom) {
        endReportOp(myOp);
        // Show the modal briefly with an error so the user gets feedback
        analysisModal.show();
        analysisModal.setStep("No Area of Interest");
        analysisModal.addLog("Cannot run analysis — no AOI geometry found. Please draw, search, or upload an Area of Interest first.", "error");
        setTimeout(() => analysisModal.hide(), 4000);
        return;
    }

    const analysisStartTime = Date.now();

    // ✅ Show analysis modal
    analysisModal.show();
    analysisModal.setProgress(0);
    analysisModal.setStep("Starting screening...");
    const totalLayers = getTotalLayerCount();
    analysisModal.addLog(`Screening ${totalLayers} datasets for coverage`);

    setBusy(true);

    let layersQueried = 0;
    let featuresFound = 0;

    try {
        // Step 1: Data Check (10% progress)
        analysisModal.setStep("Step 1/2: Checking services...");
        analysisModal.setProgress(10);
        analysisModal.addLog("Checking service availability");
        
        // Skip full health check if background check already populated status
        if (serviceStatus.size > 0) {
            analysisModal.addLog("Using cached service status (background check already ran)", "success");
        } else {
            await refreshServicesTab();
        }

        if (isReportCanceled(myOp)) {
            analysisModal.addLog("Analysis canceled by user", "error");
            analysisModal.hide();
            setStatus("canceled");
            return;
        }

        analysisModal.addLog("Service check complete", "success");
        analysisModal.setProgress(25);

        // Step 2: Query all layers for coverage (25% → 100% progress)
        analysisModal.setStep("Step 2/2: Checking layer coverage...");
        analysisModal.addLog("Checking which datasets cover your project area");

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
        const layersWithFeatures = lastReportRowsByLayer.filter(x => x.hasCoverage).length;
        analysisModal.updateStats(layersQueried, layersWithFeatures, 0);
        analysisModal.addLog(`Found ${layersWithFeatures} of ${layersQueried} layers in your project area`, "success");
        analysisModal.setProgress(100);

        // ─────────────────────────────────────────────────────────────────
        // Screening checks for layer coverage (at least 1 feature intersects).
        // Full feature data is retrieved on-demand when generating reports.
        // ─────────────────────────────────────────────────────────────────

        setStatus("Screening complete!");
        
        // ✅ Show success animation
        analysisModal.showSuccess(layersQueried, layersWithFeatures, 0, Date.now() - analysisStartTime);

        // Permitting mode: populate results and enable report button
        if (currentAppMode === "permit") {
            populatePermitResults();
            if (wizFullReport) {
                const ptDef = selectedPermitType ? PERMIT_TYPES[selectedPermitType] : null;
                const ptLabel = ptDef ? ptDef.label : 'Full';
                wizFullReport.disabled = false;
                wizFullReport.innerHTML = '📋 Generate ' + escapeHtml(ptLabel) + ' Report';
            }
            // Advance to step 3 (results)
            goToWizardStep(3);
        }

        // Enable "View Report" button (for Advanced mode, if ever re-enabled)
        if (viewReportBtn) viewReportBtn.disabled = false;

        // ── Background pre-warm for faster report generation ──
        // Silently pre-load FeatureLayer metadata + pre-fire polygon
        // coverage stats while the user reviews screening results.
        // Everything is fire-and-forget — errors are swallowed.
        if (lastReportRowsByLayer.length > 0) {
            const preWarmGeom = getReportGeometry();
            preWarmReportLayers(lastReportRowsByLayer)
                .then(() => preFireCoverageStats(lastReportRowsByLayer, preWarmGeom))
                .catch(() => {});
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
        if (reportAbortController) {
            reportAbortController = null;
        }
    }
}


// Extracted query logic - checks layer coverage (extent intersection only)
async function queryAllLayers(reportGeom, myOp, modal = null) {
    const abortSignal = reportAbortController ? reportAbortController.signal : null;
    if (resultsEl) resultsEl.innerHTML = "";
    if (exportAllBtn) exportAllBtn.disabled = true;
    lastReportRowsByLayer = [];

    // Phase 1 screening: query ONLY "core" layers for fast coverage checks.
    // Permit-type-specific layers are queried later during report generation.
    let reportLayerPool = config.reportLayers || [];
    if (selectedPermitType) {
        reportLayerPool = reportLayerPool.filter(l => {
            const pts = l.permitTypes || [];
            return pts.includes("core");
        });
    }

    const combinedCfgs = [
        ...reportLayerPool
    ];

    // PLSS State Boundaries and Parcels are NOT included in screening —
    // the report generates its own PLSS/state/county context layers independently.

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

    // Expand all service roots in parallel (with cache)
    const [fsResults, msResults] = await Promise.all([
        Promise.allSettled(featureServerRoots.map(async cfg => {
            try {
                // Use cached expansion if available
                const cacheKey = "fs:" + cfg.url;
                if (!_sublayerCache.has(cacheKey)) {
                    _sublayerCache.set(cacheKey, expandServiceToSublayers(cfg.url));
                }
                const sublayers = await _sublayerCache.get(cacheKey);
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
                // Use cached expansion if available
                const cacheKey = "ms:" + cfg.url;
                if (!_sublayerCache.has(cacheKey)) {
                    _sublayerCache.set(cacheKey, expandMapServerToSublayers(cfg.url, { polygonOnly: false }));
                }
                const subs = await _sublayerCache.get(cacheKey);
                return subs.map(sl => ({
                    title: `${cfg.title}: ${sl.title}`,
                    url: sl.url
                }));
            } catch (e) {
                return [{ title: `${cfg.title} (FAILED to expand)`, url: cfg.url, error: e }];
            }
        }))
    ]);

    // Early exit if canceled during service expansion
    if (isReportCanceled(myOp)) return;

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

        // 1b. Skip known-DOWN layers (from background health check) to avoid 30s timeouts
        const cachedStatus = serviceStatus.get(t.url);
        if (cachedStatus === "DOWN") {
            console.log(`Skipping ${t.title} — marked DOWN by health check`);
            if (modal) modal.addLog(`${t.title}: skipped (service unavailable)`, "warning");
            return {
                card: `
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(t.title)}</div>
              <div class="badge" style="background: var(--warning-bg, #f59e0b); color: #fff;">skipped (down)</div>
            </div>
            <div class="small mono">
              <a href="${escapeHtml(t.url)}" target="_blank" rel="noopener">Service URL</a>
            </div>
            <div class="small" style="margin-top:4px; color: var(--muted);">Service was unavailable at check time — excluded from analysis</div>
          </div>`,
                reportEntry: {
                    title: t.title,
                    url: t.url,
                    hasCoverage: false,
                    count: 0,
                    rows: [],
                    _layer: null,
                    _exportQuery: null,
                    fullRows: null,
                    __skippedDown: true
                }
            };
        }

        // 2. ImageServer layers — fetch metadata
        if (t.__isImageService) {
            if (isReportCanceled(myOp)) return null;
            const metaUrl = `${t.url}?f=json`;
            const resp = await fetch(metaUrl, { signal: abortSignal });
            if (!resp.ok) throw new Error(`ImageServer metadata fetch failed: HTTP ${resp.status}`);
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
                    hasCoverage: true, // ImageServer layers assumed to have coverage
                    count: 0,
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

            return {
                card: `
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(t.title)}</div>
              <div class="badge has-hits">AOI source</div>
            </div>
            <div class="small mono">
              <a href="${escapeHtml(t.url)}" target="_blank" rel="noopener">Service URL</a>
            </div>
          </div>`,
                reportEntry: {
                    title: t.title,
                    url: t.url,
                    hasCoverage: true,
                    count: feats.length,
                    rows,
                    _layer: null,
                    _exportQuery: null,
                    fullRows: rows
                }
            };
        }

        // 4. Regular feature layer — quick check if any features intersect AOI
        // Full feature queries are deferred to report generation
        if (isReportCanceled(myOp)) return null;

        let hasCoverage = false;
        let layerRef = null;
        let timedOut = false;
        const COVERAGE_TIMEOUT_MS = config.report?.coverageTimeoutMs ?? 30000;
        
        try {
            layerRef = queryEngine.getCachedLayer(t.url);
            await layerRef.load({ signal: abortSignal });
            
            if (isReportCanceled(myOp)) return null;

            // Quick check: query for just 1 feature to see if any intersect
            const checkQuery = layerRef.createQuery();
            checkQuery.geometry = reportGeom;
            checkQuery.spatialRelationship = "intersects";
            checkQuery.returnGeometry = false;
            checkQuery.num = 1; // Only need to find 1 feature to confirm coverage
            checkQuery.outFields = [layerRef.objectIdField || "OBJECTID"];
            
            // Race the query against a per-layer timeout so a slow service
            // (e.g. USFWS Critical Habitat with complex geometries) can't
            // stall the entire analysis batch indefinitely.
            let coverageTimer;
            const result = await Promise.race([
                layerRef.queryFeatures(checkQuery, { signal: abortSignal }).finally(() => clearTimeout(coverageTimer)),
                new Promise((_, reject) => {
                    coverageTimer = setTimeout(() => reject(new Error("__coverageTimeout__")), COVERAGE_TIMEOUT_MS);
                })
            ]);
            hasCoverage = result.features && result.features.length > 0;
        } catch (e) {
            if (e.name === "AbortError" || isReportCanceled(myOp)) return null;
            if (e.message === "__coverageTimeout__") {
                timedOut = true;
                console.warn(`Coverage check timed out for ${t.title} after ${COVERAGE_TIMEOUT_MS / 1000}s — assuming coverage`);
                if (modal) modal.addLog(`${t.title}: timed out — assuming coverage`, "warning");
                hasCoverage = true;
            } else {
                console.warn(`Coverage check failed for ${t.title}:`, e.message);
                // On error, assume coverage to be safe so it appears in report generation
                hasCoverage = true;
            }
        }

        // Build an export query for report generation so we don't have to re-query
        let exportQuery = null;
        if (hasCoverage && layerRef) {
            try {
                exportQuery = layerRef.createQuery();
                exportQuery.geometry = reportGeom;
                exportQuery.spatialRelationship = "intersects";
                exportQuery.outFields = ["*"];
                exportQuery.returnGeometry = false;
            } catch (e) {
                exportQuery = null;
            }
        }

        const reportEntry = {
            title: t.title,
            url: t.url,
            hasCoverage: hasCoverage,
            count: 0, // Feature count computed during report generation
            rows: [],
            _layer: layerRef,
            _exportQuery: exportQuery,
            fullRows: null
        };

        const statusBadge = timedOut ? 'assumed (timeout)' : hasCoverage ? 'has features' : 'no features';
        const statusClass = hasCoverage ? 'has-hits' : '';

        return {
            card: `
          <div class="result-card">
            <div class="result-head">
              <div class="result-title">${escapeHtml(t.title)}</div>
              <div class="badge ${statusClass}">${statusBadge}</div>
            </div>
            <div class="small mono">
              <a href="${escapeHtml(t.url)}" target="_blank" rel="noopener">Service URL</a>
            </div>
          </div>`,
            reportEntry
        };
    }
    // ── End of processOneTarget ──

    // ── Batched parallel queries ──
    // Larger batch size is safe since layers are on different domains
    const BATCH_SIZE = config.report?.queryBatchSize ?? 20;
    const cards = [];

    for (let bStart = 0; bStart < expandedTargets.length; bStart += BATCH_SIZE) {
        if (isReportCanceled(myOp)) {
            setStatus("canceled");
            break;
        }

        const bEnd = Math.min(bStart + BATCH_SIZE, expandedTargets.length);

        if (modal) {
            const progress = 25 + (70 * (bStart / expandedTargets.length));
            modal.setProgress(progress);
            modal.setStep(`Step 2/2: Checking datasets ${bStart + 1}-${bEnd} of ${expandedTargets.length}...`);
            for (let k = bStart; k < bEnd; k++) {
                modal.addLog(`Checking: ${expandedTargets[k].title}`);
            }
        }

        const batchResults = await Promise.allSettled(
            expandedTargets.slice(bStart, bEnd).map(t => processOneTarget(t))
        );

        // Check cancellation after batch completes
        if (isReportCanceled(myOp)) {
            setStatus("canceled");
            break;
        }

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

    // Only render results if analysis was not canceled
    if (!isReportCanceled(myOp)) {
        renderResults(cards.join(""));
        wireExportButtons();
        if (exportAllBtn) exportAllBtn.disabled = (lastReportRowsByLayer.length === 0);
    }
}




    // ── Shared export-all helper (used by exportAllBtn in Advanced mode) ──
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

            const myGen = ++_clickSelectGen;

            try {
                const hit = await view.hitTest(event);
                if (myGen !== _clickSelectGen) return; // superseded by newer click
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
                    if (myGen !== _clickSelectGen) return; // superseded by newer click
                    aoiSourceFeature = full?.feature || graphic || null;
                    const fullGeom = full?.geometry || null;
                    if (!fullGeom) return;

                    setAoiGeometry(fullGeom);
                    setGeometryFromSelection(fullGeom);
                    resetCoverageCacheForAoi(fullGeom);
                    aoiOriginalGeom = fullGeom; // store for optional buffering

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
                // When expanding back, scroll to the bottom so the user sees current content
                if (!isMinimized) {
                    scrollPanelToBottom();
                }
            });
        }

        config = await fetchJson("./config.json");

        // Strip excluded layers so they are invisible to screening, reporting, and mapping
        if (config.reportLayers) {
            config.reportLayers = config.reportLayers.filter(l => !l.excluded);
        }

        layerCfgByUrl = buildLayerCfgIndex(config);

        // Resolve reference layer URLs from config
        plssStateLayerUrl = config.referenceLayers?.plssStateBoundary || "https://services2.arcgis.com/FiaPA4ga0iQKduv3/arcgis/rest/services/Public_Land_Survey_System_view/FeatureServer/0";

        setLoadingState("Initializing map...", 20);

        map = new EsriMap({ basemap: config.map?.basemap || "gray-vector" });

        // --- Always-on basemap overlay: BLM SMA (BLM Only) ---
        const smaBlmOnly = new TileLayer({
            url: config.referenceLayers?.smaBLMOnly || "https://gis.blm.gov/arcgis/rest/services/lands/BLM_Natl_SMA_Cached_BLM_Only/MapServer",
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

        // State boundaries on the overview minimap for geographic context
        const overviewStateBoundaries = new FeatureLayer({
            url: config.referenceLayers?.usaStatesGeneralized || "https://services.arcgis.com/P3ePLMYs2RVChkJx/arcgis/rest/services/USA_States_Generalized_Boundaries/FeatureServer/0",
            title: "__overviewStates",
            outFields: [],
            labelsVisible: false,
            renderer: {
                type: "simple",
                symbol: {
                    type: "simple-fill",
                    color: [0, 0, 0, 0],
                    outline: { color: [200, 200, 200, 0.7], width: 1 }
                }
            }
        });

        let overviewMap = new EsriMap({
            basemap: imageryBasemapId,
            layers: [overviewStateBoundaries, overviewExtentLayer]
        });
        const overviewView = new MapView({
            container: "overviewMapView",
            map: overviewMap,
            ui: { components: [] },
            constraints: { snapToZoom: false, rotationEnabled: false },
            center: [-98.5795, 39.8283], // Center of contiguous US
            zoom: window.innerWidth <= 600 ? 1 : 2 // Lower zoom on mobile so full US is visible
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

        // Adjust overview zoom when window resizes
        window.addEventListener("resize", () => {
            const targetZoom = window.innerWidth <= 600 ? 1 : 2;
            if (overviewView.zoom !== targetZoom) {
                overviewView.zoom = targetZoom;
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
                let geom = evt.graphic?.geometry || null;
                if (!geom) return;

                const gType = geom.type; // "polygon", "polyline", "point"

                // Auto-buffer points/lines with a small buffer (0.1 mi ≈ 528 ft)
                // to convert zero-area geometry into a polygon AOI
                if (gType !== "polygon") {
                    const autoBuf = uploadAoiModule.applyBuffer([geom], 0.1);
                    if (autoBuf) geom = autoBuf;
                }

                // Finalize the drawn geometry as the AOI directly
                // (buffer option is available in the AOI confirm section)
                aoiOriginalGeom = geom;
                aoiSource = "draw";
                aoiSourceLayerTitle = null;
                aoiSourceLayerUrl = null;
                aoiSourceObjectId = null;
                aoiSourceObjectIdField = null;
                aoiSourceFeature = null;
                setAoiGeometry(geom);
                resetCoverageCacheForAoi(geom);
                setGeometryFromSelection(geom);
                setStatus("Drawn AOI ready");
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

        // Render selection layer toggles immediately (report layers will populate later)
        renderLayerToggles(map);
        ensureAoiOnTop();

        // ✅ Build report layers in BACKGROUND (don't block init)
        // This can take 30+ seconds with 50+ layer configs
        buildReportDisplayLayers().then(() => {
            console.log("[init] Report layers ready, refreshing layer toggles");
            renderLayerToggles(map);
            ensureAoiOnTop();
        }).catch(e => {
            console.error("[init] Failed to build report layers:", e);
        });

        setLoadingState("Waiting for map view...", 55);

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
                    // Township polygons are large — zoom in a bit more so
                    // individual townships are clearly distinguishable.
                    if (which === "township" && view) {
                        const extraScale = Math.floor(view.scale * 0.45);
                        if (extraScale > 0) {
                            await view.goTo(
                                { center: view.center, scale: extraScale },
                                { animate: true, duration: 300 }
                            );
                        }
                    }
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

        if (tabFinalReportBtn) tabFinalReportBtn.addEventListener("click", () => {
            setActiveTab("finalReport");
        });

        if (refreshServicesBtn) refreshServicesBtn.addEventListener("click", refreshServicesTab);
        if (viewReportBtn) viewReportBtn.addEventListener("click", generateFullProgressiveReport);

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

        // Wire search-by-name to set the selected feature as AOI
        searchModule.setOnFeatureSelected(function (feature) {
            if (!feature || !feature.geometry) {
                console.warn("Search result has no geometry");
                return;
            }

            const geomJson = feature.geometry;
            const sr = geomJson.spatialReference || view.spatialReference || { wkid: 102100 };

            // Convert raw REST API geometry JSON into auto-castable geometry with 'type'
            let geom;
            if (geomJson.rings && geomJson.rings.length > 0) {
                geom = { type: "polygon", rings: geomJson.rings, spatialReference: sr };
            } else if (geomJson.paths && geomJson.paths.length > 0) {
                geom = { type: "polyline", paths: geomJson.paths, spatialReference: sr };
            } else if (geomJson.x !== undefined && geomJson.y !== undefined) {
                geom = { type: "point", x: geomJson.x, y: geomJson.y, spatialReference: sr };
            } else {
                console.warn("Unrecognized geometry from search result", geomJson);
                return;
            }

            // For point/polyline search results, auto-cast to a proper
            // Geometry instance so generalizeAoiIfNeeded and buffer work
            var castGraphic = new Graphic({ geometry: geom });
            geom = castGraphic.geometry;

            // Set the geometry as the AOI (generalize before storing)
            geom = generalizeAoiIfNeeded(geom) || geom;
            selectionGeom = geom;
            aoiOriginalGeom = geom; // store for optional buffering
            aoiSource = "search";
            aoiSourceLayerTitle = feature.layerTitle || null;
            aoiSourceLayerUrl = feature.layerUrl || null;

            setAoiGeometry(geom);
            setGeometryFromSelection(geom);

            // Zoom to the selected feature — create a Graphic to force
            // reliable auto-casting of the plain geometry JSON, then use
            // the cast geometry's extent for a dependable goTo() target.
            try {
                const zoomGraphic = new Graphic({ geometry: geom });
                const castGeom = zoomGraphic.geometry;
                if (castGeom) {
                    if (castGeom.type === "point") {
                        view.goTo({ target: castGeom, zoom: 14 }, { animate: true, duration: 800 });
                    } else {
                        const target = castGeom.extent
                            ? castGeom.extent.expand(1.5)
                            : castGeom;
                        view.goTo(target, { animate: true, duration: 800 });
                    }
                }
            } catch (e) {
                console.warn("[Search] goTo failed:", e);
            }

            setStatus(`Selected: ${feature.layerTitle || "feature"} — ready to screen`);
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

        // AOI slide back buttons
        document.querySelectorAll(".aoi-slide-back").forEach(btn => {
            btn.addEventListener("click", () => {
                hideAoiMethodPanels();
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

                    const geomLabel = result.geometryType === "point" ? "point"
                                    : result.geometryType === "polyline" ? "line" : "polygon";
                    const geomLabelPlural = result.featureCount === 1 ? geomLabel : geomLabel + "s";

                    // Auto-buffer points/lines with a small buffer (0.1 mi ≈ 528 ft)
                    // to convert zero-area geometry into a polygon AOI
                    if (!result.hasPolygons && result.allGeometries && result.allGeometries.length) {
                        const autoBuf = uploadAoiModule.applyBuffer(result.allGeometries, 0.1);
                        if (autoBuf) {
                            result.geometry = autoBuf;
                            result.allGeometries = [autoBuf];
                            result.hasPolygons = true; // now a polygon
                        }
                    }

                    // Show the file info
                    const countLabel = result.featureCount + " " + geomLabelPlural;
                    showUploadStatus("📂 <b>" + configHelpers.escapeHtml(result.fileName) + "</b> — " + countLabel, "success");

                    // Zoom to uploaded geometry preview
                    if (view && result.geometry && result.geometry.extent) {
                        setAoiGeometry(result.geometry);
                        await view.goTo(result.geometry.extent.expand(1.5), { animate: true, duration: 600 });
                    }

                    // Show buffer panel (optional for all geometry types)
                    if (bufferPanel) {
                        const bufferMsg = document.getElementById("uploadBufferMsg");
                        const bufferInput = document.getElementById("uploadBufferInput");
                        const bufferSkipBtn = document.getElementById("uploadBufferSkip");

                        if (bufferMsg) bufferMsg.innerHTML = "Optionally add a buffer around your <strong>" + geomLabelPlural + "</strong>.";
                        if (bufferInput) bufferInput.value = "";
                        if (bufferSkipBtn) bufferSkipBtn.classList.remove("hidden");

                        bufferPanel.classList.remove("hidden");
                        scrollPanelToBottom(bufferPanel);

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
                } else {
                    // Use geometry as-is (already auto-buffered if originally point/line)
                    geometry = uploadAoiModule.applyBuffer(result.allGeometries, 0);
                    if (!geometry) {
                        showUploadStatus("❌ No polygon geometry available.", "error");
                        return;
                    }
                }

                geometry = generalizeAoiIfNeeded(geometry) || geometry;
                selectionGeom = geometry;
                aoiOriginalGeom = result.geometry; // store pre-buffer geometry for re-buffering
                aoiSource = "upload";
                aoiSourceLayerTitle = result.fileName;
                setAoiGeometry(geometry);

                if (runBtn) runBtn.disabled = false;

                // Zoom to buffered extent
                if (view && geometry && geometry.extent) {
                    view.goTo(geometry.extent.expand(1.3), { animate: true, duration: 600 });
                }

                aoiBufferMiles = bufferMiles || 0;
                const bufLabel = bufferMiles > 0 ? " with " + bufferMiles + " mi buffer" : "";
                showUploadStatus("✅ <b>" + configHelpers.escapeHtml(result.fileName) + "</b> loaded" + bufLabel, "success");

                // Hide buffer panel
                const bp = document.getElementById("uploadBufferPanel");
                if (bp) bp.classList.add("hidden");

                // Swipe to AOI confirm slide
                if (currentAppMode === "permit" && currentWizardStep === 2) {
                    populateAoiConfirmation();
                    swipeToAoiSlide(aoiSlideMap.confirm);
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

        // Watch for extent changes (debounced) — only schedule when permit panel is visible
        if (view) {
            view.watch("stationary", (stationary) => {
                if (!stationary) return;
                const pPanel = document.getElementById("aoiMethodPermit");
                if (pPanel && pPanel.classList.contains("hidden")) return;
                schedulePermitIndicatorUpdate();
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

                // Reset draw tool button active state
                document.querySelectorAll(".draw-tool-btn").forEach(b => b.classList.remove("active"));

                // Reset wizard UI
                if (wizFullReport) wizFullReport.disabled = true;

                setStatus("draw canceled");
            });
        }

        // ── Confirm-section buffer panel (universal, works for all AOI methods) ──
        const confirmBufApply = document.getElementById("aoiConfirmBufferApply");
        const confirmBufReset = document.getElementById("aoiConfirmBufferReset");

        if (confirmBufApply) {
            confirmBufApply.addEventListener("click", () => {
                const input = document.getElementById("aoiConfirmBufferInput");
                const val = parseFloat(input ? input.value : "");
                if (isNaN(val) || val <= 0) {
                    setStatus("Enter a buffer distance greater than 0.");
                    return;
                }
                if (val > 50) {
                    setStatus("Buffer cannot exceed 50 miles.");
                    return;
                }

                // Save original geometry before first buffer
                if (!aoiOriginalGeom) aoiOriginalGeom = selectionGeom;

                const buffered = uploadAoiModule.applyBuffer([aoiOriginalGeom], val);
                if (!buffered) {
                    setStatus("Buffer operation failed. Try a different value.");
                    return;
                }

                selectionGeom = generalizeAoiIfNeeded(buffered) || buffered;
                setAoiGeometry(selectionGeom);
                resetCoverageCacheForAoi(selectionGeom);
                if (view && selectionGeom.extent) {
                    view.goTo(selectionGeom.extent.expand(1.3), { animate: true, duration: 600 });
                }
                if (runBtn) runBtn.disabled = false;

                // Invalidate any cached report — geometry has changed
                Object.keys(cachedPermitReports).forEach(k => delete cachedPermitReports[k]);
                if (wizFullReport) {
                    delete wizFullReport.dataset.reportReady;
                    wizFullReport.classList.remove('ready-to-view');
                }

                // Refresh the confirmation card with new area
                populateAoiConfirmation();
                // Re-set the input value and reset button since populateAoiConfirmation clears them
                if (input) input.value = val;
                aoiBufferMiles = val;
                if (confirmBufReset) confirmBufReset.classList.remove("hidden");

                setStatus("Buffer applied (" + val + " mi)");
            });
        }

        if (confirmBufReset) {
            confirmBufReset.addEventListener("click", () => {
                if (!aoiOriginalGeom) return;

                selectionGeom = aoiOriginalGeom;
                // Keep aoiOriginalGeom so re-buffer is possible later
                setAoiGeometry(selectionGeom);
                resetCoverageCacheForAoi(selectionGeom);
                if (view && selectionGeom.extent) {
                    view.goTo(selectionGeom.extent.expand(1.3), { animate: true, duration: 600 });
                }
                if (runBtn) runBtn.disabled = false;

                // Invalidate any cached report — geometry has changed
                Object.keys(cachedPermitReports).forEach(k => delete cachedPermitReports[k]);
                if (wizFullReport) {
                    delete wizFullReport.dataset.reportReady;
                    wizFullReport.classList.remove('ready-to-view');
                }

                aoiBufferMiles = 0;
                populateAoiConfirmation();
                setStatus("Buffer removed — original geometry restored");
            });
        }

        // Confirm buffer preset buttons
        document.querySelectorAll(".confirm-buf-preset").forEach(btn => {
            btn.addEventListener("click", () => {
                const input = document.getElementById("aoiConfirmBufferInput");
                if (input) input.value = btn.dataset.val;
                document.querySelectorAll(".confirm-buf-preset").forEach(b => b.classList.remove("active"));
                btn.classList.add("active");
            });
        });

        // Location search
        if (wizLocationInput) {
            wizLocationInput.addEventListener("input", () => {
                performLocationSearch(wizLocationInput.value.trim());
            });
        }

        // ── Permit Type Card Selection (Step 1 → Step 2) ──
        document.querySelectorAll('.permit-type-card[data-permit-type]').forEach(card => {
            card.addEventListener('click', () => {
                const ptKey = card.dataset.permitType;
                if (!ptKey || !PERMIT_TYPES[ptKey]) return;

                // Set selected permit type
                selectedPermitType = ptKey;
                document.querySelectorAll('.permit-type-card').forEach(c => c.classList.remove('selected'));
                card.classList.add('selected');

                // Update badge in step 2
                const badge = document.getElementById("wizStep2PermitBadge");
                if (badge) badge.textContent = PERMIT_TYPES[ptKey].label;

                // Advance to step 2 (AOI selection)
                goToWizardStep(2);
            });
        });

        // Back to permit type — full reset (same as starting a new screening)
        const wizBackToPermitType = document.getElementById("wizBackToPermitType");
        if (wizBackToPermitType) {
            wizBackToPermitType.addEventListener("click", () => {
                clearAll();
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

        if (wizFullReport) wizFullReport.addEventListener("click", generateFullProgressiveReport);

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
                    alert("This report bookmark has expired or is not available.\n\nBookmarks are valid for 7 days and can only be opened on the device/browser where the analysis was run.");
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
        
        // Initialize report building modal
        reportModal.init();

        // Initialize service-down warning modal
        serviceDownModal.wire();

        // Background service check — delay 5s so it doesn't compete with
        // critical layer loading for browser connection slots
        setTimeout(() => {
            checkServiceStatusBackground().catch(() => {});
            // Pre-expand service roots in background so they're cached for analysis
            preExpandServiceRoots();
        }, 5000);
    }

    init().catch((e) => {
        console.error(e);
        setStatus("failed to initialize (see console)");
    });

});

// ════════════════════════════════════════════════════════════════════════════
// Scroll-Fade Indicators
// Watches scrollable containers and toggles .can-scroll-up / .can-scroll-down
// ════════════════════════════════════════════════════════════════════════════
(function initScrollFades() {
    const THRESHOLD = 4; // px from edge before showing fade

    function updateFade(el) {
        const canUp   = el.scrollTop > THRESHOLD;
        const canDown = el.scrollTop + el.clientHeight < el.scrollHeight - THRESHOLD;
        el.classList.toggle("can-scroll-up", canUp);
        el.classList.toggle("can-scroll-down", canDown);
    }

    function observe(el) {
        if (!el || el.dataset.scrollFadeInit) return;
        el.dataset.scrollFadeInit = "1";
        el.classList.add("scroll-fade");
        updateFade(el);
        el.addEventListener("scroll", () => updateFade(el), { passive: true });

        // Re-check after content changes (e.g. search results populating)
        const ro = new ResizeObserver(() => updateFade(el));
        ro.observe(el);
    }

    /**
     * Observe a container that may not exist yet.
     * Uses a MutationObserver on document.body to watch for it.
     */
    function observeWhenReady(selector, extraClasses) {
        const el = document.querySelector(selector);
        if (el) {
            observe(el);
            if (extraClasses) extraClasses.forEach(c => el.classList.add(c));
            return;
        }
        // Container does not exist yet — watch for it
        const mo = new MutationObserver(() => {
            const found = document.querySelector(selector);
            if (found) {
                mo.disconnect();
                observe(found);
                if (extraClasses) extraClasses.forEach(c => found.classList.add(c));
            }
        });
        mo.observe(document.body, { childList: true, subtree: true });
    }

    // Main panel — always in the DOM
    observeWhenReady("#panel");

    // Scrollable sub-containers (may appear dynamically)
    observeWhenReady(".feature-search-results", ["scroll-fade-sm"]);
    observeWhenReady(".feature-picker-list", ["scroll-fade-sm"]);
    observeWhenReady(".wiz-location-results", ["scroll-fade-sm"]);
    observeWhenReady(".layer-mgr-body", ["scroll-fade-dark"]);

    // ── Auto-scroll panel to bottom when new content appears ──
    // (Auto-scroll observer removed — individual scroll calls at interaction
    // points (large-AOI warning, buffer panels, bucket navigation, panel
    // expand) handle scrolling intentionally without fighting the user.)
})();

// ════════════════════════════════════════════════════════════════════════════
// Mobile gesture-conflict prevention
// ════════════════════════════════════════════════════════════════════════════
(function preventBrowserGestures() {
    // iOS Safari fires proprietary gesturestart on two-finger pinch.
    // Preventing it stops the browser from zooming the whole page,
    // so the ArcGIS map's own pinch-to-zoom works unimpeded.
    document.addEventListener("gesturestart", function (e) {
        e.preventDefault();
    }, { passive: false });

    // Prevent double-tap-to-zoom on the map surface (300 ms tap delay).
    // The ArcGIS SDK already handles double-click-zoom via its navigation.
    var viewDiv = document.getElementById("viewDiv");
    if (viewDiv) {
        var lastTap = 0;
        viewDiv.addEventListener("touchend", function (e) {
            var now = Date.now();
            if (now - lastTap < 300) { e.preventDefault(); }
            lastTap = now;
        }, { passive: false });
    }
})();

// ════════════════════════════════════════════════════════════════════════════
// Accessibility Color-Vision Widget (independent of ArcGIS modules)
// ════════════════════════════════════════════════════════════════════════════
(function initA11yWidget() {
    const STORAGE_KEY = "a11y-cv-mode";
    const toggleBtn = document.getElementById("a11yToggleBtn");
    const menu = document.getElementById("a11yMenu");
    if (!toggleBtn || !menu) return;

    const options = menu.querySelectorAll(".a11y-option");

    // In-document SVG filter references — now applied via CSS classes on <body>
    // (iOS Safari mobile silently ignores SVG url() filters set via inline JS styles;
    // CSS-class rules in styles.css work correctly across all platforms.)

    var CV_CLASSES = ["cv-protanopia", "cv-deuteranopia", "cv-tritanopia", "cv-achromatopsia", "cv-highcontrast"];

    function applyMode(mode) {
        // Remove all cv- classes then add the active one.
        // CSS-class approach works reliably on iOS Safari (inline-style SVG url()
        // filters are silently ignored on mobile WebKit).
        CV_CLASSES.forEach(function (c) { document.body.classList.remove(c); });
        if (mode && mode !== "none") {
            document.body.classList.add("cv-" + mode);
        }
        // Clear any leftover inline filter styles from prior approach
        document.body.style.webkitFilter = "";
        document.body.style.filter = "";
        // Update aria-checked on menu items
        options.forEach(btn => {
            btn.setAttribute("aria-checked", btn.dataset.cv === mode ? "true" : "false");
        });
        // Persist
        try { localStorage.setItem(STORAGE_KEY, mode || "none"); } catch (_) {}
    }

    // Restore saved preference
    let saved = "none";
    try { saved = localStorage.getItem(STORAGE_KEY) || "none"; } catch (_) {}
    applyMode(saved);

    // Toggle menu open/close
    toggleBtn.addEventListener("click", (e) => {
        e.stopPropagation(); // Prevent document click handler from immediately closing
        const isHidden = menu.classList.toggle("hidden");
        toggleBtn.setAttribute("aria-expanded", isHidden ? "false" : "true");
        if (!isHidden) {
            // Focus first menu item for keyboard navigation
            const first = menu.querySelector(".a11y-option");
            if (first) first.focus();
        }
    });

    // Menu item clicks
    options.forEach(btn => {
        btn.addEventListener("click", (e) => {
            e.stopPropagation();
            applyMode(btn.dataset.cv);
            menu.classList.add("hidden");
            toggleBtn.setAttribute("aria-expanded", "false");
        });
    });

    // Keyboard navigation inside menu
    menu.addEventListener("keydown", (e) => {
        const items = [...options];
        const idx = items.indexOf(document.activeElement);
        if (e.key === "ArrowDown" || e.key === "ArrowRight") {
            e.preventDefault();
            items[(idx + 1) % items.length].focus();
        } else if (e.key === "ArrowUp" || e.key === "ArrowLeft") {
            e.preventDefault();
            items[(idx - 1 + items.length) % items.length].focus();
        } else if (e.key === "Escape") {
            menu.classList.add("hidden");
            toggleBtn.setAttribute("aria-expanded", "false");
            toggleBtn.focus();
        }
    });

    // Close if user clicks outside the widget
    document.addEventListener("click", (e) => {
        const widget = document.getElementById("a11yWidget");
        if (!menu.classList.contains("hidden") && widget && !widget.contains(e.target)) {
            menu.classList.add("hidden");
            toggleBtn.setAttribute("aria-expanded", "false");
        }
    });
})();