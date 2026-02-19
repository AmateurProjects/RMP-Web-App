/**
 * map-utils.js — Map layer management, renderers, AOI display,
 * screenshot helpers, and layer visibility utilities.
 *
 * AMD module. Depends on config-helpers + several Esri modules.
 *
 * Usage from app.js:
 *   const mapUtils = mapUtilsFactory.init(state);
 *   // state is the shared mutable state object
 */
define([
    "app/config-helpers",
    "esri/Graphic",
    "esri/layers/FeatureLayer",
    "esri/layers/ImageryLayer"
], function (helpers, Graphic, FeatureLayer, ImageryLayer) {
    "use strict";

    // ── Module-level reference to shared state (set by init) ──
    let S = null;

    // ── Renderer / Symbology ─────────────────────────────────────────

    function getPresetRenderer(kind, cfgObj, geometryType) {
        const sym = S.config?.symbology || {};
        const defaults = sym.defaults || {};
        const presets = sym.presets || {};

        let presetId =
            (cfgObj && cfgObj.symbologyPreset) ||
            (kind === "selection" ? defaults.selectionPreset :
                kind === "report" ? defaults.reportPreset :
                    defaults.aoiPreset);

        if (kind === "report" && geometryType) {
            const gt = String(geometryType).toLowerCase();
            if (gt.includes("point")) {
                presetId = "reportPoint";
            } else if (gt.includes("line") || gt.includes("polyline")) {
                presetId = "reportLine";
            }
        }

        const r = presetId ? presets[presetId] : null;
        return r || null;
    }

    // PERF: Cache geometry type lookups to avoid refetching ?f=pjson
    const _geomTypeCache = new Map();

    async function getLayerGeometryType(layerUrl) {
        const cacheKey = layerUrl.replace(/\/$/, "");
        if (_geomTypeCache.has(cacheKey)) return _geomTypeCache.get(cacheKey);
        try {
            const pjsonUrl = cacheKey + "?f=pjson";
            const pjson = await helpers.fetchJsonWithTimeout(pjsonUrl, 5000);
            const gt = pjson?.geometryType || null;
            _geomTypeCache.set(cacheKey, gt);
            return gt;
        } catch (e) {
            return null;
        }
    }

    function makeRendererOpaque(renderer) {
        if (!renderer) return renderer;
        const r = JSON.parse(JSON.stringify(renderer));
        const forceOpaque = (c) => {
            if (Array.isArray(c) && c.length >= 4) c[3] = (c[3] <= 1) ? 1 : 255;
        };
        if (r.symbol) {
            if (r.symbol.color) forceOpaque(r.symbol.color);
            if (r.symbol.outline && r.symbol.outline.color) forceOpaque(r.symbol.outline.color);
        }
        return r;
    }

    /**
     * Thicken the symbology of a loaded FeatureLayer for clearer report maps.
     * - Polygon layers: ensures outline width >= MIN_POLYGON_OUTLINE (default 4pt)
     * - Polyline layers: ensures line width >= MIN_LINE_WIDTH (default 3pt)
     * - Point layers: left unchanged (markers are already visible)
     * Preserves the service's original colors and symbology; only overrides widths.
     * @param {FeatureLayer} layer — a loaded FeatureLayer with a renderer
     * @param {string} geomType — geometry type string from the service (e.g. "esriGeometryPolygon")
     */
    function thickenLayerSymbology(layer, geomType) {
        if (!layer || !layer.renderer || !geomType) return;
        const gt = String(geomType).toLowerCase();
        const isPolygon  = gt.includes("polygon");
        const isPolyline = gt.includes("polyline") || gt.includes("line");
        if (!isPolygon && !isPolyline) return;

        try {
            const r = layer.renderer.clone();
            const MIN_POLYGON_OUTLINE = 4;
            const MIN_LINE_WIDTH      = 3;

            function thickenSymbol(sym) {
                if (!sym) return;

                if (isPolygon) {
                    // Thicken polygon outline
                    if (sym.outline) {
                        sym.outline.width = Math.max(sym.outline.width || 0, MIN_POLYGON_OUTLINE);
                    } else {
                        sym.outline = { color: [0, 0, 0, 0.8], width: MIN_POLYGON_OUTLINE };
                    }
                } else if (isPolyline) {
                    // Thicken line width (SimpleLineSymbol stores width directly on the symbol)
                    if (sym.width != null) {
                        sym.width = Math.max(sym.width, MIN_LINE_WIDTH);
                    }
                }
            }

            if (r.symbol)         thickenSymbol(r.symbol);
            if (r.defaultSymbol)  thickenSymbol(r.defaultSymbol);
            if (r.uniqueValueInfos)  r.uniqueValueInfos.forEach(uv => thickenSymbol(uv.symbol));
            if (r.classBreakInfos)   r.classBreakInfos.forEach(cb => thickenSymbol(cb.symbol));

            layer.renderer = r;
        } catch (e) {
            console.warn("[thickenLayerSymbology] Could not thicken symbology for layer:", e);
        }
    }

    // ── AOI Layer Management ─────────────────────────────────────────

    function ensureAoiOnTop() {
        const map = S.map;
        const aoiLayer = S.aoiLayer;
        if (!map || !aoiLayer) return;
        map.reorder(aoiLayer, map.layers.length - 1);
        if (S.aoiMaskLayer) {
            map.reorder(S.aoiMaskLayer, map.layers.length - 2);
        }
    }

    function updateAoiMask(show) {
        if (show === undefined) show = true;
        const aoiMaskLayer = S.aoiMaskLayer;
        const view = S.view;

        if (!aoiMaskLayer || !view) return;

        aoiMaskLayer.removeAll();

        if (!show || !S.selectionGeom) {
            aoiMaskLayer.visible = false;
            return;
        }

        const viewExt = view.extent;
        if (!viewExt) {
            aoiMaskLayer.visible = false;
            return;
        }

        // ── Ring winding helpers ──
        // Shoelace sum: positive → clockwise, negative → counter-clockwise
        // (in projected coordinate systems where Y increases upward)
        function shoelaceSum(ring) {
            let sum = 0;
            for (let i = 0; i < ring.length - 1; i++) {
                const x1 = ring[i][0],     y1 = ring[i][1];
                const x2 = ring[i + 1][0], y2 = ring[i + 1][1];
                sum += (x2 - x1) * (y2 + y1);
            }
            return sum;
        }
        function ensureClockwise(ring) {
            return shoelaceSum(ring) >= 0 ? ring : [...ring].reverse();
        }
        function ensureCounterClockwise(ring) {
            return shoelaceSum(ring) <= 0 ? ring : [...ring].reverse();
        }

        const expandedExt = viewExt.expand(5);

        // Outer ring must be CLOCKWISE for the non-zero winding fill rule
        const outerRing = ensureClockwise([
            [expandedExt.xmin, expandedExt.ymin],
            [expandedExt.xmax, expandedExt.ymin],
            [expandedExt.xmax, expandedExt.ymax],
            [expandedExt.xmin, expandedExt.ymax],
            [expandedExt.xmin, expandedExt.ymin]
        ]);

        // ── Collect AOI rings and ensure they are COUNTER-CLOCKWISE (holes) ──
        let aoiRings = [];
        if (S.selectionGeom.rings && S.selectionGeom.rings.length > 0) {
            // Project rings to view SR if the geometry is in a different spatial reference
            let srcRings = S.selectionGeom.rings;
            const geomSR = S.selectionGeom.spatialReference;
            const viewSR = view.spatialReference;
            if (geomSR && viewSR && geomSR.wkid !== viewSR.wkid) {
                // WGS84 (4326) → Web Mercator (102100/3857) conversion
                const isGeomWgs84 = (geomSR.wkid === 4326);
                const isViewWebMerc = (viewSR.isWebMercator || viewSR.wkid === 102100 || viewSR.wkid === 3857);
                if (isGeomWgs84 && isViewWebMerc) {
                    const DEG2RAD = Math.PI / 180;
                    const EARTH_R = 6378137;
                    srcRings = srcRings.map(function (ring) {
                        return ring.map(function (pt) {
                            var x = pt[0] * EARTH_R * DEG2RAD;
                            var y = Math.log(Math.tan((90 + pt[1]) * DEG2RAD / 2)) * EARTH_R;
                            return [x, y];
                        });
                    });
                }
            }
            aoiRings = srcRings.map(ring => ensureCounterClockwise(ring));
        } else if (S.selectionGeom.type === "polygon" && S.selectionGeom.extent) {
            const ext = S.selectionGeom.extent;
            // Explicit CCW ring for the hole
            aoiRings = [ensureCounterClockwise([
                [ext.xmin, ext.ymin],
                [ext.xmax, ext.ymin],
                [ext.xmax, ext.ymax],
                [ext.xmin, ext.ymax],
                [ext.xmin, ext.ymin]
            ])];
        }

        if (aoiRings.length === 0) {
            aoiMaskLayer.visible = false;
            return;
        }

        const allRings = [outerRing, ...aoiRings];

        const maskGraphic = new Graphic({
            geometry: {
                type: "polygon",
                rings: allRings,
                spatialReference: view.spatialReference
            },
            symbol: {
                type: "simple-fill",
                color: [255, 255, 255, 0.6],
                outline: null
            }
        });

        aoiMaskLayer.add(maskGraphic);
        aoiMaskLayer.visible = true;
    }

    function hideAoiMask() {
        if (S.aoiMaskLayer) {
            S.aoiMaskLayer.visible = false;
            S.aoiMaskLayer.removeAll();
        }
    }

    function setAoiGeometry(geom) {
        const aoiLayer = S.aoiLayer;
        if (!aoiLayer) return;

        aoiLayer.removeAll();
        S.aoiGraphic = null;

        if (!geom) return;

        const aoiRenderer = getPresetRenderer("aoi", null);
        let aoiSymbol = aoiRenderer?.symbol;

        // The default AOI symbol is a polygon fill — pick an appropriate
        // symbol when the geometry is a point or polyline so it's visible.
        const gt = geom.type;
        if (gt === "point" || gt === "multipoint") {
            const outlineColor = aoiSymbol?.outline?.color || [230, 57, 70, 1];
            aoiSymbol = {
                type: "simple-marker",
                style: "circle",
                color: outlineColor,
                size: 12,
                outline: { color: [255, 255, 255, 1], width: 2 }
            };
        } else if (gt === "polyline") {
            const outlineColor = aoiSymbol?.outline?.color || [230, 57, 70, 1];
            aoiSymbol = {
                type: "simple-line",
                color: outlineColor,
                width: 3,
                style: "solid"
            };
        }

        S.aoiGraphic = new Graphic({
            geometry: geom,
            symbol: aoiSymbol || undefined
        });

        aoiLayer.add(S.aoiGraphic);
    }

    // ── Layer Visibility / Scale ─────────────────────────────────────

    async function ensureLayerVisibleAtScale(layer) {
        const view = S.view;
        if (!view || !layer) return;

        const minScale = Number(layer.minScale || 0);
        const maxScale = Number(layer.maxScale || 0);

        let targetScale = null;

        if (minScale > 0 && isFinite(minScale) && view.scale > minScale) {
            targetScale = Math.max(1, Math.floor(minScale * 0.90));
        } else if (maxScale > 0 && isFinite(maxScale) && view.scale < maxScale) {
            targetScale = Math.ceil(maxScale * 1.10);
        }

        if (targetScale && isFinite(targetScale) && targetScale > 0) {
            await view.goTo(
                { center: view.center, scale: targetScale },
                { animate: true, duration: 250 }
            );
        }
    }

    // ── Layer Updating / Spinner Helpers ──────────────────────────────

    async function wireLayerUpdatingSpinner(layer, spinnerEl) {
        const view = S.view;
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

    function waitForViewStationary(timeoutMs) {
        if (timeoutMs === undefined) timeoutMs = 1200;
        const view = S.view;
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

    // ── Background-tab resilience ────────────────────────────────────

    /**
     * Returns a Promise that resolves as soon as the page is visible.
     * If the tab is already visible it resolves immediately.  Otherwise it
     * waits for the `visibilitychange` event with a configurable timeout.
     * @param {number} timeoutMs - Max time to wait before proceeding anyway (default: 60000ms)
     */
    function waitForTabVisible(timeoutMs = 60000) {
        if (!document.hidden) return Promise.resolve();
        console.log("[map-utils] Tab hidden — pausing until visible (max " + (timeoutMs/1000) + "s)…");
        return new Promise(resolve => {
            let resolved = false;
            
            function cleanup() {
                if (!resolved) {
                    resolved = true;
                    document.removeEventListener("visibilitychange", onVis);
                    resolve();
                }
            }
            
            function onVis() {
                if (!document.hidden) {
                    console.log("[map-utils] Tab visible — resuming.");
                    cleanup();
                }
            }
            
            document.addEventListener("visibilitychange", onVis);
            
            // Timeout: proceed anyway after waiting to avoid hanging forever
            setTimeout(() => {
                if (!resolved) {
                    console.warn("[map-utils] Tab visibility wait timed out — proceeding anyway (screenshots may be blank)");
                    cleanup();
                }
            }, timeoutMs);
        });
    }

    let _wakeLock = null;

    /** Acquire a Wake Lock to prevent the device from sleeping. */
    async function acquireWakeLock() {
        try {
            if (navigator.wakeLock) {
                _wakeLock = await navigator.wakeLock.request("screen");
                // Re-acquire if released (e.g. tab switch on some browsers)
                _wakeLock.addEventListener("release", () => { _wakeLock = null; });
                console.log("[map-utils] Wake Lock acquired");
            }
        } catch (e) {
            console.warn("[map-utils] Wake Lock request failed:", e.message);
        }
    }

    /** Release the Wake Lock. */
    async function releaseWakeLock() {
        try {
            if (_wakeLock) { await _wakeLock.release(); _wakeLock = null; console.log("[map-utils] Wake Lock released"); }
        } catch (e) { /* ignore */ }
    }

    // ── Screenshot / Capture Helpers ─────────────────────────────────

    async function waitForLayerReadyToCapture(layer, view, opts) {
        if (!opts) opts = {};
        const timeoutMs = opts.timeoutMs || 8000;
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

        // Wait for updating to complete, with a stability check:
        // After updating goes false, wait 200ms and verify it's still false
        // (some layers briefly go false then true again as sub-requests fire)
        const deadline = Date.now() + timeoutMs;
        while (lv.updating && Date.now() < deadline) {
            await new Promise(resolve => {
                const remaining = Math.max(500, deadline - Date.now());
                const t = window.setTimeout(() => { h?.remove?.(); resolve(); }, remaining);
                const h = lv.watch("updating", (u) => {
                    if (!u) {
                        clearTimeout(t);
                        h.remove();
                        resolve();
                    }
                });
            });
            // Stability check: wait a moment and see if it stays non-updating
            if (!lv.updating) {
                await new Promise(r => setTimeout(r, 150));
                // If it started updating again, loop continues
            }
        }

        // Final render settle
        await new Promise(r => setTimeout(r, 150));

        // Ensure tab is visible so the canvas gets painted (short timeout to avoid hanging)
        await waitForTabVisible(5000);
    }

    async function captureScreenshotWithWait(screenConfig) {
        if (!screenConfig) screenConfig = {};
        const view = S.view;
        if (!view) return null;

        const width = screenConfig.width || (S.config?.visualReport?.screenshotWidth ?? 1400);
        const maxRetries = screenConfig.maxRetries || 3;
        // For progressive reports, use shorter timeout since user is likely viewing the popup
        const tabWaitTimeout = screenConfig.tabWaitTimeout || 5000;

        // Brief wait for tab visibility (with timeout to avoid hanging)
        await waitForTabVisible(tabWaitTimeout);

        await waitForViewStationary(1500);

        // Force a render frame so the canvas is up-to-date
        await new Promise(r => requestAnimationFrame(r));

        const ssOpts = {
            format: "png",
            quality: 100,
            width: width,
            height: Math.round(width * 0.5625)
        };

        for (let attempt = 1; attempt <= maxRetries; attempt++) {
            try {
                const ss = await view.takeScreenshot(ssOpts);
                if (ss?.dataUrl) return ss.dataUrl;
            } catch (e) {
                console.warn("[map-utils] Screenshot attempt " + attempt + " failed:", e.message);
            }
            // Back-off delay before retry
            await new Promise(r => setTimeout(r, 400 * attempt));
            // Request a fresh frame before next attempt
            await new Promise(r => requestAnimationFrame(r));
        }

        console.error("[map-utils] Screenshot capture failed after " + maxRetries + " attempts.");
        return null;
    }

    async function hardRefreshLayer(layer, opts) {
        if (!opts) opts = {};
        const timeoutMs = opts.timeoutMs || 5000;
        const view = S.view;
        if (!view || !layer) return;

        try { await layer.when(); } catch (e) { console.warn("Layer.when() error:", e); }

        let lv = null;
        try { lv = await view.whenLayerView(layer); } catch (e) {
            console.warn("whenLayerView error:", e);
            return;
        }
        if (!lv) return;

        await waitForViewStationary(1500);

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

        if (typeof lv.refresh === "function") {
            try {
                lv.refresh();
            } catch (e) {
                console.warn("Layer refresh failed:", e);
            }
        }

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

        try {
            await new Promise(r => setTimeout(r, 200));
        } catch (e) { }
    }

    // ── Build Report Display Layers ──────────────────────────────────
    //
    // PERF: Processes layers in parallel batches (BATCH_SIZE at a time)
    // instead of sequentially.  Each layer also gets a load timeout so
    // a single slow/hung service can't block the whole init.

    /** Race a promise against a timeout.  Rejects with an Error on timeout. */
    function withTimeout(promise, ms, label) {
        return new Promise((resolve, reject) => {
            const t = setTimeout(() => reject(new Error(`Timeout (${ms}ms) loading ${label}`)), ms);
            promise.then(v => { clearTimeout(t); resolve(v); },
                         e => { clearTimeout(t); reject(e); });
        });
    }

    async function buildReportDisplayLayers() {
        const map = S.map;
        if (!map) return;

        S.reportLayerViews.clear();

        const BATCH_SIZE = 10;          // concurrent layers per batch
        const LOAD_TIMEOUT_MS = 15000;  // per-layer load timeout

        // Deduplicate by URL key
        const cfgs = [];
        const seenKeys = new Set();
        for (const cfg of (S.config.reportLayers || [])) {
            const key = helpers.normalizeUrlKey(cfg.url);
            if (!key || seenKeys.has(key)) continue;
            seenKeys.add(key);
            cfgs.push({ cfg, key });
        }

        /**
         * Process a single config entry — returns { key, layers[]|layer, cfg }
         * or null on failure.  Runs entirely independently so it can be batched.
         */
        async function processOne({ cfg, key }) {
            const useServiceRenderer = cfg.useServiceRenderer === true;
            const isAlwaysVisible = cfg.alwaysVisible === true;

            try {
                // ── FeatureServer root ──
                if (helpers.isFeatureServerRoot(key)) {
                    const subs = useServiceRenderer
                        ? await helpers.expandFeatureServerToAllSublayers(key)
                        : await helpers.expandFeatureServerToPolygonSublayers(key);

                    const geomTypes = await Promise.all(
                        subs.map(sl => getLayerGeometryType(sl.url))
                    );

                    const layers = [];
                    // Load sublayers in parallel (they are on the same server — browser queues them anyway)
                    const loadResults = await Promise.allSettled(subs.map((sl, si) => {
                        const geomType = geomTypes[si];
                        const layerOpts = {
                            url: sl.url,
                            title: `${cfg.title}: ${sl.title}`,
                            outFields: ["*"],
                            visible: isAlwaysVisible
                        };
                        if (!useServiceRenderer) {
                            layerOpts.renderer = getPresetRenderer("report", cfg, geomType) || undefined;
                        }
                        if (cfg.minScale !== undefined) layerOpts.minScale = cfg.minScale;
                        if (cfg.maxScale !== undefined) layerOpts.maxScale = cfg.maxScale;
                        const lyr = new FeatureLayer(layerOpts);
                        return withTimeout(lyr.load(), LOAD_TIMEOUT_MS, sl.title)
                            .then(() => {
                                // Skip layers without query capability (avoids featurelayerview:query-not-supported)
                                if (!lyr.capabilities?.query?.supportsSqlExpression) {
                                    console.warn(`[buildReportDisplayLayers] Skipping sublayer "${sl.title}": no query capability`);
                                    return { lyr: null, sl, ok: false };
                                }
                                return { lyr, sl, ok: true };
                            })
                            .catch(e => {
                                console.warn(`[buildReportDisplayLayers] Skipping sublayer "${sl.title}":`, e.message || e);
                                return { lyr: null, sl, ok: false };
                            });
                    }));

                    for (const r of loadResults) {
                        if (r.status !== "fulfilled" || !r.value.ok) continue;
                        layers.push(r.value.lyr);
                        if (isAlwaysVisible) S.alwaysVisibleLayers.push(r.value.lyr);
                        S.layerCfgByUrl.set(r.value.sl.url, { kind: "report", cfg });
                    }

                    return { key, result: layers, cfg, type: "multi" };
                }

                // ── MapServer root ──
                if (helpers.isMapServerRoot(key)) {
                    const subs = await helpers.expandMapServerToSublayers(key, { polygonOnly: !useServiceRenderer });

                    const geomTypes = await Promise.all(
                        subs.map(sl => getLayerGeometryType(sl.url))
                    );

                    const layers = [];
                    const loadResults = await Promise.allSettled(subs.map((sl, si) => {
                        const geomType = geomTypes[si];
                        const layerOpts = {
                            url: sl.url,
                            title: `${cfg.title}: ${sl.title}`,
                            outFields: ["*"],
                            visible: isAlwaysVisible
                        };
                        if (!useServiceRenderer) {
                            layerOpts.renderer = getPresetRenderer("report", cfg, geomType) || undefined;
                        }
                        if (cfg.minScale !== undefined) layerOpts.minScale = cfg.minScale;
                        if (cfg.maxScale !== undefined) layerOpts.maxScale = cfg.maxScale;
                        const lyr = new FeatureLayer(layerOpts);
                        return withTimeout(lyr.load(), LOAD_TIMEOUT_MS, sl.title)
                            .then(() => {
                                if (!lyr.capabilities?.query?.supportsSqlExpression) {
                                    console.warn(`[buildReportDisplayLayers] Skipping sublayer "${sl.title}": no query capability`);
                                    return { lyr: null, sl, ok: false };
                                }
                                return { lyr, sl, ok: true };
                            })
                            .catch(e => {
                                console.warn(`[buildReportDisplayLayers] Skipping sublayer "${sl.title}":`, e.message || e);
                                return { lyr: null, sl, ok: false };
                            });
                    }));

                    for (const r of loadResults) {
                        if (r.status !== "fulfilled" || !r.value.ok) continue;
                        layers.push(r.value.lyr);
                        if (isAlwaysVisible) S.alwaysVisibleLayers.push(r.value.lyr);
                        S.layerCfgByUrl.set(r.value.sl.url, { kind: "report", cfg });
                    }

                    return { key, result: layers, cfg, type: "multi" };
                }

                // ── ImageServer ──
                if (cfg.imageService === true) {
                    const layerOpts = {
                        url: key,
                        title: cfg.title,
                        visible: isAlwaysVisible
                    };
                    if (cfg.renderingRule) {
                        layerOpts.rasterFunction = { functionName: cfg.renderingRule };
                    }
                    const lyr = new ImageryLayer(layerOpts);
                    if (isAlwaysVisible) S.alwaysVisibleLayers.push(lyr);
                    S.layerCfgByUrl.set(key, { kind: "report", cfg, isImageService: true });
                    return { key, result: lyr, cfg, type: "single" };
                }

                // ── Normal single layer ──
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
                if (cfg.minScale !== undefined) layerOpts.minScale = cfg.minScale;
                if (cfg.maxScale !== undefined) layerOpts.maxScale = cfg.maxScale;
                const lyr = new FeatureLayer(layerOpts);
                if (isAlwaysVisible) S.alwaysVisibleLayers.push(lyr);

                return { key, result: lyr, cfg, type: "single" };

            } catch (e) {
                console.warn(`[buildReportDisplayLayers] Failed to load "${cfg.title}" (${key}):`, e);
                return null;
            }
        }

        // Process in parallel batches of BATCH_SIZE
        const t0 = performance.now();
        for (let bi = 0; bi < cfgs.length; bi += BATCH_SIZE) {
            const batch = cfgs.slice(bi, bi + BATCH_SIZE);
            const results = await Promise.allSettled(batch.map(processOne));

            for (const r of results) {
                if (r.status !== "fulfilled" || !r.value) continue;
                const { key, result, type } = r.value;

                if (type === "multi") {
                    result.forEach(l => map.add(l));
                    S.reportLayerViews.set(key, result);
                } else {
                    map.add(result);
                    S.reportLayerViews.set(key, result);
                }
            }
        }

        console.log(`[buildReportDisplayLayers] ${cfgs.length} layer configs processed in ${((performance.now() - t0) / 1000).toFixed(1)}s`);
        ensureAoiOnTop();
    }

    // ── Public interface ─────────────────────────────────────────────

    /**
     * Initialize module with shared state.
     * @param {Object} state — mutable shared state object with view, map, config, etc.
     * @returns {Object} — all exported functions
     */
    function init(state) {
        S = state;
        return api;
    }

    const api = {
        init,

        // Renderers
        getPresetRenderer,
        getLayerGeometryType,
        makeRendererOpaque,
        thickenLayerSymbology,

        // AOI layer
        ensureAoiOnTop,
        updateAoiMask,
        hideAoiMask,
        setAoiGeometry,

        // Layer visibility
        ensureLayerVisibleAtScale,
        wireLayerUpdatingSpinner,

        // View / stationary
        waitForViewStationary,

        // Background-tab resilience
        waitForTabVisible,
        acquireWakeLock,
        releaseWakeLock,

        // Screenshot / capture
        waitForLayerReadyToCapture,
        captureScreenshotWithWait,
        hardRefreshLayer,

        // Report layers
        buildReportDisplayLayers
    };

    return api;
});
