/**
 * visual-report.js  \u2013  AMD module for Visual Report (per-layer map screenshots)
 *
 * Exports:
 *   setVisualStatus, renderVisualSummary, generateVisualReportData
 *
 * Usage (inside the main require callback):
 *   const vr = visualReportModule.init(state, deps);
 *   const { setVisualStatus, renderVisualSummary, generateVisualReportData } = vr;
 */
define([
    "app/config-helpers"
], function (configHelpers) {
    "use strict";

    const { escapeHtml, formatNumber } = configHelpers;

    // \u2500\u2500 Module-private state (set by init) \u2500\u2500
    let S;               // shared state proxy
    let mapUtils;        // map-utils API
    let queryEngine;     // query-engine API
    let ImageryLayer;    // Esri constructor
    let FeatureLayer;    // Esri constructor
    let geometryEngine;  // Esri geometryEngine

    // External callbacks injected from app.js
    let _isReportCanceled; // (myOp) => boolean

    // DOM references (resolved once in init)
    let visualReportStatusEl  = null;
    let visualReportMapWrapEl = null;
    let visualReportOutputsEl = null;
    let visualReportSummaryEl = null;

    // ── Background-tab resilience helpers ──
    let _wakeLock = null;

    /** Request a Wake Lock to keep the screen active during analysis. */
    async function requestWakeLock() {
        try {
            if (navigator.wakeLock) {
                _wakeLock = await navigator.wakeLock.request("screen");
                console.log("[visual-report] Wake Lock acquired");
            }
        } catch (e) {
            console.warn("[visual-report] Wake Lock request failed:", e.message);
        }
    }

    /** Release the Wake Lock when analysis is done. */
    async function releaseWakeLock() {
        try {
            if (_wakeLock) { await _wakeLock.release(); _wakeLock = null; console.log("[visual-report] Wake Lock released"); }
        } catch (e) { /* ignore */ }
    }

    /**
     * Retry wrapper for view.takeScreenshot.
     * Browsers throttle rendering in hidden/background tabs.  If the first
     * attempt returns no data we wait a moment and retry.
     */
    async function takeScreenshotSafe(view, opts, maxRetries) {
        if (maxRetries === undefined) maxRetries = 3;
        for (var attempt = 1; attempt <= maxRetries; attempt++) {
            try {
                var ss = await view.takeScreenshot(opts);
                if (ss && ss.dataUrl) return ss;
            } catch (e) {
                console.warn("[visual-report] takeScreenshot attempt " + attempt + " failed:", e.message);
            }
            // Wait before retrying — give browser a chance to repaint
            await new Promise(function (r) { setTimeout(r, 800 * attempt); });
            // Force a frame to encourage the browser to paint
            await new Promise(function (r) { requestAnimationFrame(r); });
        }
        // Final fallback attempt
        return await view.takeScreenshot(opts);
    }

    // \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    // Public API
    // \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    function setVisualStatus(msg) {
        if (visualReportStatusEl) visualReportStatusEl.textContent = msg || "";
    }

    function clearVisualReport() {
        if (visualReportOutputsEl) visualReportOutputsEl.innerHTML = "";
        if (visualReportMapWrapEl) visualReportMapWrapEl.classList.add("hidden");
    }

    function renderVisualSummary() {
        if (!visualReportSummaryEl) return;

        const selectionGeom = S.selectionGeom;
        const lastReportRowsByLayer = S.lastReportRowsByLayer;

        if (!selectionGeom) {
            visualReportSummaryEl.innerHTML = '<div class="small">(No AOI selected.)</div>';
            return;
        }

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            visualReportSummaryEl.innerHTML = '<div class="small">(Run the report to populate layer counts.)</div>';
            return;
        }

        const totalLayers = lastReportRowsByLayer.length;
        const layersWithHits = lastReportRowsByLayer.filter(function (x) { return (x.count || 0) > 0; });
        const totalHits = lastReportRowsByLayer.reduce(function (sum, x) { return sum + (x.count || 0); }, 0);

        const top = layersWithHits
            .slice()
            .sort(function (a, b) { return (b.count || 0) - (a.count || 0); })
            .slice(0, 12);

        var listHtml = top.length
            ? '<div style="margin-top:8px;">'
                + top.map(function (x) {
                    return '<div class="small">\u2022 ' + escapeHtml(x.title) + ' <span class="mono">(' + x.count + ')</span></div>';
                }).join("")
                + '</div>'
            : '<div class="small" style="margin-top:8px;">(No intersect hits.)</div>';

        visualReportSummaryEl.innerHTML =
            '<div class="small">Layers queried: <b>' + totalLayers + '</b></div>' +
            '<div class="small">Layers with hits: <b>' + layersWithHits.length + '</b></div>' +
            '<div class="small">Total intersecting features (sum of counts): <b>' + totalHits + '</b></div>' +
            listHtml;
    }

    async function generateVisualReportData(myOp, modal) {
        if (modal === undefined) modal = null;

        // Request Wake Lock to prevent browser throttling during screenshots
        await requestWakeLock();

        const view = S.view;
        const selectionGeom = S.selectionGeom;
        const lastReportRowsByLayer = S.lastReportRowsByLayer;
        const config = S.config;
        const aoiLayer = S.aoiLayer;
        const aoiMaskLayer = S.aoiMaskLayer;
        const alwaysVisibleLayers = S.alwaysVisibleLayers;
        const layerCfgByUrl = S.layerCfgByUrl;

        const {
            getPresetRenderer, getLayerGeometryType, makeRendererOpaque,
            ensureAoiOnTop, updateAoiMask, hideAoiMask,
            waitForViewStationary, waitForLayerReadyToCapture
        } = mapUtils;

        const {
            computeElevationStats, computeLayerCoverageStats, SQM_PER_ACRE
        } = queryEngine;

        if (!view) return;

        if (!selectionGeom) {
            setVisualStatus("No AOI selected.");
            return;
        }

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            setVisualStatus("No query results available.");
            return;
        }

        setVisualStatus("Generating maps for intersecting layers\u2026");

        if (visualReportMapWrapEl) visualReportMapWrapEl.classList.add("hidden");
        if (visualReportOutputsEl) visualReportOutputsEl.innerHTML = "";

        try {
            // Only layers with real intersect hits AND usable query objects
            // Include ImageServer layers even though they don't have _layer/_exportQuery
            var targets = lastReportRowsByLayer
                .filter(function (x) { return (x && (x.count || 0) > 0); })
                .filter(function (x) { return (x && x._layer && x._exportQuery) || (x && x.__isImageService); })
                .filter(function (x) { return !(x.title && x.title.toLowerCase().includes("state boundaries")); })
                .filter(function (x) { return !(x.title && x.title.toLowerCase().includes("administrative unit")); });

            if (!targets.length) {
                setVisualStatus("No intersecting layers to map (all counts are 0).");
                if (visualReportMapWrapEl) visualReportMapWrapEl.classList.remove("hidden");
                return;
            }

            // Zoom to AOI with padding once (we'll keep the view there)
            var paddingFactor = (config && config.visualReport && config.visualReport.paddingFactor != null)
                ? config.visualReport.paddingFactor : 1.25;
            var width = (config && config.visualReport && config.visualReport.screenshotWidth != null)
                ? config.visualReport.screenshotWidth : 1400;

            // Compute and lock a single extent for ALL screenshots
            var fixedExtent = null;
            var ext = selectionGeom && selectionGeom.extent;

            if (ext && ext.expand) {
                fixedExtent = ext.expand(paddingFactor);
                await view.goTo(fixedExtent, { animate: true, duration: 450 });
            } else {
                await view.goTo(selectionGeom, { animate: true, duration: 450 });
            }

            // Snapshot current layer visibility so we can restore after each screenshot
            var allLayers = view.map.layers.toArray();
            var visSnapshot = allLayers.map(function (l) { return { layer: l, visible: l.visible }; });

            // Helper to hide everything except AOI + basemap overlay + a temp target layer
            function setVisibilityForScreenshot(tempLayer) {
                for (var li = 0; li < allLayers.length; li++) {
                    var l = allLayers[li];
                    // Keep AOI layer visible
                    if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }
                    // Keep AOI mask layer visible for per-layer maps
                    if (aoiMaskLayer && l === aoiMaskLayer) { l.visible = true; continue; }
                    // Keep TileLayers visible by default
                    if (l && l.type === "tile") { l.visible = true; continue; }
                    // Keep always-visible layers (e.g. BLM Admin Units) visible
                    if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }
                    // Hide everything else
                    l.visible = false;
                }
                if (tempLayer) tempLayer.visible = true;
                // Show AOI mask to lighten areas outside AOI
                updateAoiMask(true);
                ensureAoiOnTop();
            }

            function restoreVisibility() {
                visSnapshot.forEach(function (s) { try { s.layer.visible = s.visible; } catch (e) { } });
                hideAoiMask();
                ensureAoiOnTop();
            }

            // AOI area in acres (used for context)
            var aoiAcres = 0;
            try {
                var aoiSqm = Math.max(0, geometryEngine.geodesicArea(selectionGeom, "square-meters"));
                aoiAcres = aoiSqm / SQM_PER_ACRE;
            } catch (e) {
                aoiAcres = 0;
            }

            var outCards = [];

            for (var i = 0; i < targets.length; i++) {
                if (_isReportCanceled(myOp)) {
                    setVisualStatus("canceled");
                    break;
                }

                var item = targets[i];

                // Update modal progress
                if (modal) {
                    var progress = 60 + (25 * (i / targets.length)); // 60% -> 85%
                    modal.setProgress(progress);
                    modal.setStep("Step 3/4: Generating map " + (i + 1) + "/" + targets.length + "...");
                    modal.addLog("Generating map for: " + item.title);
                }

                setVisualStatus("Generating map " + (i + 1) + " / " + targets.length + "\u2026");

                // Handle ImageServer layers differently
                if (item.__isImageService) {
                    var imgLayerOpts = {
                        url: item.url,
                        title: item.title,
                        visible: true
                    };
                    if (item.__renderingRule) {
                        imgLayerOpts.renderingRule = { functionName: item.__renderingRule };
                    }
                    var temp = new ImageryLayer(imgLayerOpts);
                    view.map.add(temp);

                    try {
                        setVisibilityForScreenshot(temp);
                        await waitForLayerReadyToCapture(temp, view, { timeoutMs: 10000 });
                        await view.goTo(fixedExtent, { animate: false });
                        await waitForViewStationary(2500);

                        var ss = await takeScreenshotSafe(view, { format: "png", quality: 100, width: width });
                        if (!ss || !ss.dataUrl) throw new Error("Screenshot failed");
                        var dataUrl = ss.dataUrl;

                        var meta = item.__serviceMeta || {};

                        // Compute elevation statistics for the AOI
                        var elevStats = await computeElevationStats(item.url, selectionGeom);

                        var elevStatsHtml = '';
                        if (elevStats) {
                            elevStatsHtml =
                                '<tr><td colspan="2" style="font-weight:600; padding-top:8px;">Elevation (AOI)</td></tr>' +
                                '<tr><td>Min</td><td>' + formatNumber(elevStats.minFt, 0) + ' ft</td></tr>' +
                                '<tr><td>Max</td><td>' + formatNumber(elevStats.maxFt, 0) + ' ft</td></tr>' +
                                '<tr><td>Change</td><td>' + formatNumber(elevStats.elevationChangeFt, 0) + ' ft</td></tr>';
                        }

                        outCards.push(
                            '<div class="visual-output-card">' +
                              '<div class="visual-output-title">' + escapeHtml(item.title) + '</div>' +
                              '<img class="visual-output-img" src="' + dataUrl + '" alt="' + escapeHtml(item.title) + '" />' +
                              '<div class="visual-output-meta">' +
                                '<table>' +
                                  '<tr><td>Type</td><td>Image Service</td></tr>' +
                                  '<tr><td>Service</td><td>' + escapeHtml(meta.name || item.title) + '</td></tr>' +
                                  elevStatsHtml +
                                  (meta.copyright ? '<tr><td>Source</td><td>' + escapeHtml(meta.copyright) + '</td></tr>' : '') +
                                '</table>' +
                              '</div>' +
                            '</div>'
                        );
                    } finally {
                        try { view.map.remove(temp); } catch (e) { }
                        restoreVisibility();
                    }
                    continue;
                }

                // Create a temporary layer for this URL
                var tempGeomType = await getLayerGeometryType(item.url);
                var tempOpts = {
                    url: item.url,
                    title: item.title,
                    outFields: ["*"],
                    visible: true,
                    opacity: 0.8
                };
                var temp2 = new FeatureLayer(tempOpts);

                // Always override scale to ensure layer draws at any zoom
                temp2.minScale = 0;
                temp2.maxScale = 0;

                // Add temp, hide everything else, screenshot, then remove temp
                view.map.add(temp2);
                try {
                    setVisibilityForScreenshot(temp2);

                    // Wait for layer to load
                    try { await temp2.when(); } catch (e) { }

                    // Wait until layerView is not suspended AND not updating (best effort)
                    try {
                        var lv = await view.whenLayerView(temp2);

                        // Wait for suspended -> false OR timeout
                        if (lv && lv.suspended) {
                            await new Promise(function (resolve) {
                                var h = lv.watch("suspended", function (s) {
                                    if (!s) { h.remove(); resolve(); }
                                });
                                window.setTimeout(function () { try { h.remove(); } catch (e) { } resolve(); }, 4000);
                            });
                        }

                        // Wait for updating -> false OR timeout
                        if (lv && lv.updating) {
                            await new Promise(function (resolve) {
                                var h = lv.watch("updating", function (u) {
                                    if (!u) { h.remove(); resolve(); }
                                });
                                window.setTimeout(function () { try { h.remove(); } catch (e) { } resolve(); }, 6000);
                            });
                        }
                    } catch (e) { }

                    // Re-apply locked extent to guarantee identical framing
                    if (fixedExtent) {
                        await view.goTo(fixedExtent, { animate: false });
                    }

                    var ss2 = await takeScreenshotSafe(view, { format: "png", quality: 100, width: width });
                    var dataUrl2 = ss2 && ss2.dataUrl;
                    if (!dataUrl2) throw new Error("Screenshot failed (no dataUrl).");

                    // Compute coverage stats (acres + % AOI covered)
                    var cov = await computeLayerCoverageStats(item, selectionGeom);

                    // Render a card for this layer
                    var acresCovered = cov ? cov.acresCovered : 0;
                    var pctCovered = cov ? cov.pctAoiCovered : 0;

                    // Check for low coverage warning — only for polygon layers
                    var isPolygonLayer = tempGeomType && String(tempGeomType).toLowerCase().indexOf('polygon') !== -1;
                    var isSingleFeatureLowCoverage = isPolygonLayer && (item.count === 1 && pctCovered < 3);
                    var lowCoverageWarningHtml = isSingleFeatureLowCoverage
                        ? '<div style="margin-top:8px; padding:6px; background-color:#fff3cd; border:1px solid #ffc107; border-radius:4px; font-size:11px;">' +
                            '<span style="color:#856404;">\u26a0\ufe0f Low coverage (&lt;3%) \u2014 possible sliver or boundary artifact</span>' +
                          '</div>'
                        : "";

                    outCards.push(
                        '<div class="visual-output-card">' +
                          '<div class="visual-output-title">' + escapeHtml(item.title) + '</div>' +
                          '<img class="visual-output-img" src="' + dataUrl2 + '" alt="AOI + ' + escapeHtml(item.title) + '" />' +
                          '<div class="visual-output-meta">' +
                            '<table>' +
                              '<tr><td>AOI area</td><td><span class="mono">' + formatNumber(aoiAcres, 2) + '</span> acres</td></tr>' +
                              '<tr><td>Intersecting features</td><td><span class="mono">' + escapeHtml(String(item.count || 0)) + '</span></td></tr>' +
                              '<tr><td>AOI covered by layer</td><td><span class="mono">' + formatNumber(acresCovered, 2) + '</span> acres</td></tr>' +
                              '<tr><td>% AOI covered</td><td><span class="mono">' + formatNumber(pctCovered, 2) + '</span>%' +
                                  (isSingleFeatureLowCoverage ? ' <span style="color:#856404;" title="Low coverage \u2014 possible sliver">\u26a0\ufe0f</span>' : '') +
                              '</td></tr>' +
                            '</table>' +
                            lowCoverageWarningHtml +
                          '</div>' +
                        '</div>'
                    );

                } finally {
                    // Remove temp layer and restore visibility
                    try { view.map.remove(temp2); } catch (e) { }
                    restoreVisibility();
                }
            }

            if (visualReportOutputsEl) visualReportOutputsEl.innerHTML = outCards.join("");
            if (visualReportMapWrapEl) visualReportMapWrapEl.classList.remove("hidden");

            // Keep existing summary panel behavior
            renderVisualSummary();

            setVisualStatus("Maps generated.");
        } catch (e) {
            console.error(e);
            setVisualStatus("Failed to generate maps (see console).");
        } finally {
            // Always release the Wake Lock when done
            await releaseWakeLock();
        }
    }

    // \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
    // init(state, deps)  \u2013  called once from app.js
    // \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500

    return {
        init: function (state, deps) {
            S = state;
            mapUtils        = deps.mapUtils;
            queryEngine     = deps.queryEngine;
            ImageryLayer    = deps.ImageryLayer;
            FeatureLayer    = deps.FeatureLayer;
            geometryEngine  = deps.geometryEngine;
            _isReportCanceled = deps.isReportCanceled;

            // Resolve DOM references once
            visualReportStatusEl  = document.getElementById("visualReportStatus");
            visualReportMapWrapEl = document.getElementById("visualReportMapWrap");
            visualReportOutputsEl = document.getElementById("visualReportOutputs");
            visualReportSummaryEl = document.getElementById("visualReportSummary");

            return {
                setVisualStatus: setVisualStatus,
                clearVisualReport: clearVisualReport,
                renderVisualSummary: renderVisualSummary,
                generateVisualReportData: generateVisualReportData
            };
        }
    };
});
