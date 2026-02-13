/* ────────────────────────────────────────────────────────────────
   query-engine.js  –  Query, paging, coverage & analysis helpers
   AMD module loaded via dojoConfig  →  "app/query-engine"
   ──────────────────────────────────────────────────────────────── */

define([
    "app/config-helpers",
    "esri/layers/FeatureLayer",
    "esri/geometry/geometryEngine"
], function (configHelpers, FeatureLayer, geometryEngine) {
    "use strict";

    const { escapeHtml, formatNumber } = configHelpers;

    // ── Module-private state (set via init) ──
    let S;  // shared state proxy

    // ── Coverage cache (module-private) ──
    const SQM_PER_ACRE   = 4046.8564224;
    const coverageCache  = new Map();   // `${aoiKey}||${layerUrl}` → { acresCovered, pctAoiCovered }
    let   coverageAoiKey = "";

    // ── Helpers ──────────────────────────────────────────────────

    /**
     * Stable-enough signature for an AOI geometry (extent + rounded area).
     */
    function getAoiKey(geom) {
        try {
            const ex   = geom?.extent;
            const area = geometryEngine.geodesicArea(geom, "square-meters");
            return [ex?.xmin, ex?.ymin, ex?.xmax, ex?.ymax, Math.round(area)].join("|");
        } catch (e) {
            return String(Date.now());
        }
    }

    /**
     * Clear the coverage cache when the AOI changes.
     */
    function resetCoverageCacheForAoi(geom) {
        coverageCache.clear();
        coverageAoiKey = getAoiKey(geom);
    }

    // ── Pure utilities ──────────────────────────────────────────

    /**
     * Fisher–Yates partial-shuffle, return first `n` elements.
     */
    function sampleWithoutReplacement(arr, n) {
        const a = (arr || []).slice();
        if (a.length <= n) return a;
        for (let i = a.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [a[i], a[j]] = [a[j], a[i]];
        }
        return a.slice(0, n);
    }

    /**
     * Drop polygon features that only touch AOI at an edge / vertex
     * (intersection area ≈ 0).
     */
    function filterTouchingOnly(features, aoiGeom) {
        if (!features?.length || !aoiGeom) return features || [];
        const EPS = 0.000001; // sq metres

        return features.filter(f => {
            const g = f?.geometry;
            if (!g) return false;
            try {
                const inter = geometryEngine.intersect(aoiGeom, g);
                if (!inter) return false;
                const area = geometryEngine.geodesicArea(inter, "square-meters");
                return area > EPS;
            } catch (e) {
                return true;  // keep on error
            }
        });
    }

    // ── Paging helpers ──────────────────────────────────────────

    /**
     * Page through all features (attributes only).
     */
    async function queryAllFeaturesPaged(layer, baseQuery, pageSize, maxExportFeatures) {
        const all = [];
        let offset = 0;

        while (true) {
            const q = baseQuery.clone();
            q.num   = pageSize;
            q.start = offset;
            q.returnGeometry = false;

            const fs   = await layer.queryFeatures(q);
            const feats = (fs && fs.features) ? fs.features : [];
            all.push(...feats);

            if (feats.length < pageSize) break;
            offset += pageSize;
            if (maxExportFeatures && all.length >= maxExportFeatures) break;
        }
        return all;
    }

    /**
     * Page through all features WITH geometry (no attributes).
     */
    async function queryAllFeaturesPagedWithGeometry(layer, baseQuery, pageSize, maxExportFeatures) {
        const all = [];
        let offset = 0;

        while (true) {
            const q = baseQuery.clone();
            q.num   = pageSize;
            q.start = offset;
            q.returnGeometry    = true;
            q.outFields         = [];
            q.outSpatialReference = S.view?.spatialReference;

            const fs   = await layer.queryFeatures(q);
            const feats = (fs && fs.features) ? fs.features : [];
            all.push(...feats);

            if (feats.length < pageSize) break;
            offset += pageSize;
            if (maxExportFeatures && all.length >= maxExportFeatures) break;
        }
        return all;
    }

    // ── Geometry helpers ────────────────────────────────────────

    /**
     * Return the query geometry, optionally shrunk by 1 m for PLSS selections.
     */
    function getReportGeometry() {
        const selectionGeom = S.selectionGeom;
        if (!selectionGeom) return null;

        if (S.aoiSource !== "select") return selectionGeom;
        if (selectionGeom.type !== "polygon") return selectionGeom;

        try {
            const shrunk = geometryEngine.geodesicBuffer(selectionGeom, -1, "meters");
            return shrunk || selectionGeom;
        } catch (e) {
            console.warn("AOI shrink failed; using original geometry", e);
            return selectionGeom;
        }
    }

    /**
     * Union an array of geometries in manageable chunks.
     */
    function unionGeomsChunked(geoms) {
        const CHUNK = 25;
        let acc = null;

        for (let i = 0; i < geoms.length; i += CHUNK) {
            const chunk = geoms.slice(i, i + CHUNK).filter(Boolean);
            if (!chunk.length) continue;

            const u = geometryEngine.union(chunk);
            if (!acc) acc = u;
            else      acc = geometryEngine.union([acc, u]);
        }
        return acc;
    }

    // ── Single-layer query ──────────────────────────────────────

    /**
     * Query one layer for features intersecting a geometry.
     * Returns { title, url, count, features, layer, exportQuery }.
     */
    async function querySingleLayer(layerUrl, layerTitle, geom,
                                     spatialRel = "intersects", options = {}) {
        const applyTouchFilter = !!options.applyTouchFilter;
        const objectId         = options.objectId ?? null;
        const objectIdField    = options.objectIdField || "OBJECTID";

        const layer = new FeatureLayer({ url: layerUrl, outFields: ["*"] });
        const q     = layer.createQuery();
        q.outFields = ["*"];

        // ── AOI-source layer: return the exact clicked feature ──
        if (objectId != null) {
            await layer.load();
            const trueOidField = layer.objectIdField || objectIdField || "OBJECTID";

            const oidNum       = Number(objectId);
            const oidIsNumeric = Number.isFinite(oidNum);

            q.where = oidIsNumeric
                ? `${trueOidField} = ${oidNum}`
                : `${trueOidField} = '${String(objectId).replace(/'/g, "''")}'`;

            q.returnGeometry = false;
            q.outFields      = ["*"];

            const fs    = await layer.queryFeatures(q);
            const feats = fs?.features ?? [];
            return { title: layerTitle, url: layerUrl, count: feats.length, features: feats, layer, exportQuery: q };
        }

        // ── Default: geometry intersect ──
        q.geometry            = geom;
        q.spatialRelationship = spatialRel;
        q.returnGeometry      = applyTouchFilter;

        const count      = await layer.queryFeatureCount(q);
        const maxSamples = S.config.report?.maxSampleFeaturesPerLayer ?? 25;
        let   features   = [];

        if (count > 0 && maxSamples > 0) {
            const q2 = q.clone();
            q2.num   = Math.min(maxSamples, 2000);
            const fs = await layer.queryFeatures(q2);
            const raw = fs?.features ?? [];
            features = applyTouchFilter ? filterTouchingOnly(raw, geom) : raw;
        }

        return { title: layerTitle, url: layerUrl, count, features, layer, exportQuery: q };
    }

    // ── Elevation stats ─────────────────────────────────────────

    /**
     * Compute elevation statistics (min/max/mean) for an AOI
     * via an ImageServer's computeHistograms endpoint (POST).
     */
    async function computeElevationStats(imageServerUrl, geometry) {
        if (!imageServerUrl || !geometry) return null;

        try {
            const geomJson = JSON.stringify(geometry.toJSON ? geometry.toJSON() : geometry);

            const url    = `${imageServerUrl}/computeHistograms`;
            const params = new URLSearchParams({
                f:            "json",
                geometry:     geomJson,
                geometryType: "esriGeometryPolygon"
            });

            const response = await fetch(url, {
                method:  "POST",
                headers: { "Content-Type": "application/x-www-form-urlencoded" },
                body:    params.toString()
            });
            if (!response.ok) throw new Error(`HTTP ${response.status}`);

            const data = await response.json();

            if (data.histograms && data.histograms.length > 0) {
                const hist    = data.histograms[0];
                const minElev = hist.min;
                const maxElev = hist.max;

                let mean = null;
                if (hist.counts && hist.size) {
                    const binWidth = (maxElev - minElev) / hist.counts.length;
                    let sum = 0, total = 0;
                    for (let i = 0; i < hist.counts.length; i++) {
                        const binCenter = minElev + (i + 0.5) * binWidth;
                        sum   += binCenter * hist.counts[i];
                        total += hist.counts[i];
                    }
                    mean = total > 0 ? sum / total : null;
                }

                return {
                    min: minElev,
                    max: maxElev,
                    mean,
                    minFt:             minElev * 3.28084,
                    maxFt:             maxElev * 3.28084,
                    meanFt:            mean ? mean * 3.28084 : null,
                    elevationChange:   maxElev - minElev,
                    elevationChangeFt: (maxElev - minElev) * 3.28084
                };
            }
            return null;
        } catch (e) {
            console.warn("Failed to compute elevation statistics:", e);
            return null;
        }
    }

    // ── Coverage stats ──────────────────────────────────────────

    /**
     * Compute { acresCovered, pctAoiCovered } for a single report item.
     */
    async function computeLayerCoverageStats(item, aoiGeom) {
        if (!item || !item._layer || !item._exportQuery || !aoiGeom) return null;

        const layerUrlKey = String(item.url || "").replace(/\/+$/, "");
        const aoiKey      = coverageAoiKey || getAoiKey(aoiGeom);
        const cacheKey    = `${aoiKey}||${layerUrlKey}`;

        if (coverageCache.has(cacheKey)) return coverageCache.get(cacheKey);

        // AOI area (sq m)
        let aoiAreaSqm = 0;
        try { aoiAreaSqm = Math.max(0, geometryEngine.geodesicArea(aoiGeom, "square-meters")); }
        catch (e) { aoiAreaSqm = 0; }
        if (!aoiAreaSqm) return { acresCovered: 0, pctAoiCovered: 0 };

        const pageSize  = S.config.report?.pageSize ?? 1000;
        const maxExport = S.config.report?.maxExportFeatures ?? 50000;

        const feats = await queryAllFeaturesPagedWithGeometry(item._layer, item._exportQuery, pageSize, maxExport);

        // Intersect each feature with AOI
        const interGeoms = [];
        for (const f of feats) {
            const g = f?.geometry;
            if (!g) continue;
            try {
                const inter = geometryEngine.intersect(aoiGeom, g);
                if (!inter) continue;
                const area = geometryEngine.geodesicArea(inter, "square-meters");
                if (area <= 0) continue;
                interGeoms.push(inter);
            } catch (e) { /* skip bad geom */ }
        }

        if (!interGeoms.length) return { acresCovered: 0, pctAoiCovered: 0 };

        let unionGeom = null;
        try { unionGeom = unionGeomsChunked(interGeoms); }
        catch (e) { unionGeom = null; }

        let coveredSqm = 0;
        try { coveredSqm = unionGeom ? Math.max(0, geometryEngine.geodesicArea(unionGeom, "square-meters")) : 0; }
        catch (e) { coveredSqm = 0; }

        const acresCovered  = coveredSqm / SQM_PER_ACRE;
        const pctAoiCovered = Math.min(100, Math.max(0, (coveredSqm / aoiAreaSqm) * 100));

        const out = { acresCovered, pctAoiCovered };
        coverageCache.set(cacheKey, out);
        return out;
    }

    // ── Sample table ────────────────────────────────────────────

    /**
     * Build a small HTML sample-data table from an array of features.
     */
    function makeTable(features, maxFields, totalCount) {
        if (!features || !features.length)
            return `<div class="small">No sample features fetched.</div>`;

        const picked   = sampleWithoutReplacement(features, 4);
        const attrs0   = picked[0].attributes || {};
        const keysAll  = Object.keys(attrs0);
        const keys     = keysAll;

        const defaultVisibleCols = 5;
        const colPx       = 140;
        const minTableWidth = Math.max(520, keys.length * colPx);

        const th = keys.map(k => `<th title="${escapeHtml(k)}">${escapeHtml(k)}</th>`).join("");

        const rows = picked.map(f => {
            const a   = f.attributes || {};
            const tds = keys.map(k => {
                const raw    = (a[k] == null) ? "" : String(a[k]);
                const maxLen = Math.max(4, String(k).length);
                let shown    = raw;
                if (raw.length > maxLen) shown = raw.slice(0, Math.max(1, maxLen - 1)) + "\u2026";

                const safeFull  = escapeHtml(raw);
                const safeShown = escapeHtml(shown);
                return `<td title="${safeFull}">${safeShown}</td>`;
            }).join("");
            return `<tr>${tds}</tr>`;
        }).join("");

        let moreRowHtml = "";
        const total = Number(totalCount || 0);
        const shown = picked.length;
        if (total > shown) {
            const more = total - shown;
            const msg  = `\u2026 ${more} more row${more === 1 ? "" : "s"} (see FULL export)`;
            moreRowHtml = `<tr><td colspan="${keys.length}" class="small" style="opacity:.8;">${escapeHtml(msg)}</td></tr>`;
        }

        const colHint = (keys.length > defaultVisibleCols)
            ? `<div class="small table-hint">Table has ${keys.length} columns \u2014 scroll \u2192 for more.</div>`
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

    // ── Per-feature breakdown ───────────────────────────────────

    /**
     * Build an interactive per-feature data table for the final report.
     * Each row = one feature with ALL attrs + Acres + % AOI.
     * Supports sorting, column hide/show, and horizontal scrolling.
     */
    async function buildPerFeatureTable(item, aoiGeom, tableId) {
        if (!item || !item._layer || !item._exportQuery || !aoiGeom) return "";
        if ((item.count || 0) < 1) return "";

        let aoiAreaSqm = 0;
        try { aoiAreaSqm = Math.max(0, geometryEngine.geodesicArea(aoiGeom, "square-meters")); }
        catch (e) { aoiAreaSqm = 0; }
        if (!aoiAreaSqm) return "";

        const pageSize  = S.config.report?.pageSize ?? 1000;
        const maxExport = S.config.report?.maxExportFeatures ?? 50000;

        // Query all features WITH geometry + attributes
        let feats = [];
        try {
            const q = item._exportQuery.clone();
            q.returnGeometry      = true;
            q.outFields           = ["*"];
            q.outSpatialReference = S.view?.spatialReference;

            const all = [];
            let offset = 0;
            while (true) {
                const pq = q.clone();
                pq.num   = pageSize;
                pq.start = offset;
                const fs    = await item._layer.queryFeatures(pq);
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

        if (feats.length < 1) return "";

        // Determine relevant attribute columns
        const skipPatterns = /^(objectid|oid|fid|shape|shape_area|shape_length|shape\.area|shape\.len|globalid|st_area|st_length|st_perimeter)$/i;
        const allKeys = new Set();
        for (const f of feats) {
            if (!f.attributes) continue;
            for (const k of Object.keys(f.attributes)) {
                if (!skipPatterns.test(k)) allKeys.add(k);
            }
        }

        const attrKeys = Array.from(allKeys);

        // Build rows: compute per-feature intersection area
        const tableRows = [];
        for (const f of feats) {
            const attrs = f.attributes || {};
            const geom  = f.geometry;

            let acresCovered = 0;
            let pctAoi       = 0;

            if (geom) {
                try {
                    const inter = geometryEngine.intersect(aoiGeom, geom);
                    if (inter) {
                        const areaSqm = Math.max(0, geometryEngine.geodesicArea(inter, "square-meters"));
                        acresCovered  = areaSqm / SQM_PER_ACRE;
                        pctAoi        = Math.min(100, Math.max(0, (areaSqm / aoiAreaSqm) * 100));
                    }
                } catch (e) { /* skip bad geometry */ }
            }

            tableRows.push({ attrs, acresCovered, pctAoi });
        }

        // Sort by acres descending
        tableRows.sort((a, b) => b.acresCovered - a.acresCovered);

        // Build interactive table HTML
        const tId = tableId != null ? tableId : Math.floor(Math.random() * 100000);
        const colLabels = [...attrKeys, 'Acres', '% of AOI'];
        const totalCols = colLabels.length;
        const tableTitle = feats.length === 1 ? 'Feature Attributes' : 'Per-Feature Breakdown';

        const thCells = colLabels.map((label, ci) =>
            `<th data-col="${ci}" data-label="${escapeHtml(label)}" data-sort-dir="none" onclick="sortInteractiveTable(this)" title="Click to sort">${escapeHtml(label)} <span class="sort-arrow">\u21C5</span> <button class="col-hide-btn" onclick="event.stopPropagation(); hideColumn('tbl-${tId}',${ci});" title="Hide this column">\u2716</button></th>`
        ).join("");
        const headerHtml = `<tr>${thCells}</tr>`;

        let hasSliverWarning = false;
        const bodyHtml = tableRows.map(row => {
            const tdCells = attrKeys.map((k, ci) => {
                const raw = (row.attrs[k] == null) ? "" : String(row.attrs[k]);
                const display = raw.length > 100 ? raw.slice(0, 99) + "\u2026" : raw;
                return `<td data-col="${ci}" data-sort-val="${escapeHtml(raw)}" title="${escapeHtml(raw)}">${escapeHtml(display)}</td>`;
            });
            const isSliverWarning = row.pctAoi < 3;
            if (isSliverWarning) hasSliverWarning = true;
            const acresIdx = attrKeys.length;
            const pctIdx = attrKeys.length + 1;
            tdCells.push(`<td data-col="${acresIdx}" data-sort-val="${row.acresCovered.toFixed(6)}" style="text-align:right;">${formatNumber(row.acresCovered, 2)}</td>`);
            const warningIcon = isSliverWarning ? '<span style="color:#856404;">\u26A0\uFE0F</span> ' : '';
            tdCells.push(`<td data-col="${pctIdx}" data-sort-val="${row.pctAoi.toFixed(6)}" style="text-align:right;">${warningIcon}${formatNumber(row.pctAoi, 2)}%</td>`);
            const rowStyle = isSliverWarning ? ' style="background-color:#fff3cd;"' : '';
            return `<tr${rowStyle}>${tdCells.join("")}</tr>`;
        }).join("");

        const sliverNote = hasSliverWarning
            ? `<div style="margin-top:8px; font-size:10px; color:#856404; background:#fff3cd; padding:6px 8px; border-radius:4px;">
                <b>\u26A0\uFE0F Note:</b> Highlighted rows cover less than 3% of the AOI and may represent slivers, boundary artifacts, or minor overlaps rather than meaningful intersections.
            </div>`
            : '';

        return `
            <div class="interactive-table-wrapper" id="tbl-${tId}" style="margin-top:16px;">
                <div class="table-toolbar">
                    <b>${escapeHtml(tableTitle)}</b>
                    <span style="font-size:12px;color:#5a5a5a;margin-left:8px;">(${feats.length} feature${feats.length !== 1 ? 's' : ''}, ${totalCols} columns \u2014 click \u2716 on a header to hide)</span>
                </div>
                <div class="hidden-cols-bar" style="display:none;"></div>
                <div class="table-scroll">
                    <table class="interactive-table">
                        <thead>${headerHtml}</thead>
                        <tbody>${bodyHtml}</tbody>
                    </table>
                </div>
                ${sliverNote}
            </div>
        `;
    }

    // ── Module entry point ──────────────────────────────────────

    function init(state) {
        S = state;

        return {
            // Paging
            queryAllFeaturesPaged,
            queryAllFeaturesPagedWithGeometry,

            // Geometry
            filterTouchingOnly,
            getReportGeometry,
            unionGeomsChunked,

            // Single-layer query
            querySingleLayer,

            // Elevation
            computeElevationStats,

            // Coverage
            computeLayerCoverageStats,
            buildPerFeatureTable,

            // Coverage cache
            getAoiKey,
            resetCoverageCacheForAoi,
            SQM_PER_ACRE,

            // Utility
            sampleWithoutReplacement,
            makeTable
        };
    }

    return { init };
});
