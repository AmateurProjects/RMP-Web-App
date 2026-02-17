/* ────────────────────────────────────────────────────────────────
   query-engine.js  –  Query, paging, coverage & analysis helpers
   AMD module loaded via dojoConfig  →  "app/query-engine"
   ──────────────────────────────────────────────────────────────── */

define([
    "app/config-helpers",
    "esri/layers/FeatureLayer",
    "esri/geometry/geometryEngine",
    "esri/geometry/Extent"
], function (configHelpers, FeatureLayer, geometryEngine, Extent) {
    "use strict";

    const { escapeHtml, formatNumber } = configHelpers;

    // ── Module-private state (set via init) ──
    let S;  // shared state proxy

    // ── Coverage cache (module-private) ──
    const SQM_PER_ACRE   = 4046.8564224;
    const METERS_PER_FOOT = 0.3048;
    const FEET_PER_MILE   = 5280;
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

    // ── Chunked query for large AOIs ─────────────────────────────

    /**
     * Subdivide a large AOI into a grid of cells, query each cell
     * separately, then merge & deduplicate results by ObjectID.
     * Falls back to a single unchunked query if the AOI is below threshold.
     *
     * Returns the same shape as querySingleLayer:
     *   { title, url, count, features, layer, exportQuery }
     */
    async function querySingleLayerChunked(layerUrl, layerTitle, aoiGeom,
                                            spatialRel = "intersects", options = {}) {
        const gridSize = S.config.report?.aoiChunkGridSize ?? 4;
        const ext = aoiGeom.extent;
        const cellW = (ext.xmax - ext.xmin) / gridSize;
        const cellH = (ext.ymax - ext.ymin) / gridSize;
        const sr = ext.spatialReference;

        // Build grid cells clipped to the AOI
        const cells = [];
        for (let row = 0; row < gridSize; row++) {
            for (let col = 0; col < gridSize; col++) {
                const cellExtent = new Extent({
                    xmin: ext.xmin + col * cellW,
                    ymin: ext.ymin + row * cellH,
                    xmax: ext.xmin + (col + 1) * cellW,
                    ymax: ext.ymin + (row + 1) * cellH,
                    spatialReference: sr
                });
                // Clip cell to the actual AOI polygon
                const clipped = geometryEngine.intersect(aoiGeom, cellExtent);
                if (clipped) cells.push(clipped);
            }
        }

        if (!cells.length) {
            // Fallback: use original geometry
            return querySingleLayer(layerUrl, layerTitle, aoiGeom, spatialRel, options);
        }

        const applyTouchFilter = !!options.applyTouchFilter;
        const layer = new FeatureLayer({ url: layerUrl, outFields: ["*"] });

        // Query counts for each cell in parallel (batched 6 at a time)
        const CELL_BATCH = 6;
        const oidSet = new Set();
        let allSampleFeatures = [];
        const maxSamples = S.config.report?.maxSampleFeaturesPerLayer ?? 25;

        for (let ci = 0; ci < cells.length; ci += CELL_BATCH) {
            const cellBatch = cells.slice(ci, ci + CELL_BATCH);
            const batchResults = await Promise.allSettled(cellBatch.map(async (cellGeom) => {
                const q = layer.createQuery();
                q.outFields = ["*"];
                q.geometry = cellGeom;
                q.spatialRelationship = spatialRel;
                q.returnGeometry = applyTouchFilter;

                // Get object IDs to deduplicate across cells
                let oids = [];
                try {
                    const oidResult = await layer.queryObjectIds(q);
                    oids = oidResult || [];
                } catch (e) {
                    // If queryObjectIds fails, fall back to count
                    const ct = await layer.queryFeatureCount(q);
                    return { count: ct, oids: [], features: [] };
                }

                // Fetch sample features if we still need more
                let feats = [];
                if (oids.length > 0 && allSampleFeatures.length < maxSamples) {
                    const q2 = q.clone();
                    q2.num = Math.min(maxSamples, 2000);
                    try {
                        const fs = await layer.queryFeatures(q2);
                        feats = fs?.features ?? [];
                    } catch (e) {
                        // Sample fetch failed, that's okay
                    }
                }

                return { count: 0, oids, features: feats };
            }));

            for (const r of batchResults) {
                if (r.status !== "fulfilled") continue;
                const { oids, features, count } = r.value;

                if (oids.length > 0) {
                    for (const oid of oids) oidSet.add(oid);
                }

                if (features.length > 0 && allSampleFeatures.length < maxSamples) {
                    // Deduplicate sample features by OID
                    for (const f of features) {
                        if (allSampleFeatures.length >= maxSamples) break;
                        const fOid = f.attributes?.[layer.objectIdField || "OBJECTID"];
                        if (fOid != null && oidSet.has(fOid)) {
                            // Check if we already have this feature in our sample
                            const alreadyHave = allSampleFeatures.some(sf =>
                                sf.attributes?.[layer.objectIdField || "OBJECTID"] === fOid
                            );
                            if (!alreadyHave) allSampleFeatures.push(f);
                        } else {
                            allSampleFeatures.push(f);
                        }
                    }
                }
            }
        }

        // Apply touch filter if needed
        if (applyTouchFilter && allSampleFeatures.length > 0) {
            allSampleFeatures = filterTouchingOnly(allSampleFeatures, aoiGeom);
        }

        const dedupedCount = oidSet.size;

        // Build a combined export query using the full AOI geometry
        const exportQ = layer.createQuery();
        exportQ.outFields = ["*"];
        exportQ.geometry = aoiGeom;
        exportQ.spatialRelationship = spatialRel;
        exportQ.returnGeometry = false;

        return {
            title: layerTitle,
            url: layerUrl,
            count: dedupedCount,
            features: allSampleFeatures.slice(0, maxSamples),
            layer,
            exportQuery: exportQ
        };
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

        // Determine geometry type
        let geomClass = 'unknown'; // 'polygon', 'polyline', 'point'
        try {
            const layerGeomType = (item._layer.geometryType || '').toLowerCase();
            if (layerGeomType.includes('polygon')) geomClass = 'polygon';
            else if (layerGeomType.includes('polyline') || layerGeomType.includes('line')) geomClass = 'polyline';
            else if (layerGeomType.includes('point')) geomClass = 'point';
        } catch (e) {
            if (feats[0] && feats[0].geometry && feats[0].geometry.type) {
                const ft = feats[0].geometry.type.toLowerCase();
                if (ft.includes('polygon')) geomClass = 'polygon';
                else if (ft.includes('polyline') || ft.includes('line')) geomClass = 'polyline';
                else if (ft.includes('point')) geomClass = 'point';
            }
        }
        const isPolygonLayer  = (geomClass === 'polygon');
        const isPolylineLayer = (geomClass === 'polyline');

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

        // Build rows: compute per-feature intersection metrics
        const tableRows = [];
        for (const f of feats) {
            const attrs = f.attributes || {};
            const geom  = f.geometry;

            let acresCovered = 0;
            let pctAoi       = 0;
            let lengthFeet   = 0;
            let lengthMiles  = 0;

            if (geom) {
                try {
                    const inter = geometryEngine.intersect(aoiGeom, geom);
                    if (inter) {
                        if (isPolygonLayer) {
                            const areaSqm = Math.max(0, geometryEngine.geodesicArea(inter, "square-meters"));
                            acresCovered  = areaSqm / SQM_PER_ACRE;
                            pctAoi        = Math.min(100, Math.max(0, (areaSqm / aoiAreaSqm) * 100));
                        }
                        if (isPolylineLayer) {
                            const lengthM = Math.max(0, geometryEngine.geodesicLength(inter, "meters"));
                            lengthFeet    = lengthM / METERS_PER_FOOT;
                            lengthMiles   = lengthFeet / FEET_PER_MILE;
                        }
                    }
                } catch (e) { /* skip bad geometry */ }
            }

            tableRows.push({ attrs, acresCovered, pctAoi, lengthFeet, lengthMiles });
        }

        // Sort: polygons by acres desc, lines by length desc, points by first attr
        if (isPolygonLayer) {
            tableRows.sort((a, b) => b.acresCovered - a.acresCovered);
        } else if (isPolylineLayer) {
            tableRows.sort((a, b) => b.lengthFeet - a.lengthFeet);
        }

        // Build column labels depending on geometry type
        const extraCols = [];
        if (isPolygonLayer) {
            extraCols.push('Acres', '% of AOI');
        } else if (isPolylineLayer) {
            extraCols.push('Length (ft)', 'Length (mi)');
        }
        // Points get NO extra columns

        // Build interactive table HTML
        const tId = tableId != null ? tableId : Math.floor(Math.random() * 100000);
        const colLabels = [...extraCols, ...attrKeys];
        const totalCols = colLabels.length;
        const extraOffset = extraCols.length;
        const tableTitle = feats.length === 1 ? 'Feature Attributes' : 'Per-Feature Breakdown';

        const thCells = colLabels.map((label, ci) =>
            `<th data-col="${ci}" data-label="${escapeHtml(label)}" data-sort-dir="none" onclick="sortInteractiveTable(this)" title="Click to sort">${escapeHtml(label)} <span class="sort-arrow">\u21C5</span> <button class="col-hide-btn" onclick="event.stopPropagation(); hideColumn('tbl-${tId}',${ci});" title="Hide this column">\u2716</button></th>`
        ).join("");
        const headerHtml = `<tr>${thCells}</tr>`;

        let hasSliverWarning = false;
        const bodyHtml = tableRows.map(row => {
            // Build extra metric cells first (leftmost columns)
            const extraCells = [];
            let isSliverWarning = false;
            if (isPolygonLayer) {
                isSliverWarning = row.pctAoi < 3;
                if (isSliverWarning) hasSliverWarning = true;
                extraCells.push(`<td data-col="0" data-sort-val="${row.acresCovered.toFixed(6)}" style="text-align:right;">${formatNumber(row.acresCovered, 2)}</td>`);
                const warningIcon = isSliverWarning ? '<span style="color:#856404;">\u26A0\uFE0F</span> ' : '';
                extraCells.push(`<td data-col="1" data-sort-val="${row.pctAoi.toFixed(6)}" style="text-align:right;">${warningIcon}${formatNumber(row.pctAoi, 2)}%</td>`);
            } else if (isPolylineLayer) {
                extraCells.push(`<td data-col="0" data-sort-val="${row.lengthFeet.toFixed(2)}" style="text-align:right;">${formatNumber(row.lengthFeet, 1)}</td>`);
                extraCells.push(`<td data-col="1" data-sort-val="${row.lengthMiles.toFixed(6)}" style="text-align:right;">${formatNumber(row.lengthMiles, 3)}</td>`);
            }

            // Build attribute cells (shifted by extraOffset)
            const attrCells = attrKeys.map((k, ci) => {
                const raw = (row.attrs[k] == null) ? "" : String(row.attrs[k]);
                const display = raw.length > 100 ? raw.slice(0, 99) + "\u2026" : raw;
                return `<td data-col="${ci + extraOffset}" data-sort-val="${escapeHtml(raw)}" title="${escapeHtml(raw)}">${escapeHtml(display)}</td>`;
            });

            const allCells = [...extraCells, ...attrCells].join("");
            if (isPolygonLayer && isSliverWarning) {
                return `<tr style="background-color:#fff3cd;">${allCells}</tr>`;
            }
            return `<tr>${allCells}</tr>`;
        }).join("");

        const sliverNote = hasSliverWarning
            ? `<div style="margin-top:10px; font-size:13px; color:#856404; background:#fff3cd; padding:10px 14px; border-radius:6px; line-height:1.5;">
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

            // Chunked query for large AOIs
            querySingleLayerChunked,

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
