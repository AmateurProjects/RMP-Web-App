/* ────────────────────────────────────────────────────────────────
   query-engine.js  –  Query, paging, coverage & analysis helpers
   AMD module loaded via dojoConfig  →  "app/query-engine"
   ──────────────────────────────────────────────────────────────── */

define([
    "app/config-helpers",
    "esri/layers/FeatureLayer",
    "esri/geometry/geometryEngine",
    "esri/geometry/Extent",
    "esri/geometry/Polygon"
], function (configHelpers, FeatureLayer, geometryEngine, Extent, Polygon) {
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

    // ── Layer instance cache (avoid recreating FeatureLayer for every query) ──
    const _layerCache = new Map();  // url → FeatureLayer instance

    // ── Pre-warm cache for report screenshot layers ──
    // Loaded in background after screening so report-gen .when() resolves instantly.
    const _reportLayerCache = new Map();  // lowercase url → loaded FeatureLayer

    /**
     * Get or create a FeatureLayer instance (avoids metadata re-fetch).
     */
    function getCachedLayer(url) {
        const key = String(url).replace(/\/+$/, "").toLowerCase();
        if (!_layerCache.has(key)) {
            _layerCache.set(key, new FeatureLayer({ url, outFields: ["*"] }));
        }
        return _layerCache.get(key);
    }

    // ── Geometry feature cache (shared between coverage stats and per-feature table) ──
    // Key: `${aoiKey}||${layerUrl}` → Feature[] (with geometry + attributes)
    const _geomFeatureCache = new Map();

    // ── Per-feature intersection results cache ──
    // Key: `${aoiKey}||${layerUrl}` → Map<OID, { acresCovered, pctAoi, lengthFeet, lengthMiles }>
    const _intersectionCache = new Map();

    // ── Clip-envelope helpers ────────────────────────────────────

    /**
     * Build a padded bounding-box Polygon around an AOI geometry.
     * Used as a pre-clip envelope: after features are queried with the
     * precise AOI polygon, oversized features are trimmed to this box
     * client-side before the expensive per-feature AOI intersect runs.
     *
     * Returns a Polygon (not an Extent) because geometryEngine.intersect()
     * can produce degenerate / planet-sized results when given a raw Extent.
     *
     * @param {Polygon} aoiGeom  – the AOI polygon
     * @param {number}  [factor] – expansion factor (default from config, fallback 1.5)
     * @returns {Polygon} padded bounding-box polygon
     */
    function getClipEnvelope(aoiGeom, factor) {
        const pad = factor ?? S.config?.report?.clipPaddingFactor ?? 1.5;
        const ext = aoiGeom.extent.expand(pad);
        // Convert Extent → Polygon so geometryEngine.intersect works reliably.
        // Ring must be CLOCKWISE for ArcGIS to treat it as an outer boundary.
        // (CCW would be interpreted as a hole → "everything except this box".)
        return new Polygon({
            rings: [[
                [ext.xmin, ext.ymin],
                [ext.xmax, ext.ymin],
                [ext.xmax, ext.ymax],
                [ext.xmin, ext.ymax],
                [ext.xmin, ext.ymin]
            ]],
            spatialReference: ext.spatialReference
        });
    }

    /**
     * Selectively clip feature geometries to a bounding-box Extent.
     *
     * Only features whose own extent is **larger** than the clip envelope
     * are actually clipped — these are the ones where the bbox pre-trim
     * saves meaningful work for the subsequent precise AOI intersect.
     * Small features that fit inside the clip box are passed through
     * untouched, avoiding redundant geometry operations.
     *
     * – Polygons / polylines larger than the box are trimmed at the boundary.
     * – Points pass through unchanged (containment is already guaranteed
     *   by the server-side spatial filter).
     * – Features whose geometry becomes null after clipping are removed.
     *
     * @param {Feature[]} features   – features with .geometry
     * @param {Polygon}    clipPoly   – the padded bounding-box polygon
     * @returns {Feature[]} features with (selectively) clipped geometries
     */
    function clipFeaturesToEnvelope(features, clipPoly) {
        if (!features?.length || !clipPoly) return features || [];

        // Pre-compute the clip polygon's extent area for the size comparison.
        const cExt     = clipPoly.extent;
        const clipW    = cExt.xmax - cExt.xmin;
        const clipH    = cExt.ymax - cExt.ymin;
        const clipArea = clipW * clipH;
        // Threshold: only clip features whose extent area exceeds the clip box.
        // A multiplier < 1 means "clip anything that extends meaningfully
        // beyond the AOI"; 0.8 catches features ~80 % of the box or larger.
        const THRESH = clipArea * 0.8;

        const out = [];
        for (const f of features) {
            const g = f?.geometry;
            if (!g) continue;
            try {
                const gType = (g.type || "").toLowerCase();

                // Points: always keep (server already filtered to AOI)
                if (gType === "point") {
                    out.push(f);
                    continue;
                }

                // Check whether this feature is large enough to benefit from clipping
                const fExt = g.extent;
                if (fExt) {
                    const fW = fExt.xmax - fExt.xmin;
                    const fH = fExt.ymax - fExt.ymin;
                    if (fW * fH < THRESH) {
                        // Feature is smaller than the clip box — skip clipping
                        out.push(f);
                        continue;
                    }
                }

                // Large feature: clip to the bounding box (polygon-polygon intersect)
                const clipped = geometryEngine.intersect(g, clipPoly);
                if (clipped) {
                    const fc = f.clone ? f.clone() : Object.assign(Object.create(Object.getPrototypeOf(f)), f);
                    fc.geometry = clipped;
                    out.push(fc);
                } // else feature is entirely outside the clip box — drop it
            } catch (e) {
                // On error keep the original feature untouched
                out.push(f);
            }
        }
        return out;
    }

    // ── Helpers ──────────────────────────────────────────────────

    /**
     * Stable-enough signature for an AOI geometry (extent + rounded area).
     */
    function getAoiKey(geom) {
        try {
            const ex   = geom?.extent;
            const area = geometryEngine.geodesicArea(geom, "square-meters");
            // Include vertex count to avoid collisions between different AOIs with same bbox + area
            let vCount = 0;
            if (geom?.rings) for (const r of geom.rings) vCount += r.length;
            else if (geom?.paths) for (const p of geom.paths) vCount += p.length;
            return [ex?.xmin?.toFixed(6), ex?.ymin?.toFixed(6), ex?.xmax?.toFixed(6), ex?.ymax?.toFixed(6), area.toFixed(2), vCount].join("|");
        } catch (e) {
            return String(Date.now());
        }
    }

    /**
     * Clear the coverage cache when the AOI changes.
     */
    function resetCoverageCacheForAoi(geom) {
        coverageCache.clear();
        _geomFeatureCache.clear();
        _intersectionCache.clear();
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
     * Page through all features WITH geometry AND attributes.
     * Fetching attributes here allows the geometry cache to serve
     * both computeLayerCoverageStats and buildPerFeatureTable without
     * a second round trip.
     *
     * Optional optimisations (via `opts`):
     *  • `maxAllowableOffset` – server-side geometry generalisation, reducing
     *    vertex density in the response.  Pure win for coverage calculations
     *    where sub-metre precision isn't needed.
     *  • `clipEnvelope` – after features are returned, any feature whose
     *    extent is significantly larger than the clip box is trimmed to it
     *    client-side.  This keeps the original AOI polygon as the server
     *    query geometry (most accurate hit-set) while still stripping heavy
     *    off-screen vertices from oversized features.
     *
     * @param {FeatureLayer} layer
     * @param {Query}        baseQuery
     * @param {number}       pageSize
     * @param {number}       maxExportFeatures
     * @param {Object}       [opts]
     * @param {Extent}       [opts.clipEnvelope]  – padded bbox from getClipEnvelope()
     * @param {number}       [opts.maxAllowableOffset] – server-side generalisation tolerance (metres)
     */
    async function queryAllFeaturesPagedWithGeometry(layer, baseQuery, pageSize, maxExportFeatures, opts) {
        const clipEnvelope        = opts?.clipEnvelope ?? null;
        const maxAllowableOffset  = opts?.maxAllowableOffset ?? null;

        const all = [];
        let offset = 0;

        while (true) {
            const q = baseQuery.clone();
            q.num   = pageSize;
            q.start = offset;
            q.returnGeometry    = true;
            q.outFields         = ["*"];
            q.outSpatialReference = S.view?.spatialReference;

            // ── Server-side generalisation: reduce vertex density in the
            //    response when sub-metre precision isn't needed.
            if (maxAllowableOffset != null && maxAllowableOffset > 0) {
                q.maxAllowableOffset = maxAllowableOffset;
            }

            const fs   = await layer.queryFeatures(q);
            const feats = (fs && fs.features) ? fs.features : [];
            all.push(...feats);

            if (feats.length < pageSize) break;
            offset += pageSize;
            if (maxExportFeatures && all.length >= maxExportFeatures) break;
        }

        // ── Selective client-side clip: only trim features whose extent is
        //    larger than the clip box (the ones with heavy off-screen geometry).
        //    Small features pass through untouched — no redundant work.
        if (clipEnvelope && all.length > 0) {
            return clipFeaturesToEnvelope(all, clipEnvelope);
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
        const skipExtentCheck  = !!options.skipExtentCheck;

        const layer = getCachedLayer(layerUrl);
        
        // ── Extent pre-check: skip layers with no geographic overlap ──
        // This avoids unnecessary queries for layers that don't cover the AOI region.
        // Only run when both extents share the same spatial reference to avoid
        // false negatives from coordinate system mismatches.
        if (!skipExtentCheck && objectId == null) {
            try {
                await layer.load();
                if (layer.fullExtent && geom?.extent) {
                    const layerSR = layer.fullExtent.spatialReference;
                    const geomSR  = geom.extent.spatialReference;
                    // Only compare extents when spatial references match
                    const srMatch = layerSR && geomSR && (
                        layerSR.wkid === geomSR.wkid ||
                        (layerSR.isWebMercator && geomSR.isWebMercator) ||
                        (layerSR.wkid === 4326 && geomSR.wkid === 4326)
                    );
                    if (srMatch) {
                        const overlaps = geometryEngine.intersects(layer.fullExtent, geom.extent);
                        if (!overlaps) {
                            // Layer extent doesn't overlap AOI - skip query entirely
                            return { 
                                title: layerTitle, 
                                url: layerUrl, 
                                count: 0, 
                                features: [], 
                                layer, 
                                exportQuery: null,
                                _skippedExtentCheck: true
                            };
                        }
                    }
                }
            } catch (e) {
                // If extent check fails, proceed with query anyway
                console.warn(`[querySingleLayer] Extent check failed for ${layerTitle}:`, e.message);
            }
        }
        
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

        // Fire count + sample queries in parallel (saves one full round trip)
        const maxSamples = S.config.report?.maxSampleFeaturesPerLayer ?? 25;
        const q2 = q.clone();
        q2.num = Math.min(maxSamples, 2000);

        const [countResult, sampleResult] = await Promise.all([
            layer.queryFeatureCount(q).catch(() => -1),
            maxSamples > 0 ? layer.queryFeatures(q2).catch(() => null) : Promise.resolve(null)
        ]);

        const sampleFeats = sampleResult?.features ?? [];
        // Use whichever is more authoritative: sometimes queryFeatureCount
        // returns 0 while queryFeatures returns features (observed with
        // some ArcGIS Server services like BLM NLSDB).
        const count = (countResult > 0)
            ? countResult
            : (sampleFeats.length > 0 ? sampleFeats.length : Math.max(countResult, 0));
        let features = [];

        if (count > 0 && sampleFeats.length > 0) {
            features = applyTouchFilter ? filterTouchingOnly(sampleFeats, geom) : sampleFeats;
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
        const layer = getCachedLayer(layerUrl);

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
                    // Deduplicate sample features by OID using Set for O(1) lookup
                    const existingOids = new Set(
                        allSampleFeatures.map(sf => sf.attributes?.[layer.objectIdField || "OBJECTID"])
                            .filter(v => v != null)
                    );
                    for (const f of features) {
                        if (allSampleFeatures.length >= maxSamples) break;
                        const fOid = f.attributes?.[layer.objectIdField || "OBJECTID"];
                        if (fOid != null && existingOids.has(fOid)) continue;
                        allSampleFeatures.push(f);
                        if (fOid != null) existingOids.add(fOid);
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

    // ── Elevation / 3DEP helpers ───────────────────────────────

    /**
     * Prepare geometry + adaptive pixelSize for 3DEP ImageServer calls.
     * Simplifies high-vertex polygons and scales pixel size with AOI area
     * so the server doesn't choke on huge raster requests.
     */
    function _prep3DEPGeometry(geometry) {
        const raw   = geometry.toJSON ? geometry.toJSON() : geometry;
        const ext   = geometry.extent || raw.extent;
        let geomJson, geometryType;

        // Count polygon vertices
        let verts = 0;
        if (raw.rings) { for (const r of raw.rings) verts += r.length; }

        if (verts <= 500) {
            geomJson     = JSON.stringify(raw);
            geometryType = "esriGeometryPolygon";
        } else if (verts <= 2000) {
            try {
                const simplified = geometryEngine.generalize(geometry, 100, true, "meters");
                const sJson = simplified && simplified.toJSON ? simplified.toJSON() : simplified;
                if (sJson && sJson.rings) {
                    geomJson     = JSON.stringify(sJson);
                    geometryType = "esriGeometryPolygon";
                }
            } catch (_) { /* fall through to envelope */ }
            if (!geomJson && ext) {
                geomJson     = JSON.stringify(ext.toJSON ? ext.toJSON() : ext);
                geometryType = "esriGeometryEnvelope";
            }
        } else if (ext) {
            geomJson     = JSON.stringify(ext.toJSON ? ext.toJSON() : ext);
            geometryType = "esriGeometryEnvelope";
        }
        if (!geomJson) {
            geomJson     = JSON.stringify(raw);
            geometryType = "esriGeometryPolygon";
        }

        // Adaptive pixel size: measure AOI width/height in map units,
        // cap the raster at ~2000 px per side so the server always responds.
        let pixelSize = null;
        if (ext) {
            const maxDim = Math.max(ext.width || 0, ext.height || 0);
            if (maxDim > 20000) {          // > ~20 km → coarsen pixels
                pixelSize = Math.ceil(maxDim / 2000);
            }
        }

        return { geomJson, geometryType, pixelSize };
    }

    /**
     * POST helper with timeout for 3DEP ImageServer endpoints.
     */
    async function _post3DEP(url, params, timeoutMs) {
        const controller = new AbortController();
        const tid = setTimeout(() => controller.abort(), timeoutMs || 60000);
        try {
            const resp = await fetch(url, {
                method:  "POST",
                headers: { "Content-Type": "application/x-www-form-urlencoded" },
                body:    new URLSearchParams(params).toString(),
                signal:  controller.signal
            });
            if (!resp.ok) throw new Error("HTTP " + resp.status);
            const json = await resp.json();
            if (json.error) throw new Error("ArcGIS error " + json.error.code + ": " + json.error.message);
            return json;
        } finally { clearTimeout(tid); }
    }

    /** Build envelope-only params for fallback. */
    function _envelopeParams(geometry, pixelSize) {
        const ext = geometry.extent;
        if (!ext) return null;
        const p = {
            f: "json",
            geometry:     JSON.stringify(ext.toJSON ? ext.toJSON() : ext),
            geometryType: "esriGeometryEnvelope"
        };
        if (pixelSize) p.pixelSize = JSON.stringify({ x: pixelSize, y: pixelSize });
        return p;
    }

    // ── Elevation stats ─────────────────────────────────────────

    /**
     * Compute min / max / mean elevation for an AOI from a 3DEP ImageServer.
     * Uses computeStatisticsHistograms (returns stats directly) as primary,
     * with an envelope fallback for large or complex polygons.
     */
    async function computeElevationStats(imageServerUrl, geometry) {
        if (!imageServerUrl || !geometry) return null;

        const { geomJson, geometryType, pixelSize } = _prep3DEPGeometry(geometry);
        const params = { f: "json", geometry: geomJson, geometryType: geometryType };
        if (pixelSize) params.pixelSize = JSON.stringify({ x: pixelSize, y: pixelSize });

        function _buildResult(minElev, maxElev, mean) {
            if (minElev == null || maxElev == null) return null;
            if (!isFinite(minElev) || !isFinite(maxElev)) return null;
            if (maxElev < minElev) { const t = minElev; minElev = maxElev; maxElev = t; }
            if (minElev < -500 || maxElev > 10000) return null;
            if (mean != null && !isFinite(mean)) mean = null;
            return {
                min: minElev, max: maxElev, mean,
                minFt:             minElev * 3.28084,
                maxFt:             maxElev * 3.28084,
                meanFt:            mean != null ? mean * 3.28084 : null,
                elevationChange:   maxElev - minElev,
                elevationChangeFt: (maxElev - minElev) * 3.28084
            };
        }

        function _parseStats(data) {
            const s = data.statistics && data.statistics[0];
            if (s && s.min != null && s.max != null) {
                return _buildResult(s.min, s.max, s.mean != null ? s.mean : null);
            }
            // Fall back to histogram in same response
            const h = data.histograms && data.histograms[0];
            if (h && h.min != null && h.max != null) {
                let mean = null;
                if (h.counts && h.counts.length) {
                    const bw = (h.max - h.min) / h.counts.length;
                    let sum = 0, n = 0;
                    for (let i = 0; i < h.counts.length; i++) { sum += (h.min + (i + 0.5) * bw) * h.counts[i]; n += h.counts[i]; }
                    mean = n > 0 ? sum / n : null;
                }
                return _buildResult(h.min, h.max, mean);
            }
            return null;
        }

        // Primary: computeStatisticsHistograms with polygon/envelope
        try {
            const data = await _post3DEP(imageServerUrl + "/computeStatisticsHistograms", params, 60000);
            const result = _parseStats(data);
            if (result) return result;
        } catch (e) {
            console.warn("[3DEP elevStats] Primary request failed:", e.message);
        }

        // Envelope fallback (if not already an envelope)
        if (geometryType !== "esriGeometryEnvelope") {
            const envP = _envelopeParams(geometry, pixelSize);
            if (envP) {
                try {
                    const data = await _post3DEP(imageServerUrl + "/computeStatisticsHistograms", envP, 60000);
                    const result = _parseStats(data);
                    if (result) return result;
                } catch (e2) {
                    console.warn("[3DEP elevStats] Envelope fallback failed:", e2.message);
                }
            }
        }

        return null;
    }

    // ── Slope aspect (mean slope direction) ─────────────────────

    function degToCardinal(deg) {
        const dirs = ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW'];
        return dirs[Math.round(((deg % 360) + 360) % 360 / 45) % 8];
    }

    /**
     * Compute the mean slope direction (aspect) for an AOI using the
     * Aspect raster function histogram and circular-mean math.
     */
    async function computeSlopeAspect(imageServerUrl, geometry) {
        if (!imageServerUrl || !geometry) return null;

        const { geomJson, geometryType, pixelSize } = _prep3DEPGeometry(geometry);

        function _tryParse(data) {
            if (!data.histograms || !data.histograms.length) return null;
            const hist = data.histograms[0];
            if (!hist.counts || !hist.counts.length) return null;
            const binWidth = (hist.max - hist.min) / hist.counts.length;
            let sumSin = 0, sumCos = 0, total = 0;
            for (let i = 0; i < hist.counts.length; i++) {
                const c = hist.counts[i];
                if (!c) continue;
                const deg = hist.min + (i + 0.5) * binWidth;
                if (deg < 0) continue;  // flat pixels (aspect = -1)
                const rad = deg * Math.PI / 180;
                sumSin += c * Math.sin(rad);
                sumCos += c * Math.cos(rad);
                total  += c;
            }
            if (!total) return null;
            let mean = Math.atan2(sumSin, sumCos) * 180 / Math.PI;
            if (mean < 0) mean += 360;
            return {
                meanAspectDeg:     Math.round(mean * 10) / 10,
                concentration:     Math.round(Math.sqrt(sumSin * sumSin + sumCos * sumCos) / total * 1000) / 1000,
                cardinalDirection: degToCardinal(mean)
            };
        }

        function _buildParams(gJson, gType) {
            const p = {
                f: "json", geometry: gJson, geometryType: gType,
                renderingRule: JSON.stringify({ rasterFunction: "Aspect" })
            };
            if (pixelSize) p.pixelSize = JSON.stringify({ x: pixelSize, y: pixelSize });
            return p;
        }

        // Primary attempt
        try {
            const data = await _post3DEP(imageServerUrl + "/computeHistograms", _buildParams(geomJson, geometryType), 60000);
            const result = _tryParse(data);
            if (result) return result;
        } catch (e) {
            console.warn("[3DEP aspect] Primary request failed:", e.message);
        }

        // Envelope fallback
        if (geometryType !== "esriGeometryEnvelope" && geometry.extent) {
            const envJson = JSON.stringify(geometry.extent.toJSON ? geometry.extent.toJSON() : geometry.extent);
            try {
                const data = await _post3DEP(imageServerUrl + "/computeHistograms", _buildParams(envJson, "esriGeometryEnvelope"), 60000);
                const result = _tryParse(data);
                if (result) return result;
            } catch (e2) {
                console.warn("[3DEP aspect] Envelope fallback failed:", e2.message);
            }
        }

        return null;
    }

    // ── Mean slope grade (steepness) ────────────────────────────

    /**
     * Compute the mean slope grade for an AOI using the Slope raster
     * function on a 3DEP ImageServer.  Returns { meanSlopeDeg, meanSlopePct }
     * or null on failure.
     */
    async function computeMeanSlopeGrade(imageServerUrl, geometry) {
        if (!imageServerUrl || !geometry) return null;

        const { geomJson, geometryType, pixelSize } = _prep3DEPGeometry(geometry);

        function _buildParams(gJson, gType) {
            const p = {
                f: "json", geometry: gJson, geometryType: gType,
                renderingRule: JSON.stringify({ rasterFunction: "Slope_Degrees" })
            };
            if (pixelSize) p.pixelSize = JSON.stringify({ x: pixelSize, y: pixelSize });
            return p;
        }

        function _tryParse(data) {
            // computeStatisticsHistograms returns statistics directly
            const s = data.statistics && data.statistics[0];
            if (s && s.mean != null && isFinite(s.mean)) {
                const deg = Math.abs(s.mean);
                return { meanSlopeDeg: Math.round(deg * 10) / 10, meanSlopePct: Math.round(Math.tan(deg * Math.PI / 180) * 1000) / 10 };
            }
            // Fall back to histogram
            const h = data.histograms && data.histograms[0];
            if (h && h.counts && h.counts.length) {
                const bw = (h.max - h.min) / h.counts.length;
                let sum = 0, n = 0;
                for (let i = 0; i < h.counts.length; i++) { sum += (h.min + (i + 0.5) * bw) * h.counts[i]; n += h.counts[i]; }
                if (n > 0) {
                    const deg = Math.abs(sum / n);
                    return { meanSlopeDeg: Math.round(deg * 10) / 10, meanSlopePct: Math.round(Math.tan(deg * Math.PI / 180) * 1000) / 10 };
                }
            }
            return null;
        }

        // Primary: computeStatisticsHistograms with Slope raster function
        try {
            const data = await _post3DEP(imageServerUrl + "/computeStatisticsHistograms", _buildParams(geomJson, geometryType), 60000);
            const result = _tryParse(data);
            if (result) return result;
        } catch (e) {
            console.warn("[3DEP slopeGrade] Primary request failed:", e.message);
        }

        // Try with "Slope" function name (some services use this instead)
        try {
            const altParams = _buildParams(geomJson, geometryType);
            altParams.renderingRule = JSON.stringify({ rasterFunction: "Slope" });
            const data = await _post3DEP(imageServerUrl + "/computeStatisticsHistograms", altParams, 60000);
            const result = _tryParse(data);
            if (result) return result;
        } catch (e) {
            console.warn("[3DEP slopeGrade] Alt function name failed:", e.message);
        }

        // Envelope fallback
        if (geometryType !== "esriGeometryEnvelope" && geometry.extent) {
            const envJson = JSON.stringify(geometry.extent.toJSON ? geometry.extent.toJSON() : geometry.extent);
            try {
                const data = await _post3DEP(imageServerUrl + "/computeStatisticsHistograms", _buildParams(envJson, "esriGeometryEnvelope"), 60000);
                const result = _tryParse(data);
                if (result) return result;
            } catch (e2) {
                console.warn("[3DEP slopeGrade] Envelope fallback failed:", e2.message);
            }
        }

        return null;
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
        if (!aoiAreaSqm) return { acresCovered: 0, pctAoiCovered: 0, totalLengthFeet: 0, totalLengthMiles: 0 };

        const pageSize  = S.config.report?.pageSize ?? 1000;
        const maxExport = S.config.report?.maxExportFeatures ?? 50000;

        // ── Pre-clip: query the padded bounding box and trim geometries
        //    before running the expensive per-feature AOI intersect.
        const clipEnvelope       = getClipEnvelope(aoiGeom);
        const maxAllowableOffset = S.config.report?.coverageMaxAllowableOffset ?? 10;
        const feats = await queryAllFeaturesPagedWithGeometry(
            item._layer, item._exportQuery, pageSize, maxExport,
            { clipEnvelope, maxAllowableOffset }
        );

        // Cache the (pre-clipped) geometry features for reuse by buildPerFeatureTable
        _geomFeatureCache.set(cacheKey, feats);

        // Intersect each feature with AOI and cache per-feature results
        // Yield to the main thread every YIELD_BATCH features to prevent UI freeze
        const YIELD_BATCH = 50;
        const interGeoms = [];
        const perFeatureResults = new Map(); // OID → { acresCovered, pctAoi, lengthFeet, lengthMiles }
        const oidField = item._layer?.objectIdField || "OBJECTID";
        for (let fi = 0; fi < feats.length; fi++) {
            if (fi > 0 && fi % YIELD_BATCH === 0) {
                await new Promise(r => setTimeout(r, 0));
            }
            const f = feats[fi];
            const g = f?.geometry;
            if (!g) continue;
            try {
                const inter = geometryEngine.intersect(aoiGeom, g);
                if (!inter) continue;
                const gType = (g.type || "").toLowerCase();
                const isPolygon = gType.includes("polygon");
                const isPolyline = gType.includes("polyline") || gType.includes("line");
                let areaSqm = 0;
                let lengthM = 0;
                if (isPolygon) {
                    areaSqm = Math.max(0, geometryEngine.geodesicArea(inter, "square-meters"));
                    if (areaSqm <= 0) continue;
                } else if (isPolyline) {
                    lengthM = Math.max(0, geometryEngine.geodesicLength(inter, "meters"));
                }
                interGeoms.push(inter);
                const oid = f.attributes?.[oidField];
                if (oid != null) {
                    perFeatureResults.set(oid, {
                        acresCovered: areaSqm / SQM_PER_ACRE,
                        pctAoi: aoiAreaSqm > 0 ? Math.min(100, Math.max(0, (areaSqm / aoiAreaSqm) * 100)) : 0,
                        lengthFeet: lengthM / METERS_PER_FOOT,
                        lengthMiles: lengthM / METERS_PER_FOOT / FEET_PER_MILE
                    });
                }
            } catch (e) { /* skip bad geom */ }
        }

        // Cache per-feature intersection results for buildPerFeatureTable
        _intersectionCache.set(cacheKey, perFeatureResults);

        // Sum total line lengths from per-feature results
        let totalLengthFeet = 0;
        let totalLengthMiles = 0;
        for (const [, pf] of perFeatureResults) {
            totalLengthFeet += pf.lengthFeet || 0;
            totalLengthMiles += pf.lengthMiles || 0;
        }

        if (!interGeoms.length) return { acresCovered: 0, pctAoiCovered: 0, totalLengthFeet, totalLengthMiles };

        let unionGeom = null;
        try { unionGeom = unionGeomsChunked(interGeoms); }
        catch (e) { unionGeom = null; }

        let coveredSqm = 0;
        try { coveredSqm = unionGeom ? Math.max(0, geometryEngine.geodesicArea(unionGeom, "square-meters")) : 0; }
        catch (e) { coveredSqm = 0; }

        const acresCovered  = coveredSqm / SQM_PER_ACRE;
        const pctAoiCovered = Math.min(100, Math.max(0, (coveredSqm / aoiAreaSqm) * 100));

        const out = { acresCovered, pctAoiCovered, totalLengthFeet, totalLengthMiles, intersectingFeatureCount: perFeatureResults.size };
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
        // Note: aoiAreaSqm may be 0 for line/point AOIs — that's OK, we only
        // need it for polygon-layer "% of AOI" calculations.

        const pageSize  = S.config.report?.pageSize ?? 1000;
        const maxExport = S.config.report?.maxExportFeatures ?? 50000;

        // Check geometry cache first (populated by computeLayerCoverageStats)
        const layerUrlKey = String(item.url || "").replace(/\/+$/, "");
        const aoiKey      = coverageAoiKey || getAoiKey(aoiGeom);
        const geomCacheKey = `${aoiKey}||${layerUrlKey}`;

        let feats = [];

        if (_geomFeatureCache.has(geomCacheKey)) {
            // Reuse cached geometry features — but we need attributes too.
            // Coverage stats queries geometry-only; re-query with attributes if needed.
            const cachedFeats = _geomFeatureCache.get(geomCacheKey);
            const hasAttrs = cachedFeats.length > 0 && cachedFeats[0].attributes && Object.keys(cachedFeats[0].attributes).length > 0;
            if (hasAttrs) {
                feats = cachedFeats;
            }
        }

        if (feats.length === 0) {
            // Query all features WITH geometry + attributes, using the
            // same pre-clip optimisation as computeLayerCoverageStats.
            try {
                const clipEnvelope       = getClipEnvelope(aoiGeom);
                const maxAllowableOffset = S.config.report?.coverageMaxAllowableOffset ?? 10;

                const q = item._exportQuery.clone();
                q.returnGeometry      = true;
                q.outFields           = ["*"];
                q.outSpatialReference = S.view?.spatialReference;

                feats = await queryAllFeaturesPagedWithGeometry(
                    item._layer, q, pageSize, maxExport,
                    { clipEnvelope, maxAllowableOffset }
                );
            } catch (e) {
                console.warn("buildPerFeatureTable: query failed", e);
                return "";
            }
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
        const skipPatterns = /^(objectid|oid|fid|globalid|st_area|st_length|st_perimeter|shape([-_. ].*)?|total[-_ ]?(area|length|acres|acreage))$/i;
        const allKeys = new Set();
        for (const f of feats) {
            if (!f.attributes) continue;
            for (const k of Object.keys(f.attributes)) {
                if (!skipPatterns.test(k)) allKeys.add(k);
            }
        }

        const attrKeys = Array.from(allKeys);

        // Build rows: use cached intersection results when available, else compute
        const oidField = item._layer?.objectIdField || "OBJECTID";
        const cachedIntersections = _intersectionCache.get(geomCacheKey);
        const tableRows = [];
        for (const f of feats) {
            const attrs = f.attributes || {};
            const geom  = f.geometry;
            const oid   = attrs[oidField];

            let acresCovered = 0;
            let pctAoi       = 0;
            let lengthFeet   = 0;
            let lengthMiles  = 0;

            // Try cached intersection results first (populated by computeLayerCoverageStats)
            if (cachedIntersections && oid != null && cachedIntersections.has(oid)) {
                const cached = cachedIntersections.get(oid);
                acresCovered = cached.acresCovered;
                pctAoi       = cached.pctAoi;
                lengthFeet   = cached.lengthFeet;
                lengthMiles  = cached.lengthMiles;
            } else if (geom) {
                try {
                    const inter = geometryEngine.intersect(aoiGeom, geom);
                    if (inter) {
                        if (isPolygonLayer) {
                            const areaSqm = Math.max(0, geometryEngine.geodesicArea(inter, "square-meters"));
                            acresCovered  = areaSqm / SQM_PER_ACRE;
                            pctAoi        = aoiAreaSqm > 0 ? Math.min(100, Math.max(0, (areaSqm / aoiAreaSqm) * 100)) : 0;
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

        // Limit to top 5 rows for the report table
        const totalFeatureCount = tableRows.length;
        const displayRows = tableRows.slice(0, 5);

        // Build column labels depending on geometry type
        const extraCols = [];
        if (isPolygonLayer) {
            extraCols.push('Acres in AOI', '% of AOI');
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

        // Resolve highlightFields from config for print column filtering
        let printKeepFields = null;
        if (S && S.layerCfgByUrl && item.url) {
            const cfgEntry = S.layerCfgByUrl.get(item.url)
                          || S.layerCfgByUrl.get(item.url.replace(/\/\d+$/, ''));
            const cfg = cfgEntry?.cfg || cfgEntry;
            if (cfg && Array.isArray(cfg.highlightFields) && cfg.highlightFields.length > 0) {
                printKeepFields = new Set(cfg.highlightFields.map(f => f.toLowerCase()));
            }
        }

        const thCells = colLabels.map((label, ci) => {
            // Extra cols (Acres in AOI, % of AOI, etc.) always visible in print;
            // attribute cols visible only if they match highlightFields (or no config → show first 3)
            let hideCls = '';
            if (ci >= extraOffset) {
                const attrIdx = ci - extraOffset;
                const fieldKey = attrKeys[attrIdx];
                if (printKeepFields) {
                    if (!printKeepFields.has(fieldKey.toLowerCase())) hideCls = ' print-hide-col';
                } else {
                    // No highlightFields configured: keep first 3 attribute columns
                    if (attrIdx >= 3) hideCls = ' print-hide-col';
                }
            }
            return `<th data-col="${ci}" data-label="${escapeHtml(label)}" data-sort-dir="none" class="${hideCls}" onclick="sortInteractiveTable(this)" title="Click to sort">${escapeHtml(label)} <span class="sort-arrow">\u21C5</span> <button class="col-hide-btn" onclick="event.stopPropagation(); hideColumn('tbl-${tId}',${ci});" title="Hide this column">\u2716</button></th>`;
        }).join("");
        const headerHtml = `<tr>${thCells}</tr>`;

        let hasSliverWarning = false;
        const bodyHtml = displayRows.map(row => {
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
                let hideCls = '';
                if (printKeepFields) {
                    if (!printKeepFields.has(k.toLowerCase())) hideCls = ' print-hide-col';
                } else if (ci >= 3) {
                    hideCls = ' print-hide-col';
                }
                return `<td data-col="${ci + extraOffset}" data-sort-val="${escapeHtml(raw)}" class="${hideCls}" title="${escapeHtml(raw)}">${escapeHtml(display)}</td>`;
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
                    <span style="font-size:12px;color:#5a5a5a;margin-left:8px;">(showing top ${displayRows.length} of ${totalFeatureCount} feature${totalFeatureCount !== 1 ? 's' : ''}, ${totalCols} columns \u2014 click \u2716 on a header to hide)</span>
                </div>
                <p style="font-size:12px;line-height:1.5;margin:0 0 10px;color:#5a5a5a;">For complete tables \u2014 please see the CSV file in the Report Package download.</p>
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

    // ── Report pre-warm helpers ────────────────────────────────

    /**
     * Pre-create and .load() FeatureLayer instances for every layer that
     * passed the screening coverage check.  Called once in the background
     * after screening completes so that when the user clicks "Generate
     * Report", each tempLayer.when() resolves instantly (metadata already
     * fetched).  Layers are stored by URL and reused via getPreWarmedLayer.
     */
    async function preWarmReportLayers(layerConfigs) {
        const toLoad = [];
        for (const cfg of layerConfigs) {
            if (!cfg.url || !cfg.hasCoverage || cfg.__isImageService) continue;
            const urlKey = String(cfg.url).replace(/\/+$/, '').toLowerCase();
            if (_reportLayerCache.has(urlKey)) continue;

            const fl = new FeatureLayer({
                url: cfg.url,
                outFields: ["*"],
                visible: false
            });
            fl.minScale = 0;
            fl.maxScale = 0;
            _reportLayerCache.set(urlKey, fl);
            toLoad.push(fl);
        }
        if (toLoad.length) {
            await Promise.allSettled(toLoad.map(fl => fl.load().catch(() => null)));
        }
        console.log(`[pre-warm] ${toLoad.length} report layer(s) metadata pre-loaded`);
    }

    /**
     * Return a pre-loaded FeatureLayer for the given URL, or null.
     * The caller takes ownership — the instance is removed from the cache
     * so it can be safely added to a map, mutated, and discarded.
     */
    function getPreWarmedLayer(url) {
        const urlKey = String(url).replace(/\/+$/, '').toLowerCase();
        const fl = _reportLayerCache.get(urlKey);
        if (fl) {
            _reportLayerCache.delete(urlKey);
            return fl;
        }
        return null;
    }

    /** Clear the pre-warm cache (called on clearAll / AOI change). */
    function clearPreWarmCache() {
        _reportLayerCache.clear();
    }

    /**
     * Fire computeLayerCoverageStats for every polygon layer that passed
     * screening.  Results land in the existing coverageCache so the report
     * generator gets instant cache hits instead of running the heavy
     * geometry-intersection queries during report-gen time.
     * Call this in the background (fire-and-forget) after screening.
     */
    async function preFireCoverageStats(layerConfigs, aoiGeom) {
        if (!aoiGeom) return;
        const tasks = [];
        for (const cfg of layerConfigs) {
            if (!cfg.hasCoverage || cfg.__isImageService) continue;
            if (!cfg._layer || !cfg._exportQuery) continue;
            const gt = cfg._layer.geometryType;
            if (!gt || !String(gt).toLowerCase().includes('polygon')) continue;
            tasks.push(computeLayerCoverageStats(cfg, aoiGeom).catch(() => null));
        }
        if (tasks.length) {
            console.log(`[pre-warm] Pre-firing coverage stats for ${tasks.length} polygon layer(s)`);
            await Promise.allSettled(tasks);
            console.log(`[pre-warm] Coverage stats complete`);
        }
    }

    // ── Module entry point ──────────────────────────────────────

    function init(state) {
        S = state;

        return {
            // Paging
            queryAllFeaturesPaged,
            queryAllFeaturesPagedWithGeometry,

            // Pre-clip helpers
            getClipEnvelope,
            clipFeaturesToEnvelope,

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

            // Slope direction & grade
            computeSlopeAspect,
            computeMeanSlopeGrade,

            // Coverage
            computeLayerCoverageStats,
            buildPerFeatureTable,

            // Coverage cache
            getAoiKey,
            resetCoverageCacheForAoi,
            SQM_PER_ACRE,

            // Layer cache (shared across screening + report generation)
            getCachedLayer,

            // Report pre-warm (background layer pre-loading)
            preWarmReportLayers,
            getPreWarmedLayer,
            clearPreWarmCache,
            preFireCoverageStats,

            // Utility
            sampleWithoutReplacement,
            makeTable
        };
    }

    return { init };
});
