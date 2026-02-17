/**
 * upload-aoi.js
 * ─────────────
 * Client-side parser for spatial files (Shapefile .zip, GeoJSON, KML, KMZ).
 * Converts uploaded geometries into a single ArcGIS Polygon for use as an AOI.
 *
 * External CDN libraries (loaded on demand, only when the user triggers an upload):
 *   • shpjs  – Shapefile ZIP → GeoJSON   (https://cdn.jsdelivr.net/npm/shpjs)
 *   • JSZip  – KMZ extraction             (https://cdn.jsdelivr.net/npm/jszip)
 */
define([
    "esri/geometry/Polygon",
    "esri/geometry/Polyline",
    "esri/geometry/Point",
    "esri/geometry/geometryEngine",
    "esri/geometry/support/webMercatorUtils"
], function (Polygon, Polyline, Point, geometryEngine, webMercatorUtils) {

    "use strict";

    const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB

    // Accepted extensions (lower-case)
    const ACCEPT_STRING = ".zip,.geojson,.json,.kml,.kmz";
    const FORMAT_LABELS = "Shapefile (.zip), GeoJSON, KML, or KMZ";

    // ── Dynamic CDN script loader ───────────────────────────────────

    function _loadScript(url, globalName) {
        if (window[globalName]) return Promise.resolve(window[globalName]);
        return new Promise(function (resolve, reject) {
            // Temporarily hide AMD `define` so UMD libraries fall back to
            // setting a window global instead of registering as an AMD module.
            var origDefine = window.define;
            window.define = undefined;

            var s = document.createElement("script");
            s.src = url;
            s.onload = function () {
                window.define = origDefine;   // restore AMD loader
                resolve(window[globalName]);
            };
            s.onerror = function () {
                window.define = origDefine;   // restore AMD loader
                reject(new Error("Failed to load library: " + globalName));
            };
            document.head.appendChild(s);
        });
    }

    var _shpPromise = null;
    function _loadShpJs() {
        if (!_shpPromise) _shpPromise = _loadScript(
            "https://cdn.jsdelivr.net/npm/shpjs@4.0.4/dist/shp.min.js", "shp"
        );
        return _shpPromise;
    }

    var _jszipPromise = null;
    function _loadJSZip() {
        if (!_jszipPromise) _jszipPromise = _loadScript(
            "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js", "JSZip"
        );
        return _jszipPromise;
    }

    // ── File-type detection ─────────────────────────────────────────

    function detectFileType(file) {
        var name = (file.name || "").toLowerCase();
        if (name.endsWith(".zip"))      return "shapefile";
        if (name.endsWith(".geojson") || name.endsWith(".json")) return "geojson";
        if (name.endsWith(".kml"))      return "kml";
        if (name.endsWith(".kmz"))      return "kmz";
        return null;
    }

    // ── Shapefile (.zip) → GeoJSON ──────────────────────────────────

    async function _parseShapefile(arrayBuffer) {
        var shp = await _loadShpJs();
        var result = await shp(arrayBuffer);

        // shpjs may return a single FC or an array of FCs (multi-layer zip)
        if (Array.isArray(result)) {
            var features = [];
            for (var i = 0; i < result.length; i++) {
                if (result[i] && result[i].features) features.push.apply(features, result[i].features);
            }
            return { type: "FeatureCollection", features: features };
        }
        return result;
    }

    // ── GeoJSON → normalised FeatureCollection ──────────────────────

    var SUPPORTED_GEOM_TYPES = ["Point", "MultiPoint", "LineString", "MultiLineString", "Polygon", "MultiPolygon"];

    function _parseGeoJSON(text) {
        var obj = JSON.parse(text);
        if (obj.type === "FeatureCollection") return obj;
        if (obj.type === "Feature")
            return { type: "FeatureCollection", features: [obj] };
        if (SUPPORTED_GEOM_TYPES.indexOf(obj.type) !== -1)
            return { type: "FeatureCollection", features: [{ type: "Feature", geometry: obj, properties: {} }] };
        throw new Error("Unrecognised GeoJSON structure (expected FeatureCollection, Feature, or a geometry object).");
    }

    // ── KML → GeoJSON ───────────────────────────────────────────────

    function _parseKMLCoords(text) {
        return text.trim().split(/\s+/).filter(Boolean).map(function (c) {
            var p = c.split(",").map(Number);
            return [p[0], p[1]]; // [lng, lat] — discard altitude
        });
    }

    function _parseKML(kmlText) {
        var parser = new DOMParser();
        var doc = parser.parseFromString(kmlText, "text/xml");
        var ns = doc.documentElement.namespaceURI;

        function byTag(parent, tag) {
            return ns
                ? parent.getElementsByTagNameNS(ns, tag)
                : parent.getElementsByTagName(tag);
        }

        var placemarks = byTag(doc, "Placemark");
        var features = [];

        for (var i = 0; i < placemarks.length; i++) {
            var pm = placemarks[i];

            // Points
            var kmlPoints = byTag(pm, "Point");
            for (var pi = 0; pi < kmlPoints.length; pi++) {
                var pc = byTag(kmlPoints[pi], "coordinates");
                if (pc.length === 0) continue;
                var coords = _parseKMLCoords(pc[0].textContent);
                if (coords.length > 0) {
                    features.push({ type: "Feature", geometry: { type: "Point", coordinates: coords[0] }, properties: {} });
                }
            }

            // LineStrings
            var kmlLines = byTag(pm, "LineString");
            for (var li = 0; li < kmlLines.length; li++) {
                var lc = byTag(kmlLines[li], "coordinates");
                if (lc.length === 0) continue;
                var lineCoords = _parseKMLCoords(lc[0].textContent);
                if (lineCoords.length > 1) {
                    features.push({ type: "Feature", geometry: { type: "LineString", coordinates: lineCoords }, properties: {} });
                }
            }

            // Polygons
            var polygons = byTag(pm, "Polygon");
            for (var j = 0; j < polygons.length; j++) {
                var poly = polygons[j];
                var outerEls = byTag(poly, "outerBoundaryIs");
                if (outerEls.length === 0) continue;
                var coordsEls = byTag(outerEls[0], "coordinates");
                if (coordsEls.length === 0) continue;

                var outerRing = _parseKMLCoords(coordsEls[0].textContent);
                var rings = [outerRing];

                // Inner boundaries (holes)
                var innerEls = byTag(poly, "innerBoundaryIs");
                for (var k = 0; k < innerEls.length; k++) {
                    var ic = byTag(innerEls[k], "coordinates");
                    if (ic.length > 0) rings.push(_parseKMLCoords(ic[0].textContent));
                }

                features.push({
                    type: "Feature",
                    geometry: { type: "Polygon", coordinates: rings },
                    properties: {}
                });
            }
        }

        if (features.length === 0) throw new Error("No features found in the KML file.");
        return { type: "FeatureCollection", features: features };
    }

    // ── KMZ → GeoJSON (extract KML from zip then parse) ────────────

    async function _parseKMZ(arrayBuffer) {
        var JSZip = await _loadJSZip();
        var zip = await JSZip.loadAsync(arrayBuffer);

        var kmlText = null;
        var entries = Object.keys(zip.files);
        for (var i = 0; i < entries.length; i++) {
            if (entries[i].toLowerCase().endsWith(".kml") && !zip.files[entries[i]].dir) {
                kmlText = await zip.files[entries[i]].async("string");
                break;
            }
        }
        if (!kmlText) throw new Error("No .kml file found inside the KMZ archive.");
        return _parseKML(kmlText);
    }

    // ── GeoJSON FeatureCollection → single ArcGIS Polygon ──────────

    function _extractRings(geojsonGeom) {
        if (geojsonGeom.type === "Polygon") return geojsonGeom.coordinates;
        if (geojsonGeom.type === "MultiPolygon") {
            var all = [];
            for (var i = 0; i < geojsonGeom.coordinates.length; i++) {
                all.push.apply(all, geojsonGeom.coordinates[i]);
            }
            return all;
        }
        return null;
    }

    /**
     * Convert a GeoJSON FeatureCollection to ArcGIS geometries.
     * Returns { geometry, featureCount, geometryType, hasPoints, hasLines, hasPolygons }.
     * The geometry is unified (unioned) per geometry family.
     * If the file contains only points or lines (no polygons), the raw
     * unified geometry is returned and the caller is responsible for
     * buffering it into a polygon AOI.
     */
    function _toAoiGeometry(geojsonFC, viewSR) {
        var points   = [];
        var lines    = [];
        var polygons = [];
        var SR4326   = { wkid: 4326 };

        for (var i = 0; i < (geojsonFC.features || []).length; i++) {
            var feat = geojsonFC.features[i];
            if (!feat.geometry) continue;
            var t = feat.geometry.type;

            if (t === "Point") {
                points.push(new Point({ longitude: feat.geometry.coordinates[0], latitude: feat.geometry.coordinates[1], spatialReference: SR4326 }));
            } else if (t === "MultiPoint") {
                for (var mp = 0; mp < feat.geometry.coordinates.length; mp++) {
                    points.push(new Point({ longitude: feat.geometry.coordinates[mp][0], latitude: feat.geometry.coordinates[mp][1], spatialReference: SR4326 }));
                }
            } else if (t === "LineString") {
                lines.push(new Polyline({ paths: [feat.geometry.coordinates], spatialReference: SR4326 }));
            } else if (t === "MultiLineString") {
                lines.push(new Polyline({ paths: feat.geometry.coordinates, spatialReference: SR4326 }));
            } else if (t === "Polygon" || t === "MultiPolygon") {
                var rings = _extractRings(feat.geometry);
                if (rings && rings.length > 0) {
                    polygons.push(new Polygon({ rings: rings, spatialReference: SR4326 }));
                }
            }
        }

        var totalCount = points.length + lines.length + polygons.length;
        if (totalCount === 0) {
            throw new Error("No supported geometries found in the uploaded file.");
        }

        // Determine dominant type label for UI
        var geometryType = "polygon";
        if (polygons.length > 0)      geometryType = "polygon";
        else if (lines.length > 0)    geometryType = "polyline";
        else                          geometryType = "point";

        // Union each family
        var unified = null;
        var allGeoms = [];
        if (polygons.length > 0) {
            var uPoly = polygons.length === 1 ? polygons[0] : geometryEngine.union(polygons);
            allGeoms.push(uPoly);
            unified = uPoly;
        }
        if (lines.length > 0) {
            var uLine = lines.length === 1 ? lines[0] : geometryEngine.union(lines);
            allGeoms.push(uLine);
            if (!unified) unified = uLine;
        }
        if (points.length > 0) {
            // Union points into a multipoint-like set — but geometryEngine.union
            // on points returns a single point or multipoint. We'll just keep the first
            // or union them. For buffer purposes any single point works if only 1.
            var uPoint = points.length === 1 ? points[0] : geometryEngine.union(points);
            allGeoms.push(uPoint);
            if (!unified) unified = uPoint;
        }

        // Project to view SR
        if (viewSR && viewSR.isWebMercator) {
            unified = webMercatorUtils.geographicToWebMercator(unified);
            for (var ai = 0; ai < allGeoms.length; ai++) {
                allGeoms[ai] = webMercatorUtils.geographicToWebMercator(allGeoms[ai]);
            }
        }

        return {
            geometry:     unified,
            allGeometries: allGeoms,
            featureCount: totalCount,
            geometryType: geometryType,
            hasPoints:    points.length > 0,
            hasLines:     lines.length > 0,
            hasPolygons:  polygons.length > 0
        };
    }

    /**
     * Apply a geodesic buffer (in miles) to one or more geometries
     * and union the result into a single Polygon.
     */
    function applyBuffer(geometries, bufferMiles) {
        if (!geometries || geometries.length === 0) return null;
        if (!bufferMiles || bufferMiles <= 0) {
            // No buffer — union polygons only
            var polys = geometries.filter(function (g) { return g.type === "polygon"; });
            if (polys.length === 0) return null;
            return polys.length === 1 ? polys[0] : geometryEngine.union(polys);
        }
        var buffered = [];
        for (var i = 0; i < geometries.length; i++) {
            var b = geometryEngine.geodesicBuffer(geometries[i], bufferMiles, "miles");
            if (b) buffered.push(b);
        }
        if (buffered.length === 0) return null;
        return buffered.length === 1 ? buffered[0] : geometryEngine.union(buffered);
    }

    // ── Main entry point ────────────────────────────────────────────

    /**
     * processFile(file, viewSpatialReference)
     *
     * @param  {File}   file     – the File object from <input> or drag-drop
     * @param  {Object} viewSR   – the MapView's spatialReference
     * @returns {Promise<{geometry, featureCount, fileName, fileType}>}
     */
    async function processFile(file, viewSR) {
        if (!file) throw new Error("No file provided.");

        if (file.size > MAX_FILE_SIZE) {
            var mbLimit = (MAX_FILE_SIZE / 1024 / 1024).toFixed(0);
            throw new Error("File is too large (" + (file.size / 1024 / 1024).toFixed(1)
                + " MB). Maximum allowed size is " + mbLimit + " MB.");
        }

        var fileType = detectFileType(file);
        if (!fileType) {
            throw new Error("Unsupported file type. Please upload a " + FORMAT_LABELS + " file.");
        }

        var geojsonFC;

        if (fileType === "shapefile") {
            var buf = await file.arrayBuffer();
            geojsonFC = await _parseShapefile(buf);
        } else if (fileType === "geojson") {
            var text = await file.text();
            geojsonFC = _parseGeoJSON(text);
        } else if (fileType === "kml") {
            var kmlText = await file.text();
            geojsonFC = _parseKML(kmlText);
        } else if (fileType === "kmz") {
            var kmzBuf = await file.arrayBuffer();
            geojsonFC = await _parseKMZ(kmzBuf);
        }

        if (!geojsonFC || !geojsonFC.features || geojsonFC.features.length === 0) {
            throw new Error("No features found in the uploaded file.");
        }

        var result = _toAoiGeometry(geojsonFC, viewSR);

        return {
            geometry:     result.geometry,
            featureCount: result.featureCount,
            fileName:     file.name,
            fileType:     fileType
        };
    }

    // ── Public API ──────────────────────────────────────────────────

    return {
        processFile:   processFile,
        applyBuffer:   applyBuffer,
        MAX_FILE_SIZE: MAX_FILE_SIZE,
        ACCEPT_STRING: ACCEPT_STRING,
        FORMAT_LABELS: FORMAT_LABELS
    };
});
