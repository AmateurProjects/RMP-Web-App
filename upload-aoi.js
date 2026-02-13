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
    "esri/geometry/geometryEngine",
    "esri/geometry/support/webMercatorUtils"
], function (Polygon, geometryEngine, webMercatorUtils) {

    "use strict";

    const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB

    // Accepted extensions (lower-case)
    const ACCEPT_STRING = ".zip,.geojson,.json,.kml,.kmz";
    const FORMAT_LABELS = "Shapefile (.zip), GeoJSON, KML, or KMZ";

    // ── Dynamic CDN script loader ───────────────────────────────────

    function _loadScript(url, globalName) {
        if (window[globalName]) return Promise.resolve(window[globalName]);
        return new Promise(function (resolve, reject) {
            var s = document.createElement("script");
            s.src = url;
            s.onload = function () { resolve(window[globalName]); };
            s.onerror = function () { reject(new Error("Failed to load library: " + globalName)); };
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

    function _parseGeoJSON(text) {
        var obj = JSON.parse(text);
        if (obj.type === "FeatureCollection") return obj;
        if (obj.type === "Feature")
            return { type: "FeatureCollection", features: [obj] };
        if (obj.type === "Polygon" || obj.type === "MultiPolygon")
            return { type: "FeatureCollection", features: [{ type: "Feature", geometry: obj, properties: {} }] };
        throw new Error("Unrecognised GeoJSON structure (expected FeatureCollection, Feature, Polygon, or MultiPolygon).");
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

        if (features.length === 0) throw new Error("No polygon features found in the KML file.");
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

    function _toAoiGeometry(geojsonFC, viewSR) {
        var polygons = [];

        for (var i = 0; i < (geojsonFC.features || []).length; i++) {
            var feat = geojsonFC.features[i];
            if (!feat.geometry) continue;
            var t = feat.geometry.type;
            if (t !== "Polygon" && t !== "MultiPolygon") continue;

            var rings = _extractRings(feat.geometry);
            if (!rings || rings.length === 0) continue;

            polygons.push(new Polygon({
                rings: rings,
                spatialReference: { wkid: 4326 }
            }));
        }

        if (polygons.length === 0) {
            throw new Error("No polygon geometries found in the uploaded file. "
                + "The file may contain only points or lines.");
        }

        // Union/dissolve all polygons into one AOI
        var unified = polygons.length === 1
            ? polygons[0]
            : geometryEngine.union(polygons);

        // Project to view SR (typically Web Mercator 3857)
        if (viewSR && viewSR.isWebMercator) {
            unified = webMercatorUtils.geographicToWebMercator(unified);
        }

        return { geometry: unified, featureCount: polygons.length };
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
        MAX_FILE_SIZE: MAX_FILE_SIZE,
        ACCEPT_STRING: ACCEPT_STRING,
        FORMAT_LABELS: FORMAT_LABELS
    };
});
