/**
 * search.js
 * AMD module - Feature Search Widget
 *
 * Self-contained search bar that queries all configured layers for matching
 * features by name, highlights them on the map, and zooms to them.
 *
 * Exports:
 *   init(state, deps)  - wire up DOM and event handlers
 *
 * state keys used: config, selectionLayers, view, map, reportLayerViews
 * deps: { Graphic, GraphicsLayer, escapeHtml, fetchJson, normalizeUrlKey, enableSelectionLayer }
 */
define(["app/config-helpers"], function (configHelpers) {
    "use strict";

    var escapeHtml  = configHelpers.escapeHtml;
    var fetchJson   = configHelpers.fetchJson;
    var normalizeUrlKey = configHelpers.normalizeUrlKey;
    var _getFeatureDisplayName = configHelpers.getFeatureDisplayName;

    // ── Module-private refs (set by init) ──
    var state = null;
    var GraphicsLayerCtor, GraphicCtor;
    var enableSelectionLayer; // callback from app.js

    // ── DOM refs ──
    var searchInput, searchResults, searchClear, searchIcon, searchSpinner;

    // ── Internal state ──
    var searchDebounceTimer   = null;
    var searchAbortController = null;
    var searchGeneration      = 0;
    var fieldMetadataCache    = new Map();
    var allSearchResults      = [];
    var onFeatureSelected     = null; // callback(feature) when user clicks a result

    // ── Field classification patterns ──

    var NAME_FIELD_PATTERNS = [
        /^name$/i, /name$/i, /_name$/i, /name_/i,
        /^title$/i, /^label$/i, /^description$/i, /^desc$/i,
        /comname/i, /sciname/i, /common.*name/i, /scientific.*name/i,
        /plan.*name/i, /proj.*name/i, /unit.*name/i, /area.*name/i,
        /site.*name/i, /allot.*name/i, /lup.*name/i, /permit/i
    ];

    var EXCLUDED_FIELD_PATTERNS = [
        /objectid/i, /globalid/i, /^oid$/i, /^fid$/i, /^id$/i,
        /shape/i, /geometry/i, /^guid$/i, /uuid/i,
        /(?:^|_)id$/i, /code$/i, /_code$/i, /^code/i,
        /serial/i, /row_?num/i, /unique/i, /key$/i,
        /created/i, /modified/i, /edit.*date/i, /update/i,
        /^gis_/i, /^sys_/i, /^db_/i, /^meta/i
    ];

    // ── Helpers ──

    function categorizeSearchFields(fields) {
        var nameFields = [];
        var otherFields = [];

        for (var i = 0; i < fields.length; i++) {
            var fname = fields[i].name || "";
            if (EXCLUDED_FIELD_PATTERNS.some(function (p) { return p.test(fname); })) continue;
            if (NAME_FIELD_PATTERNS.some(function (p) { return p.test(fname); })) {
                nameFields.push(fname);
            } else {
                otherFields.push(fname);
            }
        }
        return { nameFields: nameFields, otherFields: otherFields };
    }

    function getStringFieldsForLayer(url) {
        var cacheKey = url.replace(/\/$/, "");
        if (fieldMetadataCache.has(cacheKey)) {
            return Promise.resolve(fieldMetadataCache.get(cacheKey));
        }
        return fetchJson(cacheKey + "?f=pjson").then(function (info) {
            var fields = (info && info.fields) ? info.fields : [];
            var stringFields = fields.filter(function (f) { return f.type === "esriFieldTypeString"; });
            var categorized = categorizeSearchFields(stringFields);
            fieldMetadataCache.set(cacheKey, categorized);
            return categorized;
        }).catch(function (e) {
            console.warn("Failed to get fields for", url, e);
            return { nameFields: [], otherFields: [] };
        });
    }

    function getSearchableLayers() {
        var layers = [];
        var cfg = state.config;
        (cfg.reportLayers || []).forEach(function (c) {
            if (c && c.url) {
                layers.push({ title: c.title || "Unknown Layer", url: c.url, type: "report" });
            }
        });
        (state.selectionLayers || []).forEach(function (entry) {
            if (entry && entry.cfg && entry.cfg.url) {
                var exists = layers.some(function (l) { return l.url === entry.cfg.url; });
                if (!exists) {
                    layers.push({ title: entry.cfg.title || "Unknown Layer", url: entry.cfg.url, type: "selection" });
                }
            }
        });
        return layers;
    }

    function hasNameFieldMatch(attributes, searchTerm, nameFields) {
        var termLower = searchTerm.toLowerCase();
        for (var i = 0; i < nameFields.length; i++) {
            var val = attributes[nameFields[i]];
            if (val && String(val).toLowerCase().includes(termLower)) return true;
        }
        return false;
    }

    var _searchDisplayFields = [
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

    function getFeatureDisplayName(attributes) {
        return _getFeatureDisplayName(attributes, _searchDisplayFields, 80);
    }

    function calculateRelevance(attributes, searchTerm, nameFields) {
        var termLower = searchTerm.toLowerCase();
        var score = 0;
        for (var i = 0; i < nameFields.length; i++) {
            var val = String(attributes[nameFields[i]] || "").toLowerCase();
            if (val) {
                if (val === termLower) score += 100;
                else if (val.startsWith(termLower)) score += 50;
                else if (val.includes(termLower)) score += 25;
            }
        }
        var displayName = getFeatureDisplayName(attributes).toLowerCase();
        if (displayName.includes(termLower)) score += 30;
        return score;
    }

    function getFeatureDetails(attributes) {
        var details = [];
        var skipFields = ["OBJECTID", "GLOBALID", "SHAPE", "SHAPE_LENGTH", "SHAPE_AREA", "SHAPE.LEN", "SHAPE.AREA"];
        var count = 0;
        var entries = Object.entries(attributes);
        for (var i = 0; i < entries.length; i++) {
            if (count >= 2) break;
            var key = entries[i][0], val = entries[i][1];
            if (skipFields.some(function (s) { return key.toUpperCase().includes(s); })) continue;
            if (val && String(val).trim()) {
                details.push(key + ": " + String(val).trim().substring(0, 40));
                count++;
            }
        }
        return details.join(" | ");
    }

    // ── Layer search ──

    function searchLayer(layerInfo, searchTerm, signal, maxResults) {
        if (maxResults === undefined) maxResults = 5;

        return getStringFieldsForLayer(layerInfo.url).then(function (cats) {
            var nameFields  = cats.nameFields;
            var otherFields = cats.otherFields;
            if (!nameFields.length && !otherFields.length) return [];

            var escapedTerm = searchTerm.replace(/'/g, "''").replace(/%/g, "[%]").replace(/_/g, "[_]");
            var allSearchFields = nameFields.concat(otherFields);
            var whereClauses = allSearchFields.map(function (f) {
                return "UPPER(" + f + ") LIKE '%" + escapedTerm.toUpperCase() + "%'";
            });
            var where = whereClauses.join(" OR ");

            var view = state.view;
            var queryUrl = layerInfo.url.replace(/\/$/, "") + "/query";
            var params = new URLSearchParams({
                where: where,
                outFields: "*",
                returnGeometry: "true",
                outSR: String((view && view.spatialReference && view.spatialReference.wkid) || 102100),
                resultRecordCount: String(maxResults * 2),
                f: "json"
            });

            return fetch(queryUrl + "?" + params.toString(), { signal: signal, credentials: "omit" })
                .then(function (response) {
                    if (!response.ok) {
                        console.warn("Search HTTP error for", layerInfo.title, ":", response.status, response.statusText);
                        return [];
                    }
                    return response.json();
                })
                .then(function (data) {
                    var features = (data && data.features) ? data.features : [];
                    var results = features.map(function (f) {
                        var attrs = f.attributes || {};
                        return {
                            layerTitle: layerInfo.title,
                            layerUrl: layerInfo.url,
                            attributes: attrs,
                            geometry: f.geometry,
                            relevance: calculateRelevance(attrs, searchTerm, nameFields),
                            hasNameMatch: hasNameFieldMatch(attrs, searchTerm, nameFields)
                        };
                    });

                    var filtered = results.filter(function (r) {
                        if (r.hasNameMatch) return true;
                        var dn = getFeatureDisplayName(r.attributes).toLowerCase();
                        if (dn.includes(searchTerm.toLowerCase())) return true;
                        return false;
                    });

                    return filtered
                        .sort(function (a, b) { return b.relevance - a.relevance; })
                        .slice(0, maxResults);
                });
        }).catch(function (e) {
            if (e.name !== "AbortError") {
                console.warn("Search failed for layer:", layerInfo.title, e);
            }
            return [];
        });
    }

    // ── Enable a layer by URL (for search result clicks) ──

    function enableLayerByUrl(layerUrl) {
        if (!layerUrl) return Promise.resolve();

        var normalizedUrl = String(layerUrl).replace(/\/+$/, "").toLowerCase();
        var normalizedKey = normalizeUrlKey(layerUrl);
        var selectionLayers = state.selectionLayers || [];

        // Check selection layers first
        for (var i = 0; i < selectionLayers.length; i++) {
            var entry = selectionLayers[i];
            var entryUrl = String((entry && entry.cfg && entry.cfg.url) || "").replace(/\/+$/, "").toLowerCase();
            if (entryUrl === normalizedUrl) {
                if (!entry.layer.visible) {
                    return enableSelectionLayer(i);
                }
                return Promise.resolve();
            }
        }

        // Check report layers
        var reportLayerViews = state.reportLayerViews;
        if (reportLayerViews) {
            for (var _a = reportLayerViews.entries(), _entry; !(_entry = _a.next()).done; ) {
                var pair = _entry.value;
                var key = pair[0];
                var layerOrArray = pair[1];
                var layers = Array.isArray(layerOrArray) ? layerOrArray : [layerOrArray];

                var keyMatch = (key === normalizedKey);
                var urlMatch = layers.some(function (lyr) {
                    var lyrUrl = String((lyr && lyr.url) || "").replace(/\/+$/, "").toLowerCase();
                    return lyrUrl === normalizedUrl;
                });

                if (keyMatch || urlMatch) {
                    layers.forEach(function (lyr) { lyr.visible = true; });
                    var cfgIdx = (state.config.reportLayers || []).findIndex(function (cfg) {
                        return normalizeUrlKey(cfg.url) === key;
                    });
                    if (cfgIdx >= 0) {
                        var checkbox = document.getElementById("rptlayer_" + cfgIdx);
                        if (checkbox && !checkbox.checked) checkbox.checked = true;
                    }
                    return Promise.resolve();
                }
            }
        }

        console.warn("Could not find layer to enable:", layerUrl);
        return Promise.resolve();
    }

    // ── Zoom to a search result feature ──

    function zoomToFeature(feature) {
        var view = state.view;
        var map  = state.map;
        if (!view || !feature.geometry) {
            console.warn("Cannot zoom: no geometry", feature);
            return Promise.resolve();
        }

        var geomJson = feature.geometry;
        var sr = geomJson.spatialReference || view.spatialReference || { wkid: 102100 };
        var graphic = null;
        var geomType = null;

        if (geomJson.rings && geomJson.rings.length > 0) {
            geomType = "polygon";
            graphic = new GraphicCtor({
                geometry: { type: "polygon", rings: geomJson.rings, spatialReference: sr },
                symbol: { type: "simple-fill", color: [255, 255, 0, 0.4], outline: { color: [255, 100, 0], width: 3 } }
            });
        } else if (geomJson.paths && geomJson.paths.length > 0) {
            geomType = "polyline";
            graphic = new GraphicCtor({
                geometry: { type: "polyline", paths: geomJson.paths, spatialReference: sr },
                symbol: { type: "simple-line", color: [255, 255, 0], width: 6 }
            });
        } else if (geomJson.x !== undefined && geomJson.y !== undefined) {
            geomType = "point";
            graphic = new GraphicCtor({
                geometry: { type: "point", x: geomJson.x, y: geomJson.y, spatialReference: sr },
                symbol: { type: "simple-marker", color: [255, 255, 0, 0.8], size: 16, outline: { color: [255, 100, 0], width: 3 } }
            });
        }

        if (!graphic) {
            console.warn("Could not create graphic from geometry", geomJson);
            return Promise.resolve();
        }

        var goToOptions = { animate: true, duration: 800 };
        var goToPromise;

        if (geomType === "point") {
            goToPromise = view.goTo({ target: graphic.geometry, zoom: 14 }, goToOptions);
        } else {
            goToPromise = view.goTo(graphic.geometry, goToOptions);
        }

        var highlightLayer = new GraphicsLayerCtor({ title: "Search Highlight" });
        map.add(highlightLayer);
        highlightLayer.add(graphic);

        // Always clean up the highlight layer after 4s regardless of zoom outcome
        setTimeout(function () {
            try { map.remove(highlightLayer); } catch (e) {}
        }, 4000);

        return goToPromise.catch(function (e) {
            console.error("Zoom to feature failed:", e, feature);
        });
    }

    // ── Core search ──

    function performSearch(searchTerm) {
        if (!searchTerm || searchTerm.length < 2) {
            searchResults.innerHTML = '<div class="search-hint">Type at least 2 characters to search</div>';
            searchResults.classList.add("visible");
            if (searchIcon) searchIcon.style.display = "block";
            if (searchSpinner) searchSpinner.style.display = "none";
            return Promise.resolve();
        }

        if (searchAbortController) {
            try { searchAbortController.abort(); } catch (e) {}
            searchAbortController = null;
        }
        searchAbortController = new AbortController();
        var signal = searchAbortController.signal;
        var myGen = ++searchGeneration;

        if (searchIcon) searchIcon.style.display = "none";
        if (searchSpinner) searchSpinner.style.display = "block";

        var layers = getSearchableLayers();

        var searchPromises = layers.slice(0, 15).map(function (layerInfo) {
            return searchLayer(layerInfo, searchTerm, signal, 5).catch(function (e) {
                if (e.name !== "AbortError") {
                    console.warn("Search failed for layer:", layerInfo.title, e);
                }
                return [];
            });
        });

        return Promise.all(searchPromises).then(function (results) {
            if (signal.aborted || myGen !== searchGeneration) return;

            allSearchResults = [];
            results.forEach(function (layerResults) {
                layerResults.forEach(function (f) { allSearchResults.push(f); });
            });

            allSearchResults.sort(function (a, b) { return (b.relevance || 0) - (a.relevance || 0); });
            allSearchResults = allSearchResults.slice(0, 25);

            var groupedResults = new Map();
            allSearchResults.forEach(function (f) {
                if (!groupedResults.has(f.layerTitle)) groupedResults.set(f.layerTitle, []);
                groupedResults.get(f.layerTitle).push(f);
            });

            if (groupedResults.size === 0) {
                searchResults.innerHTML = '<div class="search-no-results">No matching features found</div>';
            } else {
                var html = "";
                var globalIdx = 0;
                for (var _a = groupedResults.entries(), _entry; !(_entry = _a.next()).done; ) {
                    var pair = _entry.value;
                    var layerTitle = pair[0];
                    var features = pair[1];
                    html += '<div class="search-result-group">';
                    html += '<div class="search-result-group-title">' + escapeHtml(layerTitle) + '</div>';
                    features.forEach(function (feature) {
                        var name = getFeatureDisplayName(feature.attributes);
                        var details = getFeatureDetails(feature.attributes);
                        html += '<div class="search-result-item" role="option" tabindex="-1" data-result-idx="' + globalIdx + '">';
                        html += '<div class="search-result-name">' + escapeHtml(name) + '</div>';
                        if (details) {
                            html += '<div class="search-result-details">' + escapeHtml(details) + '</div>';
                        }
                        html += '</div>';
                        globalIdx++;
                    });
                    html += '</div>';
                }

                searchResults.innerHTML = html;
                searchResults.setAttribute('role', 'listbox');

                var resultItems = searchResults.querySelectorAll(".search-result-item");
                resultItems.forEach(function (item) {
                    item.addEventListener("click", function () {
                        var idx = parseInt(item.getAttribute("data-result-idx"), 10);
                        var feature = allSearchResults[idx];
                        if (feature) {
                            if (onFeatureSelected) {
                                // AOI mode: pass the feature to the callback for direct selection
                                onFeatureSelected(feature);
                            } else {
                                enableLayerByUrl(feature.layerUrl).then(function () {
                                    return zoomToFeature(feature);
                                });
                            }
                        }
                        searchResults.classList.remove("visible");
                    });
                });
            }

            searchResults.classList.add("visible");
        }).catch(function (e) {
            if (myGen !== searchGeneration) return;
            if (e.name === "AbortError") return;
            console.error("Search error:", e);
            searchResults.innerHTML = '<div class="search-no-results">Search failed</div>';
            searchResults.classList.add("visible");
        }).then(function () {
            if (myGen === searchGeneration) {
                if (searchIcon) searchIcon.style.display = "block";
                if (searchSpinner) searchSpinner.style.display = "none";
            }
        });
    }

    // ── Init ──

    function init(appState, deps) {
        state = appState;
        GraphicsLayerCtor    = deps.GraphicsLayer;
        GraphicCtor          = deps.Graphic;
        enableSelectionLayer = deps.enableSelectionLayer;

        // DOM refs
        searchInput   = document.getElementById("featureSearchInput");
        searchResults = document.getElementById("featureSearchResults");
        searchClear   = document.getElementById("featureSearchClear");
        searchIcon    = document.getElementById("featureSearchIcon");
        searchSpinner = document.getElementById("featureSearchSpinner");

        if (!searchInput) return; // widget not present in DOM

        // Input handler with debounce
        searchInput.addEventListener("input", function (e) {
            var val = e.target.value.trim();
            if (searchClear) searchClear.style.display = val ? "flex" : "none";
            clearTimeout(searchDebounceTimer);
            searchDebounceTimer = setTimeout(function () {
                performSearch(val);
            }, 400);
        });

        // Focus handler
        searchInput.addEventListener("focus", function () {
            var val = searchInput.value.trim();
            if (val.length >= 2) {
                searchResults.classList.add("visible");
            } else if (val.length > 0) {
                searchResults.innerHTML = '<div class="search-hint">Type at least 2 characters to search</div>';
                searchResults.classList.add("visible");
            }
        });

        // Close results on outside click
        document.addEventListener("click", function (e) {
            var widget = document.getElementById("featureSearchWidget");
            if (widget && !widget.contains(e.target)) {
                searchResults.classList.remove("visible");
            }
        });

        // Clear button
        if (searchClear) {
            searchClear.addEventListener("click", function () {
                searchInput.value = "";
                searchClear.style.display = "none";
                searchResults.classList.remove("visible");
                searchResults.innerHTML = "";
                clearTimeout(searchDebounceTimer);
                if (searchAbortController) {
                    try { searchAbortController.abort(); } catch (e) {}
                    searchAbortController = null;
                }
                if (searchIcon) searchIcon.style.display = "block";
                if (searchSpinner) searchSpinner.style.display = "none";
                searchGeneration++;
            });
        }

        // Escape key + keyboard nav for results
        searchInput.addEventListener("keydown", function (e) {
            if (e.key === "Escape") {
                searchResults.classList.remove("visible");
                searchInput.blur();
                return;
            }
            var items = searchResults.querySelectorAll(".search-result-item");
            if (!items.length) return;
            var focused = searchResults.querySelector(".search-result-item.kb-focus");
            var idx = focused ? Array.prototype.indexOf.call(items, focused) : -1;
            if (e.key === "ArrowDown") {
                e.preventDefault();
                if (idx < items.length - 1) idx++; else idx = 0;
                items.forEach(function(it){ it.classList.remove('kb-focus'); });
                items[idx].classList.add('kb-focus');
                items[idx].scrollIntoView({ block: 'nearest' });
            } else if (e.key === "ArrowUp") {
                e.preventDefault();
                if (idx > 0) idx--; else idx = items.length - 1;
                items.forEach(function(it){ it.classList.remove('kb-focus'); });
                items[idx].classList.add('kb-focus');
                items[idx].scrollIntoView({ block: 'nearest' });
            } else if (e.key === "Enter" && focused) {
                e.preventDefault();
                focused.click();
            }
        });
    }

    // ── Public API ──
    return {
        init: init,
        setOnFeatureSelected: function (cb) { onFeatureSelected = cb; }
    };
});
