/**
 * summary-engine.js
 * ─────────────────
 * Generic attribute-summary engine with plugin overrides.
 *
 * Replaces the ~870-line if/else chain in final-report.js's
 * generateLayerAttributeSummary() with:
 *   1. A named-plugin registry (layer-specific field extraction).
 *   2. A generic auto-classifier that inspects field names and values.
 *
 * AMD module – no build step required.
 *
 * Usage (from final-report.js after wiring):
 *   const html = summaryEngine.generate(item);
 */
define([], function () {
    "use strict";

    /* ================================================================
     *  Utilities
     * ================================================================ */

    function escapeHtml(str) {
        if (str == null) return "";
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function formatNumber(n, decimals) {
        if (n == null || isNaN(n)) return "0";
        return Number(n).toLocaleString("en-US", {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        });
    }

    /** Render a frequency-count Map as "Label (count), …" */
    function freqSummary(map, limit) {
        limit = limit || 10;
        var items = Array.from(map.entries())
            .sort(function (a, b) { return b[1] - a[1]; })
            .slice(0, limit)
            .map(function (pair) {
                return escapeHtml(pair[0]) + " (" + pair[1] + ")";
            })
            .join(", ");
        if (map.size > limit) items += " …";
        return items;
    }

    /** Render a Set as "val, val, …" */
    function uniqueSummary(set, limit) {
        limit = limit || 10;
        var items = Array.from(set).slice(0, limit).map(function (v) {
            return escapeHtml(v);
        }).join(", ");
        if (set.size > limit) items += " …";
        return items;
    }

    /** Emit one <tr> */
    function tr(label, value) {
        return "<tr><td>" + escapeHtml(label) + "</td><td>" + value + "</td></tr>";
    }

    /* ================================================================
     *  Field-name classifiers  (regex → category)
     * ================================================================ */

    var FIELD_CLASSES = [
        { tag: "name",     rx: /^(.*_)?(NAME|NM|TITLE|LABEL|DESCRIPTION|DESC)(_.*)?$/i },
        { tag: "status",   rx: /^(.*_)?(STATUS|STAT|STATE|CONDITION|APPROVAL)(_.*)?$/i },
        { tag: "type",     rx: /^(.*_)?(TYPE|TYP|CLASS|CATEGORY|CAT|KIND|DESIGNATION)(_.*)?$/i },
        { tag: "date",     rx: /^(.*_)?(DATE|DT|CREATED|MODIFIED|EFFECTIVE|EXPIR|START|END)(_.*)?$/i },
        { tag: "year",     rx: /^(.*_)?(YEAR|YR|CY|FY|FISCAL_YEAR)(_.*)?$/i },
        { tag: "area",     rx: /^(.*_)?(ACRES|ACREAGE|GIS_ACRES|AREA|SQ_MILES|HECTARES|SQ_KM)(_.*)?$/i },
        { tag: "length",   rx: /^(.*_)?(LENGTH|MILES|FEET|METERS|DISTANCE)(_.*)?$/i },
        { tag: "count",    rx: /^(.*_)?(COUNT|TOTAL|NUM|NUMBER)(_.*)?$/i },
        { tag: "id",       rx: /^(.*_)?(CASE_NR|SERIAL_NR|CASE_NO|SERIAL_NO|CASENR|SERIALNR|CASE_ID|AUTH_NR|CASE_NUMBER|OBJECTID|FID|GLOBALID|SHAPE)(_.*)?$/i },
        { tag: "url",      rx: null } // detected by value, not field name
    ];

    /** Classify a single field name → tag string or null */
    function classifyField(fieldName) {
        for (var i = 0; i < FIELD_CLASSES.length; i++) {
            var cls = FIELD_CLASSES[i];
            if (cls.rx && cls.rx.test(fieldName)) return cls.tag;
        }
        return null;
    }

    var URL_RX = /^https?:\/\//i;
    var SKIP_FIELDS = /^(OBJECTID|FID|GLOBALID|ST_AREA|ST_LENGTH|ST_PERIMETER|SHAPE([-_. ].*)?|TOTAL[-_ ]?(AREA|LENGTH|ACRES|ACREAGE))$/i;

    /* ================================================================
     *  Generic auto-summary builder
     * ================================================================ */

    /**
     * Scan rows, auto-classify each field, aggregate into buckets.
     * Returns an object with named arrays/maps ready for rendering.
     */
    function autoClassify(rows) {
        var names     = new Set();
        var statuses  = new Map();
        var types     = new Map();
        var dates     = [];   // raw values
        var years     = new Map();
        var areaSum   = 0;
        var areaCount = 0;
        var lengthSum = 0;
        var lengthCnt = 0;
        var numericFields = {}; // fieldLabel → {min, max, sum, count}
        var urls      = new Map(); // label → Set<url>
        var ids       = new Set();
        var unclassified = new Map(); // fieldLabel → Map<value, count>

        for (var r = 0; r < rows.length; r++) {
            var row = rows[r];
            var keys = Object.keys(row);
            for (var k = 0; k < keys.length; k++) {
                var key = keys[k];
                if (SKIP_FIELDS.test(key)) continue;
                var val = row[key];
                if (val == null || val === "") continue;

                var strVal = String(val).trim();
                if (!strVal || strVal.length > 500) continue;

                // Value-based URL detection
                if (URL_RX.test(strVal)) {
                    var urlLabel = key.replace(/_/g, " ").replace(/\b\w/g, function (c) { return c.toUpperCase(); });
                    if (!urls.has(urlLabel)) urls.set(urlLabel, new Set());
                    urls.get(urlLabel).add(strVal);
                    continue;
                }

                var tag = classifyField(key);

                if (tag === "name") {
                    if (strVal.length <= 200) names.add(strVal);
                } else if (tag === "status") {
                    if (strVal.length <= 200) statuses.set(strVal, (statuses.get(strVal) || 0) + 1);
                } else if (tag === "type") {
                    if (strVal.length <= 200) types.set(strVal, (types.get(strVal) || 0) + 1);
                } else if (tag === "date") {
                    dates.push(strVal);
                } else if (tag === "year") {
                    var y = String(strVal);
                    years.set(y, (years.get(y) || 0) + 1);
                } else if (tag === "area") {
                    var av = parseFloat(val);
                    if (!isNaN(av) && av > 0) { areaSum += av; areaCount++; }
                } else if (tag === "length") {
                    var lv = parseFloat(val);
                    if (!isNaN(lv) && lv > 0) { lengthSum += lv; lengthCnt++; }
                } else if (tag === "count") {
                    // skip raw counts
                } else if (tag === "id") {
                    if (strVal.length <= 60) ids.add(strVal);
                } else {
                    // Numeric value detection
                    if (typeof val === "number" && !isNaN(val)) {
                        var nLabel = key.replace(/_/g, " ").replace(/\b\w/g, function (c) { return c.toUpperCase(); });
                        if (!numericFields[nLabel]) numericFields[nLabel] = { min: val, max: val, sum: val, count: 1 };
                        else {
                            var nf = numericFields[nLabel];
                            nf.min = Math.min(nf.min, val);
                            nf.max = Math.max(nf.max, val);
                            nf.sum += val;
                            nf.count++;
                        }
                    } else if (strVal.length <= 200) {
                        // Text field → frequency count under its label
                        var uLabel = key.replace(/_/g, " ").replace(/\b\w/g, function (c) { return c.toUpperCase(); });
                        if (!unclassified.has(uLabel)) unclassified.set(uLabel, new Map());
                        var m = unclassified.get(uLabel);
                        m.set(strVal, (m.get(strVal) || 0) + 1);
                    }
                }
            }
        }

        return {
            names: names,
            statuses: statuses,
            types: types,
            dates: dates,
            years: years,
            areaSum: areaSum,
            areaCount: areaCount,
            lengthSum: lengthSum,
            lengthCount: lengthCnt,
            numericFields: numericFields,
            urls: urls,
            ids: ids,
            unclassified: unclassified
        };
    }

    /**
     * Render classified data into HTML <tr> rows.
     */
    function renderClassified(data, headerLabel) {
        var html = "";

        var hasData = data.names.size > 0 || data.statuses.size > 0 ||
            data.types.size > 0 || data.years.size > 0 ||
            data.areaCount > 0 || data.lengthCount > 0 ||
            data.urls.size > 0 || data.ids.size > 0 ||
            Object.keys(data.numericFields).length > 0 ||
            data.unclassified.size > 0;

        if (!hasData) return "";

        html += '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";

        // Names
        if (data.names.size > 0) {
            html += tr("Names", uniqueSummary(data.names, 10));
        }

        // Types
        if (data.types.size > 0) {
            html += tr("Types", freqSummary(data.types, 10));
        }

        // Status
        if (data.statuses.size > 0) {
            html += tr("Status", freqSummary(data.statuses, 10));
        }

        // Years
        if (data.years.size > 0) {
            var sorted = Array.from(data.years.entries())
                .sort(function (a, b) { return b[0].localeCompare(a[0]); })
                .slice(0, 15)
                .map(function (p) { return escapeHtml(p[0]) + " (" + p[1] + ")"; })
                .join(", ");
            if (data.years.size > 15) sorted += " …";
            html += tr("Years", sorted);
        }

        // Area
        if (data.areaCount > 0) {
            html += tr("Total Area (acres)", formatNumber(data.areaSum, 1) +
                " (" + data.areaCount + " feature" + (data.areaCount !== 1 ? "s" : "") + ")");
        }

        // Length
        if (data.lengthCount > 0) {
            html += tr("Total Length", formatNumber(data.lengthSum, 1) +
                " (" + data.lengthCount + " feature" + (data.lengthCount !== 1 ? "s" : "") + ")");
        }

        // Numeric fields (min/max/mean for fields with >1 value)
        var numKeys = Object.keys(data.numericFields);
        for (var n = 0; n < numKeys.length; n++) {
            var nk = numKeys[n];
            var nf = data.numericFields[nk];
            if (nf.count > 1) {
                html += tr(nk, "min " + formatNumber(nf.min, 1) +
                    " / max " + formatNumber(nf.max, 1) +
                    " / avg " + formatNumber(nf.sum / nf.count, 1) +
                    " (" + nf.count + ")");
            } else {
                html += tr(nk, formatNumber(nf.sum, 1));
            }
        }

        // IDs
        if (data.ids.size > 0) {
            if (data.ids.size <= 10) {
                html += tr("Case/Serial Numbers", uniqueSummary(data.ids, 10));
            } else {
                html += tr("Case/Serial Numbers", data.ids.size + " records");
            }
        }

        // Unclassified text fields (show top 3 fields by unique-value count)
        if (data.unclassified.size > 0) {
            var sorted2 = Array.from(data.unclassified.entries())
                .sort(function (a, b) { return b[1].size - a[1].size; })
                .slice(0, 3);
            for (var u = 0; u < sorted2.length; u++) {
                var entry = sorted2[u];
                var label = entry[0];
                var valMap = entry[1];
                if (valMap.size <= 10) {
                    html += tr(label, freqSummary(valMap, 10));
                } else {
                    html += tr(label, valMap.size + " unique values");
                }
            }
        }

        // URLs
        if (data.urls.size > 0) {
            for (var iter = data.urls.entries(), step; !(step = iter.next()).done;) {
                var urlLabel = step.value[0];
                var urlSet = step.value[1];
                var links = Array.from(urlSet).slice(0, 5)
                    .map(function (u) {
                        var escaped = escapeHtml(u);
                        return '<a href="' + escaped + '" target="_blank" rel="noopener">' +
                            (escaped.length > 50 ? escaped.substring(0, 50) + "…" : escaped) + "</a>";
                    }).join("<br/>");
                if (urlSet.size > 5) links += "<br/>…";
                html += tr(urlLabel, links);
            }
        }

        return html;
    }

    /* ================================================================
     *  Plugin registry
     *  ────────────────
     *  Each plugin is a function(rows, headerLabel) → html string.
     *  If it returns "" or null, the engine falls through to generic.
     *
     *  Plugins inspect specific fields they know about. They are
     *  registered by key and matched via layer config's
     *  `summaryPlugin` string.
     * ================================================================ */

    var _plugins = {};

    /**
     * Register a named plugin.
     * @param {string}   name  Unique plugin key (e.g. "vri", "critical-habitat")
     * @param {Function} fn    function(rows, headerLabel) → html string
     */
    function registerPlugin(name, fn) {
        _plugins[name] = fn;
    }

    /* ──── Built-in plugins (migrated from the if/else chain) ──── */

    // Helper: collect unique values from a set of candidate fields
    function pick(row, candidates) {
        for (var i = 0; i < candidates.length; i++) {
            var v = row[candidates[i]];
            if (v != null && String(v).trim() !== "") return String(v).trim();
        }
        return "";
    }

    // Helper: collect frequency map from candidate fields
    function freqFromRows(rows, candidates) {
        var map = new Map();
        for (var r = 0; r < rows.length; r++) {
            var v = pick(rows[r], candidates);
            if (v) map.set(v, (map.get(v) || 0) + 1);
        }
        return map;
    }

    // Helper: collect unique set from candidate fields
    function setFromRows(rows, candidates) {
        var set = new Set();
        for (var r = 0; r < rows.length; r++) {
            var v = pick(rows[r], candidates);
            if (v) set.add(v);
        }
        return set;
    }

    // ── VRI ──
    registerPlugin("vri", function (rows, headerLabel) {
        var vriClassCounts = freqFromRows(rows, ["VRI_CLASS_CODE", "VRI_CLASS", "CLASS_CODE"]);
        var scenicRatingCounts = freqFromRows(rows, ["SL_OVRL_RT", "SCENIC_QUALITY", "SQ_RATING"]);
        if (vriClassCounts.size === 0 && scenicRatingCounts.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (vriClassCounts.size > 0) {
            var items = Array.from(vriClassCounts.entries())
                .sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])); })
                .map(function (p) { return escapeHtml(p[0]) + " (" + p[1] + ")"; }).join(", ");
            html += tr("VRI Class Codes", items);
        }
        if (scenicRatingCounts.size > 0) {
            var items2 = Array.from(scenicRatingCounts.entries())
                .sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])); })
                .map(function (p) { return escapeHtml(p[0]) + " (" + p[1] + ")"; }).join(", ");
            html += tr("Scenic Quality Ratings", items2);
        }
        return html;
    });

    // ── Critical Habitat ──
    registerPlugin("critical-habitat", function (rows, headerLabel) {
        var speciesCounts = freqFromRows(rows, ["COMNAME", "SCINAME", "SPECIES"]);
        if (speciesCounts.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        var items = Array.from(speciesCounts.entries())
            .sort(function (a, b) { return b[1] - a[1]; }).slice(0, 10)
            .map(function (p) { return escapeHtml(p[0]) + " (" + p[1] + ")"; }).join(", ");
        html += tr("Species", items + (speciesCounts.size > 10 ? " …" : ""));
        return html;
    });

    // ── Grazing Allotments ──
    registerPlugin("grazing-allotments", function (rows, headerLabel) {
        var names = setFromRows(rows, ["ALLOT_NAME", "ALLOTMENT_NAME", "NAME"]);
        if (names.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        html += tr("Allotment Names", uniqueSummary(names, 10));
        return html;
    });

    // ── Wilderness / WSA ──
    registerPlugin("wilderness", function (rows, headerLabel) {
        var names = setFromRows(rows, ["NLCS_NAME", "WSA_NAME", "NAME", "UNIT_NAME"]);
        if (names.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        html += tr("Area Names", uniqueSummary(names, 10));
        return html;
    });

    // ── ACEC ──
    registerPlugin("acec", function (rows, headerLabel) {
        var names = setFromRows(rows, ["ACEC_NAME", "NAME", "UNIT_NAME"]);
        if (names.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        html += tr("ACEC Names", uniqueSummary(names, 10));
        return html;
    });

    // ── Wild Horse / Burro ──
    registerPlugin("wild-horse-burro", function (rows, headerLabel) {
        var names = setFromRows(rows, ["HA_NAME", "HMA_NAME", "NAME"]);
        if (names.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        html += tr("Herd Area Names", uniqueSummary(names, 10));
        return html;
    });

    // ── Ungulate Migration ──
    registerPlugin("ungulate-migration", function (rows, headerLabel) {
        var speciesCounts = freqFromRows(rows, ["SPECIES", "COMMON_NAME"]);
        var useCounts = freqFromRows(rows, ["USE_TYPE", "SEASON", "MOVEMENT_TYPE"]);
        if (speciesCounts.size === 0 && useCounts.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (speciesCounts.size > 0) html += tr("Species", freqSummary(speciesCounts, 10));
        if (useCounts.size > 0) html += tr("Use Type / Season", freqSummary(useCounts, 10));
        return html;
    });

    // ── MLRS ROW / LUA ──
    registerPlugin("mlrs-row", function (rows, headerLabel) {
        var authTypes = freqFromRows(rows, ["AUTH_TYPE", "AUTHORIZATION_TYPE", "TYPE"]);
        var statusCounts = freqFromRows(rows, ["CASE_STATUS", "STATUS", "AUTH_STATUS"]);
        var caseNumbers = setFromRows(rows, ["CASE_NR", "SERIAL_NR", "CASE_NUMBER"]);
        if (authTypes.size === 0 && statusCounts.size === 0 && caseNumbers.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (authTypes.size > 0) html += tr("Authorization Types", freqSummary(authTypes, 10));
        if (statusCounts.size > 0) html += tr("Status", freqSummary(statusCounts, 10));
        if (caseNumbers.size > 0) {
            if (caseNumbers.size <= 10) html += tr("Case Numbers", uniqueSummary(caseNumbers, 10));
            else html += tr("Case Numbers", caseNumbers.size + " cases");
        }
        return html;
    });

    // ── Land Use Plans ──
    registerPlugin("land-use-plan", function (rows, headerLabel) {
        var planNames = setFromRows(rows, ["LUPName", "LUPNAME", "PLAN_NAME", "PLAN_NM", "RMP_NAME", "NAME", "LUP_NAME"]);
        var statusCounts = freqFromRows(rows, ["Status", "STATUS", "PLAN_STATUS", "APPROVAL_STATUS", "LUP_STATUS"]);
        var nepaNumbers = setFromRows(rows, ["NEPAnum", "NEPANUM", "NEPA_NUM", "NEPA_NUMBER"]);
        var rodYears = setFromRows(rows, ["RODyear", "RODYEAR", "ROD_YEAR", "ROD_YR"]);
        var epLinks = setFromRows(rows, ["ePLink", "EPLINK", "EP_LINK", "EPLANNING_LINK"]);
        if (planNames.size === 0 && statusCounts.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (planNames.size > 0) html += tr("LUP Name", uniqueSummary(planNames, 10));
        if (statusCounts.size > 0) html += tr("Status", freqSummary(statusCounts, 10));
        if (epLinks.size > 0) {
            var links = Array.from(epLinks).slice(0, 5).map(function (link) {
                var escaped = escapeHtml(link);
                return '<a href="' + escaped + '" target="_blank">' + (escaped.length > 50 ? escaped.substring(0, 50) + "…" : escaped) + "</a>";
            }).join("<br>");
            if (epLinks.size > 5) links += "<br>…";
            html += tr("ePlanning Link", links);
        }
        if (nepaNumbers.size > 0) html += tr("NEPA Number", uniqueSummary(nepaNumbers, 10));
        if (rodYears.size > 0) {
            var years = Array.from(rodYears).sort().map(function (y) { return escapeHtml(y); }).join(", ");
            html += tr("ROD Year", years);
        }
        return html;
    });

    // ── NLCS / Conservation Areas ──
    registerPlugin("nlcs", function (rows, headerLabel) {
        var names = setFromRows(rows, ["NLCS_NAME", "NCA_NAME", "NM_NAME", "NAME", "UNIT_NAME"]);
        var designations = freqFromRows(rows, ["DESIGNATION", "NLCS_TYPE", "TYPE"]);
        if (names.size === 0 && designations.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (names.size > 0) html += tr("Area Names", uniqueSummary(names, 10));
        if (designations.size > 0) html += tr("Designation Type", freqSummary(designations, 10));
        return html;
    });

    // ── Locatable Mineral Allocations ──
    registerPlugin("locatable-minerals", function (rows, headerLabel) {
        var allocations = freqFromRows(rows, ["LOC_ALLOC", "ALLOCATION", "ALLOC_TYPE", "MINERAL_ALLOCATION"]);
        var statusCounts = freqFromRows(rows, ["STATUS", "ALLOC_STATUS"]);
        if (allocations.size === 0 && statusCounts.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (allocations.size > 0) html += tr("Allocation Types", freqSummary(allocations, 10));
        if (statusCounts.size > 0) html += tr("Status", freqSummary(statusCounts, 10));
        return html;
    });

    // ── Timber Allocations ──
    registerPlugin("timber", function (rows, headerLabel) {
        var allocations = freqFromRows(rows, ["TIMBER_ALLOC", "ALLOCATION", "ALLOC_TYPE", "HARVEST_TYPE"]);
        var statusCounts = freqFromRows(rows, ["STATUS", "ALLOC_STATUS"]);
        if (allocations.size === 0 && statusCounts.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (allocations.size > 0) html += tr("Allocation Types", freqSummary(allocations, 10));
        if (statusCounts.size > 0) html += tr("Status", freqSummary(statusCounts, 10));
        return html;
    });

    // ── USFWS Regions ──
    registerPlugin("usfws-regions", function (rows, headerLabel) {
        var names = setFromRows(rows, ["REGNAME", "REGION_NAME", "REGION", "NAME"]);
        if (names.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        html += tr("Regions", uniqueSummary(names, 20));
        return html;
    });

    // ── Grazing Pastures ──
    registerPlugin("grazing-pastures", function (rows, headerLabel) {
        var pastureNames = setFromRows(rows, ["PASTURE_NAME", "PAST_NAME", "NAME"]);
        var allotmentNames = setFromRows(rows, ["ALLOT_NAME", "ALLOTMENT_NAME"]);
        if (pastureNames.size === 0 && allotmentNames.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (pastureNames.size > 0) html += tr("Pasture Names", uniqueSummary(pastureNames, 10));
        if (allotmentNames.size > 0) html += tr("Allotment Names", uniqueSummary(allotmentNames, 10));
        return html;
    });

    // ── Oil and Gas Leases ──
    registerPlugin("oil-gas", function (rows, headerLabel) {
        var statusCounts = freqFromRows(rows, ["CASE_STATUS", "LEASE_STATUS", "STATUS", "CASE_STAT", "STAT", "DISP_STATUS", "CASE_TYP", "AUTH_STAT"]);
        var typeCounts = freqFromRows(rows, ["LEASE_TYPE", "AUTH_TYPE", "TYPE", "CASE_TYPE", "TYP", "DISP_TYPE", "AUTH_TYP", "CASETYPE"]);
        var commodityCounts = freqFromRows(rows, ["COMMODITY", "CMDTY", "RESOURCE", "MINERAL", "PRODUCT"]);
        var lesseeNames = setFromRows(rows, ["HOLDER_NAME", "LESSEE", "LESSEE_NAME", "HOLDER", "COMPANY", "OPERATOR", "CUSTOMER_NAME", "CUST_NAME"]);
        var caseNumbers = setFromRows(rows, ["CASE_NR", "SERIAL_NR", "LEASE_NUMBER", "CASE_NO", "SERIAL_NO", "CASENR", "SERIALNR", "CASE_ID", "AUTH_NR"]);
        if (statusCounts.size === 0 && typeCounts.size === 0 && caseNumbers.size === 0 &&
            lesseeNames.size === 0 && commodityCounts.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (statusCounts.size > 0) html += tr("Lease Status", freqSummary(statusCounts, 10));
        if (typeCounts.size > 0) html += tr("Lease Type", freqSummary(typeCounts, 10));
        if (commodityCounts.size > 0) html += tr("Commodity", freqSummary(commodityCounts, 10));
        if (lesseeNames.size > 0) html += tr("Lessee/Holder", uniqueSummary(lesseeNames, 10));
        if (caseNumbers.size > 0) {
            if (caseNumbers.size <= 10) html += tr("Case/Serial Numbers", uniqueSummary(caseNumbers, 10));
            else html += tr("Case/Serial Numbers", caseNumbers.size + " leases");
        }
        return html;
    });

    // ── Recreation Sites ──
    registerPlugin("recreation-sites", function (rows, headerLabel) {
        var siteNames = setFromRows(rows, ["SITE_NAME", "REC_SITE_NAME", "NAME", "SITENAME", "SITE_NM", "REC_NAME"]);
        var siteTypes = freqFromRows(rows, ["SITE_TYPE", "REC_SITE_TYPE", "TYPE", "SITETYPE", "REC_TYPE", "FACILITY_TYPE"]);
        var feeCounts = freqFromRows(rows, ["FEE_YN", "FEE", "FEE_STATUS", "USER_FEE", "FEES"]);
        var activityTypes = setFromRows(rows, ["ACTIVITIES", "ACTIVITY", "REC_ACTIVITY", "PRIMARY_ACTIVITY", "USE_TYPE"]);
        if (siteNames.size === 0 && siteTypes.size === 0 && feeCounts.size === 0 && activityTypes.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (siteNames.size > 0) html += tr("Site Names", uniqueSummary(siteNames, 10));
        if (siteTypes.size > 0) html += tr("Site Types", freqSummary(siteTypes, 10));
        if (feeCounts.size > 0) html += tr("Fee Status", freqSummary(feeCounts, 10));
        if (activityTypes.size > 0) html += tr("Activities", uniqueSummary(activityTypes, 10));
        return html;
    });

    // ── LWCF ──
    registerPlugin("lwcf", function (rows, headerLabel) {
        var projectNames = setFromRows(rows, ["PROJECT_NAME", "PROJ_NAME", "NAME", "TRACT_NAME", "LWCF_NAME", "UNIT_NAME"]);
        var statusCounts = freqFromRows(rows, ["STATUS", "PROJ_STATUS", "PROJECT_STATUS", "LWCF_STATUS", "ACQ_STATUS"]);
        var purposeCounts = freqFromRows(rows, ["PURPOSE", "LWCF_PURPOSE", "USE", "PROJECT_TYPE", "PROJ_TYPE", "ACQ_TYPE"]);
        var fiscalYears = setFromRows(rows, ["FISCAL_YEAR", "FY", "YEAR", "ACQ_YEAR"]);
        if (projectNames.size === 0 && statusCounts.size === 0 && purposeCounts.size === 0 && fiscalYears.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (projectNames.size > 0) html += tr("Project/Tract Names", uniqueSummary(projectNames, 10));
        if (purposeCounts.size > 0) html += tr("Purpose/Type", freqSummary(purposeCounts, 10));
        if (statusCounts.size > 0) html += tr("Status", freqSummary(statusCounts, 10));
        if (fiscalYears.size > 0) {
            var years = Array.from(fiscalYears).sort().slice(0, 10).map(function (y) { return escapeHtml(String(y)); }).join(", ");
            if (fiscalYears.size > 10) years += " …";
            html += tr("Fiscal Years", years);
        }
        return html;
    });

    // ── ePlanning Projects ──
    registerPlugin("eplanning", function (rows, headerLabel) {
        var projectNames = setFromRows(rows, ["PROJECT_NAME", "PROJ_NAME", "NAME"]);
        var statusCounts = freqFromRows(rows, ["NEPA_STATUS", "PROJECT_STATUS", "STATUS"]);
        var typeCounts = freqFromRows(rows, ["PROJECT_TYPE", "NEPA_TYPE", "TYPE"]);
        var nepaNumbers = setFromRows(rows, ["NEPA_NUMBER", "NEPA_NO", "DOI_NUMBER"]);
        if (projectNames.size === 0 && statusCounts.size === 0 && typeCounts.size === 0 && nepaNumbers.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (projectNames.size > 0) html += tr("Project Names", uniqueSummary(projectNames, 10));
        if (typeCounts.size > 0) html += tr("Project Type", freqSummary(typeCounts, 10));
        if (statusCounts.size > 0) html += tr("NEPA Status", freqSummary(statusCounts, 10));
        if (nepaNumbers.size > 0) {
            if (nepaNumbers.size <= 10) html += tr("NEPA Numbers", uniqueSummary(nepaNumbers, 10));
            else html += tr("NEPA Numbers", nepaNumbers.size + " projects");
        }
        return html;
    });

    // ── Fire Perimeters ──
    registerPlugin("fire-perimeters", function (rows, headerLabel) {
        var fireNames = setFromRows(rows, ["INCDNT_NM"]);
        var causeCounts = freqFromRows(rows, ["FIRE_CAUSE_NM"]);
        var discoveryYears = freqFromRows(rows, ["FIRE_DSCVR_CY"]);
        var adminStates = freqFromRows(rows, ["ADMIN_ST"]);
        var complexNames = setFromRows(rows, ["CMPLX_NM"]);
        var totalAcres = [];
        for (var r = 0; r < rows.length; r++) {
            var acres = parseFloat(rows[r].GIS_ACRES || rows[r].TOTAL_RPT_ACRES_NR || 0);
            if (acres > 0) totalAcres.push(acres);
        }
        if (fireNames.size === 0 && causeCounts.size === 0 && discoveryYears.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (fireNames.size > 0) html += tr("Fire Names", uniqueSummary(fireNames, 15));
        if (causeCounts.size > 0) html += tr("Fire Cause", freqSummary(causeCounts, 10));
        if (discoveryYears.size > 0) {
            var sorted = Array.from(discoveryYears.entries())
                .sort(function (a, b) { return String(b[0]).localeCompare(String(a[0])); })
                .slice(0, 15)
                .map(function (p) { return escapeHtml(String(p[0])) + " (" + p[1] + ")"; }).join(", ");
            if (discoveryYears.size > 15) sorted += " …";
            html += tr("Discovery Years", sorted);
        }
        if (adminStates.size > 0) html += tr("Admin State", freqSummary(adminStates, 10));
        if (totalAcres.length > 0) {
            var sum = totalAcres.reduce(function (a, b) { return a + b; }, 0);
            html += tr("Total Burned Acres", formatNumber(sum, 1) + " (" + totalAcres.length + " fires)");
        }
        if (complexNames.size > 0) html += tr("Fire Complexes", uniqueSummary(complexNames, 10));
        return html;
    });

    // ── Administrative Units ──
    registerPlugin("admin-units", function (rows, headerLabel) {
        var unitNames = setFromRows(rows, ["ADMU_NAME", "Label_Full_Name", "Label"]);
        var orgTypes = freqFromRows(rows, ["BLM_ORG_TYPE"]);
        var adminStates = freqFromRows(rows, ["ADMIN_ST"]);
        var parentNames = setFromRows(rows, ["PARENT_NAME"]);
        var cityLabels = setFromRows(rows, ["City_Label"]);
        var stateUrls = setFromRows(rows, ["ADMU_ST_URL"]);
        if (unitNames.size === 0 && orgTypes.size === 0 && adminStates.size === 0) return "";
        var html = '<tr><td colspan="2" style="padding-top:12px;"><b>' + escapeHtml(headerLabel) + "</b></td></tr>";
        if (unitNames.size > 0) html += tr("Unit Names", uniqueSummary(unitNames, 15));
        if (orgTypes.size > 0) html += tr("Organization Type", freqSummary(orgTypes, 10));
        if (adminStates.size > 0) html += tr("Admin State", freqSummary(adminStates, 10));
        if (parentNames.size > 0) html += tr("Parent Office", uniqueSummary(parentNames, 10));
        if (cityLabels.size > 0) html += tr("Office Location", uniqueSummary(cityLabels, 10));
        if (stateUrls.size > 0) {
            var urls = Array.from(stateUrls).slice(0, 5).map(function (u) {
                return '<a href="' + escapeHtml(u) + '" target="_blank" rel="noopener">' + escapeHtml(u) + "</a>";
            }).join("<br/>");
            html += tr("State Office URL", urls);
        }
        return html;
    });

    /* ================================================================
     *  Title → plugin auto-match (legacy fallback)
     *  ────────────────────────────────────────────
     *  Used when a layer has no `summaryPlugin` set in config.
     *  Replicates the original title.includes() logic so that
     *  existing layers work identically without config changes.
     * ================================================================ */

    var TITLE_MATCHERS = [
        { test: function (t) { return t.includes("visual resource inventory") || t.includes("vri"); }, plugin: "vri" },
        { test: function (t) { return t.includes("critical habitat"); }, plugin: "critical-habitat" },
        { test: function (t) { return t.includes("grazing allotment"); }, plugin: "grazing-allotments" },
        { test: function (t) { return t.includes("wilderness"); }, plugin: "wilderness" },
        { test: function (t) { return t.includes("acec") || t.includes("critical environmental concern"); }, plugin: "acec" },
        { test: function (t) { return t.includes("wild horse") || t.includes("burro"); }, plugin: "wild-horse-burro" },
        { test: function (t) { return t.includes("ungulate") || t.includes("migration"); }, plugin: "ungulate-migration" },
        { test: function (t) { return t.includes("mlrs") && (t.includes("row") || t.includes("lua")); }, plugin: "mlrs-row" },
        { test: function (t) { return t.includes("land use plan") || (t.includes("revision") && t.includes("development")); }, plugin: "land-use-plan" },
        { test: function (t) { return t.includes("nlcs") || t.includes("conservation area") || t.includes("national monument"); }, plugin: "nlcs" },
        { test: function (t) { return t.includes("locatable") || t.includes("mineral allocation"); }, plugin: "locatable-minerals" },
        { test: function (t) { return t.includes("timber"); }, plugin: "timber" },
        { test: function (t) { return t.includes("usfws") && t.includes("region"); }, plugin: "usfws-regions" },
        { test: function (t) { return t.includes("grazing pasture") || t.includes("pasture polygon"); }, plugin: "grazing-pastures" },
        { test: function (t) { return t.includes("oil") && t.includes("gas"); }, plugin: "oil-gas" },
        { test: function (t) { return t.includes("recreation site") || t.includes("recreation_site"); }, plugin: "recreation-sites" },
        { test: function (t) { return t.includes("lwcf") || t.includes("land and water conservation") || t.includes("conservation fund"); }, plugin: "lwcf" },
        { test: function (t) { return t.includes("eplanning") || t.includes("epl_comment") || t.includes("nepa"); }, plugin: "eplanning" },
        { test: function (t) { return t.includes("fire perimeter"); }, plugin: "fire-perimeters" },
        { test: function (t) { return t.includes("administrative unit") || t.includes("admin unit") || t.includes("adminunit"); }, plugin: "admin-units" }
    ];

    function resolvePlugin(item, layerCfg) {
        // 1. Explicit config-level plugin name
        if (layerCfg && layerCfg.summaryPlugin && _plugins[layerCfg.summaryPlugin]) {
            return _plugins[layerCfg.summaryPlugin];
        }

        // 2. Title-based auto-match (backwards compatible)
        var title = ((item && item.title) || "").toLowerCase();
        for (var i = 0; i < TITLE_MATCHERS.length; i++) {
            if (TITLE_MATCHERS[i].test(title) && _plugins[TITLE_MATCHERS[i].plugin]) {
                return _plugins[TITLE_MATCHERS[i].plugin];
            }
        }

        return null;
    }

    /* ================================================================
     *  Append-URL pass (matches original behaviour)
     *  ─────────────────────────────────────────────
     *  After plugin or generic summary is built, scan for any
     *  URL-valued fields not already present and append them.
     * ================================================================ */

    function appendUrlRows(summaryHtml, rows) {
        if (!summaryHtml || !rows || !rows.length) return summaryHtml;
        var collectedUrls = new Map();
        for (var r = 0; r < rows.length; r++) {
            var row = rows[r];
            var keys = Object.keys(row);
            for (var k = 0; k < keys.length; k++) {
                var val = row[keys[k]];
                if (val == null || val === "") continue;
                var strVal = String(val).trim();
                if (!strVal || strVal.length > 500) continue;
                if (URL_RX.test(strVal)) {
                    var label = keys[k].replace(/_/g, " ").replace(/\b\w/g, function (c) { return c.toUpperCase(); });
                    if (!collectedUrls.has(label)) collectedUrls.set(label, new Set());
                    collectedUrls.get(label).add(strVal);
                }
            }
        }
        if (collectedUrls.size === 0) return summaryHtml;

        var urlHtml = "";
        for (var iter = collectedUrls.entries(), step; !(step = iter.next()).done;) {
            var label2 = step.value[0];
            var urlSet = step.value[1];
            var firstUrl = Array.from(urlSet)[0];
            if (summaryHtml.includes(escapeHtml(firstUrl))) continue;
            var links = Array.from(urlSet).slice(0, 5).map(function (u) {
                var escaped = escapeHtml(u);
                return '<a href="' + escaped + '" target="_blank" rel="noopener">' + escaped + "</a>";
            }).join("<br/>");
            if (urlSet.size > 5) links += "<br/>…";
            urlHtml += tr(label2, links);
        }
        if (urlHtml) {
            summaryHtml += '<tr><td colspan="2" style="padding-top:8px;"><b>Links</b></td></tr>';
            summaryHtml += urlHtml;
        }
        return summaryHtml;
    }

    /* ================================================================
     *  Public API
     * ================================================================ */

    var S = null; // shared state ref (for layerCfgByUrl lookups)

    function init(state) {
        S = state;

        return {
            /**
             * Generate HTML table-row summary for a report item.
             *
             * @param {Object} item  Report item with .title, .url,
             *   .fullRows/.rows (array of attribute objects)
             * @returns {string} HTML <tr>…</tr> string (empty if no data)
             */
            generate: function (item) {
                if (!item) return "";

                var rows = item.fullRows || item.rows || [];
                if (!rows.length) return "";

                var displayTitle = escapeHtml(item.title || "Layer");
                var headerLabel = displayTitle + " Layer Highlights";

                // Resolve layer config for plugin lookup
                var layerCfg = null;
                if (S && S.layerCfgByUrl && item.url) {
                    var entry = S.layerCfgByUrl.get(item.url);
                    if (entry) layerCfg = entry.cfg || entry;
                }

                // Try plugin first
                var pluginFn = resolvePlugin(item, layerCfg);
                var html = "";
                if (pluginFn) {
                    html = pluginFn(rows, headerLabel) || "";
                }

                // Fall back to generic auto-classifier
                if (!html) {
                    var data = autoClassify(rows);
                    html = renderClassified(data, headerLabel);
                }

                // Append URL rows that weren't already included
                html = appendUrlRows(html, rows);

                return html;
            },

            /**
             * Register a custom summary plugin.
             * Call this to add domain-specific summarizers at runtime.
             *
             * @param {string}   name  Plugin key
             * @param {Function} fn    function(rows, headerLabel) → html
             */
            registerPlugin: registerPlugin,

            /**
             * List registered plugin names (for debugging / admin UI).
             */
            listPlugins: function () {
                return Object.keys(_plugins);
            }
        };
    }

    return { init: init };
});
