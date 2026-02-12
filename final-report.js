/**
 * final-report.js  –  AMD module for Final Report generation
 *
 * Exports every function that participates in building the
 * "Final Report" HTML document, including the ~900-line
 * generateLayerAttributeSummary, the HTML template, the AOI
 * map compositor, and the main buildFinalReportHtml orchestrator.
 *
 * Usage (inside the main require callback):
 *   const fr = finalReportModule.init(state, deps);
 *   const { buildFinalReportHtml, viewFinalReport, ... } = fr;
 */
define([
    "app/config-helpers"
], function (configHelpers) {
    "use strict";

    const {
        escapeHtml, formatNumber, plssToolLabel,
        getConfiguredServices, flattenAttributes
    } = configHelpers;

    // ── Module-private state (set by init) ──
    let S;            // shared state proxy
    let mapUtils;     // map-utils API
    let queryEngine;  // query-engine API
    let ImageryLayer; // Esri constructor
    let FeatureLayer; // Esri constructor
    let geometryEngine; // Esri geometryEngine

    // Cached final report HTML (module-private)
    let cachedFinalReportHtml = null;
    // Reference to DOM element
    let finalReportStatus = null;
    // External helpers injected from app.js
    let _setStatus = () => {};

    // ────────────────────────────────────────────
    // Pure helpers (zero external deps beyond configHelpers)
    // ────────────────────────────────────────────

    function formatLegalDescription(row) {
        if (!row) return "";

        const twnshpNo  = row.TWNSHPNO  || "";
        const twnshpDir = row.TWNSHPDIR || "";
        const rangeNo   = row.RANGENO   || "";
        const rangeDir  = row.RANGEDIR  || "";
        const frstDivTxt = row.FRSTDIVTXT || "";
        const frstDivNo  = row.FRSTDIVNO  || "";
        const secDivTyp  = row.SECDIVTYP  || "";
        const secDivNo   = row.SECDIVNO   || "";

        let desc = "";
        if (twnshpNo && twnshpDir && rangeNo && rangeDir) {
            desc = `Township ${twnshpNo} ${twnshpDir} Range: ${rangeNo} ${rangeDir}`;
            if (frstDivTxt || frstDivNo) {
                desc += ` ${frstDivTxt} ${frstDivNo}`.trim();
            }
            if (secDivTyp && secDivNo) {
                desc += ` ${secDivTyp} ${secDivNo}`;
            } else if (secDivTyp) {
                desc += ` ${secDivTyp}`;
            } else if (secDivNo) {
                desc += ` ${secDivNo}`;
            }
        }
        return desc.trim();
    }

    function openHtmlInNewTab(htmlString) {
        const blob = new Blob([htmlString], { type: "text/html;charset=utf-8" });
        const url  = URL.createObjectURL(blob);
        const win  = window.open(url, "_blank", "noopener,noreferrer");
        window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
        return win;
    }

    function formatDateTimeForReport(d = new Date()) {
        try { return d.toLocaleString(); }
        catch (e) { return d.toString(); }
    }

    // ────────────────────────────────────────────
    // generateLayerAttributeSummary  (~900 lines)
    // ────────────────────────────────────────────
    function generateLayerAttributeSummary(item) {
        if (!item) return "";

        const title = (item.title || "").toLowerCase();
        const rows  = item.fullRows || item.rows || [];
        if (!rows.length) return "";

        let summaryHtml = "";

        // ── BLM National Visual Resource Inventory Classes ──
        if (title.includes("visual resource inventory") || title.includes("vri")) {
            const vriClassCounts   = new Map();
            const scenicRatingCounts = new Map();
            for (const row of rows) {
                const vriClass = row.VRI_CLASS_CODE || row.VRI_CLASS || row.CLASS_CODE || "";
                if (vriClass) vriClassCounts.set(vriClass, (vriClassCounts.get(vriClass) || 0) + 1);
                const scenicRating = row.SL_OVRL_RT || row.SCENIC_QUALITY || row.SQ_RATING || "";
                if (scenicRating) scenicRatingCounts.set(scenicRating, (scenicRatingCounts.get(scenicRating) || 0) + 1);
            }
            if (vriClassCounts.size > 0 || scenicRatingCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>VRI Attributes</b></td></tr>`;
                if (vriClassCounts.size > 0) {
                    const classItems = Array.from(vriClassCounts.entries())
                        .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
                        .map(([cls, count]) => `${escapeHtml(cls)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>VRI Class Codes</td><td>${classItems}</td></tr>`;
                }
                if (scenicRatingCounts.size > 0) {
                    const ratingItems = Array.from(scenicRatingCounts.entries())
                        .sort((a, b) => String(a[0]).localeCompare(String(b[0])))
                        .map(([rating, count]) => `${escapeHtml(rating)} (${count})`)
                        .join(", ");
                    summaryHtml += `<tr><td>Scenic Quality Ratings</td><td>${ratingItems}</td></tr>`;
                }
            }
        }

        // ── USFWS Critical Habitat ──
        if (title.includes("critical habitat")) {
            const speciesCounts = new Map();
            const statusCounts  = new Map();
            for (const row of rows) {
                const species = row.COMNAME || row.SCINAME || row.SPECIES || "";
                if (species) speciesCounts.set(species, (speciesCounts.get(species) || 0) + 1);
                const status = row.STATUS || row.LISTING_STATUS || "";
                if (status) statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
            }
            if (speciesCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Critical Habitat Details</b></td></tr>`;
                const speciesItems = Array.from(speciesCounts.entries())
                    .sort((a, b) => b[1] - a[1]).slice(0, 10)
                    .map(([sp, count]) => `${escapeHtml(sp)} (${count})`)
                    .join(", ");
                summaryHtml += `<tr><td>Species</td><td>${speciesItems}${speciesCounts.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }

        // ── BLM Grazing Allotments ──
        if (title.includes("grazing allotment")) {
            const allotmentNames = new Set();
            for (const row of rows) {
                const name = row.ALLOT_NAME || row.ALLOTMENT_NAME || row.NAME || "";
                if (name) allotmentNames.add(name);
            }
            if (allotmentNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Grazing Allotment Details</b></td></tr>`;
                const names = Array.from(allotmentNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>Allotment Names</td><td>${names}${allotmentNames.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }

        // ── BLM Wilderness Areas / WSA ──
        if (title.includes("wilderness")) {
            const areaNames = new Set();
            for (const row of rows) {
                const name = row.NLCS_NAME || row.WSA_NAME || row.NAME || row.UNIT_NAME || "";
                if (name) areaNames.add(name);
            }
            if (areaNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Wilderness Area Details</b></td></tr>`;
                const names = Array.from(areaNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>Area Names</td><td>${names}${areaNames.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }

        // ── BLM ACEC ──
        if (title.includes("acec") || title.includes("critical environmental concern")) {
            const acecNames = new Set();
            for (const row of rows) {
                const name = row.ACEC_NAME || row.NAME || row.UNIT_NAME || "";
                if (name) acecNames.add(name);
            }
            if (acecNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>ACEC Details</b></td></tr>`;
                const names = Array.from(acecNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>ACEC Names</td><td>${names}${acecNames.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }

        // ── Wild Horse and Burro Areas ──
        if (title.includes("wild horse") || title.includes("burro")) {
            const herdNames = new Set();
            for (const row of rows) {
                const name = row.HA_NAME || row.HMA_NAME || row.NAME || "";
                if (name) herdNames.add(name);
            }
            if (herdNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Herd Area Details</b></td></tr>`;
                const names = Array.from(herdNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>Herd Area Names</td><td>${names}${herdNames.size > 10 ? " ..." : ""}</td></tr>`;
            }
        }

        // ── Ungulate Migration ──
        if (title.includes("ungulate") || title.includes("migration")) {
            const speciesCounts = new Map();
            const useCounts     = new Map();
            for (const row of rows) {
                const species = row.SPECIES || row.COMMON_NAME || "";
                if (species) speciesCounts.set(species, (speciesCounts.get(species) || 0) + 1);
                const use = row.USE_TYPE || row.SEASON || row.MOVEMENT_TYPE || "";
                if (use) useCounts.set(use, (useCounts.get(use) || 0) + 1);
            }
            if (speciesCounts.size > 0 || useCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Migration Corridor Details</b></td></tr>`;
                if (speciesCounts.size > 0) {
                    const speciesItems = Array.from(speciesCounts.entries())
                        .map(([sp, count]) => `${escapeHtml(sp)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Species</td><td>${speciesItems}</td></tr>`;
                }
                if (useCounts.size > 0) {
                    const useItems = Array.from(useCounts.entries())
                        .map(([u, count]) => `${escapeHtml(u)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Use Type / Season</td><td>${useItems}</td></tr>`;
                }
            }
        }

        // ── BLM MLRS LUA ROW (Rights of Way) ──
        if (title.includes("mlrs") && (title.includes("row") || title.includes("lua"))) {
            const authTypes    = new Map();
            const statusCounts = new Map();
            const caseNumbers  = new Set();
            for (const row of rows) {
                const authType = row.AUTH_TYPE || row.AUTHORIZATION_TYPE || row.TYPE || "";
                if (authType) authTypes.set(authType, (authTypes.get(authType) || 0) + 1);
                const status = row.CASE_STATUS || row.STATUS || row.AUTH_STATUS || "";
                if (status) statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                const caseNo = row.CASE_NR || row.SERIAL_NR || row.CASE_NUMBER || "";
                if (caseNo) caseNumbers.add(caseNo);
            }
            if (authTypes.size > 0 || statusCounts.size > 0 || caseNumbers.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>ROW Authorization Details</b></td></tr>`;
                if (authTypes.size > 0) {
                    const items = Array.from(authTypes.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Authorization Types</td><td>${items}</td></tr>`;
                }
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
                if (caseNumbers.size > 0 && caseNumbers.size <= 10) {
                    const items = Array.from(caseNumbers).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Case Numbers</td><td>${items}</td></tr>`;
                } else if (caseNumbers.size > 10) {
                    summaryHtml += `<tr><td>Case Numbers</td><td>${caseNumbers.size} cases</td></tr>`;
                }
            }
        }

        // ── BLM National Land Use Plans ──
        if (title.includes("land use plan") || title.includes("revision") && title.includes("development")) {
            const planNames   = new Set();
            const statusCounts = new Map();
            const epLinks     = new Set();
            const nepaNumbers = new Set();
            const rodYears    = new Set();
            for (const row of rows) {
                const name = row.LUPName || row.LUPNAME || row.PLAN_NAME || row.PLAN_NM ||
                             row.RMP_NAME || row.NAME || row.LUP_NAME || "";
                if (name) planNames.add(name);
                const status = row.Status || row.STATUS || row.PLAN_STATUS ||
                               row.APPROVAL_STATUS || row.LUP_STATUS || "";
                if (status) statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                const epLink = row.ePLink || row.EPLINK || row.EP_LINK || row.EPLANNING_LINK || "";
                if (epLink) epLinks.add(epLink);
                const nepaNum = row.NEPAnum || row.NEPANUM || row.NEPA_NUM || row.NEPA_NUMBER || "";
                if (nepaNum) nepaNumbers.add(nepaNum);
                const rodYear = row.RODyear || row.RODYEAR || row.ROD_YEAR || row.ROD_YR || "";
                if (rodYear) rodYears.add(String(rodYear));
            }
            if (planNames.size > 0 || statusCounts.size > 0 || epLinks.size > 0 || nepaNumbers.size > 0 || rodYears.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Land Use Plan Details</b></td></tr>`;
                if (planNames.size > 0) {
                    const names = Array.from(planNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>LUP Name</td><td>${names}${planNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
                if (epLinks.size > 0) {
                    const links = Array.from(epLinks).slice(0, 5).map(link => {
                        const escaped = escapeHtml(link);
                        return `<a href="${escaped}" target="_blank">${escaped.length > 50 ? escaped.substring(0, 50) + "..." : escaped}</a>`;
                    }).join("<br>");
                    summaryHtml += `<tr><td>ePlanning Link</td><td>${links}${epLinks.size > 5 ? "<br>..." : ""}</td></tr>`;
                }
                if (nepaNumbers.size > 0) {
                    const nums = Array.from(nepaNumbers).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>NEPA Number</td><td>${nums}${nepaNumbers.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (rodYears.size > 0) {
                    const years = Array.from(rodYears).sort().map(y => escapeHtml(y)).join(", ");
                    summaryHtml += `<tr><td>ROD Year</td><td>${years}</td></tr>`;
                }
            }
        }

        // ── BLM National Conservation Areas (NLCS) ──
        if (title.includes("nlcs") || title.includes("conservation area") || title.includes("national monument")) {
            const areaNames    = new Set();
            const designations = new Map();
            for (const row of rows) {
                const name = row.NLCS_NAME || row.NCA_NAME || row.NM_NAME || row.NAME || row.UNIT_NAME || "";
                if (name) areaNames.add(name);
                const desig = row.DESIGNATION || row.NLCS_TYPE || row.TYPE || "";
                if (desig) designations.set(desig, (designations.get(desig) || 0) + 1);
            }
            if (areaNames.size > 0 || designations.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Conservation Area Details</b></td></tr>`;
                if (areaNames.size > 0) {
                    const names = Array.from(areaNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Area Names</td><td>${names}${areaNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (designations.size > 0) {
                    const items = Array.from(designations.entries())
                        .map(([d, count]) => `${escapeHtml(d)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Designation Type</td><td>${items}</td></tr>`;
                }
            }
        }

        // ── BLM Locatable Mineral Allocations ──
        if (title.includes("locatable") || title.includes("mineral allocation")) {
            const allocations  = new Map();
            const statusCounts = new Map();
            for (const row of rows) {
                const alloc = row.LOC_ALLOC || row.ALLOCATION || row.ALLOC_TYPE || row.MINERAL_ALLOCATION || "";
                if (alloc) allocations.set(alloc, (allocations.get(alloc) || 0) + 1);
                const status = row.STATUS || row.ALLOC_STATUS || "";
                if (status) statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
            }
            if (allocations.size > 0 || statusCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Mineral Allocation Details</b></td></tr>`;
                if (allocations.size > 0) {
                    const items = Array.from(allocations.entries())
                        .map(([a, count]) => `${escapeHtml(a)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Allocation Types</td><td>${items}</td></tr>`;
                }
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
            }
        }

        // ── BLM Timber Allocations ──
        if (title.includes("timber")) {
            const allocations  = new Map();
            const statusCounts = new Map();
            for (const row of rows) {
                const alloc = row.TIMBER_ALLOC || row.ALLOCATION || row.ALLOC_TYPE || row.HARVEST_TYPE || "";
                if (alloc) allocations.set(alloc, (allocations.get(alloc) || 0) + 1);
                const status = row.STATUS || row.ALLOC_STATUS || "";
                if (status) statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
            }
            if (allocations.size > 0 || statusCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Timber Allocation Details</b></td></tr>`;
                if (allocations.size > 0) {
                    const items = Array.from(allocations.entries())
                        .map(([a, count]) => `${escapeHtml(a)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Allocation Types</td><td>${items}</td></tr>`;
                }
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
            }
        }

        // ── USFWS Regions ──
        if (title.includes("usfws") && title.includes("region")) {
            const regionNames = new Set();
            for (const row of rows) {
                const region = row.REGNAME || row.REGION_NAME || row.REGION || row.NAME || "";
                if (region) regionNames.add(region);
            }
            if (regionNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>USFWS Region Details</b></td></tr>`;
                const names = Array.from(regionNames).map(n => escapeHtml(n)).join(", ");
                summaryHtml += `<tr><td>Regions</td><td>${names}</td></tr>`;
            }
        }

        // ── BLM Grazing Pasture ──
        if (title.includes("grazing pasture") || title.includes("pasture polygon")) {
            const pastureNames   = new Set();
            const allotmentNames = new Set();
            for (const row of rows) {
                const pasture = row.PASTURE_NAME || row.PAST_NAME || row.NAME || "";
                if (pasture) pastureNames.add(pasture);
                const allotment = row.ALLOT_NAME || row.ALLOTMENT_NAME || "";
                if (allotment) allotmentNames.add(allotment);
            }
            if (pastureNames.size > 0 || allotmentNames.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Grazing Pasture Details</b></td></tr>`;
                if (pastureNames.size > 0) {
                    const names = Array.from(pastureNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Pasture Names</td><td>${names}${pastureNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (allotmentNames.size > 0) {
                    const names = Array.from(allotmentNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Allotment Names</td><td>${names}${allotmentNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
            }
        }

        // ── BLM Oil and Gas Leases ──
        if (title.includes("oil") && title.includes("gas")) {
            const statusCounts   = new Map();
            const typeCounts     = new Map();
            const caseNumbers    = new Set();
            const lesseeNames    = new Set();
            const commodityCounts = new Map();
            for (const row of rows) {
                const status = row.CASE_STATUS || row.LEASE_STATUS || row.STATUS ||
                               row.CASE_STAT || row.STAT || row.DISP_STATUS ||
                               row.CASE_TYP || row.AUTH_STAT || "";
                if (status) statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                const type = row.LEASE_TYPE || row.AUTH_TYPE || row.TYPE ||
                             row.CASE_TYPE || row.TYP || row.DISP_TYPE ||
                             row.AUTH_TYP || row.CASETYPE || "";
                if (type) typeCounts.set(type, (typeCounts.get(type) || 0) + 1);
                const caseNo = row.CASE_NR || row.SERIAL_NR || row.LEASE_NUMBER ||
                               row.CASE_NO || row.SERIAL_NO || row.CASENR ||
                               row.SERIALNR || row.CASE_ID || row.AUTH_NR || "";
                if (caseNo) caseNumbers.add(caseNo);
                const lessee = row.HOLDER_NAME || row.LESSEE || row.LESSEE_NAME ||
                               row.HOLDER || row.COMPANY || row.OPERATOR ||
                               row.CUSTOMER_NAME || row.CUST_NAME || "";
                if (lessee) lesseeNames.add(lessee);
                const commodity = row.COMMODITY || row.CMDTY || row.RESOURCE ||
                                  row.MINERAL || row.PRODUCT || "";
                if (commodity) commodityCounts.set(commodity, (commodityCounts.get(commodity) || 0) + 1);
            }
            if (statusCounts.size > 0 || typeCounts.size > 0 || caseNumbers.size > 0 ||
                lesseeNames.size > 0 || commodityCounts.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Oil & Gas Lease Details</b></td></tr>`;
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Lease Status</td><td>${items}</td></tr>`;
                }
                if (typeCounts.size > 0) {
                    const items = Array.from(typeCounts.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Lease Type</td><td>${items}</td></tr>`;
                }
                if (commodityCounts.size > 0) {
                    const items = Array.from(commodityCounts.entries())
                        .map(([c, count]) => `${escapeHtml(c)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Commodity</td><td>${items}</td></tr>`;
                }
                if (lesseeNames.size > 0) {
                    const names = Array.from(lesseeNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Lessee/Holder</td><td>${names}${lesseeNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (caseNumbers.size > 0 && caseNumbers.size <= 10) {
                    const items = Array.from(caseNumbers).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Case/Serial Numbers</td><td>${items}</td></tr>`;
                } else if (caseNumbers.size > 10) {
                    summaryHtml += `<tr><td>Case/Serial Numbers</td><td>${caseNumbers.size} leases</td></tr>`;
                }
            }
        }

        // ── BLM National Recreation Sites ──
        if (title.includes("recreation site") || title.includes("recreation_site")) {
            const siteNames     = new Set();
            const siteTypes     = new Map();
            const feeCounts     = new Map();
            const activityTypes = new Set();
            for (const row of rows) {
                const name = row.SITE_NAME || row.REC_SITE_NAME || row.NAME ||
                             row.SITENAME || row.SITE_NM || row.REC_NAME || "";
                if (name) siteNames.add(name);
                const type = row.SITE_TYPE || row.REC_SITE_TYPE || row.TYPE ||
                             row.SITETYPE || row.REC_TYPE || row.FACILITY_TYPE || "";
                if (type) siteTypes.set(type, (siteTypes.get(type) || 0) + 1);
                const fee = row.FEE_YN || row.FEE || row.FEE_STATUS ||
                            row.USER_FEE || row.FEES || "";
                if (fee) feeCounts.set(fee, (feeCounts.get(fee) || 0) + 1);
                const activity = row.ACTIVITIES || row.ACTIVITY || row.REC_ACTIVITY ||
                                 row.PRIMARY_ACTIVITY || row.USE_TYPE || "";
                if (activity) activityTypes.add(activity);
            }
            if (siteNames.size > 0 || siteTypes.size > 0 || feeCounts.size > 0 || activityTypes.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Recreation Site Details</b></td></tr>`;
                if (siteNames.size > 0) {
                    const names = Array.from(siteNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Site Names</td><td>${names}${siteNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (siteTypes.size > 0) {
                    const items = Array.from(siteTypes.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Site Types</td><td>${items}</td></tr>`;
                }
                if (feeCounts.size > 0) {
                    const items = Array.from(feeCounts.entries())
                        .map(([f, count]) => `${escapeHtml(f)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Fee Status</td><td>${items}</td></tr>`;
                }
                if (activityTypes.size > 0) {
                    const activities = Array.from(activityTypes).slice(0, 10).map(a => escapeHtml(a)).join(", ");
                    summaryHtml += `<tr><td>Activities</td><td>${activities}${activityTypes.size > 10 ? " ..." : ""}</td></tr>`;
                }
            }
        }

        // ── BLM LWCF ──
        if (title.includes("lwcf") || title.includes("land and water conservation") || title.includes("conservation fund")) {
            const projectNames = new Set();
            const statusCounts = new Map();
            const purposeCounts = new Map();
            const fiscalYears   = new Set();
            for (const row of rows) {
                const name = row.PROJECT_NAME || row.PROJ_NAME || row.NAME ||
                             row.TRACT_NAME || row.LWCF_NAME || row.UNIT_NAME || "";
                if (name) projectNames.add(name);
                const status = row.STATUS || row.PROJ_STATUS || row.PROJECT_STATUS ||
                               row.LWCF_STATUS || row.ACQ_STATUS || "";
                if (status) statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                const purpose = row.PURPOSE || row.LWCF_PURPOSE || row.USE ||
                                row.PROJECT_TYPE || row.PROJ_TYPE || row.ACQ_TYPE || "";
                if (purpose) purposeCounts.set(purpose, (purposeCounts.get(purpose) || 0) + 1);
                const fy = row.FISCAL_YEAR || row.FY || row.YEAR || row.ACQ_YEAR || "";
                if (fy) fiscalYears.add(fy);
            }
            if (projectNames.size > 0 || statusCounts.size > 0 || purposeCounts.size > 0 || fiscalYears.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>LWCF Project Details</b></td></tr>`;
                if (projectNames.size > 0) {
                    const names = Array.from(projectNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Project/Tract Names</td><td>${names}${projectNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (purposeCounts.size > 0) {
                    const items = Array.from(purposeCounts.entries())
                        .map(([p, count]) => `${escapeHtml(p)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Purpose/Type</td><td>${items}</td></tr>`;
                }
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
                if (fiscalYears.size > 0) {
                    const years = Array.from(fiscalYears).sort().slice(0, 10).map(y => escapeHtml(String(y))).join(", ");
                    summaryHtml += `<tr><td>Fiscal Years</td><td>${years}${fiscalYears.size > 10 ? " ..." : ""}</td></tr>`;
                }
            }
        }

        // ── BLM ePlanning Projects ──
        if (title.includes("eplanning") || title.includes("epl_comment") || title.includes("nepa")) {
            const projectNames = new Set();
            const statusCounts = new Map();
            const typeCounts   = new Map();
            const nepaNumbers  = new Set();
            for (const row of rows) {
                const name = row.PROJECT_NAME || row.PROJ_NAME || row.NAME || "";
                if (name) projectNames.add(name);
                const status = row.NEPA_STATUS || row.PROJECT_STATUS || row.STATUS || "";
                if (status) statusCounts.set(status, (statusCounts.get(status) || 0) + 1);
                const type = row.PROJECT_TYPE || row.NEPA_TYPE || row.TYPE || "";
                if (type) typeCounts.set(type, (typeCounts.get(type) || 0) + 1);
                const nepaNo = row.NEPA_NUMBER || row.NEPA_NO || row.DOI_NUMBER || "";
                if (nepaNo) nepaNumbers.add(nepaNo);
            }
            if (projectNames.size > 0 || statusCounts.size > 0 || typeCounts.size > 0 || nepaNumbers.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>ePlanning Project Details</b></td></tr>`;
                if (projectNames.size > 0) {
                    const names = Array.from(projectNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Project Names</td><td>${names}${projectNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (typeCounts.size > 0) {
                    const items = Array.from(typeCounts.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Project Type</td><td>${items}</td></tr>`;
                }
                if (statusCounts.size > 0) {
                    const items = Array.from(statusCounts.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>NEPA Status</td><td>${items}</td></tr>`;
                }
                if (nepaNumbers.size > 0 && nepaNumbers.size <= 10) {
                    const items = Array.from(nepaNumbers).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>NEPA Numbers</td><td>${items}</td></tr>`;
                } else if (nepaNumbers.size > 10) {
                    summaryHtml += `<tr><td>NEPA Numbers</td><td>${nepaNumbers.size} projects</td></tr>`;
                }
            }
        }

        // ── BLM Fire Perimeters ──
        if (title.includes("fire perimeter")) {
            const fireNames      = new Set();
            const causeCounts    = new Map();
            const discoveryYears = new Map();
            const adminStates    = new Map();
            const totalAcres     = [];
            const complexNames   = new Set();
            for (const row of rows) {
                const name = row.INCDNT_NM || "";
                if (name) fireNames.add(name);
                const cause = row.FIRE_CAUSE_NM || "";
                if (cause) causeCounts.set(cause, (causeCounts.get(cause) || 0) + 1);
                const yr = row.FIRE_DSCVR_CY || "";
                if (yr) discoveryYears.set(String(yr), (discoveryYears.get(String(yr)) || 0) + 1);
                const adminSt = row.ADMIN_ST || "";
                if (adminSt) adminStates.set(adminSt, (adminStates.get(adminSt) || 0) + 1);
                const acres = parseFloat(row.GIS_ACRES || row.TOTAL_RPT_ACRES_NR || 0);
                if (acres > 0) totalAcres.push(acres);
                const cmplx = row.CMPLX_NM || "";
                if (cmplx) complexNames.add(cmplx);
            }
            if (fireNames.size > 0 || causeCounts.size > 0 || discoveryYears.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Fire Perimeter Details</b></td></tr>`;
                if (fireNames.size > 0) {
                    const names = Array.from(fireNames).slice(0, 15).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Fire Names</td><td>${names}${fireNames.size > 15 ? " ..." : ""}</td></tr>`;
                }
                if (causeCounts.size > 0) {
                    const items = Array.from(causeCounts.entries())
                        .map(([c, count]) => `${escapeHtml(c)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Fire Cause</td><td>${items}</td></tr>`;
                }
                if (discoveryYears.size > 0) {
                    const sorted = Array.from(discoveryYears.entries()).sort((a, b) => b[0].localeCompare(a[0]));
                    const items = sorted.slice(0, 15)
                        .map(([y, count]) => `${escapeHtml(y)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Discovery Years</td><td>${items}${sorted.length > 15 ? " ..." : ""}</td></tr>`;
                }
                if (adminStates.size > 0) {
                    const items = Array.from(adminStates.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Admin State</td><td>${items}</td></tr>`;
                }
                if (totalAcres.length > 0) {
                    const sum = totalAcres.reduce((a, b) => a + b, 0);
                    summaryHtml += `<tr><td>Total Burned Acres</td><td>${formatNumber(sum, 1)} (${totalAcres.length} fires)</td></tr>`;
                }
                if (complexNames.size > 0) {
                    const names = Array.from(complexNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Fire Complexes</td><td>${names}${complexNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
            }
        }

        // ── BLM Administrative Units ──
        if (title.includes("administrative unit") || title.includes("admin unit") || title.includes("adminunit")) {
            const unitNames   = new Set();
            const orgTypes    = new Map();
            const adminStates = new Map();
            const parentNames = new Set();
            const stateUrls   = new Set();
            const cityLabels  = new Set();
            for (const row of rows) {
                const name = row.ADMU_NAME || row.Label_Full_Name || row.Label || "";
                if (name) unitNames.add(name);
                const orgType = row.BLM_ORG_TYPE || "";
                if (orgType) orgTypes.set(orgType, (orgTypes.get(orgType) || 0) + 1);
                const adminSt = row.ADMIN_ST || "";
                if (adminSt) adminStates.set(adminSt, (adminStates.get(adminSt) || 0) + 1);
                const parent = row.PARENT_NAME || "";
                if (parent) parentNames.add(parent);
                const stUrl = row.ADMU_ST_URL || "";
                if (stUrl) stateUrls.add(stUrl);
                const city = row.City_Label || "";
                if (city) cityLabels.add(city);
            }
            if (unitNames.size > 0 || orgTypes.size > 0 || adminStates.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>BLM Administrative Unit Details</b></td></tr>`;
                if (unitNames.size > 0) {
                    const names = Array.from(unitNames).slice(0, 15).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Unit Names</td><td>${names}${unitNames.size > 15 ? " ..." : ""}</td></tr>`;
                }
                if (orgTypes.size > 0) {
                    const items = Array.from(orgTypes.entries())
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Organization Type</td><td>${items}</td></tr>`;
                }
                if (adminStates.size > 0) {
                    const items = Array.from(adminStates.entries())
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Admin State</td><td>${items}</td></tr>`;
                }
                if (parentNames.size > 0) {
                    const names = Array.from(parentNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Parent Office</td><td>${names}${parentNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (cityLabels.size > 0) {
                    const cities = Array.from(cityLabels).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Office Location</td><td>${cities}${cityLabels.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (stateUrls.size > 0) {
                    const urls = Array.from(stateUrls).slice(0, 5).map(u => `<a href="${escapeHtml(u)}" target="_blank" rel="noopener">${escapeHtml(u)}</a>`).join("<br/>");
                    summaryHtml += `<tr><td>State Office URL</td><td>${urls}</td></tr>`;
                }
            }
        }

        // ── Generic fallback ──
        if (!summaryHtml && rows.length > 0) {
            const genericNames  = new Set();
            const genericStatus = new Map();
            const genericTypes  = new Map();
            const namePatterns   = /^(.*_)?(NAME|NM|TITLE|LABEL|DESCRIPTION|DESC)(_.*)?$/i;
            const statusPatterns = /^(.*_)?(STATUS|STAT|STATE|CONDITION)(_.*)?$/i;
            const typePatterns   = /^(.*_)?(TYPE|TYP|CLASS|CATEGORY|CAT|KIND)(_.*)?$/i;
            for (const row of rows) {
                for (const [key, val] of Object.entries(row)) {
                    if (val == null || val === "" || typeof val === "number") continue;
                    const strVal = String(val).trim();
                    if (!strVal || strVal.length > 200) continue;
                    if (namePatterns.test(key)) genericNames.add(strVal);
                    else if (statusPatterns.test(key)) genericStatus.set(strVal, (genericStatus.get(strVal) || 0) + 1);
                    else if (typePatterns.test(key)) genericTypes.set(strVal, (genericTypes.get(strVal) || 0) + 1);
                }
            }
            if (genericNames.size > 0 || genericStatus.size > 0 || genericTypes.size > 0) {
                summaryHtml += `<tr><td colspan="2" style="padding-top:12px;"><b>Feature Details</b></td></tr>`;
                if (genericNames.size > 0) {
                    const names = Array.from(genericNames).slice(0, 10).map(n => escapeHtml(n)).join(", ");
                    summaryHtml += `<tr><td>Names</td><td>${names}${genericNames.size > 10 ? " ..." : ""}</td></tr>`;
                }
                if (genericTypes.size > 0) {
                    const items = Array.from(genericTypes.entries()).slice(0, 10)
                        .map(([t, count]) => `${escapeHtml(t)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Types</td><td>${items}</td></tr>`;
                }
                if (genericStatus.size > 0) {
                    const items = Array.from(genericStatus.entries()).slice(0, 10)
                        .map(([s, count]) => `${escapeHtml(s)} (${count})`).join(", ");
                    summaryHtml += `<tr><td>Status</td><td>${items}</td></tr>`;
                }
            }
        }

        return summaryHtml;
    }

    // ────────────────────────────────────────────
    // buildFinalReportHtmlDoc – HTML template
    // ────────────────────────────────────────────
    function buildFinalReportHtmlDoc({ title, createdAt, totalsHtml, aoiSectionHtml, sectionsHtml, dataSourcesHtml }) {
        const safeTitle = escapeHtml(title || "Final Report");

        return `<!doctype html>
            <html lang="en">
            <head>
            <meta charset="utf-8"/>
            <meta name="viewport" content="width=device-width,initial-scale=1"/>
            <title>${safeTitle}</title>
            <link rel="preconnect" href="https://fonts.googleapis.com">
            <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
            <link href="https://fonts.googleapis.com/css2?family=Merriweather:wght@400;700&family=Source+Sans+Pro:wght@400;600;700&display=swap" rel="stylesheet">
            <style>
                :root{
                    --blm-green: #1a472a;
                    --blm-green-light: #2d5a3d;
                    --blm-tan: #f5f0e6;
                    --blm-gold: #c5a43e;
                    --blm-brown: #5c4827;
                    --fg: #2c2c2c;
                    --muted: #5a5a5a;
                    --border: #d4cfc4;
                    --bg: #fdfcfa;
                    --white: #ffffff;
                }
                html,body{ 
                    margin:0; 
                    padding:0; 
                    background: var(--blm-tan); 
                    color:var(--fg); 
                    font-family: 'Source Sans Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    font-size: 14px;
                    line-height: 1.5;
                }
                .report-header{
                    background: linear-gradient(135deg, var(--blm-green) 0%, var(--blm-green-light) 100%);
                    color: var(--white);
                    padding: 28px 32px;
                    margin-bottom: 0;
                }
                .report-header .agency-name{
                    font-family: 'Merriweather', Georgia, serif;
                    font-size: 13px;
                    font-weight: 400;
                    letter-spacing: 1.5px;
                    text-transform: uppercase;
                    opacity: 0.9;
                    margin-bottom: 8px;
                }
                .report-header h1{
                    font-family: 'Merriweather', Georgia, serif;
                    font-size: 28px;
                    font-weight: 700;
                    margin: 0 0 6px 0;
                    letter-spacing: 0.5px;
                }
                .report-header .meta{
                    font-size: 13px;
                    opacity: 0.85;
                    margin: 0;
                }
                .wrap{ 
                    max-width: 900px; 
                    margin: 0 auto; 
                    padding: 32px 40px 60px; 
                    background: var(--bg);
                    box-shadow: 0 0 40px rgba(0,0,0,0.08);
                    min-height: 100vh;
                }
                h2{ 
                    font-family: 'Merriweather', Georgia, serif;
                    font-size: 20px; 
                    font-weight: 700;
                    color: var(--blm-green);
                    margin: 36px 0 18px; 
                    padding-bottom: 10px;
                    border-bottom: 3px solid var(--blm-gold); 
                }
                h3{
                    font-family: 'Source Sans Pro', sans-serif;
                    font-size: 16px;
                    font-weight: 700;
                    color: var(--blm-brown);
                    margin: 24px 0 12px;
                }
                .totals{ margin-top: 24px; }
                .totals .row{ display:flex; gap:14px; flex-wrap:wrap; margin-top:12px; }
                .pill{ 
                    border: 1px solid var(--border); 
                    border-radius: 6px; 
                    padding: 10px 16px; 
                    font-size: 13px; 
                    font-weight: 600;
                    background: var(--white);
                    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
                }
                .aoi-map{ 
                    margin: 18px 0; 
                    border: 2px solid var(--border); 
                    border-radius: 8px; 
                    overflow:hidden; 
                    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
                }
                .aoi-map img{ display:block; width:100%; height:auto; }
                .aoi-details{ margin-top: 20px; }
                .aoi-field{ margin: 10px 0; font-size: 14px; }
                .aoi-label{ 
                    font-weight: 600; 
                    color: var(--blm-green);
                    display: inline-block;
                    min-width: 160px;
                }
                .legal-list{ margin: 6px 0 0 20px; padding: 0; }
                .legal-list li{ margin: 5px 0; }
                table.admin-unit-tbl{
                    width:100%;
                    border-collapse: collapse;
                    margin-top: 12px;
                    font-size: 13px;
                    background: var(--white);
                    border: 1px solid var(--border);
                    border-radius: 6px;
                    overflow: hidden;
                }
                table.admin-unit-tbl th{
                    background: var(--blm-green);
                    color: var(--white);
                    padding: 10px 12px;
                    text-align: left;
                    font-weight: 600;
                    font-size: 12px;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                table.admin-unit-tbl td{
                    padding: 8px 12px;
                    border-bottom: 1px solid var(--border);
                    vertical-align: top;
                }
                table.admin-unit-tbl tr:nth-child(even){
                    background: var(--blm-tan);
                }
                table.admin-unit-tbl tr:last-child td{
                    border-bottom: none;
                }
                .section{ 
                    margin-top: 32px; 
                    padding: 24px;
                    background: var(--white);
                    border: 1px solid var(--border);
                    border-radius: 8px;
                    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
                }
                .section h3{
                    margin-top: 0;
                    padding-bottom: 8px;
                    border-bottom: 1px solid var(--border);
                }
                .section .sub{ 
                    font-size: 12px; 
                    color: var(--muted); 
                    margin-bottom: 14px;
                    font-style: italic;
                }
                .map{ 
                    width:100%; 
                    border: 1px solid var(--border); 
                    border-radius: 6px; 
                    overflow:hidden; 
                    background: var(--white); 
                    margin: 16px 0;
                    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
                }
                .map img{ display:block; width:100%; height:auto; }
                table.metaTbl{ 
                    width:100%; 
                    border-collapse: collapse; 
                    margin-top: 16px; 
                    font-size: 13px;
                    background: var(--blm-tan);
                    border-radius: 6px;
                    overflow: hidden;
                }
                table.metaTbl td{ 
                    padding: 10px 14px; 
                    border-bottom: 1px solid var(--border); 
                }
                table.metaTbl tr:last-child td{
                    border-bottom: none;
                }
                table.metaTbl td:first-child{ 
                    color: var(--blm-green); 
                    font-weight: 600;
                    width: 200px;
                    background: rgba(26,71,42,0.05);
                }
                table.data-sources-table{ 
                    width:100%; 
                    border-collapse: collapse; 
                    margin-top: 16px; 
                    font-size: 12px; 
                    table-layout: fixed;
                    background: var(--white);
                    border: 1px solid var(--border);
                    border-radius: 6px;
                    overflow: hidden;
                }
                table.data-sources-table th{ 
                    background: var(--blm-green); 
                    color: var(--white);
                    padding: 12px 14px; 
                    text-align: left; 
                    font-weight: 600;
                    font-size: 12px;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                table.data-sources-table td{ 
                    padding: 10px 14px; 
                    border-bottom: 1px solid var(--border); 
                    vertical-align: top; 
                    word-wrap: break-word; 
                }
                table.data-sources-table tr:nth-child(even){
                    background: var(--blm-tan);
                }
                table.data-sources-table tr:last-child td{
                    border-bottom: none;
                }
                table.data-sources-table .service-desc-col{ white-space: normal; line-height: 1.5; }
                table.data-sources-table .service-url-col{ 
                    font-family: 'Consolas', 'Monaco', monospace; 
                    font-size: 10px; 
                    word-break: break-all;
                    color: var(--muted);
                }    
                .mono{ font-family: 'Consolas', 'Monaco', monospace; font-size: 11px; }
                .status-up{ color: #2e7d32; font-weight: 600; }
                .status-down{ color: #c62828; font-weight: 600; }
                .actions{ 
                    margin-top: 20px; 
                    display:flex; 
                    gap:12px; 
                    flex-wrap: wrap; 
                }
                .btn{
                    display:inline-block; 
                    background: var(--blm-green); 
                    color: var(--white);
                    border: none;
                    border-radius: 6px;
                    padding: 12px 20px; 
                    font-size: 14px; 
                    font-weight: 600;
                    text-decoration:none;
                    cursor: pointer;
                    transition: background 0.2s ease;
                }
                .btn:hover{ background: var(--blm-green-light); }
                .hint{ 
                    font-size: 12px; 
                    color: var(--muted); 
                    margin-top: 10px;
                    font-style: italic;
                }
                .report-footer{
                    margin-top: 48px;
                    padding-top: 24px;
                    border-top: 2px solid var(--border);
                    font-size: 11px;
                    color: var(--muted);
                    text-align: center;
                }
                .report-footer .dept-name{
                    font-weight: 600;
                    color: var(--blm-green);
                    margin-bottom: 4px;
                }
                @media print{
                    html, body{ background: white; }
                    .actions, .hint{ display:none !important; }
                    .wrap{ 
                        max-width: none; 
                        padding: 0; 
                        box-shadow: none;
                        background: white;
                    }
                    .report-header{
                        background: var(--blm-green) !important;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                    .section{ 
                        break-inside: avoid; 
                        box-shadow: none;
                        border: 1px solid #ccc;
                    }
                    .pagebreak{ break-after: page; }
                    table.data-sources-table th{
                        background: var(--blm-green) !important;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                }
            </style>
            </head>
            <body>
            <div class="report-header">
                <div class="agency-name">U.S. Department of the Interior &bull; Bureau of Land Management</div>
                <h1>${safeTitle}</h1>
                <div class="meta">Report Generated: ${escapeHtml(createdAt || "")}</div>
            </div>
            <div class="wrap">
                <div class="actions">
                    <a class="btn" href="javascript:window.print()">&#128424; Print / Save as PDF</a>
                </div>
                <div class="hint">Use your browser's print dialog and select &ldquo;Save as PDF&rdquo; to create a permanent copy of this report.</div>

                <h2>Report Summary</h2>
                <div class="totals">
                ${totalsHtml || ""}
                </div>

                ${aoiSectionHtml || ""}

                <h2>Layer Analysis Maps</h2>
                ${sectionsHtml || ""}

                ${dataSourcesHtml || ""}
                
                <div class="report-footer">
                    <div class="dept-name">Bureau of Land Management</div>
                    <div>U.S. Department of the Interior</div>
                    <div style="margin-top:8px;">This report was generated using geospatial data from BLM and partner agency web services.</div>
                </div>
            </div>
            </body>
            </html>`;
    }

    // ────────────────────────────────────────────
    // State-dependent helpers
    // ────────────────────────────────────────────

    function getAoiSummaryForReport(aoiAcres) {
        const src = S.aoiSource === "draw" ? "Drawn AOI" : "Selected AOI";
        const tool = S.aoiSource === "select" ? plssToolLabel(S.aoiSourcePlssTool) : "";
        const srcDetail = (S.aoiSource === "select" && tool) ? ` (${tool})` : "";
        const layer = S.aoiSourceLayerTitle ? ` \u2022 Source layer: ${S.aoiSourceLayerTitle}` : "";
        return `${src}${srcDetail} \u2022 AOI area: ${formatNumber(aoiAcres, 2)} acres${layer}`;
    }

    function buildDataSourcesSection() {
        const services = getConfiguredServices(S.config);

        const rows = services.map(svc => {
            const status = S.serviceStatus.get(svc.url) || "UNKNOWN";
            const statusClass = status === "UP" ? "status-up" : "status-down";
            const desc = S.serviceStatus.get(svc.url + "::desc") || "(Description not available)";

            return `
                <tr>
                    <td class="service-name-col">${escapeHtml(svc.title)}</td>
                    <td class="service-url-col"><a href="${escapeHtml(svc.url)}" target="_blank" rel="noopener">${escapeHtml(svc.url)}</a></td>
                    <td class="service-desc-col">${escapeHtml(desc)}</td>
                    <td class="service-status-col"><span class="${statusClass}">${status}</span></td>
                </tr>
            `;
        }).join("");

        return `
            <div class="section" style="background: transparent; border: none; box-shadow: none; padding: 0;">
                <h2>Data Sources</h2>
                <p style="font-size: 13px; color: var(--muted); margin-bottom: 16px;">
                    The following geospatial web services were used to generate this report. Service availability was verified at the time of report generation.
                </p>
                <table class="data-sources-table">
                    <thead>
                        <tr>
                            <th style="width: 20%;">Service Name</th>
                            <th style="width: 25%;">Service URL</th>
                            <th style="width: 45%;">Description</th>
                            <th style="width: 10%;">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rows}
                    </tbody>
                </table>
            </div>
        `;
    }

    // ────────────────────────────────────────────
    // generateAoiMapsWithCircles
    // ────────────────────────────────────────────
    async function generateAoiMapsWithCircles() {
        const view = S.view;
        const selectionGeom = S.selectionGeom;
        if (!view || !selectionGeom) return "";

        const config = S.config;
        const aoiLayer = S.aoiLayer;
        const aoiMaskLayer = S.aoiMaskLayer;
        const alwaysVisibleLayers = S.alwaysVisibleLayers;

        const { ensureAoiOnTop, hideAoiMask, captureScreenshotWithWait } = mapUtils;

        const width  = config?.visualReport?.screenshotWidth ?? 1400;
        const height = Math.round(width * 0.5625);
        const maps   = [];

        const insetFrac = 0.22;
        const insetW = Math.round(width * insetFrac);
        const insetH = Math.round(height * insetFrac);
        const insetMargin = 12;
        const overviewZoomFactor = 8;

        const allLayers   = view.map.layers.toArray();
        const visSnapshot = allLayers.map(l => ({ layer: l, visible: l.visible }));

        let plssTownshipLayer = null;
        for (const l of allLayers) {
            if (l.title && l.title.toLowerCase().includes("township")) {
                plssTownshipLayer = l;
                break;
            }
        }

        function setVisibilityForAoi() {
            for (const l of allLayers) {
                if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }
                if (aoiMaskLayer && l === aoiMaskLayer) { l.visible = false; continue; }
                if (l?.type === "tile") { l.visible = true; continue; }
                if (plssTownshipLayer && l === plssTownshipLayer) { l.visible = true; continue; }
                if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }
                l.visible = false;
            }
            ensureAoiOnTop();
        }

        function restoreVisibility() {
            visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
            hideAoiMask();
            ensureAoiOnTop();
        }

        async function compositeWithOverview(mainDataUrl, mainExtent, scale) {
            const overviewScale = scale * overviewZoomFactor;
            await view.goTo({ target: selectionGeom.extent, scale: overviewScale }, { animate: false });
            const ovSs = await captureScreenshotWithWait({ width });
            if (!ovSs) return mainDataUrl;

            const ovExtent = view.extent;
            const [mainImg, ovImg] = await Promise.all([
                loadImageFromDataUrl(mainDataUrl),
                loadImageFromDataUrl(ovSs)
            ]);
            if (!mainImg || !ovImg) return mainDataUrl;

            const canvas = document.createElement("canvas");
            canvas.width  = width;
            canvas.height = height;
            const ctx = canvas.getContext("2d");

            ctx.drawImage(mainImg, 0, 0, width, height);

            const ix = insetMargin;
            const iy = height - insetH - insetMargin;

            ctx.save();
            ctx.shadowColor = "rgba(0,0,0,0.35)";
            ctx.shadowBlur = 8;
            ctx.shadowOffsetX = 2;
            ctx.shadowOffsetY = 2;
            ctx.fillStyle = "#fff";
            ctx.fillRect(ix - 2, iy - 2, insetW + 4, insetH + 4);
            ctx.restore();

            ctx.drawImage(ovImg, ix, iy, insetW, insetH);

            if (ovExtent && mainExtent) {
                const ovXMin = ovExtent.xmin;
                const ovYMin = ovExtent.ymin;
                const ovW = ovExtent.xmax - ovExtent.xmin;
                const ovH = ovExtent.ymax - ovExtent.ymin;

                if (ovW > 0 && ovH > 0) {
                    const rxPx = ix + ((mainExtent.xmin - ovXMin) / ovW) * insetW;
                    const ryPx = iy + ((ovExtent.ymax - mainExtent.ymax) / ovH) * insetH;
                    const rwPx = ((mainExtent.xmax - mainExtent.xmin) / ovW) * insetW;
                    const rhPx = ((mainExtent.ymax - mainExtent.ymin) / ovH) * insetH;

                    ctx.strokeStyle = "#e63946";
                    ctx.lineWidth = 2.5;
                    ctx.strokeRect(rxPx, ryPx, rwPx, rhPx);
                }
            }

            ctx.strokeStyle = "#333";
            ctx.lineWidth = 1.5;
            ctx.strokeRect(ix - 1, iy - 1, insetW + 2, insetH + 2);

            return canvas.toDataURL("image/png");
        }

        function loadImageFromDataUrl(dataUrl) {
            return new Promise(resolve => {
                const img = new Image();
                img.onload  = () => resolve(img);
                img.onerror = () => resolve(null);
                img.src = dataUrl;
            });
        }

        try {
            setVisibilityForAoi();

            const ext1 = selectionGeom.extent;
            await view.goTo({ target: ext1, scale: 900000 }, { animate: false });
            const ss1 = await captureScreenshotWithWait({ width });
            const mainExtent1 = view.extent.clone();

            if (ss1) {
                const composited1 = await compositeWithOverview(ss1, mainExtent1, 900000);
                maps.push(`<div class="aoi-map"><img src="${composited1}" alt="AOI Context (Regional 1:900,000)" /></div>`);
            }

            const ext2 = selectionGeom.extent;
            await view.goTo({ target: ext2, scale: 200000 }, { animate: false });
            const ss2 = await captureScreenshotWithWait({ width });
            const mainExtent2 = view.extent.clone();

            if (ss2) {
                const composited2 = await compositeWithOverview(ss2, mainExtent2, 200000);
                maps.push(`<div class="aoi-map"><img src="${composited2}" alt="AOI Context (County 1:200,000)" /></div>`);
            }

        } finally {
            restoreVisibility();
        }

        return maps.join("");
    }

    // ────────────────────────────────────────────
    // buildFinalReportHtml – main orchestrator
    // ────────────────────────────────────────────
    async function buildFinalReportHtml() {
        const view = S.view;
        const selectionGeom = S.selectionGeom;
        const config = S.config;
        const lastReportRowsByLayer = S.lastReportRowsByLayer;
        const aoiLayer = S.aoiLayer;
        const aoiMaskLayer = S.aoiMaskLayer;
        const alwaysVisibleLayers = S.alwaysVisibleLayers;
        const layerCfgByUrl = S.layerCfgByUrl;

        if (!view) return;

        if (!selectionGeom) {
            _setStatus("Select or draw an AOI first.");
            return;
        }

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            _setStatus("Run the report first (Tables tab) so we know which layers intersect.");
            return;
        }

        if (finalReportStatus) finalReportStatus.textContent = "Building report...";

        const {
            updateAoiMask, hideAoiMask, ensureAoiOnTop,
            waitForViewStationary, waitForLayerReadyToCapture,
            captureScreenshotWithWait,
            getLayerGeometryType, makeRendererOpaque, getPresetRenderer
        } = mapUtils;

        const {
            queryAllFeaturesPaged, computeElevationStats,
            computeLayerCoverageStats, buildPerFeatureTable, SQM_PER_ACRE
        } = queryEngine;

        try {
            // STEP 1: Compute AOI area
            let aoiAcres = 0;
            try {
                const aoiSqm = Math.max(0, geometryEngine.geodesicArea(selectionGeom, "square-meters"));
                aoiAcres = aoiSqm / SQM_PER_ACRE;
            } catch (e) {
                aoiAcres = 0;
            }

            // STEP 2: Generate AOI Section Data
            // 2a. Primary State
            let primaryState = "";
            let additionalStates = "";

            const stateItem = lastReportRowsByLayer.find(x =>
                x.title && x.title.toLowerCase().includes("state boundaries")
            );

            if (stateItem && !stateItem.fullRows && stateItem._layer && stateItem._exportQuery) {
                try {
                    const pageSize = config.report?.pageSize ?? 1000;
                    const fullFeatures = await queryAllFeaturesPaged(
                        stateItem._layer, stateItem._exportQuery, pageSize, 100
                    );
                    stateItem.fullRows = flattenAttributes(fullFeatures);
                } catch (e) {
                    console.warn("Failed to fetch state data:", e);
                }
            }

            if (stateItem && stateItem.fullRows && stateItem.fullRows.length > 0) {
                const stateNames = stateItem.fullRows
                    .map(r => r.NAME || r.STATE_NAME || r.STATE || r.STUSPS || "")
                    .filter(Boolean);
                if (stateNames.length > 0) {
                    primaryState = stateNames[0];
                    if (stateNames.length > 1) additionalStates = stateNames.slice(1).join(", ");
                }
            }

            // 2b. Legal Land Description
            const legalDescriptions = [];
            const parcelItem = lastReportRowsByLayer.find(x =>
                x.title && (x.title.toLowerCase().includes("parcel") || x.title.toLowerCase().includes("intersected"))
            );

            if (parcelItem && !parcelItem.fullRows && parcelItem._layer && parcelItem._exportQuery) {
                try {
                    const pageSize = config.report?.pageSize ?? 1000;
                    const fullFeatures = await queryAllFeaturesPaged(
                        parcelItem._layer, parcelItem._exportQuery, pageSize, 100
                    );
                    parcelItem.fullRows = flattenAttributes(fullFeatures);
                } catch (e) {
                    console.warn("Failed to fetch parcel data:", e);
                }
            }

            if (parcelItem && parcelItem.fullRows && parcelItem.fullRows.length > 0) {
                for (const row of parcelItem.fullRows) {
                    const legalDesc = formatLegalDescription(row);
                    if (legalDesc) {
                        legalDescriptions.push(legalDesc);
                    } else {
                        console.warn("Empty legal description for parcel row:", row);
                    }
                }
            }

            // 2c. AOI Method
            let aoiMethod = "Manually Drawn";
            if (S.aoiSource === "select") {
                const tool = plssToolLabel(S.aoiSourcePlssTool);
                aoiMethod = `Selected ${tool}`;
            }

            // 2d. AOI Maps
            _setStatus("building final report\u2026 (generating AOI maps)");
            const aoiMapsHtml = await generateAoiMapsWithCircles();

            // STEP 3: Generate per-layer map sections
            _setStatus("building final report\u2026 (generating layer maps)");

            const paddingFactor = config?.visualReport?.paddingFactor ?? 1.25;
            const width = config?.visualReport?.screenshotWidth ?? 1400;

            let fixedExtent = null;
            const ext = selectionGeom?.extent;
            if (ext && ext.expand) fixedExtent = ext.expand(paddingFactor);

            const targets = lastReportRowsByLayer
                .filter(x => (x?.count || 0) > 0)
                .filter(x => (x?._layer && x?._exportQuery) || x?.__isImageService)
                .filter(x => !(x.title && x.title.toLowerCase().includes("state boundaries")))
                .filter(x => !(x.title && x.title.toLowerCase().includes("administrative unit")));

            let sectionsHtml = "";

            if (!targets.length) {
                sectionsHtml = `
                  <div class="section">
                    <h3>No Intersecting Layers Found</h3>
                    <p style="color: var(--muted); font-style: italic;">The analysis found no layers with features intersecting the selected Area of Interest.</p>
                  </div>
                `;
            } else {
                const allLayers = view.map.layers.toArray();
                const visSnapshot = allLayers.map(l => ({ layer: l, visible: l.visible }));
                const originalBasemap = view.map.basemap;
                const imageryBasemapId = config?.map?.imageryBasemap || "satellite";

                let plssSectionLayer = null;
                for (const l of allLayers) {
                    if (l.title && l.title.toLowerCase().includes("section")) {
                        plssSectionLayer = l;
                        break;
                    }
                }

                function setVisibilityForScreenshot(tempLayer) {
                    for (const l of allLayers) {
                        if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }
                        if (aoiMaskLayer && l === aoiMaskLayer) { l.visible = true; continue; }
                        if (l?.type === "tile") { l.visible = true; continue; }
                        if (plssSectionLayer && l === plssSectionLayer) { l.visible = true; continue; }
                        if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }
                        l.visible = false;
                    }
                    if (tempLayer) tempLayer.visible = true;
                    updateAoiMask(true);
                    ensureAoiOnTop();
                }

                function restoreVisibility() {
                    visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
                    hideAoiMask();
                    ensureAoiOnTop();
                }

                const consistentExtent = selectionGeom.extent;

                try {
                    view.map.basemap = imageryBasemapId;
                    await new Promise(r => setTimeout(r, 2000));
                    await waitForViewStationary(3500);
                } catch (e) {
                    console.warn("Failed to switch to imagery basemap:", e);
                }

                for (let i = 0; i < targets.length; i++) {
                    const item = targets[i];
                    _setStatus(`building final report\u2026 (${i + 1}/${targets.length})`);

                    // ImageServer layers
                    if (item.__isImageService) {
                        const imgLayerOpts = {
                            url: item.url,
                            title: item.title,
                            visible: true
                        };
                        if (item.__renderingRule) {
                            imgLayerOpts.renderingRule = { functionName: item.__renderingRule };
                        }
                        const temp = new ImageryLayer(imgLayerOpts);
                        view.map.add(temp);

                        try {
                            setVisibilityForScreenshot(temp);
                            await waitForLayerReadyToCapture(temp, view, { timeoutMs: 10000 });
                            await view.goTo(consistentExtent, { animate: false });
                            await view.goTo({ center: view.center, scale: view.scale * 0.5 }, { animate: false });
                            await waitForViewStationary(2500);

                            const dataUrl = await captureScreenshotWithWait({ width });
                            if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

                            const meta = item.__serviceMeta || {};
                            const elevStats = await computeElevationStats(item.url, selectionGeom);

                            let elevStatsHtml = '';
                            if (elevStats) {
                                elevStatsHtml = `
                                  <tr><td colspan="2" style="background:#f0f0f0; font-weight:600;">Elevation Statistics (within AOI)</td></tr>
                                  <tr><td>Minimum Elevation</td><td><b>${formatNumber(elevStats.minFt, 0)}</b> ft (${formatNumber(elevStats.min, 1)} m)</td></tr>
                                  <tr><td>Maximum Elevation</td><td><b>${formatNumber(elevStats.maxFt, 0)}</b> ft (${formatNumber(elevStats.max, 1)} m)</td></tr>
                                  <tr><td>Elevation Change</td><td><b>${formatNumber(elevStats.elevationChangeFt, 0)}</b> ft (${formatNumber(elevStats.elevationChange, 1)} m)</td></tr>
                                  ${elevStats.meanFt ? `<tr><td>Mean Elevation</td><td><b>${formatNumber(elevStats.meanFt, 0)}</b> ft (${formatNumber(elevStats.mean, 1)} m)</td></tr>` : ''}
                                `;
                            }

                            sectionsHtml += `
                              <div class="section layer-section page-break">
                                <h3>${escapeHtml(item.title)}</h3>
                                <div class="layer-map-container">
                                  <img src="${dataUrl}" alt="${escapeHtml(item.title)}" style="width:100%; border-radius:8px;" />
                                </div>
                                <table class="info-table" style="margin-top:16px;">
                                  <tr><td style="width:200px;">Service Name</td><td><b>${escapeHtml(meta.name || item.title)}</b></td></tr>
                                  <tr><td>Type</td><td>Image Service (Elevation/Raster)</td></tr>
                                  ${elevStatsHtml}
                                  ${meta.copyright ? `<tr><td>Source</td><td>${escapeHtml(meta.copyright)}</td></tr>` : ''}
                                </table>
                              </div>
                            `;
                        } finally {
                            try { view.map.remove(temp); } catch (e) { }
                            restoreVisibility();
                        }
                        continue;
                    }

                    // FeatureServer layers
                    const tempGeomType = await getLayerGeometryType(item.url);
                    const itemCfg = layerCfgByUrl.get(item.url)?.cfg || null;
                    const useNativeRenderer = itemCfg?.useServiceRenderer === true;
                    const opaqueRenderer = useNativeRenderer
                        ? undefined
                        : makeRendererOpaque(getPresetRenderer("report", itemCfg, tempGeomType));
                    const tempOpts = {
                        url: item.url,
                        title: item.title,
                        outFields: ["*"],
                        visible: true
                    };
                    if (!useNativeRenderer) tempOpts.renderer = opaqueRenderer || undefined;
                    const temp = new FeatureLayer(tempOpts);

                    if (!useNativeRenderer) {
                        temp.minScale = 0;
                        temp.maxScale = 0;
                    }

                    view.map.add(temp);
                    try {
                        setVisibilityForScreenshot(temp);
                        await waitForLayerReadyToCapture(temp, view, { timeoutMs: 8000 });
                        await view.goTo(consistentExtent, { animate: false });
                        await view.goTo({ center: view.center, scale: view.scale * 0.5 }, { animate: false });
                        await waitForViewStationary(2000);

                        const dataUrl = await captureScreenshotWithWait({ width });
                        if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

                        const cov = await computeLayerCoverageStats(item, selectionGeom);
                        const acresCovered = cov ? cov.acresCovered : 0;
                        const pctCovered   = cov ? cov.pctAoiCovered : 0;

                        const layerAttrSummary = generateLayerAttributeSummary(item);
                        const perFeatureTableHtml = (item.count > 1)
                            ? await buildPerFeatureTable(item, selectionGeom)
                            : "";

                        const isSingleFeatureLowCoverage = (item.count === 1 && pctCovered < 3);
                        const lowCoverageWarningHtml = isSingleFeatureLowCoverage
                            ? `<div style="margin-top:12px; padding:10px; background-color:#fff3cd; border:1px solid #ffc107; border-radius:4px;">
                                <span style="color:#856404;">\u26A0\uFE0F <b>Low Coverage Warning:</b> This feature covers less than 3% of the AOI. This may indicate a polygon sliver or boundary artifact rather than meaningful overlap.</span>
                               </div>`
                            : "";

                        sectionsHtml += `
                        <div class="section">
                            <h3>${escapeHtml(item.title)}</h3>
                            <div class="map">
                                <img src="${dataUrl}" alt="AOI + ${escapeHtml(item.title)}"/>
                            </div>
                            <table class="metaTbl">
                                <tr><td>AOI Area</td><td><b>${formatNumber(aoiAcres, 2)}</b> acres</td></tr>
                                <tr><td>Intersecting Features</td><td><b>${escapeHtml(String(item.count || 0))}</b></td></tr>
                                <tr><td>Layer Coverage</td><td><b>${formatNumber(acresCovered, 2)}</b> acres</td></tr>
                                <tr><td>Percent of AOI Covered</td><td><b>${formatNumber(pctCovered, 2)}%</b>${isSingleFeatureLowCoverage ? ' <span style="color:#856404;" title="Low coverage \u2014 possible sliver or boundary artifact">\u26A0\uFE0F</span>' : ''}</td></tr>
                                ${layerAttrSummary}
                            </table>
                            ${lowCoverageWarningHtml}
                            ${perFeatureTableHtml}
                        </div>
                        <div class="pagebreak"></div>
                        `;
                    } finally {
                        try { view.map.remove(temp); } catch (e) { }
                        restoreVisibility();
                    }
                }

                // Restore original basemap
                try {
                    view.map.basemap = originalBasemap;
                    await new Promise(r => setTimeout(r, 1000));
                } catch (e) {
                    console.warn("Failed to restore original basemap:", e);
                }
            }

            // STEP 4: Data Sources Appendix
            const dataSourcesHtml = buildDataSourcesSection();

            // STEP 5: Build Final HTML Document
            const totalLayers    = lastReportRowsByLayer.length;
            const layersWithHits = lastReportRowsByLayer.filter(x => (x.count || 0) > 0).length;
            const totalHits      = lastReportRowsByLayer.reduce((sum, x) => sum + (x.count || 0), 0);

            const totalsHtml = `
              <div class="row">
                <div class="pill">Layers queried: <b>${escapeHtml(String(totalLayers))}</b></div>
                <div class="pill">Layers with hits: <b>${escapeHtml(String(layersWithHits))}</b></div>
                <div class="pill">Total intersecting features: <b>${escapeHtml(String(totalHits))}</b></div>
              </div>
            `;

            const stateHtml = primaryState
                ? `<div class="aoi-field"><span class="aoi-label">Primary State:</span> ${escapeHtml(primaryState)}${additionalStates ? ` (Additional: ${escapeHtml(additionalStates)})` : ""}</div>`
                : "";

            const legalHtml = legalDescriptions.length > 0
                ? `<div class="aoi-field">
                    <span class="aoi-label">Legal Land Description:</span>
                    <ul class="legal-list">
                        ${legalDescriptions.map(ld => `<li>${escapeHtml(ld)}</li>`).join("")}
                    </ul>
                   </div>`
                : "";

            // Admin Unit table for AOI section
            let adminUnitTableHtml = "";
            if (lastReportRowsByLayer && lastReportRowsByLayer.length) {
                const adminItems = lastReportRowsByLayer.filter(x =>
                    x.title && x.title.toLowerCase().includes("administrative unit") && (x.count || 0) > 0
                );
                if (adminItems.length > 0) {
                    const allAdminRows = [];
                    for (const item of adminItems) {
                        for (const row of (item.rows || [])) {
                            allAdminRows.push({
                                name: row.ADMU_NAME || row.Label_Full_Name || row.Label || "",
                                orgType: row.BLM_ORG_TYPE || "",
                                adminSt: row.ADMIN_ST || "",
                                parent: row.PARENT_NAME || "",
                                city: row.City_Label || "",
                                url: row.ADMU_ST_URL || ""
                            });
                        }
                    }
                    const seen = new Set();
                    const unique = allAdminRows.filter(r => {
                        const key = `${r.name}|${r.orgType}`;
                        if (!r.name || seen.has(key)) return false;
                        seen.add(key);
                        return true;
                    });
                    if (unique.length > 0) {
                        const trs = unique.map(r => `
                            <tr>
                                <td>${escapeHtml(r.name)}</td>
                                <td>${escapeHtml(r.orgType)}</td>
                                <td>${escapeHtml(r.adminSt)}</td>
                                <td>${escapeHtml(r.parent)}</td>
                                <td>${escapeHtml(r.city)}</td>
                                ${r.url ? `<td><a href="${escapeHtml(r.url)}" target="_blank" rel="noopener">Link</a></td>` : `<td></td>`}
                            </tr>`).join("");
                        adminUnitTableHtml = `
                            <div style="margin-top:20px;">
                                <span class="aoi-label" style="display:block; margin-bottom:8px;">BLM Administrative Units Intersecting AOI:</span>
                                <table class="admin-unit-tbl">
                                    <thead>
                                        <tr>
                                            <th>Unit Name</th>
                                            <th>Type</th>
                                            <th>State</th>
                                            <th>Parent Office</th>
                                            <th>Office Location</th>
                                            <th>URL</th>
                                        </tr>
                                    </thead>
                                    <tbody>${trs}</tbody>
                                </table>
                            </div>`;
                    }
                }
            }

            const aoiSectionHtml = `
                <h2>Area of Interest</h2>
                ${aoiMapsHtml}
                <div class="aoi-details">
                    ${stateHtml}
                    ${legalHtml}
                    <div class="aoi-field"><span class="aoi-label">Area:</span> ${formatNumber(aoiAcres, 2)} acres</div>
                    <div class="aoi-field"><span class="aoi-label">Method:</span> ${escapeHtml(aoiMethod)}</div>
                    ${adminUnitTableHtml}
                </div>
            `;

            const htmlDoc = buildFinalReportHtmlDoc({
                title: "Land & Resource Intersection Analysis Report",
                createdAt: formatDateTimeForReport(new Date()),
                totalsHtml,
                aoiSectionHtml,
                sectionsHtml,
                dataSourcesHtml
            });

            cachedFinalReportHtml = htmlDoc;

            if (finalReportStatus) finalReportStatus.textContent = "Report ready.";

        } catch (e) {
            console.error(e);
            if (finalReportStatus) finalReportStatus.textContent = "Failed to build report (see console).";
        }
    }

    function viewFinalReport() {
        if (!cachedFinalReportHtml) {
            alert("Run analysis first to generate the report.");
            return;
        }
        openHtmlInNewTab(cachedFinalReportHtml);
    }

    // ────────────────────────────────────────────
    // init – called once from app.js
    // ────────────────────────────────────────────
    function init(state, deps) {
        S              = state;
        mapUtils       = deps.mapUtils;
        queryEngine    = deps.queryEngine;
        ImageryLayer   = deps.ImageryLayer;
        FeatureLayer   = deps.FeatureLayer;
        geometryEngine = deps.geometryEngine;
        _setStatus     = deps.setStatus || _setStatus;
        finalReportStatus = deps.finalReportStatus || null;

        return {
            // Pure helpers
            formatLegalDescription,
            generateLayerAttributeSummary,
            openHtmlInNewTab,
            formatDateTimeForReport,
            buildFinalReportHtmlDoc,
            // State-dependent
            getAoiSummaryForReport,
            buildDataSourcesSection,
            generateAoiMapsWithCircles,
            buildFinalReportHtml,
            viewFinalReport,
            // Accessors for cachedFinalReportHtml
            getCachedFinalReportHtml: () => cachedFinalReportHtml,
            setCachedFinalReportHtml: (v) => { cachedFinalReportHtml = v; }
        };
    }

    return { init };
});
