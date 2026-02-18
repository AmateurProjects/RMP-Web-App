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
        getConfiguredServices, flattenAttributes,
        fetchJsonWithTimeout, normalizePjsonUrl, pickServiceDescription
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
    let _lastReportId = null;
    // Reference to DOM element
    let finalReportStatus = null;
    // External helpers injected from app.js
    let _setStatus = () => {};

    // ────────────────────────────────────────────
    // IndexedDB report storage (persist reports for sharing via URL)
    // ────────────────────────────────────────────
    const REPORT_DB_NAME = "RmpReports";
    const REPORT_STORE   = "reports";
    const REPORT_STATE_STORE = "reportState";
    const REPORT_DB_VER  = 2;
    const REPORT_TTL_DAYS = 7;

    function _openReportDb() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(REPORT_DB_NAME, REPORT_DB_VER);
            req.onupgradeneeded = () => {
                const db = req.result;
                if (!db.objectStoreNames.contains(REPORT_STORE)) {
                    const store = db.createObjectStore(REPORT_STORE, { keyPath: "id" });
                    store.createIndex("expiresAt", "expiresAt", { unique: false });
                }
                if (!db.objectStoreNames.contains(REPORT_STATE_STORE)) {
                    db.createObjectStore(REPORT_STATE_STORE, { keyPath: "reportId" });
                }
            };
            req.onsuccess = () => resolve(req.result);
            req.onerror   = () => reject(req.error);
        });
    }

    function _generateReportId() {
        // Short URL-friendly ID (8 chars)
        const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        let id = "";
        const arr = crypto.getRandomValues(new Uint8Array(8));
        for (let i = 0; i < 8; i++) id += chars[arr[i] % chars.length];
        return id;
    }

    async function saveReportToDb(html, existingId) {
        const db = await _openReportDb();
        const id = existingId || _generateReportId();
        const record = {
            id,
            html,
            createdAt: Date.now(),
            expiresAt: Date.now() + REPORT_TTL_DAYS * 86400000
        };
        return new Promise((resolve, reject) => {
            const tx = db.transaction(REPORT_STORE, "readwrite");
            tx.objectStore(REPORT_STORE).put(record);
            tx.oncomplete = () => { db.close(); resolve(id); };
            tx.onerror    = () => { db.close(); reject(tx.error); };
        });
    }

    async function loadReportFromDb(id) {
        const db = await _openReportDb();
        return new Promise((resolve, reject) => {
            const tx  = db.transaction(REPORT_STORE, "readonly");
            const req = tx.objectStore(REPORT_STORE).get(id);
            req.onsuccess = () => {
                db.close();
                const rec = req.result;
                if (!rec) return resolve(null);
                if (rec.expiresAt < Date.now()) return resolve(null); // expired
                resolve(rec.html);
            };
            req.onerror = () => { db.close(); reject(req.error); };
        });
    }

    async function cleanupExpiredReports() {
        try {
            const db = await _openReportDb();
            const tx = db.transaction(REPORT_STORE, "readwrite");
            const store = tx.objectStore(REPORT_STORE);
            const idx = store.index("expiresAt");
            const range = IDBKeyRange.upperBound(Date.now());
            const req = idx.openCursor(range);
            req.onsuccess = () => {
                const cursor = req.result;
                if (cursor) { cursor.delete(); cursor.continue(); }
            };
            tx.oncomplete = () => db.close();
            tx.onerror    = () => db.close();
        } catch (e) {
            console.warn("Report cleanup failed:", e);
        }
    }

    function getReportShareUrl(reportId) {
        const base = window.location.origin + window.location.pathname;
        return base + "?report=" + reportId;
    }

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
    // generateFindingsSummary – human-readable paragraph
    // ────────────────────────────────────────────
    function generateFindingsSummary(reportItems, aoiAcres) {
        if (!reportItems || !reportItems.length) return "";

        const totalLayers = reportItems.length;
        const layersWithHits = reportItems.filter(function (x) { return (x.count || 0) > 0; });
        const totalHits = reportItems.reduce(function (s, x) { return s + (x.count || 0); }, 0);

        // Categorize findings
        var specialDesignations = [];
        var environmentalConcerns = [];
        var existingAuthorizations = [];
        var landUsePlans = [];
        var landStatus = [];

        for (var idx = 0; idx < layersWithHits.length; idx++) {
            var item = layersWithHits[idx];
            var title = (item.title || "").toLowerCase();
            var count = item.count || 0;
            var entry = { name: item.title, count: count };

            if (title.includes("acec") || title.includes("critical environmental concern")) {
                specialDesignations.push(entry);
            } else if (title.includes("wilderness")) {
                specialDesignations.push(entry);
            } else if (title.includes("nlcs") || title.includes("conservation area") || title.includes("national monument")) {
                specialDesignations.push(entry);
            } else if (title.includes("visual resource")) {
                specialDesignations.push(entry);
            } else if (title.includes("recreation site") || title.includes("lwcf") || title.includes("conservation fund")) {
                specialDesignations.push(entry);
            } else if (title.includes("critical habitat")) {
                environmentalConcerns.push(entry);
            } else if (title.includes("migration") || title.includes("ungulate")) {
                environmentalConcerns.push(entry);
            } else if (title.includes("wild horse") || title.includes("burro")) {
                environmentalConcerns.push(entry);
            } else if (title.includes("fire perimeter") || title.includes("fire")) {
                environmentalConcerns.push(entry);
            } else if (title.includes("elevation") || title.includes("3dep")) {
                environmentalConcerns.push(entry);
            } else if (title.includes("grazing")) {
                existingAuthorizations.push(entry);
            } else if (title.includes("oil") && title.includes("gas")) {
                existingAuthorizations.push(entry);
            } else if (title.includes("row") || title.includes("right")) {
                existingAuthorizations.push(entry);
            } else if (title.includes("eplanning")) {
                existingAuthorizations.push(entry);
            } else if (title.includes("land use plan") || title.includes("revision") || title.includes("timber") || title.includes("locatable") || title.includes("mineral")) {
                landUsePlans.push(entry);
            } else if (title.includes("federal land") || title.includes("admin") || title.includes("state boundar") || title.includes("usfws region")) {
                landStatus.push(entry);
            }
        }

        var paragraphs = [];

        // Opening overview
        var acresStr = formatNumber(aoiAcres, 0);
        paragraphs.push("<p>This screening analysis examined <strong>" + totalLayers + " geospatial datasets</strong> to identify land management considerations that may be relevant to permit applications, renewals, or challenges within the approximately <strong>" + escapeHtml(acresStr) + "-acre</strong> project area. Of the datasets reviewed, <strong>" + layersWithHits.length + "</strong> contained features intersecting the area of interest, identifying a total of <strong>" + totalHits + " overlapping features</strong>.</p>");

        // Regulatory framework overview
        paragraphs.push('<p>The Bureau of Land Management administers public lands under the <strong>Federal Land Policy and Management Act of 1976 (FLPMA)</strong> (43 U.S.C. &sect;1701 et seq.), which establishes a multiple-use and sustained yield mandate for the management of public lands and their resources. All authorized uses must conform to the governing <strong>Resource Management Plan (RMP)</strong> prepared pursuant to 43 CFR Part 1600, and discretionary actions are subject to environmental review under the <strong>National Environmental Policy Act (NEPA)</strong> (42 U.S.C. &sect;4321 et seq.). The findings below identify regulatory and resource considerations applicable to the project area based on available geospatial data.</p>');

        // Land status
        if (landStatus.length > 0) {
            var lsNames = landStatus.map(function (f) { return f.name; }).join(", ");
            paragraphs.push("<p><strong>Jurisdictional Context.</strong> The project area has been identified within federal land boundaries based on the following datasets: " + escapeHtml(lsNames) + ". Under FLPMA Section 302 (43 U.S.C. &sect;1732), the BLM has authority to manage these public lands through leases, permits, and easements. Applications for use of these lands must be filed with the BLM field office having jurisdiction (43 CFR &sect;2804.11 for rights-of-way; 43 CFR &sect;2920.5-1 for leases and permits). Applicants should verify which BLM field office has authority over the project area, as this office will be the primary point of contact for all permit applications, including pre-application meetings recommended under 43 CFR &sect;2804.10.</p>");
        }

        // Special designations (high priority for permitting)
        if (specialDesignations.length > 0) {
            var sdNames = specialDesignations.map(function (f) { return "<strong>" + escapeHtml(f.name) + "</strong> (" + f.count + " feature" + (f.count !== 1 ? "s" : "") + ")"; }).join(", ");
            var sdText = "<p><strong>Special Designations.</strong> The following special designations overlap the project area: " + sdNames + ". ";
            sdText += "Under FLPMA and BLM planning regulations (43 CFR Part 1600), special designations are established through the land use planning process and impose specific management prescriptions. ";
            sdText += "<strong>Areas of Critical Environmental Concern (ACECs)</strong> are designated to protect important historical, cultural, scenic, fish and wildlife, or other natural resource values, or to address natural hazards &mdash; they can only be designated or modified through the RMP process. ";
            sdText += "<strong>Wilderness Study Areas (WSAs)</strong> are managed under FLPMA Section 603 to preserve their suitability for possible Congressional designation as Wilderness; activities that would impair wilderness character are generally prohibited. ";
            sdText += "<strong>National Conservation Lands</strong> are protected under the Omnibus Public Land Management Act of 2009 (16 U.S.C. &sect;7202). ";
            sdText += "<strong>Visual Resource Management (VRM)</strong> classifications may restrict the type and scale of surface-disturbing activities. ";
            sdText += "Lands that are specifically segregated or withdrawn from right-of-way uses are not available for grants under 43 CFR &sect;2802.10. ";
            sdText += "Applicants should closely review these designations and consult with the local BLM field office to understand applicable management direction before submitting an application.</p>";
            paragraphs.push(sdText);
        }

        // Environmental concerns
        if (environmentalConcerns.length > 0) {
            var ecNames = environmentalConcerns.map(function (f) { return "<strong>" + escapeHtml(f.name) + "</strong> (" + f.count + ")"; }).join(", ");
            var hasCriticalHabitat = environmentalConcerns.some(function (f) { return f.name.toLowerCase().includes("critical habitat"); });
            var hasMigration = environmentalConcerns.some(function (f) { return f.name.toLowerCase().includes("migration") || f.name.toLowerCase().includes("ungulate"); });
            var hasWildHorse = environmentalConcerns.some(function (f) { return f.name.toLowerCase().includes("wild horse") || f.name.toLowerCase().includes("burro"); });
            var hasFire = environmentalConcerns.some(function (f) { return f.name.toLowerCase().includes("fire"); });

            var ecText = "<p><strong>Environmental and Ecological Considerations.</strong> The following environmental factors were identified: " + ecNames + ". ";
            if (hasCriticalHabitat) {
                ecText += "The presence of federally designated Critical Habitat triggers mandatory <strong>Section 7 consultation</strong> under the Endangered Species Act (ESA, 16 U.S.C. &sect;1536). The BLM, as the action agency, must consult with the U.S. Fish and Wildlife Service (USFWS) to ensure that the proposed action is not likely to jeopardize the continued existence of listed species or destroy/adversely modify designated critical habitat. This consultation results in a Biological Opinion that may include incidental take statements, reasonable and prudent measures, and conservation recommendations, and can significantly extend review timelines. ";
            }
            if (hasMigration) {
                ecText += "Wildlife migration corridors intersecting the project area may require seasonal timing restrictions or project design modifications. Under 43 CFR &sect;2805.12(a)(8), grant holders must comply with terms and conditions designed to protect fish and wildlife habitat, including stipulations for seasonal restrictions to minimize disruption to migratory species. ";
            }
            if (hasWildHorse) {
                ecText += "Wild Horse and Burro Herd Areas or Herd Management Areas are managed under the Wild Free-Roaming Horses and Burros Act of 1971 (16 U.S.C. &sect;1331 et seq.). The presence of these areas may necessitate coordination with BLM wild horse and burro program staff and could result in additional stipulations on authorized activities. ";
            }
            if (hasFire) {
                ecText += "Historical fire perimeters indicate areas that may be subject to post-fire Emergency Stabilization and Rehabilitation (ES&amp;R) activities or Burned Area Rehabilitation (BAR) plans. Altered vegetation and soil conditions may affect project siting and design, and fire prevention obligations apply to all grant holders under 43 CFR &sect;2805.12(a)(4). ";
            }
            ecText += "These environmental factors may require additional NEPA analysis (42 U.S.C. &sect;4321 et seq.), including preparation of an Environmental Assessment (EA) or Environmental Impact Statement (EIS), as well as compliance with the National Historic Preservation Act (NHPA) Section 106 (54 U.S.C. &sect;306108) for cultural resource review.</p>";
            paragraphs.push(ecText);
        }

        // Land use plans
        if (landUsePlans.length > 0) {
            var lupNames = landUsePlans.map(function (f) { return escapeHtml(f.name); }).join(", ");
            paragraphs.push("<p><strong>Land Use Plans and Resource Allocations.</strong> The project area falls within the scope of the following BLM land use plan datasets: " + lupNames + ". Under FLPMA Section 302 and the BLM planning regulations at 43 CFR Part 1600, <strong>all proposed uses must conform to the governing Resource Management Plan (RMP)</strong>. RMPs allocate public land resources for specific uses &mdash; including minerals, timber, grazing, recreation, and conservation &mdash; and establish management prescriptions, allowable uses, and resource-specific stipulations. A proposed use that is inconsistent with the approved RMP may be denied under 43 CFR &sect;2804.26(a)(1) for rights-of-way or 43 CFR &sect;2920.2-5(b)(4) for leases and permits. An RMP amendment (43 CFR &sect;1610.5-5) may be required to accommodate non-conforming uses, which involves additional public participation and NEPA review. Applicants should review the applicable plan documents, available through the BLM <a href='https://eplanning.blm.gov/' target='_blank' rel='noopener'>ePlanning portal</a>, for relevant management direction.</p>");
        }

        // Existing authorizations
        if (existingAuthorizations.length > 0) {
            var eaNames = existingAuthorizations.map(function (f) { return "<strong>" + escapeHtml(f.name) + "</strong> (" + f.count + " feature" + (f.count !== 1 ? "s" : "") + ")"; }).join(", ");
            var eaText = "<p><strong>Existing Authorizations and Land Uses.</strong> The project area overlaps with the following existing authorizations: " + eaNames + ". ";
            eaText += "Under FLPMA, the BLM must consider existing valid rights when evaluating new applications. ";

            var hasGrazingAuth = existingAuthorizations.some(function (f) { return f.name.toLowerCase().includes('grazing'); });
            var hasOilGasAuth = existingAuthorizations.some(function (f) { return f.name.toLowerCase().includes('oil') || f.name.toLowerCase().includes('gas'); });
            var hasRowAuth = existingAuthorizations.some(function (f) { return f.name.toLowerCase().includes('row') || f.name.toLowerCase().includes('right'); });

            if (hasGrazingAuth) {
                eaText += "<strong>Grazing permits and leases</strong> are administered under 43 CFR Part 4100. The Taylor Grazing Act of 1934 and FLPMA govern grazing on public lands; permits are generally issued for 10-year terms and are renewable (43 CFR &sect;4130.2). Per 43 CFR &sect;4110.4-2(b), applicants for solar, wind, or other energy development projects must initiate early discussions with affected grazing permittees. ";
            }
            if (hasOilGasAuth) {
                eaText += "<strong>Oil and gas leases and operations</strong> are administered under the Mineral Leasing Act of 1920 (30 U.S.C. &sect;181 et seq.) and 43 CFR Parts 3100&ndash;3190. Federal leaseholders needing right-of-way access may include ROW requirements with their Application for Permit to Drill (APD) under 43 CFR &sect;2804.12(g). Proposed activities near existing wells or leases should evaluate potential conflicts with mineral rights. ";
            }
            if (hasRowAuth) {
                eaText += "<strong>Existing rights-of-way</strong> are administered under FLPMA Title V and 43 CFR Part 2800. The BLM may require common use of existing corridors (43 CFR &sect;2802.10(b)) and encourages location of new ROWs within designated rights-of-way corridors where practical. ";
            }

            eaText += "Potential conflicts between proposed and existing uses must be addressed during the application review. Coordination with existing authorization holders is recommended, and applicants should demonstrate how their proposed use will be compatible with existing authorized activities.</p>";
            paragraphs.push(eaText);
        }

        // No findings case
        if (layersWithHits.length === 0) {
            paragraphs.push("<p>No intersecting features were identified across any of the screened datasets. While this preliminary screening suggests the project area may have fewer regulatory constraints, this does not replace site-specific environmental review or a formal BLM determination under NEPA (42 U.S.C. &sect;4321 et seq.). Field conditions, unlisted species, cultural resources subject to NHPA Section 106 (54 U.S.C. &sect;306108), and other factors not captured in geospatial datasets may still require evaluation.</p>");
        }

        // Application guidance
        paragraphs.push('<p><strong>Application Guidance.</strong> Right-of-way applications are filed on Standard Form 299 (SF-299) per 43 CFR &sect;2804.12 and must include a project description, construction schedule, capability statement, and maps with GIS data. Other land use authorizations (leases, permits, easements) are governed by 43 CFR Part 2920. All applicants are subject to cost recovery fees (43 CFR &sect;2804.14) categorized by estimated federal processing hours, and must post performance and reclamation bonds before ground-disturbing activities may commence (43 CFR &sect;2805.20). A pre-application meeting with BLM staff (43 CFR &sect;2804.10) is strongly recommended to identify potential routing constraints, environmental issues, and financial obligations before formal filing.</p>');

        // Closing disclaimer
        paragraphs.push('<p><em><strong>Disclaimer.</strong> This screening report is generated automatically from publicly available geospatial datasets and is provided for informational and preliminary planning purposes only. It does not constitute a formal determination by the Bureau of Land Management, a legal opinion, or a guarantee of any permit outcome. The analysis is limited to datasets available through BLM and partner agency web services and does not account for site-specific conditions including, but not limited to: on-the-ground cultural or archaeological resources protected under the National Historic Preservation Act (54 U.S.C. &sect;300101 et seq.) and the Archaeological Resources Protection Act (16 U.S.C. &sect;470aa et seq.); unlisted candidate or sensitive species; Tribal treaty rights and trust responsibilities; state and local permitting requirements; or recent changes to Resource Management Plans. All findings should be verified through field investigation and coordination with appropriate BLM resource specialists. This report does not establish any rights or obligations under FLPMA (43 U.S.C. &sect;1701 et seq.), NEPA (42 U.S.C. &sect;4321 et seq.), or any other federal, state, or local law. Applicants are strongly encouraged to contact their local BLM field office for authoritative guidance prior to submitting permit applications, renewals, or protests.</em></p>');

        return paragraphs.join("\n");
    }

    // ────────────────────────────────────────────
    // buildFinalReportHtmlDoc – HTML template
    // ────────────────────────────────────────────
    function buildFinalReportHtmlDoc({ title, createdAt, totalsHtml, findingsSummaryHtml, aoiSectionHtml, sectionsHtml, dataSourcesHtml, reportId }) {
        const safeTitle = escapeHtml(title || "Final Report");
        const reportIdMeta = reportId ? `<meta name="report-id" content="${escapeHtml(reportId)}" />` : '';

        return `<!doctype html>
            <html lang="en">
            <head>
            <meta charset="utf-8"/>
            <meta name="viewport" content="width=device-width,initial-scale=1"/>
            <title>${safeTitle}</title>
            ${reportIdMeta}
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
                    position: relative;
                }
                .map img{ display:block; width:100%; height:auto; cursor: grab; transition: transform 0.2s ease; transform-origin: center center; }
                .map img.zoomed{ cursor: grab; }
                .map img.panning{ cursor: grabbing; }
                .map-zoom-controls {
                    position: absolute;
                    top: 10px; right: 10px;
                    display: flex; flex-direction: column; gap: 4px;
                    z-index: 10;
                }
                .map-zoom-controls button {
                    width: 32px; height: 32px;
                    border: 1px solid var(--border);
                    border-radius: 6px;
                    background: rgba(255,255,255,0.92);
                    font-size: 18px; font-weight: 700;
                    cursor: pointer;
                    color: var(--blm-green);
                    display: flex; align-items: center; justify-content: center;
                    box-shadow: 0 1px 4px rgba(0,0,0,0.12);
                    transition: background 0.15s;
                    line-height: 1;
                }
                .map-zoom-controls button:hover { background: #e8f5e9; }
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
                table.summary-stats-tbl{
                    width: auto;
                    display: inline-table;
                    margin-top: 16px;
                    border-collapse: collapse;
                    font-size: 13px;
                    background: var(--blm-tan);
                    border-radius: 6px;
                    overflow: hidden;
                }
                table.summary-stats-tbl th{
                    padding: 8px 20px;
                    color: var(--blm-green);
                    font-weight: 600;
                    font-size: 12px;
                    text-transform: uppercase;
                    letter-spacing: 0.3px;
                    border-bottom: 2px solid var(--border);
                    text-align: center;
                    background: rgba(26,71,42,0.05);
                }
                table.summary-stats-tbl td{
                    padding: 10px 20px;
                    text-align: center;
                    font-size: 14px;
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
                table.data-sources-table .desc-row td{
                    padding: 4px 14px 12px 14px;
                    font-size: 11px;
                    color: var(--muted);
                    line-height: 1.5;
                    border-bottom: 2px solid var(--border);
                    background: var(--blm-tan);
                    font-style: italic;
                }
                table.data-sources-table .feat-count{ font-weight: 600; text-align: center; }
                table.data-sources-table .feat-count-zero{ color: var(--muted); opacity: 0.6; }    
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
                /* Section hide/show toggle */
                .section-hide-btn {
                    float: right;
                    background: transparent;
                    border: 1px solid var(--border);
                    border-radius: 5px;
                    padding: 4px 12px;
                    font-size: 11px;
                    font-weight: 600;
                    color: var(--muted);
                    cursor: pointer;
                    transition: background 0.15s, color 0.15s;
                    margin-top: -2px;
                }
                .section-hide-btn:hover {
                    background: #fce4ec;
                    color: #c62828;
                    border-color: #ef9a9a;
                }
                .section {
                    transition: opacity 0.35s ease, padding 0.35s ease, background 0.35s ease;
                }
                .section .section-collapse-wrap {
                    display: grid;
                    grid-template-rows: 1fr;
                    transition: grid-template-rows 0.35s ease;
                    overflow: hidden;
                }
                .section.section-hidden .section-collapse-wrap {
                    grid-template-rows: 0fr;
                }
                .section .section-collapse-inner {
                    min-height: 0;
                    overflow: hidden;
                }
                .section.section-hidden {
                    opacity: 0.5;
                    min-height: 0;
                    padding: 16px 24px;
                    background: #f9f9f7;
                }
                .section.section-hidden h3 {
                    margin: 0;
                    padding-bottom: 0;
                    border-bottom: none;
                    font-size: 14px;
                    color: var(--muted);
                    text-decoration: line-through;
                }
                .section.section-hidden .section-hide-btn {
                    display: inline-block !important;
                    background: #e8f5e9;
                    color: #2e7d32;
                    border-color: #a5d6a7;
                }
                .section.section-hidden .section-hide-btn:hover {
                    background: #c8e6c9;
                }
                .section-hidden + .pagebreak { display: none; }
                .layer-maps-toolbar {
                    display: flex;
                    align-items: center;
                    gap: 10px;
                    margin-bottom: 8px;
                }
                .layer-maps-toolbar .toolbar-btn {
                    background: transparent;
                    border: 1px solid var(--border);
                    border-radius: 5px;
                    padding: 5px 14px;
                    font-size: 12px;
                    font-weight: 600;
                    color: var(--blm-green);
                    cursor: pointer;
                    transition: background 0.15s;
                }
                .layer-maps-toolbar .toolbar-btn:hover {
                    background: #e8f5e9;
                }
                .layer-maps-toolbar .hidden-count {
                    font-size: 12px;
                    color: var(--muted);
                    font-style: italic;
                }
                @media print{
                    html, body{ background: white; }
                    .actions, .hint{ display:none !important; }
                    .map-zoom-controls { display:none !important; }
                    .map img { transform: none !important; }
                    .section-hide-btn { display:none !important; }
                    .layer-maps-toolbar { display:none !important; }
                    .section.section-hidden { display:none !important; }
                    .section-hidden + .pagebreak { display:none !important; }
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
                /* Interactive Data Tables */
                .interactive-table-wrapper {
                    margin-top: 16px;
                }
                .table-toolbar {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    margin-bottom: 8px;
                    flex-wrap: wrap;
                }
                .col-hide-btn {
                    display: inline-block;
                    margin-left: 6px;
                    padding: 0 4px;
                    font-size: 10px;
                    line-height: 16px;
                    color: rgba(255,255,255,0.55);
                    background: transparent;
                    border: 1px solid rgba(255,255,255,0.25);
                    border-radius: 3px;
                    cursor: pointer;
                    vertical-align: middle;
                    transition: color 0.15s, border-color 0.15s;
                }
                .col-hide-btn:hover {
                    color: #fff;
                    border-color: rgba(255,255,255,0.7);
                    background: rgba(255,255,255,0.15);
                }
                .hidden-cols-bar {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 6px;
                    margin-bottom: 8px;
                    min-height: 0;
                }
                .hidden-col-pill {
                    display: inline-flex;
                    align-items: center;
                    gap: 4px;
                    padding: 3px 10px;
                    font-size: 11px;
                    font-weight: 600;
                    color: var(--blm-green);
                    background: #e8f5e9;
                    border: 1px solid var(--blm-green);
                    border-radius: 12px;
                    cursor: pointer;
                    transition: background 0.15s;
                    white-space: nowrap;
                }
                .hidden-col-pill:hover {
                    background: #c8e6c9;
                }
                .hidden-col-pill .pill-x {
                    font-size: 13px;
                    font-weight: 700;
                    line-height: 1;
                }
                .table-scroll {
                    overflow-x: auto;
                    border: 1px solid var(--border);
                    border-radius: 6px;
                    max-height: 500px;
                    overflow-y: auto;
                }
                .interactive-table {
                    width: max-content;
                    min-width: 100%;
                    border-collapse: collapse;
                    font-size: 12px;
                    background: var(--white);
                }
                .interactive-table th {
                    background: var(--blm-green);
                    color: var(--white);
                    padding: 8px 12px;
                    text-align: left;
                    cursor: pointer;
                    user-select: none;
                    white-space: nowrap;
                    position: sticky;
                    top: 0;
                    z-index: 2;
                    font-size: 11px;
                    font-weight: 600;
                    letter-spacing: 0.3px;
                    border-right: 1px solid rgba(255,255,255,0.15);
                }
                .interactive-table th:hover {
                    background: var(--blm-green-light);
                }
                .interactive-table td {
                    padding: 6px 12px;
                    border-bottom: 1px solid var(--border);
                    white-space: nowrap;
                    max-width: 320px;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    font-size: 12px;
                }
                .interactive-table tr:nth-child(even) {
                    background: var(--blm-tan);
                }
                .interactive-table tr:hover {
                    background: rgba(26,71,42,0.06);
                }
                .sort-arrow {
                    font-size: 10px;
                    opacity: 0.6;
                    margin-left: 3px;
                }
                /* Findings Summary */
                .findings-summary {
                    margin: 24px 0;
                    padding: 20px 24px;
                    background: var(--white);
                    border: 1px solid var(--border);
                    border-left: 4px solid var(--blm-green);
                    border-radius: 0 8px 8px 0;
                    line-height: 1.65;
                    font-size: 14px;
                }
                .findings-summary h3 {
                    margin-top: 0;
                    margin-bottom: 12px;
                    color: var(--blm-green);
                    font-size: 17px;
                }
                .findings-summary p {
                    margin: 10px 0;
                }
                .findings-summary em {
                    font-size: 12px;
                    color: var(--muted);
                }
                @media print {
                    .interactive-table-wrapper .table-toolbar { display: none !important; }
                    .hidden-cols-bar { display: none !important; }
                    .col-hide-btn { display: none !important; }
                    .table-scroll { max-height: none; overflow: visible; }
                    .interactive-table { font-size: 9px; }
                    .interactive-table th {
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

                ${findingsSummaryHtml ? '<div class="findings-summary"><h3>Regulatory Screening &amp; Findings Summary</h3>' + findingsSummaryHtml + '</div>' : ''}

                ${aoiSectionHtml || ""}

                <h2>Layer Analysis Maps</h2>
                <div class="layer-maps-toolbar">
                    <button class="toolbar-btn" onclick="toggleAllSections(false)">Hide All</button>
                    <button class="toolbar-btn" onclick="toggleAllSections(true)">Show All</button>
                    <span class="hidden-count" id="hiddenLayerCount"></span>
                </div>
                <div id="layerSectionsContainer">
                ${sectionsHtml || ""}
                </div>

                ${dataSourcesHtml || ""}
                
                <div class="report-footer">
                    <div class="dept-name">Bureau of Land Management</div>
                    <div>U.S. Department of the Interior</div>
                    <div style="margin-top:8px;">This report was generated using geospatial data from BLM and partner agency web services.</div>
                    <div style="margin-top:6px; font-size:11px; color: var(--muted);">Applicable authorities: Federal Land Policy and Management Act (43 U.S.C. &sect;1701 et seq.) &bull; National Environmental Policy Act (42 U.S.C. &sect;4321 et seq.) &bull; Endangered Species Act (16 U.S.C. &sect;1531 et seq.) &bull; National Historic Preservation Act (54 U.S.C. &sect;300101 et seq.) &bull; 43 CFR Parts 1600, 2800, 2920, 3100, 4100</div>
                </div>
            </div>
            <script>
            // ── Report state persistence (IndexedDB) ──
            var _REPORT_ID = (function() {
                var m = document.querySelector('meta[name="report-id"]');
                return m ? m.getAttribute('content') : null;
            })();
            var _STATE_DB_NAME = 'RmpReports';
            var _STATE_STORE = 'reportState';
            var _STATE_DB_VER = 2;
            var _saveTimer = null;

            function _openStateDb() {
                return new Promise(function(resolve, reject) {
                    var req = indexedDB.open(_STATE_DB_NAME, _STATE_DB_VER);
                    req.onupgradeneeded = function() {
                        var db = req.result;
                        if (!db.objectStoreNames.contains('reports')) {
                            var store = db.createObjectStore('reports', { keyPath: 'id' });
                            store.createIndex('expiresAt', 'expiresAt', { unique: false });
                        }
                        if (!db.objectStoreNames.contains(_STATE_STORE)) {
                            db.createObjectStore(_STATE_STORE, { keyPath: 'reportId' });
                        }
                    };
                    req.onsuccess = function() { resolve(req.result); };
                    req.onerror = function() { reject(req.error); };
                });
            }

            function _captureState() {
                var state = { hiddenSections: [], hiddenColumns: {} };
                var sections = document.querySelectorAll('#layerSectionsContainer > .section');
                for (var i = 0; i < sections.length; i++) {
                    if (sections[i].classList.contains('section-hidden')) state.hiddenSections.push(i);
                }
                var wrappers = document.querySelectorAll('.interactive-table-wrapper');
                for (var w = 0; w < wrappers.length; w++) {
                    var wId = wrappers[w].id;
                    if (!wId) continue;
                    var pills = wrappers[w].querySelectorAll('.hidden-col-pill[data-col]');
                    if (pills.length > 0) {
                        state.hiddenColumns[wId] = [];
                        for (var p = 0; p < pills.length; p++) {
                            state.hiddenColumns[wId].push(parseInt(pills[p].getAttribute('data-col')));
                        }
                    }
                }
                return state;
            }

            function _saveState() {
                if (!_REPORT_ID) return;
                if (_saveTimer) clearTimeout(_saveTimer);
                _saveTimer = setTimeout(function() {
                    _openStateDb().then(function(db) {
                        var tx = db.transaction(_STATE_STORE, 'readwrite');
                        tx.objectStore(_STATE_STORE).put({ reportId: _REPORT_ID, state: _captureState(), updatedAt: Date.now() });
                        tx.oncomplete = function() { db.close(); };
                        tx.onerror = function() { db.close(); };
                    }).catch(function() {});
                }, 500);
            }

            function _restoreState() {
                if (!_REPORT_ID) return Promise.resolve(false);
                return _openStateDb().then(function(db) {
                    return new Promise(function(resolve) {
                        var tx = db.transaction(_STATE_STORE, 'readonly');
                        var req = tx.objectStore(_STATE_STORE).get(_REPORT_ID);
                        req.onsuccess = function() {
                            db.close();
                            var rec = req.result;
                            if (!rec || !rec.state) return resolve(false);
                            var s = rec.state;
                            // Restore hidden sections
                            if (s.hiddenSections && s.hiddenSections.length) {
                                var sections = document.querySelectorAll('#layerSectionsContainer > .section');
                                for (var i = 0; i < s.hiddenSections.length; i++) {
                                    var idx = s.hiddenSections[i];
                                    if (sections[idx]) {
                                        sections[idx].classList.add('section-hidden');
                                        var btn = sections[idx].querySelector('.section-hide-btn');
                                        if (btn) btn.innerHTML = '&#x2713; Show';
                                    }
                                }
                                updateHiddenCount();
                            }
                            // Restore hidden columns
                            if (s.hiddenColumns) {
                                for (var wId in s.hiddenColumns) {
                                    var cols = s.hiddenColumns[wId];
                                    for (var c = 0; c < cols.length; c++) {
                                        hideColumn(wId, cols[c]);
                                    }
                                }
                            }
                            resolve(true);
                        };
                        req.onerror = function() { db.close(); resolve(false); };
                    });
                }).catch(function() { return false; });
            }

            // Toggle individual layer section visibility
            function toggleSection(btn) {
                var section = btn.closest('.section');
                if (!section) return;
                var isHidden = section.classList.toggle('section-hidden');
                btn.innerHTML = isHidden ? '&#x2713; Show' : '&#x2715; Hide';
                updateHiddenCount();
                _saveState();
            }
            function toggleAllSections(show) {
                var sections = document.querySelectorAll('#layerSectionsContainer > .section');
                for (var i = 0; i < sections.length; i++) {
                    var btn = sections[i].querySelector('.section-hide-btn');
                    if (show) {
                        sections[i].classList.remove('section-hidden');
                        if (btn) btn.innerHTML = '&#x2715; Hide';
                    } else {
                        sections[i].classList.add('section-hidden');
                        if (btn) btn.innerHTML = '&#x2713; Show';
                    }
                }
                updateHiddenCount();
                _saveState();
            }
            function updateHiddenCount() {
                var el = document.getElementById('hiddenLayerCount');
                if (!el) return;
                var total = document.querySelectorAll('#layerSectionsContainer > .section').length;
                var hidden = document.querySelectorAll('#layerSectionsContainer > .section.section-hidden').length;
                el.textContent = hidden > 0 ? (hidden + ' of ' + total + ' layers hidden') : '';
            }

            // Interactive table: column sorting
            function sortInteractiveTable(th) {
                var table = th.closest('table');
                if (!table) return;
                var colIdx = parseInt(th.getAttribute('data-col'));
                var tbody = table.querySelector('tbody');
                if (!tbody) return;
                var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
                var currentSort = th.getAttribute('data-sort-dir') || 'none';
                var newSort = currentSort === 'asc' ? 'desc' : 'asc';
                // Reset all headers in this table
                var allTh = table.querySelectorAll('th');
                for (var hi = 0; hi < allTh.length; hi++) {
                    allTh[hi].setAttribute('data-sort-dir', 'none');
                    var arrow = allTh[hi].querySelector('.sort-arrow');
                    if (arrow) { arrow.textContent = '\u21C5'; }
                }
                th.setAttribute('data-sort-dir', newSort);
                var sortArrow = th.querySelector('.sort-arrow');
                if (sortArrow) sortArrow.textContent = newSort === 'asc' ? '\u25B2' : '\u25BC';
                rows.sort(function(a, b) {
                    var aCell = a.querySelector('td[data-col="' + colIdx + '"]');
                    var bCell = b.querySelector('td[data-col="' + colIdx + '"]');
                    var aVal = aCell ? (aCell.getAttribute('data-sort-val') || aCell.textContent.trim()) : '';
                    var bVal = bCell ? (bCell.getAttribute('data-sort-val') || bCell.textContent.trim()) : '';
                    var aNum = parseFloat(aVal);
                    var bNum = parseFloat(bVal);
                    var cmp;
                    if (!isNaN(aNum) && !isNaN(bNum)) {
                        cmp = aNum - bNum;
                    } else {
                        cmp = aVal.localeCompare(bVal, undefined, { numeric: true, sensitivity: 'base' });
                    }
                    return newSort === 'asc' ? cmp : -cmp;
                });
                for (var ri = 0; ri < rows.length; ri++) {
                    tbody.appendChild(rows[ri]);
                }
            }
            // Interactive table: hide a column from its header button
            function hideColumn(wrapperId, colIdx) {
                var wrapper = document.getElementById(wrapperId);
                if (!wrapper) return;
                // Hide every cell/header with this data-col index
                var elements = wrapper.querySelectorAll('[data-col="' + colIdx + '"]');
                for (var ei = 0; ei < elements.length; ei++) {
                    elements[ei].style.display = 'none';
                }
                // Find column label from the header
                var th = wrapper.querySelector('thead th[data-col="' + colIdx + '"]');
                var label = th ? (th.getAttribute('data-label') || 'Column ' + colIdx) : 'Column ' + colIdx;
                // Add pill to the hidden-cols-bar
                var bar = wrapper.querySelector('.hidden-cols-bar');
                if (!bar) return;
                var pill = document.createElement('span');
                pill.className = 'hidden-col-pill';
                pill.setAttribute('data-col', colIdx);
                pill.innerHTML = label + ' <span class="pill-x">&#215;</span>';
                pill.title = 'Click to show "' + label + '" column';
                pill.onclick = function() { showColumn(wrapperId, colIdx); };
                bar.appendChild(pill);
                bar.style.display = 'flex';
                _saveState();
            }
            // Interactive table: restore a hidden column via pill click
            function showColumn(wrapperId, colIdx) {
                var wrapper = document.getElementById(wrapperId);
                if (!wrapper) return;
                var elements = wrapper.querySelectorAll('[data-col="' + colIdx + '"]');
                for (var ei = 0; ei < elements.length; ei++) {
                    elements[ei].style.display = '';
                }
                // Remove the pill
                var bar = wrapper.querySelector('.hidden-cols-bar');
                if (bar) {
                    var pills = bar.querySelectorAll('.hidden-col-pill[data-col="' + colIdx + '"]');
                    for (var pi = 0; pi < pills.length; pi++) pills[pi].remove();
                    // Hide bar if empty
                    if (!bar.querySelector('.hidden-col-pill')) bar.style.display = 'none';
                }
                _saveState();
            }

            // Auto-hide columns that are entirely null/empty on page load
            // (only runs if no saved state was restored)
            _restoreState().then(function(restored) {
                if (restored) return; // State was restored — skip auto-hide
                (function autoHideNullColumns() {
                var wrappers = document.querySelectorAll('.interactive-table-wrapper');
                for (var wi = 0; wi < wrappers.length; wi++) {
                    var wrapper = wrappers[wi];
                    var wId = wrapper.id;
                    if (!wId) continue;
                    var ths = wrapper.querySelectorAll('thead th[data-col]');
                    for (var ti = 0; ti < ths.length; ti++) {
                        var colIdx = ths[ti].getAttribute('data-col');
                        var tds = wrapper.querySelectorAll('tbody td[data-col="' + colIdx + '"]');
                        var allEmpty = true;
                        for (var di = 0; di < tds.length; di++) {
                            var val = (tds[di].getAttribute('data-sort-val') || tds[di].textContent || '').trim();
                            if (val !== '' && val !== 'null' && val !== 'Null' && val !== 'NULL' && val !== 'undefined') {
                                allEmpty = false;
                                break;
                            }
                        }
                        if (allEmpty && tds.length > 0) {
                            hideColumn(wId, parseInt(colIdx));
                        }
                    }
                }
                })();
            });

            // ── Map zoom / pan controls ──
            (function() {
                document.querySelectorAll('.map').forEach(function(container) {
                    var img = container.querySelector('img');
                    if (!img) return;
                    var scale = 1, panX = 0, panY = 0, dragging = false, startX = 0, startY = 0;
                    function applyTransform() {
                        img.style.transform = 'scale(' + scale + ') translate(' + panX + 'px,' + panY + 'px)';
                    }
                    var zoomControls = container.querySelector('.map-zoom-controls');
                    if (!zoomControls) return;
                    var btnPlus = zoomControls.querySelector('.zoom-in');
                    var btnMinus = zoomControls.querySelector('.zoom-out');
                    var btnReset = zoomControls.querySelector('.zoom-reset');
                    if (btnPlus) btnPlus.addEventListener('click', function() {
                        scale = Math.min(scale * 1.3, 8);
                        applyTransform();
                    });
                    if (btnMinus) btnMinus.addEventListener('click', function() {
                        scale = Math.max(scale / 1.3, 1);
                        if (scale === 1) { panX = 0; panY = 0; }
                        applyTransform();
                    });
                    if (btnReset) btnReset.addEventListener('click', function() {
                        scale = 1; panX = 0; panY = 0;
                        applyTransform();
                    });
                    // Mouse-wheel zoom
                    container.addEventListener('wheel', function(e) {
                        e.preventDefault();
                        var delta = e.deltaY < 0 ? 1.15 : 1/1.15;
                        scale = Math.max(1, Math.min(scale * delta, 8));
                        if (scale === 1) { panX = 0; panY = 0; }
                        applyTransform();
                    }, { passive: false });
                    // Pan via drag
                    img.addEventListener('mousedown', function(e) {
                        if (scale <= 1) return;
                        dragging = true; startX = e.clientX - panX; startY = e.clientY - panY;
                        img.classList.add('panning');
                        e.preventDefault();
                    });
                    document.addEventListener('mousemove', function(e) {
                        if (!dragging) return;
                        panX = e.clientX - startX; panY = e.clientY - startY;
                        applyTransform();
                    });
                    document.addEventListener('mouseup', function() {
                        if (dragging) { dragging = false; img.classList.remove('panning'); }
                    });
                });
            })();
            </script>
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

    async function buildDataSourcesSection() {
        const services = getConfiguredServices(S.config);
        const lastReport = S.lastReportRowsByLayer || [];

        // Build a lookup: normalised URL → feature count from the analysis
        const countByUrl = new Map();
        for (const item of lastReport) {
            if (item.url) countByUrl.set(String(item.url).replace(/\/+$/, ""), item.count || 0);
        }

        // Fetch descriptions in parallel (with timeout) for any service not already cached
        const descFetchResults = await Promise.allSettled(
            services.map(async svc => {
                // Check cache first
                const cached = S.serviceStatus.get(svc.url + "::desc");
                if (cached) return { url: svc.url, desc: cached };

                try {
                    const pjsonUrl = normalizePjsonUrl(svc.url);
                    const pjson = await fetchJsonWithTimeout(pjsonUrl, 5000);
                    const desc = pickServiceDescription(pjson) || "";
                    if (desc) S.serviceStatus.set(svc.url + "::desc", desc);
                    return { url: svc.url, desc };
                } catch (e) {
                    return { url: svc.url, desc: "" };
                }
            })
        );

        // Merge fetched descriptions
        const descByUrl = new Map();
        for (const r of descFetchResults) {
            if (r.status === "fulfilled" && r.value) {
                descByUrl.set(r.value.url, r.value.desc);
            }
        }

        const rows = services.map(svc => {
            const status = S.serviceStatus.get(svc.url) || "UNKNOWN";
            const statusClass = status === "UP" ? "status-up" : "status-down";
            const desc = descByUrl.get(svc.url) || S.serviceStatus.get(svc.url + "::desc") || "";
            const normalUrl = String(svc.url).replace(/\/+$/, "");
            const featCount = countByUrl.has(normalUrl) ? countByUrl.get(normalUrl) : null;

            const countDisplay = featCount === null
                ? '<span class="feat-count-zero">&mdash;</span>'
                : featCount > 0
                    ? `<b>${escapeHtml(String(featCount))}</b>`
                    : '<span class="feat-count-zero">0</span>';

            const descRow = desc
                ? `<tr class="desc-row"><td colspan="4">${escapeHtml(desc)}</td></tr>`
                : `<tr class="desc-row"><td colspan="4" style="opacity:0.5;">(No description available)</td></tr>`;

            return `
                <tr>
                    <td class="service-name-col">${escapeHtml(svc.title)}</td>
                    <td class="service-url-col"><a href="${escapeHtml(svc.url)}" target="_blank" rel="noopener">${escapeHtml(svc.url)}</a></td>
                    <td class="feat-count">${countDisplay}</td>
                    <td class="service-status-col"><span class="${statusClass}">${status}</span></td>
                </tr>
                ${descRow}
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
                            <th style="width: 30%;">Service Name</th>
                            <th style="width: 40%;">Service URL</th>
                            <th style="width: 15%;">Features in AOI</th>
                            <th style="width: 15%;">Status</th>
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

        const { ensureAoiOnTop, hideAoiMask, captureScreenshotWithWait, waitForTabVisible } = mapUtils;

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
        let plssTownshipOrigRenderer = null;
        for (const l of allLayers) {
            if (l.title && l.title.toLowerCase().includes("township")) {
                plssTownshipLayer = l;
                break;
            }
        }

        // Override township renderer to outline-only (no fill) for cleaner maps
        if (plssTownshipLayer) {
            plssTownshipOrigRenderer = plssTownshipLayer.renderer;
            plssTownshipLayer.renderer = {
                type: "simple",
                symbol: {
                    type: "simple-fill",
                    color: [0, 0, 0, 0],
                    outline: { color: [120, 120, 120, 0.6], width: 0.75 }
                }
            };
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
            // Restore township renderer
            if (plssTownshipLayer && plssTownshipOrigRenderer !== null) {
                try { plssTownshipLayer.renderer = plssTownshipOrigRenderer; } catch (e) { }
            }
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

            // Draw red arrow pointing at the AOI
            {
                const aoiExt = selectionGeom.extent;
                const mw = mainExtent.xmax - mainExtent.xmin;
                const mh = mainExtent.ymax - mainExtent.ymin;
                if (aoiExt && mw > 0 && mh > 0) {
                    const aoiCx = (aoiExt.xmin + aoiExt.xmax) / 2;
                    const aoiCy = (aoiExt.ymin + aoiExt.ymax) / 2;
                    const tipX = ((aoiCx - mainExtent.xmin) / mw) * width;
                    const tipY = ((mainExtent.ymax - aoiCy) / mh) * height;

                    // Shaft from upper-right, 90px at 45°
                    const shaftLen = 90;
                    const d = Math.SQRT1_2; // cos(45°)
                    const tailX = tipX + shaftLen * d;
                    const tailY = tipY - shaftLen * d;

                    // Arrowhead geometry
                    const lineAngle = Math.atan2(tipY - tailY, tipX - tailX);
                    const headLen = 18;
                    const headHalf = Math.PI / 6;
                    const h1x = tipX - headLen * Math.cos(lineAngle - headHalf);
                    const h1y = tipY - headLen * Math.sin(lineAngle - headHalf);
                    const h2x = tipX - headLen * Math.cos(lineAngle + headHalf);
                    const h2y = tipY - headLen * Math.sin(lineAngle + headHalf);

                    ctx.save();
                    ctx.lineCap = "round";
                    ctx.lineJoin = "round";

                    // White halo for contrast
                    ctx.strokeStyle = "#fff";
                    ctx.lineWidth = 6;
                    ctx.beginPath();
                    ctx.moveTo(tailX, tailY);
                    ctx.lineTo(tipX, tipY);
                    ctx.stroke();
                    ctx.beginPath();
                    ctx.moveTo(tipX, tipY);
                    ctx.lineTo(h1x, h1y);
                    ctx.lineTo(h2x, h2y);
                    ctx.closePath();
                    ctx.fillStyle = "#fff";
                    ctx.fill();
                    ctx.stroke();

                    // Red arrow
                    ctx.strokeStyle = "#e63946";
                    ctx.fillStyle = "#e63946";
                    ctx.lineWidth = 3;
                    ctx.beginPath();
                    ctx.moveTo(tailX, tailY);
                    ctx.lineTo(tipX, tipY);
                    ctx.stroke();
                    ctx.beginPath();
                    ctx.moveTo(tipX, tipY);
                    ctx.lineTo(h1x, h1y);
                    ctx.lineTo(h2x, h2y);
                    ctx.closePath();
                    ctx.fill();

                    ctx.restore();
                }
            }

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
            captureScreenshotWithWait, waitForTabVisible,
            acquireWakeLock, releaseWakeLock,
            getLayerGeometryType, makeRendererOpaque, getPresetRenderer
        } = mapUtils;

        const {
            queryAllFeaturesPaged, computeElevationStats,
            computeLayerCoverageStats, buildPerFeatureTable, SQM_PER_ACRE
        } = queryEngine;

        // Acquire Wake Lock to prevent device sleeping during report generation
        await acquireWakeLock();

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
            } else if (S.aoiSource === "upload") {
                aoiMethod = `Uploaded File: ${S.aoiSourceLayerTitle || "unknown"}`;
            }

            // 2d. AOI Maps
            _setStatus("building final report\u2026 (generating AOI maps)");
            const aoiMapsHtml = await generateAoiMapsWithCircles();

            // STEP 3: Generate per-layer map sections
            _setStatus("building final report\u2026 (generating layer maps)");

            const paddingFactor = config?.visualReport?.paddingFactor ?? 1.12;
            const width = config?.visualReport?.screenshotWidth ?? 1400;

            let fixedExtent = null;
            const ext = selectionGeom?.extent;
            if (ext && ext.expand) fixedExtent = ext.expand(paddingFactor);

            const targets = lastReportRowsByLayer
                .filter(x => (x?.count || 0) > 0)
                .filter(x => (x?._layer && x?._exportQuery) || x?.__isImageService)
                .filter(x => !(x.title && x.title.toLowerCase().includes("state boundaries")));

            // Sort so BLM Administrative Units comes first
            targets.sort((a, b) => {
                const aAdmin = a.title && a.title.toLowerCase().includes("administrative unit") ? 0 : 1;
                const bAdmin = b.title && b.title.toLowerCase().includes("administrative unit") ? 0 : 1;
                return aAdmin - bAdmin;
            });

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
                            imgLayerOpts.rasterFunction = { functionName: item.__renderingRule };
                        }
                        const temp = new ImageryLayer(imgLayerOpts);
                        view.map.add(temp);

                        try {
                            setVisibilityForScreenshot(temp);
                            await waitForLayerReadyToCapture(temp, view, { timeoutMs: 10000 });
                            if (fixedExtent) await view.goTo(fixedExtent, { animate: false });
                            else await view.goTo(selectionGeom.extent.expand(1.15), { animate: false });
                            // Re-check that layer has finished rendering at the new extent
                            await waitForLayerReadyToCapture(temp, view, { timeoutMs: 10000 });
                            await waitForViewStationary(1500);

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
                                <h3><button class="section-hide-btn" onclick="toggleSection(this)">✕ Hide</button>${escapeHtml(item.title)}</h3>
                                <div class="section-collapse-wrap"><div class="section-collapse-inner">
                                <div class="map">
                                  <div class="map-zoom-controls">
                                      <button class="zoom-in" title="Zoom in">+</button>
                                      <button class="zoom-out" title="Zoom out">&minus;</button>
                                      <button class="zoom-reset" title="Reset zoom" style="font-size:13px;">&#8634;</button>
                                  </div>
                                  <img src="${dataUrl}" alt="${escapeHtml(item.title)}" style="width:100%; border-radius:8px;" />
                                </div>
                                <table class="info-table" style="margin-top:16px;">
                                  <tr><td style="width:200px;">Service Name</td><td><b>${escapeHtml(meta.name || item.title)}</b></td></tr>
                                  <tr><td>Type</td><td>Image Service (Elevation/Raster)</td></tr>
                                  ${elevStatsHtml}
                                  ${meta.copyright ? `<tr><td>Source</td><td>${escapeHtml(meta.copyright)}</td></tr>` : ''}
                                </table>
                                </div></div>
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
                    const tempOpts = {
                        url: item.url,
                        title: item.title,
                        outFields: ["*"],
                        visible: true,
                        opacity: 0.8
                    };
                    const temp = new FeatureLayer(tempOpts);

                    // Always override scale to ensure layer draws at any zoom
                    temp.minScale = 0;
                    temp.maxScale = 0;

                    view.map.add(temp);

                    // Wait for the layer to load so its service renderer is available
                    try { await temp.when(); } catch (e) { /* continue even if load fails */ }

                    // Thicken polygon outlines while preserving the service's original symbology
                    if (tempGeomType && String(tempGeomType).toLowerCase().includes('polygon') && temp.renderer) {
                        try {
                            const r = temp.renderer.clone();
                            const MIN_OUTLINE = 3;

                            function thickenOutline(sym) {
                                if (!sym) return;
                                if (sym.outline) {
                                    sym.outline.width = Math.max(sym.outline.width || 0, MIN_OUTLINE);
                                } else {
                                    sym.outline = { color: [0, 0, 0, 0.8], width: MIN_OUTLINE };
                                }
                            }

                            if (r.symbol) thickenOutline(r.symbol);
                            if (r.defaultSymbol) thickenOutline(r.defaultSymbol);
                            if (r.uniqueValueInfos) r.uniqueValueInfos.forEach(uv => thickenOutline(uv.symbol));
                            if (r.classBreakInfos) r.classBreakInfos.forEach(cb => thickenOutline(cb.symbol));

                            temp.renderer = r;
                        } catch (e) {
                            console.warn("Could not thicken outline for", item.title, e);
                        }
                    }

                    try {
                        setVisibilityForScreenshot(temp);
                        await waitForLayerReadyToCapture(temp, view, { timeoutMs: 15000 });
                        if (fixedExtent) await view.goTo(fixedExtent, { animate: false });
                        else await view.goTo(selectionGeom.extent.expand(1.15), { animate: false });
                        // Re-check that layer has finished rendering at the new extent
                        await waitForLayerReadyToCapture(temp, view, { timeoutMs: 15000 });
                        await waitForViewStationary(1500);

                        const dataUrl = await captureScreenshotWithWait({ width });
                        if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

                        // Determine geometry class for this layer
                        const isPolygonLayer = tempGeomType && String(tempGeomType).toLowerCase().includes('polygon');

                        // Only compute coverage stats for polygon layers
                        let acresCovered = 0;
                        let pctCovered   = 0;
                        if (isPolygonLayer) {
                            const cov = await computeLayerCoverageStats(item, selectionGeom);
                            acresCovered = cov ? cov.acresCovered : 0;
                            pctCovered   = cov ? cov.pctAoiCovered : 0;
                        }

                        const layerAttrSummary = generateLayerAttributeSummary(item);
                        const perFeatureTableHtml = (item.count > 0)
                            ? await buildPerFeatureTable(item, selectionGeom, i)
                            : "";

                        // Only flag low-coverage for polygon layers (not points or lines)
                        const isSingleFeatureLowCoverage = isPolygonLayer && (item.count === 1 && pctCovered < 3);
                        const lowCoverageWarningHtml = isSingleFeatureLowCoverage
                            ? `<div style="margin-top:12px; padding:12px 16px; background-color:#fff3cd; border:1px solid #ffc107; border-radius:6px; font-size:14px; line-height:1.5;">
                                <span style="color:#856404;">\u26A0\uFE0F <b>Low Coverage Warning:</b> This feature covers less than 3% of the AOI. This may indicate a polygon sliver or boundary artifact rather than meaningful overlap.</span>
                               </div>`
                            : "";

                        // Coverage rows only for polygon layers
                        const coverageRowsHtml = isPolygonLayer
                            ? `<th>Percent of AOI</th>`
                            : "";
                        const coverageValHtml = isPolygonLayer
                            ? `<td><b>${formatNumber(pctCovered, 2)}%</b>${isSingleFeatureLowCoverage ? ' <span style="color:#856404;" title="Low coverage \u2014 possible sliver or boundary artifact">\u26A0\uFE0F</span>' : ''}</td>`
                            : "";

                        sectionsHtml += `
                        <div class="section">
                            <h3><button class="section-hide-btn" onclick="toggleSection(this)">✕ Hide</button>${escapeHtml(item.title)}</h3>
                            <div class="section-collapse-wrap"><div class="section-collapse-inner">
                            <div class="map">
                                <div class="map-zoom-controls">
                                    <button class="zoom-in" title="Zoom in">+</button>
                                    <button class="zoom-out" title="Zoom out">&minus;</button>
                                    <button class="zoom-reset" title="Reset zoom" style="font-size:13px;">&#8634;</button>
                                </div>
                                <img src="${dataUrl}" alt="AOI + ${escapeHtml(item.title)}"/>
                            </div>
                            <table class="summary-stats-tbl">
                                <thead><tr>
                                    <th>AOI Area</th>
                                    <th>Intersecting Features</th>
                                    ${coverageRowsHtml}
                                </tr></thead>
                                <tbody><tr>
                                    <td><b>${formatNumber(aoiAcres, 2)}</b> acres</td>
                                    <td><b>${escapeHtml(String(item.count || 0))}</b></td>
                                    ${coverageValHtml}
                                </tr></tbody>
                            </table>
                            ${layerAttrSummary ? `<table class="metaTbl">${layerAttrSummary}</table>` : ''}
                            ${lowCoverageWarningHtml}
                            ${perFeatureTableHtml}
                            </div></div>
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
            const dataSourcesHtml = await buildDataSourcesSection();

            // STEP 4b: Generate findings summary paragraph
            const findingsSummaryHtml = generateFindingsSummary(lastReportRowsByLayer, aoiAcres);

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



            const aoiSectionHtml = `
                <h2>Area of Interest</h2>
                ${aoiMapsHtml}
                <div class="aoi-details">
                    ${stateHtml}
                    ${legalHtml}
                    <div class="aoi-field"><span class="aoi-label">Area:</span> ${formatNumber(aoiAcres, 2)} acres</div>
                    <div class="aoi-field"><span class="aoi-label">Method:</span> ${escapeHtml(aoiMethod)}</div>
                </div>
            `;

            const htmlDoc = buildFinalReportHtmlDoc({
                title: "Land & Resource Intersection Analysis Report",
                createdAt: formatDateTimeForReport(new Date()),
                totalsHtml,
                findingsSummaryHtml,
                aoiSectionHtml,
                sectionsHtml,
                dataSourcesHtml
            });

            cachedFinalReportHtml = htmlDoc;

            // Persist report to IndexedDB for shareable URL
            try {
                const reportId = await saveReportToDb(htmlDoc);
                _lastReportId = reportId;

                // Inject reportId into the cached HTML so the report can persist its own UI state
                cachedFinalReportHtml = htmlDoc.replace(
                    '</title>',
                    `</title>\n            <meta name="report-id" content="${reportId}" />`
                );
                // Update the stored HTML with the reportId too
                await saveReportToDb(cachedFinalReportHtml, reportId);

                if (finalReportStatus) {
                    finalReportStatus.textContent = "Report ready.";
                }
            } catch (dbErr) {
                console.warn("Could not save report to IndexedDB:", dbErr);
                _lastReportId = null;
                if (finalReportStatus) finalReportStatus.textContent = "Report ready.";
            }

        } catch (e) {
            console.error(e);
            if (finalReportStatus) finalReportStatus.textContent = "Failed to build report (see console).";
        } finally {
            await releaseWakeLock();
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
            generateFindingsSummary,
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
            setCachedFinalReportHtml: (v) => { cachedFinalReportHtml = v; },
            // IndexedDB report management
            getLastReportId: () => _lastReportId,
            getReportShareUrl,
            loadReportFromDb,
            cleanupExpiredReports
        };
    }

    return { init };
});
