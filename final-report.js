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
    "app/config-helpers",
    "app/summary-engine"
], function (configHelpers, summaryEngineModule) {
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
    let summaryEngine; // summary-engine API

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

    // ────────────────────────────────────────────
    // Report Layer Buckets (same as Permit Screening UI)
    // ────────────────────────────────────────────
    const REPORT_BUCKETS = {
        "land-status": {
            label: "Land Status & Authority", icon: "🏛️",
            description: "Federal land ownership, administrative boundaries, tribal lands, and jurisdictional authority.",
            patterns: [/federal lands/i, /admin.*unit/i, /state boundar/i, /usfws.*region/i, /aoi source/i,
                        /bia.*aian/i, /indian/i, /alaska.*native/i, /tribal/i, /surface.*ownership/i,
                        /land use planning bound/i]
        },
        "land-use": {
            label: "Land Use Plans & Allocations", icon: "📑",
            description: "Resource Management Plans, timber and mineral allocations.",
            patterns: [/land use plan/i, /revision.*development/i, /timber/i, /locatable.*mineral/i,
                        /taylor grazing/i, /tga/i]
        },
        "special": {
            label: "Special Designations", icon: "⭐",
            description: "ACECs, wilderness, conservation lands, wild & scenic rivers, and other special designations.",
            patterns: [/acec/i, /critical environmental/i, /nlcs/i, /conservation area/i, /national monument/i,
                        /wilderness/i, /wsa/i, /recreation site/i, /lwcf/i, /conservation fund/i, /visual resource/i,
                        /wild.*scenic.*river/i, /roadless/i, /national forest bound/i, /national wildlife refuge/i, /nwr/i]
        },
        "environmental": {
            label: "Environmental & ESA", icon: "🌿",
            description: "Critical habitat, wetlands, hydrology, wildlife corridors, flood hazards, and fire history.",
            patterns: [/critical habitat/i, /ungulate/i, /migration/i, /wild horse/i, /burro/i, /elevation/i, /fire perim/i,
                        /wetland/i, /nwi/i, /riparian/i, /nhd/i, /hydrography/i, /watershed/i, /wbd/i,
                        /flood/i, /nfhl/i, /fema/i, /sagebrush/i, /fiat/i, /danl/i, /disturbance/i,
                        /at.risk.*species/i, /t\&e/i, /threatened/i]
        },
        "authorizations": {
            label: "Existing Authorizations", icon: "📝",
            description: "Active permits, leases, rights-of-way, mining claims, and other authorizations.",
            patterns: [/grazing allot/i, /grazing pasture/i, /oil.*gas/i, /mlrs.*row/i, /lua.*row/i, /eplanning/i, /plss.*parcel/i,
                        /mining claim/i, /lua.*lease/i, /lua.*permit/i, /lua.*easem/i, /geothermal/i, /coal case/i,
                        /oil shale/i, /non.energy/i, /mineral material/i, /locatable notice/i, /locatable plan/i,
                        /participating area/i, /agreement/i, /gtlf/i, /road.*trail/i]
        }
    };

    const BUCKET_ORDER = ["land-status", "land-use", "special", "environmental", "authorizations", "uncategorized"];

    function categorizeLayersIntoBuckets(items) {
        const buckets = {};
        for (const key of Object.keys(REPORT_BUCKETS)) buckets[key] = [];
        buckets["uncategorized"] = [];
        const cfgByUrl = S.layerCfgByUrl;
        for (const item of (items || [])) {
            const title = (item.title || "");
            // 1. Prefer explicit category from config (O(1) lookup)
            const cfgEntry = cfgByUrl?.get(String(item.url || ""));
            const cfgCategory = cfgEntry?.cfg?.category;
            if (cfgCategory && REPORT_BUCKETS[cfgCategory]) {
                buckets[cfgCategory].push(item);
                continue;
            }
            // 2. Fallback: regex title matching for layers without a category field
            let placed = false;
            for (const [bk, bd] of Object.entries(REPORT_BUCKETS)) {
                if (bd.patterns.some(p => p.test(title))) { buckets[bk].push(item); placed = true; break; }
            }
            if (!placed) buckets["uncategorized"].push(item);
        }
        return buckets;
    }

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
    // buildLayerNarrative — readable summary paragraph under each map
    // ────────────────────────────────────────────
    function buildLayerNarrative({ aoiAcres, featureCount, layerTitle, acresCovered, pctCovered, isPolygon, isImagery, elevStats }) {
        const aoiStr   = `<b>${formatNumber(aoiAcres, 2)}</b> acres`;
        const titleStr = `<b>${escapeHtml(layerTitle)}</b>`;

        let narrative = `The Area of Interest is ${aoiStr}. `;

        if (isImagery) {
            // Imagery / raster layer narrative
            narrative += `The Area of Interest intersects with ${titleStr}, which is an Imagery Layer; therefore details regarding the analysis of this layer are described below.`;

            if (elevStats) {
                narrative += `<div style="margin-top:10px;">` +
                    `Elevation within the Area of Interest ranges from a minimum of <b>${formatNumber(elevStats.minFt, 0)}</b> ft ` +
                    `(<b>${formatNumber(elevStats.min, 1)}</b> m) to a maximum of <b>${formatNumber(elevStats.maxFt, 0)}</b> ft ` +
                    `(<b>${formatNumber(elevStats.max, 1)}</b> m), an elevation change of <b>${formatNumber(elevStats.elevationChangeFt, 0)}</b> ft ` +
                    `(<b>${formatNumber(elevStats.elevationChange, 1)}</b> m).` +
                    (elevStats.meanFt ? ` The mean elevation is <b>${formatNumber(elevStats.meanFt, 0)}</b> ft (<b>${formatNumber(elevStats.mean, 1)}</b> m).` : '') +
                    `</div>`;
            }
        } else {
            // Feature layer narrative
            const countStr = `<b>${escapeHtml(String(featureCount))}</b> feature${featureCount !== 1 ? 's' : ''}`;
            narrative += `Within the Area of Interest, `;

            if (isPolygon && acresCovered > 0) {
                const covStr = `<b>${formatNumber(acresCovered, 2)}</b> acres`;
                const pctStr = `<b>${formatNumber(pctCovered, 1)}%</b>`;
                narrative += `${covStr} (${countStr}) from the ${titleStr} layer was detected. `;
                narrative += `This layer covers approximately ${pctStr} of the Area of Interest.`;
            } else {
                narrative += `${countStr} from the ${titleStr} layer ${featureCount !== 1 ? 'were' : 'was'} detected.`;
            }
        }

        return `<div class="layer-narrative">${narrative}</div>`;
    }

    // ────────────────────────────────────────────
    // generateLayerAttributeSummary
    // Delegates to summary-engine.js (plugin + generic auto-classifier).
    // ────────────────────────────────────────────
    function generateLayerAttributeSummary(item) {
        if (summaryEngine) return summaryEngine.generate(item);
        return "";
    }

    // ────────────────────────────────────────────
    // generateFindingsSummary – human-readable paragraph
    // ────────────────────────────────────────────
    function generateFindingsSummary(reportItems, aoiAcres) {
        if (!reportItems || !reportItems.length) return "";

        const totalLayers = reportItems.length;
        const layersWithHits = reportItems.filter(function (x) { return (x.count || 0) > 0; });
        const totalHits = reportItems.reduce(function (s, x) { return s + (x.count || 0); }, 0);

        // Categorize findings — use config category field first, title regex fallback
        var specialDesignations = [];
        var environmentalConcerns = [];
        var existingAuthorizations = [];
        var landUsePlans = [];
        var landStatus = [];

        // Map config category values to the local bucket arrays
        var CATEGORY_MAP = {
            "special":        specialDesignations,
            "environmental":  environmentalConcerns,
            "authorizations": existingAuthorizations,
            "land-use":       landUsePlans,
            "land-status":    landStatus
        };

        for (var idx = 0; idx < layersWithHits.length; idx++) {
            var item = layersWithHits[idx];
            var title = (item.title || "").toLowerCase();
            var count = item.count || 0;
            var entry = { name: item.title, count: count };

            // Try config-level category first
            var catBucket = null;
            if (S && S.layerCfgByUrl && item.url) {
                var cfgEntry = S.layerCfgByUrl.get(item.url);
                var cat = cfgEntry && (cfgEntry.cfg || cfgEntry).category;
                if (cat && CATEGORY_MAP[cat]) {
                    catBucket = CATEGORY_MAP[cat];
                }
            }

            if (catBucket) {
                catBucket.push(entry);
            } else if (title.includes("acec") || title.includes("critical environmental concern")) {
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
        paragraphs.push('<h4 class="findings-subhead">Regulatory Framework</h4>');
        paragraphs.push('<p>The Bureau of Land Management administers public lands under the <strong>Federal Land Policy and Management Act of 1976 (FLPMA)</strong> (43 U.S.C. &sect;1701 et seq.), which establishes a multiple-use and sustained yield mandate for the management of public lands and their resources. All authorized uses must conform to the governing <strong>Resource Management Plan (RMP)</strong> prepared pursuant to 43 CFR Part 1600, and discretionary actions are subject to environmental review under the <strong>National Environmental Policy Act (NEPA)</strong> (42 U.S.C. &sect;4321 et seq.). The findings below identify regulatory and resource considerations applicable to the project area based on available geospatial data.</p>');

        // Land status
        if (landStatus.length > 0) {
            var lsNames = landStatus.map(function (f) { return f.name; }).join(", ");
            paragraphs.push('<h4 class="findings-subhead">Jurisdictional Context</h4>');
            paragraphs.push("<p>The project area has been identified within federal land boundaries based on the following datasets: " + escapeHtml(lsNames) + ". Under FLPMA Section 302 (43 U.S.C. &sect;1732), the BLM has authority to manage these public lands through leases, permits, and easements. Applications for use of these lands must be filed with the BLM field office having jurisdiction (43 CFR &sect;2804.11 for rights-of-way; 43 CFR &sect;2920.5-1 for leases and permits). Applicants should verify which BLM field office has authority over the project area, as this office will be the primary point of contact for all permit applications, including pre-application meetings recommended under 43 CFR &sect;2804.10.</p>");
        }

        // Special designations (high priority for permitting)
        if (specialDesignations.length > 0) {
            var sdNames = specialDesignations.map(function (f) { return "<strong>" + escapeHtml(f.name) + "</strong> (" + f.count + " feature" + (f.count !== 1 ? "s" : "") + ")"; }).join(", ");
            paragraphs.push('<h4 class="findings-subhead">Special Designations</h4>');
            var sdText = "<p>The following special designations overlap the project area: " + sdNames + ". ";
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

            paragraphs.push('<h4 class="findings-subhead">Environmental &amp; Ecological Considerations</h4>');
            var ecText = "<p>The following environmental factors were identified: " + ecNames + ". ";
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
            paragraphs.push('<h4 class="findings-subhead">Land Use Plans &amp; Resource Allocations</h4>');
            paragraphs.push("<p>The project area falls within the scope of the following BLM land use plan datasets: " + lupNames + ". Under FLPMA Section 302 and the BLM planning regulations at 43 CFR Part 1600, <strong>all proposed uses must conform to the governing Resource Management Plan (RMP)</strong>. RMPs allocate public land resources for specific uses &mdash; including minerals, timber, grazing, recreation, and conservation &mdash; and establish management prescriptions, allowable uses, and resource-specific stipulations. A proposed use that is inconsistent with the approved RMP may be denied under 43 CFR &sect;2804.26(a)(1) for rights-of-way or 43 CFR &sect;2920.2-5(b)(4) for leases and permits. An RMP amendment (43 CFR &sect;1610.5-5) may be required to accommodate non-conforming uses, which involves additional public participation and NEPA review. Applicants should review the applicable plan documents, available through the BLM <a href='https://eplanning.blm.gov/' target='_blank' rel='noopener'>ePlanning portal</a>, for relevant management direction.</p>");
        }

        // Existing authorizations
        if (existingAuthorizations.length > 0) {
            var eaNames = existingAuthorizations.map(function (f) { return "<strong>" + escapeHtml(f.name) + "</strong> (" + f.count + " feature" + (f.count !== 1 ? "s" : "") + ")"; }).join(", ");
            paragraphs.push('<h4 class="findings-subhead">Existing Authorizations &amp; Land Uses</h4>');
            var eaText = "<p>The project area overlaps with the following existing authorizations: " + eaNames + ". ";
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
        paragraphs.push('<h4 class="findings-subhead">Application Guidance</h4>');
        paragraphs.push('<p>Right-of-way applications are filed on Standard Form 299 (SF-299) per 43 CFR &sect;2804.12 and must include a project description, construction schedule, capability statement, and maps with GIS data. Other land use authorizations (leases, permits, easements) are governed by 43 CFR Part 2920. All applicants are subject to cost recovery fees (43 CFR &sect;2804.14) categorized by estimated federal processing hours, and must post performance and reclamation bonds before ground-disturbing activities may commence (43 CFR &sect;2805.20). A pre-application meeting with BLM staff (43 CFR &sect;2804.10) is strongly recommended to identify potential routing constraints, environmental issues, and financial obligations before formal filing.</p>');

        // Closing disclaimer
        paragraphs.push('<h4 class="findings-subhead">Disclaimer</h4>');
        paragraphs.push('<p><em>This screening report is generated automatically from publicly available geospatial datasets and is provided for informational and preliminary planning purposes only. It does not constitute a formal determination by the Bureau of Land Management, a legal opinion, or a guarantee of any permit outcome. The analysis is limited to datasets available through BLM and partner agency web services and does not account for site-specific conditions including, but not limited to: on-the-ground cultural or archaeological resources protected under the National Historic Preservation Act (54 U.S.C. &sect;300101 et seq.) and the Archaeological Resources Protection Act (16 U.S.C. &sect;470aa et seq.); unlisted candidate or sensitive species; Tribal treaty rights and trust responsibilities; state and local permitting requirements; or recent changes to Resource Management Plans. All findings should be verified through field investigation and coordination with appropriate BLM resource specialists. This report does not establish any rights or obligations under FLPMA (43 U.S.C. &sect;1701 et seq.), NEPA (42 U.S.C. &sect;4321 et seq.), or any other federal, state, or local law. Applicants are strongly encouraged to contact their local BLM field office for authoritative guidance prior to submitting permit applications, renewals, or protests.</em></p>');

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
                .layer-narrative {
                    margin: 16px 0;
                    padding: 14px 20px;
                    background: var(--blm-tan);
                    border-left: 4px solid var(--blm-green);
                    border-radius: 4px;
                    font-size: 13.5px;
                    line-height: 1.65;
                    color: var(--text);
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
                /* Bucket category headers */
                .bucket-header {
                    background: linear-gradient(135deg, var(--blm-green) 0%, #2d5a3d 100%);
                    color: white;
                    padding: 24px 32px;
                    margin: 32px 0 20px 0;
                    border-radius: 12px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                }
                .bucket-header h2 {
                    margin: 0 0 8px 0;
                    font-size: 22px;
                    font-weight: 700;
                    border: none;
                    padding: 0;
                    color: white;
                }
                .bucket-header .bucket-description {
                    margin: 0;
                    font-size: 14px;
                    opacity: 0.9;
                    line-height: 1.5;
                }
                @media print {
                    .bucket-header {
                        background: var(--blm-green) !important;
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                        break-after: avoid;
                        page-break-after: avoid;
                    }
                }
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
                    margin: 0 0 24px 0;
                    padding: 20px 24px;
                    background: var(--white);
                    border: 1px solid var(--border);
                    border-left: 4px solid var(--blm-green);
                    border-radius: 0 8px 8px 0;
                    line-height: 1.65;
                    font-size: 14px;
                }
                .findings-summary p {
                    margin: 10px 0;
                }
                .findings-summary em {
                    font-size: 12px;
                    color: var(--muted);
                }
                /* Section numbering */
                .section-num {
                    color: var(--blm-gold);
                    margin-right: 6px;
                    font-weight: 700;
                }
                /* Section intro text */
                .section-intro {
                    font-size: 13px;
                    color: var(--muted);
                    margin: -4px 0 16px 0;
                    line-height: 1.55;
                    max-width: 700px;
                }
                /* Table of Contents */
                .report-toc {
                    margin: 20px 0 28px 0;
                    padding: 16px 24px 16px 28px;
                    background: var(--blm-tan);
                    border: 1px solid var(--border);
                    border-radius: 8px;
                }
                .toc-heading {
                    margin: 0 0 10px 0;
                    font-size: 16px;
                    color: var(--blm-green);
                    letter-spacing: 0.3px;
                }
                .toc-list {
                    margin: 0;
                    padding-left: 22px;
                    list-style: decimal;
                }
                .toc-list li {
                    margin: 4px 0;
                    font-size: 14px;
                }
                .toc-list a {
                    color: var(--blm-green);
                    text-decoration: none;
                    font-weight: 500;
                    transition: color 0.15s;
                }
                .toc-list a:hover {
                    color: var(--blm-gold);
                    text-decoration: underline;
                }
                .toc-sublist {
                    margin: 6px 0 4px 0;
                    padding-left: 20px;
                    list-style: none;
                }
                .toc-sublist li {
                    margin: 3px 0;
                    font-size: 13px;
                    font-weight: 400;
                }
                .toc-sublist a {
                    font-weight: 400;
                }
                /* Findings sub-headings */
                .findings-subhead {
                    margin: 20px 0 6px 0;
                    font-size: 15px;
                    color: var(--blm-brown);
                    border-bottom: 1px solid var(--border);
                    padding-bottom: 4px;
                    letter-spacing: 0.2px;
                }
                .findings-summary .findings-subhead:first-child {
                    margin-top: 8px;
                }
                /* Bookmark button variant */
                .btn-bookmark {
                    background: #5c4827;
                }
                .btn-bookmark:hover {
                    background: #7a6235;
                }
                /* Accessibility Widget (report-embedded) */
                .a11y-widget {
                    position: fixed;
                    bottom: 16px;
                    right: 16px;
                    z-index: 50000;
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                }
                .a11y-toggle {
                    width: 44px;
                    height: 44px;
                    border-radius: 50%;
                    border: 2px solid #1a73e8;
                    background: #fff;
                    color: #1a73e8;
                    cursor: pointer;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.25);
                    transition: background 0.2s, color 0.2s, transform 0.15s;
                }
                .a11y-toggle:hover,
                .a11y-toggle:focus-visible {
                    background: #1a73e8;
                    color: #fff;
                    transform: scale(1.08);
                    outline: 2px solid #fff;
                    outline-offset: 2px;
                }
                .a11y-toggle[aria-expanded="true"] {
                    background: #1a73e8;
                    color: #fff;
                }
                .a11y-menu {
                    position: absolute;
                    bottom: 52px;
                    right: 0;
                    min-width: 220px;
                    background: #fff;
                    border: 1px solid #d0d0d0;
                    border-radius: 10px;
                    box-shadow: 0 4px 16px rgba(0,0,0,0.18);
                    padding: 6px 0;
                    animation: a11yFadeIn 0.15s ease;
                }
                .a11y-menu.hidden { display: none; }
                @keyframes a11yFadeIn {
                    from { opacity: 0; transform: translateY(8px); }
                    to   { opacity: 1; transform: translateY(0); }
                }
                .a11y-menu-header {
                    padding: 8px 14px 6px;
                    font-size: 11px;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    color: #666;
                    border-bottom: 1px solid #eee;
                    margin-bottom: 4px;
                }
                .a11y-option {
                    display: block;
                    width: 100%;
                    text-align: left;
                    padding: 8px 14px;
                    border: none;
                    background: none;
                    font-size: 13px;
                    color: #333;
                    cursor: pointer;
                    transition: background 0.12s;
                }
                .a11y-option:hover,
                .a11y-option:focus-visible {
                    background: #e8f0fe;
                    outline: none;
                }
                .a11y-option[aria-checked="true"] {
                    font-weight: 700;
                    color: #1a73e8;
                    background: #e8f0fe;
                }
                .a11y-option[aria-checked="true"]::before {
                    content: '\u2713 ';
                }
                .a11y-option small { color: #888; font-weight: 400; }
                /* Color-vision filter classes — applied to .cv-filter-wrap so position:fixed widget stays viewport-pinned */
                .cv-filter-wrap.cv-protanopia    { -webkit-filter: url(#cv-protanopia); filter: url(#cv-protanopia); }
                .cv-filter-wrap.cv-deuteranopia  { -webkit-filter: url(#cv-deuteranopia); filter: url(#cv-deuteranopia); }
                .cv-filter-wrap.cv-tritanopia    { -webkit-filter: url(#cv-tritanopia); filter: url(#cv-tritanopia); }
                .cv-filter-wrap.cv-achromatopsia { -webkit-filter: url(#cv-achromatopsia); filter: url(#cv-achromatopsia); }
                .cv-filter-wrap.cv-highcontrast  { -webkit-filter: url(#cv-highcontrast); filter: url(#cv-highcontrast); }
                /* Back-to-top button */
                .back-to-top {
                    display: none;
                    position: fixed;
                    bottom: 28px;
                    right: 80px;
                    width: 48px;
                    height: 48px;
                    border-radius: 50%;
                    background: var(--blm-green);
                    color: var(--white);
                    font-size: 16px;
                    font-weight: 700;
                    text-decoration: none;
                    align-items: center;
                    justify-content: center;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.22);
                    z-index: 900;
                    transition: background 0.2s, transform 0.15s;
                }
                .back-to-top:hover {
                    background: var(--blm-brown);
                    transform: scale(1.08);
                }
                /* Scroll-margin for anchor targets */
                h2[id] {
                    scroll-margin-top: 16px;
                }
                @media print {
                    .report-toc { break-inside: avoid; }
                    .back-to-top { display: none !important; }
                    .a11y-widget { display: none !important; }
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
            <div class="cv-filter-wrap">
            <div class="report-header">
                <div class="agency-name">U.S. Department of the Interior &bull; Bureau of Land Management</div>
                <h1>${safeTitle}</h1>
                <div class="meta">Report Generated: ${escapeHtml(createdAt || "")}</div>
            </div>
            <div class="wrap">
                <div class="actions">
                    <a class="btn" href="javascript:window.print()">&#128424; Print / Save as PDF</a>
                    <button class="btn btn-bookmark" id="bookmarkReportBtn" title="Bookmark this report — reopens on this device/browser only">&#128278; Bookmark Report</button>
                </div>
                <div class="hint">Use your browser's print dialog to save as PDF. The bookmark saves a link to reopen this report on this device/browser only (valid for 7 days).</div>

                <!-- Table of Contents -->
                <nav class="report-toc" aria-label="Report sections">
                    <h2 class="toc-heading">Table of Contents</h2>
                    <ol class="toc-list">
                        <li><a href="#section-summary">Report Summary</a></li>
                        <li><a href="#section-findings">Regulatory Screening &amp; Findings Summary</a></li>
                        <li><a href="#section-aoi">Area of Interest</a></li>
                        <li><a href="#section-layers">Layer Analysis Maps</a>
                            <ul class="toc-sublist">
                                <li><a href="#bucket-land-status">\ud83c\udfdb\ufe0f Land Status &amp; Authority</a></li>
                                <li><a href="#bucket-land-use">\ud83d\udcd1 Land Use Plans &amp; Allocations</a></li>
                                <li><a href="#bucket-special">\u2b50 Special Designations</a></li>
                                <li><a href="#bucket-environmental">\ud83c\udf3f Environmental &amp; ESA</a></li>
                                <li><a href="#bucket-authorizations">\ud83d\udcdd Existing Authorizations</a></li>
                            </ul>
                        </li>
                        <li><a href="#section-sources">Data Sources</a></li>
                    </ol>
                </nav>

                <h2 id="section-summary"><span class="section-num">1.</span> Report Summary</h2>
                <p class="section-intro">A high-level overview of the screening analysis, including the number of datasets examined and features identified within the project area.</p>
                <div class="totals">
                ${totalsHtml || ""}
                </div>

                ${findingsSummaryHtml ? '<h2 id="section-findings"><span class="section-num">2.</span> Regulatory Screening &amp; Findings Summary</h2><p class="section-intro">A narrative summary of the regulatory framework, special designations, environmental factors, land use plans, and existing authorizations identified within the project area. Sub-sections below are shown only when relevant datasets intersect the area of interest.</p><div class="findings-summary">' + findingsSummaryHtml + '</div>' : ''}

                ${aoiSectionHtml || ""}

                <h2 id="section-layers"><span class="section-num">4.</span> Layer Analysis Maps</h2>
                <p class="section-intro">Interactive maps and attribute tables for each geospatial layer that was queried. Use the show/hide controls to focus on layers of interest, and click column headers to sort attribute data.</p>
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

            // ── Back-to-top button show/hide ──
            (function() {
                var btn = document.getElementById('backToTop');
                if (!btn) return;
                btn.addEventListener('click', function(e) {
                    e.preventDefault();
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                });
                window.addEventListener('scroll', function() {
                    btn.style.display = window.scrollY > 600 ? 'flex' : 'none';
                }, { passive: true });
            })();

            // ── Smooth-scroll for TOC anchor links ──
            (function() {
                var links = document.querySelectorAll('.report-toc a[href^="#"]');
                for (var i = 0; i < links.length; i++) {
                    links[i].addEventListener('click', function(e) {
                        e.preventDefault();
                        var target = document.querySelector(this.getAttribute('href'));
                        if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    });
                }
            })();

            // ── Bookmark Report button ──
            (function() {
                var btn = document.getElementById('bookmarkReportBtn');
                if (!btn) return;
                btn.addEventListener('click', function() {
                    var meta = document.querySelector('meta[name="report-id"]');
                    var reportId = meta ? meta.getAttribute('content') : null;
                    if (!reportId) { alert('Report ID not found. Re-run the analysis.'); return; }
                    var url = window.location.origin + window.location.pathname + '?report=' + reportId;
                    try {
                        navigator.clipboard.writeText(url).then(function() {
                            var orig = btn.textContent;
                            btn.textContent = '\u2705 Saved to Clipboard';
                            setTimeout(function() { btn.textContent = orig; }, 2000);
                        });
                    } catch(e) {
                        window.prompt('Copy this bookmark URL (works on this device only):', url);
                    }
                });
            })();
            </script>
            </div>
            <a href="#" class="back-to-top" id="backToTop" title="Back to top" aria-label="Back to top">&#8679; Top</a>

            ${getA11yWidgetBlock()}

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

    // Helper: look up service status, falling back to parent service URL
    // (serviceStatus is keyed by config-level URL, but layers may be sublayer URLs)
    function lookupServiceStatus(url) {
        if (!url) return "UNKNOWN";
        const direct = S.serviceStatus.get(url);
        if (direct) return direct;
        // Try stripping trailing /N sublayer index
        const parentUrl = url.replace(/\/\d+\/?$/, "");
        if (parentUrl !== url) {
            const parentStatus = S.serviceStatus.get(parentUrl);
            if (parentStatus) return parentStatus;
        }
        return "UNKNOWN";
    }

    /**
     * Build a data sources table scoped to the layers used in a specific report.
     * Each layer gets a row with Name, URL, Features in AOI, UP/Down,
     * followed by a description row.
     */
    async function buildLayerSourcesTable(layers) {
        if (!layers || !layers.length) return "";

        // Fetch descriptions in parallel (with timeout) — reuse cached ones
        const descFetchResults = await Promise.allSettled(
            layers.map(async item => {
                const url = item.url;
                if (!url) return { url: "", desc: "" };

                const cached = S.serviceStatus.get(url + "::desc");
                if (cached) return { url, desc: cached };

                try {
                    const pjsonUrl = normalizePjsonUrl(url);
                    const pjson = await fetchJsonWithTimeout(pjsonUrl, 5000);
                    const desc = pickServiceDescription(pjson) || "";
                    if (desc) S.serviceStatus.set(url + "::desc", desc);
                    // Service responded from the browser — mark it UP regardless of R2 cache
                    S.serviceStatus.set(url, "UP");
                    return { url, desc };
                } catch (e) {
                    return { url, desc: "" };
                }
            })
        );

        const descByUrl = new Map();
        for (const r of descFetchResults) {
            if (r.status === "fulfilled" && r.value) {
                descByUrl.set(r.value.url, r.value.desc);
            }
        }

        const rows = layers.map(item => {
            const url = item.url || "";
            const status = lookupServiceStatus(url);
            const statusClass = status === "UP" ? "status-up" : (status === "UNKNOWN" ? "status-unknown" : "status-down");
            const desc = descByUrl.get(url) || S.serviceStatus.get(url + "::desc") || "";
            const featCount = item.count || 0;

            const countDisplay = featCount > 0
                ? `<b>${escapeHtml(String(featCount))}</b>`
                : '<span class="feat-count-zero">0</span>';

            const descRow = desc
                ? `<tr class="desc-row"><td colspan="4">${escapeHtml(desc)}</td></tr>`
                : `<tr class="desc-row"><td colspan="4" style="opacity:0.5;">(No description available)</td></tr>`;

            return `
                <tr>
                    <td class="service-name-col">${escapeHtml(item.title || "Unknown")}</td>
                    <td class="service-url-col"><a href="${escapeHtml(url)}" target="_blank" rel="noopener">${escapeHtml(url)}</a></td>
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
                    The following geospatial web services were queried for this report. Service availability was verified at the time of report generation.
                </p>
                <table class="data-sources-table">
                    <thead>
                        <tr>
                            <th style="width: 30%;">Layer Name</th>
                            <th style="width: 35%;">Web Service URL</th>
                            <th style="width: 15%;">Features in AOI</th>
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
            const status = lookupServiceStatus(svc.url);
            const statusClass = status === "UP" ? "status-up" : (status === "UNKNOWN" ? "status-unknown" : "status-down");
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
                <h2 id="section-sources"><span class="section-num">5.</span> Data Sources</h2>
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

    // ── Cached state/county boundary layers for AOI maps ──
    let _stateBoundaryLayer = null;
    let _countyBoundaryLayer = null;

    function _getStateBoundaryLayer() {
        if (!_stateBoundaryLayer) {
            _stateBoundaryLayer = new FeatureLayer({
                url: S.config?.referenceLayers?.usaStatesGeneralized || "https://services.arcgis.com/P3ePLMYs2RVChkJx/arcgis/rest/services/USA_States_Generalized_Boundaries/FeatureServer/0",
                title: "__reportStateBoundaries",
                outFields: [],
                labelsVisible: true,
                labelingInfo: [{
                    labelExpressionInfo: { expression: "$feature.STATE_ABBR" },
                    symbol: {
                        type: "text",
                        color: [60, 60, 60, 0.9],
                        haloColor: [255, 255, 255, 0.85],
                        haloSize: 1.5,
                        font: { size: 11, weight: "bold", family: "Noto Sans" }
                    },
                    minScale: 25000000,
                    maxScale: 0
                }],
                renderer: {
                    type: "simple",
                    symbol: {
                        type: "simple-fill",
                        color: [0, 0, 0, 0],
                        outline: { color: [80, 80, 80, 0.7], width: 1.5 }
                    }
                },
                visible: false
            });
        }
        return _stateBoundaryLayer;
    }

    function _getCountyBoundaryLayer() {
        if (!_countyBoundaryLayer) {
            _countyBoundaryLayer = new FeatureLayer({
                url: S.config?.referenceLayers?.usaCountiesGeneralized || "https://services.arcgis.com/P3ePLMYs2RVChkJx/arcgis/rest/services/USA_Counties_Generalized_Boundaries/FeatureServer/0",
                title: "__reportCountyBoundaries",
                outFields: [],
                labelsVisible: true,
                labelingInfo: [{
                    labelExpressionInfo: { expression: "$feature.NAME" },
                    symbol: {
                        type: "text",
                        color: [90, 90, 90, 0.85],
                        haloColor: [255, 255, 255, 0.8],
                        haloSize: 1,
                        font: { size: 9, weight: "normal", family: "Noto Sans" }
                    },
                    minScale: 3000000,
                    maxScale: 0
                }],
                renderer: {
                    type: "simple",
                    symbol: {
                        type: "simple-fill",
                        color: [0, 0, 0, 0],
                        outline: { color: [140, 140, 140, 0.5], width: 0.75 }
                    }
                },
                visible: false
            });
        }
        return _countyBoundaryLayer;
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

        const { ensureAoiOnTop, hideAoiMask, captureScreenshotWithWait, waitForTabVisible, waitForLayerReadyToCapture } = mapUtils;

        const width  = config?.visualReport?.screenshotWidth ?? 1400;
        const height = Math.round(width * 0.5625);
        const maps   = [];

        const insetFrac = 0.22;
        const insetW = Math.round(width * insetFrac);
        const insetH = Math.round(height * insetFrac);
        const insetMargin = 12;
        const overviewZoomFactor = 8;

        // Add state/county boundary layers to map if not already present
        const stateLayer  = _getStateBoundaryLayer();
        const countyLayer = _getCountyBoundaryLayer();
        if (!view.map.layers.includes(stateLayer))  view.map.layers.add(stateLayer);
        if (!view.map.layers.includes(countyLayer)) view.map.layers.add(countyLayer);

        // Load them (no-op if already loaded)
        try { await stateLayer.load(); }  catch (_) { }
        try { await countyLayer.load(); } catch (_) { }

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
                if (l === stateLayer || l === countyLayer) { l.visible = true; continue; }
                if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }
                l.visible = false;
            }
            // Ensure boundary layers draw below AOI
            ensureAoiOnTop();
        }

        function restoreVisibility() {
            // Restore township renderer
            if (plssTownshipLayer && plssTownshipOrigRenderer !== null) {
                try { plssTownshipLayer.renderer = plssTownshipOrigRenderer; } catch (e) { }
            }
            // Hide boundary layers (they are only for report screenshots)
            try { stateLayer.visible = false; }  catch (_) { }
            try { countyLayer.visible = false; } catch (_) { }
            visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) { } });
            hideAoiMask();
            ensureAoiOnTop();
        }

        async function compositeWithOverview(mainDataUrl, mainExtent, scale) {
            const overviewScale = scale * overviewZoomFactor;
            await view.goTo({ target: selectionGeom.extent, scale: overviewScale }, { animate: false });

            // Wait for boundary layers to finish rendering at the new scale
            await waitForLayerReadyToCapture(stateLayer, view, { timeoutMs: 5000 });
            await waitForLayerReadyToCapture(countyLayer, view, { timeoutMs: 5000 });

            const ovSs = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });
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

            // Wait for state/county boundary layers to fully load and render
            await waitForLayerReadyToCapture(stateLayer, view, { timeoutMs: 8000 });
            await waitForLayerReadyToCapture(countyLayer, view, { timeoutMs: 8000 });

            const ext1 = selectionGeom.extent;
            await view.goTo({ target: ext1, scale: 900000 }, { animate: false });

            // Wait for boundary layers at new scale
            await waitForLayerReadyToCapture(stateLayer, view, { timeoutMs: 5000 });
            await waitForLayerReadyToCapture(countyLayer, view, { timeoutMs: 5000 });

            const ss1 = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });
            const mainExtent1 = view.extent.clone();

            if (ss1) {
                const composited1 = await compositeWithOverview(ss1, mainExtent1, 900000);
                maps.push(`<div class="aoi-map"><img src="${composited1}" alt="AOI Context (Regional 1:900,000)" /></div>`);
            }

            const ext2 = selectionGeom.extent;
            await view.goTo({ target: ext2, scale: 200000 }, { animate: false });

            // Wait for boundary layers at new scale
            await waitForLayerReadyToCapture(stateLayer, view, { timeoutMs: 5000 });
            await waitForLayerReadyToCapture(countyLayer, view, { timeoutMs: 5000 });

            const ss2 = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });
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

    // ────────────────────────────────────────────────────────────────
    // Progressive Report Builder
    // Opens a new tab immediately with a shell HTML, streams content
    // as it's generated so users see progress in real-time.
    // ────────────────────────────────────────────────────────────────

    /**
     * Returns the SVG color-vision filters, accessibility widget HTML,
     * and the inline JS that wires it up. Designed to be injected just
     * before </body> in every report template.
     * Applies cv-* classes to `.cv-filter-wrap` so the fixed-position
     * widget stays viewport-pinned (CSS filter on body breaks position:fixed).
     */
    function getA11yWidgetBlock() {
        return `
    <!-- Accessibility: SVG color-vision filters -->
    <svg xmlns="http://www.w3.org/2000/svg" style="position:absolute;width:0;height:0;overflow:hidden" aria-hidden="true">
      <defs>
        <filter id="cv-protanopia">
          <feColorMatrix type="matrix" values="0.567,0.433,0,0,0 0.558,0.442,0,0,0 0,0.242,0.758,0,0 0,0,0,1,0"/>
        </filter>
        <filter id="cv-deuteranopia">
          <feColorMatrix type="matrix" values="0.625,0.375,0,0,0 0.7,0.3,0,0,0 0,0.3,0.7,0,0 0,0,0,1,0"/>
        </filter>
        <filter id="cv-tritanopia">
          <feColorMatrix type="matrix" values="0.95,0.05,0,0,0 0,0.433,0.567,0,0 0,0.475,0.525,0,0 0,0,0,1,0"/>
        </filter>
        <filter id="cv-achromatopsia">
          <feColorMatrix type="matrix" values="0.299,0.587,0.114,0,0 0.299,0.587,0.114,0,0 0.299,0.587,0.114,0,0 0,0,0,1,0"/>
        </filter>
        <filter id="cv-highcontrast">
          <feComponentTransfer>
            <feFuncR type="linear" slope="1.8" intercept="-0.35"/>
            <feFuncG type="linear" slope="1.8" intercept="-0.35"/>
            <feFuncB type="linear" slope="1.8" intercept="-0.35"/>
          </feComponentTransfer>
        </filter>
      </defs>
    </svg>
    <!-- Accessibility floating widget -->
    <div id="a11yWidgetReport" class="a11y-widget" role="region" aria-label="Accessibility options">
      <button id="a11yToggleBtnReport" class="a11y-toggle" aria-haspopup="true" aria-expanded="false" title="Vision Assistance">
        <svg viewBox="0 0 24 24" width="22" height="22" fill="currentColor" aria-hidden="true">
          <circle cx="12" cy="4.5" r="2"/>
          <path d="M12 7.5c-3.5 0-6 1-6 1v2s2.5-.5 5-.7V12l-3.5 7h2.2l2.3-5 2.3 5h2.2L13 12v-2.2c2.5.2 5 .7 5 .7v-2s-2.5-1-6-1z"/>
        </svg>
      </button>
      <div id="a11yMenuReport" class="a11y-menu hidden" role="menu" aria-label="Color vision modes">
        <div class="a11y-menu-header">Vision Assistance</div>
        <button class="a11y-option" role="menuitem" data-cv="none">Normal Vision</button>
        <button class="a11y-option" role="menuitem" data-cv="protanopia">Protanopia <small>(no red)</small></button>
        <button class="a11y-option" role="menuitem" data-cv="deuteranopia">Deuteranopia <small>(no green)</small></button>
        <button class="a11y-option" role="menuitem" data-cv="tritanopia">Tritanopia <small>(no blue)</small></button>
        <button class="a11y-option" role="menuitem" data-cv="achromatopsia">Achromatopsia <small>(grayscale)</small></button>
        <button class="a11y-option" role="menuitem" data-cv="highcontrast">High Contrast</button>
      </div>
    </div>
    <script>
    (function() {
        var STORAGE_KEY = 'a11y-cv-mode';
        var toggleBtn = document.getElementById('a11yToggleBtnReport');
        var menu = document.getElementById('a11yMenuReport');
        if (!toggleBtn || !menu) return;
        var options = menu.querySelectorAll('.a11y-option');
        var CV_CLASSES = ['cv-protanopia','cv-deuteranopia','cv-tritanopia','cv-achromatopsia','cv-highcontrast'];
        var filterWrap = document.querySelector('.cv-filter-wrap');
        function applyMode(mode) {
            if (!filterWrap) return;
            CV_CLASSES.forEach(function(c) { filterWrap.classList.remove(c); });
            if (mode && mode !== 'none') filterWrap.classList.add('cv-' + mode);
            options.forEach(function(b) { b.setAttribute('aria-checked', b.getAttribute('data-cv') === mode ? 'true' : 'false'); });
            try { localStorage.setItem(STORAGE_KEY, mode || 'none'); } catch(_) {}
        }
        var saved = 'none';
        try { saved = localStorage.getItem(STORAGE_KEY) || 'none'; } catch(_) {}
        applyMode(saved);
        toggleBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            var isHidden = menu.classList.toggle('hidden');
            toggleBtn.setAttribute('aria-expanded', isHidden ? 'false' : 'true');
            if (!isHidden) { var first = menu.querySelector('.a11y-option'); if (first) first.focus(); }
        });
        // Touch support for iOS Safari
        toggleBtn.addEventListener('touchend', function(e) {
            e.preventDefault();
            toggleBtn.click();
        });
        options.forEach(function(b) {
            b.addEventListener('click', function(e) {
                e.stopPropagation();
                applyMode(b.getAttribute('data-cv'));
                menu.classList.add('hidden');
                toggleBtn.setAttribute('aria-expanded', 'false');
            });
        });
        menu.addEventListener('keydown', function(e) {
            var items = Array.prototype.slice.call(options);
            var idx = items.indexOf(document.activeElement);
            if (e.key === 'ArrowDown' || e.key === 'ArrowRight') { e.preventDefault(); items[(idx + 1) % items.length].focus(); }
            else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') { e.preventDefault(); items[(idx - 1 + items.length) % items.length].focus(); }
            else if (e.key === 'Escape') { menu.classList.add('hidden'); toggleBtn.setAttribute('aria-expanded','false'); toggleBtn.focus(); }
        });
        document.addEventListener('click', function(e) {
            var widget = document.getElementById('a11yWidgetReport');
            if (!menu.classList.contains('hidden') && widget && !widget.contains(e.target)) {
                menu.classList.add('hidden');
                toggleBtn.setAttribute('aria-expanded','false');
            }
        });
    })();
    </script>`;
    }

    /**
     * Get the CSS styles for the progressive report (shared with regular report)
     */
    function getReportStyles() {
        return `
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
            .progress-banner {
                background: linear-gradient(90deg, var(--blm-green) 0%, var(--blm-green-light) 100%);
                color: white;
                padding: 16px 24px;
                border-radius: 8px;
                margin-bottom: 24px;
                display: flex;
                align-items: center;
                gap: 16px;
            }
            .progress-banner.hidden { display: none; }
            .progress-spinner {
                width: 24px;
                height: 24px;
                border: 3px solid rgba(255,255,255,0.3);
                border-top-color: white;
                border-radius: 50%;
                animation: spin 1s linear infinite;
            }
            @keyframes spin { to { transform: rotate(360deg); } }
            .progress-text { flex: 1; font-weight: 500; min-width: 0; }
            .progress-status { font-size: 13px; opacity: 0.85; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
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
            .map-placeholder {
                background: linear-gradient(135deg, #f0f0f0 0%, #e0e0e0 100%);
                padding: 60px 24px;
                text-align: center;
                color: var(--muted);
                font-style: italic;
                border-radius: 6px;
                margin: 16px 0;
            }
            .map-placeholder .spinner-small {
                width: 32px;
                height: 32px;
                border: 3px solid #ccc;
                border-top-color: var(--blm-green);
                border-radius: 50%;
                animation: spin 1s linear infinite;
                margin: 0 auto 12px;
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
            .report-footer{
                margin-top: 48px;
                padding-top: 24px;
                border-top: 2px solid var(--border);
                font-size: 12px;
                color: var(--muted);
                text-align: center;
            }
            .pill{ 
                border: 1px solid var(--border); 
                border-radius: 6px; 
                padding: 10px 16px; 
                font-size: 13px; 
                font-weight: 600;
                background: var(--white);
                box-shadow: 0 1px 3px rgba(0,0,0,0.06);
            }
            .totals{ margin-top: 24px; }
            .totals .row{ display:flex; gap:14px; flex-wrap:wrap; margin-top:12px; }
            .export-btn{
                display: inline-flex;
                align-items: center;
                gap: 8px;
                background: var(--blm-green);
                color: #fff;
                border: none;
                border-radius: 6px;
                padding: 10px 18px;
                font-size: 14px;
                font-weight: 600;
                cursor: pointer;
                transition: background 0.2s;
            }
            .export-btn:hover{ background: var(--blm-green-light); }
            .report-actions{
                display: flex;
                gap: 12px;
                margin-top: 16px;
                padding-top: 16px;
                border-top: 1px solid rgba(255,255,255,0.2);
            }
            /* Layer narrative paragraph */
            .layer-narrative {
                margin: 16px 0;
                padding: 14px 20px;
                background: var(--blm-tan);
                border-left: 4px solid var(--blm-green);
                border-radius: 4px;
                font-size: 13.5px;
                line-height: 1.65;
                color: var(--text);
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
            /* Section collapse */
            .section-collapse-wrap {
                overflow: hidden;
                transition: max-height 0.3s ease;
            }
            .section.section-hidden .section-collapse-wrap {
                max-height: 0 !important;
            }
            .section-hide-btn {
                float: right;
                padding: 4px 10px;
                font-size: 12px;
                background: var(--blm-tan);
                border: 1px solid var(--border);
                border-radius: 4px;
                cursor: pointer;
                color: var(--muted);
                transition: background 0.15s;
            }
            .section-hide-btn:hover {
                background: #e0ddd4;
            }
            /* Map zoom controls */
            .map-zoom-controls {
                position: absolute;
                top: 8px;
                right: 8px;
                display: flex;
                flex-direction: column;
                gap: 4px;
                z-index: 10;
            }
            .map-zoom-controls button {
                width: 28px;
                height: 28px;
                border: 1px solid var(--border);
                background: var(--white);
                border-radius: 4px;
                font-size: 16px;
                font-weight: bold;
                cursor: pointer;
                color: var(--blm-green);
                transition: background 0.15s;
            }
            .map-zoom-controls button:hover {
                background: var(--blm-tan);
            }
            .map {
                position: relative;
            }
            /* Section collapse animation */
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
            /* Data sources table */
            table.data-sources-table {
                width: 100%;
                border-collapse: collapse;
                margin-top: 16px;
                font-size: 12px;
                table-layout: fixed;
                background: var(--white);
                border: 1px solid var(--border);
                border-radius: 6px;
                overflow: hidden;
            }
            table.data-sources-table th {
                background: var(--blm-green);
                color: var(--white);
                padding: 10px 14px;
                text-align: left;
                font-weight: 600;
                font-size: 11px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            table.data-sources-table td {
                padding: 10px 14px;
                border-bottom: 1px solid var(--border);
                vertical-align: top;
                word-wrap: break-word;
            }
            table.data-sources-table tr:last-child td { border-bottom: none; }
            table.data-sources-table .service-url-col {
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 10px;
                word-break: break-all;
                color: var(--muted);
            }
            table.data-sources-table .desc-row td {
                padding: 4px 14px 12px 14px;
                font-size: 11px;
                color: var(--muted);
                line-height: 1.5;
                border-bottom: 2px solid var(--border);
                background: var(--blm-tan);
                font-style: italic;
            }
            table.data-sources-table .feat-count { font-weight: 600; text-align: center; }
            table.data-sources-table .feat-count-zero { color: var(--muted); opacity: 0.6; }
            .status-up { color: #2e7d32; font-weight: 600; }
            .status-down { color: #c62828; font-weight: 600; }
            /* Accessibility Widget */
            .a11y-widget {
                position: fixed;
                bottom: 16px;
                right: 16px;
                z-index: 50000;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            }
            .a11y-toggle {
                width: 44px;
                height: 44px;
                border-radius: 50%;
                border: 2px solid #1a73e8;
                background: #fff;
                color: #1a73e8;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 2px 8px rgba(0,0,0,0.25);
                transition: background 0.2s, color 0.2s, transform 0.15s;
                -webkit-tap-highlight-color: transparent;
            }
            .a11y-toggle:hover,
            .a11y-toggle:focus-visible {
                background: #1a73e8;
                color: #fff;
                transform: scale(1.08);
                outline: 2px solid #fff;
                outline-offset: 2px;
            }
            .a11y-toggle[aria-expanded="true"] {
                background: #1a73e8;
                color: #fff;
            }
            .a11y-menu {
                position: absolute;
                bottom: 52px;
                right: 0;
                min-width: 220px;
                background: #fff;
                border: 1px solid #d0d0d0;
                border-radius: 10px;
                box-shadow: 0 4px 16px rgba(0,0,0,0.18);
                padding: 6px 0;
                animation: a11yFadeIn 0.15s ease;
            }
            .a11y-menu.hidden { display: none; }
            @keyframes a11yFadeIn {
                from { opacity: 0; transform: translateY(8px); }
                to   { opacity: 1; transform: translateY(0); }
            }
            .a11y-menu-header {
                padding: 8px 14px 6px;
                font-size: 11px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                color: #666;
                border-bottom: 1px solid #eee;
                margin-bottom: 4px;
            }
            .a11y-option {
                display: block;
                width: 100%;
                text-align: left;
                padding: 8px 14px;
                border: none;
                background: none;
                font-size: 13px;
                color: #333;
                cursor: pointer;
                transition: background 0.12s;
                -webkit-tap-highlight-color: transparent;
            }
            .a11y-option:hover,
            .a11y-option:focus-visible {
                background: #e8f0fe;
                outline: none;
            }
            .a11y-option[aria-checked="true"] {
                font-weight: 700;
                color: #1a73e8;
                background: #e8f0fe;
            }
            .a11y-option[aria-checked="true"]::before {
                content: '\\2713 ';
            }
            .a11y-option small { color: #888; font-weight: 400; }
            /* Color-vision filter classes — applied to .cv-filter-wrap so position:fixed widget stays viewport-pinned */
            .cv-filter-wrap.cv-protanopia    { -webkit-filter: url(#cv-protanopia); filter: url(#cv-protanopia); }
            .cv-filter-wrap.cv-deuteranopia  { -webkit-filter: url(#cv-deuteranopia); filter: url(#cv-deuteranopia); }
            .cv-filter-wrap.cv-tritanopia    { -webkit-filter: url(#cv-tritanopia); filter: url(#cv-tritanopia); }
            .cv-filter-wrap.cv-achromatopsia { -webkit-filter: url(#cv-achromatopsia); filter: url(#cv-achromatopsia); }
            .cv-filter-wrap.cv-highcontrast  { -webkit-filter: url(#cv-highcontrast); filter: url(#cv-highcontrast); }
            @media print {
                .report-actions { display: none; }
                .export-btn { display: none; }
                .a11y-widget { display: none !important; }
                .interactive-table-wrapper .table-toolbar { display: none !important; }
                .hidden-cols-bar { display: none !important; }
                .col-hide-btn { display: none !important; }
                .section-hide-btn { display: none !important; }
                .map-zoom-controls { display: none !important; }
                .interactive-table { font-size: 9px; }
                .interactive-table th {
                    background: var(--blm-green) !important;
                    -webkit-print-color-adjust: exact;
                    print-color-adjust: exact;
                }
            }
        `;
    }

    /**
     * Open a progressive report window and return an interface for streaming content
     * Returns a Promise that resolves when the window is ready
     */
    async function openProgressiveReport(options = {}) {
        const title = options.title || "Report";
        const bucketLabel = options.bucketLabel || null;
        const createdAt = formatDateTimeForReport(new Date());

        // Create initial shell HTML
        const shellHtml = `<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width,initial-scale=1"/>
    <title>${escapeHtml(title)}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Merriweather:wght@400;700&family=Source+Sans+Pro:wght@400;600;700&display=swap" rel="stylesheet">
    <style>${getReportStyles()}</style>
</head>
<body>
    <div class="cv-filter-wrap">
    <header class="report-header">
        <div class="agency-name">U.S. Department of the Interior &bull; Bureau of Land Management</div>
        <h1>${escapeHtml(title)}</h1>
        <p class="meta">Generated: ${escapeHtml(createdAt)}${bucketLabel ? ` &bull; Category: ${escapeHtml(bucketLabel)}` : ''}</p>
    </header>
    <main class="wrap">
        <div class="progress-banner" id="progressBanner">
            <div class="progress-spinner"></div>
            <div class="progress-text">
                <div id="progressTitle">Building report...</div>
                <div class="progress-status" id="progressStatus">Initializing...</div>
            </div>
        </div>
        <div id="reportContent">
            <!-- Content will be streamed here -->
        </div>
    </main>
    <script>
        // Helper to smoothly scroll to bottom as content is added
        let autoScroll = true;
        document.addEventListener('click', () => { autoScroll = false; });
        window.scrollToBottom = function() {
            if (autoScroll) window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
        };
        window.hideProgress = function() {
            const banner = document.getElementById('progressBanner');
            if (banner) banner.classList.add('hidden');
        };
        window.updateProgress = function(title, status) {
            const titleEl = document.getElementById('progressTitle');
            const statusEl = document.getElementById('progressStatus');
            if (titleEl) titleEl.textContent = title;
            if (statusEl) statusEl.textContent = status;
        };
        
        // Toggle individual layer section visibility
        function toggleSection(btn) {
            var section = btn.closest('.section');
            if (!section) return;
            var isHidden = section.classList.toggle('section-hidden');
            btn.innerHTML = isHidden ? '&#x2713; Show' : '&#x2715; Hide';
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
                if (arrow) { arrow.textContent = '\\u21C5'; }
            }
            th.setAttribute('data-sort-dir', newSort);
            var sortArrow = th.querySelector('.sort-arrow');
            if (sortArrow) sortArrow.textContent = newSort === 'asc' ? '\\u25B2' : '\\u25BC';
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
            var elements = wrapper.querySelectorAll('[data-col="' + colIdx + '"]');
            for (var ei = 0; ei < elements.length; ei++) {
                elements[ei].style.display = 'none';
            }
            var th = wrapper.querySelector('thead th[data-col="' + colIdx + '"]');
            var label = th ? (th.getAttribute('data-label') || 'Column ' + colIdx) : 'Column ' + colIdx;
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
        }
        
        // Interactive table: restore a hidden column via pill click
        function showColumn(wrapperId, colIdx) {
            var wrapper = document.getElementById(wrapperId);
            if (!wrapper) return;
            var elements = wrapper.querySelectorAll('[data-col="' + colIdx + '"]');
            for (var ei = 0; ei < elements.length; ei++) {
                elements[ei].style.display = '';
            }
            var bar = wrapper.querySelector('.hidden-cols-bar');
            if (bar) {
                var pills = bar.querySelectorAll('.hidden-col-pill[data-col="' + colIdx + '"]');
                for (var pi = 0; pi < pills.length; pi++) pills[pi].remove();
                if (!bar.querySelector('.hidden-col-pill')) bar.style.display = 'none';
            }
        }
        
        // Auto-hide columns that are entirely null/empty
        window.autoHideNullColumns = function() {
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
        };

        // Signal that the page is ready
        window.reportReady = true;
    </script>
    </div>
${getA11yWidgetBlock()}
</body>
</html>`;

        // Open the window
        const blob = new Blob([shellHtml], { type: "text/html;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        // Note: Cannot use "noopener" because we need to access win.document to stream content
        const win = window.open(url, "_blank");
        
        // Clean up blob URL after a delay
        window.setTimeout(() => URL.revokeObjectURL(url), 60000);

        if (!win) {
            console.error("Failed to open report window - popup blocked?");
            return null;
        }

        // Wait for the popup document to be ready (poll for reportReady flag)
        const maxWaitMs = 10000;
        const startTime = Date.now();
        while (Date.now() - startTime < maxWaitMs) {
            try {
                if (win.reportReady && win.document.getElementById('reportContent')) {
                    break;
                }
            } catch (e) {
                // Cross-origin or window not ready
            }
            await new Promise(r => setTimeout(r, 100));
        }

        // Final check
        try {
            if (!win.document.getElementById('reportContent')) {
                console.error("Report window did not initialize properly");
                return null;
            }
        } catch (e) {
            console.error("Cannot access report window:", e);
            return null;
        }

        // Return interface for streaming content
        return {
            win,
            
            /**
             * Append HTML content to the report
             */
            appendContent(html) {
                try {
                    const container = win.document.getElementById('reportContent');
                    if (container) {
                        container.insertAdjacentHTML('beforeend', html);
                        win.scrollToBottom?.();
                    }
                } catch (e) {
                    console.warn("Failed to append content to report:", e);
                }
            },

            /**
             * Update progress banner
             */
            updateProgress(title, status) {
                try {
                    win.updateProgress?.(title, status);
                } catch (e) {
                    // Window may be closed
                }
            },

            /**
             * Hide progress banner (report complete)
             */
            hideProgress() {
                try {
                    win.hideProgress?.();
                } catch (e) {
                    // Window may be closed
                }
            },

            /**
             * Check if window is still open
             */
            isOpen() {
                try {
                    return win && !win.closed;
                } catch (e) {
                    return false;
                }
            },

            /**
             * Add a section with a map placeholder that will be replaced
             */
            addMapPlaceholder(layerTitle, placeholderId) {
                const html = `
                    <div class="section" id="section-${escapeHtml(placeholderId)}">
                        <h3>${escapeHtml(layerTitle)}</h3>
                        <div class="map-placeholder" id="placeholder-${escapeHtml(placeholderId)}">
                            <div class="spinner-small"></div>
                            <div>Generating map...</div>
                        </div>
                    </div>
                `;
                this.appendContent(html);
            },

            /**
             * Replace a map placeholder with actual content
             */
            replaceMapPlaceholder(placeholderId, html) {
                try {
                    const placeholder = win.document.getElementById(`placeholder-${placeholderId}`);
                    if (placeholder) {
                        placeholder.outerHTML = html;
                    }
                } catch (e) {
                    console.warn("Failed to replace placeholder:", e);
                }
            },

            /**
             * Update a section's content by ID
             */
            updateSection(sectionId, html) {
                try {
                    const section = win.document.getElementById(sectionId);
                    if (section) {
                        section.innerHTML = html;
                    }
                } catch (e) {
                    console.warn("Failed to update section:", e);
                }
            },

            /**
             * Add footer when report is complete
             */
            addFooter() {
                const footerHtml = `
                    <footer class="report-footer">
                        <p>This report is for informational purposes only and does not constitute a formal BLM determination.</p>
                        <p>Generated by the BLM Permit Screening Tool &bull; ${formatDateTimeForReport(new Date())}</p>
                    </footer>
                `;
                this.appendContent(footerHtml);
            }
        };
    }

    /**
     * Build a progressive report for a specific bucket or all buckets
     * @param {Object} options
     * @param {string} options.bucketKey - e.g., "land-status", "environmental", or null for full report
     * @param {Function} options.onProgress - callback for progress updates
     */
    async function buildProgressiveReport(options = {}) {
        const bucketKey = options.bucketKey || null;
        const onProgress = options.onProgress || (() => {});

        const view = S.view;
        const selectionGeom = S.selectionGeom;
        const config = S.config;
        const lastReportRowsByLayer = S.lastReportRowsByLayer;
        const aoiLayer = S.aoiLayer;
        const aoiMaskLayer = S.aoiMaskLayer;
        const alwaysVisibleLayers = S.alwaysVisibleLayers;

        if (!view || !selectionGeom) {
            console.error("Cannot build report: no AOI selected");
            return;
        }

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            console.error("Cannot build report: no analysis data available");
            return;
        }

        // Get bucket info
        const bucketInfo = bucketKey ? REPORT_BUCKETS[bucketKey] : null;
        const reportTitle = bucketInfo 
            ? `${bucketInfo.icon} ${bucketInfo.label} Report`
            : "Land & Resource Intersection Analysis Report";

        // Filter layers to this bucket (or all for full report)
        let targetLayers;
        if (bucketKey) {
            const buckets = categorizeLayersIntoBuckets(lastReportRowsByLayer);
            targetLayers = buckets[bucketKey] || [];
        } else {
            targetLayers = lastReportRowsByLayer;
        }

        // Filter to layers with coverage that can generate maps
        // Feature counts are queried during report generation (deferred from screening)
        const mappableLayers = targetLayers
            .filter(x => x?.hasCoverage || (x?.count || 0) > 0)
            .filter(x => x?.url || x?.__isImageService)
            .filter(x => !(x.title && x.title.toLowerCase().includes("state boundaries")));

        // Open the progressive report window
        const report = await openProgressiveReport({
            title: reportTitle,
            bucketLabel: bucketInfo?.label || null
        });

        if (!report) {
            alert("Could not open report window. Please allow popups for this site.");
            return;
        }

        const {
            updateAoiMask, hideAoiMask, ensureAoiOnTop,
            waitForViewStationary, waitForLayerReadyToCapture,
            captureScreenshotWithWait, waitForTabVisible,
            acquireWakeLock, releaseWakeLock,
            getLayerGeometryType, makeRendererOpaque, getPresetRenderer,
            thickenLayerSymbology
        } = mapUtils;

        const {
            queryAllFeaturesPaged, querySingleLayer, computeLayerCoverageStats, computeElevationStats, SQM_PER_ACRE
        } = queryEngine;

        try {
            await acquireWakeLock();

            // === STEP 1: AOI Summary Section ===
            report.updateProgress("Building report...", "Generating AOI overview maps");
            onProgress("Generating AOI maps...", 10);

            // Calculate AOI info
            let aoiAcres = 0;
            try {
                const aoiSqm = Math.max(0, geometryEngine.geodesicArea(selectionGeom, "square-meters"));
                aoiAcres = aoiSqm / SQM_PER_ACRE;
            } catch (e) {
                aoiAcres = 0;
            }

            let aoiMethod = "Manually Drawn";
            if (S.aoiSource === "select") {
                const tool = plssToolLabel(S.aoiSourcePlssTool);
                aoiMethod = `Selected ${tool}`;
            } else if (S.aoiSource === "upload") {
                aoiMethod = `Uploaded File: ${S.aoiSourceLayerTitle || "unknown"}`;
            }

            // Generate AOI maps
            const aoiMapsHtml = await generateAoiMapsWithCircles();

            // Build AOI section
            const aoiSectionHtml = `
                <h2>Area of Interest</h2>
                <p style="color: var(--muted); font-style: italic;">The geographic boundary used for this analysis.</p>
                ${aoiMapsHtml}
                <div class="aoi-details">
                    <div class="aoi-field"><span class="aoi-label">Area:</span> ${formatNumber(aoiAcres, 2)} acres</div>
                    <div class="aoi-field"><span class="aoi-label">Method:</span> ${escapeHtml(aoiMethod)}</div>
                </div>
            `;
            report.appendContent(aoiSectionHtml);

            if (!report.isOpen()) return; // User closed window

            // === STEP 2: Summary stats placeholder (updated after layer queries) ===
            const layersInAoiEstimate = targetLayers.filter(x => x.hasCoverage).length;

            const summaryPlaceholderHtml = `
                <h2>Summary</h2>
                <div class="totals" id="summary-stats-pills">
                    <div class="row">
                        <div class="pill">${targetLayers.length} Layers Queried</div>
                        <div class="pill">${layersInAoiEstimate} Layers in AOI</div>
                        <div class="pill">… Features in AOI</div>
                    </div>
                </div>
            `;
            report.appendContent(summaryPlaceholderHtml);

            if (!report.isOpen()) return;

            // === STEP 3: Generate maps for each layer progressively ===
            if (mappableLayers.length === 0) {
                report.appendContent(`
                    <h2>Layer Details</h2>
                    <p style="color: var(--muted); font-style: italic;">No intersecting features found in the ${bucketInfo?.label || 'screened'} datasets.</p>
                `);
            } else {
                report.appendContent(`<h2>Layer Details</h2>`);

                const paddingFactor = config?.visualReport?.paddingFactor ?? 1.12;
                const width = config?.visualReport?.screenshotWidth ?? 1400;

                let fixedExtent = null;
                const ext = selectionGeom?.extent;
                if (ext && ext.expand) fixedExtent = ext.expand(paddingFactor);

                // Snapshot layer visibility
                const allLayers = view.map.layers.toArray();
                const visSnapshot = allLayers.map(l => ({ layer: l, visible: l.visible }));
                const originalBasemap = view.map.basemap;
                const imageryBasemapId = config?.map?.imageryBasemap || "satellite";

                function setVisibilityForScreenshot(tempLayer) {
                    for (const l of allLayers) {
                        if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }
                        if (aoiMaskLayer && l === aoiMaskLayer) { l.visible = true; continue; }
                        if (l?.type === "tile") { l.visible = true; continue; }
                        if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }
                        l.visible = false;
                    }
                    if (tempLayer) tempLayer.visible = true;
                    updateAoiMask(true);
                    ensureAoiOnTop();
                }

                function restoreVisibility() {
                    visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) {} });
                    hideAoiMask();
                    ensureAoiOnTop();
                }

                try {
                    // Switch to imagery basemap
                    view.map.basemap = imageryBasemapId;
                    await new Promise(r => setTimeout(r, 500));
                    await waitForViewStationary(800);

                    if (fixedExtent) {
                        await view.goTo(fixedExtent, { animate: false });
                        await waitForViewStationary(800);
                    }

                    // Process each layer
                    for (let i = 0; i < mappableLayers.length; i++) {
                        if (!report.isOpen()) break;

                        const item = mappableLayers[i];
                        const layerTitle = item.title || "Unknown Layer";
                        const layerId = `layer-${i}`;

                        report.updateProgress("Building report...", `Generating map ${i + 1}/${mappableLayers.length}: ${layerTitle}`);
                        onProgress(`Map ${i + 1}/${mappableLayers.length}`, 20 + (70 * i / mappableLayers.length));

                        // Add placeholder for this layer
                        report.addMapPlaceholder(layerTitle, layerId);

                        // ImageServer layers — capture screenshot + elevation narrative
                        if (item.__isImageService) {
                            let dataUrl = null;
                            const imgLayerOpts = { url: item.url, title: layerTitle, visible: true };
                            if (item.__renderingRule) {
                                imgLayerOpts.rasterFunction = { functionName: item.__renderingRule };
                            }
                            const tempImg = new ImageryLayer(imgLayerOpts);
                            try {
                                view.map.add(tempImg);
                                setVisibilityForScreenshot(tempImg);
                                await waitForLayerReadyToCapture(tempImg, view, { timeoutMs: 12000 });
                                if (fixedExtent) {
                                    await view.goTo(fixedExtent, { animate: false });
                                } else {
                                    await view.goTo(selectionGeom.extent.expand(1.15), { animate: false });
                                }
                                await waitForLayerReadyToCapture(tempImg, view, { timeoutMs: 10000 });
                                await waitForViewStationary(800);
                                // Refresh mask so outer ring matches the current view extent
                                updateAoiMask(true);
                                await waitForTabVisible(5000);
                                dataUrl = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });
                            } catch (e) {
                                console.warn(`Imagery screenshot failed for ${layerTitle}:`, e);
                            } finally {
                                try { view.map.remove(tempImg); } catch (_) {}
                                restoreVisibility();
                            }

                            let elevStats = null;
                            try {
                                elevStats = await computeElevationStats(item.url, selectionGeom);
                            } catch (e) { /* skip */ }

                            const narrativeHtml = buildLayerNarrative({
                                aoiAcres, featureCount: 0, layerTitle,
                                acresCovered: 0, pctCovered: 0,
                                isPolygon: false, isImagery: true, elevStats
                            });

                            const sectionHtml = `
                                ${dataUrl ? `<div class="map"><img src="${dataUrl}" alt="${escapeHtml(layerTitle)} map" /></div>` : ''}
                                ${narrativeHtml}
                            `;
                            report.replaceMapPlaceholder(layerId, sectionHtml);
                            continue;
                        }

                        try {
                            // === Deferred feature query (screening only checked coverage) ===
                            let featureCount = item.count || 0;
                            if (featureCount === 0 && item.url) {
                                report.updateProgress("Building report...", `Querying features: ${layerTitle}`);
                                try {
                                    const qr = await querySingleLayer(item.url, item.title, selectionGeom, "intersects");
                                    featureCount = qr.count || 0;
                                    item.count = featureCount;
                                    item.rows = qr.features ? qr.features.map(f => f.attributes) : [];
                                    item._layer = qr.layer;
                                    item._exportQuery = qr.exportQuery;
                                } catch (qe) {
                                    console.warn(`Feature query failed for ${layerTitle}:`, qe);
                                }
                            }

                            // Skip layers with no actual features
                            if (featureCount === 0) {
                                report.replaceMapPlaceholder(layerId, '');
                                continue;
                            }

                            // Create temp layer for screenshot - use service's native renderer
                            const tempGeomType = await getLayerGeometryType(item.url);

                            const tempLayer = new FeatureLayer({
                                url: item.url,
                                outFields: ["*"],
                                visible: false
                            });
                            view.map.add(tempLayer);
                            tempLayer.definitionExpression = item._exportQuery?.where || "1=1";
                            
                            // Wait for layer to load so service renderer is available
                            try { await tempLayer.when(); } catch (e) { /* continue */ }
                            
                            // Thicken borders while preserving service's original symbology
                            thickenLayerSymbology(tempLayer, tempGeomType);

                            setVisibilityForScreenshot(tempLayer);

                            await waitForLayerReadyToCapture(tempLayer, view, { timeoutMs: 10000 });
                            await waitForTabVisible(5000);

                            // Fire coverage stats in parallel with screenshot capture
                            const isPolygonLayer = tempGeomType && String(tempGeomType).toLowerCase().includes('polygon');
                            const coveragePromise = isPolygonLayer
                                ? computeLayerCoverageStats(item, selectionGeom).catch(function () { return null; })
                                : Promise.resolve(null);

                            const dataUrl = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });

                            // Clean up temp layer
                            view.map.remove(tempLayer);

                            // Collect deferred coverage result
                            let acresCovered = 0;
                            let pctCovered = 0;
                            const covStats = await coveragePromise;
                            if (covStats) {
                                acresCovered = covStats.acresCovered || 0;
                                pctCovered = covStats.pctAoiCovered || 0;
                            }

                            const narrativeHtml = buildLayerNarrative({
                                aoiAcres, featureCount, layerTitle,
                                acresCovered, pctCovered, isPolygon: isPolygonLayer
                            });

                            // Build per-feature attribute table
                            const perFeatureTableHtml = (featureCount > 0)
                                ? await buildPerFeatureTable(item, selectionGeom, i)
                                : "";

                            // Build section HTML
                            const sectionHtml = `
                                ${dataUrl ? `<div class="map"><img src="${dataUrl}" alt="${escapeHtml(layerTitle)} map" /></div>` : '<div class="sub">Map generation failed</div>'}
                                ${narrativeHtml}
                                ${generateLayerAttributeSummary(item) ? `<table class="metaTbl">${generateLayerAttributeSummary(item)}</table>` : ''}
                                ${perFeatureTableHtml}
                            `;

                            report.replaceMapPlaceholder(layerId, sectionHtml);

                        } catch (e) {
                            console.warn(`Failed to generate map for ${layerTitle}:`, e);
                            report.replaceMapPlaceholder(layerId, `<div class="sub" style="color: #c62828;">Map generation failed: ${escapeHtml(e.message)}</div>`);
                        }
                    }

                } finally {
                    // Restore original state
                    try {
                        view.map.basemap = originalBasemap;
                    } catch (e) {}
                    restoreVisibility();
                }
            }

            // === STEP 4: Update summary stats now that feature counts are known ===
            const totalLayers = targetLayers.length;
            const layersWithHits = targetLayers.filter(x => (x.count || 0) > 0).length;
            const totalFeatures = targetLayers.reduce((sum, x) => sum + (x.count || 0), 0);

            report.updateSection("summary-stats-pills", `
                <div class="row">
                    <div class="pill">${totalLayers} Layers Queried</div>
                    <div class="pill">${layersWithHits} Layers in AOI</div>
                    <div class="pill">${totalFeatures} Features in AOI</div>
                </div>
            `);

            // === STEP 5: Data Sources table and footer ===
            const dataSourcesHtml = await buildLayerSourcesTable(targetLayers);
            report.appendContent(dataSourcesHtml);
            report.hideProgress();
            report.addFooter();

            // Auto-hide all-NULL columns now that content is fully loaded
            try {
                if (report.win && !report.win.closed && report.win.autoHideNullColumns) {
                    report.win.autoHideNullColumns();
                }
            } catch (e) { /* window may be closed */ }

            onProgress("Complete", 100);

        } catch (e) {
            console.error("Progressive report error:", e);
            report.appendContent(`<div style="color: #c62828; padding: 24px;">Error generating report: ${escapeHtml(e.message)}</div>`);
            report.hideProgress();
        } finally {
            await releaseWakeLock();
        }

        return report;
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
            getLayerGeometryType, makeRendererOpaque, getPresetRenderer,
            thickenLayerSymbology
        } = mapUtils;

        const {
            queryAllFeaturesPaged, querySingleLayer, computeElevationStats,
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
                .filter(x => x?.hasCoverage || (x?.count || 0) > 0)
                .filter(x => x?.url || x?.__isImageService)
                .filter(x => !(x.title && x.title.toLowerCase().includes("state boundaries")));

            // Categorize targets into buckets for grouped output
            const bucketedTargets = categorizeLayersIntoBuckets(targets);

            let sectionsHtml = "";
            let globalLayerIndex = 0;

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
                    await new Promise(r => setTimeout(r, 500));
                    await waitForViewStationary(800);
                } catch (e) {
                    console.warn("Failed to switch to imagery basemap:", e);
                }

                // Iterate through buckets in order
                for (const bucketKey of BUCKET_ORDER) {
                    const bucketItems = bucketedTargets[bucketKey] || [];
                    if (!bucketItems.length) continue;

                    // Get bucket metadata
                    const bucketMeta = REPORT_BUCKETS[bucketKey] || { label: "Other Layers", icon: "📋", description: "" };
                    
                    // Add bucket section header
                    sectionsHtml += `
                    <div class="bucket-header" id="bucket-${bucketKey}">
                        <h2>${bucketMeta.icon} ${escapeHtml(bucketMeta.label)}</h2>
                        <p class="bucket-description">${escapeHtml(bucketMeta.description)}</p>
                    </div>
                    `;

                    // Process each layer in this bucket
                    for (let bi = 0; bi < bucketItems.length; bi++) {
                        const item = bucketItems[bi];
                        globalLayerIndex++;
                        _setStatus(`building final report\u2026 (${globalLayerIndex}/${targets.length})`);

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
                            await waitForViewStationary(800);
                            // Refresh mask so outer ring matches the current view extent
                            updateAoiMask(true);

                            const dataUrl = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });
                            if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

                            const elevStats = await computeElevationStats(item.url, selectionGeom);

                            const narrativeHtml = buildLayerNarrative({
                                aoiAcres, featureCount: 0, layerTitle: item.title,
                                acresCovered: 0, pctCovered: 0,
                                isPolygon: false, isImagery: true, elevStats
                            });

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
                                ${narrativeHtml}
                                </div></div>
                              </div>
                            `;
                        } finally {
                            try { view.map.remove(temp); } catch (e) { }
                            restoreVisibility();
                        }
                        continue;
                    }

                    // === Deferred feature query (screening only checked coverage) ===
                    let featureCount = item.count || 0;
                    if (featureCount === 0 && item.url) {
                        _setStatus(`building final report\u2026 (querying ${item.title})`);
                        try {
                            const qr = await querySingleLayer(item.url, item.title, selectionGeom, "intersects");
                            featureCount = qr.count || 0;
                            item.count = featureCount;
                            item.rows = qr.features ? qr.features.map(f => f.attributes) : [];
                            item._layer = qr.layer;
                            item._exportQuery = qr.exportQuery;
                        } catch (qe) {
                            console.warn(`Feature query failed for ${item.title}:`, qe);
                        }
                    }

                    // Skip layers with no actual features
                    if ((item.count || 0) === 0) continue;

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

                    // Thicken borders while preserving the service's original symbology
                    thickenLayerSymbology(temp, tempGeomType);

                    try {
                        setVisibilityForScreenshot(temp);
                        await waitForLayerReadyToCapture(temp, view, { timeoutMs: 15000 });
                        if (fixedExtent) await view.goTo(fixedExtent, { animate: false });
                        else await view.goTo(selectionGeom.extent.expand(1.15), { animate: false });
                        // Re-check that layer has finished rendering at the new extent
                        await waitForLayerReadyToCapture(temp, view, { timeoutMs: 15000 });
                        await waitForViewStationary(800);
                        // Refresh mask so outer ring matches the current view extent
                        updateAoiMask(true);

                        // Fire coverage stats in parallel with screenshot capture
                        const isPolygonLayer = tempGeomType && String(tempGeomType).toLowerCase().includes('polygon');
                        const coveragePromise = isPolygonLayer
                            ? computeLayerCoverageStats(item, selectionGeom).catch(function () { return null; })
                            : Promise.resolve(null);

                        const dataUrl = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });
                        if (!dataUrl) throw new Error("Screenshot failed (no dataUrl).");

                        // Collect deferred coverage result
                        let acresCovered = 0;
                        let pctCovered   = 0;
                        const covStats = await coveragePromise;
                        if (covStats) {
                            acresCovered = covStats.acresCovered || 0;
                            pctCovered   = covStats.pctAoiCovered || 0;
                        }

                        const layerAttrSummary = generateLayerAttributeSummary(item);
                        const perFeatureTableHtml = (item.count > 0)
                            ? await buildPerFeatureTable(item, selectionGeom, globalLayerIndex)
                            : "";

                        const narrativeHtml = buildLayerNarrative({
                            aoiAcres, featureCount: item.count || 0, layerTitle: item.title,
                            acresCovered, pctCovered, isPolygon: isPolygonLayer
                        });

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
                            ${narrativeHtml}
                            ${layerAttrSummary ? `<table class="metaTbl">${layerAttrSummary}</table>` : ''}
                            ${perFeatureTableHtml}
                            </div></div>
                        </div>
                        <div class="pagebreak"></div>
                        `;
                    } finally {
                        try { view.map.remove(temp); } catch (e) { }
                        restoreVisibility();
                    }
                    } // end layer loop
                } // end bucket loop

                // Restore original basemap
                try {
                    view.map.basemap = originalBasemap;
                    await new Promise(r => setTimeout(r, 1000));
                } catch (e) {
                    console.warn("Failed to restore original basemap:", e);
                }
            }

            // STEP 4: Data Sources Appendix
            const dataSourcesHtml = await buildLayerSourcesTable(lastReportRowsByLayer);

            // STEP 4b: Generate findings summary paragraph
            const findingsSummaryHtml = generateFindingsSummary(lastReportRowsByLayer, aoiAcres);

            // STEP 5: Build Final HTML Document
            const totalLayers    = lastReportRowsByLayer.length;
            const layersWithHits = lastReportRowsByLayer.filter(x => (x.count || 0) > 0).length;
            const totalHits      = lastReportRowsByLayer.reduce((sum, x) => sum + (x.count || 0), 0);

            const totalsHtml = `
              <div class="row">
                <div class="pill">Layers Queried: <b>${escapeHtml(String(totalLayers))}</b></div>
                <div class="pill">Layers in AOI: <b>${escapeHtml(String(layersWithHits))}</b></div>
                <div class="pill">Features in AOI: <b>${escapeHtml(String(totalHits))}</b></div>
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
                <h2 id="section-aoi"><span class="section-num">3.</span> Area of Interest</h2>
                <p class="section-intro">The geographic boundary used for this analysis, shown in overview and detail views with state and county reference boundaries.</p>
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

            // Persist report to IndexedDB for bookmarkable URL (device-local only)
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
    // buildReportInBackground – builds report without opening window
    // Returns complete HTML document ready to open
    // ────────────────────────────────────────────
    async function buildReportInBackground(options = {}) {
        const bucketKey = options.bucketKey || null;
        const onProgress = options.onProgress || (() => {});
        const onStep = options.onStep || (() => {});
        const isCanceled = options.isCanceled || (() => false);

        const view = S.view;
        const selectionGeom = S.selectionGeom;
        const config = S.config;
        const lastReportRowsByLayer = S.lastReportRowsByLayer;
        const aoiLayer = S.aoiLayer;
        const aoiMaskLayer = S.aoiMaskLayer;
        const alwaysVisibleLayers = S.alwaysVisibleLayers;

        if (!view || !selectionGeom) {
            throw new Error("No AOI selected");
        }

        if (!lastReportRowsByLayer || !lastReportRowsByLayer.length) {
            throw new Error("No analysis data available");
        }

        // Get bucket info
        const bucketInfo = bucketKey ? REPORT_BUCKETS[bucketKey] : null;
        const reportTitle = bucketInfo 
            ? `${bucketInfo.icon} ${bucketInfo.label} Report`
            : "Land & Resource Intersection Analysis Report";

        // Filter layers to this bucket (or all for full report)
        let targetLayers;
        if (bucketKey) {
            const buckets = categorizeLayersIntoBuckets(lastReportRowsByLayer);
            targetLayers = buckets[bucketKey] || [];
        } else {
            targetLayers = lastReportRowsByLayer;
        }

        // Filter to layers with coverage that can generate maps
        // Feature counts are queried during report generation
        const mappableLayers = targetLayers
            .filter(x => x?.hasCoverage || (x?.count || 0) > 0) // Coverage from screening
            .filter(x => x?.url || x?.__isImageService)
            .filter(x => !(x.title && x.title.toLowerCase().includes("state boundaries")));

        const { querySingleLayer } = queryEngine;

        const {
            updateAoiMask, hideAoiMask, ensureAoiOnTop,
            waitForViewStationary, waitForLayerReadyToCapture,
            captureScreenshotWithWait, waitForTabVisible,
            acquireWakeLock, releaseWakeLock,
            getLayerGeometryType, makeRendererOpaque, getPresetRenderer,
            thickenLayerSymbology
        } = mapUtils;

        const { computeLayerCoverageStats, buildPerFeatureTable, computeElevationStats, SQM_PER_ACRE } = queryEngine;

        // Accumulate content sections
        const contentParts = [];
        let mapsGenerated = 0;
        let sectionsComplete = 0;

        try {
            await acquireWakeLock();

            // === STEP 1: AOI Overview ===
            onStep("Generating AOI overview maps");
            onProgress(5, mapsGenerated, sectionsComplete);

            if (isCanceled()) throw new Error("Canceled");

            // Calculate AOI info
            let aoiAcres = 0;
            try {
                const aoiSqm = Math.max(0, geometryEngine.geodesicArea(selectionGeom, "square-meters"));
                aoiAcres = aoiSqm / SQM_PER_ACRE;
            } catch (e) {
                aoiAcres = 0;
            }

            let aoiMethod = "Manually Drawn";
            if (S.aoiSource === "select") {
                const tool = plssToolLabel(S.aoiSourcePlssTool);
                aoiMethod = `Selected ${tool}`;
            } else if (S.aoiSource === "upload") {
                aoiMethod = `Uploaded File: ${S.aoiSourceLayerTitle || "unknown"}`;
            }

            // Generate AOI maps
            const aoiMapsHtml = await generateAoiMapsWithCircles();
            mapsGenerated += 2; // Usually generates 2 AOI maps
            sectionsComplete++;
            onProgress(15, mapsGenerated, sectionsComplete);

            if (isCanceled()) throw new Error("Canceled");

            // Build AOI section
            contentParts.push(`
                <h2>Area of Interest</h2>
                <p style="color: var(--muted); font-style: italic;">The geographic boundary used for this analysis.</p>
                ${aoiMapsHtml}
                <div class="aoi-details">
                    <div class="aoi-field"><span class="aoi-label">Area:</span> ${formatNumber(aoiAcres, 2)} acres</div>
                    <div class="aoi-field"><span class="aoi-label">Method:</span> ${escapeHtml(aoiMethod)}</div>
                </div>
            `);

            // === STEP 2: Summary stats placeholder (updated after layer queries) ===
            // Actual feature counts are computed during step 3; insert placeholder index
            const summaryPlaceholderIndex = contentParts.length;
            const layersInAoiEstimate = targetLayers.filter(x => x.hasCoverage).length;
            contentParts.push(`
                <h2>Summary</h2>
                <div class="totals">
                    <div class="row">
                        <div class="pill">${targetLayers.length} Layers Queried</div>
                        <div class="pill">${layersInAoiEstimate} Layers in AOI</div>
                        <div class="pill">\u2026 Features in AOI</div>
                    </div>
                </div>
            `);
            sectionsComplete++;
            onProgress(20, mapsGenerated, sectionsComplete);

            // === STEP 3: Generate maps for each layer ===
            if (mappableLayers.length === 0) {
                contentParts.push(`
                    <h2>Layer Details</h2>
                    <p style="color: var(--muted); font-style: italic;">No intersecting features found in the ${bucketInfo?.label || 'screened'} datasets.</p>
                `);
            } else {
                contentParts.push(`<h2>Layer Details</h2>`);

                const paddingFactor = config?.visualReport?.paddingFactor ?? 1.12;
                const width = config?.visualReport?.screenshotWidth ?? 1400;

                let fixedExtent = null;
                const ext = selectionGeom?.extent;
                if (ext && ext.expand) fixedExtent = ext.expand(paddingFactor);

                // Snapshot layer visibility
                const allLayers = view.map.layers.toArray();
                const visSnapshot = allLayers.map(l => ({ layer: l, visible: l.visible }));
                const originalBasemap = view.map.basemap;
                const imageryBasemapId = config?.map?.imageryBasemap || "satellite";

                function setVisibilityForScreenshot(tempLayer) {
                    for (const l of allLayers) {
                        if (aoiLayer && l === aoiLayer) { l.visible = true; continue; }
                        if (aoiMaskLayer && l === aoiMaskLayer) { l.visible = true; continue; }
                        if (l?.type === "tile") { l.visible = true; continue; }
                        if (alwaysVisibleLayers.includes(l)) { l.visible = true; continue; }
                        l.visible = false;
                    }
                    if (tempLayer) tempLayer.visible = true;
                    updateAoiMask(true);
                    ensureAoiOnTop();
                }

                function restoreVisibility() {
                    visSnapshot.forEach(s => { try { s.layer.visible = s.visible; } catch (e) {} });
                    hideAoiMask();
                    ensureAoiOnTop();
                }

                try {
                    // Switch to imagery basemap
                    view.map.basemap = imageryBasemapId;
                    await new Promise(r => setTimeout(r, 500));
                    await waitForViewStationary(800);

                    if (fixedExtent) {
                        await view.goTo(fixedExtent, { animate: false });
                        await waitForViewStationary(800);
                    }

                    // Process each layer
                    for (let i = 0; i < mappableLayers.length; i++) {
                        if (isCanceled()) throw new Error("Canceled");

                        const item = mappableLayers[i];
                        const layerTitle = item.title || "Unknown Layer";

                        onStep(`Generating map ${i + 1}/${mappableLayers.length}: ${layerTitle}`);
                        onProgress(20 + (70 * (i + 1) / mappableLayers.length), mapsGenerated, sectionsComplete);

                        // ImageServer layers — capture screenshot + elevation narrative
                        if (item.__isImageService) {
                            let dataUrl = null;
                            const imgLayerOpts = { url: item.url, title: layerTitle, visible: true };
                            if (item.__renderingRule) {
                                imgLayerOpts.rasterFunction = { functionName: item.__renderingRule };
                            }
                            const tempImg = new ImageryLayer(imgLayerOpts);
                            try {
                                view.map.add(tempImg);
                                setVisibilityForScreenshot(tempImg);
                                await waitForLayerReadyToCapture(tempImg, view, { timeoutMs: 12000 });
                                if (fixedExtent) {
                                    await view.goTo(fixedExtent, { animate: false });
                                } else {
                                    await view.goTo(selectionGeom.extent.expand(1.15), { animate: false });
                                }
                                await waitForLayerReadyToCapture(tempImg, view, { timeoutMs: 10000 });
                                await waitForViewStationary(800);
                                // Refresh mask so outer ring matches the current view extent
                                updateAoiMask(true);
                                await waitForTabVisible(5000);
                                dataUrl = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });
                            } catch (e) {
                                console.warn(`Imagery screenshot failed for ${layerTitle}:`, e);
                            } finally {
                                try { view.map.remove(tempImg); } catch (_) {}
                                restoreVisibility();
                            }

                            let elevStats = null;
                            try {
                                elevStats = await computeElevationStats(item.url, selectionGeom);
                            } catch (e) { /* skip */ }

                            const narrativeHtml = buildLayerNarrative({
                                aoiAcres, featureCount: 0, layerTitle,
                                acresCovered: 0, pctCovered: 0,
                                isPolygon: false, isImagery: true, elevStats
                            });

                            contentParts.push(`
                                <div class="section">
                                    <h3><button class="section-hide-btn" onclick="toggleSection(this)">✕ Hide</button>${escapeHtml(layerTitle)}</h3>
                                    <div class="section-collapse-wrap"><div class="section-collapse-inner">
                                    ${dataUrl ? `<div class="map"><div class="map-zoom-controls"><button class="zoom-in" title="Zoom in">+</button><button class="zoom-out" title="Zoom out">&minus;</button><button class="zoom-reset" title="Reset zoom" style="font-size:13px;">&#8634;</button></div><img src="${dataUrl}" alt="${escapeHtml(layerTitle)}" style="width:100%; border-radius:8px;" /></div>` : ''}
                                    ${narrativeHtml}
                                    </div></div>
                                </div>
                            `);
                            sectionsComplete++;
                            continue;
                        }

                        try {
                            // === Query for actual feature count and data ===
                            // This is where feature intersection is computed (deferred from screening)
                            let queryResult = null;
                            let featureCount = item.count || 0;
                            let featureRows = item.rows || [];
                            
                            if (featureCount === 0 && item.url) {
                                // Need to query for features (deferred from screening)
                                onStep(`Querying features: ${layerTitle}`);
                                try {
                                    queryResult = await querySingleLayer(item.url, item.title, selectionGeom, "intersects");
                                    featureCount = queryResult.count || 0;
                                    featureRows = queryResult.features ? queryResult.features.map(f => f.attributes) : [];
                                    // Update item with query results for potential later use
                                    item.count = featureCount;
                                    item.rows = featureRows;
                                    item._layer = queryResult.layer;
                                    item._exportQuery = queryResult.exportQuery;
                                } catch (qe) {
                                    console.warn(`Feature query failed for ${layerTitle}:`, qe);
                                }
                            }
                            
                            // Skip layers with no actual features
                            if (featureCount === 0) {
                                sectionsComplete++;
                                continue;
                            }
                            
                            onStep(`Generating map ${i + 1}/${mappableLayers.length}: ${layerTitle}`);
                            
                            // Create temp layer for screenshot - use service's native renderer
                            const tempGeomType = await getLayerGeometryType(item.url);

                            const tempLayer = new FeatureLayer({
                                url: item.url,
                                outFields: ["*"],
                                visible: false
                            });
                            view.map.add(tempLayer);
                            tempLayer.definitionExpression = item._exportQuery?.where || "1=1";
                            
                            // Wait for layer to load so service renderer is available
                            try { await tempLayer.when(); } catch (e) { /* continue */ }
                            
                            // Thicken borders while preserving service's original symbology
                            thickenLayerSymbology(tempLayer, tempGeomType);

                            setVisibilityForScreenshot(tempLayer);

                            await waitForLayerReadyToCapture(tempLayer, view, { timeoutMs: 10000 });

                            // Fire coverage stats in parallel with screenshot capture
                            const isPolygonLayer = tempGeomType && String(tempGeomType).toLowerCase().includes('polygon');
                            const coveragePromise = isPolygonLayer
                                ? computeLayerCoverageStats(item, selectionGeom).catch(function () { return null; })
                                : Promise.resolve(null);

                            const ss = await captureScreenshotWithWait({ width, tabWaitTimeout: 5000 });
                            const dataUrl = ss || null;

                            // Clean up temp layer
                            view.map.remove(tempLayer);
                            mapsGenerated++;

                            // Collect deferred coverage result
                            let acresCovered = 0;
                            let pctCovered = 0;
                            const covStats = await coveragePromise;
                            if (covStats) {
                                acresCovered = covStats.acresCovered || 0;
                                pctCovered = covStats.pctAoiCovered || 0;
                            }

                            // Build per-feature table
                            const perFeatureTableHtml = (featureCount > 0)
                                ? await buildPerFeatureTable(item, selectionGeom, i)
                                : "";

                            const narrativeHtml = buildLayerNarrative({
                                aoiAcres, featureCount, layerTitle,
                                acresCovered, pctCovered, isPolygon: isPolygonLayer
                            });

                            // Build section HTML with hide button and zoom controls (matching old report)
                            const attrSummary = generateLayerAttributeSummary(item);
                            contentParts.push(`
                                <div class="section">
                                    <h3><button class="section-hide-btn" onclick="toggleSection(this)">✕ Hide</button>${escapeHtml(layerTitle)}</h3>
                                    <div class="section-collapse-wrap"><div class="section-collapse-inner">
                                    <div class="map">
                                        <div class="map-zoom-controls">
                                            <button class="zoom-in" title="Zoom in">+</button>
                                            <button class="zoom-out" title="Zoom out">&minus;</button>
                                            <button class="zoom-reset" title="Reset zoom" style="font-size:13px;">&#8634;</button>
                                        </div>
                                        ${dataUrl ? `<img src="${dataUrl}" alt="AOI + ${escapeHtml(layerTitle)}"/>` : '<div class="sub">Map generation failed</div>'}
                                    </div>
                                    ${narrativeHtml}
                                    ${attrSummary ? `<table class="metaTbl">${attrSummary}</table>` : ''}
                                    ${perFeatureTableHtml}
                                    </div></div>
                                </div>
                                <div class="pagebreak"></div>
                            `);
                            sectionsComplete++;

                        } catch (e) {
                            console.warn(`Failed to generate map for ${layerTitle}:`, e);
                            contentParts.push(`
                                <div class="section">
                                    <h3>${escapeHtml(layerTitle)}</h3>
                                    <div class="sub" style="color: #c62828;">Map generation failed: ${escapeHtml(e.message)}</div>
                                </div>
                            `);
                            sectionsComplete++;
                        }

                        onProgress(20 + (70 * (i + 1) / mappableLayers.length), mapsGenerated, sectionsComplete);
                    }

                } finally {
                    // Restore original state
                    try {
                        view.map.basemap = originalBasemap;
                    } catch (e) {}
                    restoreVisibility();
                }
            }

            onStep("Finalizing report...");
            onProgress(95, mapsGenerated, sectionsComplete);

            // Update summary stats now that actual feature counts are known
            const totalLayers = targetLayers.length;
            const layersWithHits = targetLayers.filter(x => (x.count || 0) > 0).length;
            const totalFeatures = targetLayers.reduce((sum, x) => sum + (x.count || 0), 0);
            contentParts[summaryPlaceholderIndex] = `
                <h2>Summary</h2>
                <div class="totals">
                    <div class="row">
                        <div class="pill">${totalLayers} Layers Queried</div>
                        <div class="pill">${layersWithHits} Layers in AOI</div>
                        <div class="pill">${totalFeatures} Features in AOI</div>
                    </div>
                </div>
            `;

            onStep("Building data sources table...");
            const dataSourcesHtml = await buildLayerSourcesTable(targetLayers);

            // Generate regulatory/findings summary (uses updated feature counts from queries)
            const findingsSummaryHtml = generateFindingsSummary(targetLayers, aoiAcres);

            // Build complete HTML document
            const createdAt = formatDateTimeForReport(new Date());
            const bucketLabel = bucketInfo?.label || null;
            
            // Prepare export data (layer title, count, and sample attribute rows)
            const exportData = targetLayers.map(layer => ({
                title: layer.title || "Unknown",
                count: layer.count || 0,
                rows: (layer.rows || []).slice(0, 100) // Limit to 100 rows per layer for embedding
            }));
            const exportDataJson = JSON.stringify(exportData).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');

            const fullHtml = `<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width,initial-scale=1"/>
    <title>${escapeHtml(reportTitle)}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Merriweather:wght@400;700&family=Source+Sans+Pro:wght@400;600;700&display=swap" rel="stylesheet">
    <style>${getReportStyles()}</style>
</head>
<body>
    <div class="cv-filter-wrap">
    <header class="report-header">
        <div class="agency-name">U.S. Department of the Interior &bull; Bureau of Land Management</div>
        <h1>${escapeHtml(reportTitle)}</h1>
        <p class="meta">Generated: ${escapeHtml(createdAt)}${bucketLabel ? ` &bull; Category: ${escapeHtml(bucketLabel)}` : ''}</p>
        <div class="report-actions">
            <button class="export-btn" onclick="exportReportCsv()">📥 Export CSV</button>
            <button class="export-btn" onclick="window.print()">🖨️ Print Report</button>
        </div>
    </header>
    <script>
        var _exportData = ${exportDataJson};
        function exportReportCsv() {
            var blocks = [];
            _exportData.forEach(function(layer) {
                if (!layer.rows || !layer.rows.length) return;
                blocks.push('\\n"=== ' + layer.title.replace(/"/g, '""') + ' (' + layer.count + ' features) ==="');
                var keys = Object.keys(layer.rows[0]);
                blocks.push(keys.map(function(k){ return '"' + String(k).replace(/"/g, '""') + '"'; }).join(','));
                layer.rows.forEach(function(row) {
                    blocks.push(keys.map(function(k){ 
                        var v = row[k]; 
                        if (v == null) return '';
                        return '"' + String(v).replace(/"/g, '""') + '"';
                    }).join(','));
                });
            });
            var csv = blocks.join('\\n');
            var blob = new Blob([csv], {type: 'text/csv;charset=utf-8'});
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url;
            a.download = 'report_data.csv';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        // Toggle individual layer section visibility
        function toggleSection(btn) {
            var section = btn.closest('.section');
            if (!section) return;
            var isHidden = section.classList.toggle('section-hidden');
            btn.innerHTML = isHidden ? '&#x2713; Show' : '&#x2715; Hide';
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
            var allTh = table.querySelectorAll('th');
            for (var i = 0; i < allTh.length; i++) allTh[i].setAttribute('data-sort-dir', 'none');
            th.setAttribute('data-sort-dir', newSort);
            rows.sort(function(a, b) {
                var aVal = (a.children[colIdx] || {}).textContent || '';
                var bVal = (b.children[colIdx] || {}).textContent || '';
                var aNum = parseFloat(aVal.replace(/[^\\d.-]/g, ''));
                var bNum = parseFloat(bVal.replace(/[^\\d.-]/g, ''));
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    return newSort === 'asc' ? aNum - bNum : bNum - aNum;
                }
                return newSort === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
            });
            rows.forEach(function(r) { tbody.appendChild(r); });
        }

        // Interactive table: hide a column
        function hideColumn(wrapperId, colIdx) {
            var wrapper = document.getElementById(wrapperId);
            if (!wrapper) return;
            var elements = wrapper.querySelectorAll('[data-col="' + colIdx + '"]');
            for (var ei = 0; ei < elements.length; ei++) {
                elements[ei].style.display = 'none';
            }
            var th = wrapper.querySelector('thead th[data-col="' + colIdx + '"]');
            var label = th ? (th.getAttribute('data-label') || 'Column ' + colIdx) : 'Column ' + colIdx;
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
        }

        // Interactive table: show a hidden column
        function showColumn(wrapperId, colIdx) {
            var wrapper = document.getElementById(wrapperId);
            if (!wrapper) return;
            var elements = wrapper.querySelectorAll('[data-col="' + colIdx + '"]');
            for (var ei = 0; ei < elements.length; ei++) {
                elements[ei].style.display = '';
            }
            var bar = wrapper.querySelector('.hidden-cols-bar');
            if (bar) {
                var pills = bar.querySelectorAll('.hidden-col-pill[data-col="' + colIdx + '"]');
                for (var pi = 0; pi < pills.length; pi++) pills[pi].remove();
                if (!bar.querySelector('.hidden-col-pill')) bar.style.display = 'none';
            }
        }

        // Auto-hide columns that are entirely null/empty
        function autoHideNullColumns() {
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
        }

        // Initialize zoom/pan controls for map images
        document.addEventListener('DOMContentLoaded', function() {
            autoHideNullColumns();
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
                container.addEventListener('wheel', function(e) {
                    e.preventDefault();
                    var delta = e.deltaY < 0 ? 1.15 : 1/1.15;
                    scale = Math.max(1, Math.min(scale * delta, 8));
                    if (scale === 1) { panX = 0; panY = 0; }
                    applyTransform();
                }, { passive: false });
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
        });
    </script>
    ${findingsSummaryHtml ? `
    <section class="findings-section wrap">
        <h2 id="section-findings"><span class="section-num">1.</span> Regulatory Screening &amp; Findings Summary</h2>
        <p class="section-intro">A narrative summary of the regulatory framework, special designations, environmental factors, land use plans, and existing authorizations identified within the project area.</p>
        <div class="findings-summary">${findingsSummaryHtml}</div>
    </section>
    ` : ''}
    <main class="wrap">
        ${contentParts.join('\n')}
        ${dataSourcesHtml}
    </main>
    <footer class="report-footer">
        <p>This report is for informational purposes only and does not constitute a formal BLM determination.</p>
        <p>Generated by the BLM Permit Screening Tool &bull; ${formatDateTimeForReport(new Date())}</p>
    </footer>
    </div>
${getA11yWidgetBlock()}
</body>
</html>`;

            onProgress(100, mapsGenerated, sectionsComplete);
            return fullHtml;

        } finally {
            await releaseWakeLock();
        }
    }

    /**
     * Open a completed report HTML document in a new tab
     */
    function openCompletedReport(htmlContent) {
        if (!htmlContent) return false;
        const blob = new Blob([htmlContent], { type: "text/html;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const win = window.open(url, "_blank");
        window.setTimeout(() => URL.revokeObjectURL(url), 60000);
        return !!win;
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

        // Initialize summary engine (loaded as AMD dep)
        summaryEngine  = summaryEngineModule.init(state);

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
            // Progressive report builder (new)
            openProgressiveReport,
            buildProgressiveReport,
            categorizeLayersIntoBuckets,
            REPORT_BUCKETS,
            // Background report builder (builds complete HTML without opening window)
            buildReportInBackground,
            openCompletedReport,
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
