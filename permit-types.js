/**
 * permit-types.js  –  AMD module defining Permit Type constants
 *
 * Shared between app.js and final-report.js so both use the same
 * permit-type definitions, report groupings, and category mappings.
 *
 * Each permit type has:
 *   - label / icon / description  – UI metadata
 *   - groups[]  – ordered array of report sub-groupings for Step 3
 *     Each group has label, icon, description, and categoryKeys[]
 *     that map to each layer's `category` field in config.json.
 *
 * A layer's `permitTypes` array in config.json controls which
 * permit types include that layer.  The special value "core" means
 * the layer is included for ALL permit types.
 */
define([], function () {
    "use strict";

    /**
     * Master list of permit types.
     * Keys must match the values stored in config.json reportLayer.permitTypes[].
     */
    var PERMIT_TYPES = {
        "oil-gas": {
            label: "Oil & Gas",
            icon: "🛢️",
            description: "Screening for oil and gas lease applications, APDs, and related energy development on federal lands.",
            groups: [
                {
                    key: "lease-availability",
                    label: "Lease Availability & Land Status",
                    icon: "🏛️",
                    description: "Federal land ownership, administrative boundaries, and land status relevant to oil & gas leasing.",
                    categoryKeys: ["land-status"]
                },
                {
                    key: "surface-constraints",
                    label: "Surface Constraints & Designations",
                    icon: "⭐",
                    description: "Special designations, land use plans, and surface constraints that may affect oil & gas operations.",
                    categoryKeys: ["special", "land-use"]
                },
                {
                    key: "environmental",
                    label: "Environmental & ESA",
                    icon: "🌿",
                    description: "Threatened and endangered species habitat, wetlands, wildlife corridors, and environmental factors.",
                    categoryKeys: ["environmental"]
                },
                {
                    key: "existing-authorizations",
                    label: "Existing Authorizations",
                    icon: "📝",
                    description: "Active permits, leases, rights-of-way, and other authorizations that overlap the project area.",
                    categoryKeys: ["authorizations"]
                }
            ]
        },
        "grazing": {
            label: "Grazing",
            icon: "🐄",
            description: "Screening for grazing permits and allotment management on BLM-administered lands.",
            groups: [
                {
                    key: "land-status",
                    label: "Land Status & Authority",
                    icon: "🏛️",
                    description: "Federal land ownership, administrative boundaries, and jurisdictional authority.",
                    categoryKeys: ["land-status"]
                },
                {
                    key: "allotment-management",
                    label: "Allotment Management & Land Use",
                    icon: "📑",
                    description: "Land use plans, grazing allocations, and management prescriptions.",
                    categoryKeys: ["land-use", "authorizations"]
                },
                {
                    key: "environmental",
                    label: "Environmental & ESA",
                    icon: "🌿",
                    description: "Critical habitat, riparian areas, wildlife corridors, and environmental factors affecting rangeland.",
                    categoryKeys: ["environmental"]
                },
                {
                    key: "special-designations",
                    label: "Special Designations",
                    icon: "⭐",
                    description: "Wilderness, ACECs, conservation lands, and other special designations within the allotment area.",
                    categoryKeys: ["special"]
                }
            ]
        },
        "mining": {
            label: "Mining",
            icon: "⛏️",
            description: "Screening for mining claims, plans of operations, and mineral exploration on federal lands.",
            groups: [
                {
                    key: "mineral-status",
                    label: "Mineral Status & Land Status",
                    icon: "🏛️",
                    description: "Federal land ownership, mineral allocations, and land status relevant to mining operations.",
                    categoryKeys: ["land-status", "land-use"]
                },
                {
                    key: "surface-constraints",
                    label: "Surface Constraints & Designations",
                    icon: "⭐",
                    description: "Special designations and surface constraints that may restrict or condition mining activities.",
                    categoryKeys: ["special"]
                },
                {
                    key: "environmental",
                    label: "Environmental & ESA",
                    icon: "🌿",
                    description: "Critical habitat, wetlands, water resources, and environmental factors.",
                    categoryKeys: ["environmental"]
                },
                {
                    key: "existing-authorizations",
                    label: "Existing Authorizations",
                    icon: "📝",
                    description: "Active mining claims, plans of operations, and other authorizations in the area.",
                    categoryKeys: ["authorizations"]
                }
            ]
        },
        "row": {
            label: "Right-of-Way (ROW)",
            icon: "🛤️",
            description: "Screening for right-of-way grants, including pipelines, transmission lines, roads, and other linear facilities.",
            groups: [
                {
                    key: "land-status",
                    label: "Land Status & Authority",
                    icon: "🏛️",
                    description: "Federal land ownership, administrative boundaries, and jurisdictional authority along the proposed route.",
                    categoryKeys: ["land-status"]
                },
                {
                    key: "corridor-constraints",
                    label: "Corridor & Route Constraints",
                    icon: "📑",
                    description: "Land use plans, designated corridors, and special designations that may affect routing.",
                    categoryKeys: ["land-use", "special"]
                },
                {
                    key: "environmental",
                    label: "Environmental & ESA",
                    icon: "🌿",
                    description: "Critical habitat, wetlands, wildlife corridors, and environmental factors along the route.",
                    categoryKeys: ["environmental"]
                },
                {
                    key: "existing-authorizations",
                    label: "Existing Authorizations",
                    icon: "📝",
                    description: "Existing rights-of-way, permits, leases, and other authorizations along the proposed route.",
                    categoryKeys: ["authorizations"]
                }
            ]
        },
        "realty": {
            label: "Realty",
            icon: "🏠",
            description: "Screening for realty actions including land sales, exchanges, withdrawals, and other dispositions.",
            groups: [
                {
                    key: "land-status",
                    label: "Land Status & Authority",
                    icon: "🏛️",
                    description: "Federal land ownership, administrative boundaries, and jurisdictional authority.",
                    categoryKeys: ["land-status"]
                },
                {
                    key: "land-use-plans",
                    label: "Land Use Plans & Allocations",
                    icon: "📑",
                    description: "Resource Management Plans and land use allocations governing the area.",
                    categoryKeys: ["land-use"]
                },
                {
                    key: "environmental",
                    label: "Environmental & ESA",
                    icon: "🌿",
                    description: "Critical habitat, wetlands, and environmental factors.",
                    categoryKeys: ["environmental"]
                },
                {
                    key: "existing-authorizations",
                    label: "Existing Authorizations & Designations",
                    icon: "📝",
                    description: "Active permits, special designations, and other encumbrances on the land.",
                    categoryKeys: ["authorizations", "special"]
                }
            ]
        }
    };

    /** Ordered list of permit type keys for consistent UI rendering */
    var PERMIT_TYPE_ORDER = ["oil-gas", "grazing", "mining", "row", "realty"];

    /**
     * The old 5-category bucket definitions.
     * Still needed for fallback title-matching when a layer has no explicit `category` field.
     * Also used by admin stats and category badges.
     */
    var CATEGORY_DEFS = {
        "land-status": {
            label: "Land Status & Authority", icon: "🏛️",
            patterns: [/federal lands/i, /admin.*unit/i, /state boundar/i, /usfws.*region/i, /aoi source/i,
                        /bia.*aian/i, /indian/i, /alaska.*native/i, /tribal/i, /surface.*ownership/i,
                        /land use planning bound/i]
        },
        "land-use": {
            label: "Land Use Plans & Allocations", icon: "📑",
            patterns: [/land use plan/i, /revision.*development/i, /timber/i, /locatable.*mineral/i,
                        /taylor grazing/i, /tga/i]
        },
        "special": {
            label: "Special Designations", icon: "⭐",
            patterns: [/acec/i, /critical environmental/i, /nlcs/i, /conservation area/i, /national monument/i,
                        /wilderness/i, /wsa/i, /recreation site/i, /lwcf/i, /conservation fund/i, /visual resource/i,
                        /wild.*scenic.*river/i, /roadless/i, /national forest bound/i, /national wildlife refuge/i, /nwr/i]
        },
        "environmental": {
            label: "Environmental & ESA", icon: "🌿",
            patterns: [/critical habitat/i, /ungulate/i, /migration/i, /wild horse/i, /burro/i, /elevation/i, /fire perim/i,
                        /wetland/i, /nwi/i, /riparian/i, /nhd/i, /hydrography/i, /watershed/i, /wbd/i,
                        /flood/i, /nfhl/i, /fema/i, /sagebrush/i, /fiat/i, /danl/i, /disturbance/i,
                        /at.risk.*species/i, /t\&e/i, /threatened/i]
        },
        "authorizations": {
            label: "Existing Authorizations", icon: "📝",
            patterns: [/grazing allot/i, /grazing pasture/i, /oil.*gas/i, /mlrs.*row/i, /lua.*row/i, /eplanning/i, /plss.*parcel/i,
                        /mining claim/i, /lua.*lease/i, /lua.*permit/i, /lua.*easem/i, /geothermal/i, /coal case/i,
                        /oil shale/i, /non.energy/i, /mineral material/i, /locatable notice/i, /locatable plan/i,
                        /participating area/i, /agreement/i, /gtlf/i, /road.*trail/i]
        }
    };

    var CATEGORY_ORDER = ["land-status", "land-use", "special", "environmental", "authorizations"];

    /**
     * Determine the category of a layer from its config or via fallback title matching.
     * @param {Object} item - { title, url } 
     * @param {Map} [layerCfgByUrl] - optional config lookup map
     * @returns {string} category key, e.g. "environmental"
     */
    /**
     * Look up a layer's config entry from layerCfgByUrl.
     * Tries the exact URL first; if not found, strips the trailing
     * sublayer ID (e.g. "/0", "/1") and retries with the parent
     * service URL. This handles layers configured as root
     * FeatureServer/MapServer URLs that get expanded to sublayers
     * at runtime.
     */
    function lookupLayerCfg(layerCfgByUrl, url) {
        if (!layerCfgByUrl || !url) return null;
        var key = String(url);
        var entry = layerCfgByUrl.get(key);
        if (entry) return entry;
        // Fallback: strip trailing /<number> (sublayer ID) and try parent URL
        var parent = key.replace(/\/\d+$/, "");
        if (parent !== key) return layerCfgByUrl.get(parent) || null;
        return null;
    }

    function resolveCategory(item, layerCfgByUrl) {
        var title = (item && item.title) || "";
        // 1. Explicit category from config
        if (layerCfgByUrl) {
            var cfgEntry = lookupLayerCfg(layerCfgByUrl, item.url);
            var cfgCategory = cfgEntry && cfgEntry.cfg && cfgEntry.cfg.category;
            if (cfgCategory && CATEGORY_DEFS[cfgCategory]) return cfgCategory;
        }
        // 2. Fallback: regex title matching
        for (var bk in CATEGORY_DEFS) {
            if (CATEGORY_DEFS.hasOwnProperty(bk)) {
                var patterns = CATEGORY_DEFS[bk].patterns;
                for (var i = 0; i < patterns.length; i++) {
                    if (patterns[i].test(title)) return bk;
                }
            }
        }
        return "uncategorized";
    }

    /**
     * Filter layers to those matching the given permitType + "core".
     * @param {Array} reportItems - all screened layers
     * @param {string} permitTypeKey - e.g. "oil-gas"
     * @param {Map} layerCfgByUrl - config lookup map
     * @returns {Array} filtered layers
     */
    function filterLayersByPermitType(reportItems, permitTypeKey, layerCfgByUrl) {
        if (!permitTypeKey) return (reportItems || []).slice();
        return (reportItems || []).filter(function (item) {
            var cfgEntry = lookupLayerCfg(layerCfgByUrl, item.url);
            var permitTypes = (cfgEntry && cfgEntry.cfg && cfgEntry.cfg.permitTypes) || [];
            return permitTypes.indexOf(permitTypeKey) !== -1 || permitTypes.indexOf("core") !== -1;
        });
    }

    /**
     * Categorize layers into the permit-type-specific report groups.
     * Returns an object keyed by group.key with arrays of layer items.
     * Also includes "all-data" and "uncategorized" keys.
     *
     * @param {Array} reportItems - all screened layers (pre-filtered or not)
     * @param {string} permitTypeKey - e.g. "oil-gas"
     * @param {Map} layerCfgByUrl - config lookup map
     * @returns {{ groups: Object<string, Array>, allData: Array, permitType: Object }}
     */
    function categorizeByPermitType(reportItems, permitTypeKey, layerCfgByUrl) {
        var ptDef = PERMIT_TYPES[permitTypeKey];
        if (!ptDef) {
            // Fallback: return all items in a single "all-data" group
            return { groups: {}, allData: (reportItems || []).slice(), permitType: null };
        }

        // Filter to permit type + core
        var filtered = filterLayersByPermitType(reportItems, permitTypeKey, layerCfgByUrl);

        // Build group buckets  
        var groups = {};
        var placed = new Set();
        for (var g = 0; g < ptDef.groups.length; g++) {
            var grp = ptDef.groups[g];
            groups[grp.key] = [];
        }
        groups["uncategorized"] = [];

        for (var i = 0; i < filtered.length; i++) {
            var item = filtered[i];
            var cat = resolveCategory(item, layerCfgByUrl);
            var wasPlaced = false;
            for (var g2 = 0; g2 < ptDef.groups.length; g2++) {
                var grp2 = ptDef.groups[g2];
                if (grp2.categoryKeys.indexOf(cat) !== -1) {
                    groups[grp2.key].push(item);
                    wasPlaced = true;
                    placed.add(item);
                    break;
                }
            }
            if (!wasPlaced) {
                groups["uncategorized"].push(item);
            }
        }

        return {
            groups: groups,
            allData: filtered,
            permitType: ptDef
        };
    }

    /**
     * Legacy: categorize layers into the flat 5-category buckets
     * (for full report or when no permit type is selected).
     */
    function categorizeIntoBuckets(reportItems, layerCfgByUrl) {
        var buckets = {};
        for (var key in CATEGORY_DEFS) {
            if (CATEGORY_DEFS.hasOwnProperty(key)) buckets[key] = [];
        }
        buckets["uncategorized"] = [];
        for (var i = 0; i < (reportItems || []).length; i++) {
            var item = reportItems[i];
            var cat = resolveCategory(item, layerCfgByUrl);
            if (buckets[cat]) {
                buckets[cat].push(item);
            } else {
                buckets["uncategorized"].push(item);
            }
        }
        return buckets;
    }

    return {
        PERMIT_TYPES: PERMIT_TYPES,
        PERMIT_TYPE_ORDER: PERMIT_TYPE_ORDER,
        CATEGORY_DEFS: CATEGORY_DEFS,
        CATEGORY_ORDER: CATEGORY_ORDER,
        lookupLayerCfg: lookupLayerCfg,
        resolveCategory: resolveCategory,
        filterLayersByPermitType: filterLayersByPermitType,
        categorizeByPermitType: categorizeByPermitType,
        categorizeIntoBuckets: categorizeIntoBuckets
    };
});
