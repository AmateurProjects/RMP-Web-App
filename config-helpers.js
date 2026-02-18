/**
 * config-helpers.js — Pure utility functions, URL helpers, network wrappers,
 * and data-transform helpers shared across the Parcel Explorer application.
 *
 * AMD module — no external Esri dependencies; only standard browser APIs.
 */
define([], function () {
    "use strict";

    // ── String / HTML helpers ────────────────────────────────────────

    function escapeHtml(s) {
        return String(s).replace(/[&<>"']/g, (c) => ({
            "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;"
        }[c]));
    }

    function normalize(s) { return String(s || "").toLowerCase(); }

    function plssToolLabel(which) {
        return (which === "intersected") ? "Parcel" :
            (which === "township") ? "Township" :
                (which === "section") ? "Section" :
                    "PLSS";
    }

    function safeFilename(name) {
        return String(name).replace(/[^\w\-]+/g, "_").replace(/_+/g, "_").replace(/^_+|_+$/g, "").slice(0, 120) || "export";
    }

    function formatNumber(n, digits = 2) {
        const x = Number(n);
        if (!isFinite(x)) return "";
        return x.toLocaleString(undefined, { maximumFractionDigits: digits, minimumFractionDigits: digits });
    }

    // ── PLSS identification helpers ──────────────────────────────────

    function isPlssLayerTitleOrUrl(title, url) {
        const t = normalize(title);
        const u = normalize(url);
        return (
            t.includes("plss") ||
            t.includes("township") ||
            t.includes("section") ||
            t.includes("intersected") ||
            u.includes("/plss") ||
            u.includes("plss")
        );
    }

    function isPlssIntersectedLayerTitle(title) {
        const t = normalize(title);
        return t.includes("intersected");
    }

    // ── URL classification ───────────────────────────────────────────

    function isFeatureServerRoot(url) {
        return /\/FeatureServer\/?$/.test(url);
    }

    function isMapServerRoot(url) {
        return /\/MapServer\/?$/.test(url);
    }

    function normalizePjsonUrl(u) {
        return u.replace(/\/$/, "") + "?f=pjson";
    }

    function normalizeUrlKey(u) {
        return String(u || "").replace(/\/+$/, "");
    }

    // ── Network helpers ──────────────────────────────────────────────

    async function fetchJson(url) {
        const res = await fetch(url, { credentials: "omit" });
        if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText} for ${url}`);
        return res.json();
    }

    async function fetchJsonWithTimeout(url, timeoutMs = 8000, opts = {}) {
        const controller = new AbortController();
        const t = window.setTimeout(() => controller.abort(), timeoutMs);

        try {
            const fetchOpts = {
                credentials: "omit",
                signal: controller.signal
            };
            // Only bypass cache when explicitly requested (e.g. health checks).
            // Metadata / pjson lookups benefit from the browser cache.
            if (opts.noCache) fetchOpts.cache = "no-store";

            const res = await fetch(url, fetchOpts);

            if (!res.ok) throw new Error(`HTTP ${res.status}`);

            const txt = await res.text();

            let json = null;
            try {
                json = JSON.parse(txt);
            } catch (e) {
                throw new Error("Non-JSON response (possible HTML error page)");
            }

            if (json && json.error) {
                const code = json.error.code != null ? json.error.code : "";
                const msg = json.error.message || "ArcGIS error";
                throw new Error(`ArcGIS error ${code}: ${msg}`);
            }

            return json;
        } finally {
            window.clearTimeout(t);
        }
    }

    function pickServiceDescription(pjson) {
        const candidates = [
            pjson?.serviceDescription,
            pjson?.description,
            pjson?.documentInfo?.Title,
            pjson?.name
        ].filter(Boolean);

        return candidates.length ? String(candidates[0]) : "";
    }

    // ── Config index builders ────────────────────────────────────────

    function buildLayerCfgIndex(cfg) {
        const m = new Map();

        const addList = (kind, arr) => {
            (arr || []).forEach(l => {
                if (!l || !l.url) return;
                m.set(String(l.url), { kind, cfg: l });
            });
        };

        addList("selection", cfg?.selectionLayers);
        addList("report", cfg?.reportLayers);

        return m;
    }

    /**
     * Returns array of { kind, title, url } for all configured services.
     * @param {Object} config — the loaded config.json object
     */
    function getConfiguredServices(config) {
        const seen = new Set();
        const out = [];

        const add = (kind, title, url) => {
            const normUrl = String(url).replace(/\/+$/, "");
            if (seen.has(normUrl)) return;
            seen.add(normUrl);
            out.push({ kind, title, url });
        };

        (config.selectionLayers || []).forEach(l => add("Selection", l.title, l.url));
        (config.reportLayers || []).forEach(l => add("Report", l.title, l.url));

        return out;
    }

    // ── Basemap helpers ──────────────────────────────────────────────

    function setBasemapBaseLayerOpacity(basemap, opacity) {
        try {
            const baseLayers = basemap?.baseLayers?.toArray ? basemap.baseLayers.toArray() : [];
            baseLayers.forEach(l => { l.opacity = opacity; });
        } catch (e) {
            // ignore
        }
    }

    function isImageryBasemap(basemap) {
        const id = (basemap && (basemap.id || basemap.portalItem?.id || basemap.title)) ? String(basemap.id || basemap.title || "") : "";
        const title = basemap?.title ? String(basemap.title).toLowerCase() : "";
        return title.includes("satellite") || title.includes("imagery") || title.includes("hybrid") || id.toLowerCase().includes("satellite") || id.toLowerCase().includes("hybrid");
    }

    // ── Service expansion helpers ────────────────────────────────────

    async function expandMapServerToSublayers(serviceUrl, { polygonOnly = true } = {}) {
        const pjsonUrl = serviceUrl.replace(/\/$/, "") + "?f=pjson";
        const info = await fetchJson(pjsonUrl);
        const layers = Array.isArray(info?.layers) ? info.layers : [];

        if (!polygonOnly) {
            // Filter out Annotation sublayers — FeatureLayer cannot render them
            return layers
                .filter(l => {
                    const t = String(l.type || l.subLayerType || "").toLowerCase();
                    return !t.includes("annotation");
                })
                .map(l => {
                    let title = String(l.name || "");
                    title = title.replace(/intersected/ig, "Parcel");
                    return { title, url: serviceUrl.replace(/\/$/, "") + "/" + l.id };
                });
        }

        // Parallel geometry type checks for polygon filtering
        const checks = await Promise.allSettled(
            layers.map(async l => {
                // Skip Annotation sublayers early — FeatureLayer cannot render them
                const parentType = String(l.type || l.subLayerType || "").toLowerCase();
                if (parentType.includes("annotation")) return null;

                const layerUrl = serviceUrl.replace(/\/$/, "") + "/" + l.id;
                try {
                    const lpjson = await fetchJson(layerUrl + "?f=pjson");
                    // Also check sublayer type from its own metadata
                    const slType = String(lpjson?.type || "").toLowerCase();
                    if (slType.includes("annotation")) return null;
                    const g = String(lpjson?.geometryType || "");
                    if (!g.toLowerCase().includes("polygon")) return null;
                    let title = String(l.name || "");
                    title = title.replace(/intersected/ig, "Parcel");
                    return { title, url: layerUrl };
                } catch (e) {
                    return null;
                }
            })
        );

        return checks
            .filter(r => r.status === "fulfilled" && r.value != null)
            .map(r => r.value);
    }

    async function expandServiceToSublayers(serviceUrl) {
        const pjsonUrl = serviceUrl.replace(/\/$/, "") + "?f=pjson";
        const info = await fetchJson(pjsonUrl);
        const layers = (info && info.layers) ? info.layers : [];
        // Filter out Annotation sublayers — FeatureLayer cannot render them
        return layers
            .filter(l => {
                const t = String(l.type || l.subLayerType || "").toLowerCase();
                return !t.includes("annotation");
            })
            .map(l => ({
                title: (l && l.name) ? String(l.name) : `Layer ${l.id}`,
                url: serviceUrl.replace(/\/$/, "") + "/" + l.id
            }));
    }

    async function expandFeatureServerToPolygonSublayers(serviceUrl) {
        const pjsonUrl = serviceUrl.replace(/\/$/, "") + "?f=pjson";
        const info = await fetchJson(pjsonUrl);
        const layers = Array.isArray(info?.layers) ? info.layers : [];

        // Parallel geometry type checks
        const checks = await Promise.allSettled(
            layers.map(async l => {
                const layerUrl = serviceUrl.replace(/\/$/, "") + "/" + l.id;
                try {
                    const lpjson = await fetchJson(layerUrl + "?f=pjson");
                    // Skip annotation sublayers
                    const slType = String(lpjson?.type || "").toLowerCase();
                    if (slType.includes("annotation")) return null;
                    const g = String(lpjson?.geometryType || "").toLowerCase();
                    if (!g.includes("polygon")) return null;
                    // Skip layers without query capability
                    const caps = String(lpjson?.capabilities || "").toLowerCase();
                    if (caps && !caps.includes("query")) return null;
                    let title = l?.name ? String(l.name) : `Layer ${l.id}`;
                    title = title.replace(/intersected/ig, "Parcel");
                    return { title, url: layerUrl };
                } catch (e) {
                    return null;
                }
            })
        );

        return checks
            .filter(r => r.status === "fulfilled" && r.value != null)
            .map(r => r.value);
    }

    async function expandFeatureServerToAllSublayers(serviceUrl) {
        const pjsonUrl = serviceUrl.replace(/\/$/, "") + "?f=pjson";
        const info = await fetchJson(pjsonUrl);
        const layers = Array.isArray(info?.layers) ? info.layers : [];

        const out = [];
        for (const l of layers) {
            // Skip Annotation sublayers — FeatureLayer cannot render them
            const t = String(l.type || l.subLayerType || "").toLowerCase();
            if (t.includes("annotation")) continue;

            const layerUrl = serviceUrl.replace(/\/$/, "") + "/" + l.id;
            let title = l?.name ? String(l.name) : `Layer ${l.id}`;
            out.push({ title, url: layerUrl });
        }

        return out;
    }

    // ── Data-transform helpers ───────────────────────────────────────

    function flattenAttributes(features) {
        return (features || []).map(f => (f && f.attributes) ? f.attributes : {});
    }

    function toCsv(rows, preferredFirstCols = []) {
        if (!rows || !rows.length) return "";

        const colSet = new Set();
        for (const r of rows) {
            if (!r) continue;
            Object.keys(r).forEach(k => colSet.add(k));
        }

        const preferred = (preferredFirstCols || []).filter(c => colSet.has(c));
        preferred.forEach(c => colSet.delete(c));

        const rest = Array.from(colSet).sort((a, b) => a.localeCompare(b));
        const cols = [...preferred, ...rest];

        const escape = (v) => {
            const s = (v == null) ? "" : String(v);
            if (/[,"\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
            return s;
        };

        const header = cols.map(escape).join(",");
        const body = rows.map(r => cols.map(c => escape(r ? r[c] : "")).join(",")).join("\n");
        return header + "\n" + body;
    }

    function downloadText(filename, text) {
        const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    }

    // ── Public interface ─────────────────────────────────────────────

    return {
        // String / HTML
        escapeHtml,
        normalize,
        plssToolLabel,
        safeFilename,
        formatNumber,

        // PLSS identification
        isPlssLayerTitleOrUrl,
        isPlssIntersectedLayerTitle,

        // URL classification
        isFeatureServerRoot,
        isMapServerRoot,
        normalizePjsonUrl,
        normalizeUrlKey,

        // Network
        fetchJson,
        fetchJsonWithTimeout,
        pickServiceDescription,

        // Config index
        buildLayerCfgIndex,
        getConfiguredServices,

        // Basemap
        setBasemapBaseLayerOpacity,
        isImageryBasemap,

        // Service expansion
        expandMapServerToSublayers,
        expandServiceToSublayers,
        expandFeatureServerToPolygonSublayers,
        expandFeatureServerToAllSublayers,

        // Data transforms
        flattenAttributes,
        toCsv,
        downloadText
    };
});
