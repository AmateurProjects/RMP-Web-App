/**
 * check-services.js — Node.js script for GitHub Actions
 *
 * Reads config.json, extracts ALL unique service URLs from
 * selectionLayers + reportLayers, pings each one with ?f=pjson,
 * and reports UP / DOWN.
 *
 * Outputs:
 *   - Console summary table
 *   - .github/scripts/health-report.md (for issue creation)
 *   - Sets GitHub Actions output `has_failures` to 'true' or 'false'
 */

const fs   = require("fs");
const path = require("path");

const TIMEOUT_MS   = 15000;           // per-service timeout
const CONCURRENCY  = 8;               // parallel pings
const DATA_PROBE_LAYER_LIMIT = 5;
const DATA_PROBE_FEATURE_LIMIT = 1;

async function main() {
    const configPath = path.join(__dirname, "..", "..", "config.json");
    const config = JSON.parse(fs.readFileSync(configPath, "utf8"));

    // Collect all unique service URLs
    const seen = new Set();
    const services = [];

    function add(kind, title, url) {
        const norm = String(url || "").replace(/\/+$/, "");
        if (!norm || seen.has(norm)) return;
        seen.add(norm);
        services.push({ kind, title, url: norm });
    }

    (config.selectionLayers || []).forEach(l => add("Selection", l.title, l.url));
    (config.reportLayers    || []).forEach(l => add("Report",    l.title, l.url));

    console.log(`\nChecking ${services.length} services …\n`);

    // Ping in batches
    const results = [];
    for (let i = 0; i < services.length; i += CONCURRENCY) {
        const batch = services.slice(i, i + CONCURRENCY);
        const batchResults = await Promise.allSettled(batch.map(checkOne));
        for (let j = 0; j < batch.length; j++) {
            const svc = batch[j];
            const r   = batchResults[j];
            if (r.status === "fulfilled") {
                results.push({ ...svc, ...r.value });
            } else {
                results.push({ ...svc, status: "DOWN", error: String(r.reason), responseTimeMs: null });
            }
        }
    }

    // Sort: DOWN first, then by title
    results.sort((a, b) => {
        if (a.status !== b.status) return a.status === "DOWN" ? -1 : 1;
        return a.title.localeCompare(b.title);
    });

    // Console output
    const down = results.filter(r => r.status === "DOWN");
    const up   = results.filter(r => r.status === "UP");
    const slow = up.filter(r => r.responseTimeMs > 5000);

    console.log("─".repeat(80));
    const pad = (s, n) => String(s).padEnd(n);
    for (const r of results) {
        const icon   = r.status === "UP" ? "✅" : "❌";
        const timing = r.responseTimeMs != null ? `${r.responseTimeMs}ms` : "—";
        const err    = r.error ? `  (${r.error.slice(0, 60)})` : "";
        console.log(`${icon} ${pad(r.status, 5)} ${pad(timing, 8)} ${pad(r.kind, 10)} ${r.title}${err}`);
    }
    console.log("─".repeat(80));
    console.log(`\nTotal: ${results.length}   UP: ${up.length}   DOWN: ${down.length}   Slow (>5s): ${slow.length}\n`);

    // Build markdown report
    const lines = [];
    lines.push(`# Service Health Report`);
    lines.push(``);
    lines.push(`**Date:** ${new Date().toISOString()}`);
    lines.push(`**Total:** ${results.length} &nbsp;|&nbsp; **UP:** ${up.length} &nbsp;|&nbsp; **DOWN:** ${down.length} &nbsp;|&nbsp; **Slow (>5s):** ${slow.length}`);
    lines.push(``);

    if (down.length > 0) {
        lines.push(`## ❌ DOWN Services`);
        lines.push(``);
        lines.push(`| Service | Type | Error |`);
        lines.push(`|---------|------|-------|`);
        for (const r of down) {
            lines.push(`| ${r.title} | ${r.kind} | ${(r.error || "").replace(/\|/g, "\\|").slice(0, 80)} |`);
        }
        lines.push(``);
    }

    if (slow.length > 0) {
        lines.push(`## ⚠️ Slow Services (>5 s)`);
        lines.push(``);
        lines.push(`| Service | Type | Response Time |`);
        lines.push(`|---------|------|---------------|`);
        for (const r of slow) {
            lines.push(`| ${r.title} | ${r.kind} | ${r.responseTimeMs}ms |`);
        }
        lines.push(``);
    }

    if (up.length > 0 && down.length === 0 && slow.length === 0) {
        lines.push(`## ✅ All Services Healthy`);
        lines.push(``);
        lines.push(`All ${up.length} configured services responded successfully.`);
    }

    lines.push(``);
    lines.push(`<details><summary>Full Results</summary>`);
    lines.push(``);
    lines.push(`| Status | Service | Type | Time | URL |`);
    lines.push(`|--------|---------|------|------|-----|`);
    for (const r of results) {
        const icon = r.status === "UP" ? "✅" : "❌";
        const time = r.responseTimeMs != null ? `${r.responseTimeMs}ms` : "—";
        lines.push(`| ${icon} | ${r.title} | ${r.kind} | ${time} | [Link](${r.url}) |`);
    }
    lines.push(``);
    lines.push(`</details>`);

    const reportPath = path.join(__dirname, "health-report.md");
    fs.writeFileSync(reportPath, lines.join("\n"), "utf8");
    console.log(`Report written to ${reportPath}`);

    // Set GitHub Actions output
    const outputFile = process.env.GITHUB_OUTPUT;
    if (outputFile) {
        fs.appendFileSync(outputFile, `has_failures=${down.length > 0 ? "true" : "false"}\n`);
    }

    // Exit with code 1 if any service is down (marks the action as failed)
    if (down.length > 0) {
        console.error(`\n⚠️  ${down.length} service(s) are DOWN.\n`);
        process.exitCode = 1;
    }
}

async function checkOne(svc) {
    const baseUrl = stripQueryAndSlash(svc.url);
    const pjsonUrl = baseUrl + "?f=pjson";
    const t0 = Date.now();

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);

    try {
        const res = await fetch(pjsonUrl, {
            signal: controller.signal,
            headers: { "Accept": "application/json" }
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);

        const text = await res.text();
        let json;
        try { json = JSON.parse(text); } catch { throw new Error("Non-JSON response"); }

        if (json.error) {
            const code = json.error.code || "";
            throw new Error(`ArcGIS ${code}: ${json.error.message || "error"}`);
        }

        const probe = await runDataProbe(baseUrl, json, controller.signal);
        if (!probe.ok) {
            throw new Error(`Data probe failed: ${probe.detail}`);
        }

        return {
            status: "UP",
            responseTimeMs: Date.now() - t0,
            error: null,
            probeMode: probe.mode,
            probeDetail: probe.detail
        };
    } catch (e) {
        const msg = e.name === "AbortError" ? `Timeout (${TIMEOUT_MS}ms)` : e.message;
        return {
            status: "DOWN",
            responseTimeMs: Date.now() - t0,
            error: msg,
            probeMode: null,
            probeDetail: null
        };
    } finally {
        clearTimeout(timer);
    }
}

function stripQueryAndSlash(url) {
    return String(url || "").split("?")[0].replace(/\/+$/, "");
}

function detectServiceKind(url) {
    const base = stripQueryAndSlash(url);
    if (/\/ImageServer$/i.test(base)) return "image-root";
    if (/\/FeatureServer\/\d+$/i.test(base)) return "feature-layer";
    if (/\/FeatureServer$/i.test(base)) return "feature-root";
    if (/\/MapServer\/\d+$/i.test(base)) return "map-layer";
    if (/\/MapServer$/i.test(base)) return "map-root";
    if (/\/GeocodeServer$/i.test(base)) return "geocode";
    return "other";
}

function pickCandidateSublayerIds(meta) {
    const layers = Array.isArray(meta?.layers) ? meta.layers : [];
    return layers
        .filter(l => Number.isFinite(Number(l?.id)))
        .map(l => Number(l.id))
        .slice(0, DATA_PROBE_LAYER_LIMIT);
}

async function fetchJsonChecked(url, signal) {
    const res = await fetch(url, {
        signal,
        headers: { "Accept": "application/json" }
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data && data.error) {
        const code = data.error.code != null ? data.error.code : "";
        throw new Error(`ArcGIS ${code}: ${data.error.message || "error"}`);
    }
    return data;
}

async function probeFeatureOrMapLayer(layerUrl, signal) {
    const base = stripQueryAndSlash(layerUrl);
    const queryUrl =
        `${base}/query?where=1%3D1` +
        `&outFields=*` +
        `&returnGeometry=true` +
        `&resultRecordCount=${DATA_PROBE_FEATURE_LIMIT}` +
        `&f=json`;

    const data = await fetchJsonChecked(queryUrl, signal);
    const features = Array.isArray(data?.features) ? data.features : [];
    if (!features.length) {
        return { ok: false, mode: "query", detail: "Query returned 0 features" };
    }

    const withGeometry = features.find(f => !!f?.geometry);
    if (!withGeometry) {
        return { ok: false, mode: "query", detail: "Query returned features without geometry" };
    }

    return { ok: true, mode: "query", detail: "Feature query returned geometry" };
}

async function probeFeatureOrMapRoot(serviceUrl, serviceMeta, signal) {
    const base = stripQueryAndSlash(serviceUrl);
    const ids = pickCandidateSublayerIds(serviceMeta);
    if (!ids.length) {
        return { ok: false, mode: "query", detail: "Service has no sublayers to probe" };
    }

    const errors = [];
    for (const id of ids) {
        const probe = await probeFeatureOrMapLayer(`${base}/${id}`, signal);
        if (probe.ok) {
            return { ok: true, mode: "query", detail: `${probe.detail} (sublayer ${id})` };
        }
        errors.push(`/${id}: ${probe.detail}`);
    }

    return { ok: false, mode: "query", detail: `No sublayer returned geometry (${errors.join("; ")})` };
}

async function probeImageService(serviceUrl, serviceMeta, signal) {
    const base = stripQueryAndSlash(serviceUrl);
    const ext = serviceMeta?.extent || serviceMeta?.fullExtent;
    const sr = ext?.spatialReference || serviceMeta?.spatialReference || {};
    const wkid = sr.latestWkid || sr.wkid || 4326;

    if (!ext || [ext.xmin, ext.ymin, ext.xmax, ext.ymax].some(v => !Number.isFinite(Number(v)))) {
        return { ok: false, mode: "exportImage", detail: "Image service extent unavailable for export probe" };
    }

    const exportUrl =
        `${base}/exportImage` +
        `?bbox=${ext.xmin},${ext.ymin},${ext.xmax},${ext.ymax}` +
        `&bboxSR=${wkid}` +
        `&imageSR=${wkid}` +
        `&size=64,64` +
        `&format=png` +
        `&f=json`;

    const data = await fetchJsonChecked(exportUrl, signal);
    if (!data?.href && !data?.url) {
        return { ok: false, mode: "exportImage", detail: "exportImage did not return an image URL" };
    }

    return { ok: true, mode: "exportImage", detail: "Image export returned raster URL" };
}

async function probeGeocodeService(serviceUrl, signal) {
    const base = stripQueryAndSlash(serviceUrl);
    const suggestUrl = `${base}/suggest?text=denver&maxSuggestions=1&f=json`;
    const data = await fetchJsonChecked(suggestUrl, signal);
    const suggestions = Array.isArray(data?.suggestions) ? data.suggestions : [];
    if (!suggestions.length) {
        return { ok: false, mode: "suggest", detail: "Geocode suggest returned 0 suggestions" };
    }
    return { ok: true, mode: "suggest", detail: "Geocode suggest returned candidates" };
}

async function runDataProbe(baseUrl, serviceMeta, signal) {
    const kind = detectServiceKind(baseUrl);
    if (kind === "feature-layer" || kind === "map-layer") {
        return probeFeatureOrMapLayer(baseUrl, signal);
    }
    if (kind === "feature-root" || kind === "map-root") {
        return probeFeatureOrMapRoot(baseUrl, serviceMeta, signal);
    }
    if (kind === "image-root") {
        return probeImageService(baseUrl, serviceMeta, signal);
    }
    if (kind === "geocode") {
        return probeGeocodeService(baseUrl, signal);
    }
    return { ok: true, mode: "metadata-only", detail: "No data probe for this endpoint type" };
}

main().catch(e => {
    console.error("Fatal error:", e);
    process.exitCode = 1;
});
