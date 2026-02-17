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
    const pjsonUrl = svc.url.replace(/\/$/, "") + "?f=pjson";
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

        return { status: "UP", responseTimeMs: Date.now() - t0, error: null };
    } catch (e) {
        const msg = e.name === "AbortError" ? `Timeout (${TIMEOUT_MS}ms)` : e.message;
        return { status: "DOWN", responseTimeMs: Date.now() - t0, error: msg };
    } finally {
        clearTimeout(timer);
    }
}

main().catch(e => {
    console.error("Fatal error:", e);
    process.exitCode = 1;
});
