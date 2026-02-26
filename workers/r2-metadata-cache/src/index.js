/**
 * permitting-web-app-cache  –  Cloudflare Worker
 *
 * Endpoints
 * ─────────
 * GET  /metadata          → returns cached metadata JSON from R2
 * GET  /metadata/:encoded → returns metadata for a single service (URL-encoded key)
 * GET  /health            → returns { ok, lastRefresh, layerCount }
 * POST /refresh           → fetches ?f=json from every service in config.json, stores in R2
 *                           Requires header:  X-Refresh-Secret: <REFRESH_SECRET>
 * POST /reports           → stores report HTML in R2, returns { id, url, expiresIn }
 * GET  /reports/:id       → serves stored report HTML (public, expires after 30 days)
 * POST /cleanup-reports   → deletes expired reports from R2 (requires X-Refresh-Secret)
 *
 * R2 keys
 * ───────
 * metadata.json           → full blob:  { lastRefresh, layers: { [url]: { ... } } }
 * reports/{id}.html       → shared report HTML documents
 */

const R2_KEY = "metadata.json";
const FETCH_TIMEOUT_MS = 30000;
const FETCH_RETRY_ATTEMPTS = 2;
const MAX_CONCURRENCY = 4;
const DATA_PROBE_LAYER_LIMIT = 5;
const DATA_PROBE_FEATURE_LIMIT = 1;

// ── Report sharing constants ─────────────────────────────────────────────────
const REPORT_KEY_PREFIX = "reports/";
const MAX_REPORT_SIZE = 100 * 1024 * 1024; // 100 MB
const REPORT_TTL_DAYS = 30;

// ── Helpers ──────────────────────────────────────────────────────────────────

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function corsHeaders(env) {
  return {
    "Access-Control-Allow-Origin": env.CORS_ORIGIN || "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-Refresh-Secret",
  };
}

function json(body, status, env) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders(env) },
  });
}

async function fetchWithTimeout(url, ms) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), ms);
  try {
    const res = await fetch(url, { signal: controller.signal });
    return res;
  } finally {
    clearTimeout(timer);
  }
}

async function fetchJsonWithTimeout(url, ms) {
  const res = await fetchWithTimeout(url, ms);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  if (json && json.error) {
    const code = json.error.code != null ? json.error.code : "";
    const msg = json.error.message || "ArcGIS error";
    throw new Error(`ArcGIS ${code}: ${msg}`);
  }
  return json;
}

function isRetryableProbeError(err) {
  const msg = String(err?.message || err || "").toLowerCase();
  return (
    msg.includes("aborted") ||
    msg.includes("abort") ||
    msg.includes("timed out") ||
    msg.includes("timeout") ||
    msg.includes("network") ||
    msg.includes("fetch failed") ||
    msg.includes("connection")
  );
}

async function fetchJsonRobust(url, ms, retries = FETCH_RETRY_ATTEMPTS) {
  let lastErr = null;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      return await fetchJsonWithTimeout(url, ms);
    } catch (err) {
      lastErr = err;
      if (attempt >= retries || !isRetryableProbeError(err)) {
        throw err;
      }
      await new Promise((resolve) => setTimeout(resolve, 300 * (attempt + 1)));
    }
  }
  throw lastErr || new Error("Unknown fetch failure");
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
  // Consider both layers and tables — some FeatureServers expose
  // data only as tables (e.g. USFWS At-Risk Species).
  const layers = Array.isArray(meta?.layers) ? meta.layers : [];
  const tables = Array.isArray(meta?.tables) ? meta.tables : [];
  const all = layers.concat(tables);
  return all
    .filter((l) => {
      if (!Number.isFinite(Number(l?.id))) return false;
      const type = String(l?.type || "").toLowerCase();
      if (type.includes("group")) return false;
      if (type.includes("feature")) return true;
      return !!l?.geometryType || tables.includes(l);
    })
    .map((l) => Number(l.id))
    .slice(0, DATA_PROBE_LAYER_LIMIT);
}

// Note: detectFeaturePresence has been merged into the data probe functions
// to avoid querying the same sublayers twice. Each probe now returns
// { ...probeResult, hasFeaturesNow: true|false|null }.

/**
 * Compute normallyHasFeatures with decay.
 * Once a service has features, it stays true for up to DECAY_THRESHOLD
 * consecutive refreshes where hasFeaturesNow !== true.  After that it
 * resets to null (unknown) so it can be re-evaluated.
 */
const FEATURES_DECAY_THRESHOLD = 6; // ~36 hours at 6-hour refresh schedule

function computeNormallyHasFeatures(previousEntry, hasFeaturesNow) {
  // Current refresh found features → sticky true, reset counter
  if (hasFeaturesNow === true) {
    return { normallyHasFeatures: true, featuresAbsentCount: 0 };
  }

  // No previous history at all → unknown
  if (!previousEntry || previousEntry.normallyHasFeatures == null) {
    return { normallyHasFeatures: null, featuresAbsentCount: 0 };
  }

  // Previous said true but current didn't find features → increment counter
  if (previousEntry.normallyHasFeatures === true) {
    const prevCount = Number(previousEntry.featuresAbsentCount) || 0;
    const newCount = prevCount + 1;
    if (newCount >= FEATURES_DECAY_THRESHOLD) {
      // Decayed — reset to unknown so we stop suppressing warnings
      return { normallyHasFeatures: null, featuresAbsentCount: 0 };
    }
    return { normallyHasFeatures: true, featuresAbsentCount: newCount };
  }

  // Previous was false or null → keep as-is
  return {
    normallyHasFeatures: previousEntry.normallyHasFeatures,
    featuresAbsentCount: Number(previousEntry.featuresAbsentCount) || 0,
  };
}

function extractServiceMetadata(body, probeResult, hasFeaturesNow, normallyHasFeatures, featuresAbsentCount) {
  return {
    dataProbe: probeResult,
    currentVersion: body.currentVersion ?? null,
    serviceDescription: (body.serviceDescription || body.description || "").slice(0, 500),
    type: body.type ?? null,
    geometryType: body.geometryType ?? null,
    capabilities: body.capabilities ?? null,
    fields: (body.fields || []).map((f) => ({
      name: f.name,
      alias: f.alias,
      type: f.type,
    })),
    layers: (body.layers || []).map((l) => ({
      id: l.id,
      name: l.name,
    })),
    spatialReference: body.spatialReference ?? null,
    extent: body.extent ?? body.fullExtent ?? null,
    maxRecordCount: body.maxRecordCount ?? null,
    supportedQueryFormats: body.supportedQueryFormats ?? null,
    advancedQueryCapabilities: body.advancedQueryCapabilities ?? null,
    hasFeaturesNow,
    normallyHasFeatures,
    featuresAbsentCount: featuresAbsentCount ?? 0,
    probedAt: new Date().toISOString(),
  };
}

async function probeFeatureOrMapLayerForGeometry(layerUrl) {
  const base = stripQueryAndSlash(layerUrl);
  const queryUrl =
    `${base}/query?where=1%3D1` +
    `&outFields=OBJECTID` +
    `&returnGeometry=true` +
    `&resultRecordCount=${DATA_PROBE_FEATURE_LIMIT}` +
    `&f=json`;

  let data;
  try {
    data = await fetchJsonRobust(queryUrl, FETCH_TIMEOUT_MS);
  } catch (queryErr) {
    const msg = String(queryErr?.message || queryErr || "").toLowerCase();
    // ArcGIS 400 "not supported" or 500 "error performing query" on large
    // services — metadata was fine, so treat as operational with a note.
    if (msg.includes("400") || msg.includes("500") || msg.includes("not supported")) {
      return {
        ok: true,
        mode: "query",
        detail: `Query not supported or errored (${queryErr?.message}); metadata OK`,
        testedUrl: base,
        hasFeaturesNow: null,
      };
    }
    throw queryErr; // re-throw genuine network failures
  }

  const features = Array.isArray(data?.features) ? data.features : [];

  if (!features.length) {
    return {
      ok: false,
      warn: true,
      mode: "query",
      detail: "Query succeeded but returned 0 features",
      testedUrl: base,
      hasFeaturesNow: false,
    };
  }

  const withGeometry = features.find((f) => !!f?.geometry);
  if (!withGeometry) {
    // Features exist but server omitted geometry from the response.
    // This is common for large polygon datasets (BLM, etc.) — the
    // service is operational, geometry renders via map tile requests.
    return {
      ok: true,
      mode: "query",
      detail: "Features returned (geometry omitted from REST response)",
      testedUrl: base,
      hasFeaturesNow: true,
    };
  }

  return {
    ok: true,
    mode: "query",
    detail: "Feature query returned geometry",
    testedUrl: base,
    hasFeaturesNow: true,
  };
}

async function probeFeatureOrMapRootForGeometry(serviceUrl, serviceMeta) {
  const base = stripQueryAndSlash(serviceUrl);
  const candidateIds = pickCandidateSublayerIds(serviceMeta);

  if (!candidateIds.length) {
    // No layers or tables to probe — metadata was fine, so mark as UP.
    // Some services are structurally valid but have no queryable sublayers.
    return {
      ok: true,
      mode: "query",
      detail: "Service has no queryable sublayers; metadata OK",
      testedUrl: base,
      hasFeaturesNow: null,
    };
  }

  const errors = [];
  const warnings = [];
  let anyHasFeatures = false;
  let anyChecked = false;
  for (const id of candidateIds) {
    const layerUrl = `${base}/${id}`;
    try {
      const pass = await probeFeatureOrMapLayerForGeometry(layerUrl);
      if (pass?.hasFeaturesNow === true) anyHasFeatures = true;
      if (pass?.hasFeaturesNow != null) anyChecked = true;
      if (pass?.ok) {
        return {
          ...pass,
          detail: `${pass.detail} (sublayer ${id})`,
          hasFeaturesNow: true,
        };
      }
      if (pass?.warn) {
        warnings.push(`/${id}: ${pass.detail}`);
        continue;
      }
      return {
        ok: false,
        warn: false,
        mode: "query",
        detail: `Sublayer ${id} probe did not return a valid result`,
        hasFeaturesNow: anyHasFeatures ? true : anyChecked ? false : null,
      };
    } catch (err) {
      errors.push(`/${id}: ${err?.message || String(err)}`);
    }
  }

  const computedHasFeatures = anyHasFeatures ? true : anyChecked ? false : null;

  if (warnings.length && !errors.length) {
    return {
      ok: false,
      warn: true,
      mode: "query",
      detail: `Sublayers reachable but no feature geometry available (${warnings.join("; ")})`,
      testedUrl: base,
      hasFeaturesNow: computedHasFeatures,
    };
  }

  if (warnings.length && errors.length) {
    return {
      ok: false,
      warn: true,
      mode: "query",
      detail: `Partial probe warnings (${warnings.join("; ")}); errors (${errors.join("; ")})`,
      testedUrl: base,
      hasFeaturesNow: computedHasFeatures,
    };
  }

  throw new Error(`No sublayer returned geometry (${errors.join("; ")})`);
}

async function probeImageServer(serviceUrl, serviceMeta) {
  const base = stripQueryAndSlash(serviceUrl);
  const ext = serviceMeta?.extent || serviceMeta?.fullExtent;
  const sr = ext?.spatialReference || serviceMeta?.spatialReference || {};
  const wkid = sr.latestWkid || sr.wkid || 4326;

  if (!ext || [ext.xmin, ext.ymin, ext.xmax, ext.ymax].some((v) => !Number.isFinite(Number(v)))) {
    throw new Error("Image service extent unavailable for export probe");
  }

  const exportUrl =
    `${base}/exportImage` +
    `?bbox=${ext.xmin},${ext.ymin},${ext.xmax},${ext.ymax}` +
    `&bboxSR=${wkid}` +
    `&imageSR=${wkid}` +
    `&size=64,64` +
    `&format=png` +
    `&f=json`;

  const data = await fetchJsonRobust(exportUrl, FETCH_TIMEOUT_MS);
  if (!data?.href && !data?.url) {
    throw new Error("exportImage did not return an image URL");
  }

  return {
    ok: true,
    mode: "exportImage",
    detail: "Image export test returned raster data URL",
    testedUrl: base,
    hasFeaturesNow: null, // N/A for imagery
  };
}

async function probeGeocodeServer(serviceUrl) {
  const base = stripQueryAndSlash(serviceUrl);
  const suggestUrl = `${base}/suggest?text=denver&maxSuggestions=1&f=json`;
  const data = await fetchJsonRobust(suggestUrl, FETCH_TIMEOUT_MS);
  const suggestions = Array.isArray(data?.suggestions) ? data.suggestions : [];

  if (!suggestions.length) {
    return {
      ok: false,
      warn: true,
      mode: "suggest",
      detail: "Geocode suggest returned 0 suggestions",
      testedUrl: base,
      hasFeaturesNow: null, // N/A for geocode
    };
  }

  return {
    ok: true,
    mode: "suggest",
    detail: "Geocode suggest returned candidates",
    testedUrl: base,
    hasFeaturesNow: null, // N/A for geocode
  };
}

async function runDataProbe(url, serviceMeta) {
  const kind = detectServiceKind(url);
  if (kind === "feature-layer" || kind === "map-layer") {
    return probeFeatureOrMapLayerForGeometry(url);
  }
  if (kind === "feature-root" || kind === "map-root") {
    return probeFeatureOrMapRootForGeometry(url, serviceMeta);
  }
  if (kind === "image-root") {
    return probeImageServer(url, serviceMeta);
  }
  if (kind === "geocode") {
    return probeGeocodeServer(url);
  }

  return {
    ok: true,
    mode: "metadata-only",
    detail: "No data probe implemented for this endpoint type",
    testedUrl: stripQueryAndSlash(url),
    hasFeaturesNow: null,
  };
}

/**
 * Run an array of async fns with bounded concurrency.
 */
async function parallelLimit(fns, limit) {
  const results = [];
  let i = 0;
  async function next() {
    while (i < fns.length) {
      const idx = i++;
      results[idx] = await fns[idx]();
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, fns.length) }, () => next()));
  return results;
}

// ── Fetch config.json from GitHub (public raw content) ───────────────────────

async function fetchConfigFromGitHub(env) {
  const repo = env.GITHUB_REPO || "AmateurProjects/RMP-Web-App";
  const branch = env.GITHUB_BRANCH || "main";
  const path = env.CONFIG_PATH || "config.json";
  const rawUrl = `https://raw.githubusercontent.com/${repo}/${branch}/${path}`;

  const res = await fetchWithTimeout(rawUrl, 15000);
  if (!res.ok) throw new Error(`GitHub raw fetch failed: ${res.status} ${res.statusText}`);
  return res.json();
}

// ── Collect all service URLs from config ─────────────────────────────────────

function collectServiceUrls(config) {
  const urls = new Set();
  for (const layer of config.reportLayers || []) {
    if (layer.url) urls.add(layer.url);
  }
  for (const layer of config.selectionLayers || []) {
    if (layer.url) urls.add(layer.url);
  }
  if (config.referenceLayers) {
    for (const val of Object.values(config.referenceLayers)) {
      if (typeof val === "string" && val.startsWith("http")) urls.add(val);
    }
  }
  return [...urls];
}

// ── Probe one service ────────────────────────────────────────────────────────

async function probeService(url, previousEntry = null) {
  const sep = url.includes("?") ? "&" : "?";
  const pjsonUrl = url.endsWith("?f=json") ? url : `${url}${sep}f=json`;

  const started = Date.now();
  try {
    let body;
    try {
      body = await fetchJsonRobust(pjsonUrl, FETCH_TIMEOUT_MS);
    } catch (metaErr) {
      try {
        // Last-chance metadata verification with a longer timeout window
        body = await fetchJsonRobust(pjsonUrl, FETCH_TIMEOUT_MS * 2);
      } catch (metaErr2) {
        const elapsed = Date.now() - started;
        const msg = metaErr2?.message || metaErr?.message || String(metaErr2 || metaErr);
        const { normallyHasFeatures, featuresAbsentCount } =
          computeNormallyHasFeatures(previousEntry, null);
        return {
          url,
          status: "DOWN",
          error: msg,
          responseMs: elapsed,
          dataProbe: {
            ok: false,
            mode: "metadata",
            detail: `Metadata request failed: ${msg}`,
          },
          hasFeaturesNow: null,
          normallyHasFeatures,
          featuresAbsentCount,
          probedAt: new Date().toISOString(),
        };
      }
    }

    // Data-level probe (query geometry / export imagery / geocode suggest)
    // The probe functions now also return hasFeaturesNow, eliminating
    // the separate detectFeaturePresence pass over the same sublayers.
    let probeResult;
    try {
      probeResult = await runDataProbe(url, body);
    } catch (probeErr) {
      const elapsed = Date.now() - started;
      const errDetail = probeErr?.message || String(probeErr);
      const failProbe = { ok: false, mode: "data", detail: errDetail };
      const { normallyHasFeatures, featuresAbsentCount } =
        computeNormallyHasFeatures(previousEntry, null);
      return {
        url,
        status: "WARN",
        error: `Data probe warning: ${errDetail}`,
        responseMs: elapsed,
        ...extractServiceMetadata(body, failProbe, null, normallyHasFeatures, featuresAbsentCount),
      };
    }

    const elapsed = Date.now() - started;
    const hasFeaturesNow = probeResult?.hasFeaturesNow ?? null;
    const { normallyHasFeatures, featuresAbsentCount } =
      computeNormallyHasFeatures(previousEntry, hasFeaturesNow);

    if (probeResult && (probeResult.warn || (probeResult.ok === false))) {
      return {
        url,
        status: "WARN",
        error: probeResult.detail || "Data probe warning",
        responseMs: elapsed,
        ...extractServiceMetadata(body, probeResult, hasFeaturesNow, normallyHasFeatures, featuresAbsentCount),
      };
    }

    // Pull out useful metadata fields
    return {
      url,
      status: "UP",
      responseMs: elapsed,
      ...extractServiceMetadata(body, probeResult, hasFeaturesNow, normallyHasFeatures, featuresAbsentCount),
    };
  } catch (err) {
    const { normallyHasFeatures, featuresAbsentCount } =
      computeNormallyHasFeatures(previousEntry, null);
    return {
      url,
      status: "DOWN",
      error: err?.message || String(err),
      responseMs: Date.now() - started,
      hasFeaturesNow: null,
      normallyHasFeatures,
      featuresAbsentCount,
      probedAt: new Date().toISOString(),
    };
  }
}

// ── Route handlers ───────────────────────────────────────────────────────────

// ── Report sharing ───────────────────────────────────────────────────────────

function generateReportId() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  const arr = new Uint8Array(12);
  crypto.getRandomValues(arr);
  let id = "";
  for (let i = 0; i < 12; i++) id += chars[arr[i] % chars.length];
  return id;
}

async function handleStoreReport(request, env) {
  // Quick reject based on Content-Length header (may be absent or forged)
  const contentLength = parseInt(request.headers.get("content-length") || "0", 10);
  if (contentLength > MAX_REPORT_SIZE) {
    return json({ error: "Report too large", maxBytes: MAX_REPORT_SIZE }, 413, env);
  }

  const html = await request.text();
  if (!html || html.length < 100) {
    return json({ error: "No report content provided" }, 400, env);
  }

  // Enforce size limit on actual body (header can be spoofed or absent)
  const actualSize = new TextEncoder().encode(html).byteLength;
  if (actualSize > MAX_REPORT_SIZE) {
    return json({ error: "Report too large", maxBytes: MAX_REPORT_SIZE, actualBytes: actualSize }, 413, env);
  }

  const id = generateReportId();
  const key = REPORT_KEY_PREFIX + id + ".html";

  await env.METADATA_BUCKET.put(key, html, {
    httpMetadata: { contentType: "text/html; charset=utf-8" },
    customMetadata: {
      createdAt: new Date().toISOString(),
      expiresAt: new Date(Date.now() + REPORT_TTL_DAYS * 86400000).toISOString(),
    },
  });

  const baseUrl = new URL(request.url).origin;
  return json({ id, url: `${baseUrl}/reports/${id}`, expiresIn: `${REPORT_TTL_DAYS} days` }, 201, env);
}

async function handleGetReport(reportId, env) {
  // Sanitize ID
  if (!/^[A-Za-z0-9]{6,20}$/.test(reportId)) {
    return new Response(notFoundPage("Invalid report link."), {
      status: 400,
      headers: { "Content-Type": "text/html; charset=utf-8", ...corsHeaders(env) },
    });
  }

  const key = REPORT_KEY_PREFIX + reportId + ".html";
  const obj = await env.METADATA_BUCKET.get(key);

  if (!obj) {
    return new Response(notFoundPage("This report may have expired or the link may be incorrect."), {
      status: 404,
      headers: { "Content-Type": "text/html; charset=utf-8", ...corsHeaders(env) },
    });
  }

  // Check TTL via custom metadata
  const expiresAt = obj.customMetadata?.expiresAt;
  if (expiresAt && new Date(expiresAt) < new Date()) {
    // Expired — delete and return 404
    await env.METADATA_BUCKET.delete(key);
    return new Response(notFoundPage("This shared report has expired."), {
      status: 404,
      headers: { "Content-Type": "text/html; charset=utf-8", ...corsHeaders(env) },
    });
  }

  const html = await obj.text();
  return new Response(html, {
    status: 200,
    headers: {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "public, max-age=86400",
      "Content-Security-Policy": "default-src 'none'; img-src data: blob:; style-src 'unsafe-inline'; script-src 'unsafe-inline'; font-src data:;",
      "X-Content-Type-Options": "nosniff",
      ...corsHeaders(env),
    },
  });
}

function notFoundPage(message) {
  const safe = escapeHtml(message);
  return `<!doctype html><html><head><meta charset="utf-8"><title>Report Not Found</title>
<style>body{font-family:'Source Sans Pro',sans-serif;background:#f5f0e6;color:#2c2c2c;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;}
.card{background:#fff;border-radius:12px;padding:48px;max-width:480px;text-align:center;box-shadow:0 4px 24px rgba(0,0,0,0.1);}
h1{color:#1a472a;font-size:24px;margin:0 0 12px;}p{color:#5a5a5a;font-size:15px;line-height:1.6;}</style>
</head><body><div class="card"><h1>Report Not Found</h1><p>${safe}</p></div></body></html>`;
}

// ── Expired report cleanup ───────────────────────────────────────────────────

const CLEANUP_BATCH_SIZE = 100;

async function handleCleanupReports(request, env) {
  // Auth check — same secret as /refresh
  const secret = request.headers.get("X-Refresh-Secret");
  if (!secret || secret !== env.REFRESH_SECRET) {
    return json({ error: "Unauthorized" }, 401, env);
  }

  const now = new Date();
  let deleted = 0;
  let checked = 0;
  let cursor = undefined;
  let truncated = true;

  while (truncated) {
    const listOpts = { prefix: REPORT_KEY_PREFIX, limit: CLEANUP_BATCH_SIZE };
    if (cursor) listOpts.cursor = cursor;
    const listed = await env.METADATA_BUCKET.list(listOpts);

    for (const obj of listed.objects) {
      checked++;
      const expiresAt = obj.customMetadata?.expiresAt;
      if (expiresAt && new Date(expiresAt) < now) {
        await env.METADATA_BUCKET.delete(obj.key);
        deleted++;
      }
    }

    truncated = listed.truncated;
    cursor = listed.truncated ? listed.cursor : undefined;
  }

  return json({ ok: true, checked, deleted }, 200, env);
}

// ── Metadata handlers ────────────────────────────────────────────────────────

async function handleGetMetadata(env) {
  const obj = await env.METADATA_BUCKET.get(R2_KEY);
  if (!obj) return json({ error: "No cached metadata yet. Trigger POST /refresh first." }, 404, env);

  const data = await obj.json();
  return json(data, 200, env);
}

async function handleGetSingleLayer(encodedUrl, env) {
  const obj = await env.METADATA_BUCKET.get(R2_KEY);
  if (!obj) return json({ error: "No cached metadata." }, 404, env);

  const data = await obj.json();
  const layerUrl = decodeURIComponent(encodedUrl);
  const entry = data.layers?.[layerUrl];
  if (!entry) return json({ error: "Layer not found in cache.", url: layerUrl }, 404, env);

  return json(entry, 200, env);
}

async function handleHealth(env) {
  const obj = await env.METADATA_BUCKET.get(R2_KEY);
  if (!obj) return json({ ok: false, lastRefresh: null, layerCount: 0 }, 200, env);

  const data = await obj.json();
  const layerCount = data.layers ? Object.keys(data.layers).length : 0;
  return json({ ok: true, lastRefresh: data.lastRefresh, layerCount }, 200, env);
}

async function handleRefresh(request, env) {
  // Auth check
  const secret = request.headers.get("X-Refresh-Secret");
  if (!secret || secret !== env.REFRESH_SECRET) {
    return json({ error: "Unauthorized" }, 401, env);
  }

  // Parse optional batch parameters from query string
  const reqUrl = new URL(request.url);
  const batchOffset = Math.max(0, parseInt(reqUrl.searchParams.get("offset") || "0", 10) || 0);
  const batchLimit  = parseInt(reqUrl.searchParams.get("limit") || "0", 10) || 0; // 0 = all

  // 1. Fetch config.json from GitHub
  let config;
  try {
    config = await fetchConfigFromGitHub(env);
  } catch (err) {
    return json({ error: "Failed to fetch config.json", detail: err.message }, 502, env);
  }

  // 2. Collect all service URLs
  const allUrls = collectServiceUrls(config);

  // Apply batch slice (offset/limit)
  const urls = batchLimit > 0
    ? allUrls.slice(batchOffset, batchOffset + batchLimit)
    : allUrls;

  const isBatch = batchLimit > 0;

  // 2b. Load previous cache for historical context + merge base
  let previousBlob = null;
  let previousLayers = {};
  try {
    const prevObj = await env.METADATA_BUCKET.get(R2_KEY);
    if (prevObj) {
      previousBlob = await prevObj.json();
      previousLayers = (previousBlob && previousBlob.layers) || {};
    }
  } catch {
    previousLayers = {};
  }

  // 3. Probe each service in this batch with bounded concurrency
  const probeFns = urls.map((url) => () => probeService(url, previousLayers[url] || null));
  const results = await parallelLimit(probeFns, MAX_CONCURRENCY);

  // 4. Build / merge metadata blob
  // For batched refreshes, merge new results into existing cache.
  // For full refreshes, start fresh.
  const layers = isBatch && previousBlob ? { ...previousBlob.layers } : {};

  let batchUp = 0;
  let batchWarn = 0;
  let batchDown = 0;
  for (const r of results) {
    layers[r.url] = r;
    if (r.status === "UP") batchUp++;
    else if (r.status === "WARN") batchWarn++;
    else batchDown++;
  }

  // Recompute totals from the full merged layer set
  let upCount = 0;
  let warnCount = 0;
  let downCount = 0;
  for (const entry of Object.values(layers)) {
    if (entry.status === "UP") upCount++;
    else if (entry.status === "WARN") warnCount++;
    else downCount++;
  }

  const blob = {
    lastRefresh: new Date().toISOString(),
    totalLayers: Object.keys(layers).length,
    upCount,
    warnCount,
    downCount,
    layers,
  };

  // 5. Store in R2
  await env.METADATA_BUCKET.put(R2_KEY, JSON.stringify(blob), {
    httpMetadata: { contentType: "application/json" },
  });

  return json(
    {
      ok: true,
      lastRefresh: blob.lastRefresh,
      totalLayers: blob.totalLayers,
      batch: isBatch ? { offset: batchOffset, limit: batchLimit, probed: urls.length } : null,
      batchResults: { up: batchUp, warn: batchWarn, down: batchDown },
      upCount,
      warnCount,
      downCount,
    },
    200,
    env,
  );
}

// ── Main fetch handler ───────────────────────────────────────────────────────

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const { pathname } = url;

    // CORS preflight
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders(env) });
    }

    // ── OAuth Device Flow proxy routes ──
    // GitHub's OAuth endpoints don't support CORS, so we proxy them here.
    // Only the public Client ID is sent — no secrets required.

    if (pathname === "/oauth/device-code" && request.method === "POST") {
      const body = await request.text();
      const ghRes = await fetch("https://github.com/login/device/code", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json",
        },
        body,
      });
      const data = await ghRes.text();
      return new Response(data, {
        status: ghRes.status,
        headers: { "Content-Type": "application/json", ...corsHeaders(env) },
      });
    }

    if (pathname === "/oauth/token" && request.method === "POST") {
      const body = await request.text();
      const ghRes = await fetch("https://github.com/login/oauth/access_token", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json",
        },
        body,
      });
      const data = await ghRes.text();
      return new Response(data, {
        status: ghRes.status,
        headers: { "Content-Type": "application/json", ...corsHeaders(env) },
      });
    }

    // Route
    if (request.method === "GET" && pathname === "/metadata") {
      return handleGetMetadata(env);
    }

    if (request.method === "GET" && pathname.startsWith("/metadata/")) {
      const encoded = pathname.slice("/metadata/".length);
      return handleGetSingleLayer(encoded, env);
    }

    if (request.method === "GET" && pathname === "/health") {
      return handleHealth(env);
    }

    if (request.method === "POST" && pathname === "/refresh") {
      return handleRefresh(request, env);
    }

    if (request.method === "POST" && pathname === "/cleanup-reports") {
      return handleCleanupReports(request, env);
    }

    // ── Report sharing routes ──
    if (request.method === "POST" && pathname === "/reports") {
      return handleStoreReport(request, env);
    }

    if (request.method === "GET" && pathname.startsWith("/reports/")) {
      const reportId = pathname.slice("/reports/".length);
      return handleGetReport(reportId, env);
    }

    return json({ error: "Not found" }, 404, env);
  },
};
