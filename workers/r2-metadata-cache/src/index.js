/**
 * rmp-metadata-cache  –  Cloudflare Worker
 *
 * Endpoints
 * ─────────
 * GET  /metadata          → returns cached metadata JSON from R2
 * GET  /metadata/:encoded → returns metadata for a single service (URL-encoded key)
 * GET  /health            → returns { ok, lastRefresh, layerCount }
 * POST /refresh           → fetches ?f=json from every service in config.json, stores in R2
 *                           Requires header:  X-Refresh-Secret: <REFRESH_SECRET>
 *
 * R2 keys
 * ───────
 * metadata.json           → full blob:  { lastRefresh, layers: { [url]: { ... } } }
 */

const R2_KEY = "metadata.json";
const FETCH_TIMEOUT_MS = 10000;
const MAX_CONCURRENCY = 8;

// ── Helpers ──────────────────────────────────────────────────────────────────

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

async function probeService(url) {
  const sep = url.includes("?") ? "&" : "?";
  const pjsonUrl = url.endsWith("?f=json") ? url : `${url}${sep}f=json`;

  const started = Date.now();
  try {
    const res = await fetchWithTimeout(pjsonUrl, FETCH_TIMEOUT_MS);
    const elapsed = Date.now() - started;

    if (!res.ok) {
      return { url, status: "DOWN", error: `HTTP ${res.status}`, responseMs: elapsed };
    }

    const body = await res.json();

    // Pull out useful metadata fields
    return {
      url,
      status: "UP",
      responseMs: elapsed,
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
      probedAt: new Date().toISOString(),
    };
  } catch (err) {
    return {
      url,
      status: "DOWN",
      error: err?.message || String(err),
      responseMs: Date.now() - started,
      probedAt: new Date().toISOString(),
    };
  }
}

// ── Route handlers ───────────────────────────────────────────────────────────

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

  // 1. Fetch config.json from GitHub
  let config;
  try {
    config = await fetchConfigFromGitHub(env);
  } catch (err) {
    return json({ error: "Failed to fetch config.json", detail: err.message }, 502, env);
  }

  // 2. Collect all service URLs
  const urls = collectServiceUrls(config);

  // 3. Probe each service with bounded concurrency
  const probeFns = urls.map((url) => () => probeService(url));
  const results = await parallelLimit(probeFns, MAX_CONCURRENCY);

  // 4. Build metadata blob
  const layers = {};
  let upCount = 0;
  let downCount = 0;
  for (const r of results) {
    layers[r.url] = r;
    if (r.status === "UP") upCount++;
    else downCount++;
  }

  const blob = {
    lastRefresh: new Date().toISOString(),
    totalLayers: urls.length,
    upCount,
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
      upCount,
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

    return json({ error: "Not found" }, 404, env);
  },
};
