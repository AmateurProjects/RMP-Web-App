/**
 * OAuth Device Flow Proxy Routes — add these to your existing Cloudflare Worker
 * ================================================================================
 *
 * GitHub's OAuth endpoints (github.com/login/device/code and
 * github.com/login/oauth/access_token) do NOT support CORS, so browser-side
 * JavaScript cannot call them directly. These two routes act as a thin CORS
 * proxy — they forward the request to GitHub and return the response with
 * proper Access-Control-Allow-Origin headers.
 *
 * NO secrets are required. The Device Flow uses only the public Client ID,
 * which the browser sends in the request body.
 *
 * ── How to add ──
 * Open your existing permitting-web-app-cache Worker in the Cloudflare Dashboard
 * (Workers & Pages → permitting-web-app-cache → Edit Code) and paste the handler
 * block below at the TOP of your fetch() handler, before your existing
 * /metadata.json and /refresh routes.
 *
 * ── One-time GitHub OAuth App setup ──
 * 1. Go to  https://github.com/settings/developers
 * 2. Click "New OAuth App"
 * 3. Fill in:
 *    - Application name:  RMP Layer Admin  (or anything)
 *    - Homepage URL:      https://amateurprojects.github.io/RMP-Web-App/admin.html
 *    - Callback URL:      https://github.com  (not used by Device Flow, but required)
 * 4. Click "Register application"
 * 5. On the app page, scroll down and check ✅ "Enable Device Flow"
 * 6. Copy the **Client ID** (starts with "Ov23li...")
 * 7. Paste the Client ID into admin.html where it says GITHUB_CLIENT_ID
 */

// ── Paste inside your Worker's fetch() handler, BEFORE existing routes ──

// CORS preflight for /oauth/* routes
if (request.method === 'OPTIONS' && url.pathname.startsWith('/oauth/')) {
  return new Response(null, {
    status: 204,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400',
    },
  });
}

// POST /oauth/device-code → proxy to GitHub Device Code endpoint
if (url.pathname === '/oauth/device-code' && request.method === 'POST') {
  const body = await request.text();
  const ghRes = await fetch('https://github.com/login/device/code', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
    },
    body: body,
  });
  const data = await ghRes.text();
  return new Response(data, {
    status: ghRes.status,
    headers: {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Origin': '*',
    },
  });
}

// POST /oauth/token → proxy to GitHub Token endpoint
if (url.pathname === '/oauth/token' && request.method === 'POST') {
  const body = await request.text();
  const ghRes = await fetch('https://github.com/login/oauth/access_token', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
    },
    body: body,
  });
  const data = await ghRes.text();
  return new Response(data, {
    status: ghRes.status,
    headers: {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Origin': '*',
    },
  });
}
