#!/usr/bin/env node
/**
 * SiYuan ONLYOFFICE Bridge Server
 *
 * Lightweight Node.js HTTP server (zero external dependencies) that acts as
 * middleware between ONLYOFFICE Document Server and SiYuan Note.
 *
 * Deployment scenario:
 *   - ONLYOFFICE + Bridge: deployed on a PUBLIC server
 *   - SiYuan: deployed on an INTERNAL network (no public IP)
 *   - Browser: can reach both
 *
 * "Push" model — the SiYuan plugin (running in the browser) reads the document
 * from SiYuan and uploads it to Bridge. Bridge serves it to ONLYOFFICE.
 * On save, ONLYOFFICE posts the callback to Bridge, which stores the saved file
 * in memory. The plugin then fetches the saved file from Bridge and writes it
 * back to SiYuan.
 *
 * Endpoints:
 *   GET  /health                Health check (?detail=true for connectivity info)
 *   POST /upload?asset=<path>   Receive file uploaded by plugin, store in memory
 *   GET  /proxy/<assetPath>     Serve stored file to ONLYOFFICE
 *   GET  /oo/<path>             Proxy ONLYOFFICE static assets (api.js, web-apps)
 *   GET  /editor                Serve ONLYOFFICE editor HTML page
 *   POST /callback              Receive ONLYOFFICE save callbacks
 *   GET  /saved?asset=<path>    Return saved file for plugin to sync back to SiYuan
 *   POST /cleanup?asset=<path>  Remove file from memory
 */

const http = require("http");
const https = require("https");
const crypto = require("crypto");

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------
function trimRightSlash(v) {
  return String(v || "").trim().replace(/\/+$/, "");
}

function normalizeBasePath(v) {
  const raw = String(v || "").trim();
  if (!raw || raw === "/") return "";
  let path = raw;
  if (/^https?:\/\//i.test(path)) {
    try { path = new URL(path).pathname || "/"; } catch { return ""; }
  }
  if (!path.startsWith("/")) path = `/${path}`;
  path = path.replace(/\/{2,}/g, "/").replace(/\/+$/, "");
  return path === "/" ? "" : path;
}

function normalizeHttpUrl(v) {
  const raw = trimRightSlash(v);
  if (!raw) return "";
  if (!/^https?:\/\//i.test(raw)) return "";
  try {
    const parsed = new URL(raw);
    if (parsed.protocol !== "http:" && parsed.protocol !== "https:") return "";
    return trimRightSlash(parsed.toString());
  } catch {
    return "";
  }
}

function inferBasePathFromUrl(v) {
  const raw = String(v || "").trim();
  if (!raw || !/^https?:\/\//i.test(raw)) return "";
  try {
    return normalizeBasePath(new URL(raw).pathname);
  } catch {
    return "";
  }
}

function normalizeBridgeUrlFromEnv(v) {
  const raw = trimRightSlash(v);
  if (!raw) return "";
  const lower = raw.toLowerCase();
  // Common placeholders in examples should not be treated as real URLs.
  if (lower.includes("your_server_ip") || lower.includes("your-server-ip")) {
    return "";
  }
  return raw;
}

const bridgeUrlRawFromEnv = trimRightSlash(process.env.BRIDGE_URL || "");
const bridgeUrlFromEnv = normalizeBridgeUrlFromEnv(bridgeUrlRawFromEnv);
const CONFIG = {
  port:                  parseInt(process.env.BRIDGE_PORT || "6789", 10),
  siyuanUrl:             trimRightSlash(process.env.SIYUAN_URL || ""),
  siyuanToken:           process.env.SIYUAN_TOKEN || "",
  // Internal address for bridge-side checks/calls (container network, localhost, etc.)
  onlyofficeInternalUrl: trimRightSlash(process.env.ONLYOFFICE_INTERNAL_URL || process.env.ONLYOFFICE_URL || "http://127.0.0.1:8080"),
  // Browser-accessible address for loading ONLYOFFICE api.js in editor page
  onlyofficePublicUrl:   trimRightSlash(process.env.ONLYOFFICE_PUBLIC_URL || process.env.ONLYOFFICE_BROWSER_URL || ""),
  bridgeUrl:             bridgeUrlFromEnv,
  bridgeBasePath:        normalizeBasePath(process.env.BRIDGE_BASE_PATH || inferBasePathFromUrl(bridgeUrlFromEnv)),
  bridgeSecret:          process.env.BRIDGE_SECRET || "",
};

// ---------------------------------------------------------------------------
// In-memory file store
// key: assetPath (e.g. "assets/example.docx")
// value: { buffer, contentType, dirty, lastAccess }
//   dirty=true means ONLYOFFICE saved a new version that the plugin has not yet
//   pulled back to SiYuan.
// ---------------------------------------------------------------------------
const fileStore = new Map();
const FILE_TTL = 2 * 60 * 60 * 1000; // 2 hours

// Auto-cleanup stale file entries every 10 minutes
setInterval(() => {
  const now = Date.now();
  for (const [key, val] of fileStore) {
    if (now - val.lastAccess > FILE_TTL) {
      fileStore.delete(key);
      log(`Auto-cleanup: expired file store entry for ${key}`);
    }
  }
}, 10 * 60 * 1000);

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------
function log(msg) {
  console.log(`[${new Date().toISOString()}] ${msg}`);
}

// ---------------------------------------------------------------------------
// URL / path helpers
// ---------------------------------------------------------------------------
function parsePath(rawUrl) {
  const safeUrl = rawUrl || "/";
  const idx = safeUrl.indexOf("?");
  const pathname = idx >= 0 ? safeUrl.slice(0, idx) : safeUrl;
  const search = idx >= 0 ? safeUrl.slice(idx + 1) : "";
  return { pathname, search, params: new URLSearchParams(search) };
}

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function isValidAssetPath(p) {
  if (!p || typeof p !== "string") return false;
  const c = p.replace(/\\/g, "/");
  return c.startsWith("assets/") && !c.includes("..");
}

function normalizePathname(pathname) {
  let p = String(pathname || "/");
  if (!p.startsWith("/")) p = `/${p}`;
  return p.replace(/\/{2,}/g, "/");
}

function isBridgeEndpointPath(pathname) {
  return pathname === "/health" ||
    pathname === "/upload" ||
    pathname === "/editor" ||
    pathname === "/callback" ||
    pathname === "/saved" ||
    pathname === "/cleanup" ||
    pathname.startsWith("/oo/") ||
    pathname.startsWith("/proxy/");
}

function splitBridgeRoute(pathname) {
  const p = normalizePathname(pathname);

  // 1) Respect configured base path first
  if (CONFIG.bridgeBasePath) {
    if (p === CONFIG.bridgeBasePath) {
      return { routedPath: "/", basePath: CONFIG.bridgeBasePath };
    }
    if (p.startsWith(`${CONFIG.bridgeBasePath}/`)) {
      const stripped = p.slice(CONFIG.bridgeBasePath.length) || "/";
      if (isBridgeEndpointPath(stripped)) {
        return { routedPath: stripped, basePath: CONFIG.bridgeBasePath };
      }
    }
  }

  // 2) Plain root paths
  if (isBridgeEndpointPath(p)) {
    return { routedPath: p, basePath: "" };
  }

  // 3) Auto-detect prefixed deployment paths (e.g. /bridge/upload)
  const exactEndpoints = ["/health", "/upload", "/editor", "/callback", "/saved", "/cleanup"];
  for (const endpoint of exactEndpoints) {
    if (!p.endsWith(endpoint)) continue;
    const prefix = normalizeBasePath(p.slice(0, p.length - endpoint.length));
    return { routedPath: endpoint, basePath: prefix };
  }

  const dynamicMarkers = ["/proxy/", "/oo/"];
  for (const marker of dynamicMarkers) {
    const markerIdx = p.indexOf(marker);
    if (markerIdx > 0 && markerIdx + marker.length < p.length) {
      const prefix = normalizeBasePath(p.slice(0, markerIdx));
      const routedPath = p.slice(markerIdx);
      return { routedPath, basePath: prefix };
    }
  }

  return { routedPath: p, basePath: CONFIG.bridgeBasePath || "" };
}

function resolveBrowserBridgeUrl(req, basePath = "") {
  const hostHeader = req.headers["x-forwarded-host"] || req.headers.host || `127.0.0.1:${CONFIG.port}`;
  const host = String(hostHeader).split(",")[0].trim();
  const protoHeader = String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim().toLowerCase();
  const proto = (protoHeader === "http" || protoHeader === "https") ? protoHeader : "http";
  const prefix = normalizeBasePath(basePath);
  return `${proto}://${host}${prefix}`;
}

function resolveBridgeUrl(req, basePath = "") {
  if (CONFIG.bridgeUrl) {
    const configuredPath = inferBasePathFromUrl(CONFIG.bridgeUrl);
    const requestPath = normalizeBasePath(basePath);
    if (!configuredPath && requestPath) {
      return `${CONFIG.bridgeUrl}${requestPath}`;
    }
    return CONFIG.bridgeUrl;
  }
  return resolveBrowserBridgeUrl(req, basePath);
}

// ---------------------------------------------------------------------------
// Security
// ---------------------------------------------------------------------------
function checkSecret(params, req) {
  if (!CONFIG.bridgeSecret) return true;
  const fromQuery = params.get("secret") || "";
  const fromHeader = (req.headers["x-bridge-secret"] || "").trim();
  return fromQuery === CONFIG.bridgeSecret || fromHeader === CONFIG.bridgeSecret;
}

// ---------------------------------------------------------------------------
// Document type helpers
// ---------------------------------------------------------------------------
const CELL_EXT = new Set(["xls","xlsx","xlsm","xltx","xltm","ods","csv"]);
const SLIDE_EXT = new Set(["ppt","pptx","pptm","potx","potm","odp"]);

function getDocumentType(ext) {
  if (CELL_EXT.has(ext)) return "cell";
  if (SLIDE_EXT.has(ext)) return "slide";
  if (ext === "pdf") return "pdf";
  return "word";
}

function generateDocKey(asset, mode) {
  if (mode === "edit") {
    return `e-${Date.now()}-${crypto.randomBytes(4).toString("hex")}`;
  }
  const day = new Date().toISOString().slice(0, 10);
  return crypto.createHash("md5").update(`${asset}-${day}`).digest("hex").slice(0, 20);
}

// ---------------------------------------------------------------------------
// HTTP helpers
// ---------------------------------------------------------------------------
function corsHeaders(req) {
  return {
    "Access-Control-Allow-Origin": req.headers.origin || "*",
    "Access-Control-Allow-Methods": "GET, HEAD, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Bridge-Secret",
    "Access-Control-Allow-Credentials": "true",
  };
}

function sendJson(res, req, code, data) {
  const body = data != null ? JSON.stringify(data) : "";
  res.writeHead(code, { "Content-Type": "application/json; charset=utf-8", ...corsHeaders(req) });
  res.end(body);
}

function sendHtml(res, req, code, html) {
  res.writeHead(code, { "Content-Type": "text/html; charset=utf-8", ...corsHeaders(req) });
  res.end(html);
}

function readBody(req, maxSize = 10 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (c) => {
      size += c.length;
      if (size > maxSize) { req.destroy(new Error("Body too large")); return; }
      chunks.push(c);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function httpGet(targetUrl, headers = {}) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(targetUrl);
    const mod = parsed.protocol === "https:" ? https : http;
    mod.get(targetUrl, { headers }, (res) => resolve(res)).on("error", reject);
  });
}

function httpRequest(options, body) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(options.url);
    const mod = parsed.protocol === "https:" ? https : http;
    const r = mod.request({
      method: options.method || "POST",
      hostname: parsed.hostname,
      port: parsed.port,
      path: parsed.pathname + parsed.search,
      headers: options.headers || {},
    }, (res) => {
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () => resolve({ status: res.statusCode, body: Buffer.concat(chunks) }));
      res.on("error", reject);
    });
    r.on("error", reject);
    if (body) r.write(body);
    r.end();
  });
}

// ---------------------------------------------------------------------------
// GET /health
// ---------------------------------------------------------------------------
async function handleHealth(req, res, params) {
  const result = {
    status: "ok",
    ts: Date.now(),
    fileStoreEntries: fileStore.size,
  };

  if (params.get("detail") === "true") {
    // Check OnlyOffice connectivity
    try {
      const ooRes = await httpGet(`${CONFIG.onlyofficeInternalUrl}/healthcheck`, {});
      const chunks = []; for await (const c of ooRes) chunks.push(c);
      const body = Buffer.concat(chunks).toString();
      result.onlyoffice = body.trim() === "true" ? "ok" : `http ${ooRes.statusCode}`;
    } catch (e) {
      result.onlyoffice = `error: ${e.message}`;
    }
    result.onlyofficeInternalUrl = CONFIG.onlyofficeInternalUrl;
    result.onlyofficePublicUrl = CONFIG.onlyofficePublicUrl || "(same as internal or editor query override)";
    result.bridgeBasePath = CONFIG.bridgeBasePath || "/";
    // Check SiYuan connectivity (optional — may not be reachable in remote setup)
    if (CONFIG.siyuanUrl) {
      try {
        const headers = {};
        if (CONFIG.siyuanToken) headers["Authorization"] = `Token ${CONFIG.siyuanToken}`;
        const syRes = await httpGet(`${CONFIG.siyuanUrl}/api/system/version`, headers);
        const chunks = []; for await (const c of syRes) chunks.push(c);
        result.siyuan = syRes.statusCode === 200 ? "ok" : `http ${syRes.statusCode}`;
      } catch (e) {
        result.siyuan = `unreachable: ${e.message}`;
      }
    } else {
      result.siyuan = "not configured (remote setup — plugin handles SiYuan sync)";
    }
  }

  sendJson(res, req, 200, result);
}

// ---------------------------------------------------------------------------
// POST /upload?asset=<path>
// Plugin uploads the document file to Bridge before opening the editor.
// ---------------------------------------------------------------------------
async function handleUpload(req, res, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  const asset = params.get("asset") || "";
  if (!isValidAssetPath(asset)) {
    return sendJson(res, req, 400, { error: "Invalid asset path" });
  }

  try {
    const buffer = await readBody(req, 200 * 1024 * 1024); // 200 MB max
    const contentType = req.headers["content-type"] || "application/octet-stream";
    fileStore.set(asset, { buffer, contentType, dirty: false, lastAccess: Date.now() });
    log(`Upload: stored "${asset}" (${buffer.length} bytes)`);
    sendJson(res, req, 200, { ok: true, size: buffer.length });
  } catch (err) {
    log(`Upload error: ${err.message}`);
    sendJson(res, req, 500, { error: err.message });
  }
}

// ---------------------------------------------------------------------------
// GET /proxy/<assetPath>
// Serve file to ONLYOFFICE. Checks in-memory store first (push model),
// then falls back to fetching from SiYuan (co-located setup).
// ---------------------------------------------------------------------------
async function handleProxy(req, res, assetPath, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  if (!isValidAssetPath(assetPath)) {
    return sendJson(res, req, 400, { error: "Invalid asset path" });
  }

  // Serve from in-memory store (push model — browser uploaded this file)
  const entry = fileStore.get(assetPath);
  if (entry) {
    log(`Proxy: serving "${assetPath}" from file store (${entry.buffer.length} bytes)`);
    entry.lastAccess = Date.now();
    res.writeHead(200, {
      "Content-Type": entry.contentType || "application/octet-stream",
      "Content-Length": entry.buffer.length,
      "Cache-Control": "no-cache",
      ...corsHeaders(req),
    });
    return res.end(entry.buffer);
  }

  // Fallback: proxy from SiYuan (for co-located setups where Bridge can reach SiYuan)
  if (!CONFIG.siyuanUrl) {
    return sendJson(res, req, 404, { error: "File not found in store and SIYUAN_URL not configured" });
  }

  const targetUrl = `${CONFIG.siyuanUrl}/${assetPath}`;
  log(`Proxy: fetching "${assetPath}" from SiYuan (fallback)`);

  const headers = {};
  if (CONFIG.siyuanToken) headers["Authorization"] = `Token ${CONFIG.siyuanToken}`;

  try {
    const proxyRes = await httpGet(targetUrl, headers);
    const respHeaders = {
      ...corsHeaders(req),
      "Content-Type": proxyRes.headers["content-type"] || "application/octet-stream",
      "Cache-Control": "no-cache",
    };
    if (proxyRes.headers["content-length"]) {
      respHeaders["Content-Length"] = proxyRes.headers["content-length"];
    }
    res.writeHead(proxyRes.statusCode, respHeaders);
    proxyRes.pipe(res);
  } catch (err) {
    log(`Proxy error: ${err.message}`);
    sendJson(res, req, 502, { error: `Proxy failed: ${err.message}` });
  }
}

// ---------------------------------------------------------------------------
// GET /oo/<path>
// Proxy ONLYOFFICE static assets through Bridge so editor can work even when
// browser cannot directly access ONLYOFFICE host/port due policy/network.
// ---------------------------------------------------------------------------
async function handleOnlyOfficeProxy(req, res, ooPath, search) {
  if (!ooPath || typeof ooPath !== "string") {
    return sendJson(res, req, 400, { error: "Missing ONLYOFFICE path" });
  }
  const normalizedPath = ooPath.replace(/^\/+/, "");
  if (!normalizedPath || normalizedPath.includes("..")) {
    return sendJson(res, req, 400, { error: "Invalid ONLYOFFICE path" });
  }
  const queryPart = search ? `?${search}` : "";
  const targetUrl = `${CONFIG.onlyofficeInternalUrl}/${normalizedPath}${queryPart}`;

  try {
    const parsed = new URL(targetUrl);
    const mod = parsed.protocol === "https:" ? https : http;
    const headers = { ...req.headers, host: parsed.host };
    // Let Node calculate correct transfer framing when request body is piped.
    delete headers["content-length"];

    const upstreamReq = mod.request({
      method: req.method || "GET",
      hostname: parsed.hostname,
      port: parsed.port || (parsed.protocol === "https:" ? 443 : 80),
      path: parsed.pathname + parsed.search,
      headers,
    }, (upstreamRes) => {
      const responseHeaders = {
        ...upstreamRes.headers,
        ...corsHeaders(req),
      };
      res.writeHead(upstreamRes.statusCode || 502, responseHeaders);
      if (req.method === "HEAD") {
        upstreamRes.resume();
        return res.end();
      }
      upstreamRes.pipe(res);
    });

    upstreamReq.on("error", (err) => {
      if (!res.headersSent) {
        sendJson(res, req, 502, { error: `ONLYOFFICE proxy failed: ${err.message}` });
      } else {
        res.end();
      }
      log(`ONLYOFFICE proxy error: ${err.message}`);
    });

    if (req.method === "GET" || req.method === "HEAD") {
      upstreamReq.end();
    } else {
      req.pipe(upstreamReq);
    }
  } catch (err) {
    log(`ONLYOFFICE proxy error: ${err.message}`);
    sendJson(res, req, 502, { error: `ONLYOFFICE proxy failed: ${err.message}` });
  }
}

// ---------------------------------------------------------------------------
// GET /editor
// ---------------------------------------------------------------------------
function handleEditor(req, res, params, routeBasePath = "") {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }

  const asset = params.get("asset") || "";
  if (!isValidAssetPath(asset)) {
    return sendHtml(res, req, 400, "<h1>Invalid or missing asset parameter</h1>");
  }

  let mode = params.get("mode") || "view";
  const lang = params.get("lang") || "en-US";
  const userId = params.get("userId") || "siyuan-user";
  const userName = params.get("userName") || "SiYuan User";

  const ext = asset.split(".").pop().toLowerCase();
  const fname = asset.split("/").pop();
  const documentType = getDocumentType(ext);

  if (ext === "pdf" && mode === "edit") mode = "view";

  // browserBase: where the browser can reach bridge (used for /oo proxy script)
  // onlyofficeBase: where ONLYOFFICE server can reach bridge (proxy/callback URLs)
  const bridgeBrowserBase = resolveBrowserBridgeUrl(req, routeBasePath);
  const bridgeOnlyofficeBase = resolveBridgeUrl(req, routeBasePath);
  const secretParam = CONFIG.bridgeSecret ? `&secret=${encodeURIComponent(CONFIG.bridgeSecret)}` : "";

  const documentUrl = `${bridgeOnlyofficeBase}/proxy/${encodeURIComponent(asset)}?t=${Date.now()}${secretParam}`;
  const callbackUrl = mode === "edit"
    ? `${bridgeOnlyofficeBase}/callback?asset=${encodeURIComponent(asset)}${secretParam}`
    : undefined;
  const key = generateDocKey(asset, mode);

  const isEdit = mode === "edit";
  const config = {
    document: {
      fileType: ext,
      key,
      title: fname,
      url: documentUrl,
      permissions: {
        edit: isEdit,
        download: true,
        print: true,
        review: false,
        comment: false,
        fillForms: isEdit,
      },
    },
    documentType,
    editorConfig: {
      mode,
      lang,
      user: { id: userId, name: userName },
      customization: {
        autosave: isEdit,
        forcesave: isEdit,
        chat: false,
        compactHeader: true,
      },
    },
    type: "desktop",
  };
  if (callbackUrl) config.editorConfig.callbackUrl = callbackUrl;

  const ooFromQuery = normalizeHttpUrl(params.get("oo") || params.get("onlyofficeUrl") || "");
  const onlyofficePublicUrl = ooFromQuery || CONFIG.onlyofficePublicUrl || CONFIG.onlyofficeInternalUrl;
  const directApiJsSrc = `${onlyofficePublicUrl}/web-apps/apps/api/documents/api.js`;
  const bridgeApiJsSrc = `${bridgeBrowserBase}/oo/web-apps/apps/api/documents/api.js`;
  const apiJsCandidates = [];
  if (ooFromQuery || CONFIG.onlyofficePublicUrl) {
    apiJsCandidates.push(directApiJsSrc);
    if (bridgeApiJsSrc !== directApiJsSrc) apiJsCandidates.push(bridgeApiJsSrc);
  } else {
    apiJsCandidates.push(bridgeApiJsSrc);
    if (bridgeApiJsSrc !== directApiJsSrc) apiJsCandidates.push(directApiJsSrc);
  }
  // Embed asset path so postMessage carries it
  const assetJson = JSON.stringify(asset);

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(fname)}</title>
  <style>
    html, body { height: 100%; margin: 0; padding: 0; overflow: hidden; background: #f4f5f7; }
    #editor { height: 100%; }
    #status {
      position: fixed; inset: 0; display: flex; align-items: center; justify-content: center;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: #555; font-size: 15px; z-index: 10;
      pointer-events: none;
    }
    #status.hidden { display: none; }
    #status.error { color: #c53030; }
  </style>
</head>
<body>
  <div id="status">Loading ONLYOFFICE editor...</div>
  <div id="editor"></div>
  <script>
    (function() {
      var statusEl = document.getElementById("status");
      var assetPath = ${assetJson};
      var hasChanges = false;
      var config = ${JSON.stringify(config)};

      config.events = {
        onDocumentReady: function() {
          statusEl.className = "hidden";
        },
        onDocumentStateChange: function(event) {
          // event.data === true  → document has unsaved changes
          // event.data === false → document is saved / no changes
          if (event && event.data === true) {
            hasChanges = true;
          } else if (event && event.data === false && hasChanges) {
            hasChanges = false;
            // Notify the parent frame (SiYuan plugin) that a save occurred
            try {
              window.parent.postMessage({ type: "oo-bridge-saved", asset: assetPath }, "*");
            } catch (e) {}
          }
        },
        onError: function(e) {
          statusEl.textContent = "Error: " + (e && e.data ? JSON.stringify(e.data) : "Unknown");
          statusEl.className = "error";
        }
      };

      function loadScript(src) {
        return new Promise(function(resolve, reject) {
          if (window.DocsAPI) { resolve(); return; }
          var s = document.createElement("script");
          s.src = src;
          s.onload = resolve;
          s.onerror = function() { reject(new Error("Failed to load " + src)); };
          document.head.appendChild(s);
        });
      }

      function loadScriptCandidates(candidates) {
        var i = 0;
        var lastErr = null;
        function next() {
          if (window.DocsAPI) return Promise.resolve();
          if (i >= candidates.length) {
            return Promise.reject(lastErr || new Error("No valid ONLYOFFICE script URL"));
          }
          var src = candidates[i++];
          return loadScript(src).catch(function(err) {
            lastErr = err;
            return next();
          });
        }
        return next();
      }

      loadScriptCandidates(${JSON.stringify(apiJsCandidates)})
        .then(function() {
          statusEl.textContent = "Initializing editor...";
          new DocsAPI.DocEditor("editor", config);
        })
        .catch(function(err) {
          statusEl.textContent = "Failed to load ONLYOFFICE: " + err.message;
          statusEl.className = "error";
        });
    })();
  </script>
</body>
</html>`;

  sendHtml(res, req, 200, html);
}

// ---------------------------------------------------------------------------
// POST /callback
// ---------------------------------------------------------------------------
async function handleCallback(req, res, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }

  const assetPath = params.get("asset") || "";

  let data;
  try {
    const bodyBuf = await readBody(req);
    data = JSON.parse(bodyBuf.toString("utf-8"));
  } catch (err) {
    log(`Callback: parse error — ${err.message}`);
    return sendJson(res, req, 200, { error: 0 });
  }

  const status = data.status;
  log(`Callback: status=${status}, asset=${assetPath}`);

  if ((status === 2 || status === 6) && data.url && assetPath) {
    try {
      await downloadAndStore(data.url, assetPath);
      log(`Callback: saved "${assetPath}"`);
    } catch (err) {
      log(`Callback: save failed — ${err.message}`);
    }
  } else if (status === 1) {
    log(`Callback: document being edited — ${assetPath}`);
  } else if (status === 4) {
    log(`Callback: closed without changes — ${assetPath}`);
  } else if (status === 3 || status === 7) {
    log(`Callback: ONLYOFFICE reported save error — status=${status}`);
  }

  sendJson(res, req, 200, { error: 0 });
}

// ---------------------------------------------------------------------------
// Download saved file from ONLYOFFICE and store in memory.
// Also attempts a direct push to SiYuan as a fallback (co-located setup).
// ---------------------------------------------------------------------------
async function downloadAndStore(downloadUrl, assetPath) {
  log(`Downloading saved file from ONLYOFFICE: ${downloadUrl}`);
  const dlRes = await httpGet(downloadUrl, {});
  const chunks = [];
  for await (const chunk of dlRes) chunks.push(chunk);
  const fileBuffer = Buffer.concat(chunks);
  log(`Downloaded ${fileBuffer.length} bytes`);

  // Always store in file store so the plugin can pull it back to SiYuan
  fileStore.set(assetPath, {
    buffer: fileBuffer,
    contentType: "application/octet-stream",
    dirty: true,        // signal: plugin has not yet synced this back to SiYuan
    lastAccess: Date.now(),
  });
  log(`Stored saved file in memory: ${assetPath}`);

  // Optional: try to push directly to SiYuan (works only if Bridge can reach SiYuan)
  if (CONFIG.siyuanUrl) {
    try {
      await pushToSiyuan(fileBuffer, assetPath);
      log(`Also pushed directly to SiYuan: ${assetPath}`);
    } catch (err) {
      log(`Could not push to SiYuan directly (expected for remote setup): ${err.message}`);
    }
  }
}

// ---------------------------------------------------------------------------
// Push file to SiYuan via /api/file/putFile (for co-located setup only)
// ---------------------------------------------------------------------------
async function pushToSiyuan(fileBuffer, assetPath) {
  const filePath = `/data/${assetPath}`;
  const fname = assetPath.split("/").pop();
  const boundary = "----BridgeBoundary" + crypto.randomBytes(8).toString("hex");

  const parts = [];
  parts.push(`--${boundary}\r\nContent-Disposition: form-data; name="path"\r\n\r\n${filePath}\r\n`);
  parts.push(`--${boundary}\r\nContent-Disposition: form-data; name="isDir"\r\n\r\nfalse\r\n`);
  parts.push(`--${boundary}\r\nContent-Disposition: form-data; name="modTime"\r\n\r\n${Date.now()}\r\n`);
  parts.push(`--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${fname}"\r\nContent-Type: application/octet-stream\r\n\r\n`);

  const head = Buffer.from(parts.join(""));
  const tail = Buffer.from(`\r\n--${boundary}--\r\n`);
  const body = Buffer.concat([head, fileBuffer, tail]);

  const headers = {
    "Content-Type": `multipart/form-data; boundary=${boundary}`,
    "Content-Length": body.length,
  };
  if (CONFIG.siyuanToken) headers["Authorization"] = `Token ${CONFIG.siyuanToken}`;

  const putUrl = `${CONFIG.siyuanUrl}/api/file/putFile`;
  const result = await httpRequest({ url: putUrl, method: "POST", headers }, body);
  let resp;
  try {
    resp = JSON.parse(result.body.toString("utf-8"));
  } catch {
    throw new Error(`SiYuan returned non-JSON: ${result.body.toString("utf-8").slice(0, 200)}`);
  }
  if (resp.code !== 0) {
    throw new Error(`SiYuan putFile error: ${resp.msg || JSON.stringify(resp)}`);
  }
}

// ---------------------------------------------------------------------------
// GET /saved?asset=<path>
// Returns the saved file if dirty (ONLYOFFICE saved a new version).
// Returns 204 if no saved changes are pending.
// ---------------------------------------------------------------------------
async function handleSaved(req, res, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  const asset = params.get("asset") || "";
  if (!isValidAssetPath(asset)) {
    return sendJson(res, req, 400, { error: "Invalid asset path" });
  }

  const entry = fileStore.get(asset);
  if (!entry || !entry.dirty) {
    // No pending saved version
    res.writeHead(204, {
      "Cache-Control": "no-store, no-cache, must-revalidate",
      Pragma: "no-cache",
      Expires: "0",
      ...corsHeaders(req),
    });
    return res.end();
  }

  // Return the saved file and clear the dirty flag
  entry.dirty = false;
  entry.lastAccess = Date.now();
  log(`Saved: serving saved version of "${asset}" to plugin`);

  res.writeHead(200, {
    "Content-Type": entry.contentType || "application/octet-stream",
    "Content-Length": entry.buffer.length,
    "Cache-Control": "no-store, no-cache, must-revalidate",
    Pragma: "no-cache",
    Expires: "0",
    ...corsHeaders(req),
  });
  res.end(entry.buffer);
}

// ---------------------------------------------------------------------------
// POST /cleanup?asset=<path>
// Plugin calls this after syncing back to SiYuan, to free memory.
// ---------------------------------------------------------------------------
async function handleCleanup(req, res, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  const asset = params.get("asset") || "";
  const deleted = fileStore.delete(asset);
  log(`Cleanup: "${asset}" ${deleted ? "removed from store" : "was not in store"}`);
  sendJson(res, req, 200, { ok: true });
}

// ---------------------------------------------------------------------------
// HTTP Server
// ---------------------------------------------------------------------------
const server = http.createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, corsHeaders(req));
    return res.end();
  }
  const isReadMethod = req.method === "GET" || req.method === "HEAD";

  const { pathname: requestPathname, search, params } = parsePath(req.url);
  const { routedPath: pathname, basePath } = splitBridgeRoute(requestPathname);

  try {
    if (isReadMethod && pathname === "/health") {
      return await handleHealth(req, res, params);
    }
    if (req.method === "POST" && pathname === "/upload") {
      return await handleUpload(req, res, params);
    }
    if (isReadMethod && pathname === "/editor") {
      return handleEditor(req, res, params, basePath);
    }
    if (isReadMethod && pathname.startsWith("/proxy/")) {
      const assetPath = decodeURIComponent(pathname.slice("/proxy/".length));
      return await handleProxy(req, res, assetPath, params);
    }
    if ((isReadMethod || req.method === "POST") && pathname.startsWith("/oo/")) {
      const ooPath = decodeURIComponent(pathname.slice("/oo/".length));
      return await handleOnlyOfficeProxy(req, res, ooPath, search);
    }
    if (req.method === "POST" && pathname === "/callback") {
      return await handleCallback(req, res, params);
    }
    if (isReadMethod && pathname === "/saved") {
      return await handleSaved(req, res, params);
    }
    if (req.method === "POST" && pathname === "/cleanup") {
      return await handleCleanup(req, res, params);
    }
    sendJson(res, req, 404, { error: "Not found" });
  } catch (err) {
    log(`Server error: ${err.stack || err.message}`);
    sendJson(res, req, 500, { error: err.message });
  }
});

server.listen(CONFIG.port, "0.0.0.0", () => {
  log("===========================================");
  log("  ONLYOFFICE Bridge Server");
  log("===========================================");
  log(`  Bridge      : http://0.0.0.0:${CONFIG.port}`);
  log(`  OnlyOffice (internal): ${CONFIG.onlyofficeInternalUrl}`);
  log(`  OnlyOffice (public)  : ${CONFIG.onlyofficePublicUrl || "(same as internal / editor query override)"}`);
  log(`  SiYuan      : ${CONFIG.siyuanUrl || "(not configured — remote setup)"}`);
  log(`  Auth token  : ${CONFIG.siyuanToken ? "configured" : "not set"}`);
  log(`  Secret      : ${CONFIG.bridgeSecret ? "configured" : "not set"}`);
  log(`  Base path   : ${CONFIG.bridgeBasePath || "/"}`);
  log(`  External URL: ${CONFIG.bridgeUrl || "(auto-detect from Host header)"}`);
  if (bridgeUrlRawFromEnv && !CONFIG.bridgeUrl) {
    log(`  NOTE: Ignored placeholder BRIDGE_URL="${bridgeUrlRawFromEnv}", using request host instead.`);
  }
  log("===========================================");
});
