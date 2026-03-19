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
 * "Push" model - the SiYuan plugin (running in the browser) reads the document
 * from SiYuan and uploads it to Bridge. Bridge serves it to ONLYOFFICE.
 * On save, ONLYOFFICE posts the callback to Bridge, which stores the saved file
 * in disk cache. The plugin then fetches the saved file from Bridge and writes it
 * back to SiYuan.
 *
 * Endpoints:
 *   GET  /health                Health check (?detail=true for connectivity info)
 *   GET  /cache?asset=<path>    Return cache metadata for an asset
 *   POST /upload?asset=<path>   Receive file uploaded by plugin, store in disk cache
 *   GET  /proxy/<assetPath>     Serve stored file to ONLYOFFICE
 *   GET  /oo/<path>             Proxy ONLYOFFICE static assets (api.js, web-apps)
 *   GET  /editor                Serve ONLYOFFICE editor HTML page
 *   POST /callback              Receive ONLYOFFICE save callbacks
 *   GET  /saved?asset=<path>    Return saved file for plugin to sync back to SiYuan
 *   POST /cleanup?asset=<path>  Clear dirty/session state (cache retained until TTL)
 *                               Add purge=1 to remove cache entry immediately
 */

const http = require("http");
const https = require("https");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const fsp = fs.promises;

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

function parsePositiveInt(v, fallback) {
  const parsed = parseInt(String(v ?? ""), 10);
  if (!Number.isFinite(parsed) || parsed <= 0) return fallback;
  return parsed;
}

function parseBool(v, fallback = false) {
  const s = String(v ?? "").trim().toLowerCase();
  if (s === "1" || s === "true" || s === "yes" || s === "on") return true;
  if (s === "0" || s === "false" || s === "no" || s === "off") return false;
  return !!fallback;
}

function normalizeTenantId(v, fallback = "") {
  const raw = String(v || "").trim();
  if (!raw) return fallback;
  const cleaned = raw
    .replace(/[^a-zA-Z0-9._-]/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 64);
  return cleaned || fallback;
}

const bridgeUrlRawFromEnv = trimRightSlash(process.env.BRIDGE_URL || "");
const bridgeUrlFromEnv = normalizeBridgeUrlFromEnv(bridgeUrlRawFromEnv);
const CONFIG = {
  port:                  parseInt(process.env.BRIDGE_PORT || "27689", 10),
  siyuanUrl:             trimRightSlash(process.env.SIYUAN_URL || ""),
  siyuanToken:           process.env.SIYUAN_TOKEN || "",
  // Internal address for bridge-side checks/calls (container network, localhost, etc.)
  onlyofficeInternalUrl: trimRightSlash(process.env.ONLYOFFICE_INTERNAL_URL || process.env.ONLYOFFICE_URL || "http://127.0.0.1:27670"),
  // Browser-accessible address for loading ONLYOFFICE api.js in editor page
  onlyofficePublicUrl:   trimRightSlash(process.env.ONLYOFFICE_PUBLIC_URL || process.env.ONLYOFFICE_BROWSER_URL || ""),
  bridgeUrl:             bridgeUrlFromEnv,
  bridgeBasePath:        normalizeBasePath(process.env.BRIDGE_BASE_PATH || inferBasePathFromUrl(bridgeUrlFromEnv)),
  bridgeSecret:          process.env.BRIDGE_SECRET || "",
  defaultTenant:         normalizeTenantId(process.env.BRIDGE_DEFAULT_TENANT || "default", "default"),
  requireTenant:         parseBool(process.env.BRIDGE_REQUIRE_TENANT || "false", false),
  maxFileMB:             parsePositiveInt(process.env.MAX_FILE_MB || process.env.MAX_UPLOAD_MB, 512),
  maxChunkMB:            parsePositiveInt(process.env.MAX_CHUNK_MB, 8),
  cacheTtlHours:         parsePositiveInt(process.env.CACHE_TTL_HOURS, 24),
};

// ---------------------------------------------------------------------------
// Chunk upload sessions stay in memory.
// File cache is persisted on disk.
// ---------------------------------------------------------------------------
const chunkUploadStore = new Map();
const FILE_TTL = CONFIG.cacheTtlHours * 60 * 60 * 1000;
const CHUNK_UPLOAD_TTL = 30 * 60 * 1000; // 30 minutes
const CACHE_ROOT_DIR = path.join(__dirname, "cache");
const CACHE_FILES_DIR = path.join(CACHE_ROOT_DIR, "files");
const CACHE_META_DIR = path.join(CACHE_ROOT_DIR, "meta");

// Auto-cleanup stale cache entries every 10 minutes
setInterval(() => {
  cleanupExpiredFileCacheEntries().catch((err) => {
    log(`Auto-cleanup: file cache cleanup failed: ${err.message}`);
  });
  const now = Date.now();
  for (const [key, val] of chunkUploadStore) {
    if (now - val.lastAccess > CHUNK_UPLOAD_TTL) {
      chunkUploadStore.delete(key);
      log(`Auto-cleanup: expired chunk upload session ${key}`);
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
    pathname === "/cache" ||
    pathname === "/upload" ||
    pathname === "/proxy" ||
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
  const exactEndpoints = ["/health", "/cache", "/upload", "/proxy", "/editor", "/callback", "/saved", "/cleanup"];
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

function resolveTenant(params, req) {
  const fromQuery = params.get("tenant") || "";
  const fromHeader = (req?.headers?.["x-bridge-tenant"] || "").trim();
  const tenant = normalizeTenantId(fromQuery || fromHeader, "");
  if (tenant) return { tenant };
  if (CONFIG.requireTenant) return { error: "Missing tenant parameter" };
  return { tenant: CONFIG.defaultTenant };
}

function fileStoreKey(tenant, asset) {
  return `${tenant}::${asset}`;
}

function cacheEntryId(tenant, asset) {
  return crypto.createHash("sha1").update(fileStoreKey(tenant, asset)).digest("hex");
}

function cacheEntryPathsById(entryId) {
  const bucket = entryId.slice(0, 2);
  return {
    filePath: path.join(CACHE_FILES_DIR, bucket, `${entryId}.bin`),
    metaPath: path.join(CACHE_META_DIR, bucket, `${entryId}.json`),
  };
}

function normalizeCacheMeta(meta) {
  const rawSourceSize = Number(meta.sourceSize);
  const sourceSize = Number.isFinite(rawSourceSize) && rawSourceSize > 0
    ? Math.floor(rawSourceSize)
    : 0;
  return {
    tenant: String(meta.tenant || ""),
    asset: String(meta.asset || ""),
    contentType: String(meta.contentType || "application/octet-stream"),
    dirty: !!meta.dirty,
    version: String(meta.version || ""),
    lastAccess: Number(meta.lastAccess || Date.now()),
    size: Number(meta.size || 0),
    sourceEtag: String(meta.sourceEtag || ""),
    sourceLastModified: String(meta.sourceLastModified || ""),
    sourceSize,
  };
}

function normalizeSourceMetaValue(v) {
  const s = String(v || "").trim();
  return s.slice(0, 512);
}

function parseSourceSize(v) {
  const parsed = Number(v);
  if (!Number.isFinite(parsed) || parsed <= 0) return 0;
  return Math.floor(parsed);
}

function extractSourceMetaFromParams(params) {
  return {
    sourceEtag: normalizeSourceMetaValue(params.get("sourceEtag") || ""),
    sourceLastModified: normalizeSourceMetaValue(params.get("sourceLastModified") || ""),
    sourceSize: parseSourceSize(params.get("sourceSize") || "0"),
  };
}

function cacheEntryLabel(meta, entryId) {
  if (meta?.tenant && meta?.asset) return fileStoreKey(meta.tenant, meta.asset);
  return entryId;
}

async function ensureCacheDirectories() {
  await Promise.all([
    fsp.mkdir(CACHE_FILES_DIR, { recursive: true }),
    fsp.mkdir(CACHE_META_DIR, { recursive: true }),
  ]);
}

async function listFilesRecursively(rootDir) {
  const files = [];
  const stack = [rootDir];
  while (stack.length > 0) {
    const dir = stack.pop();
    let items;
    try {
      items = await fsp.readdir(dir, { withFileTypes: true });
    } catch (err) {
      if (err.code === "ENOENT") continue;
      throw err;
    }
    for (const item of items) {
      const fullPath = path.join(dir, item.name);
      if (item.isDirectory()) {
        stack.push(fullPath);
      } else if (item.isFile()) {
        files.push(fullPath);
      }
    }
  }
  return files;
}

async function deleteFileIfExists(filePath) {
  try {
    await fsp.unlink(filePath);
  } catch (err) {
    if (err.code !== "ENOENT") throw err;
  }
}

async function writeFileAtomic(filePath, data, encoding = undefined) {
  await fsp.mkdir(path.dirname(filePath), { recursive: true });
  const tempPath = `${filePath}.tmp-${process.pid}-${Date.now()}-${crypto.randomBytes(4).toString("hex")}`;
  try {
    if (encoding) {
      await fsp.writeFile(tempPath, data, { encoding });
    } else {
      await fsp.writeFile(tempPath, data);
    }
    await fsp.rename(tempPath, filePath);
  } finally {
    await deleteFileIfExists(tempPath);
  }
}

async function readJsonIfExists(filePath) {
  try {
    const raw = await fsp.readFile(filePath, "utf-8");
    return JSON.parse(raw);
  } catch (err) {
    if (err.code === "ENOENT") return null;
    throw err;
  }
}

async function writeCacheMetaById(entryId, meta) {
  const normalized = normalizeCacheMeta(meta);
  const { metaPath } = cacheEntryPathsById(entryId);
  await writeFileAtomic(metaPath, JSON.stringify(normalized), "utf-8");
  return normalized;
}

async function removeCacheEntryById(entryId) {
  const { filePath, metaPath } = cacheEntryPathsById(entryId);
  await Promise.all([
    deleteFileIfExists(filePath),
    deleteFileIfExists(metaPath),
  ]);
}

function computeBufferVersion(buffer) {
  return crypto.createHash("sha1").update(buffer).digest("hex");
}

function chunkSessionKey(tenant, asset, uploadId) {
  return `${tenant}::${asset}::${uploadId}`;
}

function cleanupChunkSessionsForAsset(tenant, asset) {
  const keyPrefix = `${tenant}::${asset}::`;
  let removed = 0;
  for (const key of chunkUploadStore.keys()) {
    if (!key.startsWith(keyPrefix)) continue;
    chunkUploadStore.delete(key);
    removed += 1;
  }
  return removed;
}

async function storeAssetBuffer(
  tenant,
  asset,
  buffer,
  contentType = "application/octet-stream",
  dirty = false,
  sourceMeta = {}
) {
  const now = Date.now();
  const entryId = cacheEntryId(tenant, asset);
  const { filePath } = cacheEntryPathsById(entryId);
  const normalizedSourceMeta = {
    sourceEtag: normalizeSourceMetaValue(sourceMeta.sourceEtag || ""),
    sourceLastModified: normalizeSourceMetaValue(sourceMeta.sourceLastModified || ""),
    sourceSize: parseSourceSize(sourceMeta.sourceSize || 0),
  };
  const meta = normalizeCacheMeta({
    tenant,
    asset,
    contentType,
    dirty: !!dirty,
    version: computeBufferVersion(buffer),
    lastAccess: now,
    size: buffer.length,
    ...normalizedSourceMeta,
  });
  await writeFileAtomic(filePath, buffer);
  await writeCacheMetaById(entryId, meta);
  cleanupChunkSessionsForAsset(tenant, asset);
  return meta;
}

async function getCacheEntryMeta(tenant, asset, options = {}) {
  const touch = !!options.touch;
  const requireDirty = !!options.requireDirty;
  const entryId = cacheEntryId(tenant, asset);
  const { filePath, metaPath } = cacheEntryPathsById(entryId);
  const rawMeta = await readJsonIfExists(metaPath);
  if (!rawMeta) return null;

  const meta = normalizeCacheMeta(rawMeta);
  if (meta.tenant !== tenant || meta.asset !== asset) {
    await removeCacheEntryById(entryId);
    return null;
  }

  const now = Date.now();
  if (!Number.isFinite(meta.lastAccess) || now - meta.lastAccess > FILE_TTL) {
    await removeCacheEntryById(entryId);
    log(`Auto-cleanup: expired file cache entry for ${fileStoreKey(tenant, asset)}`);
    return null;
  }

  let stat;
  try {
    stat = await fsp.stat(filePath);
  } catch (err) {
    if (err.code === "ENOENT") {
      await removeCacheEntryById(entryId);
      return null;
    }
    throw err;
  }
  if (!stat.isFile()) {
    await removeCacheEntryById(entryId);
    return null;
  }

  meta.size = stat.size;
  if (requireDirty && !meta.dirty) return null;

  if (touch) {
    meta.lastAccess = now;
    await writeCacheMetaById(entryId, meta);
  }

  return {
    ...meta,
    entryId,
    filePath,
    metaPath,
  };
}

async function setCacheEntryDirty(tenant, asset, dirty, expectedVersion = "") {
  const entry = await getCacheEntryMeta(tenant, asset, { touch: false });
  if (!entry) return false;
  if (expectedVersion && entry.version !== expectedVersion) return false;
  entry.dirty = !!dirty;
  entry.lastAccess = Date.now();
  await writeCacheMetaById(entry.entryId, entry);
  return true;
}

async function updateCacheEntryOnCleanup(tenant, asset, sourceMeta = {}) {
  const entry = await getCacheEntryMeta(tenant, asset, { touch: false });
  if (!entry) return { exists: false, updatedSource: false };

  const etag = normalizeSourceMetaValue(sourceMeta.sourceEtag || "");
  const lastModified = normalizeSourceMetaValue(sourceMeta.sourceLastModified || "");
  const sourceSize = parseSourceSize(sourceMeta.sourceSize || 0);
  let updatedSource = false;

  if (etag) {
    entry.sourceEtag = etag;
    updatedSource = true;
  }
  if (lastModified) {
    entry.sourceLastModified = lastModified;
    updatedSource = true;
  }
  if (sourceSize > 0) {
    entry.sourceSize = sourceSize;
    updatedSource = true;
  }

  entry.dirty = false;
  entry.lastAccess = Date.now();
  await writeCacheMetaById(entry.entryId, entry);
  return { exists: true, updatedSource };
}

async function collectCacheStats() {
  const now = Date.now();
  const tenantSet = new Set();
  let entryCount = 0;
  const metaFiles = await listFilesRecursively(CACHE_META_DIR);
  for (const metaPath of metaFiles) {
    if (!metaPath.endsWith(".json")) continue;
    const entryId = path.basename(metaPath, ".json");
    if (!entryId) continue;
    const rawMeta = await readJsonIfExists(metaPath);
    if (!rawMeta) continue;
    const meta = normalizeCacheMeta(rawMeta);
    if (!meta.tenant || !meta.asset) continue;
    if (!Number.isFinite(meta.lastAccess) || now - meta.lastAccess > FILE_TTL) continue;
    const { filePath } = cacheEntryPathsById(entryId);
    let stat;
    try {
      stat = await fsp.stat(filePath);
    } catch (err) {
      if (err.code === "ENOENT") continue;
      throw err;
    }
    if (!stat.isFile()) continue;
    entryCount += 1;
    tenantSet.add(meta.tenant);
  }
  return { entryCount, tenantSet };
}

async function cleanupExpiredFileCacheEntries() {
  const now = Date.now();
  const metaFiles = await listFilesRecursively(CACHE_META_DIR);
  for (const metaPath of metaFiles) {
    if (!metaPath.endsWith(".json")) continue;
    const entryId = path.basename(metaPath, ".json");
    if (!entryId) continue;
    const rawMeta = await readJsonIfExists(metaPath);
    if (!rawMeta) continue;
    const meta = normalizeCacheMeta(rawMeta);
    const { filePath } = cacheEntryPathsById(entryId);
    let fileMissing = false;
    try {
      const stat = await fsp.stat(filePath);
      if (!stat.isFile()) fileMissing = true;
    } catch (err) {
      if (err.code === "ENOENT") {
        fileMissing = true;
      } else {
        throw err;
      }
    }
    const expired = !meta.tenant ||
      !meta.asset ||
      !Number.isFinite(meta.lastAccess) ||
      now - meta.lastAccess > FILE_TTL ||
      fileMissing;
    if (!expired) continue;
    await removeCacheEntryById(entryId);
    log(`Auto-cleanup: expired file cache entry for ${cacheEntryLabel(meta, entryId)}`);
  }
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

function generateDocKey(asset, mode, versionSeed = "", tenant = "default") {
  if (mode === "edit") {
    return `e-${Date.now()}-${crypto.randomBytes(4).toString("hex")}`;
  }
  const normalizedSeed = String(versionSeed || Date.now());
  return crypto.createHash("md5").update(`${tenant}-${asset}-${normalizedSeed}`).digest("hex").slice(0, 20);
}

// ---------------------------------------------------------------------------
// HTTP helpers
// ---------------------------------------------------------------------------
function corsHeaders(req) {
  return {
    "Access-Control-Allow-Origin": req.headers.origin || "*",
    "Access-Control-Allow-Methods": "GET, HEAD, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Bridge-Secret, X-Bridge-Tenant",
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
    let overflowed = false;
    req.on("data", (c) => {
      if (overflowed) return;
      size += c.length;
      if (size > maxSize) {
        overflowed = true;
        req.resume(); // drain remaining data so response can be sent
        const err = new Error("Body too large");
        err.statusCode = 413;
        reject(err);
        return;
      }
      chunks.push(c);
    });
    req.on("end", () => { if (!overflowed) resolve(Buffer.concat(chunks)); });
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
  const cacheStats = await collectCacheStats();
  const tenantSet = new Set(cacheStats.tenantSet);
  for (const session of chunkUploadStore.values()) {
    if (session?.tenant) tenantSet.add(session.tenant);
  }

  const result = {
    status: "ok",
    ts: Date.now(),
    features: {
      chunkUpload: true,
    },
    fileStoreEntries: cacheStats.entryCount,
    chunkUploadEntries: chunkUploadStore.size,
    tenantCount: tenantSet.size,
    defaultTenant: CONFIG.defaultTenant,
    requireTenant: CONFIG.requireTenant,
    maxFileMB: CONFIG.maxFileMB,
    maxChunkMB: CONFIG.maxChunkMB,
    cacheTtlHours: CONFIG.cacheTtlHours,
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
    // Check SiYuan connectivity (optional - may not be reachable in remote setup)
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
      result.siyuan = "not configured (remote setup - plugin handles SiYuan sync)";
    }
  }

  sendJson(res, req, 200, result);
}

// ---------------------------------------------------------------------------
// GET /cache?asset=<path>
// Returns cache metadata for plugin-side upload skip decisions.
// ---------------------------------------------------------------------------
async function handleCache(req, res, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  const tenantResult = resolveTenant(params, req);
  if (tenantResult.error) {
    return sendJson(res, req, 400, { error: tenantResult.error });
  }
  const tenant = tenantResult.tenant;

  const asset = params.get("asset") || "";
  if (!isValidAssetPath(asset)) {
    return sendJson(res, req, 400, { error: "Invalid asset path" });
  }

  const entry = await getCacheEntryMeta(tenant, asset, { touch: true });
  if (!entry) {
    return sendJson(res, req, 200, { cached: false });
  }

  return sendJson(res, req, 200, {
    cached: true,
    dirty: !!entry.dirty,
    version: entry.version,
    size: entry.size,
    contentType: entry.contentType || "application/octet-stream",
    sourceEtag: entry.sourceEtag || "",
    sourceLastModified: entry.sourceLastModified || "",
    sourceSize: Number(entry.sourceSize || 0),
  });
}

// ---------------------------------------------------------------------------
// POST /upload?asset=<path>
// Plugin uploads the document file to Bridge before opening the editor.
// ---------------------------------------------------------------------------
async function handleChunkUpload(req, res, params, asset, tenant) {
  const uploadId = String(params.get("uploadId") || "").trim();
  const chunkIndex = parseInt(params.get("chunkIndex") || "-1", 10);
  const totalChunks = parseInt(params.get("totalChunks") || "0", 10);
  const contentType = String(params.get("contentType") || "").trim() || "application/octet-stream";
  const sourceMeta = extractSourceMetaFromParams(params);

  if (!uploadId ||
      !Number.isInteger(chunkIndex) ||
      !Number.isInteger(totalChunks) ||
      chunkIndex < 0 ||
      totalChunks <= 0 ||
      totalChunks > 50000 ||
      chunkIndex >= totalChunks) {
    return sendJson(res, req, 400, { error: "Invalid chunk upload parameters" });
  }

  let buffer;
  try {
    const maxChunkBytes = CONFIG.maxChunkMB * 1024 * 1024;
    buffer = await readBody(req, maxChunkBytes);
  } catch (err) {
    if (err.statusCode === 413) {
      return sendJson(res, req, 413, { error: `Chunk too large (max ${CONFIG.maxChunkMB} MB)` });
    }
    throw err;
  }

  const key = chunkSessionKey(tenant, asset, uploadId);
  const now = Date.now();
  let session = chunkUploadStore.get(key);

  if (!session) {
    session = {
      tenant,
      asset,
      uploadId,
      totalChunks,
      chunks: new Array(totalChunks),
      receivedCount: 0,
      bytesReceived: 0,
      contentType,
      sourceMeta,
      lastAccess: now,
    };
    chunkUploadStore.set(key, session);
  } else {
    if (session.totalChunks !== totalChunks) {
      return sendJson(res, req, 409, { error: "Chunk total mismatch for this upload session" });
    }
    session.lastAccess = now;
    if (contentType && session.contentType === "application/octet-stream") {
      session.contentType = contentType;
    }
    if (!session.sourceMeta || typeof session.sourceMeta !== "object") {
      session.sourceMeta = sourceMeta;
    } else {
      if (!session.sourceMeta.sourceEtag && sourceMeta.sourceEtag) {
        session.sourceMeta.sourceEtag = sourceMeta.sourceEtag;
      }
      if (!session.sourceMeta.sourceLastModified && sourceMeta.sourceLastModified) {
        session.sourceMeta.sourceLastModified = sourceMeta.sourceLastModified;
      }
      if (!session.sourceMeta.sourceSize && sourceMeta.sourceSize) {
        session.sourceMeta.sourceSize = sourceMeta.sourceSize;
      }
    }
  }

  if (!session.chunks[chunkIndex]) {
    session.chunks[chunkIndex] = buffer;
    session.receivedCount += 1;
    session.bytesReceived += buffer.length;
  }

  const maxFileBytes = CONFIG.maxFileMB * 1024 * 1024;
  if (session.bytesReceived > maxFileBytes) {
    chunkUploadStore.delete(key);
    return sendJson(res, req, 413, { error: `File too large (max ${CONFIG.maxFileMB} MB)` });
  }

  if (session.receivedCount < session.totalChunks) {
    return sendJson(res, req, 200, {
      ok: true,
      chunked: true,
      partial: true,
      chunkIndex,
      received: session.receivedCount,
      totalChunks: session.totalChunks,
    });
  }

  for (let i = 0; i < session.totalChunks; i++) {
    if (!session.chunks[i]) {
      return sendJson(res, req, 409, { error: `Missing chunk ${i}` });
    }
  }

  const merged = Buffer.concat(session.chunks);
  chunkUploadStore.delete(key);
  await storeAssetBuffer(
    tenant,
    asset,
    merged,
    session.contentType || "application/octet-stream",
    false,
    session.sourceMeta || sourceMeta
  );
  log(`Upload(chunked): tenant=${tenant}, stored "${asset}" (${merged.length} bytes, ${session.totalChunks} chunks)`);
  return sendJson(res, req, 200, { ok: true, chunked: true, size: merged.length });
}

async function handleUpload(req, res, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  const tenantResult = resolveTenant(params, req);
  if (tenantResult.error) {
    return sendJson(res, req, 400, { error: tenantResult.error });
  }
  const tenant = tenantResult.tenant;

  const asset = params.get("asset") || "";
  if (!isValidAssetPath(asset)) {
    return sendJson(res, req, 400, { error: "Invalid asset path" });
  }

  const isChunkedMode = params.get("chunked") === "1" || params.has("uploadId") || params.has("chunkIndex");
  const sourceMeta = extractSourceMetaFromParams(params);
  if (isChunkedMode) {
    return await handleChunkUpload(req, res, params, asset, tenant);
  }

  try {
    const maxFileBytes = CONFIG.maxFileMB * 1024 * 1024;
    const buffer = await readBody(req, maxFileBytes);
    const contentType = req.headers["content-type"] || "application/octet-stream";
    await storeAssetBuffer(tenant, asset, buffer, contentType, false, sourceMeta);
    log(`Upload: tenant=${tenant}, stored "${asset}" (${buffer.length} bytes)`);
    sendJson(res, req, 200, { ok: true, size: buffer.length });
  } catch (err) {
    if (err.statusCode === 413) {
      return sendJson(res, req, 413, { error: `File too large (max ${CONFIG.maxFileMB} MB)` });
    }
    log(`Upload error: ${err.message}`);
    sendJson(res, req, 500, { error: err.message });
  }
}

// ---------------------------------------------------------------------------
// GET /proxy/<assetPath>
// Serve file to ONLYOFFICE. Checks disk cache first (push model),
// then falls back to fetching from SiYuan (co-located setup).
// ---------------------------------------------------------------------------
async function handleProxy(req, res, assetPath, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  const tenantResult = resolveTenant(params, req);
  if (tenantResult.error) {
    return sendJson(res, req, 400, { error: tenantResult.error });
  }
  const tenant = tenantResult.tenant;
  if (!isValidAssetPath(assetPath)) {
    return sendJson(res, req, 400, { error: "Invalid asset path" });
  }

  // Serve from disk cache (push model - browser uploaded this file)
  const entry = await getCacheEntryMeta(tenant, assetPath, { touch: true });
  if (entry) {
    log(`Proxy: tenant=${tenant}, serving "${assetPath}" from file cache (${entry.size} bytes)`);
    res.writeHead(200, {
      "Content-Type": entry.contentType || "application/octet-stream",
      "Content-Length": entry.size,
      "Cache-Control": "no-cache",
      ...corsHeaders(req),
    });
    if (req.method === "HEAD") return res.end();
    const stream = fs.createReadStream(entry.filePath);
    stream.on("error", (err) => {
      log(`Proxy stream error: ${err.message}`);
      if (!res.headersSent) {
        sendJson(res, req, 500, { error: `Proxy stream failed: ${err.message}` });
      } else {
        res.destroy(err);
      }
    });
    return stream.pipe(res);
  }

  // Fallback: proxy from SiYuan (for co-located setups where Bridge can reach SiYuan)
  if (!CONFIG.siyuanUrl) {
    return sendJson(res, req, 404, { error: "File not found in cache and SIYUAN_URL not configured" });
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
async function handleEditor(req, res, params, routeBasePath = "") {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  const tenantResult = resolveTenant(params, req);
  if (tenantResult.error) {
    return sendHtml(res, req, 400, "<h1>Missing tenant parameter</h1>");
  }
  const tenant = tenantResult.tenant;

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
  const storeEntry = await getCacheEntryMeta(tenant, asset, { touch: false });
  const versionSeed = storeEntry?.version || params.get("v") || Date.now();

  if (ext === "pdf" && mode === "edit") mode = "view";

  // browserBase: where the browser can reach bridge (used for /oo proxy script)
  // onlyofficeBase: where ONLYOFFICE server can reach bridge (proxy/callback URLs)
  const bridgeBrowserBase = resolveBrowserBridgeUrl(req, routeBasePath);
  const bridgeOnlyofficeBase = resolveBridgeUrl(req, routeBasePath);
  const tenantParam = `&tenant=${encodeURIComponent(tenant)}`;
  const secretParam = CONFIG.bridgeSecret ? `&secret=${encodeURIComponent(CONFIG.bridgeSecret)}` : "";

  const documentUrl = `${bridgeOnlyofficeBase}/proxy?asset=${encodeURIComponent(asset)}${tenantParam}&t=${Date.now()}${secretParam}`;
  const callbackUrl = mode === "edit"
    ? `${bridgeOnlyofficeBase}/callback?asset=${encodeURIComponent(asset)}${tenantParam}${secretParam}`
    : undefined;
  const key = generateDocKey(asset, mode, versionSeed, tenant);

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
        // Manual save mode: edits are saved when user clicks save.
        autosave: false,
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
      var docEditor = null;
      var config = ${JSON.stringify(config)};

      config.events = {
        onDocumentReady: function() {
          statusEl.className = "hidden";
        },
        onDocumentStateChange: function(event) {
          // event.data === true  -> document has unsaved changes
          // event.data === false -> document is saved / no changes
          if (event && event.data === true) {
            hasChanges = true;
            try {
              window.parent.postMessage({ type: "oo-bridge-dirty", asset: assetPath }, "*");
            } catch (e) {}
          } else if (event && event.data === false && hasChanges) {
            hasChanges = false;
            try {
              window.parent.postMessage({ type: "oo-bridge-clean", asset: assetPath }, "*");
            } catch (e) {}
            // Notify the parent frame (SiYuan plugin) that a save occurred
            try {
              window.parent.postMessage({ type: "oo-bridge-saved", asset: assetPath }, "*");
            } catch (e) {}
          }
        },
        onRequestClose: function() {
          try {
            window.parent.postMessage({ type: "oo-bridge-request-close-ok", asset: assetPath }, "*");
          } catch (e) {}
        },
        onError: function(e) {
          statusEl.textContent = "Error: " + (e && e.data ? JSON.stringify(e.data) : "Unknown");
          statusEl.className = "error";
        }
      };

      window.addEventListener("beforeunload", function(e) {
        if (!hasChanges) return;
        e.preventDefault();
        e.returnValue = "";
      });

      window.addEventListener("message", function(event) {
        var msg = event && event.data;
        if (!msg || msg.type !== "oo-bridge-request-close") return;
        if (msg.asset && msg.asset !== assetPath) return;
        try {
          if (docEditor && typeof docEditor.requestClose === "function") {
            docEditor.requestClose();
          } else {
            window.parent.postMessage({ type: "oo-bridge-request-close-ok", asset: assetPath }, "*");
          }
        } catch (e) {}
      });

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
          docEditor = new DocsAPI.DocEditor("editor", config);
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
  const tenantResult = resolveTenant(params, req);
  if (tenantResult.error) {
    return sendJson(res, req, 400, { error: tenantResult.error });
  }
  const tenant = tenantResult.tenant;

  const assetPath = params.get("asset") || "";

  let data;
  try {
    const bodyBuf = await readBody(req);
    data = JSON.parse(bodyBuf.toString("utf-8"));
  } catch (err) {
    log(`Callback: parse error - ${err.message}`);
    return sendJson(res, req, 200, { error: 0 });
  }

  const status = data.status;
  log(`Callback: tenant=${tenant}, status=${status}, asset=${assetPath}`);

  if ((status === 2 || status === 6) && data.url && assetPath) {
    try {
      await downloadAndStore(data.url, assetPath, tenant);
      log(`Callback: tenant=${tenant}, saved "${assetPath}"`);
    } catch (err) {
      log(`Callback: save failed - ${err.message}`);
    }
  } else if (status === 1) {
    log(`Callback: document being edited - ${assetPath}`);
  } else if (status === 4) {
    log(`Callback: closed without changes - ${assetPath}`);
  } else if (status === 3 || status === 7) {
    log(`Callback: ONLYOFFICE reported save error - status=${status}`);
  }

  sendJson(res, req, 200, { error: 0 });
}

// ---------------------------------------------------------------------------
// Download saved file from ONLYOFFICE and store in disk cache.
// ---------------------------------------------------------------------------
async function downloadAndStore(downloadUrl, assetPath, tenant) {
  log(`Downloading saved file from ONLYOFFICE: tenant=${tenant}, url=${downloadUrl}`);
  const DOWNLOAD_TIMEOUT = 60000; // 60 seconds

  const dlRes = await new Promise((resolve, reject) => {
    const parsed = new URL(downloadUrl);
    const mod = parsed.protocol === "https:" ? https : http;
    const req = mod.get(downloadUrl, {}, (res) => resolve(res));
    req.on("error", reject);
    req.setTimeout(DOWNLOAD_TIMEOUT, () => {
      req.destroy(new Error("Download timed out"));
    });
  });

  const chunks = [];
  let bodySize = 0;
  const MAX_SIZE = CONFIG.maxFileMB * 1024 * 1024;
  await new Promise((resolve, reject) => {
    const bodyTimer = setTimeout(() => {
      dlRes.destroy(new Error("Response body read timed out"));
    }, DOWNLOAD_TIMEOUT);
    dlRes.on("data", (chunk) => {
      bodySize += chunk.length;
      if (bodySize > MAX_SIZE) {
        clearTimeout(bodyTimer);
        dlRes.destroy(new Error("Downloaded file exceeds size limit"));
        return;
      }
      chunks.push(chunk);
    });
    dlRes.on("end", () => { clearTimeout(bodyTimer); resolve(); });
    dlRes.on("error", (err) => { clearTimeout(bodyTimer); reject(err); });
  });

  const fileBuffer = Buffer.concat(chunks);
  log(`Downloaded ${fileBuffer.length} bytes`);
  const downloadedVersion = computeBufferVersion(fileBuffer);

  const existingEntry = await getCacheEntryMeta(tenant, assetPath, { touch: false });
  if (existingEntry && existingEntry.version === downloadedVersion) {
    existingEntry.lastAccess = Date.now();
    existingEntry.size = fileBuffer.length;
    await writeCacheMetaById(existingEntry.entryId, existingEntry);
    log(
      `Callback: duplicate saved content ignored, tenant=${tenant}, asset=${assetPath}, version=${downloadedVersion.slice(0, 12)}, dirty=${existingEntry.dirty}`
    );
    return;
  }
  const sourceMeta = existingEntry ? {
    sourceEtag: existingEntry.sourceEtag || "",
    sourceLastModified: existingEntry.sourceLastModified || "",
    sourceSize: Number(existingEntry.sourceSize || 0),
  } : {};

  // Always store in file cache so the plugin can pull it back to SiYuan
  await storeAssetBuffer(tenant, assetPath, fileBuffer, "application/octet-stream", true, sourceMeta);
  log(`Stored saved file in cache: tenant=${tenant}, asset=${assetPath}`);
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
  const tenantResult = resolveTenant(params, req);
  if (tenantResult.error) {
    return sendJson(res, req, 400, { error: tenantResult.error });
  }
  const tenant = tenantResult.tenant;

  const asset = params.get("asset") || "";
  if (!isValidAssetPath(asset)) {
    return sendJson(res, req, 400, { error: "Invalid asset path" });
  }

  const entry = await getCacheEntryMeta(tenant, asset, { touch: true, requireDirty: true });
  if (!entry) {
    // No pending saved version
    res.writeHead(204, {
      "Cache-Control": "no-store, no-cache, must-revalidate",
      Pragma: "no-cache",
      Expires: "0",
      ...corsHeaders(req),
    });
    return res.end();
  }

  // Return the saved file; clear dirty flag only after response is flushed
  log(`Saved: tenant=${tenant}, serving saved version of "${asset}" to plugin`);

  res.writeHead(200, {
    "Content-Type": entry.contentType || "application/octet-stream",
    "Content-Length": entry.size,
    "Cache-Control": "no-store, no-cache, must-revalidate",
    Pragma: "no-cache",
    Expires: "0",
    ...corsHeaders(req),
  });
  if (req.method === "HEAD") {
    return res.end();
  }

  res.once("finish", () => {
    setCacheEntryDirty(tenant, asset, false, entry.version).catch((err) => {
      log(`Saved: failed to clear dirty flag for "${asset}": ${err.message}`);
    });
  });
  const stream = fs.createReadStream(entry.filePath);
  stream.on("error", (err) => {
    log(`Saved stream error: ${err.message}`);
    if (!res.headersSent) {
      sendJson(res, req, 500, { error: `Saved stream failed: ${err.message}` });
    } else {
      res.destroy(err);
    }
  });
  stream.pipe(res);
}

// ---------------------------------------------------------------------------
// POST /cleanup?asset=<path>
// Plugin calls this after syncing back to SiYuan.
// Keep cached file on disk and only clear dirty/session state by default.
// If purge=1 is provided, remove cache entry immediately.
// ---------------------------------------------------------------------------
async function handleCleanup(req, res, params) {
  if (!checkSecret(params, req)) {
    return sendJson(res, req, 403, { error: "Forbidden" });
  }
  const tenantResult = resolveTenant(params, req);
  if (tenantResult.error) {
    return sendJson(res, req, 400, { error: tenantResult.error });
  }
  const tenant = tenantResult.tenant;

  const asset = params.get("asset") || "";
  if (!isValidAssetPath(asset)) {
    return sendJson(res, req, 400, { error: "Invalid asset path" });
  }

  const purge = params.get("purge") === "1" || params.get("purge") === "true";
  if (purge) {
    await removeCacheEntryById(cacheEntryId(tenant, asset));
    const chunkRemoved = cleanupChunkSessionsForAsset(tenant, asset);
    log(`Cleanup: tenant=${tenant}, "${asset}" purged from cache, chunk sessions removed=${chunkRemoved}`);
    return sendJson(res, req, 200, { ok: true, purged: true });
  }

  const sourceMeta = extractSourceMetaFromParams(params);
  const cleanupResult = await updateCacheEntryOnCleanup(tenant, asset, sourceMeta);
  const chunkRemoved = cleanupChunkSessionsForAsset(tenant, asset);
  const stateText = cleanupResult.exists
    ? `cleared dirty flag${cleanupResult.updatedSource ? " and refreshed source meta" : ""}`
    : "not in cache";
  log(`Cleanup: tenant=${tenant}, "${asset}" ${stateText}, chunk sessions removed=${chunkRemoved}`);
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
    if (isReadMethod && pathname === "/cache") {
      return await handleCache(req, res, params);
    }
    if (req.method === "POST" && pathname === "/upload") {
      return await handleUpload(req, res, params);
    }
    if (isReadMethod && pathname === "/editor") {
      return await handleEditor(req, res, params, basePath);
    }
    if (isReadMethod && pathname === "/proxy") {
      const assetPath = params.get("asset") || "";
      return await handleProxy(req, res, assetPath, params);
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

async function startServer() {
  await ensureCacheDirectories();
  server.listen(CONFIG.port, "0.0.0.0", () => {
    log("===========================================");
    log("  ONLYOFFICE Bridge Server");
    log("===========================================");
    log(`  Bridge      : http://0.0.0.0:${CONFIG.port}`);
    log(`  OnlyOffice (internal): ${CONFIG.onlyofficeInternalUrl}`);
    log(`  OnlyOffice (public)  : ${CONFIG.onlyofficePublicUrl || "(same as internal / editor query override)"}`);
    log(`  SiYuan      : ${CONFIG.siyuanUrl || "(not configured - remote setup)"}`);
    log(`  Auth token  : ${CONFIG.siyuanToken ? "configured" : "not set"}`);
    log(`  Secret      : ${CONFIG.bridgeSecret ? "configured" : "not set"}`);
    log(`  Tenant mode : ${CONFIG.requireTenant ? "required" : "optional"} (default=${CONFIG.defaultTenant})`);
    log(`  Base path   : ${CONFIG.bridgeBasePath || "/"}`);
    log(`  Max file MB : ${CONFIG.maxFileMB}`);
    log(`  Max chunk MB: ${CONFIG.maxChunkMB}`);
    log(`  Cache TTL h : ${CONFIG.cacheTtlHours}`);
    log(`  Cache dir   : ${CACHE_ROOT_DIR}`);
    log(`  External URL: ${CONFIG.bridgeUrl || "(auto-detect from Host header)"}`);
    if (bridgeUrlRawFromEnv && !CONFIG.bridgeUrl) {
      log(`  NOTE: Ignored placeholder BRIDGE_URL="${bridgeUrlRawFromEnv}", using request host instead.`);
    }
    log("===========================================");
  });
}

startServer().catch((err) => {
  log(`Failed to start server: ${err.stack || err.message}`);
  process.exit(1);
});
