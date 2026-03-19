/**
 * SiYuan ONLYOFFICE Bridge Plugin — "Office Editor"
 *
 * "Push" model:
 *   1. Plugin reads the document from SiYuan (browser → SiYuan, internal)
 *   2. Plugin uploads it to Bridge (browser → Bridge, public)
 *   3. Bridge serves it to ONLYOFFICE (Bridge → ONLYOFFICE, public/same host)
 *   4. On save: ONLYOFFICE → Bridge callback → Bridge stores on disk cache
 *   5. Plugin pulls saved file from Bridge → writes back to SiYuan
 */

const { Plugin, Dialog, Setting, showMessage, openTab, getActiveEditor, getAllTabs } = require("siyuan");

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const STORAGE_KEY = "settings.json";
const TAB_TYPE    = "office-editor";

const ICON_PREVIEW = "iconOOPreview";
const ICON_EDIT    = "iconOOEdit";
const ICON_EMBED   = "iconOOEmbed";
const ICON_TAB     = "iconOOTab";
const SVG_ICONS = `<symbol id="${ICON_PREVIEW}" viewBox="0 0 24 24">
  <path fill="currentColor" d="M6 2h8l6 6v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2Zm7 1.5V8h4.5L13 3.5ZM8 12h8v1.5H8V12Zm0 3h8v1.5H8V15Zm0 3h5v1.5H8V18Z"/>
</symbol>
<symbol id="${ICON_EDIT}" viewBox="0 0 24 24">
  <path fill="currentColor" d="M4 17.25V20h2.75L17.8 8.94l-2.75-2.75L4 17.25Zm15.71-9.04a1 1 0 0 0 0-1.41l-1.5-1.5a1 1 0 0 0-1.41 0l-1.17 1.17 2.75 2.75 1.33-1.01Z"/>
</symbol>
<symbol id="${ICON_EMBED}" viewBox="0 0 24 24">
  <path fill="currentColor" d="M3 5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2h-6v3h2v2H9v-2h2v-3H5a2 2 0 0 1-2-2V5Zm2 0v9h14V5H5Zm4.6 7L7 9.4 8.4 8l1.2 1.2L11.8 7 13.2 8.4 9.6 12Zm6.8-4h2v4h-2V8Z"/>
</symbol>
<symbol id="${ICON_TAB}" viewBox="0 0 24 24">
  <path fill="currentColor" d="M3 5a2 2 0 0 1 2-2h6l2 2h6a2 2 0 0 1 2 2v2H3V5Zm0 5h18v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-9Z"/>
</symbol>`;

const SUPPORTED_EXTENSIONS = new Set([
  "doc","docx","docm","dotx","dotm","odt","rtf","txt","md",
  "csv","xls","xlsx","xlsm","xltx","xltm","ods",
  "ppt","pptx","pptm","potx","potm","odp","pdf",
]);

const CELL_EXTS  = new Set(["xls","xlsx","xlsm","xltx","xltm","ods","csv"]);
const SLIDE_EXTS = new Set(["ppt","pptx","pptm","potx","potm","odp"]);
const ZIP_BASED_EXTS = new Set([
  "docx","docm","dotx","dotm",
  "xlsx","xlsm","xltx","xltm",
  "pptx","pptm","potx","potm",
  "odt","ods","odp",
]);
const CFB_BASED_EXTS = new Set(["doc","xls","ppt"]);
const TEXT_BASED_EXTS = new Set(["txt","md","csv","rtf"]);

const EXT_MIME = Object.freeze({
  doc:  "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  docm: "application/vnd.ms-word.document.macroEnabled.12",
  dotx: "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
  dotm: "application/vnd.ms-word.template.macroEnabled.12",
  odt:  "application/vnd.oasis.opendocument.text",
  rtf:  "application/rtf",
  txt:  "text/plain; charset=utf-8",
  md:   "text/markdown; charset=utf-8",
  csv:  "text/csv; charset=utf-8",
  xls:  "application/vnd.ms-excel",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
  xltm: "application/vnd.ms-excel.template.macroEnabled.12",
  ods:  "application/vnd.oasis.opendocument.spreadsheet",
  ppt:  "application/vnd.ms-powerpoint",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  pptm: "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
  potx: "application/vnd.openxmlformats-officedocument.presentationml.template",
  potm: "application/vnd.ms-powerpoint.template.macroEnabled.12",
  odp:  "application/vnd.oasis.opendocument.presentation",
  pdf:  "application/pdf",
});

const DEFAULT_SETTINGS = {
  bridgeUrl:     "",
  onlyofficeUrl: "",
  bridgeSecret:  "",
  defaultMode:   "edit",
  enableEdit:    true,
};

const FALLBACK_I18N = Object.freeze({
  "message.syncPullPending":
    "Detected a save event for {{name}}, but it has not been written back to SiYuan yet. Bridge cache is kept to avoid data loss. Please retry shortly.",
  "message.syncWritebackFailed":
    "Failed to write {{name}} back to SiYuan: {{error}}. Bridge cache is kept to avoid data loss. Please resolve the issue and retry.",
  "message.closeBlockedUnsynced":
    "{{name}} has unsaved or unsynced changes, so it cannot be closed yet.",
  "message.requestSaveFailed":
    "Failed to save \"{{name}}\": {{error}}",
});

// ---------------------------------------------------------------------------
// Utility functions
// ---------------------------------------------------------------------------
function getExt(p) {
  const s = String(p || "").trim();
  const i = s.lastIndexOf(".");
  return i >= 0 && i < s.length - 1 ? s.slice(i + 1).toLowerCase() : "";
}

function isSupported(p)  { return !!p && SUPPORTED_EXTENSIONS.has(getExt(p)); }
function isPdf(p)        { return getExt(p) === "pdf"; }
function fileName(p)     { const parts = String(p || "").split("/"); return parts[parts.length - 1] || p; }
function extMime(ext)    { return EXT_MIME[ext] || "application/octet-stream"; }

function encodeAssetPath(p) {
  return String(p || "").split("/").map((seg) => encodeURIComponent(seg)).join("/");
}

function isZipSignature(head) {
  return head.length >= 4 &&
    head[0] === 0x50 && head[1] === 0x4b &&
    (head[2] === 0x03 || head[2] === 0x05 || head[2] === 0x07) &&
    (head[3] === 0x04 || head[3] === 0x06 || head[3] === 0x08);
}

function isCfbSignature(head) {
  return head.length >= 8 &&
    head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0 &&
    head[4] === 0xa1 && head[5] === 0xb1 && head[6] === 0x1a && head[7] === 0xe1;
}

function isPdfSignature(head) {
  return head.length >= 5 &&
    head[0] === 0x25 && head[1] === 0x50 && head[2] === 0x44 && head[3] === 0x46 && head[4] === 0x2d;
}

async function validateSourceBlob(asset, blob) {
  if (!blob || blob.size <= 0) {
    throw new Error("Source file is empty");
  }

  const ext = getExt(asset);
  const head = new Uint8Array(await blob.slice(0, 16).arrayBuffer());

  // If SiYuan returns an HTML/JSON fallback page, fail early with a clear error.
  if (!TEXT_BASED_EXTS.has(ext)) {
    const sniffText = await blob.slice(0, 512).text().catch(() => "");
    const looksHtml = /^\s*(<!doctype html|<html[\s>]|<head[\s>]|<body[\s>])/i.test(sniffText);
    const looksJson = /^\s*\{[\s\S]{0,260}"(code|error|msg)"\s*:/i.test(sniffText);
    if (looksHtml || looksJson) {
      throw new Error(`Source content does not match .${ext || "file"} (possible auth/path issue)`);
    }
  }

  if (ZIP_BASED_EXTS.has(ext) && !isZipSignature(head)) {
    throw new Error(`Source content does not match .${ext} (invalid ZIP signature)`);
  }
  if (CFB_BASED_EXTS.has(ext) && !isCfbSignature(head)) {
    throw new Error(`Source content does not match .${ext} (invalid legacy Office signature)`);
  }
  if (ext === "pdf" && !isPdfSignature(head)) {
    throw new Error("Source content does not match .pdf (invalid PDF signature)");
  }
}

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function parseHttpErrorDetail(raw, maxLen = 180) {
  if (!raw) return "";
  const text = String(raw).trim();
  if (!text) return "";
  try {
    const json = JSON.parse(text);
    if (json && typeof json === "object") {
      const msg = [json.error, json.message, json.msg].find((v) => typeof v === "string" && v.trim());
      if (msg) {
        return msg.trim().replace(/\s+/g, " ").slice(0, maxLen);
      }
    }
  } catch {}
  return text.replace(/\s+/g, " ").slice(0, maxLen);
}

function normUrl(v, fb) {
  let s = String(v || "").trim() || fb || "";
  if (!s) return "";
  if (!/^https?:\/\//i.test(s)) s = `http://${s}`;
  return s.replace(/\/+$/, "");
}

function normAssetPath(raw) {
  if (!raw || typeof raw !== "string") return "";
  let v = raw.trim();
  try { v = decodeURIComponent(v); } catch {}
  v = v.replace(/\\/g, "/").split("#")[0].split("?")[0].trim();
  v = v.replace(/\/\d{14}-[a-z0-9]{7}$/i, "");
  if (v.startsWith("data/assets/")) v = v.slice("data/".length);
  if (v.startsWith("./"))  v = v.slice(2);
  if (v.startsWith("/"))   v = v.slice(1);
  const ai = v.indexOf("assets/");
  if (ai > 0) v = v.slice(ai);
  if (!v.startsWith("assets/"))          return "";
  if (v.includes("../") || v.includes("..\\") || /\/\.\//.test(v)) return "";
  return v;
}

function normalizeTenantId(raw, fallback = "") {
  const input = String(raw || "").trim();
  if (!input) return fallback;
  const cleaned = input
    .replace(/[^a-zA-Z0-9._-]/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 64);
  return cleaned || fallback;
}

function hashFNV1a32Hex(input) {
  const str = String(input || "");
  let hash = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    hash ^= str.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

function extractHref(el) {
  if (!el) return "";
  if (el.tagName === "A") return el.getAttribute("href") || "";
  if (el.dataset?.href)   return el.dataset.href;
  const a = el.querySelector?.("a[href]");
  return a ? (a.getAttribute("href") || "") : "";
}

function hrefToAsset(href) {
  if (!href || typeof href !== "string") return "";
  const s = href.trim();
  if (!s || s.startsWith("siyuan://")) return "";
  if (/^https?:\/\//i.test(s)) {
    try { return normAssetPath(new URL(s).pathname); } catch { return ""; }
  }
  return normAssetPath(s);
}

function assetFromText(text) {
  if (!text) return "";
  const raw = String(text);
  const candidates = raw.match(/(?:data\/)?assets\/[^\s"'`<>]+/ig) || [];
  for (const c of candidates) {
    const cleaned = c.replace(/[),.;:!?]+$/g, "");
    const p = normAssetPath(cleaned);
    if (isSupported(p)) return p;
  }
  return "";
}

function assetFromElement(element) {
  if (!element) return "";

  const directHref = hrefToAsset(extractHref(element));
  if (isSupported(directHref)) return directHref;

  const attrs = ["href", "data-href", "data-src", "src", "data-path", "data-url", "title", "aria-label"];
  for (const key of attrs) {
    const v = element.getAttribute?.(key);
    if (!v) continue;
    const p = hrefToAsset(v) || normAssetPath(v) || assetFromText(v);
    if (isSupported(p)) return p;
  }

  const data = element.dataset || {};
  for (const key of ["href", "path", "src", "url"]) {
    const v = data[key];
    if (!v) continue;
    const p = hrefToAsset(v) || normAssetPath(v) || assetFromText(v);
    if (isSupported(p)) return p;
  }

  const textPath = assetFromText(element.textContent || element.innerText || "");
  if (isSupported(textPath)) return textPath;

  return "";
}

function loadingHtml(text) {
  return `<div class="oo-bridge-loading">
  <div class="oo-bridge-loading__spinner"></div>
  <div class="oo-bridge-loading__text">${escapeHtml(text)}</div>
</div>`;
}

function iframeHtml(frameSrc) {
  return `<div class="oo-bridge-dialog" style="height:100%;display:flex;flex-direction:column;">
  <iframe class="oo-bridge-dialog__frame" src="${escapeHtml(frameSrc)}"
    style="border:0;width:100%;height:100%;flex:1;background:#fff;"
    allow="clipboard-read; clipboard-write; fullscreen"></iframe>
</div>`;
}

// ---------------------------------------------------------------------------
// Plugin
// ---------------------------------------------------------------------------
class OnlyOfficeBridgePlugin extends Plugin {
  constructor(opts) {
    super(opts);
    this.settings = { ...DEFAULT_SETTINGS };
    this._settingEls = {};
    this._dialogs = new Set();
    this._syncInFlight = new Map();
    this._postCloseSyncing = new Set();
    this._bridgePrepInFlight = new Map();
    this._savedSyncTimers = new Map();
    this._saveSignalAssets = new Set();
    this._dirtyAssets = new Set();
    this._syncAlertAt = new Map();
    this._tabRuntime = new WeakMap();
    this._tabCloseActionOnce = new Map();
    this._nativeCloseWaiters = new Map();
    this._bridgeHealthCache = { url: "", at: 0, data: null };
    this._healthCacheTtlMs = 15000;
    this._chunkUploadHint = { url: "", expiresAt: 0 };
    this._chunkUploadHintTtlMs = 20 * 60 * 1000;
    this._chunkHintMinSizeBytes = 2 * 1024 * 1024;
    this._tabRegistered = false;
    this._isUnloading = false;
    this._settingsReady = Promise.resolve();
    this._settingsLoaded = false;
    this._tenantIdCache = "";
    this._onBridgeSavedMessage = this._onBridgeSavedMessage.bind(this);
    this._onLinkMenu = this._onLinkMenu.bind(this);
    this._onContentMenu = this._onContentMenu.bind(this);
    this._onFileAnnotationMenu = this._onFileAnnotationMenu.bind(this);
    this._onDocTreeMenu = this._onDocTreeMenu.bind(this);
    this._onProtyleLoadedStatic = this._onProtyleLoadedStatic.bind(this);
    this._onProtyleLoadedDynamic = this._onProtyleLoadedDynamic.bind(this);
    this._onSwitchProtyle = this._onSwitchProtyle.bind(this);
    this._onEditorContentClick = this._onEditorContentClick.bind(this);
  }

  t(key, params) {
    let raw = this.i18n?.[key] || FALLBACK_I18N[key] || key;
    if (params) {
      raw = raw.replace(/\{\{(\w+)\}\}/g, (_, k) =>
        Object.prototype.hasOwnProperty.call(params, k) ? String(params[k]) : ""
      );
    }
    return raw;
  }

  _isMobile() {
    return !!(
      globalThis?.siyuan?.config?.system?.isMobile ||
      globalThis?.siyuan?.config?.system?.container === "android" ||
      globalThis?.siyuan?.config?.system?.container === "ios"
    );
  }

  // -----------------------------------------------------------------------
  // Lifecycle
  // -----------------------------------------------------------------------
  async onload() {
    this._isUnloading = false;
    this.addIcons(SVG_ICONS);
    if (!this._isMobile()) {
      // addTab must be registered before async work; otherwise restored tabs in new window can be blank.
      this._registerCustomTab();
    }
    this._settingsReady = this._loadSettings().finally(() => {
      this._settingsLoaded = true;
    });
    await this._settingsReady;
    this._initSettings();

    this.addCommand({
      langKey: "command.openSettings",
      langText: this.t("command.openSettings"),
      hotkey: "",
      callback: () => this.openSetting(),
    });
    this.addCommand({
      langKey: "command.openByPath",
      langText: this.t("command.openByPath"),
      hotkey: "",
      callback: () => this._promptOpen(),
    });

    this.eventBus.on("open-menu-link", this._onLinkMenu);
    this.eventBus.on("open-menu-content", this._onContentMenu);
    this.eventBus.on("open-menu-fileannotationref", this._onFileAnnotationMenu);
    this.eventBus.on("open-menu-doctree", this._onDocTreeMenu);
    this.eventBus.on("loaded-protyle-static", this._onProtyleLoadedStatic);
    this.eventBus.on("loaded-protyle-dynamic", this._onProtyleLoadedDynamic);
    this.eventBus.on("switch-protyle", this._onSwitchProtyle);
    window.addEventListener("message", this._onBridgeSavedMessage);
  }

  onLayoutReady() {
    this._hydrateEmbedsFromRoot(document);
    setTimeout(() => this._hydrateEmbedsFromRoot(document), 520);
  }

  openSetting() {
    this._syncInputs();
    super.openSetting();
    const actionEl = this.setting?.dialog?.element?.querySelector(".b3-dialog__action");
    if (actionEl) actionEl.remove();
  }

  onunload() {
    this._isUnloading = true;
    this.eventBus.off("open-menu-link", this._onLinkMenu);
    this.eventBus.off("open-menu-content", this._onContentMenu);
    this.eventBus.off("open-menu-fileannotationref", this._onFileAnnotationMenu);
    this.eventBus.off("open-menu-doctree", this._onDocTreeMenu);
    this.eventBus.off("loaded-protyle-static", this._onProtyleLoadedStatic);
    this.eventBus.off("loaded-protyle-dynamic", this._onProtyleLoadedDynamic);
    this.eventBus.off("switch-protyle", this._onSwitchProtyle);
    window.removeEventListener("message", this._onBridgeSavedMessage);
    for (const timers of this._savedSyncTimers.values()) {
      for (const timer of timers) clearTimeout(timer);
    }
    this._savedSyncTimers.clear();
    this._saveSignalAssets.clear();
    this._dirtyAssets.clear();
    this._syncAlertAt.clear();
    this._tabCloseActionOnce.clear();
    for (const waiter of this._nativeCloseWaiters.values()) {
      if (waiter?.timer) clearTimeout(waiter.timer);
      try { waiter?.resolve?.(false); } catch {}
    }
    this._nativeCloseWaiters.clear();
    this._bridgePrepInFlight.clear();
    this._bridgeHealthCache = { url: "", at: 0, data: null };
    this._chunkUploadHint = { url: "", expiresAt: 0 };
    for (const d of this._dialogs) { try { d.destroy(); } catch {} }
    this._dialogs.clear();
  }

  uninstall() {
    // Delete persisted plugin data when uninstalling.
    this.removeData(STORAGE_KEY).catch((err) => {
      console.warn(`[Office Editor] uninstall removeData("${STORAGE_KEY}") failed:`, err);
    });
  }

  _registerCustomTab() {
    if (this._tabRegistered) return;
    const plugin = this;
    this.addTab({
      type: TAB_TYPE,
      init() {
        plugin._renderCustomTab(this);
      },
      update() {
        plugin._renderCustomTab(this);
      },
      resize() {
        plugin._renderCustomTab(this);
      },
      destroy() {
        const data = this.data || {};
        const runtime = plugin._tabRuntime.get(this);
        if (runtime?.msgHandler) {
          window.removeEventListener("message", runtime.msgHandler);
        }
        if (runtime?.liveSyncTimer) clearInterval(runtime.liveSyncTimer);
        plugin._tabRuntime.delete(this);
        if (data.asset && data.mode === "edit") {
          const frameSrc = plugin._resolveTabFrameSrc(data);
          const closeKey = plugin._tabCloseKey(data.asset, data.mode, frameSrc);
          const closeAction = plugin._tabCloseActionOnce.get(closeKey);
          if (closeAction) {
            plugin._tabCloseActionOnce.delete(closeKey);
            if (closeAction === "discard") {
              plugin._discardAssetChanges(data.asset).catch((err) => {
                console.warn("[Office Editor] discard close cleanup failed:", err);
              });
              return;
            }
            plugin._postCloseSyncAndCleanup(data.asset);
            return;
          }

          if (!plugin._isUnloading &&
              plugin._isAssetCloseBlocked(data.asset) &&
              plugin._dirtyAssets.has(data.asset)) {
            const displayName = data.displayName || fileName(data.asset);
            plugin._openTabWithFrame(data.asset, data.mode, displayName, frameSrc)
              .then(async () => {
                const approved = await plugin._requestEditorNativeClose(data.asset);
                if (!approved) return;
                if (!plugin._closeOfficeTab(data.asset, data.mode, frameSrc, "save")) {
                  plugin._postCloseSyncAndCleanup(data.asset);
                }
              })
              .catch((err) => {
                const msg = plugin.t("message.openFailed", { error: err?.message || String(err) });
                showMessage(msg, 7000, "error");
              });
            return;
          }
          plugin._postCloseSyncAndCleanup(data.asset);
        }
      },
    });
    this._tabRegistered = true;
  }

  // -----------------------------------------------------------------------
  // Settings persistence
  // -----------------------------------------------------------------------
  async _loadSettings() {
    let raw = await this.loadData(STORAGE_KEY);
    if (typeof raw === "string") { try { raw = JSON.parse(raw); } catch { raw = {}; } }
    const d = (raw && typeof raw === "object" && !Array.isArray(raw)) ? raw : {};
    this.settings = {
      bridgeUrl:     normUrl(d.bridgeUrl || d.bridgeBaseUrl || "", ""),
      onlyofficeUrl: normUrl(d.onlyofficeUrl || d.documentServerUrl || d.onlyOfficeUrl || "", ""),
      bridgeSecret:  String(d.bridgeSecret ?? ""),
      // Fixed behavior: always open in edit mode by default and keep edit entries enabled.
      defaultMode:   DEFAULT_SETTINGS.defaultMode,
      enableEdit:    DEFAULT_SETTINGS.enableEdit,
    };
    this._tenantIdCache = "";
  }

  async _saveSettings() {
    await this.saveData(STORAGE_KEY, this.settings);
  }

  // -----------------------------------------------------------------------
  // Helpers
  // -----------------------------------------------------------------------
  _userInfo() {
    const acc = globalThis?.siyuan?.config?.account || {};
    const usr = globalThis?.siyuan?.user || {};
    return {
      id:   String(usr.id   || acc.uid  || acc.id   || "siyuan-user"),
      name: String(usr.userName || usr.nickname || acc.userName || acc.name || "SiYuan User"),
    };
  }

  _lang() {
    const l = globalThis?.siyuan?.config?.lang ||
              globalThis?.siyuan?.config?.appearance?.lang || "en_US";
    return String(l).replace("_", "-");
  }

  _siyuanBaseUrl() {
    const candidates = [
      globalThis?.siyuan?.config?.api?.server,
      globalThis?.siyuan?.config?.api?.baseURL,
      globalThis?.location?.origin,
    ];
    for (const c of candidates) {
      const u = normUrl(c);
      if (/^https?:\/\//i.test(u)) return u;
    }
    return "";
  }

  _tenantId() {
    if (this._tenantIdCache) return this._tenantIdCache;

    const user = this._userInfo();
    const workspaceDir = String(globalThis?.siyuan?.config?.system?.workspaceDir || "");
    const seed = [
      this._siyuanBaseUrl(),
      workspaceDir,
      user.id,
      user.name,
      globalThis?.location?.host || "",
    ].join("|");
    const hashed = hashFNV1a32Hex(seed || "siyuan");
    this._tenantIdCache = normalizeTenantId(`sy-${hashed}`, "sy-default");
    return this._tenantIdCache;
  }

  _secretParam() {
    const params = new URLSearchParams();
    const tenant = this._tenantId();
    if (tenant) params.set("tenant", tenant);
    if (this.settings.bridgeSecret) params.set("secret", this.settings.bridgeSecret);
    const suffix = params.toString();
    return suffix ? `&${suffix}` : "";
  }

  _bridgeAuthPrefix() {
    const params = new URLSearchParams();
    const tenant = this._tenantId();
    if (tenant) params.set("tenant", tenant);
    if (this.settings.bridgeSecret) params.set("secret", this.settings.bridgeSecret);
    const query = params.toString();
    return query ? `${query}&` : "";
  }

  _ensureTabRuntime(tab) {
    let runtime = this._tabRuntime.get(tab);
    if (!runtime) {
      runtime = {};
      this._tabRuntime.set(tab, runtime);
    }
    return runtime;
  }

  _buildFrameSrc(asset, mode, displayName) {
    const user = this._userInfo();
    const params = new URLSearchParams({
      asset, mode,
      lang:     this._lang(),
      userId:   user.id,
      userName: user.name,
      title:    displayName || fileName(asset),
    });
    if (this.settings.onlyofficeUrl) params.set("oo", this.settings.onlyofficeUrl);
    params.set("tenant", this._tenantId());
    if (this.settings.bridgeSecret)  params.set("secret", this.settings.bridgeSecret);
    return `${this.settings.bridgeUrl}/editor?${params}`;
  }

  _resolveTabFrameSrc(tabData = {}) {
    if (tabData.frameSrc) return tabData.frameSrc;
    if (!tabData.asset) return "";
    const displayName = tabData.displayName || fileName(tabData.asset);
    return this._buildFrameSrc(tabData.asset, tabData.mode || "view", displayName);
  }

  async _openTabWithFrame(asset, mode, displayName, frameSrc) {
    await Promise.resolve(openTab({
      app: this.app,
      custom: {
        title: displayName,
        icon: mode === "edit" ? ICON_EDIT : ICON_PREVIEW,
        // must be plugin.name + tab.type
        id: `${this.name}${TAB_TYPE}`,
        data: { asset, mode, frameSrc, displayName },
      },
      openNewTab: true,
    }));
  }

  _renderCustomTab(tab) {
    if (!tab?.element) return;
    const runtime = this._ensureTabRuntime(tab);
    const data = tab.data || {};
    tab.element.style.height = "100%";

    if (!this._settingsLoaded) {
      tab.element.innerHTML = loadingHtml(this.t("loading.connecting"));
      const waitToken = (runtime.waitToken || 0) + 1;
      runtime.waitToken = waitToken;
      this._settingsReady.finally(() => {
        const current = this._tabRuntime.get(tab);
        if (!current || current.waitToken !== waitToken) return;
        this._renderCustomTab(tab);
      });
      return;
    }

    const frameSrc = this._resolveTabFrameSrc(data);
    if (frameSrc && data.asset) {
      const key = `${data.asset}|${data.mode || "view"}|${frameSrc}`;
      const existedFrame = tab.element.querySelector("iframe.oo-bridge-dialog__frame");
      const existedSrc = existedFrame?.getAttribute("src") || "";
      if (runtime.renderKey === key && existedFrame && existedSrc === frameSrc) return;
      if (runtime.renderKey === key && runtime.rendering) return;
      runtime.renderKey = key;
      runtime.rendering = true;
      this._renderCustomTabEditor(tab, data.asset, data.mode || "view", frameSrc, key).catch((err) => {
        this._renderTabError(tab, key, err);
      });
      return;
    }
    runtime.renderKey = "";
    runtime.rendering = false;
    tab.element.innerHTML = loadingHtml(this.t("loading.connecting"));
  }

  async _renderCustomTabEditor(tab, asset, mode, frameSrc, key) {
    const runtime = this._tabRuntime.get(tab);
    if (!runtime || runtime.renderKey !== key) return;

    const renderToken = (runtime.renderToken || 0) + 1;
    runtime.renderToken = renderToken;
    tab.element.innerHTML = loadingHtml(this.t("loading.connecting"));

    const updateLoading = (payload) => {
      const current = this._tabRuntime.get(tab);
      if (!current || current.renderKey !== key || current.renderToken !== renderToken) return;
      const textEl = tab.element.querySelector(".oo-bridge-loading__text");
      if (!textEl) return;
      if (typeof payload === "string") {
        textEl.textContent = this.t(payload);
        return;
      }
      if (payload && payload.type === "uploadProgress") {
        const percent = Math.max(0, Math.min(100, Math.round(Number(payload.percent) || 0)));
        textEl.textContent = `${this.t("loading.uploading")} ${percent}%`;
      }
    };

    await this._prepareAssetOnBridge(asset, updateLoading);

    const current = this._tabRuntime.get(tab);
    if (!current || current.renderKey !== key || current.renderToken !== renderToken) return;
    updateLoading("loading.opening");
    this._attachEditorToTab(tab, asset, mode, frameSrc);
    current.rendering = false;
  }

  _renderTabError(tab, key, err) {
    const runtime = this._tabRuntime.get(tab);
    if (!runtime || runtime.renderKey !== key) return;
    runtime.rendering = false;
    const errMsg = String(err?.message || err);
    const bridgeErr = errMsg.startsWith("HTTP ") || errMsg.includes("Failed to fetch") || errMsg.includes("aborted");
    const msg = bridgeErr
      ? this.t("message.bridgeUnreachable", { url: this.settings.bridgeUrl, error: errMsg })
      : this.t("message.uploadFailed", { error: errMsg });
    tab.element.innerHTML = `<div class="oo-bridge-loading oo-bridge-loading--error">
  <div class="oo-bridge-loading__text">${escapeHtml(msg)}</div>
</div>`;
    console.error("[Office Editor] tab render failed:", err);
    showMessage(msg, 7000, "error");
  }

  _uploadBlobWithProgress(url, bodyBlob, contentType, onProgress) {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open("POST", url, true);
      xhr.responseType = "text";
      xhr.setRequestHeader("Content-Type", contentType || "application/octet-stream");

      xhr.upload.onprogress = (event) => {
        if (!event || !event.lengthComputable) return;
        if (typeof onProgress === "function") {
          onProgress(event.loaded, event.total);
        }
      };
      xhr.onload = () => {
        const status = xhr.status || 0;
        const text = typeof xhr.response === "string"
          ? xhr.response
          : String(xhr.responseText || "");
        resolve({
          ok: status >= 200 && status < 300,
          status,
          text,
        });
      };
      xhr.onerror = () => reject(new Error("Failed to fetch"));
      xhr.onabort = () => reject(new Error("aborted"));
      xhr.send(bodyBlob);
    });
  }

  async _uploadAssetInChunks(asset, fileBlob, contentType, onProgress, sourceValidators = null) {
    const chunkSizes = [2 * 1024 * 1024, 1024 * 1024, 768 * 1024, 512 * 1024, 256 * 1024];
    let lastErr = null;
    const totalSize = Math.max(1, fileBlob.size);

    for (const chunkSize of chunkSizes) {
      const totalChunks = Math.max(1, Math.ceil(fileBlob.size / chunkSize));
      const uploadId = `${Date.now().toString(36)}-${Math.random().toString(16).slice(2, 10)}`;
      let uploadedBytes = 0;
      try {
        for (let i = 0; i < totalChunks; i++) {
          const start = i * chunkSize;
          const end = Math.min(fileBlob.size, start + chunkSize);
          const chunkBlob = fileBlob.slice(start, end);

          const params = new URLSearchParams({
            asset,
            chunked: "1",
            uploadId,
            chunkIndex: String(i),
            totalChunks: String(totalChunks),
            contentType,
          });
          this._appendSourceValidatorParams(params, sourceValidators);
          params.set("tenant", this._tenantId());
          if (this.settings.bridgeSecret) params.set("secret", this.settings.bridgeSecret);

          const resp = await this._uploadBlobWithProgress(
            `${this.settings.bridgeUrl}/upload?${params.toString()}`,
            chunkBlob,
            "application/octet-stream",
            (loaded) => {
              if (typeof onProgress === "function") {
                onProgress(Math.min(totalSize, uploadedBytes + loaded), totalSize);
              }
            }
          );
          if (!resp.ok) {
            const detail = parseHttpErrorDetail(resp.text);
            const suffix = detail ? `: ${detail}` : "";
            const err = new Error(`chunk ${i + 1}/${totalChunks} failed with HTTP ${resp.status}${suffix}`);
            err.httpStatus = resp.status;
            throw err;
          }
          uploadedBytes += chunkBlob.size;
          if (typeof onProgress === "function") {
            onProgress(uploadedBytes, totalSize);
          }
        }
        return { chunkSize, totalChunks };
      } catch (err) {
        lastErr = err;
        if (err?.httpStatus === 413) continue;
        break;
      }
    }
    throw lastErr || new Error("Chunk upload failed");
  }

  _normalizeValidatorValue(v, maxLen = 512) {
    const s = String(v || "").trim();
    if (!s) return "";
    return s.slice(0, maxLen);
  }

  _parsePositiveInt(v) {
    const n = Number(v);
    if (!Number.isFinite(n) || n <= 0) return 0;
    return Math.floor(n);
  }

  _extractSourceValidatorsFromHeaders(headers) {
    if (!headers || typeof headers.get !== "function") {
      return { etag: "", lastModified: "", size: 0 };
    }
    const etag = this._normalizeValidatorValue(headers.get("etag") || "");
    const lastModified = this._normalizeValidatorValue(headers.get("last-modified") || "");
    const size = this._parsePositiveInt(headers.get("content-length") || "0");
    return { etag, lastModified, size };
  }

  _appendSourceValidatorParams(params, validators) {
    if (!params || !validators) return;
    const etag = this._normalizeValidatorValue(validators.etag || "");
    const lastModified = this._normalizeValidatorValue(validators.lastModified || "");
    const size = this._parsePositiveInt(validators.size || 0);
    if (etag) params.set("sourceEtag", etag);
    if (lastModified) params.set("sourceLastModified", lastModified);
    if (size > 0) params.set("sourceSize", String(size));
  }

  async _getBridgeCacheMeta(asset) {
    const sp = this._bridgeAuthPrefix();
    const url = `${this.settings.bridgeUrl}/cache?${sp}asset=${encodeURIComponent(asset)}&t=${Date.now()}`;
    const resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) {
      throw new Error(`Bridge /cache returned HTTP ${resp.status}`);
    }
    let data = null;
    try {
      data = await resp.json();
    } catch {}
    if (!data || typeof data !== "object" || typeof data.cached !== "boolean") {
      throw new Error("Bridge /cache returned an unexpected response");
    }
    return data;
  }

  async _fetchSiyuanSourceValidators(asset) {
    const siyuanBase = this._siyuanBaseUrl();
    if (!siyuanBase) return null;
    const token = globalThis?.siyuan?.config?.api?.token;
    const headers = token ? { "Authorization": `Token ${token}` } : {};
    const sourceUrl = `${siyuanBase}/${encodeAssetPath(asset)}?t=${Date.now()}`;
    const ctrl = new AbortController();
    try {
      const resp = await fetch(sourceUrl, { method: "GET", headers, cache: "no-store", signal: ctrl.signal });
      if (!resp.ok) return null;
      const validators = this._extractSourceValidatorsFromHeaders(resp.headers);
      try { ctrl.abort(); } catch {}
      try { await resp.body?.cancel?.(); } catch {}
      return validators;
    } catch {
      return null;
    }
  }

  async _fetchSiyuanSourceValidatorsAfterWrite(asset) {
    let previous = null;
    try {
      previous = await this._getBridgeCacheMeta(asset);
    } catch {}
    const prevEtag = this._normalizeValidatorValue(previous?.sourceEtag || "");
    const prevLastModified = this._normalizeValidatorValue(previous?.sourceLastModified || "");
    const prevSize = this._parsePositiveInt(previous?.sourceSize || 0);
    const prevStrong = !!(prevEtag || prevLastModified);

    const delays = [0, 200, 600, 1200];
    let fallback = null;
    for (const delay of delays) {
      if (delay > 0) await new Promise((resolve) => setTimeout(resolve, delay));
      const current = await this._fetchSiyuanSourceValidators(asset);
      if (!current) continue;
      if (!fallback) fallback = current;

      const etag = this._normalizeValidatorValue(current.etag || "");
      const lastModified = this._normalizeValidatorValue(current.lastModified || "");
      const size = this._parsePositiveInt(current.size || 0);
      const strong = !!(etag || lastModified);
      if (!strong) continue;

      if (!prevStrong) return current;

      const changed =
        (!!prevEtag && !!etag && prevEtag !== etag) ||
        (!!prevLastModified && !!lastModified && prevLastModified !== lastModified) ||
        (prevSize > 0 && size > 0 && prevSize !== size);
      if (changed) return current;
    }
    return fallback;
  }

  _canReuseBridgeCache(cacheMeta, currentValidators) {
    if (!cacheMeta || !cacheMeta.cached || cacheMeta.dirty) return false;
    const etagA = this._normalizeValidatorValue(cacheMeta.sourceEtag || "");
    const etagB = this._normalizeValidatorValue(currentValidators?.etag || "");
    const lmA = this._normalizeValidatorValue(cacheMeta.sourceLastModified || "");
    const lmB = this._normalizeValidatorValue(currentValidators?.lastModified || "");
    const sizeA = this._parsePositiveInt(cacheMeta.sourceSize || 0);
    const sizeB = this._parsePositiveInt(currentValidators?.size || 0);

    let strongCompared = 0;
    if (etagA && etagB) {
      strongCompared += 1;
      if (etagA !== etagB) return false;
    }
    if (lmA && lmB) {
      strongCompared += 1;
      if (lmA !== lmB) return false;
    }
    if (strongCompared === 0) return false;
    if (sizeA > 0 && sizeB > 0 && sizeA !== sizeB) {
      return false;
    }
    return true;
  }

  async _prepareAssetOnBridge(asset, updateLoading) {
    const key = String(asset);
    const existing = this._bridgePrepInFlight.get(key);
    if (existing) return existing;

    const task = (async () => {
      let bridgeHealth = null;
      if (typeof updateLoading === "function") updateLoading("loading.connecting");
      {
        const now = Date.now();
        const cached = this._bridgeHealthCache;
        const cacheHit = (
          cached &&
          cached.url === this.settings.bridgeUrl &&
          (now - cached.at) <= this._healthCacheTtlMs &&
          cached.data &&
          cached.data.status === "ok"
        );
        if (cacheHit) {
          bridgeHealth = cached.data;
        } else {
          const ctrl  = new AbortController();
          const timer = setTimeout(() => ctrl.abort(), 5000);
          try {
            const resp = await fetch(`${this.settings.bridgeUrl}/health`, { signal: ctrl.signal });
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            const healthText = await resp.text();
            let health = null;
            try { health = JSON.parse(healthText); } catch {}
            if (!health || health.status !== "ok") {
              throw new Error("Bridge /health returned an unexpected response. Verify Bridge URL points to the bridge service.");
            }
            bridgeHealth = health;
            this._bridgeHealthCache = { url: this.settings.bridgeUrl, at: Date.now(), data: health };
          } finally {
            clearTimeout(timer);
          }
        }
      }

      if (typeof updateLoading === "function") updateLoading("loading.uploading");
      // Before uploading local file, flush any pending saved version from Bridge.
      // This prevents stale local content from overwriting unsynced remote edits.
      let preSync = await this._syncSavedFile(asset, { silentNoChange: true });
      if (preSync?.inFlight) {
        for (let i = 0; i < 6 && preSync?.inFlight; i++) {
          await new Promise((resolve) => setTimeout(resolve, 400));
          preSync = await this._syncSavedFile(asset, { silentNoChange: true });
        }
        if (preSync?.inFlight) {
          throw new Error("Previous save sync is still running, please retry in a moment");
        }
      }

      const cacheMeta = await this._getBridgeCacheMeta(asset);
      const sourceValidatorsForCacheCheck = await this._fetchSiyuanSourceValidators(asset);
      if (this._canReuseBridgeCache(cacheMeta, sourceValidatorsForCacheCheck)) {
        console.info(`[Office Editor] Reusing bridge cache for "${asset}"`);
        return;
      }

      const siyuanBase = this._siyuanBaseUrl();
      if (!siyuanBase) throw new Error("Cannot determine SiYuan URL");
      const token = globalThis?.siyuan?.config?.api?.token;
      const headers = token ? { "Authorization": `Token ${token}` } : {};
      const sourceUrl = `${siyuanBase}/${encodeAssetPath(asset)}?t=${Date.now()}`;
      const sourceResp = await fetch(sourceUrl, { headers, cache: "no-store" });
      if (!sourceResp.ok) throw new Error(`SiYuan returned HTTP ${sourceResp.status}`);
      const sourceValidators = this._extractSourceValidatorsFromHeaders(sourceResp.headers);
      const fileBlob = await sourceResp.blob();
      await validateSourceBlob(asset, fileBlob);

      const uploadParams = new URLSearchParams({ asset });
      this._appendSourceValidatorParams(uploadParams, sourceValidators);
      const uploadUrl = `${this.settings.bridgeUrl}/upload?${uploadParams.toString()}${this._secretParam()}`;
      const contentType = extMime(getExt(asset)) || fileBlob.type || "application/octet-stream";
      const canChunkUpload = !!(
        bridgeHealth &&
        (
          bridgeHealth?.features?.chunkUpload === true ||
          Number.isFinite(bridgeHealth?.maxChunkMB)
        )
      );
      const useChunkByHint = (
        canChunkUpload &&
        this._chunkUploadHint?.url === this.settings.bridgeUrl &&
        Date.now() < (this._chunkUploadHint.expiresAt || 0) &&
        fileBlob.size >= this._chunkHintMinSizeBytes
      );
      let lastUploadPercent = -1;
      const updateUploadProgress = (loaded, total) => {
        if (typeof updateLoading !== "function") return;
        const safeTotal = Number(total) > 0 ? Number(total) : fileBlob.size;
        if (!Number.isFinite(safeTotal) || safeTotal <= 0) return;
        const percent = Math.max(0, Math.min(100, Math.floor((Number(loaded) / safeTotal) * 100)));
        if (percent === lastUploadPercent) return;
        lastUploadPercent = percent;
        updateLoading({ type: "uploadProgress", percent });
      };

      if (useChunkByHint) {
        const fallback = await this._uploadAssetInChunks(
          asset,
          fileBlob,
          contentType,
          updateUploadProgress,
          sourceValidators
        );
        console.info(`[Office Editor] Chunk upload (hint) succeeded for "${asset}" (${fallback.totalChunks} chunks @ ${fallback.chunkSize} bytes)`);
        return;
      }

      updateUploadProgress(0, fileBlob.size);
      const uploadResp = await this._uploadBlobWithProgress(
        uploadUrl,
        fileBlob,
        contentType,
        updateUploadProgress
      );
      updateUploadProgress(fileBlob.size, fileBlob.size);
      if (!uploadResp.ok) {
        const detail = parseHttpErrorDetail(uploadResp.text);
        if (uploadResp.status === 404) {
          throw new Error(this.t("message.upload404Hint", { bridgeUrl: this.settings.bridgeUrl }));
        }
        if (uploadResp.status === 413) {
          this._chunkUploadHint = {
            url: this.settings.bridgeUrl,
            expiresAt: Date.now() + this._chunkUploadHintTtlMs,
          };
          if (!canChunkUpload) {
            const detailSuffix = detail ? ` (${detail})` : "";
            throw new Error(
              `Bridge returned HTTP 413${detailSuffix}. Current bridge does not advertise chunk-upload support. Upgrade/redeploy bridge and check reverse-proxy upload limit (e.g. Nginx client_max_body_size).`
            );
          }
          try {
            const fallback = await this._uploadAssetInChunks(
              asset,
              fileBlob,
              contentType,
              updateUploadProgress,
              sourceValidators
            );
            console.info(`[Office Editor] Chunk upload fallback succeeded for "${asset}" (${fallback.totalChunks} chunks @ ${fallback.chunkSize} bytes)`);
            return;
          } catch (chunkErr) {
            const detailSuffix = detail ? ` (${detail})` : "";
            const chunkText = String(chunkErr?.message || chunkErr);
            throw new Error(
              `Bridge returned HTTP 413${detailSuffix}. Chunk upload fallback failed: ${chunkText}. Check reverse-proxy upload limit (e.g. Nginx client_max_body_size).`
            );
          }
        }
        const detailSuffix = detail ? `: ${detail}` : "";
        throw new Error(`Bridge returned HTTP ${uploadResp.status}${detailSuffix}`);
      }
    })().finally(() => {
      this._bridgePrepInFlight.delete(key);
    });

    this._bridgePrepInFlight.set(key, task);
    return task;
  }

  _openAssetSafe(assetPath, mode, title, target) {
    this.openAsset(assetPath, mode, title, target).catch((err) => {
      const message = this.t("message.openFailed", { error: err?.message || String(err) });
      console.error("[Office Editor] openAsset failed:", err);
      showMessage(message, 7000, "error");
    });
  }

  _embedAssetSafe(assetPath, mode, title, protyleOrEditor) {
    this.embedAsset(assetPath, mode, title, protyleOrEditor).catch((err) => {
      const message = this.t("message.embedFailed", { error: err?.message || String(err) });
      console.error("[Office Editor] embedAsset failed:", err);
      showMessage(message, 7000, "error");
    });
  }

  _getEmbedEditor(protyleOrEditor) {
    const fromProtyle = protyleOrEditor?.getInstance?.();
    if (fromProtyle?.insert) return fromProtyle;

    if (protyleOrEditor?.insert) return protyleOrEditor;

    const nested = protyleOrEditor?.protyle?.getInstance?.();
    if (nested?.insert) return nested;

    try {
      const active = typeof getActiveEditor === "function" ? getActiveEditor(true) || getActiveEditor() : null;
      if (active?.insert) return active;
      const activeNested = active?.protyle?.getInstance?.();
      if (activeNested?.insert) return activeNested;
    } catch {}
    return null;
  }

  // -----------------------------------------------------------------------
  // Open asset — core function
  // -----------------------------------------------------------------------
  async openAsset(assetPath, mode = "view", title, target = "dialog") {
    const asset = normAssetPath(assetPath);
    if (!asset || !isSupported(asset)) {
      showMessage(this.t("message.unsupportedFile"), 5000, "error");
      return;
    }

    const isMobile = this._isMobile();

    // PDF cannot be edited; mobile is preview-only
    if (isPdf(asset) && mode === "edit") {
      showMessage(this.t("message.pdfNoEdit"), 4000, "info");
      mode = "view";
    }
    if (isMobile && mode === "edit") {
      showMessage(this.t("message.mobilePreviewOnly"), 4000, "info");
      mode = "view";
    }

    // Mobile always uses dialog
    if (isMobile) target = "dialog";

    if (!this.settings.bridgeUrl) {
      showMessage(this.t("message.notConfigured"), 7000, "error");
      this.openSetting();
      return;
    }

    const pendingTimers = this._savedSyncTimers.get(asset);
    const hasPendingSave = this._dirtyAssets.has(asset) ||
      this._saveSignalAssets.has(asset) ||
      !!(pendingTimers && pendingTimers.length);
    if (this._postCloseSyncing.has(asset) && hasPendingSave) {
      // Avoid noisy duplicate notices while post-close sync is still settling.
      // We keep the guard behavior (do not re-open yet) but stay silent.
      return;
    }

    const displayName = title || fileName(asset);
    const frameSrc = this._buildFrameSrc(asset, mode, displayName);

    if (target === "tab") {
      try {
        await this._openTabWithFrame(asset, mode, displayName, frameSrc);
      } catch (err) {
        const msg = this.t("message.openFailed", { error: err?.message || String(err) });
        showMessage(msg, 7000, "error");
      }
      return;
    }

    // Create dialog first for visible loading.
    let loadingTextEl = null;
    let dialog = null;
    let cancelled = false;

    const updateLoading = (payload) => {
      if (!loadingTextEl) return;
      if (typeof payload === "string") {
        loadingTextEl.textContent = this.t(payload);
        return;
      }
      if (payload && payload.type === "uploadProgress") {
        const percent = Math.max(0, Math.min(100, Math.round(Number(payload.percent) || 0)));
        loadingTextEl.textContent = `${this.t("loading.uploading")} ${percent}%`;
      }
    };

    if (target !== "tab") {
      const dialogTitle = mode === "edit"
        ? this.t("dialog.titleEdit", { name: displayName })
        : this.t("dialog.titleView", { name: displayName });
      const dialogW = isMobile ? "100vw" : "92vw";
      const dialogH = isMobile ? "100vh" : "90vh";

      dialog = new Dialog({
        title:  dialogTitle,
        width:  dialogW,
        height: dialogH,
        content: loadingHtml(this.t("loading.connecting")),
        destroyCallback: () => {
          cancelled = true;
          this._dialogs.delete(dialog);
          if (dialog._cleanup) dialog._cleanup();
        },
      });
      this._dialogs.add(dialog);
      loadingTextEl = dialog.element.querySelector(".oo-bridge-loading__text");
    }

    // Async preamble: health check -> read -> upload
    try {
      await this._prepareAssetOnBridge(asset, updateLoading);
    } catch (err) {
      if (cancelled) return;
      const errMsg = String(err?.message || err);
      const bridgeErr = errMsg.startsWith("HTTP ") || errMsg.includes("Failed to fetch") || errMsg.includes("aborted");
      const msg = bridgeErr
        ? this.t("message.bridgeUnreachable", { url: this.settings.bridgeUrl, error: errMsg })
        : this.t("message.uploadFailed", { error: errMsg });
      if (dialog) dialog.destroy();
      showMessage(msg, 7000, "error");
      return;
    }
    if (cancelled) return;

    // Step 4: Open editor iframe in dialog
    updateLoading("loading.opening");
    if (dialog) {
      this._attachEditorToDialog(dialog, asset, mode, frameSrc);
    }
  }

  _hydrateEmbedsFromRoot(root) {
    if (!root?.querySelectorAll) return;
    const targets = new Set();
    for (const el of root.querySelectorAll(".oo-bridge-embed")) {
      targets.add(el);
    }
    for (const iframe of root.querySelectorAll("iframe.oo-bridge-embed__frame")) {
      targets.add(iframe.closest(".oo-bridge-embed") || iframe);
    }
    for (const iframe of root.querySelectorAll("iframe")) {
      if (iframe.classList?.contains("oo-bridge-dialog__frame")) continue;
      const meta = this._parseEmbedMetaFromIframe(iframe);
      if (!meta) continue;
      targets.add(iframe.closest(".oo-bridge-embed") || iframe);
    }
    for (const ph of root.querySelectorAll("protyle-html[data-content]")) {
      const contentMeta = this._parseEmbedMetaFromHtmlContent(ph.getAttribute("data-content") || "");
      if (!contentMeta) continue;
      targets.add(ph);
      const shadowIframe = ph.shadowRoot?.querySelector?.("iframe");
      if (shadowIframe) {
        targets.add(shadowIframe.closest(".oo-bridge-embed") || shadowIframe);
      }
    }
    for (const el of targets) {
      this._hydrateSingleEmbed(el).catch((err) => {
        console.warn("[Office Editor] embed hydrate failed:", err);
      });
    }
  }

  _hydrateEmbedsInEditor(editor) {
    const root = editor?.protyle?.element || editor?.element || null;
    if (!root) return;
    setTimeout(() => this._hydrateEmbedsFromRoot(root), 80);
    setTimeout(() => this._hydrateEmbedsFromRoot(root), 520);
  }

  _decodeEmbedHtmlContent(rawContent) {
    if (!rawContent || typeof rawContent !== "string") return "";
    if (/<iframe[\s>]/i.test(rawContent) || /class\s*=\s*["'][^"']*oo-bridge-embed/i.test(rawContent)) {
      return rawContent;
    }
    if (!/&lt;/.test(rawContent)) return rawContent;
    const textarea = document.createElement("textarea");
    textarea.innerHTML = rawContent;
    return textarea.value || rawContent;
  }

  _parseEmbedMetaFromFrameSrc(rawSrc) {
    if (!rawSrc || rawSrc === "about:blank") return null;
    let url;
    try {
      url = new URL(rawSrc, globalThis?.location?.href || "http://localhost");
    } catch {
      return null;
    }
    const normalizedPath = (url.pathname || "").replace(/\/+$/, "");
    if (!/\/editor$/i.test(normalizedPath)) return null;
    const asset = normAssetPath(url.searchParams.get("asset") || "");
    if (!asset || !isSupported(asset)) return null;
    const mode = url.searchParams.get("mode") === "edit" ? "edit" : "view";
    const displayName = url.searchParams.get("title") || fileName(asset);
    return { asset, mode, displayName };
  }

  _parseEmbedMetaFromIframe(iframe) {
    if (!iframe?.getAttribute) return null;
    const host = iframe.closest?.(".oo-bridge-embed");
    const srcMeta = this._parseEmbedMetaFromFrameSrc(iframe.getAttribute("src") || iframe.src || "");

    const dataAsset = normAssetPath(
      iframe.getAttribute("data-oo-asset") ||
      host?.getAttribute?.("data-oo-asset") ||
      ""
    );
    const asset = dataAsset || srcMeta?.asset || "";
    if (!asset || !isSupported(asset)) return null;

    const modeAttr = (
      iframe.getAttribute("data-oo-mode") ||
      host?.getAttribute?.("data-oo-mode") ||
      ""
    ).toLowerCase();
    const mode = modeAttr === "edit" ? "edit" : (srcMeta?.mode === "edit" ? "edit" : "view");
    const displayName = (
      iframe.getAttribute("data-oo-title") ||
      host?.getAttribute?.("data-oo-title") ||
      srcMeta?.displayName ||
      fileName(asset)
    );

    return { asset, mode, displayName };
  }

  _parseEmbedMetaFromHtmlContent(rawContent) {
    if (!rawContent || typeof rawContent !== "string") return null;
    const html = this._decodeEmbedHtmlContent(rawContent);
    const template = document.createElement("template");
    template.innerHTML = html;
    const host = template.content.querySelector(".oo-bridge-embed");
    const iframe = host?.querySelector("iframe") || template.content.querySelector("iframe");

    const hostAsset = normAssetPath(host?.getAttribute("data-oo-asset") || "");
    const iframeMeta = iframe ? this._parseEmbedMetaFromIframe(iframe) : null;
    const srcMeta = iframe
      ? this._parseEmbedMetaFromFrameSrc(iframe.getAttribute("src") || iframe.src || "")
      : null;

    const asset = hostAsset || iframeMeta?.asset || srcMeta?.asset || "";
    if (!asset || !isSupported(asset)) return null;

    const hostMode = (host?.getAttribute("data-oo-mode") || "").toLowerCase();
    const mode = hostMode === "edit"
      ? "edit"
      : (iframeMeta?.mode === "edit" || srcMeta?.mode === "edit" ? "edit" : "view");

    const displayName = (
      host?.getAttribute("data-oo-title") ||
      iframeMeta?.displayName ||
      srcMeta?.displayName ||
      fileName(asset)
    );

    return { asset, mode, displayName };
  }

  _refreshHtmlContentEditorSrc(rawContent, meta, nextFrameSrc) {
    if (!rawContent || typeof rawContent !== "string" || !meta?.asset) return "";
    const html = this._decodeEmbedHtmlContent(rawContent);
    const template = document.createElement("template");
    template.innerHTML = html;

    let host = template.content.querySelector(".oo-bridge-embed");
    let iframe = host?.querySelector("iframe") || template.content.querySelector("iframe");

    if (!host) {
      host = document.createElement("div");
      host.className = "oo-bridge-embed";
      if (iframe && iframe.parentNode) {
        iframe.parentNode.insertBefore(host, iframe);
        host.appendChild(iframe);
      } else {
        template.content.appendChild(host);
      }
    }

    if (!iframe) {
      iframe = document.createElement("iframe");
      host.appendChild(iframe);
    }

    if (!iframe.classList.contains("oo-bridge-embed__frame")) {
      iframe.classList.add("oo-bridge-embed__frame");
    }
    this._setIframeSrcIfNeeded(iframe, nextFrameSrc);
    iframe.setAttribute("allow", "clipboard-read; clipboard-write; fullscreen");
    iframe.setAttribute("data-oo-asset", meta.asset);
    iframe.setAttribute("data-oo-mode", meta.mode);
    iframe.setAttribute("data-oo-title", meta.displayName || fileName(meta.asset));

    host.setAttribute("data-oo-asset", meta.asset);
    host.setAttribute("data-oo-mode", meta.mode);
    host.setAttribute("data-oo-title", meta.displayName || fileName(meta.asset));

    return template.innerHTML;
  }

  _setIframeSrcIfNeeded(iframe, nextFrameSrc) {
    if (!iframe || !nextFrameSrc) return;
    const currentSrc = String(iframe.getAttribute("src") || "").trim();
    const nextSrc = String(nextFrameSrc || "").trim();
    if (!nextSrc || currentSrc === nextSrc) return;
    iframe.setAttribute("src", nextSrc);
  }

  async _hydrateSingleEmbed(el) {
    if (!el || el.dataset.ooHydrating === "1") return;
    if (el.tagName === "PROTYLE-HTML") {
      const lastHydratedAt = Number(el.dataset.ooHydratedAt || 0);
      if (lastHydratedAt > 0 && (Date.now() - lastHydratedAt) < 15000) return;

      const rawContent = el.getAttribute("data-content") || "";
      const meta = this._parseEmbedMetaFromHtmlContent(rawContent);
      if (!meta || !this.settings.bridgeUrl) return;

      let mode = meta.mode;
      if (isPdf(meta.asset) && mode === "edit") mode = "view";
      if (this._isMobile() && mode === "edit") mode = "view";
      const displayName = meta.displayName || fileName(meta.asset);

      el.dataset.ooHydrating = "1";
      try {
        await this._prepareAssetOnBridge(meta.asset);
        const frameSrc = this._buildFrameSrc(meta.asset, mode, displayName);

        const nextContent = this._refreshHtmlContentEditorSrc(rawContent, {
          asset: meta.asset,
          mode,
          displayName,
        }, frameSrc);
        if (nextContent && nextContent !== rawContent) {
          el.setAttribute("data-content", nextContent);
        }

        const renderedIframes = [];
        if (el.shadowRoot?.querySelectorAll) {
          renderedIframes.push(...el.shadowRoot.querySelectorAll("iframe"));
        }
        if (typeof el.querySelectorAll === "function") {
          renderedIframes.push(...el.querySelectorAll("iframe"));
        }
        const uniqueIframes = Array.from(new Set(renderedIframes));
        for (const iframe of uniqueIframes) {
          const currentMeta = this._parseEmbedMetaFromIframe(iframe);
          const isTarget = (
            (currentMeta && currentMeta.asset === meta.asset) ||
            iframe.classList?.contains("oo-bridge-embed__frame") ||
            (!currentMeta && uniqueIframes.length === 1)
          );
          if (!isTarget) continue;
          this._setIframeSrcIfNeeded(iframe, frameSrc);
          iframe.setAttribute("allow", "clipboard-read; clipboard-write; fullscreen");
          iframe.setAttribute("data-oo-asset", meta.asset);
          iframe.setAttribute("data-oo-mode", mode);
          iframe.setAttribute("data-oo-title", displayName);
          const host = iframe.closest?.(".oo-bridge-embed");
          if (host) {
            host.setAttribute("data-oo-asset", meta.asset);
            host.setAttribute("data-oo-mode", mode);
            host.setAttribute("data-oo-title", displayName);
          }
        }
        el.dataset.ooHydratedAt = String(Date.now());
      } finally {
        delete el.dataset.ooHydrating;
      }
      return;
    }

    const iframe = el.tagName === "IFRAME"
      ? el
      : el.matches?.("iframe.oo-bridge-embed__frame")
      ? el
      : el.querySelector?.(".oo-bridge-embed__frame");
    if (!iframe || iframe.classList?.contains("oo-bridge-dialog__frame") || !this.settings.bridgeUrl) return;
    if (iframe.getAttribute("loading") === "lazy") {
      iframe.setAttribute("loading", "eager");
    }

    const host = el.matches?.(".oo-bridge-embed")
      ? el
      : iframe.closest(".oo-bridge-embed");
    const lastHydratedAt = Number(
      iframe.dataset?.ooHydratedAt ||
      host?.dataset?.ooHydratedAt ||
      0
    );
    if (lastHydratedAt > 0 && (Date.now() - lastHydratedAt) < 15000) return;

    let asset = normAssetPath(host?.dataset?.ooAsset || "");
    let mode = host?.dataset?.ooMode === "edit" ? "edit" : "view";
    let displayName = host?.dataset?.ooTitle || "";

    if (!asset || !isSupported(asset)) {
      const meta = this._parseEmbedMetaFromIframe(iframe);
      if (meta) {
        asset = meta.asset;
        mode = meta.mode;
        if (!displayName) displayName = meta.displayName;
      }
    }
    if (!asset || !isSupported(asset)) return;

    if (isPdf(asset) && mode === "edit") mode = "view";
    if (this._isMobile() && mode === "edit") mode = "view";
    if (!displayName) displayName = fileName(asset);

    el.dataset.ooHydrating = "1";
    try {
      await this._prepareAssetOnBridge(asset);
      const frameSrc = this._buildFrameSrc(asset, mode, displayName);
      this._setIframeSrcIfNeeded(iframe, frameSrc);
      iframe.setAttribute("data-oo-asset", asset);
      iframe.setAttribute("data-oo-mode", mode);
      iframe.setAttribute("data-oo-title", displayName);
      iframe.setAttribute("allow", "clipboard-read; clipboard-write; fullscreen");
      const hydratedAt = String(Date.now());
      iframe.dataset.ooHydratedAt = hydratedAt;
      if (host?.dataset) {
        host.dataset.ooAsset = asset;
        host.dataset.ooMode = mode;
        host.dataset.ooTitle = displayName;
        host.dataset.ooHydratedAt = hydratedAt;
      }
    } finally {
      delete el.dataset.ooHydrating;
    }
  }

  _onProtyleLoadedStatic({ detail }) {
    const root = detail?.protyle?.element;
    this._hydrateEmbedsFromRoot(root);
    setTimeout(() => this._hydrateEmbedsFromRoot(root), 520);
  }

  _onProtyleLoadedDynamic({ detail }) {
    const root = detail?.protyle?.element;
    this._hydrateEmbedsFromRoot(root);
    setTimeout(() => this._hydrateEmbedsFromRoot(root), 520);
  }

  _onSwitchProtyle({ detail }) {
    const root = detail?.protyle?.element;
    if (!root) return;
    setTimeout(() => this._hydrateEmbedsFromRoot(root), 120);
    setTimeout(() => this._hydrateEmbedsFromRoot(root), 640);
  }

  _onEditorContentClick({ detail }) {
    const root = detail?.protyle?.element;
    this._hydrateEmbedsFromRoot(root);
    setTimeout(() => this._hydrateEmbedsFromRoot(root), 360);
  }

  _clearSavedSyncTimers(asset) {
    const key = String(asset || "");
    if (!key) return;
    const timers = this._savedSyncTimers.get(key);
    if (!timers?.length) return;
    for (const timer of timers) clearTimeout(timer);
    this._savedSyncTimers.delete(key);
  }

  _notifySyncIssue(asset, errorText = "", pendingOnly = false) {
    const key = String(asset || "");
    if (!key) return;
    const now = Date.now();
    const prev = this._syncAlertAt.get(key) || 0;
    if (now - prev < 5000) return;
    this._syncAlertAt.set(key, now);

    const name = fileName(key);
    const msg = pendingOnly
      ? this.t("message.syncPullPending", { name })
      : this.t("message.syncWritebackFailed", {
          name,
          error: errorText || "unknown error",
        });
    if (pendingOnly) {
      // Pending-only means callback lag / eventual consistency; don't toast.
      return;
    }
    showMessage(msg, 10000, "error");
  }

  _isAssetCloseBlocked(asset) {
    const key = String(asset || "");
    if (!key) return false;
    if (this._dirtyAssets.has(key)) return true;
    if (this._saveSignalAssets.has(key)) return true;
    if (this._postCloseSyncing.has(key)) return true;
    const timers = this._savedSyncTimers.get(key);
    return !!(timers && timers.length);
  }

  _tabCloseKey(asset, mode = "view", frameSrc = "") {
    return `${String(asset || "")}|${String(mode || "view")}|${String(frameSrc || "")}`;
  }

  _markAssetSettled(asset) {
    const key = String(asset || "");
    if (!key) return;
    this._clearSavedSyncTimers(key);
    this._saveSignalAssets.delete(key);
    this._dirtyAssets.delete(key);
    this._syncAlertAt.delete(key);
  }

  async _discardAssetChanges(asset) {
    const key = String(asset || "");
    if (!key) return;
    this._markAssetSettled(key);
    await this._cleanupBridgeFile(key, null, { purge: true });
  }

  _postMessageToAssetFrames(asset, payload) {
    const key = String(asset || "");
    if (!key || !payload) return 0;
    const frames = document.querySelectorAll("iframe.oo-bridge-dialog__frame, iframe.oo-bridge-embed__frame");
    let sent = 0;
    for (const frame of frames) {
      let frameAsset = normAssetPath(
        frame.getAttribute("data-oo-asset") ||
        frame.dataset?.ooAsset ||
        ""
      );
      if (!frameAsset) {
        const src = frame.getAttribute("src") || "";
        if (src) {
          try {
            const parsed = new URL(src, window.location.href);
            frameAsset = normAssetPath(parsed.searchParams.get("asset") || "");
          } catch {}
        }
      }
      if (frameAsset !== key) continue;
      try {
        frame.contentWindow?.postMessage(payload, "*");
        sent++;
      } catch {}
    }
    return sent;
  }

  _resolveNativeClose(asset, approved) {
    const key = String(asset || "");
    if (!key) return;
    const waiter = this._nativeCloseWaiters.get(key);
    if (!waiter) return;
    if (waiter.timer) clearTimeout(waiter.timer);
    this._nativeCloseWaiters.delete(key);
    try { waiter.resolve(!!approved); } catch {}
  }

  async _requestEditorNativeClose(asset, timeoutMs = 180000) {
    const key = String(asset || "");
    if (!key) return false;
    let posted = this._postMessageToAssetFrames(key, { type: "oo-bridge-request-close", asset: key });
    for (let i = 0; !posted && i < 8; i++) {
      await new Promise((resolve) => setTimeout(resolve, 250));
      posted = this._postMessageToAssetFrames(key, { type: "oo-bridge-request-close", asset: key });
    }
    if (!posted) return false;

    const prev = this._nativeCloseWaiters.get(key);
    if (prev?.timer) clearTimeout(prev.timer);
    if (prev?.resolve) {
      try { prev.resolve(false); } catch {}
    }

    return await new Promise((resolve) => {
      const timer = setTimeout(() => {
        this._resolveNativeClose(key, false);
      }, timeoutMs);
      this._nativeCloseWaiters.set(key, { resolve, timer });
    });
  }

  _findOfficeTab(asset, mode = "view", frameSrc = "") {
    if (typeof getAllTabs !== "function") return null;
    const tabs = getAllTabs() || [];
    for (const tab of tabs) {
      const data = tab?.model?.data;
      if (!data || typeof data !== "object") continue;
      if (String(data.asset || "") !== String(asset || "")) continue;
      if (String(data.mode || "view") !== String(mode || "view")) continue;
      const currentSrc = String(data.frameSrc || this._resolveTabFrameSrc(data) || "");
      if (frameSrc && currentSrc !== String(frameSrc)) continue;
      return tab;
    }
    return null;
  }

  _closeOfficeTab(asset, mode, frameSrc, action = "save") {
    const key = this._tabCloseKey(asset, mode, frameSrc);
    const tab = this._findOfficeTab(asset, mode, frameSrc);
    if (!tab) return false;
    this._tabCloseActionOnce.set(key, action);
    try {
      tab.close();
      return true;
    } catch {
      this._tabCloseActionOnce.delete(key);
      return false;
    }
  }

  _scheduleSavedSync(asset) {
    const key = String(asset || "");
    if (!key) return;

    const timers = [];
    this._clearSavedSyncTimers(key);
    this._saveSignalAssets.add(key);
    this._savedSyncTimers.set(key, timers);
    let done = false;
    const runAttempt = async (isLast) => {
      if (done) return;
      try {
        const result = await this._syncSavedFile(key, { silentNoChange: true });
        if (result?.changed) {
          done = true;
          this._clearSavedSyncTimers(key);
          this._syncAlertAt.delete(key);
          return;
        }
        if (isLast) {
          done = true;
          this._notifySyncIssue(key, "", true);
          this._clearSavedSyncTimers(key);
        }
      } catch (err) {
        if (!isLast) return;
        done = true;
        const errText = String(err?.message || err);
        this._notifySyncIssue(key, errText, false);
        this._clearSavedSyncTimers(key);
      }
    };

    runAttempt(false);
    const delays = [1500, 4500, 9000];
    for (let i = 0; i < delays.length; i++) {
      const delay = delays[i];
      const timer = setTimeout(() => {
        runAttempt(i === delays.length - 1);
      }, delay);
      timers.push(timer);
    }
  }

  _onBridgeSavedMessage(event) {
    const data = event?.data;
    if (!data || typeof data.type !== "string") return;
    const asset = normAssetPath(String(data.asset || ""));
    if (!asset || !isSupported(asset)) return;
    if (data.type === "oo-bridge-request-close-ok") {
      this._resolveNativeClose(asset, true);
      return;
    }
    if (data.type === "oo-bridge-dirty") {
      this._dirtyAssets.add(asset);
      return;
    }
    if (data.type === "oo-bridge-clean") {
      this._dirtyAssets.delete(asset);
      return;
    }
    if (data.type !== "oo-bridge-saved") return;
    this._dirtyAssets.delete(asset);
    this._scheduleSavedSync(asset);
  }

  async embedAsset(assetPath, mode = "view", title, protyleOrEditor) {
    const asset = normAssetPath(assetPath);
    if (!asset || !isSupported(asset)) {
      showMessage(this.t("message.unsupportedFile"), 5000, "error");
      return;
    }
    if (!this.settings.bridgeUrl) {
      showMessage(this.t("message.notConfigured"), 7000, "error");
      this.openSetting();
      return;
    }
    const editor = this._getEmbedEditor(protyleOrEditor);
    if (!editor) {
      showMessage(this.t("message.embedNoEditor"), 5000, "error");
      return;
    }
    if (isPdf(asset) && mode === "edit") {
      showMessage(this.t("message.pdfNoEdit"), 3000, "info");
      mode = "view";
    }
    const displayName = title || fileName(asset);
    await this._prepareAssetOnBridge(asset);
    const frameSrc = this._buildFrameSrc(asset, mode, displayName);
    const html = `<iframe src="${escapeHtml(frameSrc)}"></iframe>`;
    editor.insert(html, true, true);
    showMessage(this.t("message.embedInserted"), 2500, "info");
  }

  // -----------------------------------------------------------------------
  // Attach editor iframe to Dialog
  // -----------------------------------------------------------------------
  _attachEditorToDialog(dialog, asset, mode, frameSrc) {
    // Replace loading content with iframe
    const container = dialog.element.querySelector(".b3-dialog__body");
    if (container) container.innerHTML = iframeHtml(frameSrc);

    // Set up sync
    const syncSaved = () => this._syncSavedFile(asset, { silentNoChange: true });
    let liveSyncTimer = null;

    if (mode === "edit") {
      liveSyncTimer = setInterval(() => {
        syncSaved().catch((err) => {
          this._notifySyncIssue(asset, String(err?.message || err), false);
        });
      }, 30000);
    }

    const originalDestroy = dialog.destroy.bind(dialog);
    let closeFlow = null;

    const requestClose = async () => {
      if (closeFlow) return closeFlow;
      closeFlow = (async () => {
        if (mode !== "edit" || this._isUnloading || !this._isAssetCloseBlocked(asset)) {
          originalDestroy();
          return;
        }
        if (!this._dirtyAssets.has(asset)) {
          originalDestroy();
          return;
        }
        const approved = await this._requestEditorNativeClose(asset);
        if (!approved) return;
        originalDestroy();
      })().finally(() => {
        closeFlow = null;
      });
      return closeFlow;
    };

    dialog.destroy = () => {
      requestClose().catch((err) => {
        showMessage(this.t("message.requestSaveFailed", {
          name: fileName(asset),
          error: String(err?.message || err),
        }), 9000, "error");
      });
    };

    dialog._cleanup = () => {
      dialog.destroy = originalDestroy;
      if (liveSyncTimer) clearInterval(liveSyncTimer);
      if (mode === "edit") {
        this._postCloseSyncAndCleanup(asset);
      }
    };
  }

  // -----------------------------------------------------------------------
  // Attach editor iframe to Tab
  // -----------------------------------------------------------------------
  _attachEditorToTab(tab, asset, mode, frameSrc) {
    tab.element.style.height = "100%";
    const runtime = this._ensureTabRuntime(tab);
    if (runtime.msgHandler) window.removeEventListener("message", runtime.msgHandler);
    if (runtime.liveSyncTimer) clearInterval(runtime.liveSyncTimer);
    tab.element.innerHTML = iframeHtml(frameSrc);

    const syncSaved = () => this._syncSavedFile(asset, { silentNoChange: true });

    let liveSyncTimer = null;
    if (mode === "edit") {
      liveSyncTimer = setInterval(() => {
        syncSaved().catch((err) => {
          this._notifySyncIssue(asset, String(err?.message || err), false);
        });
      }, 30000);
    }

    runtime.msgHandler = null;
    runtime.liveSyncTimer = liveSyncTimer;
  }

  // -----------------------------------------------------------------------
  // Sync saved file from Bridge back to SiYuan
  // -----------------------------------------------------------------------
  async _syncSavedFile(asset, opts = {}) {
    const silentNoChange = !!opts?.silentNoChange;
    if (this._syncInFlight.get(asset)) return { changed: false, inFlight: true };
    this._syncInFlight.set(asset, true);

    try {
      const sp = this._bridgeAuthPrefix();
      const url = `${this.settings.bridgeUrl}/saved?${sp}asset=${encodeURIComponent(asset)}&t=${Date.now()}`;
      let resp;
      try {
        resp = await fetch(url, { cache: "no-store" });
      } catch (err) {
        throw new Error(`Bridge /saved request failed: ${String(err?.message || err)}`);
      }

      if (resp.status === 204 || resp.status === 404) {
        if (!silentNoChange) {
          console.info(`[Office Editor] No pending saved data for "${asset}"`);
        }
        return { changed: false };
      }
      if (!resp.ok) {
        throw new Error(`Bridge /saved returned HTTP ${resp.status}`);
      }

      const savedBlob = await resp.blob();
      const siyuanBase = this._siyuanBaseUrl();
      if (!siyuanBase) {
        throw new Error("Cannot determine SiYuan URL");
      }

      const token = globalThis?.siyuan?.config?.api?.token;
      const headers = {};
      if (token) headers["Authorization"] = `Token ${token}`;

      let syncOk = false;
      let lastErr = null;
      const candidates = [`/data/${asset}`, `/${asset}`];

      for (const targetPath of candidates) {
        const form = new FormData();
        form.append("path",    targetPath);
        form.append("isDir",   "false");
        form.append("modTime", String(Date.now()));
        form.append("file",    savedBlob, fileName(asset));

        const putResp = await fetch(`${siyuanBase}/api/file/putFile`, {
          method: "POST", body: form, headers, cache: "no-store",
        });
        let putJson = null;
        try { putJson = await putResp.json(); } catch {}

        if (putResp.ok && putJson && typeof putJson === "object" && putJson.code === 0) {
          syncOk = true;
          break;
        }
        lastErr = new Error(`path=${targetPath}, http=${putResp.status}, body=${JSON.stringify(putJson)}`);
      }

      if (!syncOk) {
        throw lastErr || new Error("SiYuan putFile failed");
      }

      this._saveSignalAssets.delete(asset);
      this._dirtyAssets.delete(asset);
      this._syncAlertAt.delete(asset);
      const sourceValidators = await this._fetchSiyuanSourceValidatorsAfterWrite(asset);
      await this._cleanupBridgeFile(asset, sourceValidators);
      console.info(`[Office Editor] Synced "${asset}" back to SiYuan`);
      return { changed: true, sourceValidators };
    } finally {
      this._syncInFlight.delete(asset);
    }
  }

  // -----------------------------------------------------------------------
  // Post-close sync (3 targeted syncs instead of 90s polling)
  // -----------------------------------------------------------------------
  async _postCloseSyncAndCleanup(asset) {
    this._postCloseSyncing.add(asset);
    let synced = false;
    let lastErr = null;
    try {
      // If we did not observe an explicit save signal, keep close-sync short.
      // If we did observe a save signal, allow longer retries for callback lag.
      const hadSaveSignalAtStart = this._saveSignalAssets.has(asset);
      const retryDelays = hadSaveSignalAtStart ? [0, 2500, 7000] : [0, 800];

      for (const delay of retryDelays) {
        if (delay > 0) {
          await new Promise((resolve) => setTimeout(resolve, delay));
        }
        try {
          const result = await this._syncSavedFile(asset, { silentNoChange: true });
          if (result?.changed) {
            synced = true;
            break;
          }
        } catch (err) {
          lastErr = err;
        }
      }

      const hadSaveSignal = this._saveSignalAssets.has(asset);
      if (synced || (!hadSaveSignal && !lastErr)) {
        await this._cleanupBridgeFile(asset);
        this._saveSignalAssets.delete(asset);
        this._dirtyAssets.delete(asset);
        this._syncAlertAt.delete(asset);
        return;
      }

      const errText = lastErr ? String(lastErr?.message || lastErr) : "";
      this._notifySyncIssue(asset, errText, !lastErr);
    } finally {
      this._postCloseSyncing.delete(asset);
    }
  }

  // -----------------------------------------------------------------------
  // Clear Bridge session state after sync (cache retained until TTL)
  // -----------------------------------------------------------------------
  async _cleanupBridgeFile(asset, sourceValidators = null, options = {}) {
    try {
      const sp = this._bridgeAuthPrefix();
      const base = sp ? sp.slice(0, -1) : "";
      const params = new URLSearchParams(base);
      params.set("asset", String(asset || ""));
      if (options?.purge) params.set("purge", "1");
      this._appendSourceValidatorParams(params, sourceValidators);
      await fetch(
        `${this.settings.bridgeUrl}/cleanup?${params.toString()}`,
        { method: "POST" }
      );
    } catch {
      // Silently ignore
    }
  }

  // -----------------------------------------------------------------------
  // Prompt open by manual path (Dialog-based, works in Electron)
  // -----------------------------------------------------------------------
  _promptOpen() {
    const dialog = new Dialog({
      title: this.t("command.openByPath"),
      width: "520px",
      content: `<div class="b3-dialog__content" style="padding: 16px;">
  <input class="b3-text-field" style="width: 100%;"
    placeholder="assets/example.docx" value="assets/" id="oo-prompt-input">
</div>
<div class="b3-dialog__action">
  <button class="b3-button b3-button--cancel" id="oo-prompt-cancel">${escapeHtml(this.t("prompt.cancel"))}</button>
  <div class="fn__space"></div>
  <button class="b3-button b3-button--text" id="oo-prompt-ok">${escapeHtml(this.t("prompt.ok"))}</button>
</div>`,
    });

    const inputEl   = dialog.element.querySelector("#oo-prompt-input");
    const okBtn     = dialog.element.querySelector("#oo-prompt-ok");
    const cancelBtn = dialog.element.querySelector("#oo-prompt-cancel");

    const doOpen = () => {
      const raw = inputEl.value;
      dialog.destroy();
      const p = normAssetPath(raw);
      if (!p)              { showMessage(this.t("message.invalidAssetPath"), 5000, "error"); return; }
      if (!isSupported(p)) { showMessage(this.t("message.unsupportedFile"), 5000, "error"); return; }
      this.openAsset(p, this.settings.defaultMode);
    };

    okBtn.addEventListener("click", doOpen);
    cancelBtn.addEventListener("click", () => dialog.destroy());
    inputEl.addEventListener("keydown", (e) => { if (e.key === "Enter") doOpen(); });
    setTimeout(() => { inputEl.focus(); inputEl.setSelectionRange(inputEl.value.length, inputEl.value.length); }, 100);
  }

  // -----------------------------------------------------------------------
  // Context menu integration
  // -----------------------------------------------------------------------
  _addMenuItems(menu, assetPath, label, protyle = null) {
    if (!menu || !assetPath || !isSupported(assetPath)) return;
    if (!this.settings.bridgeUrl) return;

    const isMobile  = this._isMobile();
    const canEdit   = this.settings.enableEdit && !isPdf(assetPath) && !isMobile;
    const embedEditor = this._getEmbedEditor(protyle);
    const canEmbed = !!embedEditor && !isMobile;

    // Dialog items
    menu.addItem({
      icon: ICON_PREVIEW, label: this.t("menu.preview"),
      click: () => this._openAssetSafe(assetPath, "view", label, "dialog"),
    });
    if (canEdit) {
      menu.addItem({
        icon: ICON_EDIT,   label: this.t("menu.edit"),
        click: () => this._openAssetSafe(assetPath, "edit", label, "dialog"),
      });
    }

    if (canEmbed) {
      menu.addItem({
        icon: ICON_EMBED, label: this.t("menu.embedPreview"),
        click: () => this._embedAssetSafe(assetPath, "view", label, embedEditor),
      });
      if (canEdit) {
        menu.addItem({
          icon: ICON_EMBED, label: this.t("menu.embedEdit"),
          click: () => this._embedAssetSafe(assetPath, "edit", label, embedEditor),
        });
      }
    }

    // Tab items (desktop only)
    if (!isMobile) {
      menu.addItem({
        icon: ICON_TAB, label: this.t("menu.previewInTab"),
        click: () => this._openAssetSafe(assetPath, "view", label, "tab"),
      });
      if (canEdit) {
        menu.addItem({
          icon: ICON_TAB, label: this.t("menu.editInTab"),
          click: () => this._openAssetSafe(assetPath, "edit", label, "tab"),
        });
      }
    }
  }

  _onLinkMenu({ detail }) {
    try {
      const { menu, element } = detail || {};
      if (!menu || !element) return;
      const asset = assetFromElement(element);
      if (!isSupported(asset)) return;
      this._addMenuItems(menu, asset, element.innerText?.trim() || fileName(asset), detail.protyle);
    } catch (e) { console.error(e); }
  }

  _onContentMenu({ detail }) {
    try {
      const { menu, element } = detail || {};
      if (!menu || !element) return;
      const asset = assetFromElement(element);
      if (!isSupported(asset)) return;
      this._addMenuItems(menu, asset, element.innerText?.trim() || fileName(asset), detail.protyle);
    } catch (e) { console.error(e); }
  }

  _onFileAnnotationMenu({ detail }) {
    try {
      const { menu, element } = detail || {};
      if (!menu || !element) return;
      const asset = assetFromElement(element);
      if (!isSupported(asset)) return;
      this._addMenuItems(menu, asset, element.innerText?.trim() || fileName(asset), detail.protyle);
    } catch (e) { console.error(e); }
  }

  _onDocTreeMenu({ detail }) {
    try {
      const { menu, elements } = detail || {};
      if (!menu || !elements || !elements.length) return;
      for (const el of elements) {
        const asset = assetFromElement(el);
        if (!isSupported(asset)) continue;
        this._addMenuItems(menu, asset, el.innerText?.trim() || fileName(asset));
        return;
      }
    } catch (e) { console.error(e); }
  }

  // -----------------------------------------------------------------------
  // Settings panel
  // -----------------------------------------------------------------------
  _initSettings() {
    const mkInput = (placeholder, type = "text") => {
      const el = document.createElement("input");
      el.className   = "b3-text-field";
      el.type        = type;
      el.placeholder = placeholder;
      el.style.width = "280px";
      return el;
    };

    const els = {
      bridgeUrl:     mkInput("http://your-public-server:27689"),
      onlyofficeUrl: mkInput("http://your-public-server:27670"),
      bridgeSecret:  mkInput("", "password"),
    };
    this._settingEls = els;

    this.setting = new Setting({
      width: "760px",
      destroyCallback: () => {
        this._onSave({ silentNoChange: true }).catch((err) => {
          console.warn("[Office Editor] Failed to save settings on close:", err);
        });
      },
    });

    this.setting.addItem({ title: this.t("settings.bridgeUrl"),     description: this.t("settings.bridgeUrlDesc"),     createActionElement: () => els.bridgeUrl });
    this.setting.addItem({ title: this.t("settings.onlyofficeUrl"), description: this.t("settings.onlyofficeUrlDesc"), createActionElement: () => els.onlyofficeUrl });
    this.setting.addItem({ title: this.t("settings.bridgeSecret"),  description: this.t("settings.bridgeSecretDesc"),  createActionElement: () => els.bridgeSecret });

    this._syncInputs();
  }

  _syncInputs() {
    const e = this._settingEls;
    if (!e.bridgeUrl) return;
    e.bridgeUrl.value     = this.settings.bridgeUrl;
    e.onlyofficeUrl.value = this.settings.onlyofficeUrl;
    e.bridgeSecret.value  = this.settings.bridgeSecret;
  }

  async _onSave(opts = {}) {
    const silentNoChange = !!opts.silentNoChange;
    const e = this._settingEls;
    if (!e.bridgeUrl) return;
    const nextSettings = {
      bridgeUrl:     normUrl(e.bridgeUrl.value, ""),
      onlyofficeUrl: normUrl(e.onlyofficeUrl.value, ""),
      bridgeSecret:  e.bridgeSecret.value.trim(),
      // Keep fixed defaults hidden from settings UI.
      defaultMode:   DEFAULT_SETTINGS.defaultMode,
      enableEdit:    DEFAULT_SETTINGS.enableEdit,
    };
    const changed =
      nextSettings.bridgeUrl !== this.settings.bridgeUrl ||
      nextSettings.onlyofficeUrl !== this.settings.onlyofficeUrl ||
      nextSettings.bridgeSecret !== this.settings.bridgeSecret ||
      nextSettings.defaultMode !== this.settings.defaultMode ||
      nextSettings.enableEdit !== this.settings.enableEdit;
    this.settings = nextSettings;
    if (!changed) {
      this._syncInputs();
      return;
    }
    await this._saveSettings();
    this._syncInputs();
    if (!silentNoChange) showMessage(this.t("message.settingsSaved"), 2500, "info");
  }
}

module.exports = OnlyOfficeBridgePlugin;
module.exports.default = OnlyOfficeBridgePlugin;
