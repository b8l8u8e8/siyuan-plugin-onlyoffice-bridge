/**
 * SiYuan ONLYOFFICE Bridge Plugin
 *
 * Supports the scenario where ONLYOFFICE is on a public server and SiYuan
 * is on an internal network with no public IP.
 *
 * "Push" model:
 *   1. Plugin reads the document from SiYuan (browser → SiYuan, internal)
 *   2. Plugin uploads it to Bridge (browser → Bridge, public)
 *   3. Bridge serves it to ONLYOFFICE (Bridge → ONLYOFFICE, public/same host)
 *   4. On save: ONLYOFFICE → Bridge callback → Bridge stores in memory
 *   5. Plugin pulls saved file from Bridge → writes back to SiYuan
 *
 * Requires: Bridge service configured. Bridge and ONLYOFFICE must be on a
 * public server reachable by both the browser and ONLYOFFICE itself.
 */

const { Plugin, Dialog, Setting, showMessage } = require("siyuan");

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const STORAGE_KEY = "settings.json";

const ICON_PREVIEW = "iconOOPreview";
const ICON_EDIT    = "iconOOEdit";
const SVG_ICONS = `<symbol id="${ICON_PREVIEW}" viewBox="0 0 24 24">
  <path fill="currentColor" d="M3 4a2 2 0 0 1 2-2h10l6 6v12a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4Zm11 0v5h5l-5-5Zm-4 9h2v6h-2v-6Zm-4 2h2v4H6v-4Zm8-3h2v7h-2v-7Z"/>
</symbol>
<symbol id="${ICON_EDIT}" viewBox="0 0 24 24">
  <path fill="currentColor" d="M3 4a2 2 0 0 1 2-2h10l6 6v12a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4Zm11 0v5h5l-5-5ZM14.06 12.94l-5 5L8 21l3.06-1.06 5-5-2-2Zm3.41-1.41 2-2-2-2-2 2 2 2Z"/>
</symbol>`;

const SUPPORTED_EXTENSIONS = new Set([
  "doc","docx","docm","dotx","dotm","odt","rtf","txt","md",
  "csv","xls","xlsx","xlsm","xltx","xltm","ods",
  "ppt","pptx","pptm","potx","potm","odp","pdf",
]);

const CELL_EXTS  = new Set(["xls","xlsx","xlsm","xltx","xltm","ods","csv"]);
const SLIDE_EXTS = new Set(["ppt","pptx","pptm","potx","potm","odp"]);

const DEFAULT_SETTINGS = {
  bridgeUrl:                  "",
  onlyofficeUrl:              "",
  bridgeSecret:               "",
  defaultMode:                "view",
  enableEdit:                 true,
  showInfoBar:                false,
  dialogWidth:                "90vw",
  dialogHeight:               "85vh",
  enableLinkMenu:             true,
  enableFileAnnotationMenu:   true,
  enableDocTreeMenu:          true,
};

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

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function normUrl(v, fb) {
  let s = String(v || "").trim() || fb || "";
  if (!s) return "";
  if (!/^https?:\/\//i.test(s)) s = `http://${s}`;
  return s.replace(/\/+$/, "");
}

function normSize(v, fb) {
  const s = String(v ?? "").trim();
  return s && /^\d+(px|vw|vh|%)$/i.test(s) ? s : fb;
}

function normBool(v, fb) {
  if (typeof v === "boolean") return v;
  if (v === "true")  return true;
  if (v === "false") return false;
  return !!fb;
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

function toArray(v) {
  if (!v) return [];
  if (Array.isArray(v)) return v;
  if (typeof v.length === "number") return Array.from(v);
  return [v];
}

function tryAttr(el, attr) {
  if (!el?.getAttribute) return "";
  const v = el.getAttribute(attr);
  return typeof v === "string" ? v : "";
}

const DOC_TREE_ATTRS = ["data-path","data-url","data-href","data-src","data-id","href"];

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
    this._onLinkMenu          = this._onLinkMenu.bind(this);
    this._onFileAnnotationMenu = this._onFileAnnotationMenu.bind(this);
    this._onDocTreeMenu       = this._onDocTreeMenu.bind(this);
  }

  t(key, params) {
    let raw = this.i18n?.[key] || key;
    if (params) {
      raw = raw.replace(/\{\{(\w+)\}\}/g, (_, k) =>
        Object.prototype.hasOwnProperty.call(params, k) ? String(params[k]) : ""
      );
    }
    return raw;
  }

  // -----------------------------------------------------------------------
  // Lifecycle
  // -----------------------------------------------------------------------
  async onload() {
    this.addIcons(SVG_ICONS);
    await this._loadSettings();
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

    this.eventBus.on("open-menu-link",              this._onLinkMenu);
    this.eventBus.on("open-menu-fileannotationref", this._onFileAnnotationMenu);
    this.eventBus.on("open-menu-doctree",           this._onDocTreeMenu);
  }

  onunload() {
    this.eventBus.off("open-menu-link",              this._onLinkMenu);
    this.eventBus.off("open-menu-fileannotationref", this._onFileAnnotationMenu);
    this.eventBus.off("open-menu-doctree",           this._onDocTreeMenu);
    for (const d of this._dialogs) { try { d.destroy(); } catch {} }
    this._dialogs.clear();
  }

  // -----------------------------------------------------------------------
  // Settings persistence
  // -----------------------------------------------------------------------
  async _loadSettings() {
    let raw = await this.loadData(STORAGE_KEY);
    if (typeof raw === "string") { try { raw = JSON.parse(raw); } catch { raw = {}; } }
    const d = (raw && typeof raw === "object" && !Array.isArray(raw)) ? raw : {};
    this.settings = {
      bridgeUrl:                normUrl(d.bridgeUrl || d.bridgeBaseUrl || "", ""),
      onlyofficeUrl:            normUrl(d.onlyofficeUrl || d.documentServerUrl || d.onlyOfficeUrl || "", ""),
      bridgeSecret:             String(d.bridgeSecret ?? ""),
      defaultMode:              d.defaultMode === "edit" ? "edit" : "view",
      enableEdit:               normBool(d.enableEdit, DEFAULT_SETTINGS.enableEdit),
      showInfoBar:              normBool(d.showInfoBar, DEFAULT_SETTINGS.showInfoBar),
      dialogWidth:              normSize(d.dialogWidth,  DEFAULT_SETTINGS.dialogWidth),
      dialogHeight:             normSize(d.dialogHeight, DEFAULT_SETTINGS.dialogHeight),
      enableLinkMenu:           normBool(d.enableLinkMenu,           DEFAULT_SETTINGS.enableLinkMenu),
      enableFileAnnotationMenu: normBool(d.enableFileAnnotationMenu, DEFAULT_SETTINGS.enableFileAnnotationMenu),
      enableDocTreeMenu:        normBool(d.enableDocTreeMenu,        DEFAULT_SETTINGS.enableDocTreeMenu),
    };
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

  // URL of SiYuan that the browser can access (used for fetch calls to SiYuan)
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

  _secretParam() {
    return this.settings.bridgeSecret
      ? `&secret=${encodeURIComponent(this.settings.bridgeSecret)}`
      : "";
  }

  // -----------------------------------------------------------------------
  // Open asset — core function
  // -----------------------------------------------------------------------
  async openAsset(assetPath, mode = "view", title) {
    const asset = normAssetPath(assetPath);
    if (!asset || !isSupported(asset)) {
      showMessage(this.t("message.unsupportedFile"), 5000, "error");
      return;
    }

    // PDF cannot be edited
    if (isPdf(asset) && mode === "edit") {
      showMessage(this.t("message.pdfNoEdit"), 4000, "info");
      mode = "view";
    }

    if (!this.settings.bridgeUrl) {
      showMessage(this.t("message.notConfigured"), 7000, "error");
      this.openSetting();
      return;
    }

    // Step 1: Health check
    try {
      const ctrl  = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 5000);
      const resp  = await fetch(`${this.settings.bridgeUrl}/health`, { signal: ctrl.signal });
      clearTimeout(timer);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    } catch (err) {
      showMessage(
        this.t("message.bridgeUnreachable", { url: this.settings.bridgeUrl, error: err.message }),
        7000, "error"
      );
      return;
    }

    // Step 2: Read file from SiYuan (browser → internal SiYuan, always reachable)
    let fileBlob;
    try {
      const siyuanBase = this._siyuanBaseUrl();
      if (!siyuanBase) throw new Error("Cannot determine SiYuan URL");
      const token = globalThis?.siyuan?.config?.api?.token;
      const headers = token ? { "Authorization": `Token ${token}` } : {};
      const sourceUrl = `${siyuanBase}/${asset}${asset.includes("?") ? "&" : "?"}t=${Date.now()}`;
      const resp = await fetch(sourceUrl, { headers, cache: "no-store" });
      if (!resp.ok) throw new Error(`SiYuan returned HTTP ${resp.status}`);
      fileBlob = await resp.blob();
    } catch (err) {
      showMessage(this.t("message.uploadFailed", { error: err.message }), 7000, "error");
      return;
    }

    // Step 3: Upload to Bridge (browser → public Bridge)
    try {
      const uploadUrl = `${this.settings.bridgeUrl}/upload?asset=${encodeURIComponent(asset)}${this._secretParam()}`;
      const resp = await fetch(uploadUrl, {
        method: "POST",
        body: fileBlob,
        headers: { "Content-Type": fileBlob.type || "application/octet-stream" },
      });
      if (!resp.ok) {
        const hint = resp.status === 404
          ? this.t("message.upload404Hint", { bridgeUrl: this.settings.bridgeUrl })
          : `Bridge returned HTTP ${resp.status}`;
        throw new Error(hint);
      }
    } catch (err) {
      showMessage(this.t("message.uploadFailed", { error: err.message }), 7000, "error");
      return;
    }

    // Step 4: Build editor iframe URL
    const user = this._userInfo();
    const params = new URLSearchParams({
      asset,
      mode,
      lang:     this._lang(),
      userId:   user.id,
      userName: user.name,
      title:    title || fileName(asset),
    });
    if (this.settings.onlyofficeUrl) params.set("oo", this.settings.onlyofficeUrl);
    if (this.settings.bridgeSecret) params.set("secret", this.settings.bridgeSecret);
    const frameSrc = `${this.settings.bridgeUrl}/editor?${params}`;

    // Step 5: Build dialog
    const displayName = title || fileName(asset);
    const modeLabel   = mode === "edit" ? this.t("badge.edit") : this.t("badge.view");
    const dialogTitle = mode === "edit"
      ? this.t("dialog.titleEdit", { name: displayName })
      : this.t("dialog.titleView", { name: displayName });

    // Step 6: Set up save-back message handler
    const syncSaved = () => this._syncSavedFile(asset);
    let liveSyncTimer = null;
    const msgHandler = (event) => {
      if (event.data?.type === "oo-bridge-saved" && event.data?.asset === asset) {
        syncSaved().catch(() => {});
      }
    };
    window.addEventListener("message", msgHandler);

    // While editing, poll saved state to tolerate delayed callback delivery.
    if (mode === "edit") {
      liveSyncTimer = setInterval(() => {
        syncSaved().catch(() => {});
      }, 4000);
    }

    const toolbarHtml = this.settings.showInfoBar
      ? `<div class="oo-bridge-dialog__toolbar">
    <span class="oo-bridge-badge oo-bridge-badge--${mode}">${escapeHtml(modeLabel)}</span>
    <code class="oo-bridge-dialog__path" title="${escapeHtml(asset)}">${escapeHtml(asset)}</code>
  </div>`
      : "";

    const dialog = new Dialog({
      title:  dialogTitle,
      width:  this.settings.dialogWidth,
      height: this.settings.dialogHeight,
      content: `<div class="oo-bridge-dialog">
  ${toolbarHtml}
  <iframe class="oo-bridge-dialog__frame" src="${escapeHtml(frameSrc)}"
    allow="clipboard-read; clipboard-write; fullscreen"></iframe>
</div>`,
      destroyCallback: () => {
        window.removeEventListener("message", msgHandler);
        this._dialogs.delete(dialog);
        if (liveSyncTimer) clearInterval(liveSyncTimer);

        if (mode === "edit") {
          // Keep syncing for a while after close because status=2 callback can be delayed.
          this._postCloseSyncAndCleanup(asset);
        } else {
          this._cleanupBridgeFile(asset);
        }
      },
    });
    this._dialogs.add(dialog);
  }

  // -----------------------------------------------------------------------
  // Sync saved file from Bridge back to SiYuan
  // -----------------------------------------------------------------------
  async _syncSavedFile(asset) {
    if (this._syncInFlight.get(asset)) return;
    this._syncInFlight.set(asset, true);

    const sp = this.settings.bridgeSecret
      ? `secret=${encodeURIComponent(this.settings.bridgeSecret)}&`
      : "";
    const url = `${this.settings.bridgeUrl}/saved?${sp}asset=${encodeURIComponent(asset)}&t=${Date.now()}`;
    let resp;
    try {
      resp = await fetch(url, { cache: "no-store" });
    } catch {
      this._syncInFlight.delete(asset);
      return; // Bridge unreachable — silently skip
    }

    if (resp.status === 204 || resp.status === 404) {
      this._syncInFlight.delete(asset);
      return; // No saved changes
    }
    if (!resp.ok) {
      console.warn(`[OnlyOffice Bridge] Bridge /saved returned HTTP ${resp.status} for "${asset}"`);
      this._syncInFlight.delete(asset);
      return;
    }

    const savedBlob = await resp.blob();
    const siyuanBase = this._siyuanBaseUrl();
    if (!siyuanBase) {
      this._syncInFlight.delete(asset);
      return;
    }

    try {
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
          method: "POST",
          body: form,
          headers,
          cache: "no-store",
        });
        let putJson = null;
        try { putJson = await putResp.json(); } catch {}

        if (putResp.ok && putJson && typeof putJson === "object" && putJson.code === 0) {
          syncOk = true;
          break;
        }
        lastErr = new Error(`path=${targetPath}, http=${putResp.status}, body=${JSON.stringify(putJson)}`);
      }

      if (!syncOk) throw lastErr || new Error("SiYuan putFile failed");
      console.info(`[OnlyOffice Bridge] Synced "${asset}" back to SiYuan`);
    } catch (err) {
      console.warn(`[OnlyOffice Bridge] Sync failed for "${asset}":`, err);
    } finally {
      this._syncInFlight.delete(asset);
    }
  }

  // -----------------------------------------------------------------------
  // Continue syncing after dialog close to avoid missing delayed callbacks.
  // -----------------------------------------------------------------------
  _postCloseSyncAndCleanup(asset) {
    const intervalMs = 3000;
    const maxDurationMs = 90000; // 90s grace window
    const deadline = Date.now() + maxDurationMs;

    const tick = () => {
      this._syncSavedFile(asset)
        .catch(() => {})
        .finally(() => {
          if (Date.now() < deadline) {
            setTimeout(tick, intervalMs);
          } else {
            this._cleanupBridgeFile(asset);
          }
        });
    };

    tick();
  }

  // -----------------------------------------------------------------------
  // Clean up Bridge memory after session
  // -----------------------------------------------------------------------
  async _cleanupBridgeFile(asset) {
    try {
      const sp = this.settings.bridgeSecret
        ? `secret=${encodeURIComponent(this.settings.bridgeSecret)}&`
        : "";
      await fetch(
        `${this.settings.bridgeUrl}/cleanup?${sp}asset=${encodeURIComponent(asset)}`,
        { method: "POST" }
      );
    } catch {
      // Silently ignore
    }
  }

  // -----------------------------------------------------------------------
  // Prompt open by manual path
  // -----------------------------------------------------------------------
  _promptOpen() {
    const input = globalThis?.prompt?.(this.t("prompt.assetPath"), "assets/");
    if (input == null) return;
    const p = normAssetPath(input);
    if (!p)            { showMessage(this.t("message.invalidAssetPath"), 5000, "error"); return; }
    if (!isSupported(p)) { showMessage(this.t("message.unsupportedFile"), 5000, "error"); return; }
    this.openAsset(p, this.settings.defaultMode);
  }

  // -----------------------------------------------------------------------
  // Context menu integration
  // -----------------------------------------------------------------------
  _addMenuItems(menu, assetPath, label) {
    if (!menu || !assetPath || !isSupported(assetPath)) return;
    if (!this.settings.bridgeUrl) return;

    const canEdit    = this.settings.enableEdit && !isPdf(assetPath);
    const defaultEdit = canEdit && this.settings.defaultMode === "edit";

    if (defaultEdit) {
      menu.addItem({
        icon:  ICON_EDIT,
        label: this.t("menu.edit"),
        click: () => this.openAsset(assetPath, "edit", label),
      });
    }
    menu.addItem({
      icon:  ICON_PREVIEW,
      label: this.t("menu.preview"),
      click: () => this.openAsset(assetPath, "view", label),
    });
    if (canEdit && !defaultEdit) {
      menu.addItem({
        icon:  ICON_EDIT,
        label: this.t("menu.edit"),
        click: () => this.openAsset(assetPath, "edit", label),
      });
    }
  }

  _onLinkMenu({ detail }) {
    try {
      if (!this.settings.enableLinkMenu) return;
      const { menu, element } = detail || {};
      if (!menu || !element) return;
      const asset = hrefToAsset(extractHref(element));
      if (!isSupported(asset)) return;
      this._addMenuItems(menu, asset, element.innerText?.trim() || fileName(asset));
    } catch (e) { console.error(e); }
  }

  _onFileAnnotationMenu({ detail }) {
    try {
      if (!this.settings.enableFileAnnotationMenu) return;
      const { menu, element } = detail || {};
      if (!menu || !element) return;
      const rawId = element.dataset?.id || tryAttr(element, "data-id") || "";
      const asset = normAssetPath(rawId);
      if (!isSupported(asset)) return;
      this._addMenuItems(menu, asset, element.innerText?.trim() || fileName(asset));
    } catch (e) { console.error(e); }
  }

  _onDocTreeMenu({ detail }) {
    try {
      if (!this.settings.enableDocTreeMenu) return;
      const { menu } = detail || {};
      if (!menu) return;
      const elems = toArray(detail?.elements ?? detail?.element);
      if (!elems.length) return;

      let asset = "", label = "";
      for (const el of elems) {
        if (!el) continue;
        const candidates = [];
        for (const attr of DOC_TREE_ATTRS) {
          const v = tryAttr(el, attr);
          if (v) candidates.push(v);
          const closest = el.closest?.("[data-path],[data-url],[data-id],[data-href]");
          if (closest) {
            const cv = tryAttr(closest, attr);
            if (cv) candidates.push(cv);
          }
        }
        for (const c of candidates) {
          const n = normAssetPath(c);
          if (isSupported(n)) { asset = n; break; }
        }
        if (asset) {
          label = tryAttr(el, "aria-label") || el.textContent?.trim() || fileName(asset);
          break;
        }
      }
      if (asset) this._addMenuItems(menu, asset, label);
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

    const mkSelect = (options) => {
      const el = document.createElement("select");
      el.className   = "b3-select";
      el.style.width = "160px";
      for (const [value, text] of options) {
        const opt = document.createElement("option");
        opt.value       = value;
        opt.textContent = text;
        el.appendChild(opt);
      }
      return el;
    };

    const mkSwitch = () => {
      const el = document.createElement("input");
      el.className = "b3-switch fn__flex-center";
      el.type      = "checkbox";
      return el;
    };

    const els = {
      bridgeUrl:                mkInput("http://your-public-server:6789"),
      onlyofficeUrl:            mkInput("http://your-public-server:8080"),
      bridgeSecret:             mkInput("", "password"),
      defaultMode:              mkSelect([
        ["view", this.t("settings.defaultModeView")],
        ["edit", this.t("settings.defaultModeEdit")],
      ]),
      enableEdit:               mkSwitch(),
      showInfoBar:              mkSwitch(),
      dialogWidth:              mkInput(DEFAULT_SETTINGS.dialogWidth),
      dialogHeight:             mkInput(DEFAULT_SETTINGS.dialogHeight),
      enableLinkMenu:           mkSwitch(),
      enableFileAnnotationMenu: mkSwitch(),
      enableDocTreeMenu:        mkSwitch(),
    };
    this._settingEls = els;

    this.setting = new Setting({
      width: "760px",
      confirmCallback: () => this._onSave(),
    });

    // Connection
    this.setting.addItem({
      title:               this.t("settings.bridgeUrl"),
      description:         this.t("settings.bridgeUrlDesc"),
      createActionElement: () => els.bridgeUrl,
    });
    this.setting.addItem({
      title:               this.t("settings.onlyofficeUrl"),
      description:         this.t("settings.onlyofficeUrlDesc"),
      createActionElement: () => els.onlyofficeUrl,
    });
    this.setting.addItem({
      title:               this.t("settings.bridgeSecret"),
      description:         this.t("settings.bridgeSecretDesc"),
      createActionElement: () => els.bridgeSecret,
    });

    // Mode
    this.setting.addItem({
      title:               this.t("settings.defaultMode"),
      description:         this.t("settings.defaultModeDesc"),
      createActionElement: () => els.defaultMode,
    });
    this.setting.addItem({
      title:               this.t("settings.enableEdit"),
      description:         this.t("settings.enableEditDesc"),
      createActionElement: () => els.enableEdit,
    });
    this.setting.addItem({
      title:               this.t("settings.showInfoBar"),
      description:         this.t("settings.showInfoBarDesc"),
      createActionElement: () => els.showInfoBar,
    });

    // Dialog size
    this.setting.addItem({
      title:               this.t("settings.dialogWidth"),
      description:         this.t("settings.dialogSizeDesc"),
      createActionElement: () => els.dialogWidth,
    });
    this.setting.addItem({
      title:               this.t("settings.dialogHeight"),
      description:         this.t("settings.dialogSizeDesc"),
      createActionElement: () => els.dialogHeight,
    });

    // Menu toggles
    this.setting.addItem({
      title:               this.t("settings.enableLinkMenu"),
      description:         this.t("settings.enableLinkMenuDesc"),
      createActionElement: () => els.enableLinkMenu,
    });
    this.setting.addItem({
      title:               this.t("settings.enableFileAnnotationMenu"),
      description:         this.t("settings.enableFileAnnotationMenuDesc"),
      createActionElement: () => els.enableFileAnnotationMenu,
    });
    this.setting.addItem({
      title:               this.t("settings.enableDocTreeMenu"),
      description:         this.t("settings.enableDocTreeMenuDesc"),
      createActionElement: () => els.enableDocTreeMenu,
    });

    // Auto-save on change
    for (const el of Object.values(els)) el.addEventListener("change", () => this._onSave());

    this._syncInputs();
  }

  _syncInputs() {
    const e = this._settingEls;
    if (!e.bridgeUrl) return;
    e.bridgeUrl.value                  = this.settings.bridgeUrl;
    e.onlyofficeUrl.value              = this.settings.onlyofficeUrl;
    e.bridgeSecret.value               = this.settings.bridgeSecret;
    e.defaultMode.value                = this.settings.defaultMode;
    e.enableEdit.checked               = this.settings.enableEdit;
    e.showInfoBar.checked              = this.settings.showInfoBar;
    e.dialogWidth.value                = this.settings.dialogWidth;
    e.dialogHeight.value               = this.settings.dialogHeight;
    e.enableLinkMenu.checked           = this.settings.enableLinkMenu;
    e.enableFileAnnotationMenu.checked = this.settings.enableFileAnnotationMenu;
    e.enableDocTreeMenu.checked        = this.settings.enableDocTreeMenu;
  }

  async _onSave() {
    const e = this._settingEls;
    if (!e.bridgeUrl) return;
    this.settings = {
      bridgeUrl:                normUrl(e.bridgeUrl.value, ""),
      onlyofficeUrl:            normUrl(e.onlyofficeUrl.value, ""),
      bridgeSecret:             e.bridgeSecret.value.trim(),
      defaultMode:              e.defaultMode.value === "edit" ? "edit" : "view",
      enableEdit:               !!e.enableEdit.checked,
      showInfoBar:              !!e.showInfoBar.checked,
      dialogWidth:              normSize(e.dialogWidth.value,  DEFAULT_SETTINGS.dialogWidth),
      dialogHeight:             normSize(e.dialogHeight.value, DEFAULT_SETTINGS.dialogHeight),
      enableLinkMenu:           !!e.enableLinkMenu.checked,
      enableFileAnnotationMenu: !!e.enableFileAnnotationMenu.checked,
      enableDocTreeMenu:        !!e.enableDocTreeMenu.checked,
    };
    await this._saveSettings();
    this._syncInputs();
    showMessage(this.t("message.settingsSaved"), 2500, "info");
  }
}

module.exports = OnlyOfficeBridgePlugin;
module.exports.default = OnlyOfficeBridgePlugin;
