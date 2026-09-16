(function () {
    "use strict";

    const GRAPH_RESOURCE = "https://graph.microsoft.com";
    const GRAPH_VERSION = "v1.0";
    const CLIENT_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

    const dom = {
        pickBtn: document.getElementById("btn-pick"),
        clearBtn: document.getElementById("btn-clear"),
        confirmBar: document.getElementById("confirm-clear"),
        confirmYes: document.getElementById("confirm-clear-yes"),
        confirmNo: document.getElementById("confirm-clear-no"),
        errorBanner: document.getElementById("error-banner"),
        errorText: document.getElementById("error-text"),
        errorDismiss: document.getElementById("error-dismiss"),
        fileDisplay: document.getElementById("file-display"),
        debugStatus: document.getElementById("debug-status"),
        container: document.getElementById("root-container"),
        browserPanel: document.getElementById("browser-panel"),
        siteSelectorRow: document.getElementById("site-selector-row"),
        siteSelector: document.getElementById("site-selector"),
        breadcrumb: document.getElementById("breadcrumb"),
        browserLoading: document.getElementById("browser-loading"),
        browserError: document.getElementById("browser-error"),
        browserList: document.getElementById("browser-list"),
        browserConfirm: document.getElementById("browser-confirm"),
        browserCancel: document.getElementById("browser-cancel"),
        browserSelectionCount: document.getElementById("browser-selection-count"),
    };

    // Kontent.ai always embeds custom elements as a child iframe. Bail out
    // early (before any auth/browser wiring) if this page was opened directly.
    if (window === window.top) {
        if (dom.container) {
            dom.container.textContent = "This page only works embedded as a Kontent.ai custom element.";
        }
        return;
    }

    let config = null;
    let isDisabled = false;
    let msalInstance = null;
    let msalReady = false;
    let resizeObserver = null;

    // Normalized in-memory copy of the last rendered value (kept in sync by
    // renderFiles), so refreshStoredLinks() always has something to re-check
    // against Graph without re-parsing element.value.
    let storedFiles = [];

    // Per-site Graph state, keyed by the site's config index. Populated lazily
    // the first time each site is opened, so opening the browser never fetches
    // more than the site currently in view.
    const siteState = new Map();

    // --- status / error UI -------------------------------------------------

    function setStatus(msg) {
        if (config && config.debug) {
            dom.debugStatus.textContent = msg;
            dom.debugStatus.classList.remove("hidden");
        }
        console.log("[SharePointPickerAlt]", msg);
    }

    function showError(msg) {
        dom.errorText.textContent = msg;
        dom.errorBanner.classList.remove("hidden");
        console.error("[SharePointPickerAlt]", msg);
    }

    function clearErrorBanner() {
        dom.errorBanner.classList.add("hidden");
        dom.errorText.textContent = "";
    }

    // --- config --------------------------------------------------------------

    // Unlike the original picker, there is no "sharePointTenant" that implies
    // browsing every site under that tenant. Each entry in `sites` is a site
    // an Entra admin has explicitly granted this app's Sites.Selected
    // permission against (see README) - nothing else is reachable.
    function validateConfig(rawConfig) {
        const cfg = rawConfig || {};
        const errors = [];

        if (!cfg.clientId || !CLIENT_ID_PATTERN.test(cfg.clientId)) {
            errors.push("\"clientId\" is missing or is not a valid Azure AD application (client) ID.");
        }
        if (!cfg.tenant || typeof cfg.tenant !== "string") {
            errors.push("\"tenant\" is missing. Use your Entra tenant ID (GUID) or verified domain, e.g. \"contoso.onmicrosoft.com\".");
        }
        if (!Array.isArray(cfg.sites) || cfg.sites.length === 0) {
            errors.push("\"sites\" must be a non-empty array of { label, hostname, path } objects.");
        } else {
            cfg.sites.forEach((site, i) => {
                if (!site || typeof site.hostname !== "string" || !site.hostname) {
                    errors.push(`sites[${i}].hostname is missing, e.g. "contoso.sharepoint.com".`);
                }
                if (!site || typeof site.path !== "string" || !site.path.startsWith("/")) {
                    errors.push(`sites[${i}].path is missing or must start with "/", e.g. "/sites/SalesFirst".`);
                }
            });
        }
        if (cfg.selectionMode && cfg.selectionMode !== "single" && cfg.selectionMode !== "multiple") {
            errors.push("\"selectionMode\" must be either \"single\" or \"multiple\".");
        }

        if (errors.length) {
            return { valid: false, errors };
        }

        return {
            valid: true,
            clientId: cfg.clientId,
            tenant: cfg.tenant,
            sites: cfg.sites.map((site, i) => ({
                key: String(i),
                label: site.label || site.path,
                hostname: site.hostname,
                path: site.path,
            })),
            selectionMode: cfg.selectionMode === "single" ? "single" : "multiple",
            debug: !!cfg.debug,
        };
    }

    // --- MSAL / auth -----------------------------------------------------------
    // Only ever requests the Graph "Sites.Selected" scope. There is no
    // SharePoint-resource token (AllSites.Read / MyFiles.Read) and no
    // Sites.Read.All / Files.Read.All - access to a site only works at all
    // once a tenant admin has granted this app's service principal a role on
    // that specific site (see README "Grant per-site access").

    async function initMsal() {
        try {
            msalInstance = new msal.PublicClientApplication({
                auth: {
                    clientId: config.clientId,
                    authority: `https://login.microsoftonline.com/${config.tenant}`,
                    redirectUri: window.location.origin + window.location.pathname,
                },
                cache: { cacheLocation: "sessionStorage" },
            });
            await msalInstance.initialize();
            msalReady = true;
            setStatus("MSAL ready | accounts: " + msalInstance.getAllAccounts().length);
        } catch (e) {
            showError("Microsoft authentication library failed to initialize: " + e.message);
        }
    }

    function graphScopes() {
        return [`${GRAPH_RESOURCE}/Sites.Selected`];
    }

    async function ensureAuthenticated() {
        let account = msalInstance.getAllAccounts()[0];
        if (!account) {
            const authRes = await msalInstance.loginPopup({ scopes: graphScopes() });
            account = authRes.account;
        }
        try {
            await msalInstance.acquireTokenSilent({ scopes: graphScopes(), account });
        } catch (err) {
            await msalInstance.acquireTokenPopup({ scopes: graphScopes(), account });
        }
    }

    async function getGraphToken() {
        const account = msalInstance.getAllAccounts()[0];
        if (!account) throw new Error("No signed-in Microsoft account.");
        try {
            const r = await msalInstance.acquireTokenSilent({ scopes: graphScopes(), account });
            return r.accessToken;
        } catch (err) {
            const r = await msalInstance.acquireTokenPopup({ scopes: graphScopes(), account });
            return r.accessToken;
        }
    }

    async function graphGet(path) {
        const token = await getGraphToken();
        const res = await fetch(`${GRAPH_RESOURCE}/${GRAPH_VERSION}${path}`, {
            headers: { Authorization: `Bearer ${token}` },
        });
        if (!res.ok) {
            let detail = "";
            try {
                const body = await res.json();
                detail = body?.error?.message ? `: ${body.error.message}` : "";
            } catch (e) { /* ignore parse failure, use status only */ }

            if (res.status === 403 || res.status === 401) {
                throw new Error(
                    "Access denied by SharePoint/Graph" + detail +
                    ". Check that a tenant admin has granted this app's Sites.Selected permission on this site (see README)."
                );
            }
            throw new Error(`Graph request failed (${res.status})${detail}`);
        }
        return res.json();
    }

    // --- site / drive browsing ------------------------------------------------

    async function getSiteState(siteConfig) {
        if (siteState.has(siteConfig.key)) return siteState.get(siteConfig.key);

        setStatus(`Resolving site ${siteConfig.hostname}${siteConfig.path} ...`);
        const site = await graphGet(`/sites/${siteConfig.hostname}:${siteConfig.path}`);
        const state = { siteId: site.id, siteName: site.displayName || siteConfig.label };
        siteState.set(siteConfig.key, state);
        return state;
    }

    async function listChildren(siteId, folderId) {
        const base = folderId
            ? `/sites/${siteId}/drive/items/${folderId}/children`
            : `/sites/${siteId}/drive/root/children`;
        const select = "id,name,file,folder,webUrl,lastModifiedDateTime,lastModifiedBy,parentReference";
        const data = await graphGet(`${base}?$select=${select}&$top=200`);
        return data.value || [];
    }

    // --- stored value / rendering (same shape as the original element) -------

    function renderFiles(fileData) {
        const files = fileData ? (Array.isArray(fileData) ? fileData : [fileData]) : [];
        storedFiles = files;

        dom.fileDisplay.innerHTML = "";

        if (files.length === 0) {
            dom.fileDisplay.classList.add("hidden");
            dom.pickBtn.textContent = "Select SharePoint Files";
            dom.clearBtn.classList.add("hidden");
            hideConfirmClear();
            return;
        }

        files.forEach((file) => {
            const fileItem = document.createElement("div");
            fileItem.className = "file-item";

            const linkEl = document.createElement("a");
            linkEl.className = "file-link";
            linkEl.href = isSafeUrl(file.url) ? file.url : "#";
            linkEl.target = "_blank";
            linkEl.rel = "noopener noreferrer";
            linkEl.innerText = file.name || "View linked SharePoint file";

            const metaEl = document.createElement("div");
            metaEl.className = "file-meta";
            const dateStr = file.lastModified
                ? new Date(file.lastModified).toLocaleDateString(undefined, {
                    year: "numeric",
                    month: "short",
                    day: "numeric",
                    hour: "2-digit",
                    minute: "2-digit",
                })
                : "Unknown date";
            metaEl.innerText = `Updated ${dateStr} by ${file.author || "Unknown author"}`;

            fileItem.appendChild(linkEl);
            fileItem.appendChild(metaEl);
            dom.fileDisplay.appendChild(fileItem);
        });

        dom.fileDisplay.classList.remove("hidden");
        dom.pickBtn.textContent = "Change Files";
        if (!isDisabled) dom.clearBtn.classList.remove("hidden");
    }

    function isSafeUrl(url) {
        if (typeof url !== "string") return false;
        try {
            return new URL(url).protocol === "https:";
        } catch (e) {
            return false;
        }
    }

    function saveValue(fileData) {
        CustomElement.setValue(fileData ? JSON.stringify(fileData) : null);
        renderFiles(fileData);
    }

    // --- self-healing links --------------------------------------------------
    // The stored value's `url` is a snapshot from whenever the file was picked.
    // Renaming or moving the file in SharePoint changes its webUrl, so that
    // snapshot goes stale (404) even though the file itself is untouched. The
    // stored `id` + `driveId` are stable identifiers Graph can always resolve
    // back to the file's *current* location, so we use them to re-fetch fresh
    // metadata and, if anything changed, re-save the corrected value.

    async function resolveLiveMetadata(file) {
        if (!file || !file.driveId || !file.id) return file; // nothing to resolve against
        try {
            const item = await graphGet(
                `/drives/${encodeURIComponent(file.driveId)}/items/${encodeURIComponent(file.id)}` +
                "?$select=id,name,webUrl,lastModifiedDateTime,lastModifiedBy"
            );
            return {
                ...file,
                name: item.name || file.name,
                url: item.webUrl || file.url,
                author: item.lastModifiedBy?.user?.displayName || file.author,
                lastModified: item.lastModifiedDateTime || file.lastModified,
            };
        } catch (e) {
            // File may have been deleted, or this account may no longer have
            // access - keep the last-known-good stored value rather than
            // erroring the whole element out over a background refresh.
            setStatus(`Could not refresh link for "${file.name || file.id}": ${e.message}`);
            return file;
        }
    }

    // silent: true => never prompt for sign-in, only use an already-cached
    // MSAL session (used opportunistically on load). Called without options
    // once the editor has just interactively signed in (opening the picker),
    // where a prompt is already expected and acceptable.
    async function refreshStoredLinks(options) {
        const silent = !!(options && options.silent);
        if (!storedFiles.length || !msalReady) return;

        const account = msalInstance.getAllAccounts()[0];
        if (!account) return; // never prompt from here - best effort only

        if (silent) {
            try {
                await msalInstance.acquireTokenSilent({ scopes: graphScopes(), account });
            } catch (e) {
                return; // no usable cached session; don't interrupt with a popup
            }
        }

        const refreshed = await Promise.all(storedFiles.map(resolveLiveMetadata));
        if (JSON.stringify(refreshed) !== JSON.stringify(storedFiles)) {
            setStatus(`Refreshed ${refreshed.length} file link(s) from SharePoint.`);
            saveValue(refreshed);
        }
    }

    function updateHeight() {
        if (dom.container) {
            CustomElement.setHeight(dom.container.getBoundingClientRect().height);
        }
    }

    // --- clear confirmation (no window.confirm - not available inside the
    // sandboxed custom element iframe) ---------------------------------------

    function showConfirmClear() {
        dom.confirmBar.classList.remove("hidden");
    }

    function hideConfirmClear() {
        dom.confirmBar.classList.add("hidden");
    }

    // --- inline browser panel -------------------------------------------------
    // Replaces the Microsoft File Picker v8 popup entirely. Runs inside the
    // element's own iframe against Microsoft Graph only, scoped to whichever
    // site is currently selected.

    const browser = {
        open: false,
        siteConfig: null,
        siteId: null,
        // Breadcrumb stack of {id, name}; id === null means the site's root.
        path: [],
        items: [],
        selected: new Map(), // id -> graph drive item
    };

    function currentFolderId() {
        return browser.path.length ? browser.path[browser.path.length - 1].id : null;
    }

    function renderSiteSelector() {
        dom.siteSelector.innerHTML = "";
        config.sites.forEach((site) => {
            const opt = document.createElement("option");
            opt.value = site.key;
            opt.textContent = site.label;
            dom.siteSelector.appendChild(opt);
        });
        dom.siteSelectorRow.classList.toggle("hidden", config.sites.length <= 1);
    }

    function renderBreadcrumb() {
        dom.breadcrumb.innerHTML = "";

        const rootBtn = document.createElement("button");
        rootBtn.type = "button";
        rootBtn.className = "breadcrumb-item";
        rootBtn.textContent = browser.siteConfig ? browser.siteConfig.label : "Site";
        if (browser.path.length === 0) rootBtn.disabled = true;
        rootBtn.addEventListener("click", () => navigateTo(0));
        dom.breadcrumb.appendChild(rootBtn);

        browser.path.forEach((crumb, index) => {
            const sep = document.createElement("span");
            sep.className = "breadcrumb-sep";
            sep.textContent = "/";
            dom.breadcrumb.appendChild(sep);

            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "breadcrumb-item";
            btn.textContent = crumb.name;
            if (index === browser.path.length - 1) btn.disabled = true;
            btn.addEventListener("click", () => navigateTo(index + 1));
            dom.breadcrumb.appendChild(btn);
        });
    }

    // depth === 0 means "back to the site root"; depth === N means
    // "truncate the breadcrumb to its first N entries".
    function navigateTo(depth) {
        browser.path = browser.path.slice(0, depth);
        loadCurrentFolder();
    }

    function updateSelectionUi() {
        const count = browser.selected.size;
        dom.browserSelectionCount.textContent = count === 0
            ? "Nothing selected yet"
            : `${count} file${count === 1 ? "" : "s"} selected`;
        dom.browserConfirm.disabled = count === 0;
    }

    function toggleSelect(item) {
        if (browser.selected.has(item.id)) {
            browser.selected.delete(item.id);
        } else {
            if (config.selectionMode === "single") browser.selected.clear();
            browser.selected.set(item.id, item);
        }
        renderList();
        updateSelectionUi();
    }

    function renderList() {
        dom.browserList.innerHTML = "";

        if (browser.items.length === 0) {
            const empty = document.createElement("div");
            empty.className = "browser-row-empty";
            empty.textContent = "This folder is empty.";
            dom.browserList.appendChild(empty);
            return;
        }

        // Folders first, then files, both alphabetical - easier to scan than
        // Graph's default order.
        const sorted = [...browser.items].sort((a, b) => {
            const aFolder = !!a.folder, bFolder = !!b.folder;
            if (aFolder !== bFolder) return aFolder ? -1 : 1;
            return a.name.localeCompare(b.name);
        });

        sorted.forEach((item) => {
            const row = document.createElement("div");
            row.className = "browser-row" + (item.folder ? " is-folder" : "");

            const icon = document.createElement("span");
            icon.className = "browser-row-icon";
            icon.textContent = item.folder ? "📁" : "📄";
            row.appendChild(icon);

            const name = document.createElement("span");
            name.className = "browser-row-name";
            name.textContent = item.name;
            row.appendChild(name);

            if (!item.folder) {
                const checkbox = document.createElement("input");
                checkbox.type = config.selectionMode === "single" ? "radio" : "checkbox";
                checkbox.checked = browser.selected.has(item.id);
                checkbox.addEventListener("click", (e) => e.stopPropagation());
                checkbox.addEventListener("change", () => toggleSelect(item));
                row.insertBefore(checkbox, row.firstChild);
            }

            row.addEventListener("click", () => {
                if (item.folder) {
                    browser.path.push({ id: item.id, name: item.name });
                    loadCurrentFolder();
                } else {
                    toggleSelect(item);
                }
            });

            dom.browserList.appendChild(row);
        });
    }

    async function loadCurrentFolder() {
        dom.browserError.classList.add("hidden");
        dom.browserLoading.classList.remove("hidden");
        dom.browserList.innerHTML = "";
        renderBreadcrumb();
        updateHeight();

        try {
            browser.items = await listChildren(browser.siteId, currentFolderId());
            renderList();
        } catch (e) {
            dom.browserError.textContent = e.message;
            dom.browserError.classList.remove("hidden");
        } finally {
            dom.browserLoading.classList.add("hidden");
            updateHeight();
        }
    }

    async function openSite(siteConfig) {
        browser.siteConfig = siteConfig;
        browser.path = [];
        browser.siteId = null;
        dom.browserError.classList.add("hidden");
        dom.browserLoading.classList.remove("hidden");
        dom.browserList.innerHTML = "";
        renderBreadcrumb();
        updateHeight();

        try {
            const state = await getSiteState(siteConfig);
            browser.siteId = state.siteId;
            await loadCurrentFolder();
        } catch (e) {
            dom.browserLoading.classList.add("hidden");
            dom.browserError.textContent = e.message;
            dom.browserError.classList.remove("hidden");
            updateHeight();
        }
    }

    function closeBrowser() {
        browser.open = false;
        browser.selected.clear();
        dom.browserPanel.classList.add("hidden");
        dom.pickBtn.disabled = isDisabled;
        dom.pickBtn.textContent = dom.fileDisplay.classList.contains("hidden") ? "Select SharePoint Files" : "Change Files";
        updateHeight();
    }

    async function openBrowser() {
        clearErrorBanner();
        browser.open = true;
        browser.selected.clear();
        dom.browserPanel.classList.remove("hidden");
        updateSelectionUi();
        renderSiteSelector();
        dom.siteSelector.value = config.sites[0].key;
        updateHeight();
        await openSite(config.sites[0]);
    }

    dom.siteSelector.addEventListener("change", () => {
        // Selections are keyed by Graph item id, which is unique per drive, so
        // switching sites does not risk collisions - keep whatever the editor
        // already checked in other sites instead of silently dropping it.
        const site = config.sites.find((s) => s.key === dom.siteSelector.value);
        if (site) openSite(site);
    });

    dom.browserCancel.addEventListener("click", closeBrowser);

    dom.browserConfirm.addEventListener("click", () => {
        // Each picker session fully replaces the stored value, same contract
        // as the original element: the browser has no way to show files
        // picked in an earlier session as pre-checked, so silently merging
        // with the previous value here would only let editors add files,
        // never remove one, without using "Remove" to clear everything.
        const picked = Array.from(browser.selected.values()).map((item) => ({
            name: item.name,
            url: item.webUrl || null,
            id: item.id,
            driveId: item.parentReference?.driveId || null,
            author: item.lastModifiedBy?.user?.displayName || "Unknown",
            lastModified: item.lastModifiedDateTime || null,
        }));

        saveValue(picked);
        closeBrowser();
    });

    // --- event wiring ---------------------------------------------------------

    dom.errorDismiss.addEventListener("click", clearErrorBanner);

    dom.clearBtn.addEventListener("click", () => {
        clearErrorBanner();
        showConfirmClear();
    });

    dom.confirmNo.addEventListener("click", hideConfirmClear);

    dom.confirmYes.addEventListener("click", () => {
        hideConfirmClear();
        saveValue(null);
    });

    dom.pickBtn.addEventListener("click", async () => {
        clearErrorBanner();

        if (!msalReady) {
            showError("Authentication is still initializing. Please try again in a moment.");
            return;
        }

        dom.pickBtn.disabled = true;
        dom.pickBtn.textContent = "Authenticating...";

        try {
            await ensureAuthenticated();
            await refreshStoredLinks(); // heal any renamed/moved links now that we have a token
            await openBrowser();
            dom.pickBtn.textContent = "Browsing...";
        } catch (error) {
            showError("Microsoft authentication failed. Please ensure popups are allowed and try again.");
            dom.pickBtn.disabled = isDisabled;
            dom.pickBtn.textContent = dom.fileDisplay.classList.contains("hidden") ? "Select SharePoint Files" : "Change Files";
        }
    });

    // --- bootstrap -------------------------------------------------------------

    function applyDisabledState(disabled) {
        isDisabled = disabled;
        dom.pickBtn.disabled = disabled;
        if (disabled) {
            dom.clearBtn.classList.add("hidden");
            hideConfirmClear();
            if (browser.open) closeBrowser();
        } else if (!dom.fileDisplay.classList.contains("hidden")) {
            dom.clearBtn.classList.remove("hidden");
        }
    }

    CustomElement.init((element, context) => {
        const result = validateConfig(element.config);

        if (!result.valid) {
            showError(
                "This custom element is not configured correctly:\n" +
                result.errors.join("\n") +
                "\nSee the README for the expected configuration JSON."
            );
            dom.pickBtn.disabled = true;
            updateHeight();
            return;
        }

        config = result;
        if (config.debug) dom.debugStatus.classList.remove("hidden");
        setStatus("Origin: " + window.location.origin + " | MSAL: loading...");

        if (element.value) {
            try {
                renderFiles(JSON.parse(element.value));
            } catch (e) {
                console.error("[SharePointPickerAlt] Failed to parse saved value", e);
            }
        }

        applyDisabledState(element.disabled);
        // Best-effort, silent link healing: only runs if a cached MSAL session
        // already exists (e.g. the editor signed in earlier in this browser
        // tab) - never prompts on its own.
        initMsal().then(() => refreshStoredLinks({ silent: true }));
    });

    if (typeof CustomElement.onDisabledChanged === "function") {
        CustomElement.onDisabledChanged(applyDisabledState);
    }

    resizeObserver = new ResizeObserver(() => updateHeight());
    resizeObserver.observe(dom.container);
})();
