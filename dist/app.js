"use strict";

(function bootstrapInvoiceAccessManager() {
  const STORAGE_PREFIX = "rocketlane-invoice-access";
  const STORAGE_VERSION = "v2";
  const LOG_LIMIT = 200;
  const SDK_WAIT_MS = 4500;
  const SDK_POLL_INTERVAL_MS = 120;
  const MAX_PDF_SCRUB_PAGES = 6;
  const PDF_WORKER_CDN = "./pdf.worker.min.js";

  const SOURCE_PROJECT_NAMES = [
    "expert advisor program invoices",
    "expert advisors program invoices",
  ];

  const SAMPLE_PDF_DATA_URL =
    "data:application/pdf;base64,JVBERi0xLjQKJeLjz9MKMSAwIG9iago8PCAvVHlwZSAvQ2F0YWxvZyAvUGFnZXMgMiAwIFIgPj4KZW5kb2JqCjIgMCBvYmoKPDwgL1R5cGUgL1BhZ2VzIC9LaWRzIFszIDAgUl0gL0NvdW50IDEgPj4KZW5kb2JqCjMgMCBvYmoKPDwgL1R5cGUgL1BhZ2UgL1BhcmVudCAyIDAgUiAvTWVkaWFCb3ggWzAgMCA2MTIgNzkyXSAvQ29udGVudHMgNCAwIFIgL1Jlc291cmNlcyA8PCAvRm9udCA8PCAvRjEgNSAwIFIgPj4gPj4gPj4KZW5kb2JqCjQgMCBvYmoKPDwgL0xlbmd0aCAxMzYgPj4Kc3RyZWFtCkJUCi9GMSAxOCBUZgo3MiA3MzAgVGQKKFNhbXBsZSBJbnZvaWNlIElOVi0wMDAxKSBUagowIC0yOCBUZAooUHJldmlldyBmcm9tIEludm9pY2UgQWNjZXNzIE1hbmFnZXIpIFRqCjAgLTIyIFRkCihEYXRlOiAyMDI2LTAzLTAzKSBUagpFVAplbmRzdHJlYW0KZW5kb2JqCjUgMCBvYmoKPDwgL1R5cGUgL0ZvbnQgL1N1YnR5cGUgL1R5cGUxIC9CYXNlRm9udCAvSGVsdmV0aWNhID4+CmVuZG9iagp4cmVmCjAgNgowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwMTUgMDAwMDAgbiAKMDAwMDAwMDA2NCAwMDAwMCBuIAowMDAwMDAwMTIxIDAwMDAwIG4gCjAwMDAwMDAyNDcgMDAwMDAgbiAKMDAwMDAwMDQzMyAwMDAwMCBuIAp0cmFpbGVyCjw8IC9TaXplIDYgL1Jvb3QgMSAwIFIgPj4Kc3RhcnR4cmVmCjUwMwolJUVPRgo=";

  const EMAIL_PATTERN = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  const EMAIL_VALIDATION_PATTERN = /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i;

  const SUGGESTIONS = {
    RUNTIME_INIT_FAILED:
      "Verify the app is opened inside Rocketlane and the SDK script is loaded correctly.",
    SOURCE_PROJECTS_NOT_FOUND:
      "Create the source project named: Expert Advisor Program Invoices.",
    SOURCE_FETCH_FAILED:
      "Confirm Rocketlane API endpoints are reachable from this workspace and retry.",
    PERMISSION_FETCH_FAILED:
      "Confirm team member APIs are accessible so permission-set based access can be enforced.",
    SOURCE_INVOICES_NOT_FOUND:
      "Add PDF invoice files/documents to the source project to populate this table.",
    PDF_LIB_UNAVAILABLE:
      "Allow access to PDF.js or bundle PDF.js assets with this app package.",
    PDF_PREVIEW_FAILED:
      "Verify the invoice file URL is valid and accessible for this user session.",
    PDF_SCRUB_FAILED:
      "The PDF could not be scanned for email addresses. Verify file access and PDF text content.",
  };

  const state = {
    client: null,
    connected: false,
    context: null,
    rawUser: null,
    rawAccount: null,
    rawProject: null,
    storageKey: "",
    logs: [],
    invoices: [],
    sourceProjects: [],
    teamMembers: [],
    syncDiagnostics: {},
    permissionHint: null,
    selectedInvoiceId: null,
    searchQuery: "",
    searchInsight: "",
    searchServerCheckedQuery: "",
    searchServerMatchedCount: null,
    searchVerifyTimer: 0,
    pageSize: 25,
    currentPage: 1,
    sortBy: "issueDate",
    sortDirection: "desc",
    filters: {
      invoiceStatus: "",
      invoiceNumber: "",
      ownerName: "",
      amount: "",
      accountName: "",
      issueDate: "",
      dueDate: "",
    },
    activeTab: "invoices",
    syncStatus: "Initializing...",
    invoicePreviewCache: {},
    exportMode: "one",
    exportInsight: "",
    access: {
      role: "non_admin",
      roleLabel: "Restricted",
      isAdmin: false,
      email: "",
      displayName: "",
    },
  };

  window.__invoiceAccessBuild = "search-preview-20260305";
  window.__invoiceAccessDebug = {
    reason: "booting",
    connected: false,
    access: state.access,
  };

  const refs = {};

  document.addEventListener("DOMContentLoaded", () => {
    initializeApp().catch((error) => {
      console.error("Invoice app failed to initialize", error);
      if (refs.connectionBadge) {
        refs.connectionBadge.textContent = "Initialization failed";
      }
      appendLog(
        "RUNTIME_INIT_FAILED",
        "Invoice manager initialization failed.",
        error
      );
      updateDebugState("init-catch-error");
      renderAll();
    });
  });

  async function initializeApp() {
    cacheDomReferences();
    bindEvents();

    const runtime = await initializeRuntime();
    state.client = runtime.client;
    state.connected = runtime.connected;
    state.context = runtime.context;
    state.rawUser = runtime.rawUser;
    state.rawAccount = runtime.rawAccount;
    state.rawProject = runtime.rawProject;
    state.storageKey = createLogsStorageKey(runtime.context);
    loadLogs();

    const permissionHint = await fetchPermissionHintFromSdk();
    state.permissionHint = permissionHint;
    state.access = deriveAccessProfile(
      state.rawUser,
      state.rawAccount,
      runtime.context,
      permissionHint
    );
    updateDebugState("post-init");

    updateHeader();
    configureUiForAccess();
    await refreshInvoicesFromSource();
    updateDebugState("post-sync");
    ensureSelectedInvoice();
    renderAll();
  }

  function cacheDomReferences() {
    refs.scopeText = document.getElementById("scopeText");
    refs.connectionBadge = document.getElementById("connectionBadge");
    refs.roleBadge = document.getElementById("roleBadge");
    refs.tabInvoicesButton = document.getElementById("tabInvoicesButton");
    refs.tabLogsButton = document.getElementById("tabLogsButton");
    refs.tabInvoices = document.getElementById("tabInvoices");
    refs.tabLogs = document.getElementById("tabLogs");
    refs.layoutGrid = document.getElementById("layoutGrid");
    refs.contextPanel = document.getElementById("contextPanel");
    refs.syncStatus = document.getElementById("syncStatus");
    refs.searchInput = document.getElementById("searchInput");
    refs.visibilitySummary = document.getElementById("visibilitySummary");
    refs.searchInsight = document.getElementById("searchInsight");
    refs.filterInvoiceStatus = document.getElementById("filterInvoiceStatus");
    refs.filterInvoiceNumber = document.getElementById("filterInvoiceNumber");
    refs.filterProjectManager = document.getElementById("filterProjectManager");
    refs.filterAmount = document.getElementById("filterAmount");
    refs.filterAccount = document.getElementById("filterAccount");
    refs.filterIssueDate = document.getElementById("filterIssueDate");
    refs.filterDueDate = document.getElementById("filterDueDate");
    refs.clearFiltersButton = document.getElementById("clearFiltersButton");
    refs.exportModeSelect = document.getElementById("exportModeSelect");
    refs.downloadZipButton = document.getElementById("downloadZipButton");
    refs.exportInsight = document.getElementById("exportInsight");
    refs.invoiceStats = document.getElementById("invoiceStats");
    refs.sortButtons = Array.from(document.querySelectorAll(".sort-button"));
    refs.invoiceTableBody = document.getElementById("invoiceTableBody");
    refs.invoiceEmptyState = document.getElementById("invoiceEmptyState");
    refs.invoiceEmptyTitle = refs.invoiceEmptyState.querySelector("h3");
    refs.invoiceEmptyBody = refs.invoiceEmptyState.querySelector("p");
    refs.paginationBar = document.getElementById("paginationBar");
    refs.paginationPrevButton = document.getElementById("paginationPrevButton");
    refs.paginationInfo = document.getElementById("paginationInfo");
    refs.paginationNextButton = document.getElementById("paginationNextButton");
    refs.selectedInvoiceSummary = document.getElementById("selectedInvoiceSummary");
    refs.sourceProjectsText = document.getElementById("sourceProjectsText");
    refs.logsList = document.getElementById("logsList");
    refs.logsEmptyState = document.getElementById("logsEmptyState");
    refs.logsInfoText = document.getElementById("logsInfoText");
    refs.logsAccessNotice = document.getElementById("logsAccessNotice");
    refs.clearLogsButton = document.getElementById("clearLogsButton");
    refs.pdfModal = document.getElementById("pdfModal");
    refs.modalTitle = document.getElementById("modalTitle");
    refs.modalPdfFrame = document.getElementById("modalPdfFrame");
    refs.modalInvoicePreview = document.getElementById("modalInvoicePreview");
    refs.closeModalButton = document.getElementById("closeModalButton");
  }

  function bindEvents() {
    refs.tabInvoicesButton.addEventListener("click", () => setActiveTab("invoices"));
    refs.tabLogsButton.addEventListener("click", () => setActiveTab("logs"));
    refs.searchInput.addEventListener("input", (event) => {
      state.searchQuery = String(event.target.value || "").trim().toLowerCase();
      state.searchServerCheckedQuery = "";
      state.searchServerMatchedCount = null;
      state.currentPage = 1;
      ensureSelectedInvoice();
      renderAll();
      scheduleSearchVerification();
    });
    bindFilterEvents();
    bindSortEvents();
    refs.paginationPrevButton.addEventListener("click", () => changePage(-1));
    refs.paginationNextButton.addEventListener("click", () => changePage(1));
    refs.downloadZipButton.addEventListener("click", onDownloadZip);
    refs.exportModeSelect.addEventListener("change", (event) => {
      state.exportMode = String(event.target.value || "one");
      renderExportInsight();
    });
    refs.clearLogsButton.addEventListener("click", onClearLogs);
    refs.closeModalButton.addEventListener("click", closePdfModal);
    refs.pdfModal.addEventListener("click", (event) => {
      if (event.target === refs.pdfModal) {
        closePdfModal();
      }
    });
    window.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        closePdfModal();
      }
    });
  }

  function bindFilterEvents() {
    const handleFilterChange = () => {
      state.filters.invoiceStatus = String(refs.filterInvoiceStatus.value || "");
      state.filters.invoiceNumber = String(refs.filterInvoiceNumber.value || "")
        .trim()
        .toLowerCase();
      state.filters.ownerName = String(refs.filterProjectManager.value || "");
      state.filters.amount = String(refs.filterAmount.value || "");
      state.filters.accountName = String(refs.filterAccount.value || "");
      state.filters.issueDate = String(refs.filterIssueDate.value || "");
      state.filters.dueDate = String(refs.filterDueDate.value || "");
      state.currentPage = 1;
      ensureSelectedInvoice();
      renderAll();
    };
    refs.filterInvoiceStatus.addEventListener("change", handleFilterChange);
    refs.filterInvoiceNumber.addEventListener("input", handleFilterChange);
    refs.filterProjectManager.addEventListener("change", handleFilterChange);
    refs.filterAmount.addEventListener("change", handleFilterChange);
    refs.filterAccount.addEventListener("change", handleFilterChange);
    refs.filterIssueDate.addEventListener("change", handleFilterChange);
    refs.filterDueDate.addEventListener("change", handleFilterChange);
    refs.clearFiltersButton.addEventListener("click", clearFilters);
  }

  function bindSortEvents() {
    refs.sortButtons.forEach((button) => {
      button.addEventListener("click", () => {
        const sortKey = String(button.getAttribute("data-sort") || "").trim();
        if (!sortKey) {
          return;
        }
        if (state.sortBy === sortKey) {
          state.sortDirection = state.sortDirection === "asc" ? "desc" : "asc";
        } else {
          state.sortBy = sortKey;
          state.sortDirection = defaultSortDirectionForColumn(sortKey);
        }
        state.currentPage = 1;
        renderAll();
      });
    });
  }

  async function initializeRuntime() {
    const query = new URLSearchParams(window.location.search);
    const fallback = getContextFromQuery(query);
    const sdk = await waitForSdkBridge(SDK_WAIT_MS);

    if (!sdk && hasLegacyRocketlaneApp()) {
      try {
        return await initializeLegacyRuntime(fallback);
      } catch (error) {
        console.warn("Legacy Rocketlane runtime init failed.", error);
      }
    }

    if (!sdk || typeof sdk.init !== "function") {
      return {
        client: null,
        connected: false,
        context: fallback,
        rawUser: null,
        rawAccount: null,
        rawProject: null,
        error: null,
      };
    }

    try {
      const client = await sdk.init({});
      const [account, user, project] = await Promise.all([
        safeGetClientData(client, "account"),
        safeGetClientData(client, "user"),
        safeGetClientData(client, "project"),
      ]);

      return {
        client,
        connected: true,
        context: mergeContextData(fallback, account, user, project),
        rawUser: user,
        rawAccount: account,
        rawProject: project,
        error: null,
      };
    } catch (error) {
      appendLog(
        "RUNTIME_INIT_FAILED",
        "Rocketlane SDK initialization failed. Falling back to local mode.",
        error
      );
      return {
        client: null,
        connected: false,
        context: fallback,
        rawUser: null,
        rawAccount: null,
        rawProject: null,
        error,
      };
    }
  }

  function hasLegacyRocketlaneApp() {
    return Boolean(
      window.rocketlaneApp &&
        typeof window.rocketlaneApp.init === "function"
    );
  }

  async function initializeLegacyRuntime(fallback) {
    const app = await window.rocketlaneApp.init();
    const contextEnvelope =
      app && typeof app.context === "function" ? await app.context() : {};
    const account =
      contextEnvelope.account || contextEnvelope.currentAccount || null;
    const user = contextEnvelope.user || contextEnvelope.currentUser || null;
    const project =
      contextEnvelope.project || contextEnvelope.currentProject || null;

    return {
      client: app,
      connected: true,
      context: mergeContextData(fallback, account, user, project),
      rawUser: user,
      rawAccount: account,
      rawProject: project,
      error: null,
    };
  }

  function resolveSdkBridge() {
    if (window.rliSdk && typeof window.rliSdk.init === "function") {
      return window.rliSdk;
    }

    try {
      if (
        window.parent &&
        window.parent !== window &&
        window.parent.rliSdk &&
        typeof window.parent.rliSdk.init === "function"
      ) {
        return window.parent.rliSdk;
      }
    } catch (_error) {
      // Ignore cross-origin parent access issues.
    }

    return null;
  }

  function waitForSdkBridge(timeoutMs) {
    const existing = resolveSdkBridge();
    if (existing) {
      return Promise.resolve(existing);
    }

    return new Promise((resolve) => {
      const startedAt = Date.now();
      const timer = window.setInterval(() => {
        const sdk = resolveSdkBridge();
        if (sdk) {
          window.clearInterval(timer);
          resolve(sdk);
          return;
        }
        if (Date.now() - startedAt >= timeoutMs) {
          window.clearInterval(timer);
          resolve(null);
        }
      }, SDK_POLL_INTERVAL_MS);
    });
  }

  async function safeGetClientData(client, objectName) {
    if (!client || !client.data || typeof client.data.get !== "function") {
      return null;
    }

    const attempted = new Set();
    const fetchByIdentifier = async (identifier) => {
      const key = String(identifier || "").trim();
      if (!key || attempted.has(key)) {
        return null;
      }
      attempted.add(key);
      try {
        const value = await client.data.get(key);
        const normalized = normalizeRuntimeEntity(objectName, value);
        return normalized != null ? normalized : null;
      } catch (_error) {
        return null;
      }
    };

    try {
      const direct = await client.data.get(objectName);
      const normalized = normalizeRuntimeEntity(objectName, direct);
      if (normalized != null) {
        return normalized;
      }
    } catch (_error) {
      // Continue to broader lookup candidates.
    }

    const directAliases = {
      account: ["account", "current_account", "currentAccount", "workspace"],
      user: ["user", "current_user", "currentUser", "viewer"],
      project: ["project", "current_project", "currentProject"],
    };
    const aliasKeys = {
      account: ["GET_ACCOUNT_DATA", "CURRENT_ACCOUNT", "ACCOUNT"],
      user: ["GET_USER_DATA", "CURRENT_USER", "USER", "ACCOUNT_USER"],
      project: ["GET_PROJECT_DATA", "CURRENT_PROJECT", "PROJECT"],
    };
    const matcherMap = {
      account: /(ACCOUNT|WORKSPACE)/i,
      user: /(USER|MEMBER|VIEWER)/i,
      project: /(PROJECT)/i,
    };

    const directList = directAliases[objectName] || [];
    for (let i = 0; i < directList.length; i += 1) {
      const found = await fetchByIdentifier(directList[i]);
      if (found) {
        return found;
      }
    }

    const identifiers = client.data.dataIdentifiers || {};
    const primaryAliasKeys = aliasKeys[objectName] || [];
    for (let i = 0; i < primaryAliasKeys.length; i += 1) {
      const found = await fetchByIdentifier(identifiers[primaryAliasKeys[i]]);
      if (found) {
        return found;
      }
    }

    const matcher = matcherMap[objectName];
    if (matcher) {
      const keys = Object.keys(identifiers);
      for (let i = 0; i < keys.length; i += 1) {
        if (!matcher.test(keys[i])) {
          continue;
        }
        const found = await fetchByIdentifier(identifiers[keys[i]]);
        if (found) {
          return found;
        }
      }
    }

    return null;
  }

  function normalizeRuntimeEntity(objectName, rawValue) {
    if (rawValue == null) {
      return null;
    }

    if (objectName === "user") {
      return normalizeUserPayload(rawValue);
    }
    if (objectName === "account") {
      return normalizeAccountPayload(rawValue);
    }
    if (objectName === "project") {
      return normalizeProjectPayload(rawValue);
    }

    return rawValue;
  }

  function normalizeUserPayload(rawValue) {
    const candidate = unwrapEntityObject(rawValue, [
      "user",
      "currentUser",
      "viewer",
      "accountUser",
      "member",
      "me",
    ]);
    if (!candidate) {
      return null;
    }
    if (
      Array.isArray(candidate.users) ||
      Array.isArray(candidate.members) ||
      Array.isArray(candidate.teamMembers) ||
      Array.isArray(candidate.accountUsers)
    ) {
      const fromCollection = selectCurrentUserFromCollection(
        candidate.users ||
          candidate.members ||
          candidate.teamMembers ||
          candidate.accountUsers
      );
      return fromCollection || null;
    }
    return candidate;
  }

  function normalizeAccountPayload(rawValue) {
    const candidate = unwrapEntityObject(rawValue, [
      "account",
      "currentAccount",
      "workspace",
      "organization",
    ]);
    if (!candidate) {
      return null;
    }
    if (Array.isArray(candidate.accounts) || Array.isArray(candidate.workspaces)) {
      return null;
    }
    return candidate;
  }

  function normalizeProjectPayload(rawValue) {
    const candidate = unwrapEntityObject(rawValue, [
      "project",
      "currentProject",
      "engagement",
    ]);
    if (!candidate) {
      return null;
    }
    if (Array.isArray(candidate.projects)) {
      return null;
    }
    return candidate;
  }

  function unwrapEntityObject(rawValue, preferredKeys) {
    const direct = asPlainObject(rawValue);
    if (!direct) {
      const fromArray = selectCurrentUserFromCollection(rawValue);
      return fromArray ? asPlainObject(fromArray) : null;
    }

    const candidates = [
      direct.response,
      direct.data,
      direct.payload,
      direct.result,
      direct.item,
      direct,
    ];

    for (let i = 0; i < candidates.length; i += 1) {
      const candidate = candidates[i];
      const asObject = asPlainObject(candidate);
      if (asObject) {
        const nested = pickNestedEntity(asObject, preferredKeys);
        if (nested) {
          return nested;
        }
        if (asObject !== direct) {
          return asObject;
        }
        continue;
      }

      if (Array.isArray(candidate)) {
        const fromData = selectCurrentUserFromCollection(candidate);
        if (fromData) {
          return fromData;
        }
      }
    }

    return direct;
  }

  function pickNestedEntity(payload, preferredKeys) {
    for (let i = 0; i < preferredKeys.length; i += 1) {
      const value = asPlainObject(payload[preferredKeys[i]]);
      if (value) {
        return value;
      }
    }
    return null;
  }

  function asPlainObject(value) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return null;
    }
    return value;
  }

  function selectCurrentUserFromCollection(collection) {
    if (!Array.isArray(collection)) {
      return null;
    }
    for (let i = 0; i < collection.length; i += 1) {
      const candidate = collection[i];
      if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) {
        continue;
      }
      if (
        candidate.isCurrentUser === true ||
        candidate.current === true ||
        candidate.me === true ||
        candidate.self === true
      ) {
        return candidate;
      }
    }
    return null;
  }

  function getContextFromQuery(query) {
    return {
      accountId: query.get("accountId") || "",
      accountName: query.get("account") || "Rocketlane Workspace",
      userId: query.get("userId") || "",
      userName: query.get("user") || "Rocketlane User",
      userEmail: normalizeEmail(query.get("email") || ""),
      userRole: query.get("role") || "",
      projectId: query.get("projectId") || "",
      projectName: query.get("project") || "",
    };
  }

  function mergeContextData(fallback, account, user, project) {
    const accountView = unwrapTopLevelObject(account);
    const userView = unwrapTopLevelObject(user);
    const projectView = unwrapTopLevelObject(project);

    return {
      accountId:
        pickFirst(
          accountView &&
            (accountView.id ||
              accountView.accountId ||
              accountView._id ||
              accountView.companyId)
        ) ||
        fallback.accountId,
      accountName:
        pickFirst(
          accountView &&
            (accountView.name ||
              accountView.accountName ||
              accountView.displayName ||
              accountView.companyName)
        ) || fallback.accountName,
      userId:
        pickFirst(
          userView &&
            (userView.id || userView.userId || userView._id || userView.userID)
        ) || fallback.userId,
      userName:
        pickFirst(
          userView &&
            (userView.name ||
              userView.fullName ||
              userView.displayName ||
              userView.userName ||
              userView.email ||
              [
                pickFirst(userView.firstName),
                pickFirst(userView.lastName),
              ]
                .filter(Boolean)
                .join(" "))
        ) || fallback.userName,
      userEmail:
        normalizeEmail(
          pickFirst(
            userView &&
              (userView.email ||
                userView.workEmail ||
                userView.userEmail ||
                userView.userName ||
                (userView.profile && userView.profile.email))
          ) || fallback.userEmail
        ),
      userRole:
        extractRoleLabel(userView) ||
        extractPermissionLabel(userView) ||
        pickFirst(userView && userView.userType) ||
        fallback.userRole,
      projectId:
        pickFirst(
          projectView &&
            (projectView.id || projectView.projectId || projectView._id)
        ) ||
        fallback.projectId,
      projectName:
        pickFirst(
          projectView &&
            (projectView.name ||
              projectView.projectName ||
              projectView.engagementName)
        ) ||
        fallback.projectName,
    };
  }

  async function fetchPermissionHintFromSdk() {
    if (!state.connected || !state.client || !state.client.data) {
      return null;
    }

    const direct = normalizeTeamMemberFromAny(state.rawUser);
    if (direct && hasPermissionSignals(direct)) {
      return direct;
    }

    const permissionPayloadHint = await fetchPermissionMetadataFromSdk(state.client);
    if (permissionPayloadHint) {
      return mergeObjects(
        {
          id: pickFirst((state.rawUser && (state.rawUser.id || state.rawUser.userId || state.rawUser._id)) || state.context.userId),
          email: normalizeEmail(extractPrimaryEmail(state.rawUser) || state.context.userEmail || ""),
          permission: "",
          roleLabel: "",
        },
        permissionPayloadHint
      );
    }

    try {
      state.teamMembers = await fetchTeamMembersFromSdk(state.client);
    } catch (error) {
      appendLog(
        "PERMISSION_FETCH_FAILED",
        "Unable to fetch team members via SDK data identifiers.",
        error
      );
      state.teamMembers = [];
    }

    return resolveCurrentUserPermission(
      state.teamMembers,
      state.rawUser,
      state.context
    );
  }

  async function fetchPermissionMetadataFromSdk(client) {
    if (!client || !client.data) {
      return null;
    }

    const parsePayload = (payload) => {
      if (!payload || typeof payload !== "object") {
        return null;
      }
      const permission = extractPermissionLabel(payload);
      const roleLabel = extractRoleLabel(payload);
      if (!permission && !roleLabel) {
        return null;
      }
      return {
        permission: permission || "",
        roleLabel: roleLabel || permission || "",
      };
    };

    const directCandidates = [
      "permission",
      "permissions",
      "permissionSet",
      "permission_set",
      "role",
      "roles",
      "currentUserRole",
      "currentUserPermission",
    ];
    for (let i = 0; i < directCandidates.length; i += 1) {
      const payload = await invokeSdkDataGet(client, directCandidates[i]);
      const parsed = parsePayload(payload);
      if (parsed) {
        return parsed;
      }
    }

    const identifiers = (client.data && client.data.dataIdentifiers) || {};
    const keys = Object.keys(identifiers).filter((key) => {
      const upper = key.toUpperCase();
      return (
        upper.includes("PERMISSION") ||
        upper.includes("ROLE") ||
        upper.includes("ACCESS")
      );
    });
    for (let i = 0; i < keys.length; i += 1) {
      const payload = await invokeSdkDataGet(client, identifiers[keys[i]]);
      const parsed = parsePayload(payload);
      if (parsed) {
        return parsed;
      }
    }

    return null;
  }

  async function fetchTeamMembersFromSdk(client) {
    const members = [];
    const seen = new Set();

    const addMembersFromPayload = (payload) => {
      const rows = extractCollection(payload, [
        "users",
        "members",
        "teamMembers",
        "accountUsers",
        "account_users",
        "data",
        "items",
        "results",
      ]);
      rows.forEach((row) => {
        const normalized = normalizeTeamMemberFromAny(row);
        if (!normalized) {
          return;
        }
        const key = normalized.email || normalized.id;
        if (!key || seen.has(key)) {
          return;
        }
        seen.add(key);
        members.push(normalized);
      });
    };

    const directCandidates = [
      "users",
      "members",
      "teamMembers",
      "accountUsers",
      "account_users",
      "allUsers",
      "all_users",
    ];
    for (let i = 0; i < directCandidates.length; i += 1) {
      const payload = await invokeSdkDataGet(client, directCandidates[i]);
      if (payload) {
        addMembersFromPayload(payload);
      }
    }

    const identifiers = (client.data && client.data.dataIdentifiers) || {};
    const identifierKeys = Object.keys(identifiers).filter((key) => {
      const upper = key.toUpperCase();
      return (
        (upper.includes("USER") || upper.includes("MEMBER") || upper.includes("TEAM")) &&
        !upper.includes("CURRENT_PROJECT") &&
        !upper.includes("GET_PROJECT")
      );
    });

    for (let i = 0; i < identifierKeys.length; i += 1) {
      const key = identifiers[identifierKeys[i]];
      const payload = await invokeSdkDataGet(client, key);
      if (payload) {
        addMembersFromPayload(payload);
      }
    }

    return members;
  }

  async function invokeSdkDataGet(client, key) {
    const value = String(key || "").trim();
    if (!value || !client || !client.data || typeof client.data.get !== "function") {
      return null;
    }
    try {
      return await client.data.get(value);
    } catch (_error) {
      return null;
    }
  }

  function normalizeTeamMemberFromAny(raw) {
    if (!raw || typeof raw !== "object") {
      return null;
    }

    const id = pickFirst(raw.id || raw.userId || raw._id);
    const email = normalizeEmail(
      pickFirst(
        raw.email ||
          raw.emailId ||
          raw.userEmail ||
          raw.workEmail ||
          (raw.user && (raw.user.email || raw.user.emailId)) ||
          (raw.user && raw.user.email) ||
          (raw.profile && raw.profile.email)
      )
    );
    if (!email && !id) {
      return null;
    }

    const permission = extractPermissionLabel(raw);
    const roleLabel = extractRoleLabel(raw);

    return {
      id,
      email,
      permission,
      roleLabel,
      isAdmin: Boolean(
        raw.isAdmin === true ||
          raw.admin === true ||
          raw.isAccountAdmin === true ||
          raw.accountAdmin === true
      ),
    };
  }

  function resolveCurrentUserPermission(members, rawUser, context) {
    if (!Array.isArray(members) || !members.length) {
      return null;
    }

    const targetEmail = normalizeEmail(
      extractPrimaryEmail(rawUser) || context.userEmail || ""
    );
    const targetId = pickFirst(
      (rawUser && (rawUser.id || rawUser.userId || rawUser._id)) || context.userId
    );

    if (targetEmail) {
      const byEmail = members.find((member) => member.email === targetEmail);
      if (byEmail) {
        return byEmail;
      }
    }

    if (targetId) {
      const byId = members.find((member) => member.id && member.id === targetId);
      if (byId) {
        return byId;
      }
    }

    return null;
  }

  async function refreshInvoicesFromSource() {
    state.syncStatus = "Fetching invoices from source projects...";
    renderSyncStatus();

    let invoices = [];
    state.syncDiagnostics = mergeObjects(state.syncDiagnostics, {
      lastRefreshAt: new Date().toISOString(),
      usedServerAction: false,
      serverActionAttempted: false,
      usedSdkFallback: false,
      serverActionError: "",
    });
    try {
      if (state.connected) {
        const serverInvoices = await fetchInvoicesFromServerAction();
        state.syncDiagnostics.usedServerAction = true;
        if (serverInvoices && serverInvoices.length) {
          invoices = serverInvoices;
        } else {
          state.syncDiagnostics.usedSdkFallback = true;
          invoices = await fetchInvoicesFromSourceProjects();
        }
      }
    } catch (error) {
      appendLog(
        "SOURCE_FETCH_FAILED",
        "Unable to fetch invoices from source projects.",
        error
      );
    }

    if (!invoices.length) {
      state.syncStatus =
        "No invoices found in source project(s) yet. Add approved invoices in Expert Advisor Program Invoices and refresh.";
      appendLog(
        "SOURCE_INVOICES_NOT_FOUND",
        "No invoices were discovered from source project data."
      );
    } else {
      state.syncStatus =
        "Loaded " +
        invoices.length +
        " invoice(s) from source project(s): " +
        state.sourceProjects.join(", ");
    }

    state.invoices = invoices.map(normalizeInvoice).filter(Boolean);
    state.currentPage = 1;
  }

  async function fetchInvoicesFromServerAction() {
    if (
      !state.client ||
      !state.client.data ||
      typeof state.client.data.invoke !== "function"
    ) {
      return [];
    }

    try {
      state.syncDiagnostics = mergeObjects(state.syncDiagnostics, {
        serverActionAttempted: true,
      });
      const workspaceCandidates = [
        "https://blink.rocketlane.com",
        "https://innovate-calgary.rocketlane.com",
      ];
      const accountDomain = pickFirst(
        state.rawAccount &&
          (state.rawAccount.domainName ||
            state.rawAccount.primaryDomainName ||
            state.rawAccount.domain)
      );
      if (accountDomain) {
        workspaceCandidates.unshift("https://" + accountDomain.replace(/^https?:\/\//i, ""));
      }
      const payload = await state.client.data.invoke("syncInvoicesFromSource", {
        sourceProjectNames: SOURCE_PROJECT_NAMES.slice(),
        accountName: state.context.accountName || "",
        workspaceBaseUrl: workspaceCandidates[0],
        workspaceCandidates,
        viewerContext: {
          userId: state.context.userId || "",
          userEmail: state.context.userEmail || "",
          userRole: state.context.userRole || "",
          userName: state.context.userName || "",
        },
      });
      const result = unwrapServerActionResponse(payload);
      if (!result || result.ok === false) {
        if (result && result.error) {
          appendLog("SOURCE_FETCH_FAILED", result.error);
          state.syncDiagnostics = mergeObjects(state.syncDiagnostics, {
            serverActionError: String(result.error || ""),
          });
        }
        if (result && result.diagnostics) {
          state.syncDiagnostics = mergeObjects(state.syncDiagnostics, result.diagnostics);
        }
        return [];
      }

      if (Array.isArray(result.sourceProjects) && result.sourceProjects.length) {
        state.sourceProjects = result.sourceProjects
          .map((project) => normalizeProjectRecord(project))
          .filter(Boolean)
          .map((project) => project.name);
      }

      if (Array.isArray(result.teamMembers) && result.teamMembers.length) {
        state.teamMembers = result.teamMembers
          .map((member) => normalizeTeamMemberFromAny(member))
          .filter(Boolean);
        const permissionHint = resolveCurrentUserPermission(
          state.teamMembers,
          state.rawUser,
          state.context
        );
        if (permissionHint && hasPermissionSignals(permissionHint)) {
          const nextAccess = deriveAccessProfile(
            state.rawUser,
            state.rawAccount,
            state.context,
            permissionHint
          );
          state.permissionHint = permissionHint;
          // Never demote an already-detected admin using a weaker hint payload.
          if (!state.access.isAdmin || nextAccess.isAdmin) {
            state.access = nextAccess;
          }
          updateHeader();
          configureUiForAccess();
        }
      }

      if (result.viewer && typeof result.viewer === "object") {
        const viewerHint = normalizeTeamMemberFromAny(result.viewer) || {};
        if (result.viewer.isAdmin === true) {
          viewerHint.isAdmin = true;
        }
        if (!viewerHint.permission) {
          viewerHint.permission = pickFirst(result.viewer.permission);
        }
        if (!viewerHint.roleLabel) {
          viewerHint.roleLabel = pickFirst(result.viewer.roleLabel);
        }
        if (!viewerHint.email) {
          viewerHint.email = normalizeEmail(pickFirst(result.viewer.email || result.viewer.emailId));
        }
        if (!viewerHint.id) {
          viewerHint.id = pickFirst(result.viewer.id || result.viewer.userId || result.viewer._id);
        }
        if (hasPermissionSignals(viewerHint)) {
          state.permissionHint = mergeObjects(state.permissionHint || {}, viewerHint);
          const nextAccess = deriveAccessProfile(
            state.rawUser,
            state.rawAccount,
            state.context,
            state.permissionHint
          );
          if (!state.access.isAdmin || nextAccess.isAdmin) {
            state.access = nextAccess;
          }
          updateHeader();
          configureUiForAccess();
        }
      }

      state.syncDiagnostics = mergeObjects(state.syncDiagnostics, result.diagnostics || {});
      return Array.isArray(result.invoices) ? result.invoices : [];
    } catch (error) {
      appendLog(
        "SOURCE_FETCH_FAILED",
        "Server action invoice sync failed; falling back to SDK-only discovery.",
        error
      );
      state.syncDiagnostics = mergeObjects(state.syncDiagnostics, {
        serverActionError: simplifyError(error),
      });
      return [];
    }
  }

  function unwrapServerActionResponse(payload) {
    let current = payload;
    for (let i = 0; i < 6; i += 1) {
      if (!current) {
        return null;
      }
      if (Array.isArray(current)) {
        return { ok: true, invoices: current };
      }
      if (typeof current !== "object") {
        return null;
      }
      if (
        current.ok !== undefined ||
        current.error ||
        current.invoices ||
        current.sourceProjects ||
        current.teamMembers
      ) {
        return current;
      }
      current =
        current.data ||
        current.response ||
        current.result ||
        current.payload ||
        current.body ||
        null;
    }
    return null;
  }

  async function fetchInvoicesFromSourceProjects() {
    const projects = await fetchSourceProjects();
    state.sourceProjects = projects.map((item) => item.name);

    const invoices = [];
    if (projects.length) {
      for (let i = 0; i < projects.length; i += 1) {
        const perProject = await fetchInvoicesForProject(projects[i]);
        invoices.push(...perProject);
      }
    }

    if (!invoices.length) {
      const sdkFallbackInvoices = await fetchInvoicesByScanningSdkArtifacts();
      if (sdkFallbackInvoices.length) {
        return dedupeInvoices(sdkFallbackInvoices);
      }
    }

    return dedupeInvoices(invoices);
  }

  async function fetchSourceProjects() {
    const projects = [];
    const byKey = new Set();

    const runtimeProject = normalizeProjectRecord(state.rawProject);
    if (
      runtimeProject &&
      isSourceProjectName(runtimeProject.name)
    ) {
      const key = runtimeProject.id || runtimeProject.name.toLowerCase();
      byKey.add(key);
      projects.push(runtimeProject);
    }

    const sdkProjects = await fetchProjectsFromSdk();
    sdkProjects.forEach((record) => {
      const project = normalizeProjectRecord(record);
      if (!project) {
        return;
      }
      if (!isSourceProjectName(project.name)) {
        return;
      }
      const key = project.id || project.name.toLowerCase();
      if (byKey.has(key)) {
        return;
      }
      byKey.add(key);
      projects.push(project);
    });

    const payload = await requestCollection(
      [
        "/api/1.0/projects?size=200",
        "/api/1.0/projects?limit=200",
        "/api/1.0/projects",
      ],
      ["projects", "data", "content", "results", "items"]
    );

    payload.forEach((record) => {
      const project = normalizeProjectRecord(record);
      if (!project) {
        return;
      }
      if (!isSourceProjectName(project.name)) {
        return;
      }
      const key = project.id || project.name.toLowerCase();
      if (byKey.has(key)) {
        return;
      }
      byKey.add(key);
      projects.push(project);
    });

    return projects;
  }

  async function fetchInvoicesForProject(project) {
    const projectId = project.id;
    if (!projectId && !project.name) {
      return [];
    }

    const records = [];
    if (project.raw && typeof project.raw === "object") {
      records.push(project.raw);
    }
    const sdkArtifacts = await fetchProjectArtifactsFromSdk(project);
    records.push(...sdkArtifacts);

    const endpoints = [];
    if (projectId) {
      endpoints.push(
        "/api/1.0/projects/" + encodeURIComponent(projectId) + "/documents?size=200",
        "/api/1.0/projects/" + encodeURIComponent(projectId) + "/documents",
        "/api/1.0/projects/" + encodeURIComponent(projectId) + "/files?size=200",
        "/api/1.0/projects/" + encodeURIComponent(projectId) + "/files",
        "/api/1.0/documents?projectId=" + encodeURIComponent(projectId) + "&size=200",
        "/api/1.0/files?projectId=" + encodeURIComponent(projectId) + "&size=200",
        "/api/1.0/tasks?projectId=" + encodeURIComponent(projectId) + "&size=200"
      );
    }

    const apiRecords = await requestCollection(endpoints, [
      "documents",
      "files",
      "tasks",
      "data",
      "content",
      "results",
      "items",
    ]);
    records.push(...apiRecords);

    const invoices = [];
    for (let i = 0; i < records.length; i += 1) {
      const candidates = extractInvoiceCandidates(records[i]);
      for (let j = 0; j < candidates.length; j += 1) {
        const invoice = await buildInvoiceFromCandidate(candidates[j], project);
        if (invoice) {
          invoices.push(invoice);
        }
      }
    }

    if (!invoices.length) {
      appendLog(
        "SOURCE_INVOICES_NOT_FOUND",
        'No invoice records were found under source project "' +
          project.name +
          '".'
      );
    }

    return invoices;
  }

  async function fetchProjectsFromSdk() {
    if (!state.client || !state.client.data) {
      return [];
    }

    const records = [];
    const seen = new Set();
    const addRows = (payload) => {
      const rows = extractCollection(payload, [
        "projects",
        "data",
        "content",
        "results",
        "items",
      ]);
      if (
        !rows.length &&
        payload &&
        typeof payload === "object" &&
        !Array.isArray(payload)
      ) {
        rows.push(payload);
      }
      rows.forEach((row) => {
        const key = buildRowKey(row);
        if (seen.has(key)) {
          return;
        }
        seen.add(key);
        records.push(row);
      });
    };

    const directCandidates = ["projects", "projectList", "allProjects"];
    for (let i = 0; i < directCandidates.length; i += 1) {
      const payload = await invokeSdkDataGet(state.client, directCandidates[i]);
      if (payload) {
        addRows(payload);
      }
    }

    const identifiers = (state.client.data && state.client.data.dataIdentifiers) || {};
    const keys = Object.keys(identifiers).filter((key) => {
      const upper = key.toUpperCase();
      return (
        upper.includes("PROJECT") &&
        !upper.includes("CURRENT_PROJECT") &&
        (upper.includes("PROJECTS") ||
          upper.includes("ALL") ||
          upper.includes("LIST") ||
          upper.includes("SEARCH"))
      );
    });

    for (let i = 0; i < keys.length; i += 1) {
      const payload = await invokeSdkDataGet(state.client, identifiers[keys[i]]);
      if (payload) {
        addRows(payload);
      }
    }

    return records;
  }

  async function fetchProjectArtifactsFromSdk(project) {
    if (!state.client || !state.client.data) {
      return [];
    }

    const records = [];
    const seen = new Set();
    const attemptedKeys = [];
    const addRows = (payload) => {
      const rows = extractCollection(payload, [
        "documents",
        "files",
        "tasks",
        "invoices",
        "attachments",
        "data",
        "content",
        "results",
        "items",
      ]);
      if (
        !rows.length &&
        payload &&
        typeof payload === "object" &&
        !Array.isArray(payload)
      ) {
        rows.push(payload);
      }
      rows.forEach((row) => {
        if (project && !matchesProject(row, project)) {
          return;
        }
        const key = buildRowKey(row);
        if (seen.has(key)) {
          return;
        }
        seen.add(key);
        records.push(row);
      });
    };

    const directCandidates = [
      "documents",
      "files",
      "attachments",
      "tasks",
      "invoices",
      "invoiceDocuments",
      "projectInvoices",
      "projectArtifacts",
    ];
    for (let i = 0; i < directCandidates.length; i += 1) {
      attemptedKeys.push(directCandidates[i]);
      const payload = await invokeSdkDataGet(state.client, directCandidates[i]);
      if (payload) {
        addRows(payload);
      }
    }

    const identifiers = (state.client.data && state.client.data.dataIdentifiers) || {};
    const keys = Object.keys(identifiers).filter((key) =>
      isPotentialArtifactIdentifierKey(key)
    );

    for (let i = 0; i < keys.length; i += 1) {
      attemptedKeys.push(String(keys[i]));
      const payload = await invokeSdkDataGet(state.client, identifiers[keys[i]]);
      if (payload) {
        addRows(payload);
      }
    }

    if (!records.length) {
      const broadKeys = Object.keys(identifiers).filter(
        (key) => !isExcludedFromBroadArtifactScan(key)
      );
      for (let i = 0; i < broadKeys.length; i += 1) {
        attemptedKeys.push(String(broadKeys[i]));
        const payload = await invokeSdkDataGet(state.client, identifiers[broadKeys[i]]);
        if (!payload || !payloadLikelyContainsArtifacts(payload)) {
          continue;
        }
        addRows(payload);
      }
    }

    state.syncDiagnostics = mergeObjects(state.syncDiagnostics, {
      attemptedKeys: dedupeStrings(attemptedKeys),
      artifactRecordsFound: records.length,
      sourceProjectsFound: state.sourceProjects.slice(),
      dataIdentifiersCount: Object.keys(identifiers || {}).length,
    });

    return records;
  }

  async function fetchInvoicesByScanningSdkArtifacts() {
    const records = await fetchProjectArtifactsFromSdk(null);
    if (!records.length) {
      return [];
    }

    const invoices = [];
    const sourceProjectNames = new Set();
    for (let i = 0; i < records.length; i += 1) {
      const project = resolveProjectFromArtifact(records[i]);
      const projectFromText = detectSourceProjectNameInValue(records[i]);
      const resolvedProject =
        project ||
        (projectFromText
          ? createSyntheticSourceProject(projectFromText, records[i])
          : null);
      if (!resolvedProject) {
        continue;
      }
      if (!isSourceProjectName(resolvedProject.name)) {
        continue;
      }
      sourceProjectNames.add(resolvedProject.name);
      const candidates = extractInvoiceCandidates(records[i]);
      for (let j = 0; j < candidates.length; j += 1) {
        const invoice = await buildInvoiceFromCandidate(candidates[j], resolvedProject);
        if (invoice) {
          invoices.push(invoice);
        }
      }
    }

    if (sourceProjectNames.size) {
      state.sourceProjects = Array.from(sourceProjectNames);
    }
    return dedupeInvoices(invoices);
  }

  function resolveProjectFromArtifact(record) {
    if (!record || typeof record !== "object") {
      return null;
    }

    const nested =
      normalizeProjectRecord(record.project) ||
      normalizeProjectRecord(record.parentProject) ||
      normalizeProjectRecord(record.engagement);
    if (nested) {
      return nested;
    }

    const detectedSourceName = detectSourceProjectNameInValue(record);
    const name =
      pickFirst(
      record.projectName ||
        record.project_name ||
        record.projectTitle ||
        record.projectDisplayName ||
        record.engagementName ||
        record.parentProjectName ||
        (record.meta && (record.meta.projectName || record.meta.project))
    ) || detectedSourceName;
    if (!name) {
      return null;
    }

    const id = pickFirst(
      record.projectId ||
        record.project_id ||
        record.parentProjectId ||
        record.engagementId
    );
    const accountName =
      pickFirst(
        record.accountName ||
          record.companyName ||
          (record.account && record.account.name) ||
          (record.customer && record.customer.name)
      ) || state.context.accountName;
    const contacts = extractContacts(record);

    return {
      id,
      name,
      accountName,
      ownerName: contacts.names[0] || "Unassigned",
      ownerEmails: contacts.emails,
      raw: record,
    };
  }

  function createSyntheticSourceProject(name, record) {
    const contacts = extractContacts(record);
    return {
      id: "",
      name: String(name || SOURCE_PROJECT_NAMES[0] || "").trim(),
      accountName: state.context.accountName,
      ownerName: contacts.names[0] || "Unassigned",
      ownerEmails: contacts.emails,
      raw: record,
    };
  }

  function isSourceProjectName(name) {
    const normalized = canonicalProjectName(name);
    if (!normalized) {
      return false;
    }
    return SOURCE_PROJECT_NAMES.some((candidate) => {
      const target = canonicalProjectName(candidate);
      return (
        normalized === target ||
        normalized.includes(target) ||
        target.includes(normalized)
      );
    });
  }

  function normalizeProjectName(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function canonicalProjectName(value) {
    return normalizeProjectName(value)
      .split(/\s+/)
      .filter(Boolean)
      .map((token) => (token.length > 3 && token.endsWith("s") ? token.slice(0, -1) : token))
      .join(" ");
  }

  function detectSourceProjectNameInValue(value) {
    let serialized = "";
    try {
      serialized = JSON.stringify(value || {});
    } catch (_error) {
      serialized = String(value || "");
    }
    const haystack = normalizeProjectName(serialized);
    if (!haystack) {
      return "";
    }
    for (let i = 0; i < SOURCE_PROJECT_NAMES.length; i += 1) {
      const candidate = SOURCE_PROJECT_NAMES[i];
      const normalized = normalizeProjectName(candidate);
      if (normalized && haystack.includes(normalized)) {
        return candidate;
      }
    }
    return "";
  }

  function isPotentialArtifactIdentifierKey(key) {
    const upper = String(key || "").toUpperCase();
    if (!upper || upper.includes("CURRENT_USER")) {
      return false;
    }
    return (
      upper.includes("DOCUMENT") ||
      upper.includes("FILE") ||
      upper.includes("ATTACHMENT") ||
      upper.includes("INVOICE") ||
      upper.includes("TASK") ||
      upper.includes("ARTIFACT") ||
      upper.includes("BILL") ||
      upper.includes("RECEIPT")
    );
  }

  function isExcludedFromBroadArtifactScan(key) {
    const upper = String(key || "").toUpperCase();
    return (
      upper.includes("CURRENT_USER") ||
      upper.includes("CURRENT_ACCOUNT") ||
      upper.includes("CURRENT_PROJECT") ||
      upper.includes("GET_USER_DATA") ||
      upper.includes("GET_ACCOUNT_DATA") ||
      upper.includes("GET_PROJECT_DATA") ||
      upper.includes("USER") ||
      upper.includes("ACCOUNT") ||
      upper.includes("PERMISSION") ||
      upper.includes("ROLE") ||
      upper.includes("TEAM") ||
      upper.includes("MEMBER")
    );
  }

  function payloadLikelyContainsArtifacts(payload) {
    if (!payload) {
      return false;
    }
    const rows = extractCollection(payload, [
      "documents",
      "files",
      "invoices",
      "attachments",
      "tasks",
      "items",
      "results",
      "data",
    ]);
    if (!rows.length) {
      return false;
    }
    return rows.some((row) => looksLikeInvoiceRecordNode(row) || extractInvoiceCandidates(row).length > 0);
  }

  function extractAssociatedEmails(node) {
    if (!node || typeof node !== "object") {
      return [];
    }
    const direct = [
      node.projectManagerEmail,
      node.pmEmail,
      node.expertAdvisorEmail,
      node.ownerEmail,
      node.assigneeEmail,
      node.createdByEmail,
      node.submittedByEmail,
      node.approvedByEmail,
      node.userEmail,
      node.email,
    ].map((value) => normalizeEmail(value));
    return dedupeEmails(direct);
  }

  function extractAssociatedUserIds(node) {
    if (!node || typeof node !== "object") {
      return [];
    }
    const ids = [];
    const push = (value) => {
      const text = pickFirst(value);
      if (text) {
        ids.push(text);
      }
    };
    push(node.userId);
    push(node.userID);
    push(node.ownerId);
    push(node.assigneeId);
    push(node.projectManagerId);
    push(node.expertAdvisorId);
    push(node.createdByUserId);
    push(node.submittedByUserId);
    push(node.approvedByUserId);
    if (node.createdBy && typeof node.createdBy === "object") {
      push(node.createdBy.id || node.createdBy.userId || node.createdBy.userID);
    }
    if (node.user && typeof node.user === "object") {
      push(node.user.id || node.user.userId || node.user.userID);
    }
    return dedupeStrings(ids);
  }

  async function requestCollection(endpoints, preferredKeys) {
    if (!canUseDirectApiFetch()) {
      return [];
    }

    const collected = [];
    const seen = new Set();

    for (let i = 0; i < endpoints.length; i += 1) {
      try {
        const payload = await requestJson(endpoints[i]);
        const rows = extractCollection(payload, preferredKeys);
        rows.forEach((row) => {
          const key = buildRowKey(row);
          if (seen.has(key)) {
            return;
          }
          seen.add(key);
          collected.push(row);
        });
      } catch (_error) {
        // Continue with next endpoint candidate.
      }
    }

    return collected;
  }

  function canUseDirectApiFetch() {
    const origin = String((window.location && window.location.origin) || "").toLowerCase();
    if (!origin) {
      return false;
    }
    if (origin.includes("amazonaws.com") || origin.includes("cloudfront.net")) {
      return false;
    }
    return origin.includes("rocketlane.com") || origin.includes("localhost");
  }

  function matchesProject(record, project) {
    if (!record || typeof record !== "object" || !project) {
      return false;
    }

    const targetId = String(project.id || "").toLowerCase();
    const targetName = String(project.name || "").toLowerCase();
    const recordProjectId = pickFirst(
      record.projectId ||
        record.project_id ||
        record.parentProjectId ||
        (record.project && (record.project.id || record.project.projectId))
    );
    if (targetId && recordProjectId && String(recordProjectId).toLowerCase() === targetId) {
      return true;
    }

    const recordProjectName = pickFirst(
      record.projectName ||
        record.project_name ||
        (record.project && (record.project.name || record.project.projectName))
    );
    if (targetName && recordProjectName) {
      return String(recordProjectName).toLowerCase() === targetName;
    }

    return false;
  }

  async function requestJson(path) {
    const response = await fetch(path, {
      method: "GET",
      credentials: "include",
      headers: { Accept: "application/json" },
    });

    if (!response.ok) {
      throw new Error("Request failed (" + response.status + ") for " + path);
    }

    const text = await response.text();
    if (!text) {
      return null;
    }

    try {
      return JSON.parse(text);
    } catch (_error) {
      throw new Error("Expected JSON payload from " + path);
    }
  }

  function extractCollection(payload, preferredKeys) {
    if (!payload) {
      return [];
    }
    if (Array.isArray(payload)) {
      return payload;
    }

    const preferred = [];
    const pushIfArray = (value) => {
      if (Array.isArray(value)) {
        preferred.push(value);
      }
    };

    preferredKeys.forEach((key) => {
      if (payload && typeof payload === "object") {
        pushIfArray(payload[key]);
        if (payload.data && typeof payload.data === "object") {
          pushIfArray(payload.data[key]);
        }
      }
    });

    if (preferred.length) {
      return preferred.sort((a, b) => b.length - a.length)[0];
    }

    const discovered = [];
    collectArrayNodes(payload, discovered, 0, 3);
    if (!discovered.length) {
      return [];
    }
    return discovered.sort((a, b) => b.length - a.length)[0];
  }

  function collectArrayNodes(value, output, depth, maxDepth) {
    if (depth > maxDepth || value == null) {
      return;
    }

    if (Array.isArray(value)) {
      if (value.some((entry) => entry && typeof entry === "object")) {
        output.push(value);
      }
      value.forEach((entry) => collectArrayNodes(entry, output, depth + 1, maxDepth));
      return;
    }

    if (typeof value !== "object") {
      return;
    }

    Object.keys(value).forEach((key) => {
      collectArrayNodes(value[key], output, depth + 1, maxDepth);
    });
  }

  function normalizeProjectRecord(raw) {
    if (!raw || typeof raw !== "object") {
      return null;
    }

    const id = pickFirst(raw.id || raw._id || raw.projectId);
    const name = pickFirst(raw.name || raw.projectName || raw.title);
    if (!name) {
      return null;
    }

    const accountName =
      pickFirst(
        raw.accountName ||
          (raw.account && raw.account.name) ||
          (raw.customer && raw.customer.name) ||
          raw.companyName
      ) || state.context.accountName;

    const contacts = extractContacts(raw);

    return {
      id,
      name,
      accountName,
      ownerName: contacts.names[0] || "Unassigned",
      ownerEmails: contacts.emails,
      raw,
    };
  }

  function extractPdfCandidates(record) {
    const candidates = [];
    const visited = new Set();

    function walk(value, depth) {
      if (depth > 5 || value == null) {
        return;
      }

      if (typeof value !== "object") {
        return;
      }

      if (Array.isArray(value)) {
        value.forEach((entry) => walk(entry, depth + 1));
        return;
      }

      if (visited.has(value)) {
        return;
      }
      visited.add(value);

      if (looksLikePdfNode(value)) {
        candidates.push(value);
      }

      Object.keys(value).forEach((key) => {
        walk(value[key], depth + 1);
      });
    }

    walk(record, 0);
    return candidates;
  }

  function extractInvoiceCandidates(record) {
    const pdfCandidates = extractPdfCandidates(record);
    const seen = new Set();
    const candidates = [];

    const addCandidate = (node) => {
      if (!node || typeof node !== "object" || Array.isArray(node)) {
        return;
      }
      const key = buildRowKey(node);
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      candidates.push(node);
    };

    pdfCandidates.forEach(addCandidate);

    const visited = new Set();
    function walk(value, depth) {
      if (depth > 5 || value == null) {
        return;
      }
      if (typeof value !== "object") {
        return;
      }
      if (Array.isArray(value)) {
        value.forEach((entry) => walk(entry, depth + 1));
        return;
      }
      if (visited.has(value)) {
        return;
      }
      visited.add(value);
      if (looksLikeInvoiceRecordNode(value)) {
        addCandidate(value);
      }
      Object.keys(value).forEach((key) => walk(value[key], depth + 1));
    }

    walk(record, 0);
    return candidates;
  }

  function looksLikeInvoiceRecordNode(node) {
    const invoiceNumber = pickFirst(
      node.invoiceNumber ||
        node.invoiceNo ||
        node.invoiceId ||
        node.billNumber ||
        node.referenceNumber
    );
    const status = String(node.status || node.invoiceStatus || node.state || "").toLowerCase();
    const category = String(node.category || node.kind || node.type || "").toLowerCase();
    const name = String(
      node.name || node.title || node.invoiceName || node.invoiceTitle || ""
    ).toLowerCase();
    const hasMoneySignal = Boolean(
      node.totalAmount || node.amount || node.netAmount || node.grossAmount || node.currencyCode
    );

    if (invoiceNumber) {
      return true;
    }
    if (name.includes("invoice") || category.includes("invoice")) {
      return true;
    }
    if (status.includes("approved") && hasMoneySignal) {
      return true;
    }
    return false;
  }

  function looksLikePdfNode(node) {
    const mime = String(
      node.mimeType || node.contentType || node.fileType || node.type || ""
    ).toLowerCase();
    const name = String(node.name || node.fileName || node.title || "").toLowerCase();
    const url = String(
      node.url ||
        node.fileUrl ||
        node.downloadUrl ||
        node.signedUrl ||
        node.href ||
        node.previewUrl ||
        node.attachmentUrl ||
        node.documentUrl ||
        (node.file && (node.file.url || node.file.downloadUrl || node.file.signedUrl)) ||
        ""
    ).toLowerCase();
    const category = String(
      node.category || node.kind || node.documentType || node.recordType || ""
    ).toLowerCase();
    const hasInvoiceMarker = Boolean(
      pickFirst(node.invoiceNumber || node.invoiceId || node.billNumber || node.referenceNumber)
    );
    const hasFileUrl = Boolean(url);

    if (mime.includes("pdf")) {
      return true;
    }
    if (name.endsWith(".pdf")) {
      return true;
    }
    if (url.includes(".pdf")) {
      return true;
    }
    if (
      hasFileUrl &&
      (name.includes("invoice") ||
        category.includes("invoice") ||
        category.includes("bill") ||
        category.includes("receipt") ||
        hasInvoiceMarker)
    ) {
      return true;
    }
    return false;
  }

  async function buildInvoiceFromCandidate(node, project) {
    const pdfUrlRaw = String(
      node.signedUrl ||
        node.downloadUrl ||
        node.fileUrl ||
        node.url ||
        node.href ||
        node.previewUrl ||
        node.attachmentUrl ||
        node.documentUrl ||
        (node.file && (node.file.signedUrl || node.file.downloadUrl || node.file.url)) ||
        ""
    ).trim();
    const pdfUrl = pdfUrlRaw ? toAbsoluteUrl(pdfUrlRaw) : "";
    const invoiceName =
      pickFirst(
        node.name ||
          node.fileName ||
          node.title ||
          node.subject ||
          node.invoiceName ||
          node.invoiceTitle
      ) || "Invoice";
    const invoiceNumber =
      pickFirst(
        node.invoiceNumber ||
          node.invoiceNo ||
          node.invoiceId ||
          node.number ||
          node.docNumber ||
          node.documentNumber ||
          node.billNumber ||
          node.referenceNumber ||
          extractInvoiceNumberFromText(invoiceName)
      ) || "INV-" + createShortId();
    const invoiceDate =
      pickFirst(
        node.invoiceDate ||
          node.dateOfIssue ||
          node.date ||
          node.issuedOn ||
          node.issuedDate ||
          node.approvedAt ||
          node.submittedAt ||
          node.createdAt ||
          node.updatedAt
      ) || new Date().toISOString();
    const dueDate =
      pickFirst(node.dueDate || node.dueOn || node.paymentDueDate || node.paymentDueOn) || "";
    const invoiceStatus = pickFirst(node.status || node.invoiceStatus || node.state || "Unknown");
    const amount = Number(
      pickFirst(node.amount || node.totalAmount || node.netAmount || node.grossAmount || node.subTotal) || 0
    );
    const currencyCode = pickFirst(node.currencyCode || (node.currency && node.currency.currencyCode));
    const currencySymbol = pickFirst(
      node.currencySymbol || (node.currency && node.currency.currencySymbol)
    );

    const nodeContacts = extractContacts(node);
    const ownerName =
      pickFirst(
        node.projectManagerName ||
          node.expertAdvisorName ||
          node.pmName ||
          node.expertAdvisor ||
          node.projectManager
      ) ||
      nodeContacts.names[0] ||
      project.ownerName ||
      "Unassigned";
    const scrubbedEmails = pdfUrl ? await scrubPdfForEmails(pdfUrl, invoiceNumber) : [];
    const associatedUserIds = dedupeStrings(
      extractAssociatedUserIds(node).concat(extractAssociatedUserIds(project.raw || {}))
    );
    const associatedEmails = dedupeEmails(
      nodeContacts.emails
        .concat(project.ownerEmails || [])
        .concat(scrubbedEmails)
        .concat(extractAssociatedEmails(node))
        .concat(
          associatedUserIds.includes(String(state.context.userId || "").trim()) &&
            state.access.email
            ? [state.access.email]
            : []
        )
    );

    return normalizeInvoice({
      id:
        pickFirst(
          node.id ||
            node._id ||
            node.fileId ||
            node.documentId ||
            node.invoiceId ||
            node.invoiceNumber
        ) ||
        createId(),
      invoiceNumber,
      invoiceName,
      ownerName,
      accountName:
        pickFirst(
          node.accountName ||
            node.companyName ||
            (node.company && (node.company.companyName || node.company.name)) ||
            (node.account && node.account.name) ||
            (node.customer && node.customer.name)
        ) || project.accountName || state.context.accountName || "Rocketlane Account",
      invoiceDate,
      issueDate: invoiceDate,
      dueDate,
      invoiceStatus,
      amount: Number.isFinite(amount) ? amount : 0,
      currencyCode,
      currencySymbol,
      pdfUrl,
      associatedEmails,
      associatedUserIds,
      sourceProjectName: project.name,
    });
  }

  async function scrubPdfForEmails(pdfUrl, invoiceRef) {
    try {
      ensurePdfJsAvailable();
    } catch (error) {
      appendLog("PDF_LIB_UNAVAILABLE", "PDF.js was unavailable for PDF email scrubbing.", error);
      return [];
    }

    try {
      const buffer = await fetchPdfArrayBuffer(pdfUrl);
      if (!buffer) {
        return [];
      }
      const text = await extractTextFromPdfBuffer(buffer, MAX_PDF_SCRUB_PAGES);
      return extractEmailsFromText(text);
    } catch (error) {
      appendLog(
        "PDF_SCRUB_FAILED",
        "Failed to scrub PDF for access emails (" + invoiceRef + ").",
        error
      );
      return [];
    }
  }

  async function fetchPdfArrayBuffer(pdfUrl) {
    if (pdfUrl.startsWith("data:application/pdf;base64,")) {
      const encoded = pdfUrl.slice("data:application/pdf;base64,".length);
      const bytes = window.atob(encoded);
      const len = bytes.length;
      const array = new Uint8Array(len);
      for (let i = 0; i < len; i += 1) {
        array[i] = bytes.charCodeAt(i);
      }
      return array.buffer;
    }

    const response = await fetch(pdfUrl, {
      method: "GET",
      credentials: "include",
    });
    if (!response.ok) {
      throw new Error("Unable to fetch PDF (" + response.status + ")");
    }
    return await response.arrayBuffer();
  }

  async function extractTextFromPdfBuffer(buffer, maxPages) {
    const loadingTask = window.pdfjsLib.getDocument({ data: buffer });
    try {
      const pdfDocument = await loadingTask.promise;
      const pages = Math.min(pdfDocument.numPages || 0, maxPages || pdfDocument.numPages);
      const fragments = [];
      for (let pageNumber = 1; pageNumber <= pages; pageNumber += 1) {
        const page = await pdfDocument.getPage(pageNumber);
        const content = await page.getTextContent();
        fragments.push(content.items.map((item) => item.str).join(" "));
      }
      if (typeof pdfDocument.cleanup === "function") {
        pdfDocument.cleanup();
      }
      return fragments.join("\n");
    } finally {
      if (loadingTask && typeof loadingTask.destroy === "function") {
        loadingTask.destroy();
      }
    }
  }

  function ensurePdfJsAvailable() {
    if (!window.pdfjsLib || typeof window.pdfjsLib.getDocument !== "function") {
      throw new Error("PDF.js is not available.");
    }
    if (window.pdfjsLib.GlobalWorkerOptions) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_CDN;
    }
  }

  function extractEmailsFromText(text) {
    return dedupeEmails(String(text || "").match(EMAIL_PATTERN) || []);
  }

  function normalizeInvoice(invoice) {
    if (!invoice || typeof invoice !== "object") {
      return null;
    }
    const pdfUrl = String(invoice.pdfUrl || "").trim();

    return {
      id: String(invoice.id || createId()),
      invoiceId: String(invoice.invoiceId || invoice.id || "").trim(),
      invoiceNumber: String(invoice.invoiceNumber || "Unknown").trim(),
      invoiceName: String(invoice.invoiceName || "Untitled invoice").trim(),
      ownerName: String(invoice.ownerName || "Unassigned").trim(),
      accountName: String(invoice.accountName || "Rocketlane Account").trim(),
      invoiceDate: String(invoice.invoiceDate || invoice.issueDate || new Date().toISOString()),
      issueDate: String(invoice.issueDate || invoice.invoiceDate || new Date().toISOString()),
      dueDate: String(invoice.dueDate || ""),
      invoiceStatus: String(invoice.invoiceStatus || invoice.status || "Unknown").trim(),
      amount: Number(invoice.amount || 0),
      currencyCode: String(invoice.currencyCode || "").trim(),
      currencySymbol: String(invoice.currencySymbol || "").trim(),
      pdfUrl,
      associatedEmails: dedupeEmails(invoice.associatedEmails || []),
      associatedUserIds: dedupeStrings(invoice.associatedUserIds || []),
      sourceProjectName: String(invoice.sourceProjectName || "").trim(),
    };
  }

  function dedupeInvoices(invoices) {
    const seen = new Set();
    const result = [];
    invoices.forEach((invoice) => {
      const normalized = normalizeInvoice(invoice);
      if (!normalized) {
        return;
      }
      const key =
        normalized.invoiceNumber.toLowerCase() +
        "|" +
        normalized.pdfUrl.toLowerCase() +
        "|" +
        normalized.sourceProjectName.toLowerCase();
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      result.push(normalized);
    });
    return result.sort(
      (a, b) => timestampValue(b.issueDate || b.invoiceDate) - timestampValue(a.issueDate || a.invoiceDate)
    );
  }

  function toAbsoluteUrl(value) {
    try {
      return new URL(value, window.location.origin).toString();
    } catch (_error) {
      return value;
    }
  }

  function extractInvoiceNumberFromText(text) {
    const match = String(text || "").match(/INV[-\s]?[\w\d-]+/i);
    return match ? match[0].toUpperCase().replace(/\s+/g, "-") : "";
  }

  function extractContacts(value) {
    const names = [];
    const emails = [];
    const visited = new Set();

    function walk(node, depth, parentKey) {
      if (depth > 6 || node == null) {
        return;
      }
      if (typeof node === "string") {
        const maybeEmail = normalizeEmail(node);
        if (isValidEmail(maybeEmail)) {
          emails.push(maybeEmail);
        } else if (
          parentKey &&
          /(name|manager|advisor|owner|lead|member|assignee|user)/i.test(parentKey)
        ) {
          names.push(node.trim());
        }
        return;
      }

      if (typeof node !== "object") {
        return;
      }
      if (visited.has(node)) {
        return;
      }
      visited.add(node);

      if (Array.isArray(node)) {
        node.forEach((entry) => walk(entry, depth + 1, parentKey));
        return;
      }

      const directName = pickFirst(
        node.name || node.fullName || node.displayName || node.userName
      );
      const directEmail = normalizeEmail(
        pickFirst(node.email || node.workEmail || node.userEmail)
      );
      if (directName) {
        names.push(directName);
      }
      if (isValidEmail(directEmail)) {
        emails.push(directEmail);
      }

      Object.keys(node).forEach((key) => {
        walk(node[key], depth + 1, key);
      });
    }

    walk(value, 0, "");
    return {
      names: dedupeStrings(names),
      emails: dedupeEmails(emails),
    };
  }

  function updateHeader() {
    const view = state.context.projectName || "Cross-project view";
    refs.scopeText.textContent =
      "Scope: " + state.context.accountName + " / " + view + " invoices";

    if (state.connected) {
      refs.connectionBadge.className = "badge badge-ok";
      refs.connectionBadge.textContent = "Connected to Rocketlane";
    } else {
      refs.connectionBadge.className = "badge badge-local";
      refs.connectionBadge.textContent = "Local preview mode";
    }

    refs.roleBadge.className =
      "badge " + (state.access.isAdmin ? "badge-admin" : "badge-muted");
    const identity = state.access.email || state.access.displayName || "Unknown user";
    refs.roleBadge.textContent = state.access.roleLabel + " - " + identity;
  }

  function configureUiForAccess() {
    if (state.access.isAdmin) {
      refs.tabLogsButton.classList.remove("hidden");
      refs.tabLogsButton.textContent = "Diagnostics";
      refs.tabLogsButton.classList.remove("locked");
      refs.clearLogsButton.classList.remove("hidden");
      if (refs.contextPanel) {
        refs.contextPanel.classList.remove("hidden");
      }
      if (refs.layoutGrid) {
        refs.layoutGrid.classList.remove("single-column");
      }
      refs.logsInfoText.textContent =
        "Admin users can view full diagnostics, error context, and remediation suggestions.";
    } else {
      refs.tabLogsButton.classList.add("hidden");
      refs.tabLogsButton.classList.remove("active");
      refs.tabLogsButton.setAttribute("aria-selected", "false");
      refs.clearLogsButton.classList.add("hidden");
      if (refs.contextPanel) {
        refs.contextPanel.classList.add("hidden");
      }
      if (refs.layoutGrid) {
        refs.layoutGrid.classList.add("single-column");
      }
      if (state.activeTab === "logs") {
        state.activeTab = "invoices";
      }
      refs.tabLogs.classList.add("hidden");
      refs.tabInvoices.classList.remove("hidden");
      refs.tabInvoicesButton.classList.add("active");
      refs.tabInvoicesButton.setAttribute("aria-selected", "true");
      refs.logsInfoText.textContent =
        "Diagnostics details are admin-only. This tab still provides access guidance.";
    }
  }

  function setActiveTab(tab) {
    if (tab === "logs" && !state.access.isAdmin) {
      tab = "invoices";
    }
    state.activeTab = tab;
    const showInvoices = tab === "invoices";
    refs.tabInvoicesButton.classList.toggle("active", showInvoices);
    refs.tabInvoicesButton.setAttribute("aria-selected", String(showInvoices));
    refs.tabInvoices.classList.toggle("hidden", !showInvoices);

    const showLogs = !showInvoices;
    refs.tabLogsButton.classList.toggle("active", showLogs);
    refs.tabLogsButton.setAttribute("aria-selected", String(showLogs));
    refs.tabLogs.classList.toggle("hidden", !showLogs);

    if (showLogs) {
      renderLogs();
    }
  }

  function renderAll() {
    renderSyncStatus();
    renderVisibilitySummary();
    renderFilterControls();
    renderSearchInsight();
    renderSortHeaders();
    renderInvoiceStats();
    renderInvoiceTable();
    renderPagination();
    renderExportInsight();
    renderSelectedSummary();
    renderSourceProjects();
    renderLogs();
  }

  function renderSyncStatus() {
    refs.syncStatus.textContent = state.syncStatus || "";
  }

  function renderVisibilitySummary() {
    if (state.access.isAdmin) {
      refs.visibilitySummary.textContent =
        "Admin access: all invoices from source projects are visible.";
      return;
    }
    if (state.access.email) {
      refs.visibilitySummary.textContent =
        "Restricted access: only invoices associated with " +
        state.access.email +
        " are visible.";
      return;
    }
    refs.visibilitySummary.textContent =
      "Restricted access: sign-in email was not detected, so no invoices are visible.";
  }

  function renderInvoiceStats() {
    const matched = getVisibleInvoices();
    const paged = getCurrentPageInvoices(matched);
    refs.invoiceStats.textContent =
      paged.length +
      " shown / " +
      matched.length +
      " matched / " +
      state.invoices.length +
      " total";
  }

  function renderInvoiceTable() {
    const matched = getVisibleInvoices();
    const visible = getCurrentPageInvoices(matched);
    refs.invoiceTableBody.innerHTML = "";

    if (!matched.length) {
      if (state.searchQuery || countActiveFilters() > 0) {
        if (refs.invoiceEmptyTitle) {
          refs.invoiceEmptyTitle.textContent = "No invoices match your search/filter";
        }
        if (refs.invoiceEmptyBody) {
          refs.invoiceEmptyBody.textContent =
            "No invoices were found with the current filters. Try a different search input or clear filters.";
        }
      } else {
        if (refs.invoiceEmptyTitle) {
          refs.invoiceEmptyTitle.textContent = "No invoices found";
        }
        if (refs.invoiceEmptyBody) {
          refs.invoiceEmptyBody.textContent =
            "No invoices were discovered in the source projects.";
        }
      }
      refs.invoiceEmptyState.classList.remove("hidden");
      return;
    }
    refs.invoiceEmptyState.classList.add("hidden");

    visible.forEach((invoice) => {
      const row = document.createElement("tr");
      row.className = "invoice-row";
      if (invoice.id === state.selectedInvoiceId) {
        row.classList.add("selected");
      }

      row.addEventListener("click", () => {
        state.selectedInvoiceId = invoice.id;
        renderInvoiceTable();
        renderSelectedSummary();
      });

      const numberCell = document.createElement("td");
      const numberButton = document.createElement("button");
      numberButton.type = "button";
      numberButton.className = "invoice-link";
      numberButton.textContent = invoice.invoiceNumber;
      numberButton.addEventListener("click", (event) => {
        event.stopPropagation();
        state.selectedInvoiceId = invoice.id;
        renderInvoiceTable();
        renderSelectedSummary();
        openPdfModal(invoice);
      });
      numberCell.appendChild(numberButton);

      const statusCell = document.createElement("td");
      statusCell.textContent = formatStatus(invoice.invoiceStatus);

      const amountCell = document.createElement("td");
      amountCell.className = "amount-cell";
      amountCell.textContent = formatAmount(invoice.amount, invoice.currencyCode, invoice.currencySymbol);

      const ownerCell = document.createElement("td");
      ownerCell.textContent = invoice.ownerName || "Unassigned";

      const accountCell = document.createElement("td");
      accountCell.textContent = invoice.accountName;

      const issueDateCell = document.createElement("td");
      issueDateCell.textContent = formatDate(invoice.issueDate || invoice.invoiceDate);

      const dueDateCell = document.createElement("td");
      dueDateCell.textContent = formatDate(invoice.dueDate);

      row.appendChild(statusCell);
      row.appendChild(numberCell);
      row.appendChild(ownerCell);
      row.appendChild(amountCell);
      row.appendChild(accountCell);
      row.appendChild(issueDateCell);
      row.appendChild(dueDateCell);
      refs.invoiceTableBody.appendChild(row);
    });
  }

  function renderSelectedSummary() {
    const invoice = getSelectedVisibleInvoice();
    if (!invoice) {
      refs.selectedInvoiceSummary.textContent =
        "Click an invoice number to preview invoice details.";
      return;
    }

    refs.selectedInvoiceSummary.textContent =
      formatStatus(invoice.invoiceStatus) +
      " · " +
      invoice.invoiceNumber +
      " · " +
      formatAmount(invoice.amount, invoice.currencyCode, invoice.currencySymbol) +
      " · " +
      formatDate(invoice.issueDate || invoice.invoiceDate) +
      (invoice.pdfUrl ? "" : " · showing invoice details preview");
  }

  function renderSourceProjects() {
    if (!state.sourceProjects.length) {
      refs.sourceProjectsText.textContent = "No source project found yet.";
      return;
    }
    refs.sourceProjectsText.textContent = state.sourceProjects.join(", ");
  }

  function openPdfModal(invoice) {
    if (!invoice) {
      return;
    }
    refs.modalTitle.textContent = invoice.invoiceNumber + " · " + invoice.invoiceName;
    refs.pdfModal.classList.remove("hidden");
    refs.pdfModal.setAttribute("aria-hidden", "false");
    if (invoice.pdfUrl) {
      refs.modalInvoicePreview.classList.add("hidden");
      refs.modalInvoicePreview.innerHTML = "";
      refs.modalPdfFrame.classList.remove("hidden");
      refs.modalPdfFrame.setAttribute("src", invoice.pdfUrl);
      return;
    }
    refs.modalPdfFrame.classList.add("hidden");
    refs.modalPdfFrame.removeAttribute("src");
    refs.modalInvoicePreview.classList.remove("hidden");
    renderInvoicePreviewContent(invoice, null, true);
    loadInvoicePreview(invoice);
  }

  function closePdfModal() {
    refs.pdfModal.classList.add("hidden");
    refs.pdfModal.setAttribute("aria-hidden", "true");
    refs.modalPdfFrame.removeAttribute("src");
    refs.modalPdfFrame.classList.remove("hidden");
    refs.modalInvoicePreview.classList.add("hidden");
    refs.modalInvoicePreview.innerHTML = "";
  }

  function renderSearchInsight() {
    if (!refs.searchInsight) {
      return;
    }
    const matchedCount = getVisibleInvoices().length;
    const totalCount = state.invoices.length;
    const activeFilterCount = countActiveFilters();
    if (!state.searchQuery && activeFilterCount === 0) {
      refs.searchInsight.textContent =
        "Search and filters run across all invoice columns (status, number, project manager, amount, account, issue date, due date).";
      return;
    }
    const queryText = state.searchQuery ? '"' + state.searchQuery + '"' : "current filters";
    if (matchedCount > 0) {
      refs.searchInsight.textContent =
        "Showing " +
        matchedCount +
        " matching invoice(s) for " +
        queryText +
        " from " +
        totalCount +
        " loaded invoice(s).";
      return;
    }
    if (state.searchQuery && state.searchServerCheckedQuery === state.searchQuery) {
      refs.searchInsight.textContent =
        "No matches for " +
        queryText +
        ". Server verified " +
        (state.searchServerMatchedCount == null ? 0 : state.searchServerMatchedCount) +
        " match(es). Try a different search input.";
      return;
    }
    refs.searchInsight.textContent = state.searchQuery
      ? "No local matches for " +
        queryText +
        ". Verifying against source projects..."
      : "No invoices match the current filters. Try a different filter combination.";
  }

  function renderFilterControls() {
    populateSelectOptions(refs.filterInvoiceStatus, uniqueFieldValues("invoiceStatus").map(formatStatus));
    populateSelectOptions(refs.filterProjectManager, uniqueFieldValues("ownerName"));
    populateSelectOptions(
      refs.filterAmount,
      uniqueFieldValues("amount").map((value) =>
        formatAmount(Number(value || 0), "", "$")
      )
    );
    populateSelectOptions(refs.filterAccount, uniqueFieldValues("accountName"));
    populateSelectOptions(
      refs.filterIssueDate,
      uniqueFieldValues("issueDate", "invoiceDate").map((value) => formatDate(value))
    );
    populateSelectOptions(
      refs.filterDueDate,
      uniqueFieldValues("dueDate").map((value) => formatDate(value))
    );
    refs.filterInvoiceStatus.value = state.filters.invoiceStatus;
    refs.filterInvoiceNumber.value = state.filters.invoiceNumber;
    refs.filterProjectManager.value = state.filters.ownerName;
    refs.filterAmount.value = state.filters.amount;
    refs.filterAccount.value = state.filters.accountName;
    refs.filterIssueDate.value = state.filters.issueDate;
    refs.filterDueDate.value = state.filters.dueDate;
  }

  function uniqueFieldValues(primaryField, fallbackField) {
    const scoped = getAccessibleInvoices();
    const values = scoped
      .map((invoice) => pickFirst(invoice[primaryField] || (fallbackField ? invoice[fallbackField] : "")))
      .filter(Boolean);
    return dedupeStrings(values).sort((a, b) => a.localeCompare(b));
  }

  function populateSelectOptions(selectRef, values) {
    if (!selectRef) {
      return;
    }
    const currentValue = String(selectRef.value || "");
    const seen = new Set([""]);
    const options = ['<option value="">All</option>'];
    values.forEach((value) => {
      const normalized = String(value || "").trim();
      if (!normalized || seen.has(normalized)) {
        return;
      }
      seen.add(normalized);
      options.push(
        '<option value="' + escapeHtml(normalized) + '">' + escapeHtml(normalized) + "</option>"
      );
    });
    selectRef.innerHTML = options.join("");
    if (currentValue && seen.has(currentValue)) {
      selectRef.value = currentValue;
    }
  }

  function countActiveFilters() {
    return [
      state.filters.invoiceStatus,
      state.filters.invoiceNumber,
      state.filters.ownerName,
      state.filters.amount,
      state.filters.accountName,
      state.filters.issueDate,
      state.filters.dueDate,
    ].filter(Boolean).length;
  }

  function clearFilters() {
    state.filters = {
      invoiceStatus: "",
      invoiceNumber: "",
      ownerName: "",
      amount: "",
      accountName: "",
      issueDate: "",
      dueDate: "",
    };
    state.currentPage = 1;
    ensureSelectedInvoice();
    renderAll();
  }

  function renderSortHeaders() {
    refs.sortButtons.forEach((button) => {
      const key = String(button.getAttribute("data-sort") || "");
      const label = String(button.textContent || "").replace(/\s*[↑↓]$/, "");
      if (key === state.sortBy) {
        button.textContent = label + (state.sortDirection === "asc" ? " ↑" : " ↓");
      } else {
        button.textContent = label;
      }
    });
  }

  function defaultSortDirectionForColumn(column) {
    if (column === "issueDate" || column === "dueDate" || column === "amount") {
      return "desc";
    }
    return "asc";
  }

  function renderPagination() {
    const matched = getVisibleInvoices();
    const totalPages = Math.max(1, Math.ceil(matched.length / state.pageSize));
    if (state.currentPage > totalPages) {
      state.currentPage = totalPages;
    }
    if (!matched.length || matched.length <= state.pageSize) {
      refs.paginationBar.classList.add("hidden");
      return;
    }
    refs.paginationBar.classList.remove("hidden");
    refs.paginationInfo.textContent =
      "Page " + state.currentPage + " of " + totalPages + " (" + matched.length + " matches)";
    refs.paginationPrevButton.disabled = state.currentPage <= 1;
    refs.paginationNextButton.disabled = state.currentPage >= totalPages;
  }

  function changePage(delta) {
    const matched = getVisibleInvoices();
    const totalPages = Math.max(1, Math.ceil(matched.length / state.pageSize));
    const nextPage = Math.min(totalPages, Math.max(1, state.currentPage + delta));
    if (nextPage === state.currentPage) {
      return;
    }
    state.currentPage = nextPage;
    ensureSelectedInvoice();
    renderAll();
  }

  function getCurrentPageInvoices(visibleInvoices) {
    const rows = Array.isArray(visibleInvoices) ? visibleInvoices : getVisibleInvoices();
    if (!rows.length) {
      return [];
    }
    const totalPages = Math.max(1, Math.ceil(rows.length / state.pageSize));
    if (state.currentPage > totalPages) {
      state.currentPage = totalPages;
    }
    const start = (state.currentPage - 1) * state.pageSize;
    const end = start + state.pageSize;
    return rows.slice(start, end);
  }

  function renderExportInsight() {
    if (!refs.exportInsight) {
      return;
    }
    const matched = getVisibleInvoices();
    const one = getSelectedVisibleInvoice();
    refs.downloadZipButton.disabled = false;
    if (state.exportMode === "one") {
      refs.exportInsight.textContent = one
        ? "Will download selected invoice " + one.invoiceNumber + " as ZIP."
        : "Select an invoice row to export one invoice.";
      refs.downloadZipButton.disabled = !one;
      return;
    }
    if (state.exportMode === "filtered") {
      refs.exportInsight.textContent =
        "Will download " + matched.length + " filtered/search invoice(s) as ZIP.";
      refs.downloadZipButton.disabled = matched.length === 0;
      return;
    }
    refs.exportInsight.textContent =
      "Will download all " + getAccessibleInvoices().length + " accessible invoice(s) as ZIP.";
    refs.downloadZipButton.disabled = getAccessibleInvoices().length === 0;
  }

  function scheduleSearchVerification() {
    if (state.searchVerifyTimer) {
      window.clearTimeout(state.searchVerifyTimer);
      state.searchVerifyTimer = 0;
    }
    if (!state.connected || !state.searchQuery || state.searchQuery.length < 2) {
      return;
    }
    const queryAtSchedule = state.searchQuery;
    state.searchVerifyTimer = window.setTimeout(() => {
      verifySearchAgainstSourceProjects(queryAtSchedule).catch((_error) => {
        // Search verification is best-effort and should not block UI interactions.
      });
    }, 420);
  }

  async function verifySearchAgainstSourceProjects(query) {
    const normalizedQuery = String(query || "").trim().toLowerCase();
    if (!normalizedQuery || normalizedQuery !== state.searchQuery) {
      return;
    }
    const result = await fetchSearchMatchesFromServerAction(normalizedQuery);
    if (!result || normalizedQuery !== state.searchQuery) {
      return;
    }
    state.searchServerCheckedQuery = normalizedQuery;
    state.searchServerMatchedCount = Number(result.matchedInvoices || 0);
    renderSearchInsight();
  }

  async function fetchSearchMatchesFromServerAction(searchQuery) {
    if (
      !state.client ||
      !state.client.data ||
      typeof state.client.data.invoke !== "function"
    ) {
      return null;
    }
    const workspaceCandidates = [
      "https://blink.rocketlane.com",
      "https://innovate-calgary.rocketlane.com",
    ];
    const accountDomain = pickFirst(
      state.rawAccount &&
        (state.rawAccount.domainName ||
          state.rawAccount.primaryDomainName ||
          state.rawAccount.domain)
    );
    if (accountDomain) {
      workspaceCandidates.unshift(
        "https://" + accountDomain.replace(/^https?:\/\//i, "")
      );
    }
    try {
      const payload = await state.client.data.invoke("syncInvoicesFromSource", {
        sourceProjectNames: SOURCE_PROJECT_NAMES.slice(),
        accountName: state.context.accountName || "",
        workspaceBaseUrl: workspaceCandidates[0],
        workspaceCandidates,
        searchQuery,
        searchOnly: true,
      });
      const result = unwrapServerActionResponse(payload);
      if (!result || result.ok === false) {
        return null;
      }
      if (result.search && typeof result.search === "object") {
        return result.search;
      }
      const matches = Array.isArray(result.invoices) ? result.invoices.length : 0;
      return { query: searchQuery, matchedInvoices: matches };
    } catch (_error) {
      return null;
    }
  }

  async function loadInvoicePreview(invoice) {
    const cacheKey = String(invoice.id || invoice.invoiceId || invoice.invoiceNumber || "");
    if (cacheKey && state.invoicePreviewCache[cacheKey]) {
      renderInvoicePreviewContent(invoice, state.invoicePreviewCache[cacheKey], false);
      return;
    }
    try {
      const preview = await fetchInvoicePreviewFromServerAction(invoice);
      if (cacheKey && preview) {
        state.invoicePreviewCache[cacheKey] = preview;
      }
      renderInvoicePreviewContent(invoice, preview, false);
    } catch (_error) {
      renderInvoicePreviewContent(
        invoice,
        null,
        false,
        "Unable to load invoice preview details right now."
      );
    }
  }

  async function fetchInvoicePreviewFromServerAction(invoice) {
    if (
      !state.connected ||
      !state.client ||
      !state.client.data ||
      typeof state.client.data.invoke !== "function"
    ) {
      return null;
    }
    const previewInvoiceId = pickFirst(
      invoice.invoiceId || invoice.id || invoice.invoiceNumber
    );
    if (!previewInvoiceId) {
      return null;
    }
    const workspaceCandidates = [
      "https://blink.rocketlane.com",
      "https://innovate-calgary.rocketlane.com",
    ];
    const accountDomain = pickFirst(
      state.rawAccount &&
        (state.rawAccount.domainName ||
          state.rawAccount.primaryDomainName ||
          state.rawAccount.domain)
    );
    if (accountDomain) {
      workspaceCandidates.unshift(
        "https://" + accountDomain.replace(/^https?:\/\//i, "")
      );
    }
    const payload = await state.client.data.invoke("syncInvoicesFromSource", {
      sourceProjectNames: SOURCE_PROJECT_NAMES.slice(),
      accountName: state.context.accountName || "",
      workspaceBaseUrl: workspaceCandidates[0],
      workspaceCandidates,
      previewInvoiceId,
    });
    const result = unwrapServerActionResponse(payload);
    if (!result || result.ok === false) {
      return null;
    }
    return result.preview || null;
  }

  function renderInvoicePreviewContent(invoice, preview, isLoading, errorText) {
    if (!refs.modalInvoicePreview) {
      return;
    }
    if (isLoading) {
      refs.modalInvoicePreview.innerHTML =
        '<p class="muted">Loading invoice preview details...</p>';
      return;
    }
    if (errorText) {
      refs.modalInvoicePreview.innerHTML =
        '<p class="muted">' + escapeHtml(errorText) + "</p>";
      return;
    }
    const previewData = preview || {};
    const summaryRows = [
      ["Status", formatStatus(previewData.status || invoice.invoiceStatus)],
      ["Invoice #", invoice.invoiceNumber],
      [
        "Amount",
        formatAmount(
          previewData.amount != null ? previewData.amount : invoice.amount,
          previewData.currencyCode || invoice.currencyCode,
          previewData.currencySymbol || invoice.currencySymbol
        ),
      ],
      ["Account", previewData.accountName || invoice.accountName],
      [
        "Issue date",
        formatDate(previewData.issueDate || invoice.issueDate || invoice.invoiceDate),
      ],
      ["Due date", formatDate(previewData.dueDate || invoice.dueDate)],
    ];
    const lineItems = Array.isArray(previewData.lineItems) ? previewData.lineItems : [];
    const payments = Array.isArray(previewData.payments) ? previewData.payments : [];
    const lineRowsHtml = lineItems.length
      ? lineItems
          .map(
            (line) =>
              "<tr><td>" +
              escapeHtml(line.description || "Line item") +
              "</td><td>" +
              escapeHtml(String(line.quantity || 0)) +
              "</td><td>" +
              escapeHtml(
                formatAmount(
                  line.unitPrice || 0,
                  previewData.currencyCode || invoice.currencyCode,
                  previewData.currencySymbol || invoice.currencySymbol
                )
              ) +
              "</td><td>" +
              escapeHtml(
                formatAmount(
                  line.amount || 0,
                  previewData.currencyCode || invoice.currencyCode,
                  previewData.currencySymbol || invoice.currencySymbol
                )
              ) +
              "</td></tr>"
          )
          .join("")
      : '<tr><td colspan="4">No invoice line items returned by API.</td></tr>';
    const paymentRowsHtml = payments.length
      ? payments
          .map(
            (payment) =>
              "<tr><td>" +
              escapeHtml(payment.recordType || "Payment") +
              "</td><td>" +
              escapeHtml(formatDate(payment.paymentDate)) +
              "</td><td>" +
              escapeHtml(
                formatAmount(
                  payment.amount || 0,
                  previewData.currencyCode || invoice.currencyCode,
                  previewData.currencySymbol || invoice.currencySymbol
                )
              ) +
              "</td><td>" +
              escapeHtml(payment.notes || "-") +
              "</td></tr>"
          )
          .join("")
      : '<tr><td colspan="4">No payment records returned by API.</td></tr>';
    refs.modalInvoicePreview.innerHTML =
      '<div class="modal-preview-summary">' +
      summaryRows
        .map(
          (row) =>
            '<div class="modal-preview-row"><strong>' +
            escapeHtml(row[0]) +
            ":</strong> " +
            escapeHtml(row[1] || "Unknown") +
            "</div>"
        )
        .join("") +
      "</div>" +
      '<h4 class="modal-preview-section-title">Invoice line items</h4>' +
      '<table class="modal-preview-table"><thead><tr><th>Description</th><th>Qty</th><th>Unit price</th><th>Amount</th></tr></thead><tbody>' +
      lineRowsHtml +
      "</tbody></table>" +
      '<h4 class="modal-preview-section-title">Payments</h4>' +
      '<table class="modal-preview-table"><thead><tr><th>Type</th><th>Date</th><th>Amount</th><th>Notes</th></tr></thead><tbody>' +
      paymentRowsHtml +
      "</tbody></table>";
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  async function onDownloadZip() {
    const JSZipCtor = window.JSZip;
    if (!JSZipCtor) {
      appendLog(
        "SOURCE_FETCH_FAILED",
        "ZIP library missing. jszip.min.js did not load."
      );
      if (refs.exportInsight) {
        refs.exportInsight.textContent =
          "Unable to export: ZIP library unavailable.";
      }
      return;
    }
    let invoicesToExport = [];
    if (state.exportMode === "one") {
      const selected = getSelectedVisibleInvoice() || getCurrentPageInvoices()[0];
      invoicesToExport = selected ? [selected] : [];
    } else if (state.exportMode === "filtered") {
      invoicesToExport = getVisibleInvoices();
    } else {
      invoicesToExport = getAccessibleInvoices();
    }
    if (!invoicesToExport.length) {
      refs.exportInsight.textContent =
        "No invoices available for the selected export option.";
      return;
    }

    refs.exportInsight.textContent = "Preparing ZIP export...";
    const zip = new JSZipCtor();
    const csvRows = [
      [
        "Invoice Status",
        "Invoice Number",
        "Project Manager",
        "Amount",
        "Account",
        "Issue Date",
        "Due Date",
      ],
    ];

    for (let i = 0; i < invoicesToExport.length; i += 1) {
      const invoice = invoicesToExport[i];
      csvRows.push([
        formatStatus(invoice.invoiceStatus),
        invoice.invoiceNumber,
        invoice.ownerName || "",
        formatAmount(invoice.amount, invoice.currencyCode, invoice.currencySymbol),
        invoice.accountName,
        formatDate(invoice.issueDate || invoice.invoiceDate),
        formatDate(invoice.dueDate),
      ]);
      const exportRecord = {
        invoiceStatus: formatStatus(invoice.invoiceStatus),
        invoiceNumber: invoice.invoiceNumber,
        projectManager: invoice.ownerName || "",
        amount: invoice.amount || 0,
        currencyCode: invoice.currencyCode || "",
        currencySymbol: invoice.currencySymbol || "",
        account: invoice.accountName,
        issueDate: invoice.issueDate || invoice.invoiceDate,
        dueDate: invoice.dueDate || "",
        sourceProjectName: invoice.sourceProjectName || "",
        associatedEmails: invoice.associatedEmails || [],
        associatedUserIds: invoice.associatedUserIds || [],
      };
      if (state.exportMode === "one") {
        try {
          const preview = await fetchInvoicePreviewFromServerAction(invoice);
          if (preview) {
            exportRecord.preview = preview;
          }
        } catch (_error) {
          // Keep export resilient even if preview detail lookup fails.
        }
      }
      zip.file(
        "invoices/" + safeFileName(invoice.invoiceNumber || invoice.id || "invoice") + ".json",
        JSON.stringify(exportRecord, null, 2)
      );
    }

    zip.file("invoices.csv", toCsv(csvRows));
    const modeLabel = state.exportMode === "one" ? "single" : state.exportMode;
    const blob = await zip.generateAsync({ type: "blob" });
    const fileName =
      "invoice-export-" + modeLabel + "-" + new Date().toISOString().slice(0, 10) + ".zip";
    downloadBlob(blob, fileName);
    refs.exportInsight.textContent =
      "Downloaded " + invoicesToExport.length + " invoice(s) to " + fileName + ".";
  }

  function downloadBlob(blob, fileName) {
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    window.setTimeout(() => {
      window.URL.revokeObjectURL(url);
    }, 500);
  }

  function safeFileName(value) {
    return String(value || "invoice")
      .replace(/[^a-z0-9._-]+/gi, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 80);
  }

  function toCsv(rows) {
    return rows
      .map((row) =>
        row
          .map((cell) => {
            const text = String(cell == null ? "" : cell);
            if (/[",\n]/.test(text)) {
              return '"' + text.replace(/"/g, '""') + '"';
            }
            return text;
          })
          .join(",")
      )
      .join("\n");
  }

  function getVisibleInvoices() {
    let invoices = getAccessibleInvoices();
    invoices = applySearchAndFilters(invoices);
    invoices = sortInvoices(invoices, state.sortBy, state.sortDirection);
    return invoices;
  }

  function getAccessibleInvoices() {
    let invoices = state.invoices.slice();
    if (state.access.isAdmin) {
      return invoices;
    }
    const email = state.access.email;
    const userId = String(state.context.userId || "").trim();
    if (!email && !userId) {
      return [];
    }
    return invoices.filter(
      (invoice) =>
        (email && invoice.associatedEmails.includes(email)) ||
        (userId && invoice.associatedUserIds.includes(userId))
    );
  }

  function applySearchAndFilters(invoices) {
    const query = String(state.searchQuery || "").trim().toLowerCase();
    const filters = state.filters || {};
    return invoices.filter((invoice) => {
      const status = formatStatus(invoice.invoiceStatus);
      const issueDateLabel = formatDate(invoice.issueDate || invoice.invoiceDate);
      const dueDateLabel = formatDate(invoice.dueDate);
      const amountLabel = formatAmount(invoice.amount, invoice.currencyCode, invoice.currencySymbol);
      if (filters.invoiceStatus && status !== filters.invoiceStatus) {
        return false;
      }
      if (filters.invoiceNumber && !String(invoice.invoiceNumber || "").toLowerCase().includes(filters.invoiceNumber)) {
        return false;
      }
      if (filters.ownerName && String(invoice.ownerName || "") !== filters.ownerName) {
        return false;
      }
      if (filters.amount && amountLabel !== filters.amount) {
        return false;
      }
      if (filters.accountName && String(invoice.accountName || "") !== filters.accountName) {
        return false;
      }
      if (filters.issueDate && issueDateLabel !== filters.issueDate) {
        return false;
      }
      if (filters.dueDate && dueDateLabel !== filters.dueDate) {
        return false;
      }
      if (!query) {
        return true;
      }
      const haystack = (
        status +
        " " +
        String(invoice.invoiceNumber || "") +
        " " +
        String(invoice.ownerName || "") +
        " " +
        amountLabel +
        " " +
        String(invoice.accountName || "") +
        " " +
        issueDateLabel +
        " " +
        dueDateLabel
      ).toLowerCase();
      return haystack.includes(query);
    });
  }

  function sortInvoices(invoices, sortBy, sortDirection) {
    const direction = sortDirection === "asc" ? 1 : -1;
    const rows = invoices.slice();
    rows.sort((a, b) => {
      if (sortBy === "amount") {
        return direction * ((Number(a.amount || 0) - Number(b.amount || 0)) || 0);
      }
      if (sortBy === "issueDate" || sortBy === "dueDate") {
        const aTime = timestampValue(a[sortBy] || (sortBy === "issueDate" ? a.invoiceDate : ""));
        const bTime = timestampValue(b[sortBy] || (sortBy === "issueDate" ? b.invoiceDate : ""));
        return direction * (aTime - bTime);
      }
      const aText = String(a[sortBy] || "").toLowerCase();
      const bText = String(b[sortBy] || "").toLowerCase();
      if (aText < bText) {
        return -1 * direction;
      }
      if (aText > bText) {
        return 1 * direction;
      }
      return 0;
    });
    return rows;
  }

  function hasPermissionSignals(hint) {
    if (!hint || typeof hint !== "object") {
      return false;
    }
    if (hint.isAdmin === true) {
      return true;
    }
    const permission = String(hint.permission || "").trim();
    const roleLabel = String(hint.roleLabel || "").trim();
    return Boolean(permission || roleLabel);
  }

  function ensureSelectedInvoice() {
    const visible = getCurrentPageInvoices();
    if (!visible.length) {
      state.selectedInvoiceId = null;
      return;
    }
    const present = visible.some((invoice) => invoice.id === state.selectedInvoiceId);
    if (!present) {
      state.selectedInvoiceId = visible[0].id;
    }
  }

  function getSelectedVisibleInvoice() {
    const visible = getCurrentPageInvoices();
    return visible.find((invoice) => invoice.id === state.selectedInvoiceId) || null;
  }

  function createLogsStorageKey(context) {
    const accountScope = slug(context.accountId || context.accountName || "workspace");
    return [STORAGE_PREFIX, "logs", STORAGE_VERSION, accountScope].join(":");
  }

  function loadLogs() {
    const raw = safeStorageGet(state.storageKey);
    if (!raw) {
      state.logs = [];
      return;
    }
    try {
      const parsed = JSON.parse(raw);
      state.logs = Array.isArray(parsed)
        ? parsed.map(normalizeLog).filter(Boolean).slice(0, LOG_LIMIT)
        : [];
    } catch (_error) {
      state.logs = [];
    }
  }

  function persistLogs() {
    try {
      safeStorageSet(state.storageKey, JSON.stringify(state.logs.slice(0, LOG_LIMIT)));
    } catch (_error) {
      // Ignore storage failures for logs.
    }
  }

  function appendLog(code, message, error, suggestion) {
    const entry = {
      id: createId(),
      createdAt: new Date().toISOString(),
      code: code || "UNKNOWN",
      message: message || "Unexpected app event.",
      details: simplifyError(error),
      suggestion: suggestion || SUGGESTIONS[code] || "Review app configuration and retry.",
    };

    const duplicate = state.logs.find(
      (item) => item.code === entry.code && item.message === entry.message
    );
    if (duplicate) {
      return;
    }

    state.logs.unshift(entry);
    if (state.logs.length > LOG_LIMIT) {
      state.logs = state.logs.slice(0, LOG_LIMIT);
    }
    persistLogs();
  }

  function renderLogs() {
    if (!state.access.isAdmin) {
      refs.logsList.innerHTML = "";
      refs.logsEmptyState.classList.add("hidden");
      refs.logsAccessNotice.classList.remove("hidden");
      const identity = state.access.email || state.access.displayName || "this user";
      refs.logsAccessNotice.textContent =
        "You are signed in as " +
        identity +
        ". Full diagnostics are visible to admins only. Contact a workspace admin to review detailed logs and solution suggestions.";
      return;
    }

    refs.logsAccessNotice.classList.add("hidden");
    refs.logsList.innerHTML = "";
    if (!state.logs.length) {
      refs.logsEmptyState.classList.remove("hidden");
      return;
    }
    refs.logsEmptyState.classList.add("hidden");

    state.logs.forEach((entry) => {
      const card = document.createElement("article");
      card.className = "log-item";

      const header = document.createElement("div");
      header.className = "log-item-header";

      const code = document.createElement("span");
      code.className = "log-code";
      code.textContent = entry.code;

      const time = document.createElement("span");
      time.className = "log-time";
      time.textContent = formatDateTime(entry.createdAt);

      header.appendChild(code);
      header.appendChild(time);

      const message = document.createElement("p");
      message.className = "log-message";
      message.textContent = entry.message;

      const suggestion = document.createElement("p");
      suggestion.className = "log-suggestion";
      suggestion.textContent = "Suggested fix: " + entry.suggestion;

      card.appendChild(header);
      card.appendChild(message);
      if (entry.details) {
        const details = document.createElement("p");
        details.className = "log-suggestion";
        details.textContent = "Details: " + entry.details;
        card.appendChild(details);
      }
      card.appendChild(suggestion);
      refs.logsList.appendChild(card);
    });
  }

  function onClearLogs() {
    if (!state.access.isAdmin || !state.logs.length) {
      return;
    }
    const confirmed = window.confirm("Clear all diagnostic logs?");
    if (!confirmed) {
      return;
    }
    state.logs = [];
    persistLogs();
    renderLogs();
  }

  function deriveAccessProfile(rawUser, rawAccount, context, permissionHint) {
    const displayName =
      pickFirst(
        rawUser &&
          (rawUser.name || rawUser.fullName || rawUser.displayName || rawUser.email)
      ) || context.userName;
    const email = normalizeEmail(extractPrimaryEmail(rawUser) || context.userEmail || "");
    const permissionLabel =
      (permissionHint && permissionHint.permission) || extractPermissionLabel(rawUser);
    const roleLabelHint =
      (permissionHint && permissionHint.roleLabel) || extractRoleLabel(rawUser) || context.userRole;
    const permissionRole = normalizePermissionRole(permissionLabel || roleLabelHint);
    const inferredRole = inferRole(rawUser, rawAccount, context.userRole);
    const forcedAdmin = Boolean(permissionHint && permissionHint.isAdmin === true);
    const role = forcedAdmin ? "admin" : permissionRole || inferredRole || "non_admin";
    const permissionValue = permissionLabel || roleLabelHint || "";
    const isAdmin = role === "admin";
    const roleLabel = resolveRoleLabel(role, { permission: permissionValue });

    return {
      role,
      roleLabel,
      isAdmin,
      email,
      displayName,
      permission: permissionValue,
    };
  }

  function resolveRoleLabel(role, permissionHint) {
    if (role === "admin") {
      const label = String((permissionHint && permissionHint.permission) || "").trim();
      if (label && isLikelyAdminLabel(label)) {
        return label;
      }
      return "Account Admin";
    }
    if (permissionHint && permissionHint.permission) {
      return permissionHint.permission;
    }
    if (role === "collaborator") {
      return "Collaborator";
    }
    if (role === "expert_advisor") {
      return "Expert Advisor";
    }
    return "Restricted";
  }

  function inferRole(user, account, fallbackRole) {
    const normalizedFallback = normalizeRole(fallbackRole);
    if (normalizedFallback) {
      return normalizedFallback;
    }

    if (isAccountOwner(user, account)) {
      return "admin";
    }

    if (user && typeof user === "object") {
      if (
        user.isAdmin === true ||
        user.admin === true ||
        user.isAccountAdmin === true ||
        user.accountAdmin === true
      ) {
        return "admin";
      }
    }

    const tokens = [];
    collectRoleTokens(user, tokens, 0);
    collectRoleTokens(account, tokens, 0);
    const haystack = tokens
      .join(" ")
      .toLowerCase()
      .replace(/[_-]+/g, " ");

    if (
      /(^|\b)(account|workspace)\s*admin(istrator)?(\b|$)/.test(haystack) ||
      /(^|\b)admin(\b|$)/.test(haystack)
    ) {
      return "admin";
    }
    if (/(^|\b)(account|workspace|company)\s*owner(\b|$)/.test(haystack)) {
      return "admin";
    }
    if (/expert[\s_-]*advisor/.test(haystack)) {
      return "expert_advisor";
    }
    if (/(^|\b)collaborator(\b|$)/.test(haystack)) {
      return "collaborator";
    }
    return "non_admin";
  }

  function normalizePermissionRole(value) {
    const text = String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[_-]+/g, " ");
    if (!text) {
      return "";
    }
    if (
      text.includes("account admin") ||
      text.includes("workspace admin") ||
      text.includes("owner") ||
      text === "admin"
    ) {
      return "admin";
    }
    return "non_admin";
  }

  function isLikelyAdminLabel(value) {
    const text = String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[_-]+/g, " ");
    if (!text) {
      return false;
    }
    return (
      text.includes("account admin") ||
      text.includes("workspace admin") ||
      text.includes("account owner") ||
      text.includes("workspace owner") ||
      text.includes("company owner") ||
      text === "admin"
    );
  }

  function collectRoleTokens(value, target, depth) {
    if (depth > 6 || value == null) {
      return;
    }
    if (typeof value === "string" || typeof value === "number") {
      target.push(String(value));
      return;
    }
    if (Array.isArray(value)) {
      value.forEach((item) => collectRoleTokens(item, target, depth + 1));
      return;
    }
    if (typeof value !== "object") {
      return;
    }

    const directLabel = readTextValue(value);
    if (directLabel) {
      target.push(directLabel);
    }

    Object.keys(value).forEach((key) => {
      const lowered = key.toLowerCase();
      const next = value[key];
      if (
        lowered.includes("role") ||
        lowered.includes("permission") ||
        lowered.includes("type") ||
        lowered.includes("group") ||
        lowered.includes("admin") ||
        lowered === "response" ||
        lowered === "data" ||
        lowered === "payload" ||
        lowered === "result" ||
        lowered === "item" ||
        lowered === "user" ||
        lowered === "account"
      ) {
        collectRoleTokens(next, target, depth + 1);
      }
    });
  }

  function extractPrimaryEmail(user) {
    if (!user || typeof user !== "object") {
      return "";
    }
    const emails = extractContacts(user).emails;
    return emails[0] || "";
  }

  function normalizeRole(value) {
    const text = String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[_-]+/g, " ");
    if (!text) {
      return "";
    }
    if (
      text === "admin" ||
      text.includes("account admin") ||
      text.includes("workspace admin") ||
      text.includes("account administrator") ||
      text.includes("workspace administrator") ||
      text.includes("account owner") ||
      text.includes("workspace owner") ||
      text.includes("company owner")
    ) {
      return "admin";
    }
    if (text.includes("expert") && text.includes("advisor")) {
      return "expert_advisor";
    }
    if (text.includes("collaborator")) {
      return "collaborator";
    }
    return "";
  }

  function buildRowKey(row) {
    if (!row || typeof row !== "object") {
      return createId();
    }
    return (
      pickFirst(row.id || row._id || row.documentId || row.fileId) ||
      pickFirst(row.name || row.fileName || row.title || row.url || row.href) ||
      createId()
    );
  }

  function normalizeLog(entry) {
    if (!entry || typeof entry !== "object") {
      return null;
    }
    return {
      id: String(entry.id || createId()),
      createdAt: String(entry.createdAt || new Date().toISOString()),
      code: String(entry.code || "UNKNOWN"),
      message: String(entry.message || "Unexpected app event."),
      details: String(entry.details || ""),
      suggestion: String(entry.suggestion || "Review app configuration and retry."),
    };
  }

  function dedupeEmails(values) {
    const seen = new Set();
    const result = [];
    values.forEach((value) => {
      const email = normalizeEmail(value);
      if (!isValidEmail(email) || seen.has(email)) {
        return;
      }
      seen.add(email);
      result.push(email);
    });
    return result;
  }

  function dedupeStrings(values) {
    const seen = new Set();
    const result = [];
    values.forEach((value) => {
      const normalized = String(value || "").trim();
      if (!normalized) {
        return;
      }
      const key = normalized.toLowerCase();
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      result.push(normalized);
    });
    return result;
  }

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function isValidEmail(value) {
    return EMAIL_VALIDATION_PATTERN.test(String(value || "").trim());
  }

  function safeStorageGet(key) {
    try {
      return window.localStorage.getItem(key);
    } catch (_error) {
      return null;
    }
  }

  function safeStorageSet(key, value) {
    window.localStorage.setItem(key, value);
  }

  function simplifyError(error) {
    if (!error) {
      return "";
    }
    if (typeof error === "string") {
      return error.slice(0, 240);
    }
    if (error && typeof error.message === "string") {
      return error.message.slice(0, 240);
    }
    try {
      return JSON.stringify(error).slice(0, 240);
    } catch (_error) {
      return "Unknown error";
    }
  }

  function formatDate(value) {
    if (!value) {
      return "Unknown";
    }
    const timestamp = new Date(value).getTime();
    if (Number.isNaN(timestamp)) {
      return "Unknown";
    }
    return new Date(timestamp).toLocaleDateString([], {
      year: "numeric",
      month: "short",
      day: "numeric",
    });
  }

  function formatStatus(value) {
    const raw = String(value || "").trim();
    if (!raw) {
      return "Unknown";
    }
    return raw
      .toLowerCase()
      .replace(/[_-]+/g, " ")
      .replace(/\b\w/g, (ch) => ch.toUpperCase());
  }

  function formatAmount(amount, currencyCode, currencySymbol) {
    const numeric = Number(amount || 0);
    const code = String(currencyCode || "").trim().toUpperCase();
    if (!Number.isFinite(numeric)) {
      return "0.00";
    }
    if (code) {
      try {
        return new Intl.NumberFormat(undefined, {
          style: "currency",
          currency: code,
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        }).format(numeric);
      } catch (_error) {
        // Fall back when currency code is invalid.
      }
    }
    const symbol = String(currencySymbol || "").trim();
    return (symbol || "$") + numeric.toFixed(2);
  }

  function formatDateTime(value) {
    const timestamp = new Date(value).getTime();
    if (Number.isNaN(timestamp)) {
      return "Unknown";
    }
    return new Date(timestamp).toLocaleString([], {
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function timestampValue(value) {
    const timestamp = new Date(value).getTime();
    return Number.isNaN(timestamp) ? 0 : timestamp;
  }

  function slug(value) {
    return (
      String(value || "")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/(^-|-$)/g, "") || "default"
    );
  }

  function pickFirst(value) {
    if (typeof value === "string") {
      return value.trim();
    }
    if (typeof value === "number") {
      return String(value);
    }
    return "";
  }

  function readTextValue(value) {
    if (typeof value === "string" || typeof value === "number") {
      return pickFirst(value);
    }
    if (Array.isArray(value)) {
      for (let i = 0; i < value.length; i += 1) {
        const text = readTextValue(value[i]);
        if (text) {
          return text;
        }
      }
      return "";
    }
    if (!value || typeof value !== "object") {
      return "";
    }

    const wrapped = unwrapTopLevelObject(value);
    if (wrapped !== value) {
      const wrappedText = readTextValue(wrapped);
      if (wrappedText) {
        return wrappedText;
      }
    }

    return (
      pickFirst(
        value.name ||
          value.label ||
          value.displayName ||
          value.title ||
          value.value ||
        value.permissionName ||
        value.roleName ||
          value.role ||
          value.permission
      ) || ""
    );
  }

  function extractPermissionLabel(raw) {
    if (!raw || typeof raw !== "object") {
      return "";
    }
    const source = unwrapTopLevelObject(raw);
    const permissionsList = Array.isArray(raw.permissions)
      ? raw.permissions
          .map((item) => readTextValue(item))
          .filter(Boolean)
          .join(", ")
      : "";
    return readTextValue(
      source.permission ||
        (source.permission && source.permission.permissionName) ||
        (source.permission && source.permission.name) ||
        source.permissionName ||
        source.permissionSet ||
        (source.permissionSet && source.permissionSet.name) ||
        source.permissionSetObj ||
        source.accountPermission ||
        (source.accountPermission && source.accountPermission.permissionName) ||
        source.access ||
        source.permissions ||
        permissionsList
    );
  }

  function extractRoleLabel(raw) {
    if (!raw || typeof raw !== "object") {
      return "";
    }
    const source = unwrapTopLevelObject(raw);
    return readTextValue(
      source.role ||
        (source.role && source.role.roleName) ||
        (source.role && source.role.name) ||
        source.userRole ||
        source.designation ||
        source.title ||
        source.userType ||
        source.roleInfo ||
        (source.role && source.role.roleName)
    );
  }

  function isAccountOwner(user, account) {
    if (!user || !account) {
      return false;
    }

    const userView = unwrapTopLevelObject(user);
    const accountView = unwrapTopLevelObject(account);
    const userId = pickFirst(
      userView && (userView.id || userView.userId || userView.userID || userView._id)
    );
    const userEmail = normalizeEmail(
      pickFirst(
        userView &&
          (userView.email ||
            userView.emailId ||
            userView.userEmail ||
            userView.workEmail ||
            userView.userName)
      )
    );

    const candidates = [
      accountView && accountView.owner,
      accountView && accountView.accountOwner,
      accountView && accountView.primaryOwner,
      accountView && accountView.createdBy,
      accountView && accountView.createdByUser,
      accountView && accountView.createdByUserId,
      accountView && accountView.ownerId,
      accountView && accountView.accountOwnerId,
      accountView &&
        accountView.company &&
        (accountView.company.owner || accountView.company.createdBy),
    ];

    const listCandidates = [
      accountView && accountView.owners,
      accountView && accountView.accountOwners,
      accountView && accountView.admins,
      accountView && accountView.accountAdmins,
      accountView && accountView.administrators,
    ];
    listCandidates.forEach((entry) => {
      if (Array.isArray(entry)) {
        candidates.push(...entry);
      }
    });

    for (let i = 0; i < candidates.length; i += 1) {
      if (principalMatchesUser(candidates[i], userId, userEmail)) {
        return true;
      }
    }

    return false;
  }

  function principalMatchesUser(principal, userId, userEmail) {
    if (principal == null) {
      return false;
    }

    if (typeof principal === "string" || typeof principal === "number") {
      const text = pickFirst(principal);
      if (!text) {
        return false;
      }
      const asEmail = normalizeEmail(text);
      if (userEmail && asEmail && asEmail === userEmail) {
        return true;
      }
      return Boolean(userId && text === userId);
    }

    if (typeof principal !== "object") {
      return false;
    }

    const objectPrincipal = unwrapTopLevelObject(principal);
    const principalId = pickFirst(
      objectPrincipal &&
        (objectPrincipal.id ||
          objectPrincipal.userId ||
          objectPrincipal.userID ||
          objectPrincipal.ownerId ||
          objectPrincipal.createdByUserId ||
          objectPrincipal._id)
    );
    if (userId && principalId && principalId === userId) {
      return true;
    }

    const principalEmail = normalizeEmail(
      pickFirst(
        objectPrincipal &&
          (objectPrincipal.email ||
            objectPrincipal.emailId ||
            objectPrincipal.userEmail ||
            objectPrincipal.workEmail ||
            objectPrincipal.userName)
      )
    );
    if (userEmail && principalEmail && principalEmail === userEmail) {
      return true;
    }

    return false;
  }

  function unwrapTopLevelObject(value) {
    const source = asPlainObject(value);
    if (!source) {
      return value;
    }
    const candidates = [
      source.response,
      source.data,
      source.payload,
      source.result,
      source.item,
    ];
    for (let i = 0; i < candidates.length; i += 1) {
      const asObject = asPlainObject(candidates[i]);
      if (asObject) {
        return asObject;
      }
    }
    return source;
  }

  function updateDebugState(reason) {
    window.__invoiceAccessDebug = {
      reason: reason || "",
      connected: state.connected,
      access: state.access,
      ownerBasedAdmin: isAccountOwner(state.rawUser, state.rawAccount),
      inferredRoleFromRawUser: inferRole(state.rawUser, state.rawAccount, ""),
      extractedRoleLabel: extractRoleLabel(state.rawUser),
      extractedPermissionLabel: extractPermissionLabel(state.rawUser),
      syncStatus: state.syncStatus,
      sourceProjects: state.sourceProjects,
      invoiceCount: Array.isArray(state.invoices) ? state.invoices.length : 0,
      syncDiagnostics: state.syncDiagnostics,
      permissionHint: state.permissionHint,
      context: state.context,
      rawUser: state.rawUser,
      rawAccount: state.rawAccount,
      teamMembersCount: Array.isArray(state.teamMembers) ? state.teamMembers.length : 0,
      teamMembersPreview: Array.isArray(state.teamMembers) ? state.teamMembers.slice(0, 5) : [],
    };
  }

  function mergeObjects(a, b) {
    if (!a) {
      return b || null;
    }
    if (!b) {
      return a || null;
    }
    return Object.assign({}, a, b);
  }

  function createId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }
    return "id-" + Math.random().toString(16).slice(2) + "-" + Date.now().toString(16);
  }

  function createShortId() {
    return Math.random().toString(16).slice(2, 7).toUpperCase();
  }

  // Ensure PDF.js worker path is configured when available.
  if (window.pdfjsLib && window.pdfjsLib.GlobalWorkerOptions) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_CDN;
  }

  // Extract standalone emails from any string that may contain free text.
  window.__extractEmailsForInvoiceManager = function (text) {
    return dedupeEmails(String(text || "").match(EMAIL_PATTERN) || []);
  };
})();
