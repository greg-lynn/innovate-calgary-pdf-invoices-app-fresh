"use strict";

(function bootstrapInvoiceAccessManager() {
  const STORAGE_PREFIX = "rocketlane-invoice-access";
  const STORAGE_VERSION = "v2";
  const LOG_LIMIT = 200;
  const SDK_WAIT_MS = 4500;
  const SDK_POLL_INTERVAL_MS = 120;
  const PDF_WORKER_CDN =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  const SOURCE_PROJECT_NAMES = [
    "expert advisor program invoices",
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
    SOURCE_INVOICES_NOT_FOUND:
      "Add PDF invoice files/documents to the source project to populate this table.",
    PDF_LIB_UNAVAILABLE:
      "Allow access to PDF.js or bundle PDF.js assets with this app package.",
    PDF_PREVIEW_FAILED:
      "Verify the invoice file URL is valid and accessible for this user session.",
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
    selectedInvoiceId: null,
    searchQuery: "",
    activeTab: "invoices",
    syncStatus: "Initializing...",
    access: {
      role: "non_admin",
      roleLabel: "Restricted",
      isAdmin: false,
      email: "",
      displayName: "",
    },
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

    if (runtime.connected) {
      const apiUser = await fetchCurrentUserProfile();
      if (apiUser && typeof apiUser === "object") {
        state.rawUser = mergeObjects(state.rawUser, apiUser);
      }
    }

    state.access = deriveAccessProfile(state.rawUser, state.rawAccount, runtime.context);

    updateHeader();
    configureUiForAccess();
    await refreshInvoicesFromSource();
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
    refs.syncStatus = document.getElementById("syncStatus");
    refs.searchInput = document.getElementById("searchInput");
    refs.visibilitySummary = document.getElementById("visibilitySummary");
    refs.invoiceStats = document.getElementById("invoiceStats");
    refs.invoiceTableBody = document.getElementById("invoiceTableBody");
    refs.invoiceEmptyState = document.getElementById("invoiceEmptyState");
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
    refs.closeModalButton = document.getElementById("closeModalButton");
  }

  function bindEvents() {
    refs.tabInvoicesButton.addEventListener("click", () => setActiveTab("invoices"));
    refs.tabLogsButton.addEventListener("click", () => setActiveTab("logs"));
    refs.searchInput.addEventListener("input", (event) => {
      state.searchQuery = String(event.target.value || "").trim().toLowerCase();
      ensureSelectedInvoice();
      renderInvoiceTable();
      renderInvoiceStats();
      renderSelectedSummary();
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
        return value != null ? value : null;
      } catch (_error) {
        return null;
      }
    };

    try {
      const direct = await client.data.get(objectName);
      if (direct != null) {
        return direct;
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
    return {
      accountId:
        pickFirst(account && (account.id || account.accountId || account._id)) ||
        fallback.accountId,
      accountName:
        pickFirst(
          account &&
            (account.name ||
              account.accountName ||
              account.displayName ||
              account.companyName)
        ) || fallback.accountName,
      userId:
        pickFirst(user && (user.id || user.userId || user._id)) || fallback.userId,
      userName:
        pickFirst(
          user && (user.name || user.fullName || user.displayName || user.email)
        ) || fallback.userName,
      userEmail:
        normalizeEmail(
          pickFirst(
            user &&
              (user.email ||
                user.workEmail ||
                user.userEmail ||
                (user.profile && user.profile.email))
          ) || fallback.userEmail
        ),
      userRole:
        pickFirst(
          user &&
            (user.role ||
              user.userRole ||
              user.permissionSet ||
              user.permissionLevel ||
              user.userType ||
              user.accountPermission)
        ) || fallback.userRole,
      projectId:
        pickFirst(project && (project.id || project.projectId || project._id)) ||
        fallback.projectId,
      projectName:
        pickFirst(project && (project.name || project.projectName)) ||
        fallback.projectName,
    };
  }

  async function fetchCurrentUserProfile() {
    const endpoints = ["/api/1.0/users/me", "/api/1.0/me", "/api/1.0/users/current"];
    for (let i = 0; i < endpoints.length; i += 1) {
      try {
        const payload = await requestJson(endpoints[i]);
        if (payload && typeof payload === "object") {
          if (payload.data && typeof payload.data === "object" && !Array.isArray(payload.data)) {
            return payload.data;
          }
          return payload;
        }
      } catch (_error) {
        // Try next endpoint.
      }
    }
    return null;
  }

  async function refreshInvoicesFromSource() {
    state.syncStatus = "Fetching invoices from source projects...";
    renderSyncStatus();

    let invoices = [];
    try {
      if (state.connected) {
        invoices = await fetchInvoicesFromSourceProjects();
      }
    } catch (error) {
      appendLog(
        "SOURCE_FETCH_FAILED",
        "Unable to fetch invoices from source projects.",
        error
      );
    }

    if (!invoices.length) {
      state.sourceProjects = [];
      invoices = [createSampleInvoice()];
      state.syncStatus =
        "No source project found yet. Showing the sample invoice entry for preview.";
      appendLog(
        "SOURCE_PROJECTS_NOT_FOUND",
        "No source-of-truth project was discovered. Added a sample invoice row."
      );
    } else {
      state.syncStatus =
        "Loaded " +
        invoices.length +
        " invoice(s) from source project(s): " +
        state.sourceProjects.join(", ");
    }

    state.invoices = invoices.map(normalizeInvoice).filter(Boolean);
  }

  async function fetchInvoicesFromSourceProjects() {
    const projects = await fetchSourceProjects();
    state.sourceProjects = projects.map((item) => item.name);

    if (!projects.length) {
      return [];
    }

    const invoices = [];
    for (let i = 0; i < projects.length; i += 1) {
      const perProject = await fetchInvoicesForProject(projects[i]);
      invoices.push(...perProject);
    }

    return dedupeInvoices(invoices);
  }

  async function fetchSourceProjects() {
    const projects = [];
    const byKey = new Set();

    const runtimeProject = normalizeProjectRecord(state.rawProject);
    if (
      runtimeProject &&
      SOURCE_PROJECT_NAMES.includes(runtimeProject.name.toLowerCase())
    ) {
      const key = runtimeProject.id || runtimeProject.name.toLowerCase();
      byKey.add(key);
      projects.push(runtimeProject);
    }

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
      if (!SOURCE_PROJECT_NAMES.includes(project.name.toLowerCase())) {
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
    if (!projectId) {
      return [];
    }

    const endpoints = [
      "/api/1.0/projects/" + encodeURIComponent(projectId) + "/documents?size=200",
      "/api/1.0/projects/" + encodeURIComponent(projectId) + "/documents",
      "/api/1.0/projects/" + encodeURIComponent(projectId) + "/files?size=200",
      "/api/1.0/projects/" + encodeURIComponent(projectId) + "/files",
      "/api/1.0/documents?projectId=" + encodeURIComponent(projectId) + "&size=200",
      "/api/1.0/files?projectId=" + encodeURIComponent(projectId) + "&size=200",
      "/api/1.0/tasks?projectId=" + encodeURIComponent(projectId) + "&size=200",
    ];

    const records = await requestCollection(endpoints, [
      "documents",
      "files",
      "tasks",
      "data",
      "content",
      "results",
      "items",
    ]);

    const invoices = [];
    for (let i = 0; i < records.length; i += 1) {
      const candidates = extractPdfCandidates(records[i]);
      candidates.forEach((candidate) => {
        const invoice = buildInvoiceFromCandidate(candidate, project);
        if (invoice) {
          invoices.push(invoice);
        }
      });
    }

    if (!invoices.length) {
      appendLog(
        "SOURCE_INVOICES_NOT_FOUND",
        'No PDF invoice documents were found under source project "' +
          project.name +
          '".'
      );
    }

    return invoices;
  }

  async function requestCollection(endpoints, preferredKeys) {
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

  function looksLikePdfNode(node) {
    const mime = String(
      node.mimeType || node.contentType || node.fileType || node.type || ""
    ).toLowerCase();
    const name = String(node.name || node.fileName || node.title || "").toLowerCase();
    const url = String(
      node.url || node.fileUrl || node.downloadUrl || node.signedUrl || node.href || ""
    ).toLowerCase();

    if (mime.includes("pdf")) {
      return true;
    }
    if (name.endsWith(".pdf")) {
      return true;
    }
    if (url.includes(".pdf")) {
      return true;
    }
    return false;
  }

  function buildInvoiceFromCandidate(node, project) {
    const pdfUrlRaw = String(
      node.signedUrl || node.downloadUrl || node.fileUrl || node.url || node.href || ""
    ).trim();
    if (!pdfUrlRaw) {
      return null;
    }

    const pdfUrl = toAbsoluteUrl(pdfUrlRaw);
    const invoiceName =
      pickFirst(node.name || node.fileName || node.title) || "Invoice PDF";
    const invoiceNumber =
      pickFirst(
        node.invoiceNumber ||
          node.number ||
          node.docNumber ||
          node.referenceNumber ||
          extractInvoiceNumberFromText(invoiceName)
      ) || "INV-" + createShortId();
    const invoiceDate =
      pickFirst(
        node.invoiceDate ||
          node.date ||
          node.issuedOn ||
          node.issuedDate ||
          node.createdAt ||
          node.updatedAt
      ) || new Date().toISOString();

    const nodeContacts = extractContacts(node);
    const ownerName = nodeContacts.names[0] || project.ownerName || "Unassigned";
    const associatedEmails = dedupeEmails(
      nodeContacts.emails.concat(project.ownerEmails || [])
    );

    return normalizeInvoice({
      id:
        pickFirst(node.id || node._id || node.fileId || node.documentId) ||
        createId(),
      invoiceNumber,
      invoiceName,
      ownerName,
      accountName: project.accountName || state.context.accountName || "Rocketlane Account",
      invoiceDate,
      pdfUrl,
      associatedEmails,
      sourceProjectName: project.name,
    });
  }

  function createSampleInvoice() {
    const fallbackEmail = state.access.email || "sample@rocketlane.com";
    return normalizeInvoice({
      id: "sample-invoice-preview",
      invoiceNumber: "INV-0001-SAMPLE",
      invoiceName: "Sample invoice preview",
      ownerName: "Sample Expert Advisor",
      accountName: state.context.accountName || "Rocketlane Workspace",
      invoiceDate: "2026-03-03T00:00:00.000Z",
      pdfUrl: SAMPLE_PDF_DATA_URL,
      associatedEmails: [fallbackEmail, "sample@rocketlane.com"],
      sourceProjectName: "Sample (fallback)",
    });
  }

  function normalizeInvoice(invoice) {
    if (!invoice || typeof invoice !== "object") {
      return null;
    }
    const pdfUrl = String(invoice.pdfUrl || "").trim();
    if (!pdfUrl) {
      return null;
    }

    return {
      id: String(invoice.id || createId()),
      invoiceNumber: String(invoice.invoiceNumber || "Unknown").trim(),
      invoiceName: String(invoice.invoiceName || "Untitled invoice").trim(),
      ownerName: String(invoice.ownerName || "Unassigned").trim(),
      accountName: String(invoice.accountName || "Rocketlane Account").trim(),
      invoiceDate: String(invoice.invoiceDate || new Date().toISOString()),
      pdfUrl,
      associatedEmails: dedupeEmails(invoice.associatedEmails || []),
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
      (a, b) => timestampValue(b.invoiceDate) - timestampValue(a.invoiceDate)
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
    refs.tabLogsButton.classList.remove("hidden");
    if (state.access.isAdmin) {
      refs.tabLogsButton.textContent = "Diagnostics";
      refs.tabLogsButton.classList.remove("locked");
      refs.clearLogsButton.classList.remove("hidden");
      refs.logsInfoText.textContent =
        "Admin users can view full diagnostics, error context, and remediation suggestions.";
    } else {
      refs.tabLogsButton.textContent = "Diagnostics (Admin)";
      refs.tabLogsButton.classList.add("locked");
      refs.clearLogsButton.classList.add("hidden");
      refs.logsInfoText.textContent =
        "Diagnostics details are admin-only. This tab still provides access guidance.";
    }
  }

  function setActiveTab(tab) {
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
    renderInvoiceStats();
    renderInvoiceTable();
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
    const visible = getVisibleInvoices().length;
    refs.invoiceStats.textContent =
      visible + " visible / " + state.invoices.length + " total";
  }

  function renderInvoiceTable() {
    const visible = getVisibleInvoices();
    refs.invoiceTableBody.innerHTML = "";

    if (!visible.length) {
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

      const nameCell = document.createElement("td");
      nameCell.textContent = invoice.invoiceName;

      const ownerCell = document.createElement("td");
      ownerCell.textContent = invoice.ownerName;

      const accountCell = document.createElement("td");
      accountCell.textContent = invoice.accountName;

      const dateCell = document.createElement("td");
      dateCell.textContent = formatDate(invoice.invoiceDate);

      row.appendChild(numberCell);
      row.appendChild(nameCell);
      row.appendChild(ownerCell);
      row.appendChild(accountCell);
      row.appendChild(dateCell);
      refs.invoiceTableBody.appendChild(row);
    });
  }

  function renderSelectedSummary() {
    const invoice = getSelectedVisibleInvoice();
    if (!invoice) {
      refs.selectedInvoiceSummary.textContent =
        "Click an invoice number to preview its PDF.";
      return;
    }

    refs.selectedInvoiceSummary.textContent =
      invoice.invoiceNumber +
      " · " +
      invoice.invoiceName +
      " · " +
      invoice.ownerName +
      " · " +
      formatDate(invoice.invoiceDate);
  }

  function renderSourceProjects() {
    if (!state.sourceProjects.length) {
      refs.sourceProjectsText.textContent =
        "No source project found yet. Sample invoice is shown for preview.";
      return;
    }
    refs.sourceProjectsText.textContent = state.sourceProjects.join(", ");
  }

  function openPdfModal(invoice) {
    if (!invoice || !invoice.pdfUrl) {
      appendLog("PDF_PREVIEW_FAILED", "Attempted to preview an invoice without a PDF URL.");
      return;
    }
    refs.modalTitle.textContent = invoice.invoiceNumber + " · " + invoice.invoiceName;
    refs.modalPdfFrame.setAttribute("src", invoice.pdfUrl);
    refs.pdfModal.classList.remove("hidden");
    refs.pdfModal.setAttribute("aria-hidden", "false");
  }

  function closePdfModal() {
    refs.pdfModal.classList.add("hidden");
    refs.pdfModal.setAttribute("aria-hidden", "true");
    refs.modalPdfFrame.removeAttribute("src");
  }

  function getVisibleInvoices() {
    let invoices = state.invoices.slice();

    if (!state.access.isAdmin) {
      const email = state.access.email;
      if (!email) {
        return [];
      }
      invoices = invoices.filter((invoice) =>
        invoice.associatedEmails.includes(email)
      );
    }

    if (state.searchQuery) {
      const search = state.searchQuery;
      invoices = invoices.filter((invoice) => {
        const haystack = (
          invoice.invoiceNumber +
          " " +
          invoice.invoiceName +
          " " +
          invoice.ownerName +
          " " +
          invoice.accountName +
          " " +
          invoice.invoiceDate
        ).toLowerCase();
        return haystack.includes(search);
      });
    }

    invoices.sort(
      (a, b) => timestampValue(b.invoiceDate) - timestampValue(a.invoiceDate)
    );
    return invoices;
  }

  function ensureSelectedInvoice() {
    const visible = getVisibleInvoices();
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
    const visible = getVisibleInvoices();
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

  function deriveAccessProfile(rawUser, rawAccount, context) {
    const displayName =
      pickFirst(
        rawUser &&
          (rawUser.name || rawUser.fullName || rawUser.displayName || rawUser.email)
      ) || context.userName;
    const email = normalizeEmail(extractPrimaryEmail(rawUser) || context.userEmail || "");
    const role = inferRole(rawUser, rawAccount, context.userRole);
    const isAdmin = role === "admin";
    const roleLabel =
      role === "admin"
        ? "Admin"
        : role === "collaborator"
          ? "Collaborator"
          : role === "expert_advisor"
            ? "Expert Advisor"
            : "Restricted";

    return {
      role,
      roleLabel,
      isAdmin,
      email,
      displayName,
    };
  }

  function inferRole(user, account, fallbackRole) {
    const normalizedFallback = normalizeRole(fallbackRole);
    if (normalizedFallback) {
      return normalizedFallback;
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
    const haystack = tokens.join(" ").toLowerCase();

    if (/(^|\b)(account|workspace)?\s*admin(istrator)?(\b|$)/.test(haystack)) {
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

    Object.keys(value).forEach((key) => {
      const lowered = key.toLowerCase();
      const next = value[key];
      if (
        lowered.includes("role") ||
        lowered.includes("permission") ||
        lowered.includes("type") ||
        lowered.includes("group") ||
        lowered.includes("admin")
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
    const text = String(value || "").trim().toLowerCase();
    if (!text) {
      return "";
    }
    if (text.includes("admin")) {
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
