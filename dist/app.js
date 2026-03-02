"use strict";

(function bootstrapInvoiceAccessManager() {
  const STORAGE_PREFIX = "rocketlane-invoice-access";
  const STORAGE_VERSION = "v1";
  const LOG_LIMIT = 150;
  const PDF_WORKER_CDN =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  const EMAIL_PATTERN = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  const EMAIL_VALIDATION_PATTERN = /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i;

  const SUGGESTIONS = {
    RUNTIME_INIT_FAILED:
      "Verify the app is opened inside Rocketlane and the installation has required permissions.",
    PDF_LIB_UNAVAILABLE:
      "Allow access to the PDF.js CDN or bundle PDF.js with the app assets.",
    PDF_PARSE_FAILED:
      "Upload a text-based PDF invoice (or run OCR first) so the PM email can be extracted.",
    PDF_EMAIL_NOT_FOUND:
      "Confirm the project manager email exists in the PDF text content.",
    INVALID_IMPORT:
      "Fill all required fields and ensure PM email appears in the detected PDF emails.",
    ACCESS_DENIED:
      "Non-admin users can only import/view invoices where PM email matches their login email.",
    STORAGE_READ_FAILED:
      "Check browser storage permissions and clear corrupted local app data if needed.",
    STORAGE_WRITE_FAILED:
      "Storage quota may be full. Remove older invoices or upload smaller PDFs.",
    VIEWER_OPEN_FAILED:
      "Re-import the invoice PDF and retry opening in a new tab.",
  };

  const state = {
    client: null,
    context: null,
    rawUser: null,
    rawProject: null,
    storageKeys: null,
    invoices: [],
    logs: [],
    selectedInvoiceId: null,
    searchQuery: "",
    detectedEmails: [],
    activeTab: "invoices",
    parseToken: "",
    access: {
      role: "non_admin",
      roleLabel: "Restricted",
      isAdmin: false,
      email: "",
      displayName: "",
      canImport: false,
    },
  };

  const refs = {};

  document.addEventListener("DOMContentLoaded", () => {
    initializeApp().catch((error) => {
      console.error("Unable to initialize invoice app", error);
      if (refs.connectionBadge) {
        refs.connectionBadge.textContent = "Initialization failed";
      }
      appendLog(
        "RUNTIME_INIT_FAILED",
        "Invoice app failed to initialize.",
        error
      );
      renderLogs();
    });
  });

  async function initializeApp() {
    cacheDomReferences();
    bindEvents();

    const runtime = await initializeRuntime();
    state.client = runtime.client;
    state.context = runtime.context;
    state.rawUser = runtime.rawUser;
    state.rawProject = runtime.rawProject;
    state.storageKeys = createStorageKeys(runtime.context);
    state.access = deriveAccessProfile(runtime.rawUser, runtime.context);

    updateHeader(runtime.connected);
    configureUiForAccess();
    hydrateImportDefaults();

    loadLogs();
    loadInvoices();

    if (runtime.error) {
      appendLog(
        "RUNTIME_INIT_FAILED",
        "Rocketlane SDK was unavailable. Running in local preview mode.",
        runtime.error
      );
    }

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
    refs.importPolicyText = document.getElementById("importPolicyText");
    refs.invoiceFileInput = document.getElementById("invoiceFileInput");
    refs.invoiceTitleInput = document.getElementById("invoiceTitleInput");
    refs.projectNameInput = document.getElementById("projectNameInput");
    refs.projectIdInput = document.getElementById("projectIdInput");
    refs.pmEmailInput = document.getElementById("pmEmailInput");
    refs.detectedEmails = document.getElementById("detectedEmails");
    refs.importInvoiceButton = document.getElementById("importInvoiceButton");
    refs.importStatus = document.getElementById("importStatus");
    refs.searchInput = document.getElementById("searchInput");
    refs.visibilitySummary = document.getElementById("visibilitySummary");
    refs.invoiceStats = document.getElementById("invoiceStats");
    refs.invoiceList = document.getElementById("invoiceList");
    refs.invoiceEmptyState = document.getElementById("invoiceEmptyState");
    refs.viewerEmptyState = document.getElementById("viewerEmptyState");
    refs.viewerPanel = document.getElementById("viewerPanel");
    refs.viewerTitle = document.getElementById("viewerTitle");
    refs.viewerMeta = document.getElementById("viewerMeta");
    refs.pdfViewerFrame = document.getElementById("pdfViewerFrame");
    refs.openInNewTabButton = document.getElementById("openInNewTabButton");
    refs.logsList = document.getElementById("logsList");
    refs.logsEmptyState = document.getElementById("logsEmptyState");
    refs.clearLogsButton = document.getElementById("clearLogsButton");
  }

  function bindEvents() {
    refs.tabInvoicesButton.addEventListener("click", () => setActiveTab("invoices"));
    refs.tabLogsButton.addEventListener("click", () => setActiveTab("logs"));

    refs.invoiceFileInput.addEventListener("change", onInvoiceFileSelected);
    refs.pmEmailInput.addEventListener("blur", () => {
      refs.pmEmailInput.value = normalizeEmail(refs.pmEmailInput.value);
    });
    refs.importInvoiceButton.addEventListener("click", onImportInvoice);
    refs.searchInput.addEventListener("input", (event) => {
      state.searchQuery = String(event.target.value || "").trim().toLowerCase();
      ensureSelectedInvoice();
      renderInvoiceList();
      renderInvoiceStats();
      renderViewer();
    });
    refs.openInNewTabButton.addEventListener("click", onOpenInNewTab);
    refs.clearLogsButton.addEventListener("click", onClearLogs);
  }

  async function initializeRuntime() {
    const query = new URLSearchParams(window.location.search);
    const fallback = getContextFromQuery(query);

    if (!window.rliSdk || typeof window.rliSdk.init !== "function") {
      return {
        client: null,
        connected: false,
        context: fallback,
        rawUser: null,
        rawProject: null,
        error: null,
      };
    }

    try {
      const client = await window.rliSdk.init({});
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
        rawProject: project,
        error: null,
      };
    } catch (error) {
      console.warn("Rocketlane SDK init failed; falling back to query context.", error);
      return {
        client: null,
        connected: false,
        context: fallback,
        rawUser: null,
        rawProject: null,
        error,
      };
    }
  }

  async function safeGetClientData(client, objectName) {
    if (!client || !client.data || typeof client.data.get !== "function") {
      return null;
    }

    try {
      return await client.data.get(objectName);
    } catch (_error) {
      const aliases = {
        account: "GET_ACCOUNT_DATA",
        user: "GET_USER_DATA",
        project: "GET_PROJECT_DATA",
      };
      const identifier =
        client.data.dataIdentifiers &&
        aliases[objectName] &&
        client.data.dataIdentifiers[aliases[objectName]];

      if (!identifier) {
        return null;
      }

      try {
        return await client.data.get(identifier);
      } catch (_error2) {
        return null;
      }
    }
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
              user.userType)
        ) || fallback.userRole,
      projectId:
        pickFirst(project && (project.id || project.projectId || project._id)) ||
        fallback.projectId,
      projectName:
        pickFirst(project && (project.name || project.projectName)) ||
        fallback.projectName,
    };
  }

  function updateHeader(connected) {
    const view = state.context.projectName || "Cross-project view";
    refs.scopeText.textContent =
      "Scope: " + state.context.accountName + " / " + view + " invoices";

    if (connected) {
      refs.connectionBadge.className = "badge badge-ok";
      refs.connectionBadge.textContent = "Connected to Rocketlane";
    } else {
      refs.connectionBadge.className = "badge badge-local";
      refs.connectionBadge.textContent = "Local preview mode";
    }

    refs.roleBadge.className = "badge " + (state.access.isAdmin ? "badge-admin" : "badge-muted");
    const identity = state.access.email || state.access.displayName || "Unknown user";
    refs.roleBadge.textContent = state.access.roleLabel + " - " + identity;
  }

  function configureUiForAccess() {
    if (state.access.isAdmin) {
      refs.tabLogsButton.classList.remove("hidden");
      refs.pmEmailInput.readOnly = false;
      refs.importPolicyText.textContent =
        "Admin mode: import invoices for any project manager email that is detected in the PDF.";
      refs.importInvoiceButton.disabled = false;
    } else {
      refs.tabLogsButton.classList.add("hidden");
      refs.pmEmailInput.readOnly = true;
      if (state.access.email) {
        refs.pmEmailInput.value = state.access.email;
        refs.importPolicyText.textContent =
          "Restricted mode: you can import and view only invoices where PM email matches your Rocketlane sign-in email.";
        refs.importInvoiceButton.disabled = false;
      } else {
        refs.importPolicyText.textContent =
          "Restricted mode: no sign-in email was detected, so invoice import is disabled.";
        refs.importInvoiceButton.disabled = true;
      }
    }
  }

  function hydrateImportDefaults() {
    if (state.context.projectName) {
      refs.projectNameInput.value = state.context.projectName;
    }
    if (state.context.projectId) {
      refs.projectIdInput.value = state.context.projectId;
    }
    if (!refs.pmEmailInput.value && state.access.email) {
      refs.pmEmailInput.value = state.access.email;
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

  function createStorageKeys(context) {
    const accountScope = slug(context.accountId || context.accountName || "workspace");
    return {
      invoices: [
        STORAGE_PREFIX,
        "invoices",
        STORAGE_VERSION,
        accountScope,
      ].join(":"),
      logs: [STORAGE_PREFIX, "logs", STORAGE_VERSION, accountScope].join(":"),
    };
  }

  function loadInvoices() {
    const raw = safeStorageGet(state.storageKeys.invoices);
    if (!raw) {
      state.invoices = [];
      return;
    }

    try {
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        state.invoices = [];
        return;
      }

      state.invoices = parsed
        .map((invoice) => normalizeInvoice(invoice))
        .filter(Boolean);
    } catch (error) {
      state.invoices = [];
      appendLog(
        "STORAGE_READ_FAILED",
        "Could not parse saved invoice data. Existing local invoice cache was reset.",
        error
      );
    }
  }

  function persistInvoices() {
    try {
      safeStorageSet(state.storageKeys.invoices, JSON.stringify(state.invoices));
      return true;
    } catch (error) {
      appendLog(
        "STORAGE_WRITE_FAILED",
        "Failed to save invoice data to browser storage.",
        error
      );
      return false;
    }
  }

  function loadLogs() {
    const raw = safeStorageGet(state.storageKeys.logs);
    if (!raw) {
      state.logs = [];
      return;
    }

    try {
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        state.logs = [];
        return;
      }
      state.logs = parsed
        .map((item) => normalizeLog(item))
        .filter(Boolean)
        .slice(0, LOG_LIMIT);
    } catch (_error) {
      state.logs = [];
    }
  }

  function persistLogs() {
    try {
      safeStorageSet(state.storageKeys.logs, JSON.stringify(state.logs.slice(0, LOG_LIMIT)));
    } catch (_error) {
      // Avoid recursive logging if storage itself is unavailable.
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

    state.logs.unshift(entry);
    if (state.logs.length > LOG_LIMIT) {
      state.logs = state.logs.slice(0, LOG_LIMIT);
    }
    persistLogs();
    if (state.access.isAdmin && refs.logsList) {
      renderLogs();
    }
  }

  async function onInvoiceFileSelected() {
    const file = refs.invoiceFileInput.files && refs.invoiceFileInput.files[0];
    state.detectedEmails = [];
    renderDetectedEmails();

    if (!file) {
      setImportStatus("", null);
      return;
    }

    if (!isPdfFile(file)) {
      setImportStatus("Only PDF files are supported.", "error");
      appendLog("INVALID_IMPORT", "Rejected non-PDF file during import validation.");
      return;
    }

    if (!refs.invoiceTitleInput.value.trim()) {
      refs.invoiceTitleInput.value = file.name.replace(/\.pdf$/i, "");
    }

    const token = createId();
    state.parseToken = token;
    setImportStatus("Scanning PDF for email addresses...", null);

    try {
      const emails = await extractEmailsFromPdf(file);
      if (token !== state.parseToken) {
        return;
      }

      state.detectedEmails = emails;
      renderDetectedEmails();

      if (!emails.length) {
        setImportStatus(
          "No email addresses detected in this PDF. A PM email must be extractable.",
          "error"
        );
        appendLog(
          "PDF_EMAIL_NOT_FOUND",
          "No email addresses were found while parsing selected PDF invoice."
        );
        return;
      }

      if (!refs.pmEmailInput.value.trim()) {
        refs.pmEmailInput.value = emails[0];
      }

      setImportStatus(
        "Detected " + emails.length + " email address(es) from PDF text.",
        "success"
      );
    } catch (error) {
      if (token !== state.parseToken) {
        return;
      }
      setImportStatus("Unable to parse the PDF file.", "error");
      appendLog("PDF_PARSE_FAILED", "Unable to parse uploaded PDF invoice.", error);
    }
  }

  async function onImportInvoice() {
    const file = refs.invoiceFileInput.files && refs.invoiceFileInput.files[0];
    if (!file) {
      setImportStatus("Select a PDF invoice first.", "error");
      appendLog("INVALID_IMPORT", "Import attempted without selecting a file.");
      return;
    }

    if (!isPdfFile(file)) {
      setImportStatus("Only PDF files can be imported.", "error");
      appendLog("INVALID_IMPORT", "Non-PDF file blocked from import.");
      return;
    }

    if (!state.access.isAdmin && !state.access.email) {
      setImportStatus("Sign-in email is required for restricted imports.", "error");
      appendLog(
        "ACCESS_DENIED",
        "Import blocked because non-admin email context is unavailable."
      );
      return;
    }

    const title = String(refs.invoiceTitleInput.value || "").trim() || file.name;
    const projectName =
      String(refs.projectNameInput.value || "").trim() ||
      state.context.projectName ||
      "Unspecified project";
    const projectId =
      String(refs.projectIdInput.value || "").trim() || state.context.projectId || "";
    const pmEmail = normalizeEmail(refs.pmEmailInput.value);

    if (!isValidEmail(pmEmail)) {
      setImportStatus("Enter a valid project manager email.", "error");
      appendLog("INVALID_IMPORT", "Import blocked due to invalid PM email.");
      return;
    }

    if (!state.detectedEmails.length) {
      setImportStatus("Parsing PDF for emails before import...", null);
      try {
        state.detectedEmails = await extractEmailsFromPdf(file);
        renderDetectedEmails();
      } catch (error) {
        setImportStatus("Unable to parse the PDF file.", "error");
        appendLog("PDF_PARSE_FAILED", "PDF parsing failed during import.", error);
        return;
      }
    }

    if (!state.detectedEmails.length) {
      setImportStatus(
        "No email addresses detected in PDF. Import cancelled.",
        "error"
      );
      appendLog(
        "PDF_EMAIL_NOT_FOUND",
        "Invoice import failed because no PM email could be extracted from PDF."
      );
      return;
    }

    const detectedSet = new Set(state.detectedEmails.map(normalizeEmail));
    if (!detectedSet.has(pmEmail)) {
      setImportStatus(
        "PM email must be present in the PDF's detected email list.",
        "error"
      );
      appendLog(
        "INVALID_IMPORT",
        "PM email did not match detected PDF emails during import."
      );
      return;
    }

    if (!state.access.isAdmin && pmEmail !== state.access.email) {
      setImportStatus(
        "Restricted users can only import invoices assigned to their own email.",
        "error"
      );
      appendLog(
        "ACCESS_DENIED",
        "Non-admin import blocked because PM email did not match signed-in email."
      );
      return;
    }

    setImportStatus("Importing invoice...", null);

    try {
      const pdfDataUrl = await readFileAsDataUrl(file);
      const now = new Date().toISOString();
      const invoice = normalizeInvoice({
        id: createId(),
        title,
        fileName: file.name,
        projectName,
        projectId,
        projectManagerEmail: pmEmail,
        detectedEmails: Array.from(detectedSet),
        uploadedByEmail: state.access.email,
        uploadedByName: state.access.displayName,
        uploadedAt: now,
        sizeBytes: Number(file.size || 0),
        mimeType: file.type || "application/pdf",
        pdfDataUrl,
      });

      if (!invoice) {
        throw new Error("Failed to normalize invoice payload.");
      }

      state.invoices.unshift(invoice);
      if (!persistInvoices()) {
        state.invoices = state.invoices.filter((item) => item.id !== invoice.id);
        setImportStatus("Import failed while saving invoice.", "error");
        return;
      }

      state.selectedInvoiceId = invoice.id;
      resetImportInputs();
      setImportStatus("Invoice imported successfully.", "success");
      renderAll();
    } catch (error) {
      setImportStatus("Import failed. Please retry with a valid PDF invoice.", "error");
      appendLog("STORAGE_WRITE_FAILED", "Invoice import failed.", error);
    }
  }

  function resetImportInputs() {
    refs.invoiceFileInput.value = "";
    refs.invoiceTitleInput.value = "";
    state.detectedEmails = [];
    renderDetectedEmails();

    if (state.access.isAdmin) {
      refs.pmEmailInput.value = "";
    } else {
      refs.pmEmailInput.value = state.access.email;
    }

    if (state.context.projectName && !refs.projectNameInput.value.trim()) {
      refs.projectNameInput.value = state.context.projectName;
    }
    if (state.context.projectId && !refs.projectIdInput.value.trim()) {
      refs.projectIdInput.value = state.context.projectId;
    }
  }

  function setImportStatus(text, tone) {
    refs.importStatus.textContent = text || "";
    refs.importStatus.classList.remove("error", "success");
    if (tone === "error") {
      refs.importStatus.classList.add("error");
    } else if (tone === "success") {
      refs.importStatus.classList.add("success");
    }
  }

  async function extractEmailsFromPdf(file) {
    ensurePdfJsAvailable();
    const text = await extractTextFromPdf(file);
    return extractEmailsFromText(text);
  }

  function ensurePdfJsAvailable() {
    if (!window.pdfjsLib || typeof window.pdfjsLib.getDocument !== "function") {
      appendLog(
        "PDF_LIB_UNAVAILABLE",
        "PDF.js was not loaded, so PDF parsing cannot proceed."
      );
      throw new Error("PDF.js library is not available.");
    }
    if (window.pdfjsLib.GlobalWorkerOptions) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_CDN;
    }
  }

  async function extractTextFromPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    let loadingTask;
    try {
      loadingTask = window.pdfjsLib.getDocument({ data: arrayBuffer });
      const pdfDocument = await loadingTask.promise;
      const parts = [];

      for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
        const page = await pdfDocument.getPage(pageNumber);
        const content = await page.getTextContent();
        const pageText = content.items.map((item) => item.str).join(" ");
        parts.push(pageText);
      }

      if (typeof pdfDocument.cleanup === "function") {
        pdfDocument.cleanup();
      }

      return parts.join("\n");
    } catch (error) {
      appendLog("PDF_PARSE_FAILED", "PDF text extraction failed.", error);
      throw error;
    } finally {
      if (loadingTask && typeof loadingTask.destroy === "function") {
        loadingTask.destroy();
      }
    }
  }

  function extractEmailsFromText(text) {
    const matches = String(text || "").match(EMAIL_PATTERN) || [];
    const deduped = [];
    const seen = new Set();
    matches.forEach((value) => {
      const email = normalizeEmail(value);
      if (!isValidEmail(email) || seen.has(email)) {
        return;
      }
      seen.add(email);
      deduped.push(email);
    });
    return deduped;
  }

  function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(reader.error || new Error("File read failed."));
      reader.readAsDataURL(file);
    });
  }

  function renderAll() {
    renderVisibilitySummary();
    renderInvoiceStats();
    renderInvoiceList();
    renderViewer();
    renderLogs();
  }

  function renderVisibilitySummary() {
    if (state.access.isAdmin) {
      refs.visibilitySummary.textContent =
        "Admin access: all imported invoices are visible.";
      return;
    }

    if (state.access.email) {
      refs.visibilitySummary.textContent =
        "Restricted access: only invoices where PM email = " +
        state.access.email +
        " are visible.";
      return;
    }

    refs.visibilitySummary.textContent =
      "Restricted access: sign-in email not detected, so no invoices are visible.";
  }

  function renderInvoiceStats() {
    const visible = getVisibleInvoices().length;
    const total = state.invoices.length;
    refs.invoiceStats.textContent = visible + " visible / " + total + " total";
  }

  function renderInvoiceList() {
    const visible = getVisibleInvoices();
    refs.invoiceList.innerHTML = "";

    if (!visible.length) {
      refs.invoiceEmptyState.classList.remove("hidden");
      return;
    }
    refs.invoiceEmptyState.classList.add("hidden");

    visible.forEach((invoice) => {
      const item = document.createElement("article");
      item.className = "invoice-item";
      if (invoice.id === state.selectedInvoiceId) {
        item.classList.add("selected");
      }
      item.setAttribute("role", "button");
      item.setAttribute("tabindex", "0");
      item.setAttribute("aria-label", "Open invoice " + invoice.title);

      const title = document.createElement("p");
      title.className = "invoice-item-title";
      title.textContent = invoice.title;

      const meta = document.createElement("p");
      meta.className = "invoice-item-meta";
      meta.textContent =
        "Project: " +
        invoice.projectName +
        " | PM: " +
        invoice.projectManagerEmail +
        " | Uploaded " +
        formatTime(invoice.uploadedAt);

      item.appendChild(title);
      item.appendChild(meta);

      item.addEventListener("click", () => {
        state.selectedInvoiceId = invoice.id;
        renderInvoiceList();
        renderViewer();
      });
      item.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          state.selectedInvoiceId = invoice.id;
          renderInvoiceList();
          renderViewer();
        }
      });

      refs.invoiceList.appendChild(item);
    });
  }

  function renderViewer() {
    const invoice = getSelectedVisibleInvoice();
    if (!invoice) {
      refs.viewerPanel.classList.add("hidden");
      refs.viewerEmptyState.classList.remove("hidden");
      refs.pdfViewerFrame.removeAttribute("src");
      return;
    }

    refs.viewerEmptyState.classList.add("hidden");
    refs.viewerPanel.classList.remove("hidden");
    refs.viewerTitle.textContent = invoice.title;
    refs.viewerMeta.textContent =
      "File: " +
      invoice.fileName +
      " | Project: " +
      invoice.projectName +
      " | PM: " +
      invoice.projectManagerEmail +
      " | Uploaded by: " +
      (invoice.uploadedByEmail || "Unknown") +
      " | Size: " +
      formatBytes(invoice.sizeBytes);

    if (refs.pdfViewerFrame.getAttribute("src") !== invoice.pdfDataUrl) {
      refs.pdfViewerFrame.setAttribute("src", invoice.pdfDataUrl);
    }
  }

  function onOpenInNewTab() {
    const invoice = getSelectedVisibleInvoice();
    if (!invoice) {
      return;
    }

    try {
      window.open(invoice.pdfDataUrl, "_blank", "noopener,noreferrer");
    } catch (error) {
      appendLog(
        "VIEWER_OPEN_FAILED",
        "Could not open selected invoice in a new tab.",
        error
      );
    }
  }

  function renderDetectedEmails() {
    refs.detectedEmails.innerHTML = "";
    if (!state.detectedEmails.length) {
      const empty = document.createElement("span");
      empty.className = "chip";
      empty.textContent = "No emails detected yet";
      refs.detectedEmails.appendChild(empty);
      return;
    }

    state.detectedEmails.forEach((email) => {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = email;
      refs.detectedEmails.appendChild(chip);
    });
  }

  function renderLogs() {
    if (!state.access.isAdmin) {
      refs.logsList.innerHTML = "";
      refs.logsEmptyState.classList.add("hidden");
      return;
    }

    refs.logsList.innerHTML = "";
    if (!state.logs.length) {
      refs.logsEmptyState.classList.remove("hidden");
      return;
    }
    refs.logsEmptyState.classList.add("hidden");

    state.logs.forEach((entry) => {
      const item = document.createElement("article");
      item.className = "log-item";

      const header = document.createElement("div");
      header.className = "log-item-header";

      const code = document.createElement("span");
      code.className = "log-code";
      code.textContent = entry.code;

      const time = document.createElement("span");
      time.className = "log-time";
      time.textContent = formatTime(entry.createdAt);

      header.appendChild(code);
      header.appendChild(time);

      const message = document.createElement("p");
      message.className = "log-message";
      message.textContent = entry.message;

      const suggestion = document.createElement("p");
      suggestion.className = "log-suggestion";
      suggestion.textContent = "Suggested fix: " + entry.suggestion;

      item.appendChild(header);
      item.appendChild(message);
      if (entry.details) {
        const detail = document.createElement("p");
        detail.className = "log-suggestion";
        detail.textContent = "Details: " + entry.details;
        item.appendChild(detail);
      }
      item.appendChild(suggestion);
      refs.logsList.appendChild(item);
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

  function getVisibleInvoices() {
    const query = state.searchQuery;
    let invoices = state.invoices.slice();

    if (!state.access.isAdmin) {
      const currentEmail = state.access.email;
      if (!currentEmail) {
        return [];
      }
      invoices = invoices.filter(
        (invoice) => normalizeEmail(invoice.projectManagerEmail) === currentEmail
      );
    }

    invoices.sort((a, b) => timestampValue(b.uploadedAt) - timestampValue(a.uploadedAt));

    if (!query) {
      return invoices;
    }

    return invoices.filter((invoice) => {
      const haystack = (
        invoice.title +
        " " +
        invoice.fileName +
        " " +
        invoice.projectName +
        " " +
        invoice.projectManagerEmail +
        " " +
        invoice.detectedEmails.join(" ")
      ).toLowerCase();
      return haystack.includes(query);
    });
  }

  function ensureSelectedInvoice() {
    const visible = getVisibleInvoices();
    if (!visible.length) {
      state.selectedInvoiceId = null;
      return;
    }

    const hasSelected = visible.some((item) => item.id === state.selectedInvoiceId);
    if (!hasSelected) {
      state.selectedInvoiceId = visible[0].id;
    }
  }

  function getSelectedVisibleInvoice() {
    const visible = getVisibleInvoices();
    return visible.find((item) => item.id === state.selectedInvoiceId) || null;
  }

  function deriveAccessProfile(rawUser, context) {
    const displayName =
      pickFirst(
        rawUser &&
          (rawUser.name || rawUser.fullName || rawUser.displayName || rawUser.email)
      ) || context.userName;
    const email = normalizeEmail(extractPrimaryEmail(rawUser) || context.userEmail || "");
    const role = inferRole(rawUser, context.userRole);
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
      canImport: isAdmin || Boolean(email),
    };
  }

  function inferRole(user, fallbackRole) {
    const normalizedFallback = normalizeRole(fallbackRole);
    if (normalizedFallback) {
      return normalizedFallback;
    }

    if (!user || typeof user !== "object") {
      return "non_admin";
    }

    if (user.isAdmin === true || user.admin === true) {
      return "admin";
    }

    const tokens = [];
    collectRoleTokens(user, tokens, 0);
    const haystack = tokens.join(" ").toLowerCase();

    if (/(^|\b)admin(istrator)?(\b|$)/.test(haystack)) {
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
      const lowerKey = key.toLowerCase();
      const nestedValue = value[key];
      if (
        lowerKey.includes("role") ||
        lowerKey.includes("permission") ||
        lowerKey.includes("type") ||
        lowerKey.includes("group")
      ) {
        collectRoleTokens(nestedValue, target, depth + 1);
      }
    });
  }

  function extractPrimaryEmail(user) {
    if (!user || typeof user !== "object") {
      return "";
    }

    const candidates = [];
    collectEmailsFromObject(user, candidates, 0);

    for (let i = 0; i < candidates.length; i += 1) {
      const email = normalizeEmail(candidates[i]);
      if (isValidEmail(email)) {
        return email;
      }
    }

    return "";
  }

  function collectEmailsFromObject(value, target, depth) {
    if (depth > 6 || value == null) {
      return;
    }

    if (typeof value === "string") {
      if (isValidEmail(value.trim())) {
        target.push(value.trim());
      }
      return;
    }

    if (Array.isArray(value)) {
      value.forEach((item) => collectEmailsFromObject(item, target, depth + 1));
      return;
    }

    if (typeof value !== "object") {
      return;
    }

    Object.keys(value).forEach((key) => {
      const nestedValue = value[key];
      if (String(key).toLowerCase().includes("email")) {
        if (typeof nestedValue === "string") {
          target.push(nestedValue);
        } else {
          collectEmailsFromObject(nestedValue, target, depth + 1);
        }
      }
    });
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

  function normalizeInvoice(invoice) {
    if (!invoice || typeof invoice !== "object") {
      return null;
    }

    const projectManagerEmail = normalizeEmail(invoice.projectManagerEmail);
    const pdfDataUrl = String(invoice.pdfDataUrl || "");
    if (!projectManagerEmail || !pdfDataUrl) {
      return null;
    }

    const detected = Array.isArray(invoice.detectedEmails)
      ? invoice.detectedEmails.map(normalizeEmail).filter(isValidEmail)
      : [];

    return {
      id: String(invoice.id || createId()),
      title: String(invoice.title || invoice.fileName || "Untitled invoice").trim(),
      fileName: String(invoice.fileName || "invoice.pdf").trim(),
      projectName: String(invoice.projectName || "Unspecified project").trim(),
      projectId: String(invoice.projectId || "").trim(),
      projectManagerEmail,
      detectedEmails: dedupeEmails(detected),
      uploadedByEmail: normalizeEmail(invoice.uploadedByEmail),
      uploadedByName: String(invoice.uploadedByName || "").trim(),
      uploadedAt: String(invoice.uploadedAt || new Date().toISOString()),
      sizeBytes: Number(invoice.sizeBytes || 0),
      mimeType: String(invoice.mimeType || "application/pdf"),
      pdfDataUrl,
    };
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

  function isPdfFile(file) {
    if (!file) {
      return false;
    }
    const type = String(file.type || "").toLowerCase();
    const name = String(file.name || "").toLowerCase();
    return type.includes("pdf") || name.endsWith(".pdf");
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
    } catch (error) {
      appendLog("STORAGE_READ_FAILED", "Unable to read from browser storage.", error);
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
      return "Unknown error object";
    }
  }

  function formatTime(isoString) {
    const value = new Date(isoString).getTime();
    if (Number.isNaN(value)) {
      return "just now";
    }
    return new Date(value).toLocaleString([], {
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function timestampValue(isoString) {
    const value = new Date(isoString).getTime();
    return Number.isNaN(value) ? 0 : value;
  }

  function formatBytes(value) {
    const bytes = Number(value || 0);
    if (!Number.isFinite(bytes) || bytes <= 0) {
      return "0 B";
    }
    if (bytes < 1024) {
      return bytes + " B";
    }
    if (bytes < 1024 * 1024) {
      return (bytes / 1024).toFixed(1) + " KB";
    }
    return (bytes / (1024 * 1024)).toFixed(2) + " MB";
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

  function createId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }
    return "id-" + Math.random().toString(16).slice(2) + "-" + Date.now().toString(16);
  }
})();
