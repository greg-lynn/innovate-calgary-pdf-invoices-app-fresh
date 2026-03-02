"use strict";

(function bootstrapBulletinBoard() {
  const STORAGE_PREFIX = "rocketlane-bulletin-board";
  const NOTE_COLORS = [
    { id: "yellow", label: "Sunny", hex: "#fff8a6" },
    { id: "blue", label: "Ocean", hex: "#cce8ff" },
    { id: "green", label: "Mint", hex: "#d8f5d2" },
    { id: "pink", label: "Rose", hex: "#ffd7ea" },
    { id: "purple", label: "Lavender", hex: "#ecd8ff" },
  ];

  const state = {
    client: null,
    context: null,
    storageKey: "",
    notes: [],
    selectedNoteId: null,
    searchQuery: "",
    filter: "all",
    saveTimer: null,
  };

  const refs = {};

  document.addEventListener("DOMContentLoaded", () => {
    initializeApp().catch((error) => {
      console.error("Unable to initialize app", error);
      if (refs.connectionBadge) {
        refs.connectionBadge.textContent = "Initialization failed";
      }
    });
  });

  async function initializeApp() {
    cacheDomReferences();
    bindEvents();
    buildColorPicker();

    const runtime = await initializeRuntime();
    state.client = runtime.client;
    state.context = runtime.context;
    state.storageKey = createStorageKey(runtime.context);

    updateHeader(runtime);
    loadNotes();

    const query = new URLSearchParams(window.location.search);
    if (state.notes.length === 0 && shouldSeedDemo(query)) {
      seedDemoNotes();
    }

    if (state.notes.length === 0) {
      createDefaultWelcomeNote();
    }

    if (!state.selectedNoteId && state.notes.length > 0) {
      state.selectedNoteId = state.notes[0].id;
    }

    renderAll();
  }

  function cacheDomReferences() {
    refs.boardScope = document.getElementById("boardScope");
    refs.connectionBadge = document.getElementById("connectionBadge");
    refs.addNoteButton = document.getElementById("addNoteButton");
    refs.seedDemoButton = document.getElementById("seedDemoButton");
    refs.clearBoardButton = document.getElementById("clearBoardButton");
    refs.searchInput = document.getElementById("searchInput");
    refs.filterSelect = document.getElementById("filterSelect");
    refs.boardStats = document.getElementById("boardStats");
    refs.notesGrid = document.getElementById("notesGrid");
    refs.notesEmptyState = document.getElementById("notesEmptyState");
    refs.editorEmptyState = document.getElementById("editorEmptyState");
    refs.editorPanel = document.getElementById("editorPanel");
    refs.noteTitleInput = document.getElementById("noteTitleInput");
    refs.noteBodyInput = document.getElementById("noteBodyInput");
    refs.colorPicker = document.getElementById("colorPicker");
    refs.saveState = document.getElementById("saveState");
    refs.pinNoteButton = document.getElementById("pinNoteButton");
    refs.deleteNoteButton = document.getElementById("deleteNoteButton");
  }

  function bindEvents() {
    refs.addNoteButton.addEventListener("click", onAddNote);
    refs.seedDemoButton.addEventListener("click", () => {
      seedDemoNotes();
      renderAll();
    });
    refs.clearBoardButton.addEventListener("click", onClearBoard);
    refs.searchInput.addEventListener("input", (event) => {
      state.searchQuery = event.target.value.trim();
      renderNotesGrid();
      renderStats();
    });
    refs.filterSelect.addEventListener("change", (event) => {
      state.filter = event.target.value;
      renderNotesGrid();
      renderStats();
    });

    refs.noteTitleInput.addEventListener("input", onTitleInput);
    refs.noteBodyInput.addEventListener("input", onBodyInput);
    refs.noteBodyInput.addEventListener("paste", onEditorPaste);

    refs.pinNoteButton.addEventListener("click", onTogglePinFromEditor);
    refs.deleteNoteButton.addEventListener("click", onDeleteFromEditor);

    refs.colorPicker.addEventListener("click", onColorSwatchClick);

    document.querySelectorAll(".format-button").forEach((button) => {
      button.addEventListener("mousedown", (event) => event.preventDefault());
      button.addEventListener("click", () => {
        const command = button.getAttribute("data-command");
        if (command) {
          executeFormattingCommand(command);
        }
      });
    });
  }

  async function initializeRuntime() {
    const query = new URLSearchParams(window.location.search);
    const fallback = getContextFromQuery(query);

    if (!window.rliSdk || typeof window.rliSdk.init !== "function") {
      return {
        client: null,
        connected: false,
        context: fallback,
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
      };
    } catch (error) {
      console.warn("Rocketlane SDK unavailable, using local preview mode.", error);
      return {
        client: null,
        connected: false,
        context: fallback,
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
      const key = aliases[objectName];
      const identifier =
        client.data.dataIdentifiers &&
        key &&
        client.data.dataIdentifiers[key];

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
      projectId:
        pickFirst(project && (project.id || project.projectId || project._id)) ||
        fallback.projectId,
      projectName:
        pickFirst(project && (project.name || project.projectName)) ||
        fallback.projectName,
    };
  }

  function getContextFromQuery(query) {
    return {
      accountId: query.get("accountId") || "",
      accountName: query.get("account") || "Rocketlane Workspace",
      userId: query.get("userId") || "",
      userName: query.get("user") || "Rocketlane User",
      projectId: query.get("projectId") || "",
      projectName: query.get("project") || "",
    };
  }

  function updateHeader(runtime) {
    const context = runtime.context;
    const view = context.projectName || "Customer home";

    refs.boardScope.textContent =
      "Scope: " + context.accountName + " / " + view + " bulletin board";

    if (runtime.connected) {
      refs.connectionBadge.className = "badge badge-ok";
      refs.connectionBadge.textContent = "Connected to Rocketlane";
    } else {
      refs.connectionBadge.className = "badge badge-local";
      refs.connectionBadge.textContent = "Local preview mode";
    }
  }

  function createStorageKey(context) {
    const accountScope = slug(context.accountId || context.accountName || "workspace");
    const viewScope = slug(
      context.projectId || context.projectName || "customer-home-view"
    );
    return STORAGE_PREFIX + ":" + accountScope + ":" + viewScope;
  }

  function loadNotes() {
    const raw = window.localStorage.getItem(state.storageKey);
    if (!raw) {
      state.notes = [];
      return;
    }

    try {
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        state.notes = [];
        return;
      }

      state.notes = parsed.map((note) => normalizeNote(note));
    } catch (error) {
      console.warn("Unable to parse saved notes", error);
      state.notes = [];
    }
  }

  function persistNotes() {
    window.localStorage.setItem(state.storageKey, JSON.stringify(state.notes));
    setSaveState("Saved " + formatTime(new Date().toISOString()));
  }

  function queueSave() {
    setSaveState("Saving...");
    if (state.saveTimer) {
      window.clearTimeout(state.saveTimer);
    }

    state.saveTimer = window.setTimeout(() => {
      persistNotes();
    }, 240);
  }

  function setSaveState(text) {
    refs.saveState.textContent = text;
  }

  function createDefaultWelcomeNote() {
    const note = normalizeNote({
      title: "Welcome to your customer bulletin board",
      pinned: true,
      color: "yellow",
      content:
        "<p>This sticky board is ready for your Rocketlane home page updates.</p>" +
        "<ul><li>Share announcements</li><li>Track open action items</li></ul>" +
        "<p>Use the toolbar for bullet and numbered lists like Gmail.</p>",
    });

    state.notes = [note];
    state.selectedNoteId = note.id;
    persistNotes();
  }

  function seedDemoNotes() {
    const now = Date.now();
    const demo = [
      {
        title: "Go-live countdown",
        pinned: true,
        color: "yellow",
        updatedAt: new Date(now - 1000 * 60 * 4).toISOString(),
        content:
          "<p>Final checks for launch:</p>" +
          "<ol><li>Confirm SSO handoff</li><li>Validate training links</li><li>Publish kickoff update</li></ol>",
      },
      {
        title: "Customer asks this week",
        color: "blue",
        updatedAt: new Date(now - 1000 * 60 * 30).toISOString(),
        content:
          "<p>Top requests from customer standup:</p>" +
          "<ul><li>Add billing FAQ in portal</li><li>Share onboarding checklist</li><li>Post ETA for migration</li></ul>",
      },
      {
        title: "Ops reminders",
        color: "green",
        updatedAt: new Date(now - 1000 * 60 * 70).toISOString(),
        content:
          "<p>Keep these visible:</p>" +
          "<ul><li>Daily sync at 10:00 AM</li><li>Escalation channel monitored</li><li>Weekly recap every Friday</li></ul>",
      },
      {
        title: "Template for announcements",
        color: "pink",
        updatedAt: new Date(now - 1000 * 60 * 95).toISOString(),
        content:
          "<p><strong>Update:</strong> <em>What changed?</em></p>" +
          "<p><strong>Impact:</strong> Who is affected?</p>" +
          "<p><strong>Next step:</strong></p>" +
          "<ol><li>Owner</li><li>Timeline</li><li>Follow-up link</li></ol>",
      },
    ];

    state.notes = demo.map((item) => normalizeNote(item));
    state.selectedNoteId = state.notes[0] ? state.notes[0].id : null;
    persistNotes();
  }

  function onAddNote() {
    const next = normalizeNote({
      title: "Untitled note",
      color: NOTE_COLORS[0].id,
      content: "",
    });
    state.notes.unshift(next);
    state.selectedNoteId = next.id;
    persistNotes();
    renderAll();
    refs.noteTitleInput.focus();
    refs.noteTitleInput.select();
  }

  function onClearBoard() {
    if (!state.notes.length) {
      return;
    }

    const confirmed = window.confirm(
      "Clear all sticky notes for this board scope? This cannot be undone."
    );

    if (!confirmed) {
      return;
    }

    state.notes = [];
    state.selectedNoteId = null;
    persistNotes();
    renderAll();
  }

  function onTitleInput(event) {
    const note = getSelectedNote();
    if (!note) {
      return;
    }

    note.title = event.target.value.slice(0, 120);
    note.updatedAt = new Date().toISOString();
    queueSave();
    renderNotesGrid();
    renderStats();
  }

  function onBodyInput() {
    const note = getSelectedNote();
    if (!note) {
      return;
    }

    note.content = normalizeEditorHtml(refs.noteBodyInput.innerHTML);
    note.updatedAt = new Date().toISOString();
    queueSave();
    renderNotesGrid();
    renderStats();
  }

  function onEditorPaste(event) {
    event.preventDefault();
    const text = (event.clipboardData || window.clipboardData).getData("text");
    document.execCommand("insertText", false, text);
  }

  function onTogglePinFromEditor() {
    const note = getSelectedNote();
    if (!note) {
      return;
    }

    note.pinned = !note.pinned;
    note.updatedAt = new Date().toISOString();
    persistNotes();
    renderAll();
  }

  function onDeleteFromEditor() {
    const note = getSelectedNote();
    if (!note) {
      return;
    }

    const confirmed = window.confirm(
      'Delete note "' + (note.title || "Untitled note") + '"?'
    );
    if (!confirmed) {
      return;
    }

    state.notes = state.notes.filter((item) => item.id !== note.id);
    const visible = getVisibleNotes();
    state.selectedNoteId = visible[0] ? visible[0].id : null;
    persistNotes();
    renderAll();
  }

  function onColorSwatchClick(event) {
    const button = event.target.closest(".color-swatch");
    if (!button) {
      return;
    }

    const nextColor = button.getAttribute("data-color");
    const note = getSelectedNote();
    if (!note || !nextColor) {
      return;
    }

    note.color = nextColor;
    note.updatedAt = new Date().toISOString();
    persistNotes();
    renderAll();
  }

  function executeFormattingCommand(command) {
    refs.noteBodyInput.focus();
    const worked = document.execCommand(command, false, null);
    if (worked === false) {
      setSaveState("Formatting command not supported in this browser");
      return;
    }
    refs.noteBodyInput.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function renderAll() {
    renderNotesGrid();
    renderEditor();
    renderStats();
  }

  function renderStats() {
    const total = state.notes.length;
    const pinned = state.notes.filter((note) => note.pinned).length;
    const visible = getVisibleNotes().length;

    refs.boardStats.textContent =
      visible +
      " visible / " +
      total +
      " total notes (" +
      pinned +
      " pinned)";
  }

  function renderNotesGrid() {
    const notes = getVisibleNotes();
    refs.notesGrid.innerHTML = "";

    if (!notes.length) {
      refs.notesEmptyState.classList.remove("hidden");
      return;
    }

    refs.notesEmptyState.classList.add("hidden");

    notes.forEach((note, index) => {
      const card = document.createElement("article");
      card.className = "note-card note-color-" + note.color;
      if (note.id === state.selectedNoteId) {
        card.classList.add("selected");
      }
      card.style.setProperty("--tilt", tiltFor(note.id, index) + "deg");
      card.setAttribute("tabindex", "0");
      card.setAttribute("role", "button");
      card.setAttribute("aria-label", "Open note " + (note.title || "Untitled"));

      card.addEventListener("click", () => {
        state.selectedNoteId = note.id;
        renderAll();
      });

      card.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          state.selectedNoteId = note.id;
          renderAll();
        }
      });

      const header = document.createElement("div");
      header.className = "note-header";

      const title = document.createElement("p");
      title.className = "note-title";
      title.textContent = note.title || "Untitled note";
      header.appendChild(title);

      if (note.pinned) {
        const pin = document.createElement("span");
        pin.className = "pin-chip";
        pin.textContent = "Pinned";
        header.appendChild(pin);
      }

      const preview = document.createElement("p");
      preview.className = "note-preview";
      preview.textContent = summarize(note.content, 170);

      const meta = document.createElement("p");
      meta.className = "note-meta";
      meta.textContent = "Updated " + formatTime(note.updatedAt);

      card.appendChild(header);
      card.appendChild(preview);
      card.appendChild(meta);
      refs.notesGrid.appendChild(card);
    });
  }

  function renderEditor() {
    const note = getSelectedNote();

    if (!note) {
      refs.editorPanel.classList.add("hidden");
      refs.editorEmptyState.classList.remove("hidden");
      return;
    }

    refs.editorEmptyState.classList.add("hidden");
    refs.editorPanel.classList.remove("hidden");

    if (refs.noteTitleInput.value !== note.title) {
      refs.noteTitleInput.value = note.title;
    }

    if (document.activeElement !== refs.noteBodyInput) {
      refs.noteBodyInput.innerHTML = note.content || "";
    }

    refs.pinNoteButton.textContent = note.pinned ? "Unpin" : "Pin";
    renderColorPickerSelection(note.color);
    setSaveState("Saved");
  }

  function buildColorPicker() {
    refs.colorPicker.innerHTML = "";
    NOTE_COLORS.forEach((color) => {
      const swatch = document.createElement("button");
      swatch.type = "button";
      swatch.className = "color-swatch";
      swatch.style.background = color.hex;
      swatch.setAttribute("data-color", color.id);
      swatch.setAttribute("title", color.label);
      swatch.setAttribute("aria-label", "Set color " + color.label);
      refs.colorPicker.appendChild(swatch);
    });
  }

  function renderColorPickerSelection(colorId) {
    refs.colorPicker.querySelectorAll(".color-swatch").forEach((swatch) => {
      const isSelected = swatch.getAttribute("data-color") === colorId;
      swatch.classList.toggle("selected", isSelected);
    });
  }

  function getVisibleNotes() {
    const search = state.searchQuery.toLowerCase();
    let notes = state.notes.slice();

    if (state.filter === "pinned") {
      notes = notes.filter((note) => note.pinned);
    }

    const sorted = notes.sort((a, b) => {
      if (state.filter !== "recent" && a.pinned !== b.pinned) {
        return a.pinned ? -1 : 1;
      }
      return timestampValue(b.updatedAt) - timestampValue(a.updatedAt);
    });

    if (!search) {
      return sorted;
    }

    return sorted.filter((note) => {
      const haystack = (note.title + " " + stripHtml(note.content)).toLowerCase();
      return haystack.includes(search);
    });
  }

  function getSelectedNote() {
    return state.notes.find((note) => note.id === state.selectedNoteId) || null;
  }

  function normalizeNote(note) {
    const createdAt = note && note.createdAt ? note.createdAt : new Date().toISOString();
    const updatedAt = note && note.updatedAt ? note.updatedAt : createdAt;

    return {
      id: (note && note.id) || createId(),
      title: (note && String(note.title || "").trim()) || "Untitled note",
      content: normalizeEditorHtml((note && note.content) || ""),
      color: resolveColor((note && note.color) || NOTE_COLORS[0].id),
      pinned: Boolean(note && note.pinned),
      createdAt,
      updatedAt,
    };
  }

  function resolveColor(colorId) {
    return NOTE_COLORS.some((item) => item.id === colorId)
      ? colorId
      : NOTE_COLORS[0].id;
  }

  function createId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }

    return (
      "note-" + Math.random().toString(16).slice(2) + "-" + Date.now().toString(16)
    );
  }

  function normalizeEditorHtml(html) {
    const cleaned = String(html || "").replace(/^\s+|\s+$/g, "");
    if (!cleaned || cleaned === "<br>" || cleaned === "<div><br></div>") {
      return "";
    }
    return cleaned;
  }

  function summarize(html, maxLength) {
    const text = stripHtml(html).replace(/\s+/g, " ").trim();
    if (!text) {
      return "No content yet.";
    }
    return text.length > maxLength ? text.slice(0, maxLength - 1) + "…" : text;
  }

  function stripHtml(html) {
    const container = document.createElement("div");
    container.innerHTML = html || "";
    return container.textContent || container.innerText || "";
  }

  function tiltFor(id, index) {
    let hash = 0;
    const source = String(id) + ":" + String(index);
    for (let i = 0; i < source.length; i += 1) {
      hash = (hash << 5) - hash + source.charCodeAt(i);
      hash |= 0;
    }
    const raw = ((Math.abs(hash) % 8) - 4) * 0.55;
    return raw.toFixed(2);
  }

  function formatTime(isoString) {
    if (!isoString) {
      return "just now";
    }
    const date = new Date(isoString);
    if (Number.isNaN(date.getTime())) {
      return "just now";
    }
    return date.toLocaleString([], { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
  }

  function timestampValue(isoString) {
    const value = new Date(isoString).getTime();
    return Number.isNaN(value) ? 0 : value;
  }

  function shouldSeedDemo(query) {
    const value = query.get("demo");
    return value === "1" || value === "true";
  }

  function slug(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/(^-|-$)/g, "") || "default";
  }

  function pickFirst(value) {
    if (typeof value === "string") {
      const trimmed = value.trim();
      return trimmed || "";
    }
    if (typeof value === "number") {
      return String(value);
    }
    return "";
  }
})();
