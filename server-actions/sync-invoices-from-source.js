"use strict";

const DEFAULT_SOURCE_PROJECTS = ["Expert Advisor Program Invoices"];
// Production override: embed API key here so app works without installer prompt.
// Replace before shipping to users if needed.
const EMBEDDED_ROCKETLANE_API_KEY = "rl-6657ce9e-ee84-465d-b4df-d97b1239a343";
const EMBEDDED_ROCKETLANE_API_KEY_WORKSPACE = "innovate-calgary.rocketlane.com";
const ROCKETLANE_API_BASE_URL = "https://api.rocketlane.com";
const FIELD_ALIAS_GROUPS = {
  contractName: ["contract name", "contract", "contractname"],
  hub: ["hub"],
  program: ["program"],
  accountName: ["account", "client", "customer", "company"],
  createdBy: ["created by", "createdby", "creator"],
  quantityHours: ["quantity", "qty", "hour", "hours"],
};

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

function isSourceProjectName(name, sourceProjectNames) {
  const normalized = canonicalProjectName(name);
  if (!normalized) {
    return false;
  }
  return sourceProjectNames.some((candidate) => {
    const target = canonicalProjectName(candidate);
    return (
      normalized === target ||
      normalized.includes(target) ||
      target.includes(normalized)
    );
  });
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

function fullName(value) {
  if (!value || typeof value !== "object") {
    return "";
  }
  const first = pickFirst(value.firstName || value.first_name);
  const last = pickFirst(value.lastName || value.last_name);
  const combined = `${first} ${last}`.trim();
  return combined || pickFirst(value.name || value.displayName || value.userName);
}

function normalizeDateValue(value) {
  if (value == null) {
    return "";
  }
  if (typeof value === "number") {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? "" : date.toISOString();
  }
  const text = String(value).trim();
  if (!text) {
    return "";
  }
  if (/^\d{10,13}$/.test(text)) {
    const asNumber = Number(text);
    const date = new Date(asNumber);
    if (!Number.isNaN(date.getTime())) {
      return date.toISOString();
    }
  }
  const date = new Date(text);
  if (!Number.isNaN(date.getTime())) {
    return date.toISOString();
  }
  return text;
}

function normalizeAmount(value) {
  if (value == null || value === "") {
    return 0;
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  const numeric = Number(String(value).replace(/[^0-9.-]/g, ""));
  return Number.isFinite(numeric) ? numeric : 0;
}

function isLikelyIdentifierValue(value) {
  const text = pickFirst(value);
  if (!text) {
    return false;
  }
  if (/^\d+$/.test(text)) {
    return true;
  }
  if (/^[0-9a-f]{16,}$/i.test(text)) {
    return true;
  }
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(text)) {
    return true;
  }
  return false;
}

function isLikelyDisplayName(value) {
  const text = pickFirst(value);
  if (!text) {
    return false;
  }
  if (isLikelyIdentifierValue(text)) {
    return false;
  }
  return /[a-z]/i.test(text);
}

function pickPreferredAccountName(candidates) {
  const values = (Array.isArray(candidates) ? candidates : []).map((value) => pickFirst(value));
  const informative = values.find(
    (value) =>
      value &&
      normalizeFieldLabel(value) !== "rocketlane account" &&
      normalizeFieldLabel(value) !== "rocketlane workspace"
  );
  if (informative) {
    return informative;
  }
  return values.find(Boolean) || "";
}

function isGenericAccountName(value) {
  const normalized = normalizeFieldLabel(value);
  return (
    !normalized ||
    normalized === "rocketlane account" ||
    normalized === "rocketlane workspace" ||
    normalized === "account"
  );
}

function normalizeFieldLabel(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function humanizeFieldLabel(value) {
  return String(value || "")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function fieldLabelMatchesAlias(label, aliases) {
  const normalizedLabel = normalizeFieldLabel(label);
  if (!normalizedLabel || !Array.isArray(aliases)) {
    return false;
  }
  return aliases.some((alias) => {
    const normalizedAlias = normalizeFieldLabel(alias);
    return (
      normalizedAlias &&
      (normalizedLabel === normalizedAlias ||
        normalizedLabel.includes(normalizedAlias) ||
        normalizedAlias.includes(normalizedLabel))
    );
  });
}

function toFieldText(value) {
  if (value == null) {
    return "";
  }
  if (Array.isArray(value)) {
    return dedupeStrings(value.map((entry) => toFieldText(entry))).join(", ");
  }
  if (typeof value === "object") {
    return pickFirst(
      value.fieldValueLabel ||
        value.fieldValue ||
        value.label ||
        value.name ||
        value.value ||
        value.displayValue
    );
  }
  return String(value).trim();
}

function extractNamedCustomFieldValues(fields) {
  const output = {
    contractName: [],
    hub: [],
    program: [],
    accountName: [],
    createdBy: [],
    quantityHours: [],
  };
  const entries = extractFieldDisplayEntries(fields);
  if (!entries.length) {
    return output;
  }
  entries.forEach((entry) => {
    if (!entry || typeof entry !== "object") {
      return;
    }
    const label = normalizeFieldLabel(entry.label);
    if (!label) {
      return;
    }
    const value = toFieldText(entry.value);
    if (!value) {
      return;
    }
    Object.keys(FIELD_ALIAS_GROUPS).forEach((targetKey) => {
      if (!fieldLabelMatchesAlias(label, FIELD_ALIAS_GROUPS[targetKey])) {
        return;
      }
      value
        .split(",")
        .map((part) => pickFirst(part))
        .filter(Boolean)
        .forEach((part) => output[targetKey].push(part));
    });
  });
  output.contractName = dedupeStrings(output.contractName);
  output.hub = dedupeStrings(output.hub);
  output.program = dedupeStrings(output.program);
  output.accountName = dedupeStrings(output.accountName);
  output.createdBy = dedupeStrings(output.createdBy);
  output.quantityHours = dedupeStrings(output.quantityHours);
  return output;
}

function extractFieldDisplayEntries(fields) {
  const entries = [];
  const seen = new Set();
  const pushEntry = (label, value) => {
    const cleanLabel = pickFirst(label);
    const cleanValue = toFieldText(value);
    if (!cleanLabel || !cleanValue) {
      return;
    }
    const dedupeKey = `${normalizeFieldLabel(cleanLabel)}|${cleanValue}`;
    if (seen.has(dedupeKey)) {
      return;
    }
    seen.add(dedupeKey);
    entries.push({ label: cleanLabel, value: cleanValue });
  };
  const walk = (node, depth, fallbackLabel) => {
    if (depth > 5 || node == null) {
      return;
    }
    if (Array.isArray(node)) {
      node.forEach((entry) => walk(entry, depth + 1, fallbackLabel));
      return;
    }
    if (typeof node !== "object") {
      if (fallbackLabel) {
        pushEntry(fallbackLabel, node);
      }
      return;
    }
    const label = pickFirst(
      node.fieldLabel ||
        node.fieldName ||
        node.label ||
        node.name ||
        node.key ||
        node.title ||
        fallbackLabel
    );
    const directValue =
      pickFirst(
        node.fieldValueLabel ||
          node.fieldValue ||
          node.displayValue ||
          node.value ||
          (node.metaFieldValue && (node.metaFieldValue.label || node.metaFieldValue.value)) ||
          (node.option && (node.option.label || node.option.value))
      ) || node.values;
    if (label && directValue !== undefined) {
      pushEntry(label, directValue);
    }
    Object.keys(node).forEach((key) => {
      if (!Object.prototype.hasOwnProperty.call(node, key)) {
        return;
      }
      const value = node[key];
      if (value == null) {
        return;
      }
      if (typeof value !== "object") {
        if (
          key !== "id" &&
          key !== "_id" &&
          key !== "createdAt" &&
          key !== "updatedAt" &&
          key !== "deletedAt"
        ) {
          pushEntry(humanizeFieldLabel(key), value);
        }
        return;
      }
      walk(value, depth + 1, humanizeFieldLabel(key));
    });
  };
  walk(fields, 0, "");
  return entries;
}

function mergeFieldDisplayEntries(primary, secondary) {
  const result = [];
  const seen = new Set();
  const pushEntries = (entries) => {
    if (!Array.isArray(entries)) {
      return;
    }
    entries.forEach((entry) => {
      if (!entry || typeof entry !== "object") {
        return;
      }
      const label = pickFirst(entry.label);
      const value = pickFirst(entry.value);
      if (!label || !value) {
        return;
      }
      const key = `${normalizeFieldLabel(label)}|${value}`;
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      result.push({ label, value });
    });
  };
  pushEntries(primary);
  pushEntries(secondary);
  return result;
}

function compactJoined(values) {
  return dedupeStrings((Array.isArray(values) ? values : []).map((value) => pickFirst(value))).join(
    ", "
  );
}

function collectCustomFieldSources(record) {
  if (!record || typeof record !== "object") {
    return [];
  }
  return [
    record.fields,
    record.customFields,
    record.customFieldValues,
    record.fieldValues,
    record.projectFields,
    record.invoiceFields,
    record.metadataFields,
  ].filter((value) => value && (Array.isArray(value) || typeof value === "object"));
}

function extractCustomFieldAliases(record) {
  const merged = {
    contractName: [],
    hub: [],
    program: [],
    accountName: [],
    createdBy: [],
    quantityHours: [],
  };
  collectCustomFieldSources(record).forEach((source) => {
    const extracted = extractNamedCustomFieldValues(source);
    merged.contractName.push(...(extracted.contractName || []));
    merged.hub.push(...(extracted.hub || []));
    merged.program.push(...(extracted.program || []));
    merged.accountName.push(...(extracted.accountName || []));
    merged.createdBy.push(...(extracted.createdBy || []));
    merged.quantityHours.push(...(extracted.quantityHours || []));
  });
  merged.contractName = dedupeStrings(merged.contractName);
  merged.hub = dedupeStrings(merged.hub);
  merged.program = dedupeStrings(merged.program);
  merged.accountName = dedupeStrings(merged.accountName);
  merged.createdBy = dedupeStrings(merged.createdBy);
  merged.quantityHours = dedupeStrings(merged.quantityHours);
  return merged;
}

function collectInvoiceLineItems(record) {
  const lineItems = [];
  const pushLines = (items) => {
    if (!Array.isArray(items)) {
      return;
    }
    items.forEach((item) => {
      if (item && typeof item === "object") {
        lineItems.push(item);
      }
    });
  };
  pushLines(record && record.invoiceLineItems);
  if (record && Array.isArray(record.invoiceToSourceMappings)) {
    record.invoiceToSourceMappings.forEach((mapping) => {
      pushLines(mapping && mapping.invoiceLineItems);
    });
  }
  return lineItems;
}

function sumLineItemQuantity(lineItems) {
  return (Array.isArray(lineItems) ? lineItems : []).reduce((sum, line) => {
    const lineFields = extractNamedCustomFieldValues(line && line.fields);
    const fieldQuantity = normalizeAmount(
      (lineFields.quantityHours && lineFields.quantityHours[0]) || ""
    );
    const quantity = normalizeAmount(
      line &&
        (line.quantity ||
          line.qty ||
          line.qtyHours ||
          line.hours ||
          line.billableHours ||
          line.billableQuantity ||
          line.quantityHours ||
          fieldQuantity)
    );
    return sum + (Number.isFinite(quantity) ? quantity : 0);
  }, 0);
}

function invoiceMatchesQuery(invoice, query) {
  const normalizedQuery = String(query || "").trim().toLowerCase();
  if (!normalizedQuery) {
    return true;
  }
  const haystack = (
    String(invoice.invoiceStatus || "") +
    " " +
    String(invoice.invoiceNumber || "") +
    " " +
    String(invoice.invoiceName || "") +
    " " +
    String(invoice.ownerName || "") +
    " " +
    String(invoice.accountName || "") +
    " " +
    String(invoice.issueDate || invoice.invoiceDate || "") +
    " " +
    String(invoice.dueDate || "") +
    " " +
    String(invoice.amount || "") +
    " " +
    String(invoice.contractName || "") +
    " " +
    String(invoice.hub || "") +
    " " +
    String(invoice.program || "") +
    " " +
    String(invoice.quantityHours || "") +
    " " +
    String(invoice.sourceProjectName || "") +
    " " +
    (Array.isArray(invoice.associatedEmails) ? invoice.associatedEmails.join(" ") : "")
  ).toLowerCase();
  return haystack.includes(normalizedQuery);
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function mergeObjects(a, b) {
  return Object.assign({}, a || {}, b || {});
}

function ensureAbsoluteUrl(baseUrl, path) {
  try {
    return new URL(path, baseUrl).toString();
  } catch (_error) {
    return "";
  }
}

function workspaceHost(value) {
  const text = pickFirst(value);
  if (!text) {
    return "";
  }
  try {
    return String(new URL(text).hostname || "").toLowerCase();
  } catch (_error) {
    return "";
  }
}

function hostMatchesTarget(host, targetHost) {
  const left = pickFirst(host).toLowerCase();
  const right = pickFirst(targetHost).toLowerCase();
  if (!left || !right) {
    return false;
  }
  return left === right || left.endsWith(`.${right}`);
}

function canUseEmbeddedToken(request, workspaceCandidates) {
  const targetHost = String(EMBEDDED_ROCKETLANE_API_KEY_WORKSPACE || "").toLowerCase();
  if (!targetHost) {
    return false;
  }
  const runtimeWorkspaceHost = workspaceHost(
    (request && request.viewerContext && request.viewerContext.workspaceBaseUrl) ||
      (request && request.workspaceBaseUrl)
  );
  if (runtimeWorkspaceHost) {
    return runtimeWorkspaceHost === targetHost;
  }
  return (Array.isArray(workspaceCandidates) ? workspaceCandidates : [])
    .map((candidate) => workspaceHost(candidate))
    .some((host) => host === targetHost);
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
      if (payload.response && typeof payload.response === "object") {
        pushIfArray(payload.response[key]);
      }
    }
  });

  if (preferred.length) {
    return preferred.sort((a, b) => b.length - a.length)[0];
  }

  return [];
}

function extractRecordObject(payload, preferredKeys) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    return null;
  }
  const keys = Array.isArray(preferredKeys) ? preferredKeys : [];
  for (let i = 0; i < keys.length; i += 1) {
    const key = keys[i];
    const direct = payload[key];
    if (direct && typeof direct === "object" && !Array.isArray(direct)) {
      return direct;
    }
    if (payload.data && payload.data[key] && typeof payload.data[key] === "object") {
      return payload.data[key];
    }
    if (payload.response && payload.response[key] && typeof payload.response[key] === "object") {
      return payload.response[key];
    }
    if (payload.result && payload.result[key] && typeof payload.result[key] === "object") {
      return payload.result[key];
    }
  }
  if (payload.data && typeof payload.data === "object" && !Array.isArray(payload.data)) {
    return payload.data;
  }
  if (payload.response && typeof payload.response === "object" && !Array.isArray(payload.response)) {
    return payload.response;
  }
  if (payload.result && typeof payload.result === "object" && !Array.isArray(payload.result)) {
    return payload.result;
  }
  if (payload.payload && typeof payload.payload === "object" && !Array.isArray(payload.payload)) {
    return payload.payload;
  }
  return payload;
}

async function requestJson(url, headers) {
  const response = await fetch(url, {
    method: "GET",
    headers,
  });
  if (!response.ok) {
    throw new Error(`Request failed (${response.status}) for ${url}`);
  }
  const text = await response.text();
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch (_error) {
    throw new Error(`Expected JSON payload for ${url}`);
  }
}

async function requestBinary(url, headers) {
  const response = await fetch(url, {
    method: "GET",
    headers,
  });
  if (!response.ok) {
    throw new Error(`Request failed (${response.status}) for ${url}`);
  }
  const buffer = await response.arrayBuffer();
  if (!buffer || !buffer.byteLength) {
    return null;
  }
  return new Uint8Array(buffer);
}

function bytesToPdfDataUrl(bytes) {
  if (!(bytes instanceof Uint8Array) || !bytes.byteLength) {
    return "";
  }
  return `data:application/pdf;base64,${Buffer.from(bytes).toString("base64")}`;
}

function canonicalInvoiceNumber(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

function buildInvoiceDisplayNumber(record) {
  if (!record || typeof record !== "object") {
    return "";
  }
  const full = pickFirst(record.invoiceNumber);
  if (full && /^INV[-\s]?\d+/i.test(full)) {
    return full;
  }
  const prefix = pickFirst(record.invoiceNumberPrefix || "INV-");
  const suffix = pickFirst(record.invoiceNumber);
  if (!suffix) {
    return "";
  }
  return `${prefix}${suffix}`;
}

function extractInvoiceProjectIds(record) {
  if (!record || typeof record !== "object") {
    return [];
  }
  const ids = [];
  const mappings = Array.isArray(record.invoiceToSourceMappings)
    ? record.invoiceToSourceMappings
    : [];
  mappings.forEach((mapping) => {
    ids.push(
      pickFirst(
        mapping &&
          (mapping.sourceId ||
            (mapping.project && (mapping.project.projectId || mapping.project.id)))
      )
    );
  });
  if (Array.isArray(record.projects)) {
    record.projects.forEach((project) => {
      ids.push(
        pickFirst(
          project && (project.projectId || project.id || project._id || project.projectID)
        )
      );
    });
  }
  if (record.projects && typeof record.projects === "object" && !Array.isArray(record.projects)) {
    ids.push(
      pickFirst(
        record.projects.projectId ||
          record.projects.id ||
          record.projects._id ||
          record.projects.projectID
      )
    );
  }
  return dedupeStrings(ids);
}

function buildProjectLookup(projects) {
  const rows = Array.isArray(projects) ? projects : [];
  const byId = new Map();
  const byCanonicalName = new Map();
  rows.forEach((project) => {
    if (!project || typeof project !== "object") {
      return;
    }
    const id = pickFirst(project.id);
    const name = pickFirst(project.name);
    if (id) {
      byId.set(id, project);
    }
    if (name) {
      const key = canonicalProjectName(name);
      if (key && !byCanonicalName.has(key)) {
        byCanonicalName.set(key, project);
      }
    }
  });
  return { byId, byCanonicalName };
}

function resolveProjectForInvoice(record, projectLookup) {
  if (!record || typeof record !== "object") {
    return null;
  }
  const lookup = projectLookup || { byId: new Map(), byCanonicalName: new Map() };
  const projectIds = dedupeStrings(
    extractInvoiceProjectIds(record).concat([
      pickFirst(record.projectId || record.projectID),
      pickFirst(
        record.project &&
          (record.project.projectId ||
            record.project.id ||
            record.project.projectID)
      ),
    ])
  );
  for (let i = 0; i < projectIds.length; i += 1) {
    const id = projectIds[i];
    if (id && lookup.byId.has(id)) {
      return lookup.byId.get(id);
    }
  }

  const mappings = Array.isArray(record.invoiceToSourceMappings)
    ? record.invoiceToSourceMappings
    : [];
  for (let i = 0; i < mappings.length; i += 1) {
    const mapping = mappings[i] || {};
    const mappedProject = normalizeProject(mapping.project || mapping.sourceProject);
    if (mappedProject) {
      return mappedProject;
    }
    const mappedName = pickFirst(
      (mapping.project && (mapping.project.projectName || mapping.project.name)) ||
        mapping.projectName
    );
    const canonicalName = canonicalProjectName(mappedName);
    if (canonicalName && lookup.byCanonicalName.has(canonicalName)) {
      return lookup.byCanonicalName.get(canonicalName);
    }
  }

  const directProject = normalizeProject(record.project);
  if (directProject) {
    return directProject;
  }
  if (Array.isArray(record.projects)) {
    for (let i = 0; i < record.projects.length; i += 1) {
      const candidate = record.projects[i] || {};
      const normalizedCandidate = normalizeProject(candidate);
      if (normalizedCandidate) {
        return normalizedCandidate;
      }
      const candidateName = pickFirst(
        candidate.projectName || candidate.name || candidate.projectTitle || candidate.label
      );
      const candidateCanonical = canonicalProjectName(candidateName);
      if (candidateCanonical && lookup.byCanonicalName.has(candidateCanonical)) {
        return lookup.byCanonicalName.get(candidateCanonical);
      }
    }
  }
  const directName = pickFirst(record.projectName || record.projectTitle);
  const directCanonical = canonicalProjectName(directName);
  if (directCanonical && lookup.byCanonicalName.has(directCanonical)) {
    return lookup.byCanonicalName.get(directCanonical);
  }
  return null;
}

async function resolveInvoiceIdForPreview(
  baseUrl,
  headers,
  previewInvoiceId,
  previewInvoiceNumber,
  previewSourceProjectId
) {
  const explicitId = pickFirst(previewInvoiceId);
  if (/^\d+$/.test(explicitId)) {
    return explicitId;
  }

  const targetNumber = canonicalInvoiceNumber(
    previewInvoiceNumber || previewInvoiceId || ""
  );
  if (!targetNumber) {
    return explicitId;
  }

  const containsToken = targetNumber.replace(/^INV/, "");
  const lookupUrl = ensureAbsoluteUrl(
    baseUrl,
    `/api/v1/invoices?invoiceNumber.contains=${encodeURIComponent(containsToken)}`
  );
  if (!lookupUrl) {
    return explicitId;
  }
  const lookupPayload = await requestJson(lookupUrl, headers);
  const rows = Array.isArray(lookupPayload)
    ? lookupPayload
    : extractCollection(lookupPayload, ["data", "invoices", "items", "results"]);
  if (!rows.length) {
    return explicitId;
  }

  const targetProjectId = pickFirst(previewSourceProjectId);
  const scored = rows
    .map((row) => {
      const invoiceId = pickFirst(row && row.invoiceId);
      if (!invoiceId) {
        return null;
      }
      const displayNumber = canonicalInvoiceNumber(buildInvoiceDisplayNumber(row));
      const projectIds = extractInvoiceProjectIds(row);
      const projectMatch =
        !targetProjectId || projectIds.includes(String(targetProjectId));
      let score = 0;
      if (displayNumber === targetNumber) {
        score += 10;
      } else if (displayNumber.includes(targetNumber) || targetNumber.includes(displayNumber)) {
        score += 5;
      }
      if (projectMatch) {
        score += 3;
      }
      if (explicitId && invoiceId === explicitId) {
        score += 2;
      }
      return { invoiceId, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  return (scored[0] && scored[0].invoiceId) || explicitId;
}

async function requestCollection(baseUrl, headers, paths, preferredKeys) {
  const rows = [];
  const seen = new Set();
  const errors = [];

  for (let i = 0; i < paths.length; i += 1) {
    const url = ensureAbsoluteUrl(baseUrl, paths[i]);
    if (!url) {
      continue;
    }
    try {
      const payload = await requestJson(url, headers);
      const records = extractCollection(payload, preferredKeys);
      records.forEach((record) => {
        const key =
          pickFirst(
            record &&
              (record.id ||
                record._id ||
                record.projectId ||
                record.documentId ||
                record.fileId ||
                record.invoiceId ||
                record.invoiceNumber)
          ) ||
          JSON.stringify(record || {});
        if (seen.has(key)) {
          return;
        }
        seen.add(key);
        rows.push(record);
      });
    } catch (error) {
      errors.push(String(error && error.message ? error.message : error));
    }
  }

  return { rows, errors };
}

function extractViewerCandidate(payload) {
  if (!payload) {
    return null;
  }
  const objectCandidates = [
    payload,
    payload.data,
    payload.response,
    payload.result,
    payload.payload,
    payload.user,
    payload.data && payload.data.user,
    payload.response && payload.response.user,
  ];
  for (let i = 0; i < objectCandidates.length; i += 1) {
    const candidate = objectCandidates[i];
    if (!candidate) {
      continue;
    }
    if (Array.isArray(candidate)) {
      for (let j = 0; j < candidate.length; j += 1) {
        const normalized = normalizeMember(candidate[j]);
        if (normalized) {
          return normalized;
        }
      }
      continue;
    }
    const normalized = normalizeMember(candidate);
    if (normalized) {
      return normalized;
    }
  }
  return null;
}

async function requestCurrentViewer(baseUrl, headers) {
  const paths = [
    "/api/1.0/users/me?includeFields=permission,role,company",
    "/api/1.0/users/me?includeFields=permission",
    "/api/1.0/users/me",
    "/api/v1/users/me",
    "/api/v1/users/current",
    "/api/1.0/account-users/me",
  ];
  for (let i = 0; i < paths.length; i += 1) {
    const path = paths[i];
    try {
      const payload = await requestJson(ensureAbsoluteUrl(baseUrl, path), headers);
      const viewer = extractViewerCandidate(payload);
      if (viewer) {
        return viewer;
      }
    } catch (_error) {
      // Ignore and continue with next endpoint candidate.
    }
  }
  return null;
}

function normalizeProject(record) {
  if (!record || typeof record !== "object") {
    return null;
  }
  const id = pickFirst(record.id || record._id || record.projectId);
  const name = pickFirst(
    record.name || record.projectName || record.projectTitle || record.title
  );
  if (!name) {
    return null;
  }
  const accountName = pickFirst(
    record.accountName ||
      record.companyName ||
      (record.company && record.company.companyName) ||
      (record.company && record.company.name) ||
      (record.account && record.account.name) ||
      (record.customer && record.customer.name) ||
      (record.customer && record.customer.companyName)
  );
  const owner = record.owner && typeof record.owner === "object" ? record.owner : null;
  const teamMembers = [];
  if (record.teamMembers && typeof record.teamMembers === "object") {
    if (Array.isArray(record.teamMembers.members)) {
      teamMembers.push(...record.teamMembers.members);
    }
    if (Array.isArray(record.teamMembers.customers)) {
      teamMembers.push(...record.teamMembers.customers);
    }
  }
  if (Array.isArray(record.members)) {
    teamMembers.push(...record.members);
  }
  const ownerName = fullName(owner);
  const ownerEmail = normalizeEmail(
    pickFirst(owner && (owner.email || owner.emailId || owner.userEmail))
  );
  const ownerUserId = pickFirst(owner && (owner.userId || owner.id || owner._id));
  const memberNames = dedupeStrings(teamMembers.map((member) => fullName(member)));
  const memberEmails = dedupeStrings(
    teamMembers.map((member) =>
      normalizeEmail(
        pickFirst(member && (member.email || member.emailId || member.userEmail))
      )
    )
  );
  const memberUserIds = dedupeStrings(
    teamMembers.map((member) => pickFirst(member && (member.userId || member.id || member._id)))
  );
  const customFieldValues = extractCustomFieldAliases(record);
  return {
    id,
    name,
    accountName,
    ownerName: ownerName || memberNames[0] || "",
    ownerEmail: ownerEmail || memberEmails[0] || "",
    ownerUserId: ownerUserId || memberUserIds[0] || "",
    memberEmails,
    memberUserIds,
    contractName: compactJoined(customFieldValues.contractName),
    hub: compactJoined(customFieldValues.hub),
    program: compactJoined(customFieldValues.program),
  };
}

function extractEmails(value, output, depth) {
  if (depth > 5 || value == null) {
    return;
  }
  if (typeof value === "string") {
    const email = normalizeEmail(value);
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      output.push(email);
    }
    return;
  }
  if (typeof value !== "object") {
    return;
  }
  if (Array.isArray(value)) {
    value.forEach((entry) => extractEmails(entry, output, depth + 1));
    return;
  }
  Object.keys(value).forEach((key) => extractEmails(value[key], output, depth + 1));
}

function dedupeStrings(values) {
  const seen = new Set();
  const result = [];
  values.forEach((value) => {
    const text = pickFirst(value);
    if (!text || seen.has(text)) {
      return;
    }
    seen.add(text);
    result.push(text);
  });
  return result;
}

function normalizeInvoiceRecord(record, project, fallbackAccountName) {
  if (!record || typeof record !== "object") {
    return null;
  }
  const projectInfo =
    project && typeof project === "object"
      ? project
      : {
          id: "",
          name: "",
          accountName: "",
          ownerName: "",
          ownerEmail: "",
          ownerUserId: "",
          memberEmails: [],
          memberUserIds: [],
          contractName: "",
          hub: "",
          program: "",
        };

  const invoiceNumber =
    pickFirst(
      buildInvoiceDisplayNumber(record) ||
        record.invoiceNumber ||
        record.invoiceNo ||
        record.invoiceId ||
        record.billNumber ||
        record.referenceNumber ||
        record.docNumber ||
        record.number
    ) || "";
  const invoiceName =
    pickFirst(
      record.invoiceName ||
        record.invoiceTitle ||
        record.name ||
        record.fileName ||
        record.title ||
        record.subject
    ) || (invoiceNumber ? `Invoice ${invoiceNumber}` : "Invoice");
  const invoiceDate =
    normalizeDateValue(
      record.invoiceDate ||
        record.dateOfIssue ||
        record.issuedDate ||
        record.issuedOn ||
        record.date ||
        record.approvedAt ||
        record.submittedAt ||
        record.createdAt ||
        record.updatedAt
    ) || normalizeDateValue(new Date().toISOString());
  const issueDate = invoiceDate;
  const dueDate = normalizeDateValue(
    record.dueDate ||
      record.dueOn ||
      record.paymentDueDate ||
      record.paymentDueOn ||
      record.expectedPaymentDate
  );
  const invoiceStatus = pickFirst(record.status || record.invoiceStatus || record.state || "Unknown");
  const amount = normalizeAmount(
    record.amount ||
      record.totalAmount ||
      record.netAmount ||
      record.grossAmount ||
      record.subTotal
  );
  const currencyCode = pickFirst(
    record.currencyCode || (record.currency && record.currency.currencyCode)
  );
  const currencySymbol = pickFirst(
    record.currencySymbol || (record.currency && record.currency.currencySymbol)
  );
  const pdfUrl = pickFirst(
    record.signedUrl ||
      record.downloadUrl ||
      record.fileUrl ||
      record.url ||
      record.href ||
      record.previewUrl ||
      record.attachmentUrl ||
      record.documentUrl ||
      (record.file && (record.file.signedUrl || record.file.downloadUrl || record.file.url))
  );

  const emails = [];
  extractEmails(record, emails, 0);
  const associatedEmails = dedupeStrings(emails.map((email) => normalizeEmail(email)));

  const associatedUserIds = dedupeStrings([
    record.userId,
    record.userID,
    record.ownerId,
    record.assigneeId,
    record.projectManagerId,
    record.expertAdvisorId,
    record.createdByUserId,
    record.submittedByUserId,
    record.approvedByUserId,
    record.submittedBy && record.submittedBy.id,
    record.submittedBy && record.submittedBy.userId,
    record.submittedByUser && record.submittedByUser.id,
    record.submittedByUser && record.submittedByUser.userId,
    extractUserIdValue(record.createdBy),
    extractUserIdValue(record.submittedBy),
    extractUserIdValue(record.updatedBy),
    record.createdBy && record.createdBy.id,
    record.createdBy && record.createdBy.userId,
    record.user && record.user.id,
    record.user && record.user.userId,
  ]);

  const projectUserIds = dedupeStrings([
    projectInfo.ownerUserId,
  ].concat(projectInfo.memberUserIds || []));
  const projectEmails = dedupeStrings([
    normalizeEmail(projectInfo.ownerEmail),
  ].concat(projectInfo.memberEmails || []));
  const customFieldValues = extractCustomFieldAliases(record);
  const lineItems = collectInvoiceLineItems(record);
  const lineItemQuantity = sumLineItemQuantity(lineItems);
  const quantityHoursFromFields = customFieldValues.quantityHours.reduce(
    (sum, value) => sum + normalizeAmount(value),
    0
  );
  const quantityHours = quantityHoursFromFields || lineItemQuantity;
  const contractName = compactJoined(
    customFieldValues.contractName.concat([projectInfo.contractName])
  );
  const hub = compactJoined(customFieldValues.hub.concat([projectInfo.hub]));
  const program = compactJoined(
    customFieldValues.program.concat([projectInfo.program])
  );
  const submittedByName = pickFirst(
    record.submittedByName ||
      fullName(record.submittedBy) ||
      fullName(record.submittedByUser) ||
      (record.submittedBy && record.submittedBy.name) ||
      (record.submittedByUser && record.submittedByUser.name)
  );
  const createdByFieldName = isLikelyDisplayName(customFieldValues.createdBy[0])
    ? pickFirst(customFieldValues.createdBy[0])
    : "";
  const createdByName =
    pickFirst(
      (isLikelyDisplayName(record.createdByName) ? record.createdByName : "") ||
        extractUserDisplayName(record.createdBy) ||
        extractUserDisplayName(record.submittedBy) ||
        extractUserDisplayName(record.createdByUser) ||
        fullName(record.createdBy) ||
        (record.createdByUser &&
          (record.createdByUser.name ||
            [pickFirst(record.createdByUser.firstName), pickFirst(record.createdByUser.lastName)]
              .filter(Boolean)
              .join(" ")))
    ) || createdByFieldName;

  if (!invoiceNumber && !invoiceName) {
    return null;
  }

  return {
    id:
      pickFirst(
        record.id ||
          record._id ||
          record.invoiceId ||
          record.documentId ||
          record.fileId ||
          record.invoiceNumber
      ) || `${projectInfo.id || projectInfo.name || "invoice"}-${invoiceNumber || invoiceName}`,
    invoiceNumber: invoiceNumber || `INV-${Math.random().toString(16).slice(2, 8).toUpperCase()}`,
    invoiceId: pickFirst(record.invoiceId || record.id || record._id),
    invoiceName,
    ownerName: pickFirst(
      createdByName ||
        createdByFieldName ||
        submittedByName ||
        record.projectManagerName ||
        record.expertAdvisorName ||
        record.pmName ||
        record.ownerName ||
        record.assigneeName ||
        fullName(record.createdBy) ||
        fullName(record.owner) ||
        (record.owner && record.owner.name)
    ) || projectInfo.ownerName || "Unassigned",
    accountName:
      pickPreferredAccountName([
        record.toCompany && (record.toCompany.companyName || record.toCompany.name),
        record.accountName,
        record.account && (record.account.accountName || record.account.name),
        record.customer && (record.customer.accountName || record.customer.companyName || record.customer.name),
        record.company && (record.company.companyName || record.company.name),
        record.companyName,
        customFieldValues.accountName[0],
        projectInfo.accountName,
        fallbackAccountName,
        "Rocketlane Account",
      ]) || "Rocketlane Account",
    invoiceDate,
    issueDate,
    dueDate,
    invoiceStatus,
    amount,
    currencyCode,
    currencySymbol,
    pdfUrl,
    createdByName,
    createdByUserId: pickFirst(
      record.createdByUserId ||
        extractUserIdValue(record.createdBy) ||
        extractUserIdValue(record.createdByUser) ||
        extractUserIdValue(record.submittedBy) ||
        (record.createdBy && (record.createdBy.userId || record.createdBy.id)) ||
        (record.createdByUser && (record.createdByUser.userId || record.createdByUser.id))
    ),
    createdByEmail: normalizeEmail(
      pickFirst(
        extractUserEmailValue(record.createdBy) ||
          extractUserEmailValue(record.createdByUser) ||
          extractUserEmailValue(record.submittedBy) ||
        (record.createdBy &&
          (record.createdBy.email || record.createdBy.emailId || record.createdBy.userEmail)) ||
          (record.createdByUser &&
            (record.createdByUser.email ||
              record.createdByUser.emailId ||
              record.createdByUser.userEmail))
      )
    ),
    submittedByUserId: pickFirst(
      record.submittedByUserId ||
        (record.submittedBy && (record.submittedBy.userId || record.submittedBy.id)) ||
        (record.submittedByUser && (record.submittedByUser.userId || record.submittedByUser.id))
    ),
    submittedByEmail: normalizeEmail(
      pickFirst(
        (record.submittedBy &&
          (record.submittedBy.email || record.submittedBy.emailId || record.submittedBy.userEmail)) ||
          (record.submittedByUser &&
            (record.submittedByUser.email ||
              record.submittedByUser.emailId ||
              record.submittedByUser.userEmail))
      )
    ),
    contractName,
    hub,
    program,
    quantityHours,
    associatedEmails: dedupeStrings(associatedEmails.concat(projectEmails)),
    associatedUserIds: dedupeStrings(associatedUserIds.concat(projectUserIds)),
    sourceProjectId: projectInfo.id || "",
    sourceProjectName: projectInfo.name || "",
  };
}

function normalizeInvoicePreview(invoiceRecord, lineRecords, paymentRecords) {
  const previewFieldSources = [
    invoiceRecord && invoiceRecord.fields,
    invoiceRecord && invoiceRecord.customFields,
    invoiceRecord && invoiceRecord.customFieldValues,
    invoiceRecord && invoiceRecord.fieldValues,
    invoiceRecord && invoiceRecord.projectFields,
    invoiceRecord && invoiceRecord.invoiceFields,
  ];
  const previewFields = mergeFieldDisplayEntries(
    mergeFieldDisplayEntries(
      previewFieldSources.reduce(
        (acc, source) => mergeFieldDisplayEntries(acc, extractFieldDisplayEntries(source)),
        []
      ),
      extractFieldDisplayEntries(
        invoiceRecord && invoiceRecord.additionalFields ? invoiceRecord.additionalFields : []
      )
    ),
    extractFieldDisplayEntries(invoiceRecord || {})
  ).filter((entry) => pickFirst(entry && entry.value));
  const invoiceNumber = pickFirst(invoiceRecord && invoiceRecord.invoiceNumber);
  const projectName = pickFirst(
    invoiceRecord &&
      Array.isArray(invoiceRecord.projects) &&
      invoiceRecord.projects[0] &&
      (invoiceRecord.projects[0].projectName || invoiceRecord.projects[0].name)
  );
  const submitter =
    (invoiceRecord && invoiceRecord.submittedBy) ||
    (invoiceRecord && invoiceRecord.submittedByUser) ||
    (invoiceRecord && invoiceRecord.createdBy) ||
    null;
  return {
    invoiceId: pickFirst(invoiceRecord && (invoiceRecord.invoiceId || invoiceRecord.id || invoiceRecord._id)),
    invoiceNumber,
    status: pickFirst(invoiceRecord && (invoiceRecord.status || invoiceRecord.invoiceStatus)),
    amount: normalizeAmount(invoiceRecord && (invoiceRecord.amount || invoiceRecord.totalAmount || invoiceRecord.subTotal)),
    currencyCode: pickFirst(
      invoiceRecord &&
        (invoiceRecord.currencyCode ||
          (invoiceRecord.currency && invoiceRecord.currency.currencyCode))
    ),
    currencySymbol: pickFirst(
      invoiceRecord &&
        (invoiceRecord.currencySymbol ||
          (invoiceRecord.currency && invoiceRecord.currency.currencySymbol))
    ),
    issueDate: normalizeDateValue(invoiceRecord && (invoiceRecord.dateOfIssue || invoiceRecord.invoiceDate || invoiceRecord.createdAt)),
    dueDate: normalizeDateValue(invoiceRecord && invoiceRecord.dueDate),
    accountName: pickFirst(
      invoiceRecord &&
        (invoiceRecord.accountName ||
          invoiceRecord.companyName ||
          (invoiceRecord.company && (invoiceRecord.company.companyName || invoiceRecord.company.name)))
    ),
    projectName,
    fromName: pickFirst(
      invoiceRecord &&
        (invoiceRecord.workspaceName ||
          invoiceRecord.accountName ||
          (invoiceRecord.company && (invoiceRecord.company.workspaceName || invoiceRecord.company.companyName)))
    ),
    billToName:
      pickFirst(invoiceRecord && (invoiceRecord.submittedByName || invoiceRecord.createdByName)) ||
      fullName(submitter),
    billToEmail: normalizeEmail(
      pickFirst(submitter && (submitter.email || submitter.emailId || submitter.userEmail))
    ),
    customFields: previewFields,
    allFields: previewFields,
    lineItems: Array.isArray(lineRecords)
      ? lineRecords.map((line) => ({
          id: pickFirst(line && (line.invoiceLineItemId || line.id || line._id)),
          description: pickFirst(line && line.description),
          quantity: normalizeAmount(line && (line.quantity || line.qty || line.hours || line.qtyHours)),
          unitPrice: normalizeAmount(line && line.unitPrice),
          amount: normalizeAmount(line && line.amount),
          fields: extractFieldDisplayEntries(line && line.fields),
        }))
      : [],
    payments: Array.isArray(paymentRecords)
      ? paymentRecords.map((payment) => ({
          id: pickFirst(payment && (payment.paymentId || payment.id || payment._id)),
          recordType: pickFirst(payment && (payment.paymentRecordType || payment.type || payment.status)),
          paymentDate: normalizeDateValue(payment && (payment.paymentDate || payment.date || payment.createdAt)),
          amount: normalizeAmount(payment && payment.amount),
          notes: pickFirst(payment && payment.notes),
        }))
      : [],
  };
}

function invoiceBelongsToProject(record, project) {
  if (!record || typeof record !== "object" || !project) {
    return false;
  }
  const targetId = pickFirst(project.id);
  const targetName = normalizeProjectName(project.name);
  const candidates = [];

  if (Array.isArray(record.projects)) {
    candidates.push(...record.projects);
  }
  if (record.project && typeof record.project === "object") {
    candidates.push(record.project);
  }
  if (record.projects && typeof record.projects === "object" && !Array.isArray(record.projects)) {
    candidates.push(record.projects);
  }

  for (let i = 0; i < candidates.length; i += 1) {
    const item = candidates[i] || {};
    const projectId = pickFirst(
      item.projectId || item.id || item._id || item.projectID || item.value
    );
    const projectName = normalizeProjectName(
      pickFirst(item.projectName || item.name || item.projectTitle || item.label)
    );
    if (targetId && projectId && targetId === projectId) {
      return true;
    }
    if (targetName && projectName && projectName.includes(targetName)) {
      return true;
    }
  }

  const directProjectId = pickFirst(record.projectId || record.projectID);
  if (targetId && directProjectId && targetId === directProjectId) {
    return true;
  }
  const directProjectName = normalizeProjectName(
    pickFirst(record.projectName || record.projectTitle)
  );
  if (targetName && directProjectName && directProjectName.includes(targetName)) {
    return true;
  }

  return false;
}

function normalizeMember(record) {
  if (!record || typeof record !== "object") {
    return null;
  }
  const nestedUser = record.user && typeof record.user === "object" ? record.user : null;
  const nestedProfile =
    record.profile && typeof record.profile === "object" ? record.profile : null;
  const nestedPermission =
    record.permission && typeof record.permission === "object" ? record.permission : null;
  const email = normalizeEmail(
    pickFirst(
      record.email ||
        record.userEmail ||
        record.workEmail ||
        (nestedUser && (nestedUser.email || nestedUser.userEmail || nestedUser.workEmail)) ||
        (nestedProfile && (nestedProfile.email || nestedProfile.workEmail))
    )
  );
  const id = pickFirst(
    record.id ||
      record.userId ||
      record._id ||
      (nestedUser && (nestedUser.id || nestedUser.userId || nestedUser._id))
  );
  if (!email && !id) {
    return null;
  }
  const permissionList = Array.isArray(record.permissions)
    ? record.permissions
        .map((entry) =>
          pickFirst(
            entry &&
              (entry.permissionName ||
                entry.name ||
                entry.permission ||
                entry.label)
          )
        )
        .filter(Boolean)
        .join(", ")
    : "";
  return {
    id,
    email,
    name: pickFirst(
      record.name ||
        record.displayName ||
        record.fullName ||
        record.userName ||
        (nestedUser &&
          (nestedUser.name ||
            nestedUser.displayName ||
            nestedUser.fullName ||
            nestedUser.userName)) ||
        [pickFirst(record.firstName), pickFirst(record.lastName)].filter(Boolean).join(" ")
    ),
    permission: pickFirst(
      (nestedPermission &&
        (nestedPermission.permissionName || nestedPermission.name)) ||
        (nestedUser &&
          nestedUser.permission &&
          (nestedUser.permission.permissionName || nestedUser.permission.name)) ||
        (nestedUser &&
          nestedUser.permissionSet &&
          (nestedUser.permissionSet.permissionName || nestedUser.permissionSet.name)) ||
        (nestedUser &&
          nestedUser.permissionSetObj &&
          (nestedUser.permissionSetObj.permissionName || nestedUser.permissionSetObj.name)) ||
        (nestedUser &&
          nestedUser.accountPermission &&
          (nestedUser.accountPermission.permissionName || nestedUser.accountPermission.name)) ||
        (nestedUser &&
          nestedUser.accountPermissionSet &&
          (nestedUser.accountPermissionSet.permissionName ||
            nestedUser.accountPermissionSet.name)) ||
        record.permission ||
        record.permissionSet ||
        (record.permissionSet &&
          (record.permissionSet.permissionName || record.permissionSet.name)) ||
        (record.accountPermission &&
          (record.accountPermission.permissionName || record.accountPermission.name)) ||
        (record.accountPermissionSet &&
          (record.accountPermissionSet.permissionName || record.accountPermissionSet.name)) ||
        (record.permissionSetObj &&
          (record.permissionSetObj.permissionName || record.permissionSetObj.name)) ||
        permissionList
    ) || extractPermissionFromAny(record),
    roleLabel: pickFirst(
      (record.role && (record.role.roleName || record.role.name)) ||
        (nestedUser &&
          nestedUser.role &&
          (nestedUser.role.roleName || nestedUser.role.name)) ||
        record.role ||
        record.userRole ||
        record.designation ||
        record.title
    ),
  };
}

function buildMemberLookups(members) {
  const byId = new Map();
  const byEmail = new Map();
  (Array.isArray(members) ? members : []).forEach((member) => {
    if (!member || typeof member !== "object") {
      return;
    }
    const id = pickFirst(member.id);
    const email = normalizeEmail(member.email);
    const name = pickFirst(member.name);
    if (id && name) {
      byId.set(id, name);
    }
    if (email && name) {
      byEmail.set(email, name);
    }
  });
  return { byId, byEmail };
}

function resolveMemberDisplayNameFromInvoice(invoice, lookups) {
  const memberLookups = lookups || { byId: new Map(), byEmail: new Map() };
  const idCandidates = dedupeStrings([
    pickFirst(invoice.createdByUserId),
    pickFirst(invoice.submittedByUserId),
  ].concat(Array.isArray(invoice.associatedUserIds) ? invoice.associatedUserIds : []));
  for (let i = 0; i < idCandidates.length; i += 1) {
    const id = idCandidates[i];
    if (id && memberLookups.byId.has(id)) {
      return pickFirst(memberLookups.byId.get(id));
    }
  }
  const emailCandidates = dedupeStrings([
    normalizeEmail(invoice.createdByEmail),
    normalizeEmail(invoice.submittedByEmail),
  ].concat(Array.isArray(invoice.associatedEmails) ? invoice.associatedEmails : []));
  for (let i = 0; i < emailCandidates.length; i += 1) {
    const email = normalizeEmail(emailCandidates[i]);
    if (email && memberLookups.byEmail.has(email)) {
      return pickFirst(memberLookups.byEmail.get(email));
    }
  }
  return "";
}

function resolveProjectAccountForInvoice(invoice, projectLookup) {
  const lookup = projectLookup || { byId: new Map(), byCanonicalName: new Map() };
  const sourceProjectId = pickFirst(invoice && invoice.sourceProjectId);
  if (sourceProjectId && lookup.byId.has(sourceProjectId)) {
    const project = lookup.byId.get(sourceProjectId) || {};
    const account = pickFirst(project.accountName);
    if (!isGenericAccountName(account)) {
      return account;
    }
  }
  const sourceProjectName = pickFirst(invoice && invoice.sourceProjectName);
  const sourceProjectCanonical = canonicalProjectName(sourceProjectName);
  if (sourceProjectCanonical && lookup.byCanonicalName.has(sourceProjectCanonical)) {
    const project = lookup.byCanonicalName.get(sourceProjectCanonical) || {};
    const account = pickFirst(project.accountName);
    if (!isGenericAccountName(account)) {
      return account;
    }
  }
  return "";
}

function enrichInvoiceDisplayData(invoice, memberLookups, projectLookup) {
  if (!invoice || typeof invoice !== "object") {
    return invoice;
  }
  const next = Object.assign({}, invoice);
  const ownerNeedsOverride =
    !isLikelyDisplayName(next.ownerName) || normalizeFieldLabel(next.ownerName) === "unassigned";
  if (ownerNeedsOverride) {
    const fromMember = resolveMemberDisplayNameFromInvoice(next, memberLookups);
    const fallbackOwner = pickFirst(next.createdByName || next.ownerName);
    const resolvedOwner = pickFirst(fromMember || fallbackOwner);
    if (resolvedOwner) {
      next.ownerName = resolvedOwner;
    }
  }
  if (isGenericAccountName(next.accountName)) {
    const projectAccount = resolveProjectAccountForInvoice(next, projectLookup);
    if (projectAccount) {
      next.accountName = projectAccount;
    }
  }
  return next;
}

function extractUserIdValue(value) {
  if (value == null) {
    return "";
  }
  if (typeof value === "number") {
    return String(value);
  }
  if (typeof value === "string") {
    return /^\d+$/.test(value.trim()) ? value.trim() : "";
  }
  if (typeof value === "object") {
    return pickFirst(value.userId || value.id || value._id);
  }
  return "";
}

function extractUserEmailValue(value) {
  if (!value || typeof value !== "object") {
    return "";
  }
  return normalizeEmail(pickFirst(value.email || value.emailId || value.userEmail || value.workEmail));
}

function extractUserDisplayName(value) {
  if (!value || typeof value !== "object") {
    return "";
  }
  const direct = fullName(value) || pickFirst(value.userName || value.displayName || value.name);
  return isLikelyDisplayName(direct) ? direct : "";
}

function enrichInvoiceFromDetailPayload(invoice, detailRecord, projectLookup) {
  if (!invoice || typeof invoice !== "object") {
    return invoice;
  }
  const next = Object.assign({}, invoice);
  const detail = detailRecord && typeof detailRecord === "object" ? detailRecord : {};
  const creator = detail.createdBy;
  const submitter = detail.submittedBy || detail.submittedByUser;

  const creatorUserId = pickFirst(
    extractUserIdValue(creator) ||
      extractUserIdValue(submitter) ||
      detail.createdByUserId ||
      detail.submittedByUserId
  );
  const creatorEmail = normalizeEmail(
    pickFirst(
      extractUserEmailValue(creator) ||
        extractUserEmailValue(submitter) ||
        detail.createdByEmail ||
        detail.submittedByEmail
    )
  );
  const creatorName = pickFirst(
    extractUserDisplayName(creator) ||
      extractUserDisplayName(submitter) ||
      detail.createdByName ||
      detail.submittedByName
  );

  if (creatorUserId) {
    next.createdByUserId = creatorUserId;
    next.associatedUserIds = dedupeStrings(
      [creatorUserId].concat(Array.isArray(next.associatedUserIds) ? next.associatedUserIds : [])
    );
  }
  if (creatorEmail) {
    next.createdByEmail = creatorEmail;
    next.associatedEmails = dedupeStrings(
      [creatorEmail].concat(Array.isArray(next.associatedEmails) ? next.associatedEmails : [])
    );
  }
  if (creatorName) {
    next.createdByName = creatorName;
    if (!isLikelyDisplayName(next.ownerName) || normalizeFieldLabel(next.ownerName) === "unassigned") {
      next.ownerName = creatorName;
    }
  }

  const detailAccountName = pickPreferredAccountName([
    detail.accountName,
    detail.company && (detail.company.companyName || detail.company.name),
    detail.toCompany && (detail.toCompany.companyName || detail.toCompany.name),
    detail.customer && (detail.customer.companyName || detail.customer.name),
    detail.account && (detail.account.accountName || detail.account.companyName || detail.account.name),
  ]);
  if (detailAccountName && !isGenericAccountName(detailAccountName)) {
    next.accountName = detailAccountName;
  }

  const detailProjects = Array.isArray(detail.projects) ? detail.projects : [];
  if (detailProjects.length) {
    const detailProject = detailProjects[0] || {};
    const detailProjectId = pickFirst(
      detailProject.projectId || detailProject.id || detailProject._id || detailProject.projectID
    );
    const detailProjectName = pickFirst(
      detailProject.projectName || detailProject.name || detailProject.projectTitle
    );
    if (detailProjectId) {
      next.sourceProjectId = detailProjectId;
    }
    if (detailProjectName) {
      next.sourceProjectName = detailProjectName;
    }
  }

  if (isGenericAccountName(next.accountName)) {
    const projectAccount = resolveProjectAccountForInvoice(next, projectLookup);
    if (projectAccount) {
      next.accountName = projectAccount;
    }
  }

  return next;
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
  const directLabel = pickFirst(
    value.name ||
      value.label ||
      value.displayName ||
      value.title ||
      value.value ||
      value.role ||
      value.permission
  );
  if (directLabel) {
    target.push(directLabel);
  }
  Object.keys(value).forEach((key) => {
    const lowered = key.toLowerCase();
    if (
      lowered.includes("role") ||
      lowered.includes("permission") ||
      lowered.includes("admin") ||
      lowered.includes("owner") ||
      lowered.includes("type") ||
      lowered.includes("group") ||
      lowered === "data" ||
      lowered === "response" ||
      lowered === "result" ||
      lowered === "payload" ||
      lowered === "user" ||
      lowered === "account"
    ) {
      collectRoleTokens(value[key], target, depth + 1);
    }
  });
}

function isAdminToken(text) {
  const value = String(text || "")
    .toLowerCase()
    .replace(/[_-]+/g, " ");
  if (!value) {
    return false;
  }
  return (
    /(^|\b)account\s*admin(istrator)?(\b|$)/.test(value) ||
    value === "admin" ||
    value === "administrator"
  );
}

function collectPermissionTokens(value, target, depth) {
  if (depth > 6 || value == null) {
    return;
  }
  if (typeof value === "string" || typeof value === "number") {
    const text = pickFirst(value);
    if (text) {
      target.push(text);
    }
    return;
  }
  if (Array.isArray(value)) {
    value.forEach((item) => collectPermissionTokens(item, target, depth + 1));
    return;
  }
  if (typeof value !== "object") {
    return;
  }
  Object.keys(value).forEach((key) => {
    const lowered = String(key || "").toLowerCase();
    if (
      lowered.includes("permission") ||
      lowered.includes("access") ||
      lowered.includes("role")
    ) {
      collectPermissionTokens(value[key], target, depth + 1);
    }
  });
}

function extractPermissionFromAny(record) {
  if (!record || typeof record !== "object") {
    return "";
  }
  const tokens = [];
  collectPermissionTokens(record, tokens, 0);
  const deduped = dedupeStrings(tokens);
  const adminToken = deduped.find((token) => isAdminToken(token));
  if (adminToken) {
    return adminToken;
  }
  return pickFirst(deduped[0]);
}

function deriveViewerAccess(request, context) {
  const ctxUser = (context && context.user) || {};
  const ctxAccount = (context && context.account) || {};
  const viewerContext = (request && request.viewerContext) || {};
  const id = pickFirst(
    ctxUser.id || ctxUser.userId || ctxUser._id || viewerContext.userId
  );
  const email = normalizeEmail(
    pickFirst(
      ctxUser.email ||
        ctxUser.emailId ||
        ctxUser.userEmail ||
        viewerContext.userEmail
    )
  );
  const displayName =
    fullName(ctxUser) ||
    pickFirst(
      ctxUser.name || ctxUser.displayName || ctxUser.userName || viewerContext.userName
    );
  const permission = pickFirst(
    (ctxUser.permission &&
      (ctxUser.permission.permissionName || ctxUser.permission.name)) ||
      ctxUser.permission ||
      ctxUser.permissionSet ||
      (ctxUser.permissionSet &&
        (ctxUser.permissionSet.permissionName || ctxUser.permissionSet.name)) ||
      (ctxUser.permissionSetObj &&
        (ctxUser.permissionSetObj.permissionName || ctxUser.permissionSetObj.name)) ||
      (ctxUser.accountPermission &&
        (ctxUser.accountPermission.permissionName || ctxUser.accountPermission.name)) ||
      (ctxUser.accountPermissionSet &&
        (ctxUser.accountPermissionSet.permissionName ||
          ctxUser.accountPermissionSet.name)) ||
      (Array.isArray(ctxUser.permissions) &&
        ctxUser.permissions
          .map((entry) =>
            pickFirst(
              entry &&
                (entry.permissionName ||
                  entry.name ||
                  entry.permission ||
                  entry.label)
            )
          )
          .filter(Boolean)
          .join(", ")) ||
      (viewerContext.permission &&
        (viewerContext.permission.permissionName || viewerContext.permission.name)) ||
      viewerContext.permission
  ) || extractPermissionFromAny(viewerContext) || extractPermissionFromAny(ctxUser);
  const roleLabel = pickFirst(
    ctxUser.role ||
      ctxUser.userRole ||
      ctxUser.type ||
      ctxUser.userType ||
      ctxUser.designation ||
      viewerContext.userRole
  );

  const isAdmin = isAdminToken(permission);

  return {
    id,
    email,
    displayName,
    permission,
    roleLabel,
    isAdmin,
  };
}

function canonicalIdentityName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function extractEmailLocalPart(email) {
  const text = normalizeEmail(email);
  const index = text.indexOf("@");
  return index > 0 ? text.slice(0, index) : "";
}

function resolveViewerFromMembers(baseViewer, members, request) {
  const viewer = baseViewer && typeof baseViewer === "object" ? baseViewer : {};
  const list = Array.isArray(members) ? members : [];
  if (!list.length) {
    return viewer;
  }
  const viewerContext = (request && request.viewerContext) || {};
  const targetId = pickFirst(viewer.id || viewerContext.userId);
  const targetEmail = normalizeEmail(pickFirst(viewer.email || viewerContext.userEmail));
  const targetName = canonicalIdentityName(
    pickFirst(viewer.displayName || viewerContext.userName)
  );
  const targetEmailLocalPart = extractEmailLocalPart(targetEmail);

  const matched = list.find((member) => {
    if (!member || typeof member !== "object") {
      return false;
    }
    const memberId = pickFirst(member.id);
    const memberEmail = normalizeEmail(pickFirst(member.email));
    const memberName = canonicalIdentityName(pickFirst(member.name));
    if (targetId && memberId && targetId === memberId) {
      return true;
    }
    if (targetEmail && memberEmail && targetEmail === memberEmail) {
      return true;
    }
    if (targetName && memberName && (targetName === memberName || targetName.includes(memberName) || memberName.includes(targetName))) {
      return true;
    }
    const memberEmailLocalPart = extractEmailLocalPart(memberEmail);
    if (targetEmailLocalPart && memberEmailLocalPart && targetEmailLocalPart === memberEmailLocalPart) {
      return true;
    }
    return false;
  });

  if (!matched) {
    return viewer;
  }
  const permission = pickFirst(
    matched.permission ||
      viewer.permission ||
      (isAdminToken(matched.roleLabel) ? matched.roleLabel : "")
  );
  const roleLabel = pickFirst(matched.roleLabel || viewer.roleLabel);
  const isAdmin = isAdminToken(permission) || viewer.isAdmin === true;
  return {
    id: pickFirst(matched.id || viewer.id),
    email: normalizeEmail(pickFirst(matched.email || viewer.email)),
    displayName: pickFirst(matched.name || viewer.displayName),
    permission,
    roleLabel,
    isAdmin,
  };
}

module.exports = {
  syncInvoicesFromSource: async (request = {}, context = {}) => {
    const installation = context.installation || {};
    const iParams = installation.iparams || {};
    const secureParams = installation.secureParams || {};
    const workspaceCandidates = dedupeStrings([
      request.workspaceBaseUrl,
      ...(Array.isArray(request.workspaceCandidates) ? request.workspaceCandidates : []),
      request.viewerContext && request.viewerContext.workspaceBaseUrl,
      iParams.workspaceBaseUrl,
      iParams.workspaceUrl,
      secureParams.workspaceBaseUrl,
      secureParams.workspaceUrl,
    ]);
    const embeddedWorkspaceHost = workspaceHost(
      `https://${String(EMBEDDED_ROCKETLANE_API_KEY_WORKSPACE || "").trim()}`
    );
    const embeddedWorkspaceBaseUrl = embeddedWorkspaceHost
      ? `https://${embeddedWorkspaceHost}`
      : "";
    const workspaceCandidatesForFetch = embeddedWorkspaceHost
      ? dedupeStrings([
          embeddedWorkspaceBaseUrl,
          ...workspaceCandidates.filter((candidate) =>
            hostMatchesTarget(workspaceHost(candidate), embeddedWorkspaceHost)
          ),
        ])
      : workspaceCandidates;
    const tokenCandidates = [
      // Hard-prioritize embedded key for the configured workspace so
      // stale install params cannot redirect invoice reads to another account.
      { value: EMBEDDED_ROCKETLANE_API_KEY, source: "embedded" },
      { value: request.apiToken, source: "request.apiToken" },
      {
        value:
          secureParams.rocketlaneApiToken || secureParams.apiToken || secureParams.apiKey,
        source: "installation.secureParams",
      },
      {
        value: iParams.rocketlaneApiToken || iParams.apiToken,
        source: "installation.iparams",
      },
      { value: context.apiKey, source: "context.apiKey" },
    ];
    const selectedToken = tokenCandidates.find((candidate) => pickFirst(candidate.value));
    const apiToken = selectedToken ? pickFirst(selectedToken.value) : "";

    const apiBaseCandidates = dedupeStrings([
      request.apiBaseUrl,
      secureParams.apiBaseUrl,
      iParams.apiBaseUrl,
      ROCKETLANE_API_BASE_URL,
    ]);
    const normalizedSearchQuery = String(request.searchQuery || "").trim();

    if (!apiBaseCandidates.length || !apiToken) {
      return {
        ok: false,
        error:
          "Missing workspace/API key configuration. Set EMBEDDED_ROCKETLANE_API_KEY in server-actions/sync-invoices-from-source.js or provide token via request/install settings.",
        invoices: [],
        sourceProjects: [],
        teamMembers: [],
      };
    }

    const headers = {
      Accept: "application/json",
      "api-key": apiToken,
    };

    const diagnostics = {
      workspaceCandidates: workspaceCandidatesForFetch,
      projectErrors: [],
      invoiceErrors: [],
      memberErrors: [],
      invoiceFetchMode: "account-wide-invoices-type-all",
      workspaceUsed: "",
      apiBaseUsed: "",
      hasApiToken: Boolean(apiToken),
      tokenSource: selectedToken ? selectedToken.source : "none",
      workspaceLockHost: embeddedWorkspaceHost,
      contextUserKeys: context && context.user ? Object.keys(context.user) : [],
      searchQuery: normalizedSearchQuery,
    };
    const viewer = deriveViewerAccess(request, context);
    let resolvedViewer = viewer;

    if (request.previewInvoiceId || request.previewInvoiceNumber) {
      for (let i = 0; i < apiBaseCandidates.length; i += 1) {
        const baseUrl = apiBaseCandidates[i];
        try {
          const resolvedInvoiceId = await resolveInvoiceIdForPreview(
            baseUrl,
            headers,
            request.previewInvoiceId,
            request.previewInvoiceNumber,
            request.previewSourceProjectId
          );
          const previewInvoiceId = encodeURIComponent(String(resolvedInvoiceId || ""));
          if (!previewInvoiceId) {
            throw new Error("Invoice ID could not be resolved for preview.");
          }
          let invoicePayload = null;
          try {
            invoicePayload = await requestJson(
              ensureAbsoluteUrl(baseUrl, `/api/v1/invoices/${previewInvoiceId}`),
              headers
            );
          } catch (_error) {
            invoicePayload = await requestJson(
              ensureAbsoluteUrl(baseUrl, `/api/1.0/invoices/${previewInvoiceId}`),
              headers
            );
          }
          const linePayload = await requestJson(
            ensureAbsoluteUrl(baseUrl, `/api/1.0/invoices/${previewInvoiceId}/lines`),
            headers
          );
          const paymentPayload = await requestJson(
            ensureAbsoluteUrl(baseUrl, `/api/1.0/invoices/${previewInvoiceId}/payments`),
            headers
          );
          const lineItems = extractCollection(linePayload, ["data", "lines", "items", "results"]);
          const payments = extractCollection(paymentPayload, [
            "data",
            "payments",
            "items",
            "results",
          ]);
          let generatedPdfDataUrl = "";
          try {
            const pdfBytes = await requestBinary(
              ensureAbsoluteUrl(baseUrl, `/api/v1/invoices/${previewInvoiceId}/generate`),
              mergeObjects(headers, { Accept: "*/*" })
            );
            generatedPdfDataUrl = bytesToPdfDataUrl(pdfBytes);
          } catch (_error) {
            generatedPdfDataUrl = "";
          }
          const invoiceRecord = extractRecordObject(invoicePayload || {}, [
            "invoice",
            "data",
            "result",
            "payload",
            "response",
          ]) || {};
          const preview = normalizeInvoicePreview(invoiceRecord, lineItems, payments);
          if (generatedPdfDataUrl) {
            preview.pdfDataUrl = generatedPdfDataUrl;
            preview.pdfSource = "api-v1-generate";
          }
          return {
            ok: true,
            preview,
            viewer,
            diagnostics: mergeObjects(diagnostics, {
              apiBaseUsed: baseUrl,
              previewInvoiceResolvedId: String(resolvedInvoiceId || ""),
            }),
          };
        } catch (error) {
          diagnostics.invoiceErrors.push(
            String(error && error.message ? error.message : error)
          );
        }
      }
      return {
        ok: false,
        error: "Unable to load invoice preview details.",
        preview: null,
        viewer,
        diagnostics,
      };
    }

    let sourceProjects = [];
    let invoices = [];
    let members = [];

    for (let w = 0; w < apiBaseCandidates.length; w += 1) {
      const baseUrl = apiBaseCandidates[w];

      const projectsResult = await requestCollection(
        baseUrl,
        headers,
        ["/api/1.0/projects"],
        ["projects", "data", "content", "results", "items"]
      );

      diagnostics.projectErrors.push(...projectsResult.errors);
      const allProjects = projectsResult.rows
        .map(normalizeProject)
        .filter(Boolean);
      const projectLookup = buildProjectLookup(allProjects);

      const allInvoicesResult = await requestCollection(
        baseUrl,
        headers,
        [
          "/api/v1/invoices",
          "/api/1.0/invoices",
        ],
        ["invoices", "data", "content", "results", "items"]
      );
      diagnostics.invoiceErrors.push(...allInvoicesResult.errors);
      const globalInvoices = allInvoicesResult.rows;

      const collectedInvoices = [];
      const quantityBackfillTasks = [];
      for (let idx = 0; idx < globalInvoices.length; idx += 1) {
        const row = globalInvoices[idx];
        const project = resolveProjectForInvoice(row, projectLookup);
        let normalized = normalizeInvoiceRecord(
          row,
          project,
          request.accountName || iParams.accountName || ""
        );
        if (normalized) {
          const invoiceId = encodeURIComponent(
            pickFirst(normalized.invoiceId || normalized.id || row.invoiceId || row.id || row._id)
          );
          if (Number(normalized.quantityHours || 0) <= 0) {
            if (invoiceId) {
              const targetInvoice = normalized;
              quantityBackfillTasks.push(
                requestCollection(
                  baseUrl,
                  headers,
                  [`/api/1.0/invoices/${invoiceId}/lines`],
                  ["data", "lines", "items", "results"]
                )
                  .then((lineFetch) => {
                    diagnostics.invoiceErrors.push(...lineFetch.errors);
                    const fetchedQuantity = sumLineItemQuantity(lineFetch.rows);
                    if (fetchedQuantity > 0) {
                      targetInvoice.quantityHours = fetchedQuantity;
                    }
                  })
                  .catch((lineError) => {
                    diagnostics.invoiceErrors.push(
                      String(lineError && lineError.message ? lineError.message : lineError)
                    );
                  })
              );
            }
          }
          collectedInvoices.push(normalized);
        }
      }
      if (quantityBackfillTasks.length) {
        await Promise.all(quantityBackfillTasks);
      }

      const membersResult = await requestCollection(
        baseUrl,
        headers,
        ["/api/1.0/users?includeFields=permission,role,company", "/api/1.0/users?includeFields=permission", "/api/1.0/users"],
        ["users", "members", "teamMembers", "data", "results", "items"]
      );
      diagnostics.memberErrors.push(...membersResult.errors);
      const normalizedMembers = membersResult.rows.map(normalizeMember).filter(Boolean);
      const viewerFromApi = await requestCurrentViewer(baseUrl, headers);
      const membersWithViewer = viewerFromApi
        ? [viewerFromApi].concat(normalizedMembers)
        : normalizedMembers;
      const memberLookups = buildMemberLookups(membersWithViewer);
      const enrichedInvoices = collectedInvoices.map((invoice) =>
        enrichInvoiceDisplayData(invoice, memberLookups, projectLookup)
      );
      const viewerSeed = mergeObjects(viewer, viewerFromApi || {});
      const resolvedLoopViewer = resolveViewerFromMembers(
        viewerSeed,
        membersWithViewer,
        request
      );

      if (allProjects.length || globalInvoices.length || normalizedMembers.length) {
        diagnostics.workspaceUsed = workspaceCandidatesForFetch[0] || "";
        diagnostics.apiBaseUsed = baseUrl;
        sourceProjects = dedupeStrings(
          enrichedInvoices.map((invoice) => pickFirst(invoice && invoice.sourceProjectName))
        );
        invoices = enrichedInvoices;
        members = normalizedMembers;
        resolvedViewer = resolvedLoopViewer;
        break;
      }
    }

    const dedupedInvoices = [];
    const seen = new Set();
    invoices.forEach((invoice) => {
      const key =
        `${invoice.invoiceNumber}|${invoice.sourceProjectName}|${invoice.id}`.toLowerCase();
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      dedupedInvoices.push(invoice);
    });
    const searchMatches = normalizedSearchQuery
      ? dedupedInvoices.filter((invoice) => invoiceMatchesQuery(invoice, normalizedSearchQuery))
      : dedupedInvoices;

    resolvedViewer = resolveViewerFromMembers(resolvedViewer, members, request);

    return {
      ok: true,
      sourceProjects,
      invoices: request.searchOnly ? searchMatches : dedupedInvoices,
      teamMembers: dedupeStrings(members.map((m) => `${m.email}|${m.id}`))
        .map((key) => members.find((m) => `${m.email}|${m.id}` === key))
        .filter(Boolean),
      search: {
        query: normalizedSearchQuery,
        totalInvoices: dedupedInvoices.length,
        matchedInvoices: searchMatches.length,
      },
      viewer: resolvedViewer,
      diagnostics,
    };
  },
};
