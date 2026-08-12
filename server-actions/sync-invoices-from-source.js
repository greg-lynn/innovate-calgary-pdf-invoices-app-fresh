"use strict";

const { PDFDocument, rgb } = require("pdf-lib");
const http = require("http");
const https = require("https");

const DEFAULT_SOURCE_PROJECTS = ["Expert Advisor Program Invoices"];
// Production override: embed API key here so app works without installer prompt.
// Replace before shipping to users if needed.
const EMBEDDED_ROCKETLANE_API_KEY = "rl-6657ce9e-ee84-465d-b4df-d97b1239a343";
const EMBEDDED_ROCKETLANE_API_KEY_WORKSPACE = "innovate-calgary.rocketlane.com";
const ROCKETLANE_API_BASE_URL = "https://api.rocketlane.com";
const REQUEST_TIMEOUT_MS = 30000;
const MAX_HTTP_REDIRECTS = 5;
const FIELD_ALIAS_GROUPS = {
  contractName: ["contract name", "contract", "contractname"],
  hub: ["hub"],
  program: ["program"],
  accountName: ["account"],
  createdBy: ["created by", "createdby", "creator"],
  quantityHours: ["quantity", "qty", "hour", "hours"],
};

const HUB_FALLBACK_LABEL_ALIASES = ["ea address", "address", "hub location", "location"];
const PROGRAM_FALLBACK_LABEL_ALIASES = [
  "program",
  "ea program",
  "expert advisor program",
];

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

function parseJsonObject(value) {
  const text = String(value || "").trim();
  if (!text || (text[0] !== "{" && text[0] !== "[")) {
    return null;
  }
  try {
    const parsed = JSON.parse(text);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch (_error) {
    return null;
  }
}

function normalizeLookupKey(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function collectNestedCandidates(value, maxDepth) {
  const limit = Number(maxDepth || 6);
  const queue = [{ node: value, depth: 0 }];
  const visited = typeof WeakSet === "function" ? new WeakSet() : null;
  const out = [];
  while (queue.length) {
    const current = queue.shift();
    if (!current) {
      continue;
    }
    const node = current.node;
    const depth = Number(current.depth || 0);
    if (depth > limit || node == null) {
      continue;
    }
    if (typeof node === "string") {
      const parsed = parseJsonObject(node);
      if (parsed) {
        queue.push({ node: parsed, depth: depth + 1 });
      }
      continue;
    }
    if (typeof node !== "object") {
      continue;
    }
    if (visited) {
      if (visited.has(node)) {
        continue;
      }
      visited.add(node);
    }
    out.push(node);
    if (Array.isArray(node)) {
      node.forEach((item) => queue.push({ node: item, depth: depth + 1 }));
      continue;
    }
    Object.keys(node).forEach((key) => {
      queue.push({ node: node[key], depth: depth + 1 });
    });
  }
  return out;
}

function pickFirstValueFromAny(root, keys) {
  const searchKeys = Array.isArray(keys) ? keys : [];
  if (!searchKeys.length) {
    return "";
  }
  const searchTokens = new Set(searchKeys.map((key) => normalizeLookupKey(key)).filter(Boolean));
  const nodes = collectNestedCandidates(root, 6);
  for (let i = 0; i < nodes.length; i += 1) {
    const node = nodes[i];
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      continue;
    }
    const nodeKeys = Object.keys(node);
    for (let j = 0; j < nodeKeys.length; j += 1) {
      const nodeKey = nodeKeys[j];
      if (!searchTokens.has(normalizeLookupKey(nodeKey))) {
        continue;
      }
      const value = pickFirst(node[nodeKey]);
      if (value) {
        return value;
      }
    }
  }
  return "";
}

function pickFirstArrayFromAny(root, keys) {
  const searchKeys = Array.isArray(keys) ? keys : [];
  if (!searchKeys.length) {
    return [];
  }
  const searchTokens = new Set(searchKeys.map((key) => normalizeLookupKey(key)).filter(Boolean));
  const nodes = collectNestedCandidates(root, 6);
  for (let i = 0; i < nodes.length; i += 1) {
    const node = nodes[i];
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      continue;
    }
    const nodeKeys = Object.keys(node);
    for (let j = 0; j < nodeKeys.length; j += 1) {
      const nodeKey = nodeKeys[j];
      if (!searchTokens.has(normalizeLookupKey(nodeKey))) {
        continue;
      }
      const value = node[nodeKey];
      if (Array.isArray(value)) {
        return value;
      }
    }
  }
  return [];
}

function pickFirstObjectFromAny(root, keys) {
  const searchKeys = Array.isArray(keys) ? keys : [];
  if (!searchKeys.length) {
    return null;
  }
  const searchTokens = new Set(searchKeys.map((key) => normalizeLookupKey(key)).filter(Boolean));
  const nodes = collectNestedCandidates(root, 6);
  for (let i = 0; i < nodes.length; i += 1) {
    const node = nodes[i];
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      continue;
    }
    const nodeKeys = Object.keys(node);
    for (let j = 0; j < nodeKeys.length; j += 1) {
      const nodeKey = nodeKeys[j];
      if (!searchTokens.has(normalizeLookupKey(nodeKey))) {
        continue;
      }
      const value = node[nodeKey];
      if (value && typeof value === "object" && !Array.isArray(value)) {
        return value;
      }
    }
  }
  return null;
}

function parseBooleanFromAny(root, keys, fallback) {
  const defaultValue = Boolean(fallback);
  const direct = pickFirstValueFromAny(root, keys);
  if (!direct) {
    return defaultValue;
  }
  const normalized = String(direct).trim().toLowerCase();
  if (normalized === "true" || normalized === "1" || normalized === "yes") {
    return true;
  }
  if (normalized === "false" || normalized === "0" || normalized === "no") {
    return false;
  }
  return defaultValue;
}

function normalizeIncomingRequest(request) {
  const source = request && typeof request === "object" ? request : {};
  const previewObject = pickFirstObjectFromAny(source, ["preview"]) || {};
  const viewerContext = mergeObjects(
    pickFirstObjectFromAny(source, ["viewerContext"]) || {},
    source.viewerContext && typeof source.viewerContext === "object" ? source.viewerContext : {}
  );
  const invoiceId = pickFirst(
    source.previewInvoiceId ||
      source.invoiceId ||
      previewObject.invoiceId ||
      pickFirstValueFromAny(source, ["previewInvoiceId", "invoiceId"])
  );
  const invoiceNumberForPreview = pickFirst(
    source.previewInvoiceNumber ||
      source.invoiceNumberForPreview ||
      source.invoiceNumber ||
      previewObject.invoiceNumber ||
      pickFirstValueFromAny(source, [
        "previewInvoiceNumber",
        "invoiceNumberForPreview",
        "invoiceNumber",
      ])
  );
  const prefetchInvoiceId = pickFirst(
    source.prefetchInvoiceId ||
      source.invoiceId ||
      source.previewInvoiceId ||
      previewObject.invoiceId ||
      pickFirstValueFromAny(source, [
        "prefetchInvoiceId",
        "previewInvoiceId",
        "invoiceId",
      ])
  );
  const prefetchInvoiceNumber = pickFirst(
    source.prefetchInvoiceNumber ||
      source.invoiceNumberForPreview ||
      source.previewInvoiceNumber ||
      source.invoiceNumber ||
      previewObject.invoiceNumber ||
      pickFirstValueFromAny(source, [
        "prefetchInvoiceNumber",
        "previewInvoiceNumber",
        "invoiceNumberForPreview",
        "invoiceNumber",
      ])
  );
  const workspaceCandidates = Array.isArray(source.workspaceCandidates)
    ? source.workspaceCandidates
    : pickFirstArrayFromAny(source, ["workspaceCandidates"]);
  return mergeObjects(source, {
    requestMode: pickFirst(
      source.requestMode ||
        source.mode ||
        pickFirstValueFromAny(source, ["requestMode", "mode"])
    ),
    mode: pickFirst(
      source.mode || source.requestMode || pickFirstValueFromAny(source, ["mode", "requestMode"])
    ),
    workspaceBaseUrl: pickFirst(
      source.workspaceBaseUrl ||
        source.workspaceUrl ||
        pickFirstValueFromAny(source, ["workspaceBaseUrl", "workspaceUrl"])
    ),
    workspaceCandidates,
    apiBaseUrl: pickFirst(
      source.apiBaseUrl || pickFirstValueFromAny(source, ["apiBaseUrl"])
    ),
    apiToken: pickFirst(
      source.apiToken ||
        pickFirstValueFromAny(source, ["apiToken", "rocketlaneApiToken", "apiKey"])
    ),
    accountName: pickFirst(
      source.accountName || pickFirstValueFromAny(source, ["accountName"])
    ),
    searchQuery: pickFirst(
      source.searchQuery || pickFirstValueFromAny(source, ["searchQuery"])
    ),
    invoiceId,
    previewInvoiceId: invoiceId,
    invoiceNumberForPreview,
    previewInvoiceNumber: invoiceNumberForPreview,
    prefetchInvoiceId,
    prefetchInvoiceNumber,
    previewSourceProjectId: pickFirst(
      source.previewSourceProjectId ||
        pickFirstValueFromAny(source, ["previewSourceProjectId"])
    ),
    searchOnly:
      source.searchOnly === true || parseBooleanFromAny(source, ["searchOnly"], false),
    prefetchPreviewPdfs:
      source.prefetchPreviewPdfs === true ||
      parseBooleanFromAny(source, ["prefetchPreviewPdfs"], false),
    disablePreviewMode:
      source.disablePreviewMode === true ||
      parseBooleanFromAny(source, ["disablePreviewMode"], false),
    viewerContext,
    preview: mergeObjects(previewObject, {
      invoiceId: pickFirst(previewObject.invoiceId || invoiceId),
      invoiceNumber: pickFirst(previewObject.invoiceNumber || invoiceNumberForPreview),
    }),
  });
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

function extractFieldValuesByLabelAliases(record, aliases) {
  if (!record || typeof record !== "object" || !Array.isArray(aliases) || !aliases.length) {
    return [];
  }
  const values = [];
  collectCustomFieldSources(record).forEach((source) => {
    extractFieldDisplayEntries(source).forEach((entry) => {
      if (!entry || typeof entry !== "object") {
        return;
      }
      if (!fieldLabelMatchesAlias(entry.label, aliases)) {
        return;
      }
      const value = toFieldText(entry.value);
      if (value) {
        values.push(value);
      }
    });
  });
  return dedupeStrings(values);
}

function deriveHubFromAddressText(text) {
  const raw = pickFirst(text);
  if (!raw) {
    return "";
  }
  const parts = raw
    .split(",")
    .map((part) => pickFirst(part))
    .filter(Boolean);
  if (parts.length < 2) {
    return "";
  }
  const cityCandidate = pickFirst(parts[1]);
  if (
    cityCandidate &&
    /^[A-Za-z][A-Za-z .'-]+$/.test(cityCandidate) &&
    !/^[A-Z]{2,3}$/.test(cityCandidate)
  ) {
    return cityCandidate;
  }
  return "";
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

function cleanExtractedFieldValue(value) {
  return pickFirst(value)
    .replace(/\s+/g, " ")
    .replace(/^[|,:;\-]+/g, "")
    .replace(/[|,:;\-]+$/g, "")
    .trim();
}

function extractHubProgramFromLineItems(lineItems) {
  const hubValues = [];
  const programValues = [];
  (Array.isArray(lineItems) ? lineItems : []).forEach((line) => {
    if (!line || typeof line !== "object") {
      return;
    }
    const description = pickFirst(
      line.description ||
        line.itemDescription ||
        line.itemName ||
        line.notes ||
        line.text ||
        line.name
    );
    if (description) {
      const hubPattern = /(?:^|\||\n|\r)\s*hub\s*:\s*([^|\n\r]+)/gi;
      const programPattern = /(?:^|\||\n|\r)\s*program\s*:\s*([^|\n\r]+)/gi;
      let match = null;
      while ((match = hubPattern.exec(description)) != null) {
        const parsed = cleanExtractedFieldValue(match[1]);
        if (parsed) {
          hubValues.push(parsed);
        }
      }
      while ((match = programPattern.exec(description)) != null) {
        const parsed = cleanExtractedFieldValue(match[1]);
        if (parsed) {
          programValues.push(parsed);
        }
      }
    }
    const lineFieldValues = extractNamedCustomFieldValues(line.fields);
    if (Array.isArray(lineFieldValues.hub)) {
      lineFieldValues.hub.forEach((value) => {
        const parsed = cleanExtractedFieldValue(value);
        if (parsed) {
          hubValues.push(parsed);
        }
      });
    }
    if (Array.isArray(lineFieldValues.program)) {
      lineFieldValues.program.forEach((value) => {
        const parsed = cleanExtractedFieldValue(value);
        if (parsed) {
          programValues.push(parsed);
        }
      });
    }
  });
  return {
    hub: dedupeStrings(hubValues),
    program: dedupeStrings(programValues),
  };
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
  const response = await requestBuffer(url, {
    method: "GET",
    headers,
    timeoutMs: REQUEST_TIMEOUT_MS,
  });
  if (!response.ok) {
    throw new Error(`Request failed (${response.status}) for ${url}`);
  }
  const text = bufferToUtf8Text(response.body);
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
  const response = await requestBuffer(url, {
    method: "GET",
    headers,
    timeoutMs: REQUEST_TIMEOUT_MS,
  });
  if (!response.ok) {
    throw new Error(`Request failed (${response.status}) for ${url}`);
  }
  const body = response.body;
  if (!body || !body.length) {
    return null;
  }
  return new Uint8Array(body);
}

async function requestBinaryWithMethods(url, headers, methods) {
  const methodList = Array.isArray(methods) && methods.length ? methods : ["GET"];
  let lastError = null;
  for (let i = 0; i < methodList.length; i += 1) {
    const method = String(methodList[i] || "GET").toUpperCase();
    try {
      const response = await requestBuffer(url, {
        method,
        headers,
        timeoutMs: REQUEST_TIMEOUT_MS,
      });
      if (!response.ok) {
        throw new Error(`Request failed (${response.status}) for ${url}`);
      }
      const body = response.body;
      if (!body || !body.length) {
        throw new Error(`Empty binary payload for ${url}`);
      }
      return new Uint8Array(body);
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error(`Unable to fetch binary payload for ${url}`);
}

function hasFetchRuntime() {
  return typeof fetch === "function";
}

function normalizeRequestHeaders(headers) {
  const output = {};
  if (!headers || typeof headers !== "object") {
    return output;
  }
  Object.keys(headers).forEach((key) => {
    if (!Object.prototype.hasOwnProperty.call(headers, key)) {
      return;
    }
    const value = headers[key];
    if (value == null) {
      return;
    }
    output[String(key)] = String(value);
  });
  if (!output["Accept"] && !output["accept"]) {
    output.Accept = "application/json";
  }
  return output;
}

function isRedirectStatus(statusCode) {
  return statusCode === 301 || statusCode === 302 || statusCode === 303 || statusCode === 307 || statusCode === 308;
}

function bufferToUtf8Text(buffer) {
  if (!buffer) {
    return "";
  }
  try {
    return Buffer.from(buffer).toString("utf8");
  } catch (_error) {
    return "";
  }
}

async function requestBuffer(url, options) {
  const settings = options && typeof options === "object" ? options : {};
  if (hasFetchRuntime()) {
    try {
      const response = await fetch(url, {
        method: String(settings.method || "GET").toUpperCase(),
        headers: normalizeRequestHeaders(settings.headers),
      });
      const bodyArrayBuffer = await response.arrayBuffer();
      const body = Buffer.from(bodyArrayBuffer || new ArrayBuffer(0));
      return {
        ok: response.ok,
        status: Number(response.status || 0),
        body,
      };
    } catch (_error) {
      // Fall back to the Node transport when fetch is unavailable or blocked.
    }
  }
  return requestBufferViaNode(url, settings, 0);
}

function requestBufferViaNode(url, options, redirectCount) {
  const settings = options && typeof options === "object" ? options : {};
  const redirects = Number(redirectCount || 0);
  return new Promise((resolve, reject) => {
    let parsed = null;
    try {
      parsed = new URL(String(url || ""));
    } catch (_error) {
      reject(new Error(`Invalid URL ${url}`));
      return;
    }
    const method = String(settings.method || "GET").toUpperCase();
    const headers = normalizeRequestHeaders(settings.headers);
    const transport = parsed.protocol === "http:" ? http : https;
    const request = transport.request(
      {
        protocol: parsed.protocol,
        hostname: parsed.hostname,
        port: parsed.port || undefined,
        path: `${parsed.pathname || ""}${parsed.search || ""}`,
        method,
        headers,
      },
      (response) => {
        const chunks = [];
        response.on("data", (chunk) => {
          chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
        });
        response.on("end", async () => {
          const statusCode = Number(response.statusCode || 0);
          const location = pickFirst(response.headers && response.headers.location);
          if (isRedirectStatus(statusCode) && location) {
            if (redirects >= MAX_HTTP_REDIRECTS) {
              reject(new Error(`Too many redirects for ${url}`));
              return;
            }
            const nextUrl = ensureAbsoluteUrl(parsed.toString(), location);
            if (!nextUrl) {
              reject(new Error(`Invalid redirect location for ${url}`));
              return;
            }
            const nextMethod = statusCode === 303 ? "GET" : method;
            try {
              const redirected = await requestBufferViaNode(
                nextUrl,
                mergeObjects(settings, { method: nextMethod }),
                redirects + 1
              );
              resolve(redirected);
            } catch (redirectError) {
              reject(redirectError);
            }
            return;
          }
          resolve({
            ok: statusCode >= 200 && statusCode < 300,
            status: statusCode,
            body: Buffer.concat(chunks),
          });
        });
      }
    );
    request.on("error", (error) => {
      reject(error);
    });
    request.setTimeout(Number(settings.timeoutMs || REQUEST_TIMEOUT_MS), () => {
      request.destroy(new Error(`Request timed out for ${url}`));
    });
    if (settings.body) {
      request.write(settings.body);
    }
    request.end();
  });
}

function bytesToPdfDataUrl(bytes) {
  if (!(bytes instanceof Uint8Array) || !bytes.byteLength) {
    return "";
  }
  return `data:application/pdf;base64,${Buffer.from(bytes).toString("base64")}`;
}

function looksLikePdfBytes(bytes) {
  return (
    bytes instanceof Uint8Array &&
    bytes.byteLength >= 5 &&
    bytes[0] === 0x25 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x44 &&
    bytes[3] === 0x46 &&
    bytes[4] === 0x2d
  );
}

function parseJsonSafe(value) {
  try {
    return JSON.parse(String(value || ""));
  } catch (_error) {
    return null;
  }
}

function extractLikelyUrlFromObject(input) {
  const queue = [input];
  const visited = typeof WeakSet === "function" ? new WeakSet() : null;
  while (queue.length) {
    const current = queue.shift();
    if (!current) {
      continue;
    }
    if (typeof current === "string") {
      const trimmed = current.trim();
      if (/^https?:\/\/\S+/i.test(trimmed)) {
        return trimmed;
      }
      continue;
    }
    if (typeof current !== "object") {
      continue;
    }
    if (visited) {
      if (visited.has(current)) {
        continue;
      }
      visited.add(current);
    }
    if (Array.isArray(current)) {
      for (let i = 0; i < current.length; i += 1) {
        queue.push(current[i]);
      }
      continue;
    }
    const directUrl = pickFirst(
      current.url ||
        current.downloadUrl ||
        current.pdfUrl ||
        current.signedUrl ||
        current.fileUrl ||
        current.location
    );
    if (directUrl && /^https?:\/\/\S+/i.test(directUrl.trim())) {
      return directUrl.trim();
    }
    const nestedValues = [
      current.data,
      current.response,
      current.result,
      current.payload,
      current.body,
    ];
    const keys = Object.keys(current);
    for (let i = 0; i < keys.length; i += 1) {
      nestedValues.push(current[keys[i]]);
    }
    for (let i = 0; i < nestedValues.length; i += 1) {
      if (nestedValues[i] !== undefined) {
        queue.push(nestedValues[i]);
      }
    }
  }
  return "";
}

function extractLikelyPdfUrlFromBytes(bytes) {
  if (!(bytes instanceof Uint8Array) || !bytes.byteLength) {
    return "";
  }
  const text = Buffer.from(bytes).toString("utf8").trim();
  if (!text) {
    return "";
  }
  if (/^https?:\/\/\S+/i.test(text)) {
    return text;
  }
  const parsed = parseJsonSafe(text);
  if (typeof parsed === "string" && /^https?:\/\/\S+/i.test(parsed.trim())) {
    return parsed.trim();
  }
  if (parsed && typeof parsed === "object") {
    const fromObject = extractLikelyUrlFromObject(parsed);
    if (fromObject) {
      return fromObject;
    }
  }
  return "";
}

function isLikelySignedPdfUrl(value) {
  const text = pickFirst(value);
  if (!text) {
    return false;
  }
  let parsed = null;
  try {
    parsed = new URL(text);
  } catch (_error) {
    return false;
  }
  const query = String(parsed.search || "").toLowerCase();
  if (!query) {
    return false;
  }
  return (
    query.includes("x-amz-signature=") ||
    query.includes("x-amz-credential=") ||
    query.includes("x-amz-security-token=") ||
    query.includes("signature=") ||
    query.includes("token=") ||
    query.includes("policy=") ||
    query.includes("expires=")
  );
}

async function removeLogoFromPdfBytes(pdfBytes) {
  if (!(pdfBytes instanceof Uint8Array) || !looksLikePdfBytes(pdfBytes)) {
    return pdfBytes;
  }
  try {
    const pdfDocument = await PDFDocument.load(pdfBytes, {
      ignoreEncryption: true,
      updateMetadata: false,
    });
    const pages = pdfDocument.getPages();
    if (!pages.length) {
      return pdfBytes;
    }
    const firstPage = pages[0];
    const pageSize = firstPage.getSize();
    const pageWidth = Number(pageSize.width || 0);
    const pageHeight = Number(pageSize.height || 0);
    if (!pageWidth || !pageHeight) {
      return pdfBytes;
    }
    const maskWidth = Math.max(70, pageWidth * 0.18);
    const maskHeight = Math.max(70, pageHeight * 0.16);
    // Nudge mask slightly up/left to align with Rocketlane logo bounds.
    const x = Math.max(0, pageWidth - maskWidth - pageWidth * 0.075);
    const y = Math.max(0, pageHeight - maskHeight - pageHeight * 0.065);
    firstPage.drawRectangle({
      x,
      y,
      width: maskWidth,
      height: maskHeight,
      color: rgb(1, 1, 1),
      borderWidth: 0,
      opacity: 1,
    });
    const savedBytes = await pdfDocument.save({
      useObjectStreams: false,
      updateFieldAppearances: false,
    });
    const normalized = savedBytes instanceof Uint8Array ? savedBytes : new Uint8Array(savedBytes);
    return looksLikePdfBytes(normalized) ? normalized : pdfBytes;
  } catch (_error) {
    return pdfBytes;
  }
}

async function fetchPreviewPdfData(baseUrl, headers, previewInvoiceId, invoiceRecord) {
  const encodedId = String(previewInvoiceId || "").trim();
  if (!encodedId) {
    return { pdfDataUrl: "", pdfBase64: "", pdfSource: "", pdfUrl: "" };
  }
  const requestHeaders = mergeObjects(headers, { Accept: "*/*" });
  const pdfPaths = [
    {
      path: `/invoices/${encodedId}/attachments/download`,
      source: "web-attachments-download",
      methods: ["GET"],
    },
    {
      path: `/api/v1/invoices/${encodedId}/attachments/download`,
      source: "api-v1-attachments-download",
      methods: ["GET"],
    },
    {
      path: `/api/1.0/invoices/${encodedId}/attachments/download`,
      source: "api-1-attachments-download",
      methods: ["GET"],
    },
    {
      path: `/api/v1/invoices/${encodedId}/generate`,
      source: "api-v1-generate",
      methods: ["GET", "POST"],
    },
    {
      path: `/api/1.0/invoices/${encodedId}/generate`,
      source: "api-1-generate",
      methods: ["GET", "POST"],
    },
  ];
  for (let i = 0; i < pdfPaths.length; i += 1) {
    const candidate = pdfPaths[i];
    try {
      const pdfBytes = await requestBinaryWithMethods(
        ensureAbsoluteUrl(baseUrl, candidate.path),
        requestHeaders,
        candidate.methods
      );
      if (looksLikePdfBytes(pdfBytes)) {
        const logoMaskedPdfBytes = await removeLogoFromPdfBytes(pdfBytes);
        const pdfBase64 = Buffer.from(logoMaskedPdfBytes).toString("base64");
        return {
          pdfDataUrl: `data:application/pdf;base64,${pdfBase64}`,
          pdfBase64,
          pdfSource: `${candidate.source}-logo-masked`,
          pdfUrl: "",
        };
      }
      const redirectedPdfUrl = extractLikelyPdfUrlFromBytes(pdfBytes);
      if (redirectedPdfUrl) {
        try {
          const redirectedBytes = await requestBinaryWithMethods(
            redirectedPdfUrl,
            { Accept: "*/*" },
            ["GET"]
          );
          if (looksLikePdfBytes(redirectedBytes)) {
            const logoMaskedRedirectedBytes = await removeLogoFromPdfBytes(redirectedBytes);
            const redirectedBase64 = Buffer.from(logoMaskedRedirectedBytes).toString("base64");
            return {
              pdfDataUrl: `data:application/pdf;base64,${redirectedBase64}`,
              pdfBase64: redirectedBase64,
              pdfSource: `${candidate.source}-redirect-url-logo-masked`,
              pdfUrl: "",
            };
          }
        } catch (_redirectError) {
          if (isLikelySignedPdfUrl(redirectedPdfUrl)) {
            return {
              pdfDataUrl: "",
              pdfBase64: "",
              pdfSource: `${candidate.source}-redirect-url-signed`,
              pdfUrl: redirectedPdfUrl,
            };
          }
        }
      }
    } catch (_error) {
      // Ignore and continue with next candidate endpoint.
    }
  }
  const invoiceUrlPdf = await fetchPreviewPdfDataFromUrlCandidates(
    baseUrl,
    headers,
    invoiceRecord
  );
  if (invoiceUrlPdf && (invoiceUrlPdf.pdfDataUrl || invoiceUrlPdf.pdfUrl)) {
    return invoiceUrlPdf;
  }
  return { pdfDataUrl: "", pdfBase64: "", pdfSource: "", pdfUrl: "" };
}

function collectPreviewPdfUrlCandidates(invoiceRecord) {
  if (!invoiceRecord || typeof invoiceRecord !== "object") {
    return [];
  }
  const urls = [];
  const push = (value) => {
    const text = pickFirst(value);
    if (text) {
      urls.push(text);
    }
  };
  push(invoiceRecord.signedUrl);
  push(invoiceRecord.downloadUrl);
  push(invoiceRecord.fileUrl);
  push(invoiceRecord.url);
  push(invoiceRecord.href);
  push(invoiceRecord.previewUrl);
  push(invoiceRecord.attachmentUrl);
  push(invoiceRecord.documentUrl);
  push(invoiceRecord.pdfUrl);
  if (invoiceRecord.file && typeof invoiceRecord.file === "object") {
    push(invoiceRecord.file.signedUrl);
    push(invoiceRecord.file.downloadUrl);
    push(invoiceRecord.file.url);
  }
  const attachmentLists = [
    invoiceRecord.attachments,
    invoiceRecord.attachmentFiles,
    invoiceRecord.files,
    invoiceRecord.documents,
  ];
  attachmentLists.forEach((list) => {
    if (!Array.isArray(list)) {
      return;
    }
    list.forEach((entry) => {
      if (!entry || typeof entry !== "object") {
        return;
      }
      push(entry.signedUrl || entry.downloadUrl || entry.fileUrl || entry.url || entry.href);
      if (entry.file && typeof entry.file === "object") {
        push(entry.file.signedUrl || entry.file.downloadUrl || entry.file.url);
      }
    });
  });
  return dedupeStrings(urls);
}

async function fetchPreviewPdfDataFromUrlCandidates(baseUrl, headers, invoiceRecord) {
  const requestHeaders = mergeObjects(headers, { Accept: "*/*" });
  const candidates = collectPreviewPdfUrlCandidates(invoiceRecord);
  let signedUrlFallback = "";
  for (let i = 0; i < candidates.length; i += 1) {
    const candidateUrl = ensureAbsoluteUrl(baseUrl, candidates[i]);
    if (!candidateUrl) {
      continue;
    }
    if (!signedUrlFallback && isLikelySignedPdfUrl(candidateUrl)) {
      signedUrlFallback = candidateUrl;
    }
    try {
      const pdfBytes = await requestBinaryWithMethods(candidateUrl, requestHeaders, ["GET"]);
      if (looksLikePdfBytes(pdfBytes)) {
        const logoMaskedPdfBytes = await removeLogoFromPdfBytes(pdfBytes);
        const pdfBase64 = Buffer.from(logoMaskedPdfBytes).toString("base64");
        return {
          pdfDataUrl: `data:application/pdf;base64,${pdfBase64}`,
          pdfBase64,
          pdfSource: "invoice-url-logo-masked",
          pdfUrl: "",
        };
      }
      const redirectedPdfUrl = extractLikelyPdfUrlFromBytes(pdfBytes);
      if (redirectedPdfUrl) {
        if (!signedUrlFallback && isLikelySignedPdfUrl(redirectedPdfUrl)) {
          signedUrlFallback = redirectedPdfUrl;
        }
        try {
          const redirectedBytes = await requestBinaryWithMethods(
            redirectedPdfUrl,
            { Accept: "*/*" },
            ["GET"]
          );
          if (looksLikePdfBytes(redirectedBytes)) {
            const logoMaskedRedirectedBytes = await removeLogoFromPdfBytes(redirectedBytes);
            const redirectedBase64 = Buffer.from(logoMaskedRedirectedBytes).toString("base64");
            return {
              pdfDataUrl: `data:application/pdf;base64,${redirectedBase64}`,
              pdfBase64: redirectedBase64,
              pdfSource: "invoice-url-redirect-logo-masked",
              pdfUrl: "",
            };
          }
        } catch (_redirectError) {
          if (isLikelySignedPdfUrl(redirectedPdfUrl)) {
            return {
              pdfDataUrl: "",
              pdfBase64: "",
              pdfSource: "invoice-url-redirect-signed",
              pdfUrl: redirectedPdfUrl,
            };
          }
        }
      }
    } catch (_error) {
      if (isLikelySignedPdfUrl(candidateUrl)) {
        return {
          pdfDataUrl: "",
          pdfBase64: "",
          pdfSource: "invoice-url-signed-fallback",
          pdfUrl: candidateUrl,
        };
      }
      // Try the next URL candidate.
    }
  }
  if (signedUrlFallback) {
    return {
      pdfDataUrl: "",
      pdfBase64: "",
      pdfSource: "invoice-url-signed-fallback",
      pdfUrl: signedUrlFallback,
    };
  }
  return { pdfDataUrl: "", pdfBase64: "", pdfSource: "", pdfUrl: "" };
}

async function attachPreviewPdfToInvoices(baseUrl, headers, invoices, diagnostics) {
  if (!baseUrl || !Array.isArray(invoices) || !invoices.length) {
    return;
  }
  const maxConcurrency = 4;
  const queue = invoices
    .map((invoice) => ({
      invoice,
      invoiceId: pickFirst(invoice && (invoice.invoiceId || invoice.id)),
    }))
    .filter((entry) => entry.invoice && entry.invoiceId);
  if (!queue.length) {
    diagnostics.previewPdfPrefetch = {
      attempted: 0,
      succeeded: 0,
      failed: 0,
    };
    return;
  }

  let index = 0;
  let succeeded = 0;
  let failed = 0;

  const workers = Array.from({ length: Math.min(maxConcurrency, queue.length) }, () =>
    (async () => {
      while (index < queue.length) {
        const currentIndex = index;
        index += 1;
        const entry = queue[currentIndex];
        const invoiceId = encodeURIComponent(entry.invoiceId);
        try {
          const previewPdf = await fetchPreviewPdfData(baseUrl, headers, invoiceId, entry.invoice);
          if (previewPdf && previewPdf.pdfBase64) {
            entry.invoice.previewPdfBase64 = previewPdf.pdfBase64;
            entry.invoice.previewPdfDataUrl = previewPdf.pdfDataUrl;
            entry.invoice.previewPdfSource = previewPdf.pdfSource;
            succeeded += 1;
          } else {
            failed += 1;
          }
        } catch (_error) {
          failed += 1;
        }
      }
    })()
  );

  await Promise.all(workers);
  diagnostics.previewPdfPrefetch = {
    attempted: queue.length,
    succeeded,
    failed,
  };
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
  previewSourceProjectId,
  options
) {
  const settings = options && typeof options === "object" ? options : {};
  const skipCollectionFallback = settings.skipCollectionFallback === true;
  const explicitId = pickFirst(previewInvoiceId);
  const targetNumber = canonicalInvoiceNumber(
    previewInvoiceNumber || previewInvoiceId || ""
  );
  // Do not short-circuit on numeric explicit IDs when invoice number is provided.
  // Some list payloads can expose numeric IDs that are not invoice IDs for preview APIs.
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
  let rows = [];
  try {
    const lookupPayload = await requestJson(lookupUrl, headers);
    rows = Array.isArray(lookupPayload)
      ? lookupPayload
      : extractCollection(lookupPayload, ["data", "invoices", "items", "results"]);
  } catch (_error) {
    rows = [];
  }
  if (!rows.length && !skipCollectionFallback) {
    const fallbackLookup = await requestCollection(
      baseUrl,
      headers,
      ["/api/v1/invoices", "/api/1.0/invoices"],
      ["data", "invoices", "items", "results"]
    );
    rows = fallbackLookup.rows || [];
  }
  if (!rows.length) {
    return explicitId;
  }

  const targetProjectId = pickFirst(previewSourceProjectId);
  const scored = rows
    .map((row) => {
      const invoiceId = pickFirst(
        row &&
          (row.invoiceId ||
            row.id ||
            row._id ||
            (row.invoice && (row.invoice.invoiceId || row.invoice.id || row.invoice._id)) ||
            (row.data && (row.data.invoiceId || row.data.id || row.data._id)))
      );
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

async function requestInvoiceRecordForPreview(baseUrl, headers, previewInvoiceId) {
  const encodedId = encodeURIComponent(String(previewInvoiceId || "").trim());
  if (!encodedId) {
    return {};
  }
  try {
    const payload = await requestJson(
      ensureAbsoluteUrl(baseUrl, `/api/v1/invoices/${encodedId}`),
      headers
    );
    return (
      extractRecordObject(payload || {}, [
        "invoice",
        "data",
        "result",
        "payload",
        "response",
      ]) || {}
    );
  } catch (_error) {
    try {
      const payload = await requestJson(
        ensureAbsoluteUrl(baseUrl, `/api/1.0/invoices/${encodedId}`),
        headers
      );
      return (
        extractRecordObject(payload || {}, [
          "invoice",
          "data",
          "result",
          "payload",
          "response",
        ]) || {}
      );
    } catch (_fallbackError) {
      return {};
    }
  }
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
      if (records.length > 0) {
        break;
      }
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
  const hubFallbackValues = extractFieldValuesByLabelAliases(record, HUB_FALLBACK_LABEL_ALIASES);
  const programFallbackValues = extractFieldValuesByLabelAliases(
    record,
    PROGRAM_FALLBACK_LABEL_ALIASES
  );
  const derivedHubFromAddress = hubFallbackValues
    .map((value) => deriveHubFromAddressText(value))
    .find(Boolean);
  const contractName = compactJoined(
    customFieldValues.contractName.concat([projectInfo.contractName])
  );
  const hub = compactJoined(
    customFieldValues.hub.concat([derivedHubFromAddress], [projectInfo.hub])
  );
  const program = compactJoined(
    customFieldValues.program.concat(programFallbackValues, [projectInfo.program])
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
        record.accountName,
        record.account && (record.account.accountName || record.account.name),
        record.customer && (record.customer.accountName || record.customer.companyName || record.customer.name),
        record.company && (record.company.companyName || record.company.name),
        record.companyName,
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

function hasFullInvoiceAccessToken(text) {
  const value = String(text || "")
    .toLowerCase()
    .replace(/[_-]+/g, " ");
  if (!value) {
    return false;
  }
  return (
    /(^|\b)account\s*admin(istrator)?(\b|$)/.test(value) ||
    /(^|\b)finance(\b|$)/.test(value) ||
    value === "admin" ||
    value === "administrator"
  );
}

function isAdminToken(text) {
  return hasFullInvoiceAccessToken(text);
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
  const adminToken = deduped.find((token) => hasFullInvoiceAccessToken(token));
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

  const isAdmin =
    hasFullInvoiceAccessToken(permission) ||
    hasFullInvoiceAccessToken(roleLabel);

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
      (hasFullInvoiceAccessToken(matched.roleLabel) ? matched.roleLabel : "")
  );
  const roleLabel = pickFirst(matched.roleLabel || viewer.roleLabel);
  const isAdmin =
    hasFullInvoiceAccessToken(permission) ||
    hasFullInvoiceAccessToken(roleLabel) ||
    viewer.isAdmin === true;
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
    request = normalizeIncomingRequest(request);
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
    const workspaceApiBaseCandidates = dedupeStrings(
      workspaceCandidatesForFetch
        .map((candidate) => pickFirst(candidate).replace(/\/+$/, ""))
        .filter(Boolean)
    );
    const previewApiBaseCandidates = dedupeStrings([
      ...apiBaseCandidates,
      ...workspaceApiBaseCandidates,
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
      previewApiBaseCandidates,
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
    const previewInvoiceIdRequested = pickFirst(
      request.previewInvoiceId ||
        request.invoiceId ||
        (request.preview && request.preview.invoiceId)
    );
    const previewInvoiceNumberRequested = pickFirst(
      request.previewInvoiceNumber ||
        request.invoiceNumberForPreview ||
        request.invoiceNumber ||
        (request.preview && request.preview.invoiceNumber)
    );
    const targetPrefetchInvoiceId = pickFirst(
      request.prefetchInvoiceId || previewInvoiceIdRequested || request.invoiceId
    );
    const targetPrefetchInvoiceNumber = canonicalInvoiceNumber(
      pickFirst(
        request.prefetchInvoiceNumber ||
          request.previewInvoiceNumber ||
          request.invoiceNumberForPreview ||
          request.invoiceNumber ||
          (request.preview && request.preview.invoiceNumber)
      )
    );
    const requestMode = String(request.requestMode || request.mode || "").toLowerCase();
    const disablePreviewMode = request.disablePreviewMode === true;
    const isPreviewPdfRequest = requestMode === "preview-pdf";
    const isPreviewRequest =
      !disablePreviewMode &&
      (requestMode === "preview" ||
        isPreviewPdfRequest ||
        Boolean(previewInvoiceIdRequested || previewInvoiceNumberRequested));
    const isTargetedPreviewPrefetchRequest =
      request.searchOnly !== true &&
      request.prefetchPreviewPdfs === true &&
      disablePreviewMode &&
      Boolean(targetPrefetchInvoiceId || targetPrefetchInvoiceNumber);
    const isAnyPreviewPrefetchRequest =
      request.searchOnly !== true &&
      request.prefetchPreviewPdfs === true &&
      disablePreviewMode;
    diagnostics.previewRequest = {
      isPreviewRequest,
      disablePreviewMode,
      previewInvoiceIdRequested,
      previewInvoiceNumberRequested,
      requestMode: String(request.requestMode || request.mode || ""),
      targetedPrefetch: isTargetedPreviewPrefetchRequest,
      anyPrefetchPreview: isAnyPreviewPrefetchRequest,
      targetPrefetchInvoiceId,
      targetPrefetchInvoiceNumber,
    };
    const viewer = deriveViewerAccess(request, context);
    let resolvedViewer = viewer;

    if (isPreviewRequest) {
      if (isPreviewPdfRequest) {
        const previewPdfAttempts = [];
        for (let i = 0; i < previewApiBaseCandidates.length; i += 1) {
          const baseUrl = previewApiBaseCandidates[i];
          try {
            const resolvedInvoiceIdFast = await resolveInvoiceIdForPreview(
              baseUrl,
              headers,
              previewInvoiceIdRequested,
              previewInvoiceNumberRequested,
              request.previewSourceProjectId,
              { skipCollectionFallback: true }
            );
            let resolvedInvoiceId = pickFirst(resolvedInvoiceIdFast);
            if (!resolvedInvoiceId && previewInvoiceNumberRequested) {
              resolvedInvoiceId = pickFirst(
                await resolveInvoiceIdForPreview(
                  baseUrl,
                  headers,
                  previewInvoiceIdRequested,
                  previewInvoiceNumberRequested,
                  request.previewSourceProjectId
                )
              );
            }
            const previewInvoiceIds = dedupeStrings([
              previewInvoiceIdRequested,
              pickFirst(resolvedInvoiceId),
            ]).filter(Boolean);
            if (!previewInvoiceIds.length) {
              throw new Error("Invoice ID could not be resolved for preview PDF.");
            }
            for (let candidateIdx = 0; candidateIdx < previewInvoiceIds.length; candidateIdx += 1) {
              const previewInvoiceId = pickFirst(previewInvoiceIds[candidateIdx]);
              if (!previewInvoiceId) {
                continue;
              }
              const invoiceRecord = await requestInvoiceRecordForPreview(
                baseUrl,
                headers,
                previewInvoiceId
              );
              const previewPdf = await fetchPreviewPdfData(
                baseUrl,
                headers,
                encodeURIComponent(previewInvoiceId),
                invoiceRecord
              );
              if (previewPdf && (previewPdf.pdfDataUrl || previewPdf.pdfUrl)) {
                previewPdfAttempts.push({
                  baseUrl,
                  invoiceId: previewInvoiceId,
                  outcome: "success",
                  source: previewPdf.pdfSource || "",
                });
                return {
                  ok: true,
                  previewPdf: {
                    invoiceId: previewInvoiceId,
                    pdfDataUrl: previewPdf.pdfDataUrl,
                    pdfBase64: previewPdf.pdfBase64 || "",
                    pdfSource: previewPdf.pdfSource || "",
                    pdfUrl: previewPdf.pdfUrl || "",
                  },
                  viewer,
                  diagnostics: mergeObjects(diagnostics, {
                    apiBaseUsed: baseUrl,
                    previewInvoiceResolvedId: previewInvoiceId,
                    previewInvoiceResolvedFrom: String(resolvedInvoiceId || ""),
                    previewPdfRequest: true,
                    previewPdfAttempts,
                  }),
                };
              }
              previewPdfAttempts.push({
                baseUrl,
                invoiceId: previewInvoiceId,
                outcome: "no-pdf",
              });
            }
          } catch (error) {
            previewPdfAttempts.push({
              baseUrl,
              invoiceId: "",
              outcome: "error",
              error: String(error && error.message ? error.message : error),
            });
            diagnostics.invoiceErrors.push(
              String(error && error.message ? error.message : error)
            );
          }
        }
        return {
          ok: false,
          error: "Unable to load invoice preview PDF.",
          previewPdf: null,
          viewer,
          diagnostics: mergeObjects(diagnostics, { previewPdfAttempts }),
        };
      }
      let previewFallbackResponse = null;
      for (let i = 0; i < previewApiBaseCandidates.length; i += 1) {
        const baseUrl = previewApiBaseCandidates[i];
        try {
          const resolvedInvoiceId = await resolveInvoiceIdForPreview(
            baseUrl,
            headers,
            previewInvoiceIdRequested,
            previewInvoiceNumberRequested,
            request.previewSourceProjectId
          );
          const previewInvoiceIds = dedupeStrings([
            pickFirst(resolvedInvoiceId),
            previewInvoiceIdRequested,
          ])
            .map((value) => encodeURIComponent(String(value || "").trim()))
            .filter(Boolean);
          if (!previewInvoiceIds.length) {
            throw new Error("Invoice ID could not be resolved for preview.");
          }
          let lastPreviewError = null;
          for (let candidateIdx = 0; candidateIdx < previewInvoiceIds.length; candidateIdx += 1) {
            const previewInvoiceId = previewInvoiceIds[candidateIdx];
            try {
              let invoicePayload = null;
              try {
                invoicePayload = await requestJson(
                  ensureAbsoluteUrl(baseUrl, `/api/v1/invoices/${previewInvoiceId}`),
                  headers
                );
              } catch (_error) {
                try {
                  invoicePayload = await requestJson(
                    ensureAbsoluteUrl(baseUrl, `/api/1.0/invoices/${previewInvoiceId}`),
                    headers
                  );
                } catch (_fallbackError) {
                  invoicePayload = null;
                }
              }
              let lineItems = [];
              try {
                const linePayload = await requestJson(
                  ensureAbsoluteUrl(baseUrl, `/api/1.0/invoices/${previewInvoiceId}/lines`),
                  headers
                );
                lineItems = extractCollection(linePayload, ["data", "lines", "items", "results"]);
              } catch (_error) {
                lineItems = [];
              }
              let payments = [];
              try {
                const paymentPayload = await requestJson(
                  ensureAbsoluteUrl(baseUrl, `/api/1.0/invoices/${previewInvoiceId}/payments`),
                  headers
                );
                payments = extractCollection(paymentPayload, [
                  "data",
                  "payments",
                  "items",
                  "results",
                ]);
              } catch (_error) {
                payments = [];
              }
              const invoiceRecord = extractRecordObject(invoicePayload || {}, [
                "invoice",
                "data",
                "result",
                "payload",
                "response",
              ]) || {};
              const previewPdf = await fetchPreviewPdfData(
                baseUrl,
                headers,
                previewInvoiceId,
                invoiceRecord
              );
              const preview = normalizeInvoicePreview(invoiceRecord, lineItems, payments);
              if (previewPdf.pdfDataUrl) {
                preview.pdfDataUrl = previewPdf.pdfDataUrl;
                preview.pdfBase64 = previewPdf.pdfBase64 || "";
                preview.pdfSource = previewPdf.pdfSource;
              }
              if (previewPdf.pdfUrl) {
                preview.pdfUrl = previewPdf.pdfUrl;
                preview.pdfSource = previewPdf.pdfSource || preview.pdfSource || "";
              }
              if (preview.pdfDataUrl || preview.pdfUrl) {
                return {
                  ok: true,
                  preview,
                  viewer,
                  diagnostics: mergeObjects(diagnostics, {
                    apiBaseUsed: baseUrl,
                    previewInvoiceResolvedId: decodeURIComponent(previewInvoiceId),
                    previewInvoiceResolvedFrom: String(resolvedInvoiceId || ""),
                  }),
                };
              }
              if (
                !previewFallbackResponse &&
                (lineItems.length || payments.length || Object.keys(invoiceRecord).length)
              ) {
                previewFallbackResponse = {
                  ok: true,
                  preview,
                  viewer,
                  diagnostics: mergeObjects(diagnostics, {
                    apiBaseUsed: baseUrl,
                    previewInvoiceResolvedId: decodeURIComponent(previewInvoiceId),
                    previewInvoiceResolvedFrom: String(resolvedInvoiceId || ""),
                    previewPdfUnavailable: true,
                  }),
                };
              }
              lastPreviewError = new Error(
                `No usable preview payload returned for invoice ${decodeURIComponent(previewInvoiceId)}`
              );
            } catch (previewError) {
              lastPreviewError = previewError;
              diagnostics.invoiceErrors.push(
                String(previewError && previewError.message ? previewError.message : previewError)
              );
            }
          }
          if (lastPreviewError) {
            throw lastPreviewError;
          }
          throw new Error("Invoice ID could not be resolved for preview.");
        } catch (error) {
          diagnostics.invoiceErrors.push(
            String(error && error.message ? error.message : error)
          );
        }
      }
      if (previewFallbackResponse) {
        return previewFallbackResponse;
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
      const lineEnrichmentQueue = [];
      for (let idx = 0; idx < globalInvoices.length; idx += 1) {
        const row = globalInvoices[idx];
        const project = resolveProjectForInvoice(row, projectLookup);
        let normalized = normalizeInvoiceRecord(
          row,
          project,
          request.accountName || iParams.accountName || ""
        );
        if (normalized) {
          if (isTargetedPreviewPrefetchRequest) {
            const normalizedInvoiceId = pickFirst(normalized.invoiceId || normalized.id);
            const normalizedInvoiceNumber = canonicalInvoiceNumber(
              pickFirst(normalized.invoiceNumber)
            );
            const idMatches =
              Boolean(targetPrefetchInvoiceId) &&
              Boolean(normalizedInvoiceId) &&
              targetPrefetchInvoiceId === normalizedInvoiceId;
            const numberMatches =
              Boolean(targetPrefetchInvoiceNumber) &&
              Boolean(normalizedInvoiceNumber) &&
              targetPrefetchInvoiceNumber === normalizedInvoiceNumber;
            if (!idMatches && !numberMatches) {
              continue;
            }
          }
          const invoiceId = encodeURIComponent(
            pickFirst(normalized.invoiceId || normalized.id || row.invoiceId || row.id || row._id)
          );
          if (invoiceId) {
            const targetInvoice = normalized;
            const shouldEnrichFromLines =
                !isAnyPreviewPrefetchRequest &&
                (Number(targetInvoice.quantityHours || 0) <= 0 ||
                  !pickFirst(targetInvoice.hub) ||
                  !pickFirst(targetInvoice.program));
            if (shouldEnrichFromLines) {
              lineEnrichmentQueue.push({
                invoiceId,
                targetInvoice,
              });
            }
          }
          collectedInvoices.push(normalized);
        }
      }
      if (lineEnrichmentQueue.length) {
        const maxLineEnrichmentConcurrency = 6;
        let lineEnrichmentIndex = 0;
        const workers = Array.from(
          { length: Math.min(maxLineEnrichmentConcurrency, lineEnrichmentQueue.length) },
          () =>
            (async () => {
              while (lineEnrichmentIndex < lineEnrichmentQueue.length) {
                const currentIndex = lineEnrichmentIndex;
                lineEnrichmentIndex += 1;
                const entry = lineEnrichmentQueue[currentIndex];
                try {
                  const lineFetch = await requestCollection(
                    baseUrl,
                    headers,
                    [`/api/1.0/invoices/${entry.invoiceId}/lines`],
                    ["data", "lines", "items", "results"]
                  );
                  diagnostics.invoiceErrors.push(...lineFetch.errors);
                  const fetchedQuantity = sumLineItemQuantity(lineFetch.rows);
                  if (fetchedQuantity > 0) {
                    entry.targetInvoice.quantityHours = fetchedQuantity;
                  }
                  const extractedLineValues = extractHubProgramFromLineItems(lineFetch.rows);
                  if (Array.isArray(extractedLineValues.hub) && extractedLineValues.hub.length) {
                    entry.targetInvoice.hub = compactJoined(extractedLineValues.hub);
                  }
                  if (
                    Array.isArray(extractedLineValues.program) &&
                    extractedLineValues.program.length
                  ) {
                    entry.targetInvoice.program = compactJoined(extractedLineValues.program);
                  }
                } catch (lineError) {
                  diagnostics.invoiceErrors.push(
                    String(lineError && lineError.message ? lineError.message : lineError)
                  );
                }
              }
            })()
        );
        await Promise.all(workers);
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

    const shouldPrefetchPreviewPdfs =
      request.searchOnly !== true && request.prefetchPreviewPdfs === true;
    if (shouldPrefetchPreviewPdfs) {
      const targetPreviewInvoiceId = pickFirst(
        request.previewInvoiceId ||
          request.invoiceId ||
          request.prefetchInvoiceId ||
          (request.preview && request.preview.invoiceId)
      );
      const targetPreviewInvoiceNumber = canonicalInvoiceNumber(
        pickFirst(
          request.previewInvoiceNumber ||
            request.invoiceNumberForPreview ||
            request.invoiceNumber ||
            request.prefetchInvoiceNumber ||
            (request.preview && request.preview.invoiceNumber)
        )
      );
      let invoicesToPrefetch = dedupedInvoices;
      if (targetPreviewInvoiceId || targetPreviewInvoiceNumber) {
        const targeted = dedupedInvoices.filter((invoice) => {
          const invoiceId = pickFirst(invoice && (invoice.invoiceId || invoice.id));
          const invoiceNumber = canonicalInvoiceNumber(
            pickFirst(invoice && invoice.invoiceNumber)
          );
          if (targetPreviewInvoiceId && invoiceId && invoiceId === targetPreviewInvoiceId) {
            return true;
          }
          if (
            targetPreviewInvoiceNumber &&
            invoiceNumber &&
            invoiceNumber === targetPreviewInvoiceNumber
          ) {
            return true;
          }
          return false;
        });
        if (targeted.length) {
          invoicesToPrefetch = targeted;
        }
      }
      await attachPreviewPdfToInvoices(
        previewApiBaseCandidates[0] || diagnostics.apiBaseUsed || apiBaseCandidates[0] || "",
        headers,
        invoicesToPrefetch,
        diagnostics
      );
    }

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
