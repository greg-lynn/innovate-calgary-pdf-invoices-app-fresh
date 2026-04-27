"use strict";

const DEFAULT_SOURCE_PROJECTS = ["Expert Advisor Program Invoices"];
// Production override: embed API key here so app works without installer prompt.
// Replace before shipping to users if needed.
const EMBEDDED_ROCKETLANE_API_KEY = "rl-7e0f30b5-1aab-4faf-837c-6a3ec5cbfde7";
const ROCKETLANE_API_BASE_URL = "https://api.rocketlane.com";

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

function normalizeFieldLabel(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
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
  };
  if (!Array.isArray(fields)) {
    return output;
  }
  fields.forEach((field) => {
    if (!field || typeof field !== "object") {
      return;
    }
    const label = normalizeFieldLabel(
      pickFirst(
        field.fieldLabel || field.fieldName || field.label || field.name || field.key
      )
    );
    if (!label) {
      return;
    }
    let targetKey = "";
    if (label === "contract name" || label === "contractname" || label === "contract") {
      targetKey = "contractName";
    } else if (label === "hub") {
      targetKey = "hub";
    } else if (label === "program") {
      targetKey = "program";
    }
    if (!targetKey) {
      return;
    }
    const candidates = [
      field.fieldValueLabel,
      field.fieldValue,
      field.displayValue,
      field.value,
      field.values,
      field.metaFieldValue && field.metaFieldValue.label,
      field.metaFieldValue && field.metaFieldValue.value,
      field.option && field.option.label,
      field.option && field.option.value,
    ];
    candidates.forEach((candidate) => {
      const text = toFieldText(candidate);
      if (!text) {
        return;
      }
      text
        .split(",")
        .map((part) => pickFirst(part))
        .filter(Boolean)
        .forEach((part) => output[targetKey].push(part));
    });
  });
  output.contractName = dedupeStrings(output.contractName);
  output.hub = dedupeStrings(output.hub);
  output.program = dedupeStrings(output.program);
  return output;
}

function extractFieldDisplayEntries(fields) {
  if (!Array.isArray(fields)) {
    return [];
  }
  const entries = [];
  const seen = new Set();
  fields.forEach((field) => {
    if (!field || typeof field !== "object") {
      return;
    }
    const label = pickFirst(
      field.fieldLabel || field.fieldName || field.label || field.name || field.key
    );
    if (!label) {
      return;
    }
    const value = toFieldText(
      pickFirst(
        field.fieldValueLabel ||
          field.fieldValue ||
          field.displayValue ||
          field.value ||
          (field.metaFieldValue &&
            (field.metaFieldValue.label || field.metaFieldValue.value))
      ) || field.values
    );
    if (!value) {
      return;
    }
    const dedupeKey = `${normalizeFieldLabel(label)}|${value}`;
    if (seen.has(dedupeKey)) {
      return;
    }
    seen.add(dedupeKey);
    entries.push({ label, value });
  });
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
    const quantity = normalizeAmount(
      line && (line.quantity || line.hours || line.billableHours || line.billableQuantity)
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
  const customFieldValues = extractNamedCustomFieldValues(record.fields);
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
  const customFieldValues = extractNamedCustomFieldValues(record.fields);
  const lineItems = collectInvoiceLineItems(record);
  const quantityHours = sumLineItemQuantity(lineItems);
  const contractName = compactJoined(
    customFieldValues.contractName.concat([projectInfo.contractName])
  );
  const hub = compactJoined(customFieldValues.hub.concat([projectInfo.hub]));
  const program = compactJoined(
    customFieldValues.program.concat([projectInfo.program])
  );
  const createdByName = fullName(record.createdBy) || pickFirst(record.createdByName);

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
      pickFirst(
        record.accountName ||
          record.companyName ||
          (record.company && (record.company.companyName || record.company.name)) ||
          (record.account && record.account.name) ||
          (record.customer && (record.customer.companyName || record.customer.name))
      ) ||
      projectInfo.accountName ||
      fallbackAccountName ||
      "Rocketlane Account",
    invoiceDate,
    issueDate,
    dueDate,
    invoiceStatus,
    amount,
    currencyCode,
    currencySymbol,
    pdfUrl,
    createdByName,
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
  const invoiceNumber = pickFirst(invoiceRecord && invoiceRecord.invoiceNumber);
  const projectName = pickFirst(
    invoiceRecord &&
      Array.isArray(invoiceRecord.projects) &&
      invoiceRecord.projects[0] &&
      (invoiceRecord.projects[0].projectName || invoiceRecord.projects[0].name)
  );
  const createdBy = invoiceRecord && invoiceRecord.createdBy ? invoiceRecord.createdBy : null;
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
    billToName: fullName(createdBy),
    billToEmail: normalizeEmail(
      pickFirst(createdBy && (createdBy.email || createdBy.emailId || createdBy.userEmail))
    ),
    customFields: extractFieldDisplayEntries(invoiceRecord && invoiceRecord.fields),
    lineItems: Array.isArray(lineRecords)
      ? lineRecords.map((line) => ({
          id: pickFirst(line && (line.invoiceLineItemId || line.id || line._id)),
          description: pickFirst(line && line.description),
          quantity: normalizeAmount(line && line.quantity),
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
  const email = normalizeEmail(
    pickFirst(
      record.email ||
        record.userEmail ||
        record.workEmail ||
        (record.user && record.user.email) ||
        (record.profile && record.profile.email)
    )
  );
  const id = pickFirst(record.id || record.userId || record._id);
  if (!email && !id) {
    return null;
  }
  return {
    id,
    email,
    permission: pickFirst(
      (record.permission && record.permission.permissionName) ||
        (record.permission && record.permission.name) ||
        record.permission ||
        record.permissionSet ||
        (record.permissionSet && record.permissionSet.name) ||
        record.accountPermission ||
        (record.permissionSetObj && record.permissionSetObj.name)
    ),
    roleLabel: pickFirst(
      (record.role && (record.role.roleName || record.role.name)) ||
        record.role ||
        record.userRole ||
        record.designation ||
        record.title
    ),
  };
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
    /(^|\b)(account|workspace|company)\s*admin(istrator)?(\b|$)/.test(value) ||
    /(^|\b)(account|workspace|company)\s*owner(\b|$)/.test(value) ||
    /(^|\b)admin(\b|$)/.test(value)
  );
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
    ctxUser.permission ||
      ctxUser.permissionSet ||
      (ctxUser.permissionSet && ctxUser.permissionSet.name) ||
      (ctxUser.permissionSetObj && ctxUser.permissionSetObj.name) ||
      ctxUser.accountPermission ||
      (ctxUser.accountPermissionSet && ctxUser.accountPermissionSet.name)
  );
  const roleLabel = pickFirst(
    ctxUser.role ||
      ctxUser.userRole ||
      ctxUser.type ||
      ctxUser.userType ||
      ctxUser.designation ||
      viewerContext.userRole
  );

  const tokens = [];
  collectRoleTokens(ctxUser, tokens, 0);
  collectRoleTokens(ctxAccount, tokens, 0);
  collectRoleTokens(viewerContext, tokens, 0);
  const uniqueTokens = dedupeStrings(tokens);
  const tokenText = uniqueTokens.join(" ");
  const isAdmin = Boolean(
    ctxUser.isAdmin === true ||
      ctxUser.admin === true ||
      ctxUser.isAccountAdmin === true ||
      ctxUser.accountAdmin === true ||
      isAdminToken(permission) ||
      isAdminToken(roleLabel) ||
      isAdminToken(tokenText)
  );

  return {
    id,
    email,
    displayName,
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
      iParams.workspaceBaseUrl,
      iParams.workspaceUrl,
      secureParams.workspaceBaseUrl,
      secureParams.workspaceUrl,
      "https://blink.rocketlane.com",
      "https://innovate-calgary.rocketlane.com",
    ]);
    const apiToken =
      EMBEDDED_ROCKETLANE_API_KEY ||
      request.apiToken ||
      secureParams.rocketlaneApiToken ||
      secureParams.apiToken ||
      secureParams.apiKey ||
      iParams.rocketlaneApiToken ||
      iParams.apiToken ||
      context.apiKey ||
      "";

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
      workspaceCandidates,
      projectErrors: [],
      invoiceErrors: [],
      memberErrors: [],
      invoiceFetchMode: "account-wide-invoices-type-all",
      workspaceUsed: "",
      apiBaseUsed: "",
      hasApiToken: Boolean(apiToken),
      tokenSource: EMBEDDED_ROCKETLANE_API_KEY
        ? "embedded"
        : request.apiToken
        ? "request.apiToken"
        : secureParams.rocketlaneApiToken || secureParams.apiToken || secureParams.apiKey
        ? "installation.secureParams"
        : iParams.rocketlaneApiToken || iParams.apiToken
        ? "installation.iparams"
        : context.apiKey
        ? "context.apiKey"
        : "none",
      contextUserKeys: context && context.user ? Object.keys(context.user) : [],
      searchQuery: normalizedSearchQuery,
    };
    const viewer = deriveViewerAccess(request, context);

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
          const preview = normalizeInvoicePreview(invoicePayload || {}, lineItems, payments);
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
          "/api/v1/invoices?type=all",
          "/api/v1/invoices",
          "/api/1.0/invoices?type=all",
          "/api/1.0/invoices",
        ],
        ["invoices", "data", "content", "results", "items"]
      );
      diagnostics.invoiceErrors.push(...allInvoicesResult.errors);
      const globalInvoices = allInvoicesResult.rows;

      const collectedInvoices = [];
      globalInvoices.forEach((row) => {
        const project = resolveProjectForInvoice(row, projectLookup);
        const normalized = normalizeInvoiceRecord(
          row,
          project,
          request.accountName || iParams.accountName || ""
        );
        if (normalized) {
          collectedInvoices.push(normalized);
        }
      });

      const membersResult = await requestCollection(
        baseUrl,
        headers,
        ["/api/1.0/users?includeFields=permission,role,company", "/api/1.0/users?includeFields=permission", "/api/1.0/users"],
        ["users", "members", "teamMembers", "data", "results", "items"]
      );
      diagnostics.memberErrors.push(...membersResult.errors);
      const normalizedMembers = membersResult.rows.map(normalizeMember).filter(Boolean);

      if (allProjects.length || globalInvoices.length || normalizedMembers.length) {
        diagnostics.workspaceUsed = workspaceCandidates[0] || "";
        diagnostics.apiBaseUsed = baseUrl;
        sourceProjects = dedupeStrings(
          collectedInvoices.map((invoice) => pickFirst(invoice && invoice.sourceProjectName))
        );
        invoices = collectedInvoices;
        members = normalizedMembers;
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
      viewer,
      diagnostics,
    };
  },
};
