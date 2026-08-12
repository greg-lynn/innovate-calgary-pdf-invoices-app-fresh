"use strict";

const { syncInvoicesFromSource } = require("./sync-invoices-from-source");

function pickFirst(value) {
  if (typeof value === "string") {
    return value.trim();
  }
  if (typeof value === "number") {
    return String(value);
  }
  return "";
}

function canonicalInvoiceNumber(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

function normalizeLookupKey(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
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
  const tokens = new Set(searchKeys.map((key) => normalizeLookupKey(key)).filter(Boolean));
  if (!tokens.size) {
    return "";
  }
  const nodes = collectNestedCandidates(root, 6);
  for (let i = 0; i < nodes.length; i += 1) {
    const node = nodes[i];
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      continue;
    }
    const nodeKeys = Object.keys(node);
    for (let j = 0; j < nodeKeys.length; j += 1) {
      const nodeKey = nodeKeys[j];
      if (!tokens.has(normalizeLookupKey(nodeKey))) {
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

function pickFirstObjectFromAny(root, keys) {
  const searchKeys = Array.isArray(keys) ? keys : [];
  const tokens = new Set(searchKeys.map((key) => normalizeLookupKey(key)).filter(Boolean));
  if (!tokens.size) {
    return null;
  }
  const nodes = collectNestedCandidates(root, 6);
  for (let i = 0; i < nodes.length; i += 1) {
    const node = nodes[i];
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      continue;
    }
    const nodeKeys = Object.keys(node);
    for (let j = 0; j < nodeKeys.length; j += 1) {
      const nodeKey = nodeKeys[j];
      if (!tokens.has(normalizeLookupKey(nodeKey))) {
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

function pickFirstArrayFromAny(root, keys) {
  const searchKeys = Array.isArray(keys) ? keys : [];
  const tokens = new Set(searchKeys.map((key) => normalizeLookupKey(key)).filter(Boolean));
  if (!tokens.size) {
    return [];
  }
  const nodes = collectNestedCandidates(root, 6);
  for (let i = 0; i < nodes.length; i += 1) {
    const node = nodes[i];
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      continue;
    }
    const nodeKeys = Object.keys(node);
    for (let j = 0; j < nodeKeys.length; j += 1) {
      const nodeKey = nodeKeys[j];
      if (!tokens.has(normalizeLookupKey(nodeKey))) {
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

function normalizePreviewRequest(request) {
  const base = request && typeof request === "object" ? request : {};
  const preview =
    pickFirstObjectFromAny(base, ["preview"]) ||
    (base.preview && typeof base.preview === "object" ? base.preview : {});
  const invoiceId = pickFirst(
    base.previewInvoiceId ||
      base.invoiceId ||
      preview.invoiceId ||
      pickFirstValueFromAny(base, ["previewInvoiceId", "invoiceId"])
  );
  const invoiceNumber = pickFirst(
    base.previewInvoiceNumber ||
      base.invoiceNumberForPreview ||
      base.invoiceNumber ||
      preview.invoiceNumber ||
      pickFirstValueFromAny(base, [
        "previewInvoiceNumber",
        "invoiceNumberForPreview",
        "invoiceNumber",
      ])
  );
  const viewerContext =
    pickFirstObjectFromAny(base, ["viewerContext"]) ||
    (base.viewerContext && typeof base.viewerContext === "object" ? base.viewerContext : {});
  return Object.assign({}, base, {
    requestMode: "preview-pdf",
    mode: "preview-pdf",
    searchOnly: false,
    prefetchPreviewPdfs: false,
    disablePreviewMode: false,
    prefetchInvoiceId: invoiceId,
    prefetchInvoiceNumber: invoiceNumber,
    workspaceBaseUrl: pickFirst(
      base.workspaceBaseUrl ||
        base.workspaceUrl ||
        pickFirstValueFromAny(base, ["workspaceBaseUrl", "workspaceUrl"])
    ),
    workspaceCandidates: Array.isArray(base.workspaceCandidates)
      ? base.workspaceCandidates
      : pickFirstArrayFromAny(base, ["workspaceCandidates"]),
    previewInvoiceId: invoiceId,
    invoiceId,
    previewInvoiceNumber: invoiceNumber,
    invoiceNumberForPreview: invoiceNumber,
    viewerContext,
    preview: Object.assign({}, preview, {
      invoiceId: pickFirst(preview.invoiceId || invoiceId),
      invoiceNumber: pickFirst(preview.invoiceNumber || invoiceNumber),
    }),
  });
}

function hasPreviewPdf(result) {
  if (!result || typeof result !== "object") {
    return false;
  }
  const previewPdf = result.previewPdf && typeof result.previewPdf === "object" ? result.previewPdf : null;
  if (!previewPdf) {
    return false;
  }
  return Boolean(
    pickFirst(previewPdf.pdfDataUrl || "") ||
      pickFirst(previewPdf.pdfBase64 || "") ||
      pickFirst(previewPdf.pdfUrl || "")
  );
}

function buildTargetedPrefetchRequest(request) {
  const targetInvoiceId = pickFirst(request && (request.previewInvoiceId || request.invoiceId));
  const targetInvoiceNumber = pickFirst(
    request &&
      (request.previewInvoiceNumber || request.invoiceNumberForPreview || request.invoiceNumber)
  );
  return Object.assign({}, request, {
    requestMode: "",
    mode: "",
    searchOnly: false,
    prefetchPreviewPdfs: true,
    disablePreviewMode: true,
    prefetchInvoiceId: targetInvoiceId,
    prefetchInvoiceNumber: targetInvoiceNumber,
    previewInvoiceId: targetInvoiceId,
    invoiceId: targetInvoiceId,
    previewInvoiceNumber: targetInvoiceNumber,
    invoiceNumberForPreview: targetInvoiceNumber,
    invoiceNumber: targetInvoiceNumber,
    preview: Object.assign({}, request && request.preview, {
      invoiceId: targetInvoiceId,
      invoiceNumber: targetInvoiceNumber,
    }),
  });
}

function normalizePreviewPdfFromResult(result, request) {
  if (!result || typeof result !== "object" || result.ok === false) {
    return result;
  }
  if (result.previewPdf && typeof result.previewPdf === "object") {
    return result;
  }
  const targetInvoiceId = pickFirst(request && (request.previewInvoiceId || request.invoiceId));
  const targetInvoiceNumber = canonicalInvoiceNumber(
    pickFirst(
    request &&
      (request.previewInvoiceNumber || request.invoiceNumberForPreview || request.invoiceNumber)
    )
  );
  const invoices = Array.isArray(result.invoices) ? result.invoices : [];
  const matched = invoices.find((invoice) => {
    if (!invoice || typeof invoice !== "object") {
      return false;
    }
    const invoiceId = pickFirst(invoice.invoiceId || invoice.id);
    const invoiceNumber = canonicalInvoiceNumber(pickFirst(invoice.invoiceNumber));
    if (targetInvoiceId && invoiceId && targetInvoiceId === invoiceId) {
      return true;
    }
    if (targetInvoiceNumber && invoiceNumber && targetInvoiceNumber === invoiceNumber) {
      return true;
    }
    return false;
  });
  if (!matched) {
    return result;
  }
  const previewPdf = {
    invoiceId: pickFirst(matched.invoiceId || matched.id || targetInvoiceId),
    pdfDataUrl: pickFirst(matched.previewPdfDataUrl || matched.pdfDataUrl),
    pdfBase64: pickFirst(matched.previewPdfBase64 || matched.pdfBase64),
    pdfSource: pickFirst(matched.previewPdfSource || matched.pdfSource || "invoices-list"),
    pdfUrl: pickFirst(matched.previewPdfUrl || matched.pdfUrl || ""),
  };
  if (!previewPdf.pdfDataUrl && !previewPdf.pdfBase64 && !previewPdf.pdfUrl) {
    return result;
  }
  return Object.assign({}, result, { previewPdf });
}

module.exports = {
  syncInvoicePreviewPayload: async (request = {}, context = {}) => {
    const normalizedRequest = normalizePreviewRequest(request);
    const startedAt = Date.now();
    let strictResult = null;
    let strictError = "";
    try {
      strictResult = await syncInvoicesFromSource(normalizedRequest, context);
    } catch (error) {
      strictError = String(error && error.message ? error.message : error);
    }
    const strictNormalized = normalizePreviewPdfFromResult(strictResult, normalizedRequest);
    if (hasPreviewPdf(strictNormalized)) {
      return Object.assign({}, strictNormalized, {
        diagnostics: Object.assign({}, strictNormalized.diagnostics || {}, {
          previewPipeline: "strict-preview-pdf",
          previewPipelineElapsedMs: Date.now() - startedAt,
          previewPipelineStrictError: strictError,
        }),
      });
    }

    const fallbackRequest = buildTargetedPrefetchRequest(normalizedRequest);
    let fallbackResult = null;
    let fallbackError = "";
    try {
      fallbackResult = await syncInvoicesFromSource(fallbackRequest, context);
    } catch (error) {
      fallbackError = String(error && error.message ? error.message : error);
    }
    const fallbackNormalized = normalizePreviewPdfFromResult(fallbackResult, fallbackRequest);
    if (hasPreviewPdf(fallbackNormalized)) {
      return Object.assign({}, fallbackNormalized, {
        diagnostics: Object.assign({}, fallbackNormalized.diagnostics || {}, {
          previewPipeline: "targeted-prefetch-fallback",
          previewPipelineElapsedMs: Date.now() - startedAt,
          previewPipelineStrictError: strictError,
          previewPipelineFallbackError: fallbackError,
        }),
      });
    }

    const baseResult =
      fallbackNormalized && typeof fallbackNormalized === "object"
        ? fallbackNormalized
        : strictNormalized && typeof strictNormalized === "object"
          ? strictNormalized
          : { ok: false, previewPdf: null };
    return Object.assign({}, baseResult, {
      diagnostics: Object.assign({}, (baseResult && baseResult.diagnostics) || {}, {
        previewPipeline: "no-pdf",
        previewPipelineElapsedMs: Date.now() - startedAt,
        previewPipelineStrictError: strictError,
        previewPipelineFallbackError: fallbackError,
      }),
    });
  },
};

