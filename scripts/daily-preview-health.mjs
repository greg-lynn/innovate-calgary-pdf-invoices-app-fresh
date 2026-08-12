#!/usr/bin/env node

import fs from "fs";
import path from "path";
import { createRequire } from "module";

const require = createRequire(import.meta.url);
const { syncInvoicePreviewPayload } = require("../server-actions/sync-invoice-preview.js");

function pickFirst(value) {
  if (typeof value === "string") {
    return value.trim();
  }
  if (typeof value === "number") {
    return String(value);
  }
  return "";
}

function parseIntegerEnv(name, fallback) {
  const raw = Number(process.env[name]);
  if (!Number.isFinite(raw) || raw <= 0) {
    return fallback;
  }
  return Math.floor(raw);
}

function hasPreviewPdf(result) {
  const previewPdf = result && typeof result === "object" ? result.previewPdf : null;
  if (!previewPdf || typeof previewPdf !== "object") {
    return false;
  }
  return Boolean(
    pickFirst(previewPdf.pdfDataUrl) ||
      pickFirst(previewPdf.pdfBase64) ||
      pickFirst(previewPdf.pdfUrl)
  );
}

async function runPreviewProbe(invoiceNumber, workspaceBaseUrl) {
  const startedAt = Date.now();
  const result = await syncInvoicePreviewPayload({
    previewInvoiceNumber: invoiceNumber,
    invoiceNumberForPreview: invoiceNumber,
    workspaceBaseUrl,
    workspaceCandidates: [workspaceBaseUrl],
  });
  const elapsedMs = Date.now() - startedAt;
  const ok = Boolean(result && result.ok && hasPreviewPdf(result));
  return {
    invoiceNumber,
    ok,
    elapsedMs,
    previewPipeline:
      result && result.diagnostics ? pickFirst(result.diagnostics.previewPipeline) : "",
    previewSource:
      result && result.previewPdf ? pickFirst(result.previewPdf.pdfSource || "") : "",
    diagnostics: result && result.diagnostics ? result.diagnostics : null,
    error: !ok ? pickFirst(result && result.error) || "Missing preview PDF payload." : "",
  };
}

async function main() {
  const workspaceBaseUrl =
    pickFirst(process.env.PREVIEW_HEALTH_WORKSPACE_BASE_URL) ||
    "https://innovate-calgary.rocketlane.com";
  const maxAllowedMs = parseIntegerEnv("PREVIEW_HEALTH_MAX_MS", 5000);
  const attemptsPerInvoice = parseIntegerEnv("PREVIEW_HEALTH_ATTEMPTS_PER_INVOICE", 2);
  const invoiceNumbers = String(
    process.env.PREVIEW_HEALTH_INVOICE_NUMBERS || "INV-000194,INV-000164"
  )
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  if (!invoiceNumbers.length) {
    throw new Error("No invoice numbers configured for preview health check.");
  }

  const report = {
    generatedAt: new Date().toISOString(),
    workspaceBaseUrl,
    maxAllowedMs,
    attemptsPerInvoice,
    invoiceNumbers,
    checks: [],
    failures: [],
  };

  for (const invoiceNumber of invoiceNumbers) {
    for (let attempt = 1; attempt <= attemptsPerInvoice; attempt += 1) {
      try {
        const probe = await runPreviewProbe(invoiceNumber, workspaceBaseUrl);
        report.checks.push({
          invoiceNumber,
          attempt,
          ok: probe.ok,
          elapsedMs: probe.elapsedMs,
          previewPipeline: probe.previewPipeline,
          previewSource: probe.previewSource,
          error: probe.error,
        });
        if (!probe.ok || probe.elapsedMs > maxAllowedMs) {
          report.failures.push({
            invoiceNumber,
            attempt,
            reason: !probe.ok
              ? "native-preview-missing"
              : `preview-too-slow-${probe.elapsedMs}ms`,
            detail: probe,
          });
        }
      } catch (error) {
        report.failures.push({
          invoiceNumber,
          attempt,
          reason: "preview-check-threw",
          detail: String(error && error.message ? error.message : error),
        });
      }
    }
  }

  const reportPath = path.resolve(
    process.cwd(),
    process.env.PREVIEW_HEALTH_REPORT_PATH || "artifacts/preview-health-report.json"
  );
  fs.mkdirSync(path.dirname(reportPath), { recursive: true });
  fs.writeFileSync(reportPath, JSON.stringify(report, null, 2), "utf8");

  // Print concise report for CI logs.
  console.log(
    JSON.stringify(
      {
        generatedAt: report.generatedAt,
        workspaceBaseUrl: report.workspaceBaseUrl,
        checks: report.checks.length,
        failures: report.failures.length,
        maxAllowedMs: report.maxAllowedMs,
      },
      null,
      2
    )
  );

  if (report.failures.length > 0) {
    const summary = report.failures
      .slice(0, 5)
      .map((failure) => `${failure.invoiceNumber}#${failure.attempt}:${failure.reason}`)
      .join(", ");
    throw new Error(
      `Preview health check failed (${report.failures.length} issue(s)): ${summary}`
    );
  }
}

main().catch((error) => {
  console.error(String(error && error.stack ? error.stack : error));
  process.exit(1);
});
