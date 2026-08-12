#!/usr/bin/env node

import crypto from "crypto";
import fs from "fs";
import path from "path";

const LOCK_PATH = path.resolve(process.cwd(), ".app-lock.json");
const LOCKED_FILES = [
  "dist/app.js",
  "dist/index.html",
  "dist/styles.css",
  "index.js",
  "server-actions/sync-invoices-from-source.js",
  "server-actions/sync-invoice-preview.js",
];

function readFileUtf8(relativePath) {
  const absolutePath = path.resolve(process.cwd(), relativePath);
  return fs.readFileSync(absolutePath, "utf8");
}

function hashText(content) {
  return crypto.createHash("sha256").update(content).digest("hex");
}

function computeLockEntries() {
  const entries = {};
  LOCKED_FILES.forEach((relativePath) => {
    const content = readFileUtf8(relativePath);
    entries[relativePath] = hashText(content);
  });
  return entries;
}

function writeLockFile(entries) {
  const payload = {
    version: 1,
    generatedAt: new Date().toISOString(),
    files: entries,
  };
  fs.writeFileSync(LOCK_PATH, JSON.stringify(payload, null, 2) + "\n", "utf8");
}

function verifyLockFile() {
  if (!fs.existsSync(LOCK_PATH)) {
    throw new Error(
      "Missing .app-lock.json. Run `npm run lock:update` to create the baseline lock file."
    );
  }
  const raw = fs.readFileSync(LOCK_PATH, "utf8");
  let parsed = null;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    throw new Error(`Invalid .app-lock.json JSON: ${String(error.message || error)}`);
  }
  const baselineFiles = parsed && parsed.files && typeof parsed.files === "object" ? parsed.files : {};
  const currentEntries = computeLockEntries();
  const allPaths = Array.from(
    new Set(Object.keys(currentEntries).concat(Object.keys(baselineFiles)))
  ).sort();
  const mismatches = [];
  allPaths.forEach((relativePath) => {
    const expected = String(baselineFiles[relativePath] || "");
    const actual = String(currentEntries[relativePath] || "");
    if (expected !== actual) {
      mismatches.push({
        file: relativePath,
        expected,
        actual,
      });
    }
  });
  if (mismatches.length) {
    const details = mismatches
      .map((entry) => {
        return (
          `- ${entry.file}\n` +
          `  expected: ${entry.expected || "(missing)"}\n` +
          `  actual:   ${entry.actual || "(missing)"}`
        );
      })
      .join("\n");
    throw new Error(
      "App lock verification failed. Locked files changed.\n" +
        details +
        "\nIf these changes are intentional, run `npm run lock:update` and commit the updated .app-lock.json."
    );
  }
  console.log("App lock verification passed.");
}

function main() {
  const mode = String(process.argv[2] || "verify").trim().toLowerCase();
  if (mode === "update") {
    const entries = computeLockEntries();
    writeLockFile(entries);
    console.log(`Updated lock file: ${path.relative(process.cwd(), LOCK_PATH)}`);
    return;
  }
  if (mode === "verify") {
    verifyLockFile();
    return;
  }
  throw new Error("Unknown mode. Use `verify` or `update`.");
}

try {
  main();
} catch (error) {
  console.error(String(error && error.stack ? error.stack : error));
  process.exit(1);
}
