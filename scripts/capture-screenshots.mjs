import path from "node:path";
import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");

const appFile = path.join(rootDir, "dist", "index.html");
const screenshotDir = path.join(rootDir, "docs", "screenshots");
const appUrl =
  "file://" +
  appFile +
  "?account=Acme%20Implementation%20Team&project=Customer%20Portal%20Onboarding&email=admin%40acme.com&role=admin";

await fs.mkdir(screenshotDir, { recursive: true });

const browser = await chromium.launch({
  headless: true,
});

const page = await browser.newPage({
  viewport: { width: 1620, height: 980 },
});

await page.goto(appUrl);
await page.waitForTimeout(700);

await page.screenshot({
  path: path.join(screenshotDir, "invoice-access-overview.png"),
  fullPage: true,
});

await page.locator("#tabLogsButton").click();
await page.waitForTimeout(250);

await page.screenshot({
  path: path.join(screenshotDir, "invoice-access-admin-diagnostics.png"),
  fullPage: true,
});

await browser.close();

console.log("Screenshots captured in docs/screenshots");
