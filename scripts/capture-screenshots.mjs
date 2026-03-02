import path from "node:path";
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
  "?demo=1&account=Acme%20Implementation%20Team&project=Customer%20Home";

const browser = await chromium.launch({
  headless: true,
});

const page = await browser.newPage({
  viewport: { width: 1620, height: 980 },
});

await page.goto(appUrl);
await page.waitForTimeout(700);

await page.screenshot({
  path: path.join(screenshotDir, "bulletin-board-overview.png"),
  fullPage: true,
});

const secondCard = page.locator(".note-card").nth(1);
if ((await secondCard.count()) > 0) {
  await secondCard.click();
}

await page.locator("#noteBodyInput").click();
await page.waitForTimeout(250);

await page.screenshot({
  path: path.join(screenshotDir, "bulletin-board-lists-and-formatting.png"),
  fullPage: true,
});

await browser.close();

console.log("Screenshots captured in docs/screenshots");
