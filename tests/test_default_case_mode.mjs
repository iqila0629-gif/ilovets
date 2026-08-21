import assert from "node:assert/strict";
import { createRequire } from "node:module";
import fs from "node:fs";


const require = createRequire(import.meta.url);
const playwrightPackage = process.env.PLAYWRIGHT_PACKAGE || "playwright";
const { chromium } = require(playwrightPackage);
const browserCandidates = [
  process.env.CHROME_PATH,
  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
].filter(Boolean);
const browserExecutable = browserCandidates.find((candidate) => fs.existsSync(candidate));


const pageUrl = new URL("../index.html", import.meta.url).href;


async function getKeys(page, text, mode) {
  return page.evaluate(
    ({ text: inputText, mode: inputMode }) => window.keysForLine(inputText, inputMode)
      .map((word) => word.map((item) => item.key)),
    { text, mode },
  );
}


async function run() {
  const browser = await chromium.launch({
    headless: true,
    ...(browserExecutable ? { executablePath: browserExecutable } : {}),
  });
  const page = await browser.newPage();
  await page.goto(pageUrl);
  await page.waitForLoadState("networkidle");

  assert.equal(
    await page.locator('input[name="caseMode"][value="default"]').isChecked(),
    true,
  );
  assert.deepEqual(await getKeys(page, "The BLACK Dog", "default"), [
    ["T", "h", "e"],
    ["B", "L", "A", "C", "K"],
    ["D", "o", "g"],
  ]);

  assert.deepEqual(await getKeys(page, "The BLACK Dog", "upper"), [
    ["T", "H", "E"],
    ["B", "L", "A", "C", "K"],
    ["D", "O", "G"],
  ]);
  assert.deepEqual(await getKeys(page, "The BLACK Dog", "lower"), [
    ["t", "h", "e"],
    ["b", "l", "a", "c", "k"],
    ["d", "o", "g"],
  ]);

  await browser.close();
  console.log("default case mode tests passed");
}


run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
