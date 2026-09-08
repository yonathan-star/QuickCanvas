const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright");

const root = path.resolve(__dirname, "..");
const css = fs.readFileSync(path.join(root, "extension", "popup.css"), "utf8");
const html = fs
  .readFileSync(path.join(root, "extension", "popup.html"), "utf8")
  .replace('<link rel="stylesheet" href="popup.css" />', `<style>${css}</style>`)
  .replace(/<script[^>]*><\/script>/g, "");

(async () => {
  const browser = await chromium.launch({
    headless: true,
    executablePath: "/usr/bin/google-chrome",
    args: ["--no-sandbox"],
  });
  const page = await browser.newPage({ viewport: { width: 392, height: 600 } });
  try {
    await page.setContent(html);
    await page.evaluate(() => {
      document.querySelectorAll("[data-pane]").forEach((pane) => { pane.hidden = true; });
      document.querySelector(".auth-panel").hidden = false;
    });
    const signedOut = await page.evaluate(() => {
      const top = document.querySelector(".top").getBoundingClientRect();
      const auth = document.querySelector(".auth-panel").getBoundingClientRect();
      const primary = document.querySelector(".auth-panel .primary").getBoundingClientRect();
      return {
        tabs: getComputedStyle(document.querySelector(".tabs")).display,
        gap: Math.round(auth.top - top.bottom),
        overflow: document.documentElement.scrollWidth > document.documentElement.clientWidth,
        primaryWidth: Math.round(primary.width),
        panelInnerWidth: Math.round(auth.width - 34),
      };
    });
    assert.equal(signedOut.tabs, "none");
    assert.equal(signedOut.gap, 10);
    assert.equal(signedOut.overflow, false);
    assert.equal(signedOut.primaryWidth, signedOut.panelInnerWidth);
    console.log("Popup signed-out layout checks passed.");
  } finally {
    await browser.close();
  }
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
