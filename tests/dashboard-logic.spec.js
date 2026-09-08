const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright");

const root = path.resolve(__dirname, "..");
const script = fs.readFileSync(path.join(root, "extension", "content.js"), "utf8");
const styles = fs.readFileSync(path.join(root, "extension", "content.css"), "utf8");

const shell = `<!doctype html><html><head><meta charset="utf-8"><style>${styles}</style></head><body>
  <header id="header"><a href="/">Canvas</a></header>
  <main id="main"><div id="content">Native dashboard</div></main>
</body></html>`;

(async () => {
  const browser = await chromium.launch({
    headless: true,
    executablePath: "/usr/bin/google-chrome",
    args: ["--no-sandbox"],
  });
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });

  try {
    await page.addInitScript(() => {
      const stores = {
        sync: {
          canvasSettings: { enabled: true, baseUrl: "https://canvas.test" },
          popupTheme: { mode: "light", accent: "#1f5f8b" },
          cfeAuthGateMirror: { authenticated: true, userId: "test-user" },
        },
        local: {
          cfeAuthState: { authenticated: true, userId: "test-user" },
        },
      };
      const listeners = new Set();
      const select = (store, keys) => {
        if (keys == null) return { ...store };
        if (typeof keys === "string") return { [keys]: store[keys] };
        if (Array.isArray(keys)) {
          return Object.fromEntries(keys.map((key) => [key, store[key]]));
        }
        return Object.fromEntries(
          Object.entries(keys).map(([key, fallback]) => [
            key,
            store[key] === undefined ? fallback : store[key],
          ]),
        );
      };
      const area = (name) => ({
        get: async (keys) => select(stores[name], keys),
        set: async (values) => {
          const changes = {};
          Object.entries(values || {}).forEach(([key, value]) => {
            changes[key] = { oldValue: stores[name][key], newValue: value };
            stores[name][key] = value;
          });
          listeners.forEach((listener) => listener(changes, name));
        },
        remove: async (keys) => {
          for (const key of [].concat(keys || [])) delete stores[name][key];
        },
      });
      window.chrome = {
        runtime: {
          id: "quickcanvas-test",
          getURL: (value) => `https://extension.test/${value}`,
          sendMessage: async () => ({}),
          onMessage: { addListener() {} },
        },
        storage: {
          sync: area("sync"),
          local: area("local"),
          onChanged: {
            addListener: (listener) => listeners.add(listener),
            removeListener: (listener) => listeners.delete(listener),
          },
        },
      };
      window.__setQuickCanvasAuth = async (authenticated) => {
        await window.chrome.storage.local.set({
          cfeAuthState: {
            authenticated,
            userId: authenticated ? "test-user" : "",
          },
        });
      };
    });

    let plannerRequests = 0;
    await page.route("https://canvas.test/**", async (route) => {
      const url = new URL(route.request().url());
      if (!url.pathname.startsWith("/api/v1/")) {
        return route.fulfill({ contentType: "text/html", body: shell });
      }
      const tomorrow = new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString();
      let body = [];
      if (url.pathname === "/api/v1/planner/items") {
        plannerRequests += 1;
        if (plannerRequests > 1) await new Promise((resolve) => setTimeout(resolve, 250));
        body = [{
          plannable_type: "assignment",
          plannable_id: 77,
          course_id: 1,
          plannable: { id: 77, title: "Lab report", due_at: tomorrow, points_possible: 10 },
          html_url: "https://canvas.test/courses/1/assignments/77",
          submissions: { workflow_state: "unsubmitted" },
        }];
      } else if (url.pathname === "/api/v1/courses") {
        body = [{ id: 1, name: "Biology" }];
      } else if (url.pathname === "/api/v1/users/self/course_nicknames") {
        body = [];
      } else if (url.pathname.includes("/assignments/77/submissions/self")) {
        body = { workflow_state: "unsubmitted" };
      }
      return route.fulfill({ contentType: "application/json", body: JSON.stringify(body) });
    });

    await page.goto("https://canvas.test/");
    await page.addScriptTag({ content: script });
    await page.locator("#cfe-dashboard").waitFor({ timeout: 10000 });
    await page.getByText("Lab report", { exact: true }).first().waitFor();

    const refresh = page.locator(".cfe-refresh");
    await refresh.click();
    await page.waitForFunction(() => {
      const button = document.querySelector(".cfe-refresh");
      return button?.disabled && button.getAttribute("aria-busy") === "true";
    });
    await page.waitForFunction(() => !document.querySelector(".cfe-refresh")?.disabled);
    assert.ok(plannerRequests >= 2, "manual refresh did not bypass the data cache");

    await page.locator('[data-filter="all"]').first().click();
    assert.equal(await page.locator('[data-filter="all"]').first().getAttribute("class").then((v) => v.includes("is-active")), true);

    await page.evaluate(() => window.__setQuickCanvasAuth(false));
    await page.locator("#cfe-auth-wall").waitFor({ timeout: 5000 });
    assert.equal(await page.locator("#cfe-dashboard").count(), 0);

    await page.evaluate(() => window.__setQuickCanvasAuth(true));
    await page.locator("#cfe-dashboard").waitFor({ timeout: 8000 });

    await page.evaluate(() => history.pushState({}, "", "/courses/1/modules"));
    await page.waitForFunction(() => !document.querySelector("#cfe-dashboard"), null, { timeout: 5000 });
    await page.evaluate(() => history.pushState({}, "", "/"));
    await page.locator("#cfe-dashboard").waitFor({ timeout: 8000 });

    const metrics = await page.evaluate(() => ({
      horizontalOverflow: document.documentElement.scrollWidth > document.documentElement.clientWidth,
      dashboardCount: document.querySelectorAll("#cfe-dashboard").length,
      build: document.documentElement.dataset.cfeBuild,
    }));
    assert.deepEqual(metrics, {
      horizontalOverflow: false,
      dashboardCount: 1,
      build: "0.9.11",
    });
    console.log("Dashboard logic regression checks passed.");
  } finally {
    await browser.close();
  }
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
