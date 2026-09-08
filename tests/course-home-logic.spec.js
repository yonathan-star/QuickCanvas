const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright");

const root = path.resolve(__dirname, "..");
const script = fs.readFileSync(path.join(root, "extension", "content.js"), "utf8");
const styles = fs.readFileSync(path.join(root, "extension", "content.css"), "utf8");
const fixture = `<!doctype html><html><head><title>Syllabus · Biology</title><style>${styles}</style></head><body>
  <div class="ic-app-crumbs"><ol class="ic-app-crumbs__crumbs"><li><a href="/courses/1">Biology</a></li><li class="ic-app-crumbs__crumb--current">Syllabus</li></ol></div>
  <aside id="left-side"><nav class="course-navigation"><ul id="section-tabs"><li class="section"><a href="/courses/1">Home</a></li><li class="section active"><a href="/courses/1/assignments/syllabus" aria-current="page">Syllabus</a></li></ul></nav></aside>
  <div class="ic-Layout-contentWrapper"><main id="content"><h1>Course Syllabus</h1><article id="course_syllabus">Teacher content</article></main><aside id="right-side"></aside></div>
</body></html>`;

(async () => {
  const browser = await chromium.launch({ headless: true, executablePath: "/usr/bin/google-chrome", args: ["--no-sandbox"] });
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  try {
    await page.addInitScript(() => {
      const sync = {
        canvasSettings: { enabled: true, baseUrl: "https://canvas.test" },
        popupTheme: { mode: "light", accent: "#1f5f8b" },
        cfeAuthGateMirror: { authenticated: true, userId: "test-user" },
      };
      const get = async (keys) => typeof keys === "string"
        ? { [keys]: sync[keys] }
        : Object.fromEntries((keys || []).map((key) => [key, sync[key]]));
      window.chrome = {
        runtime: { id: "test", getURL: (value) => value, sendMessage: async () => ({}), onMessage: { addListener() {} } },
        storage: {
          sync: { get, set: async (values) => Object.assign(sync, values) },
          local: { get: async () => ({ cfeAuthState: { authenticated: true, userId: "test-user" } }), set: async () => {} },
          onChanged: { addListener() {}, removeListener() {} },
        },
      };
    });
    await page.route("https://canvas.test/**", async (route) => {
      const url = new URL(route.request().url());
      if (!url.pathname.startsWith("/api/v1/")) return route.fulfill({ contentType: "text/html", body: fixture });
      if (url.pathname === "/api/v1/courses/1") {
        return route.fulfill({ contentType: "application/json", body: JSON.stringify({ id: 1, name: "Biology", course_code: "BIO 101" }) });
      }
      if (url.pathname.endsWith("/users/self/progress")) {
        return route.fulfill({ contentType: "application/json", body: JSON.stringify({ requirement_count: 10, requirement_completed_count: 4, next_requirement_url: "https://canvas.test/courses/1/modules/items/8" }) });
      }
      await new Promise((resolve) => setTimeout(resolve, 1800));
      const body = url.pathname.endsWith("/assignments")
        ? [{ id: 9, name: "Field notes", due_at: new Date(Date.now() + 86400000).toISOString(), html_url: "https://canvas.test/courses/1/assignments/9" }]
        : [];
      return route.fulfill({ contentType: "application/json", body: JSON.stringify(body) });
    });

    await page.goto("https://canvas.test/courses/1");
    const startedAt = Date.now();
    await page.addScriptTag({ content: script });
    await page.locator('[data-cfe-experience="course-home"]').waitFor({ timeout: 1200 });
    assert.ok(Date.now() - startedAt < 1200, "essential Course Home shell waited for secondary APIs");
    assert.equal(await page.locator(".cfe-course-context-panel section:first-child header strong").textContent(), "40%");
    assert.equal(await page.locator("#section-tabs .section.active").textContent(), "Home");
    assert.equal(await page.locator(".cfe-course-home-experience").count(), 1);
    await page.getByText("Field notes", { exact: true }).waitFor({ timeout: 4000 });
    assert.equal(await page.locator(".cfe-course-home-experience").count(), 1);
    console.log("Course Home routing, progress, and staged-loading checks passed.");
  } finally {
    await browser.close();
  }
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
