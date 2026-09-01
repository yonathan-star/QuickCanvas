const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright");

const root = path.resolve(__dirname, "..");
const contentScript = fs.readFileSync(
  path.join(root, "extension", "content.js"),
  "utf8",
);
const contentStyles = fs.readFileSync(
  path.join(root, "extension", "content.css"),
  "utf8",
);

function fixture(activeSection, duplicateCount = 6) {
  const identities = Array.from(
    { length: duplicateCount },
    () => `
      <div class="cfe-course-identity">
        <span class="cfe-course-identity-label">Course</span>
        <strong class="cfe-course-identity-title">AP Calculus AB</strong>
        <small class="cfe-course-identity-status">Active course</small>
      </div>`,
  ).join("");
  return `<!doctype html>
    <html><head><title>${activeSection} · AP Calculus AB</title><style>${contentStyles}</style></head><body>
      <div class="ic-app-crumbs"><ol class="ic-app-crumbs__crumbs">
        <li><a href="/courses/10585">AP Calculus AB</a></li>
        <li class="ic-app-crumbs__crumb--current">${activeSection}</li>
      </ol></div>
      <div class="ic-Layout-columns">
        <aside id="left-side"><div class="course-navigation">${identities}
          <ul id="section-tabs">
            <li class="section"><a>Home</a></li>
            <li class="section ${activeSection === "Announcements" ? "active" : ""}"><a>Announcements</a></li>
            <li class="section ${activeSection === "Syllabus" ? "active" : ""}"><a>Syllabus</a></li>
          </ul>
        </div></aside>
        <main id="content" class="ic-Layout-contentMain">
          <div class="ic-Action-header"><h1>AP Calculus AB Per C-1233</h1></div>
          <article id="course_syllabus"><h2>Course information</h2><p>Policies and expectations.</p></article>
        </main>
        <aside id="right-side"><section id="cfe-course-widget-board">Wrong home widgets</section></aside>
      </div>
    </body></html>`;
}

async function runCase(browser, url, activeSection, options = {}) {
  const page = await browser.newPage();
  await page.addInitScript(() => {
    const settings = {
      canvasSettings: { enabled: true, baseUrl: "https://canvas.test" },
      popupTheme: { mode: "light", accent: "#1f5f8b" },
      cfeAuthGateMirror: { authenticated: true, userId: "test-user" },
    };
    const getValues = async (keys) => {
      if (Array.isArray(keys)) {
        return Object.fromEntries(keys.map((key) => [key, settings[key]]));
      }
      return { [keys]: settings[keys] };
    };
    window.chrome = {
      runtime: {
        id: "quickcanvas-test",
        getURL: (value) => `https://extension.test/${value}`,
        onMessage: { addListener() {} },
        sendMessage: async () => ({}),
      },
      storage: {
        sync: { get: getValues, set: async () => {} },
        local: {
          get: async (key) =>
            key === "cfeAuthState"
              ? { cfeAuthState: { authenticated: true, userId: "test-user" } }
              : { [key]: null },
          set: async () => {},
        },
        onChanged: { addListener() {}, removeListener() {} },
      },
    };
  });
  await page.route("https://canvas.test/**", (route) => {
    const requestUrl = new URL(route.request().url());
    if (options.apiFailure && requestUrl.pathname.startsWith("/api/v1/")) {
      return route.fulfill({ status: 403, body: "Forbidden" });
    }
    if (requestUrl.pathname.endsWith("/api/v1/courses/10585/discussion_topics")) {
      return route.fulfill({
        contentType: "application/json",
        body: JSON.stringify([
          {
            id: 51,
            title: "Field lab moved",
            message: "<p>Meet at the east entrance.</p>",
            posted_at: "2026-08-31T13:00:00Z",
            pinned: true,
            read_state: "unread",
            author: { display_name: "Dr. Test" },
            html_url: "https://canvas.test/courses/10585/announcements/51",
          },
          {
            id: 52,
            title: "Week five resources",
            message: "<p>The review guide is ready.</p>",
            posted_at: "2026-08-30T13:00:00Z",
            pinned: false,
            read_state: "read",
            author: { display_name: "Course Team" },
            html_url: "https://canvas.test/courses/10585/announcements/52",
          },
        ]),
      });
    }
    if (requestUrl.pathname.endsWith("/api/v1/courses/10585")) {
      return route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({
          id: 10585,
          name: "AP Calculus AB",
          course_code: "AP CALC AB",
          workflow_state: "available",
          updated_at: "2026-08-31T13:00:00Z",
          term: { name: "2026 ALL" },
          teachers: [{ display_name: "Mrs. Kibler" }],
          syllabus_body:
            "<h2>Course overview</h2><p>Limits, derivatives, and integrals.</p><h2>Course policies</h2><p>Submit work through Canvas.</p>",
        }),
      });
    }
    return route.fulfill({
      contentType: "text/html",
      body: fixture(activeSection),
    });
  });
  await page.goto(url);
  await page.addScriptTag({ content: contentScript });
  await page.waitForTimeout(2600);
  await page.evaluate(() => {
    window.dispatchEvent(new Event("focus"));
    window.dispatchEvent(new PageTransitionEvent("pageshow"));
  });
  await page.waitForTimeout(1900);

  if (process.env.CFE_CAPTURE_DIR) {
    fs.mkdirSync(process.env.CFE_CAPTURE_DIR, { recursive: true });
    await page.screenshot({
      path: path.join(
        process.env.CFE_CAPTURE_DIR,
        `${activeSection.toLowerCase()}-experience.png`,
      ),
      fullPage: true,
    });
  }

  const result = await page.evaluate(() => {
    const identities = [...document.querySelectorAll(".cfe-course-identity")];
    const tabs = document.querySelector("#section-tabs");
    return {
      identityCount: identities.length,
      identityIsImmediatelyBeforeTabs:
        identities[0]?.parentElement === tabs?.parentElement &&
        identities[0]?.nextElementSibling === tabs,
      syllabus: document.body.classList.contains("cfe-page-syllabus"),
      announcements: document.body.classList.contains(
        "cfe-page-announcements",
      ),
      courseHome: document.body.classList.contains("cfe-page-course-home"),
      title: document.querySelector("#content h1")?.textContent,
      hasWrongHomeWidgets: Boolean(
        document.querySelector("#cfe-course-widget-board"),
      ),
      dataExperience:
        document.querySelector(".cfe-course-data-experience")?.getAttribute(
          "data-cfe-experience",
        ) || "",
      nativeHidden: Boolean(
        document.querySelector("#content > .cfe-native-course-hidden"),
      ),
      syllabusBody: document.querySelector(".cfe-syllabus-body")?.textContent,
      announcementDetail: document.querySelector(
        ".cfe-announcement-detail h2",
      )?.textContent,
      announcementRows: document.querySelectorAll(
        ".cfe-announcement-row",
      ).length,
    };
  });
  await page.close();
  return result;
}

(async () => {
  const launchOptions = { headless: true };
  const chromePath = process.env.CHROME_BIN || "/usr/bin/google-chrome";
  if (fs.existsSync(chromePath)) launchOptions.executablePath = chromePath;
  const browser = await chromium.launch(launchOptions);
  try {
    const syllabus = await runCase(
      browser,
      "https://canvas.test/courses/10585",
      "Syllabus",
    );
    assert.equal(syllabus.identityCount, 1);
    assert.equal(syllabus.identityIsImmediatelyBeforeTabs, true);
    assert.equal(syllabus.syllabus, true);
    assert.equal(syllabus.courseHome, false);
    assert.equal(syllabus.title, "Course Syllabus");
    assert.equal(syllabus.hasWrongHomeWidgets, false);
    assert.equal(syllabus.dataExperience, "syllabus");
    assert.equal(syllabus.nativeHidden, true);
    assert.match(syllabus.syllabusBody, /Limits, derivatives/);

    const announcements = await runCase(
      browser,
      "https://canvas.test/courses/10585/announcements",
      "Announcements",
    );
    assert.equal(announcements.identityCount, 1);
    assert.equal(announcements.identityIsImmediatelyBeforeTabs, true);
    assert.equal(announcements.announcements, true);
    assert.equal(announcements.courseHome, false);
    assert.equal(announcements.dataExperience, "announcements");
    assert.equal(announcements.nativeHidden, true);
    assert.equal(announcements.announcementDetail, "Field lab moved");
    assert.equal(announcements.announcementRows, 2);

    const fallback = await runCase(
      browser,
      "https://canvas.test/courses/10585",
      "Syllabus",
      { apiFailure: true },
    );
    assert.equal(fallback.dataExperience, "");
    assert.equal(fallback.nativeHidden, false);
    console.log("Course adapter regression checks passed.");
  } finally {
    await browser.close();
  }
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
