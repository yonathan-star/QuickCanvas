const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright");

const root = path.resolve(__dirname, "..");
const contentScript = fs.readFileSync(
  path.join(root, "extension", "content.js"),
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
    <html><head><title>${activeSection} · AP Calculus AB</title></head><body>
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

async function runCase(browser, url, activeSection) {
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
  await page.route("https://canvas.test/**", (route) =>
    route.fulfill({ contentType: "text/html", body: fixture(activeSection) }),
  );
  await page.goto(url);
  await page.addScriptTag({ content: contentScript });
  await page.waitForTimeout(2600);
  await page.evaluate(() => {
    window.dispatchEvent(new Event("focus"));
    window.dispatchEvent(new PageTransitionEvent("pageshow"));
  });
  await page.waitForTimeout(1900);

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

    const announcements = await runCase(
      browser,
      "https://canvas.test/courses/10585/announcements",
      "Announcements",
    );
    assert.equal(announcements.identityCount, 1);
    assert.equal(announcements.identityIsImmediatelyBeforeTabs, true);
    assert.equal(announcements.announcements, true);
    assert.equal(announcements.courseHome, false);
    console.log("Course adapter regression checks passed.");
  } finally {
    await browser.close();
  }
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
