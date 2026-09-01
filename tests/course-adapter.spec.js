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
  const sections = [
    "Home", "Announcements", "Syllabus", "Modules", "Assignments",
    "Discussions", "Grades", "People", "Pages", "Files", "Quizzes",
  ];
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
            ${sections.map((section) => `<li class="section ${activeSection === section ? "active" : ""}"><a>${section}</a></li>`).join("")}
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
  const page = await browser.newPage({ viewport: options.viewport || { width: 1280, height: 800 } });
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
      const announcements = requestUrl.searchParams.get("only_announcements") === "true";
      return route.fulfill({
        contentType: "application/json",
        body: JSON.stringify([
          {
            id: 51,
            title: announcements ? "Field lab moved" : "Derivative strategies",
            message: announcements ? "<p>Meet at the east entrance.</p>" : "<p>Compare two solution methods.</p>",
            posted_at: "2026-08-31T13:00:00Z",
            pinned: true,
            read_state: "unread",
            author: { display_name: "Dr. Test" },
            html_url: `https://canvas.test/courses/10585/${announcements ? "announcements" : "discussion_topics"}/51`,
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
    const collectionFixtures = {
      "/api/v1/courses/10585/modules": [{ id: 1, name: "Limits", published: true, items: [{ id: 11, title: "Limits overview", type: "Page", html_url: "https://canvas.test/courses/10585/pages/limits" }] }],
      "/api/v1/courses/10585/assignment_groups": [{ id: 4, name: "Practice" }],
      "/api/v1/courses/10585/assignments": [{ id: 7, name: "Chapter review", assignment_group_id: 4, due_at: "2026-09-08T13:00:00Z", points_possible: 20, submission_types: ["online_upload"], html_url: "https://canvas.test/courses/10585/assignments/7", submission: { workflow_state: "graded", score: 18, grade: "18" } }],
      "/api/v1/courses/10585/users": [{ id: 9, display_name: "Ada Student", sortable_name: "Student, Ada", enrollments: [{ type: "StudentEnrollment", enrollment_state: "active", course_section_id: 3 }] }],
      "/api/v1/courses/10585/quizzes": [{ id: 10, title: "Limits check", due_at: "2026-09-10T13:00:00Z", question_count: 8, points_possible: 10, html_url: "https://canvas.test/courses/10585/quizzes/10" }],
      "/api/v1/courses/10585/files": [{ id: 12, display_name: "Review guide.pdf", size: 20480, modified_at: "2026-09-01T13:00:00Z", content_type: "application/pdf", url: "https://canvas.test/files/12/download" }],
      "/api/v1/courses/10585/pages": [{ url: "limits", title: "Limits overview", front_page: true, updated_at: "2026-09-01T13:00:00Z", html_url: "https://canvas.test/courses/10585/pages/limits" }],
      "/api/v1/courses/10585/pages/limits": { url: "limits", title: "Limits overview", updated_at: "2026-09-01T13:00:00Z", body: "<h2>Learning goals</h2><p>Evaluate limits graphically.</p>", html_url: "https://canvas.test/courses/10585/pages/limits" },
    };
    if (Object.hasOwn(collectionFixtures, requestUrl.pathname)) {
      return route.fulfill({ contentType: "application/json", body: JSON.stringify(collectionFixtures[requestUrl.pathname]) });
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

  if (activeSection === "Modules") {
    await page.locator("[data-cfe-module-toggle]").first().click();
  }
  if (["Assignments", "Grades", "People", "Files", "Quizzes"].includes(activeSection)) {
    await page.locator("[data-cfe-collection-search]").fill("definitely-no-match");
  }

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
      collectionRows: document.querySelectorAll(
        "[data-cfe-collection-row], [data-cfe-announcement-id]",
      ).length,
      experienceTitle: document.querySelector(".cfe-course-data-experience h1")?.textContent,
      moduleCollapsed:
        document.querySelector("[data-cfe-module-toggle]")?.getAttribute("aria-expanded") === "false" &&
        Boolean(document.querySelector("[data-cfe-module-items]")?.hidden),
      searchFiltered:
        !document.querySelector("[data-cfe-collection-search]") ||
        Array.from(document.querySelectorAll("[data-cfe-collection-row]"))
          .every((row) => row.hidden),
      horizontalOverflow:
        document.documentElement.scrollWidth > document.documentElement.clientWidth + 1,
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

    const routeCases = [
      ["modules", "Modules", "Modules"],
      ["assignments", "Assignments", "Assignments"],
      ["discussion_topics", "Discussions", "Discussions"],
      ["grades", "Grades", "Grades"],
      ["users", "People", "People"],
      ["pages", "Pages", "Pages & Files"],
      ["files", "Files", "Course Files"],
      ["quizzes", "Quizzes", "Quizzes & Assessments"],
    ];
    for (const [pathName, section, title] of routeCases) {
      const current = await runCase(
        browser,
        `https://canvas.test/courses/10585/${pathName}`,
        section,
      );
      assert.equal(current.identityCount, 1, `${section}: duplicate identity`);
      assert.equal(current.dataExperience, section.toLowerCase(), `${section}: wrong experience`);
      assert.equal(current.nativeHidden, true, `${section}: native content remains visible`);
      assert.equal(current.experienceTitle, title, `${section}: wrong title`);
      assert.ok(current.collectionRows >= 1, `${section}: no data rows rendered`);
      assert.equal(current.horizontalOverflow, false, `${section}: horizontal overflow`);
      if (section === "Modules") {
        assert.equal(current.moduleCollapsed, true, "Modules: collapse control failed");
      }
      if (["Assignments", "Grades", "People", "Files", "Quizzes"].includes(section)) {
        assert.equal(current.searchFiltered, true, `${section}: search control failed`);
      }
    }

    const mobile = await runCase(
      browser,
      "https://canvas.test/courses/10585/pages",
      "Pages",
      { viewport: { width: 390, height: 844 } },
    );
    assert.equal(mobile.dataExperience, "pages");
    assert.equal(mobile.horizontalOverflow, false, "Pages: mobile horizontal overflow");

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
