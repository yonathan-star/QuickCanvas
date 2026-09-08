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
        <li><a href="/courses/10585">AP Calculus AB Per C-1233-ALL-Kibler</a></li>
        <li class="ic-app-crumbs__crumb--current">${activeSection}</li>
      </ol></div>
      <div class="ic-Layout-columns">
        <aside id="left-side"><div class="course-navigation">${identities}
          <ul id="section-tabs">
            ${sections.map((section) => `<li class="section ${activeSection === section ? "active" : ""}"><a href="/courses/10585/${section === "Home" ? "" : section.toLowerCase()}">${section}</a></li>`).join("")}
          </ul>
        </div></aside>
        <div class="ic-Layout-contentWrapper">
          <main id="content" class="ic-Layout-contentMain">
            <div class="ic-Action-header"><h1>AP Calculus AB Per C-1233</h1></div>
            <article id="course_syllabus"><h2>Course information</h2><p>Policies and expectations.</p></article>
            ${/Panopto Recordings|Google Drive|Mystery Tool/.test(activeSection) ? '<div class="tool_content_wrapper"><iframe id="tool_content_642" name="tool_content_642" src="about:blank"></iframe></div>' : ""}
          </main>
          <aside id="right-side"><section id="cfe-course-widget-board">Wrong home widgets</section></aside>
        </div>
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
    if (requestUrl.pathname.endsWith("/api/v1/courses/10585/discussion_topics/51/view")) {
      return route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({
          view: [
            {
              id: 501,
              user_name: "Ada Student",
              created_at: "2026-09-01T13:00:00Z",
              message: "<p>I compared both strategies.</p>",
            },
          ],
        }),
      });
    }
    if (requestUrl.pathname.endsWith("/api/v1/courses/10585/discussion_topics/51")) {
      return route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({
          id: 51,
          title: "Field lab moved",
          message: "<p>Meet at the east entrance.</p>",
          posted_at: "2026-08-31T13:00:00Z",
          pinned: true,
          read_state: "unread",
          author: { display_name: "Dr. Test" },
          html_url: "https://canvas.test/courses/10585/announcements/51",
          attachments: [
            {
              display_name: "field-map.pdf",
              content_type: "application/pdf",
              url: "https://canvas.test/files/71/download",
            },
          ],
        }),
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
      "/api/v1/courses/10585/groups": [{ id: 31, name: "Calculus Study Group", members_count: 4 }],
    };
    if (Object.hasOwn(collectionFixtures, requestUrl.pathname)) {
      return route.fulfill({ contentType: "application/json", body: JSON.stringify(collectionFixtures[requestUrl.pathname]) });
    }
    if (requestUrl.pathname.endsWith("/api/v1/courses/10585")) {
      return route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({
          id: 10585,
          name: "AP Calculus AB Per C-1233-ALL-Kibler",
          course_code: "AP Calculus AB Per C-1233-ALL-Kibler",
          workflow_state: "available",
          updated_at: "2026-08-31T13:00:00Z",
          term: { name: "2026 ALL" },
          sections: [{ id: 3, name: "Period C-1233" }],
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

  const isExternalTool = /\/external_tools\//.test(new URL(url).pathname);
  const isNativeWorkflow = options.apiFailure || isExternalTool || /\/(?:assignments|discussion_topics|quizzes)\/\d+/.test(new URL(url).pathname);
  const dataExperienceKinds = {
    Home: "course-home",
    Modules: "modules",
    Assignments: "assignments",
    Discussions: "discussions",
    Grades: "grades",
    People: "people",
    Pages: "pages",
    Files: "files",
    Quizzes: "quizzes",
    Announcements: "announcements",
    Syllabus: "syllabus",
  };
  const expectedDataKind = new URL(url).searchParams.get("quickcanvas_home") === "1"
    ? "course-home"
    : dataExperienceKinds[activeSection];
  const expectedExperience = isExternalTool
    ? ".cfe-tool-toolbar"
    : isNativeWorkflow
      ? ".cfe-course-page-header"
      : `[data-cfe-experience="${expectedDataKind}"]`;
  try {
    await page.locator(expectedExperience).first().waitFor({
      state: "visible",
      timeout: 12000,
    });
  } catch (error) {
    const adapterState = await page.evaluate(() => ({
      href: window.location.href,
      title: document.title,
      activeSection: document.querySelector("#section-tabs .section.active")?.textContent,
      currentCrumb: document.querySelector(".ic-app-crumbs__crumb--current")?.textContent,
      bodyClasses: document.body.className,
      experiences: Array.from(
        document.querySelectorAll(".cfe-course-data-experience"),
      ).map((node) => ({
        kind: node.getAttribute("data-cfe-experience"),
        className: node.className,
        text: String(node.textContent || "").trim().slice(0, 120),
      })),
    }));
    console.error("Course adapter wait state:", adapterState);
    throw error;
  }
  if (activeSection === "Announcements") {
    await page.locator(".cfe-announcement-attachment").first().waitFor({
      state: "visible",
      timeout: 5000,
    });
  }
  if (activeSection === "Discussions" && !isNativeWorkflow) {
    await page.locator(".cfe-discussion-reply").first().waitFor({
      state: "visible",
      timeout: 5000,
    });
  }

  if (process.env.CFE_CAPTURE_DIR) {
    fs.mkdirSync(process.env.CFE_CAPTURE_DIR, { recursive: true });
    await page.screenshot({
      path: path.join(
        process.env.CFE_CAPTURE_DIR,
        `${expectedDataKind || activeSection.toLowerCase()}-experience.png`,
      ),
      fullPage: true,
    });
  }

  if (activeSection === "Modules") {
    await page.locator("[data-cfe-module-toggle]").first().click();
  }
  if (["Assignments", "Grades", "People", "Quizzes"].includes(activeSection)) {
    const search = page.locator("[data-cfe-collection-search]");
    if (await search.count()) await search.fill("definitely-no-match");
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
      contentWrapperBackground: getComputedStyle(
        document.querySelector(".ic-Layout-contentWrapper"),
      ).backgroundColor,
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
      announcementAttachments: document.querySelectorAll(
        ".cfe-announcement-attachment",
      ).length,
      discussionReplies: document.querySelectorAll(
        ".cfe-discussion-reply",
      ).length,
      collectionRows: document.querySelectorAll(
        "[data-cfe-collection-row], [data-cfe-announcement-id]",
      ).length,
      experienceTitle: document.querySelector(".cfe-course-data-experience h1")?.textContent,
      courseIdentityTitle: document.querySelector(".cfe-course-identity-title")?.textContent,
      courseHomeSvgIcons: document.querySelectorAll(
        ".cfe-course-home-experience svg",
      ).length,
      courseHomeMessageAction: Boolean(
        document.querySelector(".cfe-context-inline-action"),
      ),
      injectedHeaderCopies: document.querySelectorAll(
        ".cfe-course-data-experience .cfe-course-page-copy",
      ).length,
      adaptedHeaderCount: document.querySelectorAll(".cfe-course-page-header").length,
      adaptedCopyCount: document.querySelectorAll(".cfe-course-page-copy").length,
      visibleExternalNativeHeaders: Array.from(
        document.querySelectorAll("#content > .cfe-external-native-hidden"),
      ).filter((node) => getComputedStyle(node).display !== "none").length,
      toolbars: document.querySelectorAll(".cfe-tool-toolbar").length,
      toolFrameHosts: document.querySelectorAll(".cfe-tool-frame-host").length,
      toolStatePanels: document.querySelectorAll(".cfe-tool-states").length,
      pageTypeClasses: Array.from(document.body.classList).filter((name) =>
        name.startsWith("cfe-page-"),
      ),
      courseNavLabels: Array.from(
        document.querySelectorAll("#section-tabs > .section, #section-tabs > .cfe-course-tools-label"),
      ).map((node) => String(node.textContent || "").trim()),
      courseNavLinkHeight:
        document.querySelector("#section-tabs .section a")?.getBoundingClientRect().height || 0,
      activeCourseNav:
        document.querySelector("#section-tabs .section.active")?.textContent.trim() || "",
      identityLabelVisible:
        document.querySelector(".cfe-course-identity-label")?.getBoundingClientRect().height > 0,
      assignmentFiltersOverlap: (() => {
        const input = document.querySelector("[data-cfe-collection-search]");
        const firstFilter = document.querySelector("[data-cfe-assignment-filter]");
        if (!input || !firstFilter) return false;
        return input.getBoundingClientRect().right > firstFilter.getBoundingClientRect().left;
      })(),
      moduleCollapsed:
        document.querySelector("[data-cfe-module-toggle]")?.getAttribute("aria-expanded") === "false" &&
        Boolean(document.querySelector("[data-cfe-module-items]")?.hidden),
      searchFiltered:
        !document.querySelector("[data-cfe-collection-search]") ||
        Array.from(document.querySelectorAll("[data-cfe-collection-row]"))
          .every((row) => row.hidden),
      assignmentControls:
        document.querySelectorAll("[data-cfe-assignment-group-toggle]").length &&
        Boolean(document.querySelector("[data-cfe-more-filters]")),
      horizontalOverflow:
        document.documentElement.scrollWidth > document.documentElement.clientWidth + 1,
      bodyClasses: document.body.className,
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
    assert.equal(syllabus.courseNavLinkHeight, 36);
    assert.equal(syllabus.identityLabelVisible, false);
    assert.equal(syllabus.syllabus, true);
    assert.equal(syllabus.courseHome, false);
    assert.equal(syllabus.title, "Course Syllabus");
    assert.equal(syllabus.hasWrongHomeWidgets, false);
    assert.equal(syllabus.dataExperience, "syllabus");
    assert.equal(syllabus.nativeHidden, true);
    assert.equal(syllabus.contentWrapperBackground, "rgb(255, 255, 255)");
    assert.match(syllabus.syllabusBody, /Limits, derivatives/);
    assert.deepEqual(syllabus.courseNavLabels.slice(0, 10), [
      "Home",
      "Announcements",
      "Modules",
      "Assignments",
      "Quizzes",
      "Discussions",
      "Grades",
      "People",
      "Pages",
      "Files",
    ]);

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
    assert.equal(announcements.announcementAttachments, 1);

    const courseHome = await runCase(
      browser,
      "https://canvas.test/courses/10585",
      "Home",
    );
    assert.equal(courseHome.dataExperience, "course-home");
    assert.equal(courseHome.nativeHidden, true);
    assert.equal(courseHome.experienceTitle, "AP Calculus AB");
    assert.equal(courseHome.courseIdentityTitle, "AP Calculus AB");
    assert.ok(courseHome.courseHomeSvgIcons >= 8);
    assert.equal(courseHome.courseHomeMessageAction, true);
    assert.equal(courseHome.horizontalOverflow, false);

    const forcedCourseHome = await runCase(
      browser,
      "https://canvas.test/courses/10585?quickcanvas_home=1",
      "Syllabus",
    );
    assert.equal(forcedCourseHome.courseHome, true);
    assert.equal(forcedCourseHome.syllabus, false);
    assert.equal(forcedCourseHome.dataExperience, "course-home");
    assert.equal(forcedCourseHome.activeCourseNav, "Home");
    assert.equal(forcedCourseHome.experienceTitle, "AP Calculus AB");
    assert.equal(forcedCourseHome.horizontalOverflow, false);

    const routeCases = [
      ["modules", "Modules", "Modules"],
      ["assignments", "Assignments", "Assignments"],
      ["discussion_topics", "Discussions", "Discussions"],
      ["grades", "Grades", "Grades"],
      ["users", "People", "People"],
      ["pages", "Pages", "Pages & Files"],
      ["files", "Files", "Pages & Files"],
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
      assert.equal(
        current.contentWrapperBackground,
        "rgb(255, 255, 255)",
        `${section}: course workspace is not flush white`,
      );
      if (section === "Assignments") {
        assert.ok(current.assignmentControls, "Assignments: approved group/filter controls missing");
        assert.equal(current.assignmentFiltersOverlap, false, "Assignments: search and filters overlap");
      }
      if (section === "Discussions") {
        assert.equal(current.discussionReplies, 1, "Discussions: replies were not hydrated");
      }
      assert.equal(
        current.injectedHeaderCopies,
        0,
        `${section}: data header was reprocessed`,
      );
      const expectedPageClass = {
        Modules: "cfe-page-modules",
        Assignments: "cfe-page-assignments",
        Discussions: "cfe-page-discussions",
        Grades: "cfe-page-grades",
        People: "cfe-page-people",
        Pages: "cfe-page-pages",
        Files: "cfe-page-files",
        Quizzes: "cfe-page-assessments",
      }[section];
      assert.deepEqual(
        current.pageTypeClasses.sort(),
        ["cfe-page-course", expectedPageClass].sort(),
        `${section}: stale route classes leaked`,
      );
      if (section === "Modules") {
        assert.equal(current.moduleCollapsed, true, "Modules: collapse control failed");
      }
      if (["Assignments", "Grades", "People", "Quizzes"].includes(section)) {
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

    const canonicalSyllabus = await runCase(
      browser,
      "https://canvas.test/courses/10585/assignments/syllabus",
      "Syllabus",
    );
    assert.equal(canonicalSyllabus.dataExperience, "syllabus");
    assert.deepEqual(canonicalSyllabus.pageTypeClasses.sort(), [
      "cfe-page-course",
      "cfe-page-syllabus",
    ]);

    const nativeCases = [
      ["assignments/7", "Assignments", "cfe-page-assignment-detail"],
      ["discussion_topics/51", "Discussions", "cfe-page-discussion-detail"],
      ["quizzes/10", "Quizzes", "cfe-quiz-page"],
      ["external_tools/65", "Panopto Recordings", "cfe-page-media-tool"],
      ["external_tools/66", "Google Drive", "cfe-page-collaboration-tool"],
      ["external_tools/67", "Mystery Tool", "cfe-page-custom-tool"],
    ];
    for (const [pathName, section, expectedClass] of nativeCases) {
      const current = await runCase(
        browser,
        `https://canvas.test/courses/10585/${pathName}`,
        section,
      );
      assert.equal(current.dataExperience, "", `${section}: native workflow was replaced`);
      assert.equal(current.nativeHidden, false, `${section}: native workflow was hidden`);
      assert.match(current.bodyClasses, new RegExp(`(?:^|\\s)${expectedClass}(?:\\s|$)`));
      assert.equal(current.horizontalOverflow, false, `${section}: horizontal overflow`);
      if (pathName.startsWith("external_tools/")) {
        assert.equal(current.adaptedHeaderCount, 0, `${section}: duplicate page header remains`);
        assert.equal(current.adaptedCopyCount, 0, `${section}: duplicate page copy remains`);
        assert.equal(current.visibleExternalNativeHeaders, 0, `${section}: native header is visible`);
        assert.equal(current.toolbars, 1, `${section}: toolbar missing or repeated`);
        assert.equal(current.toolFrameHosts, 1, `${section}: frame host missing or repeated`);
        assert.equal(current.toolStatePanels, 1, `${section}: state panel missing or repeated`);
      } else {
        assert.equal(current.adaptedHeaderCount, 1, `${section}: header adapter repeated`);
        assert.equal(current.adaptedCopyCount, 1, `${section}: header copy repeated`);
      }
    }

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
