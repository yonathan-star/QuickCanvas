# QuickCanvas Workspace

QuickCanvas connects to Canvas through the user's existing signed-in Canvas
browser session. It does not ask students to create or paste personal Canvas
access tokens. Enter the school's HTTPS Canvas URL in the extension, grant site
access, and sign in to Canvas normally.

Canvas OAuth is still the supported option for a standalone or multi-school
service. It requires an administrator-enabled Developer Key and a trusted
backend for the client secret; never embed that secret in this extension.

Version 0.9.14 keeps the popup on its last selected tab across closes and
browser restarts. Community themes are now always discoverable on the Themes
tab, the preset expander controls only the built-in preset list, and signed-in
users can access their Cloud Themes from Account again.

Version 0.9.13 implements the approved Superdesign course workspace as a
structural Canvas adapter rather than a generic skin. Native routes now use the
same 80px global rail, 208px course rail, 48px breadcrumb bar, content bounds,
page hierarchy, and compact card system as the reference. Panopto and other LTI
tools retain their real content inside the matching full-width tool shell. The
dashboard, popup, and existing Canvas interactions remain intact. The 0.8.8
course shell uses the approved route-specific widths, typography, list density,
side panels, and responsive navigation. Its bounded runtime prevents observer
feedback and duplicate LTI toolbars without recurring theme reapplication. It
also repairs duplicated course navigation identities and recognizes Canvas
sections such as Syllabus when a school serves them at the course-home URL.
Syllabus, Announcements, Modules, Assignments, Discussions, Grades, People,
Pages, Files, and Quizzes now use route-specific, data-backed Superdesign
layouts built from the user's signed-in Canvas session. Detail pages, active
assessments, submissions, and external tools preserve Canvas's native
functionality, and every enhanced route automatically restores the native page
if a school blocks a session endpoint. Course Home now follows the approved
editorial workspace with progress, announcements, due work, contacts, and
quick links. Modules, Assignments, Grades, People, Discussions, and the unified
Pages/Files workspace include the corresponding reference controls, status
panels, filters, and responsive structures. Office documents and module-file
items open through Canvas's native viewer with an explicit download fallback;
hidden Files routes, omitted module items, paginated lists, and slow Canvas
responses are handled without trapping or truncating the interface.
Dashboard checkmarks now update every matching task view and progress ring
immediately. Course and module progress use Canvas requirements when configured
and fall back to real assignment submission/manual-completion state otherwise.
The Modules tracker also merges assignment and quiz submission state, excludes
untracked informational pages from its denominator, and clearly distinguishes
submitted, completed, and remaining work.

- `extension/` contains the browser extension source.
- `ad-kit/` contains the TikTok ad project and recording scripts.
- `oauth-branding/` contains the OAuth homepage + policy pages.
- `.github/workflows/pages.yml` auto-deploys `oauth-branding/` to GitHub Pages.

## Release the extension

1. Increase `version` in `extension/manifest.json`.
2. Test the unpacked extension locally.
3. Commit and push the change to `main`.
4. ZIP the contents of `extension/` so `manifest.json` is at the root of the
   ZIP file. Exclude the older ZIP archives stored in that folder.
5. In the Chrome Developer Dashboard, open QuickCanvas, choose **Package**, and
   upload the new ZIP.

Ad recording:
1. `cd ad-kit`
2. `powershell -ExecutionPolicy Bypass -File .\record-ad.ps1`

GitHub Pages (OAuth Branding):
1. Create a new GitHub repo.
2. Run:
   - `git init`
   - `git add .`
   - `git commit -m "Initial QuickCanvas workspace"`
   - `git branch -M main`
   - `git remote add origin https://github.com/<your-username>/<your-repo>.git`
   - `git push -u origin main`
3. In GitHub repo settings -> Pages:
   - Source: `GitHub Actions`
4. After workflow completes, use:
   - Home: `https://<your-username>.github.io/<your-repo>/index.html`
   - Privacy: `https://<your-username>.github.io/<your-repo>/privacy.html`
   - Terms: `https://<your-username>.github.io/<your-repo>/terms.html`
