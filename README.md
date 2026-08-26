# QuickCanvas Workspace

QuickCanvas connects to Canvas through the user's existing signed-in Canvas
browser session. It does not ask students to create or paste personal Canvas
access tokens. Enter the school's HTTPS Canvas URL in the extension, grant site
access, and sign in to Canvas normally.

Canvas OAuth is still the supported option for a standalone or multi-school
service. It requires an administrator-enabled Developer Key and a trusted
backend for the client secret; never embed that secret in this extension.

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
