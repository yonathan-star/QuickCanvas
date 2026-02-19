# QuickCanvas Workspace

- `extension/` contains the browser extension source.
- `ad-kit/` contains the TikTok ad project and recording scripts.
- `oauth-branding/` contains the OAuth homepage + policy pages.
- `.github/workflows/pages.yml` auto-deploys `oauth-branding/` to GitHub Pages.

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
