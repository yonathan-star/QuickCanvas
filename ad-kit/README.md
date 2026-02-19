# QuickCanvas Ad Kit

Drop your ad screenshots into `assets/` with these names:
- `assets/shot1.png`
- `assets/shot2.png`
- `assets/shot3.png`
- `assets/shot4.png`
- `assets/shot5.png`

Quick import from Downloads (latest 5 images):
1. Save the 5 screenshots from chat to your `Downloads` folder.
2. Run:
   `powershell -ExecutionPolicy Bypass -File .\import-shots.ps1`

Run locally:
1. Open `index.html` in your browser.
2. (Optional) Open `captions.html` for ready-to-copy TikTok captions.
3. Record MP4:
   `powershell -ExecutionPolicy Bypass -File .\record-ad.ps1`

Notes:
- The slideshow auto-loads available images and skips any missing files.
- If no screenshots are found, a styled fallback slide appears.
