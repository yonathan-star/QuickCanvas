const baseUrlInput = document.getElementById("baseUrl");
const enabledInput = document.getElementById("enabled");
const saveBtn = document.getElementById("save");
const clearBtn = document.getElementById("clear");
const statusEl = document.getElementById("status");
const CONTENT_SCRIPT_ID = "cfe-canvas-content";
const CONTENT_SCRIPT_ORIGIN_KEY = "cfeCanvasContentOrigin";

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#b3261e" : "#1f5f8b";
}

function normalizeBaseUrl(url) {
  if (!url) return "";
  const trimmed = url.trim();
  const withProtocol = /^https?:\/\//i.test(trimmed)
    ? trimmed
    : `https://${trimmed}`;
  try {
    const parsed = new URL(withProtocol);
    if (parsed.protocol !== "https:") return "";
    return parsed.origin;
  } catch (error) {
    return "";
  }
}

function baseUrlToMatchPattern(baseUrl) {
  const origin = normalizeBaseUrl(baseUrl);
  return origin ? `${origin}/*` : "";
}

async function unregisterCanvasContentScript() {
  try {
    await chrome.scripting.unregisterContentScripts({
      ids: [CONTENT_SCRIPT_ID],
    });
  } catch (error) {
    // The script may not have been registered yet.
  }
}

async function ensureCanvasSiteAccess(baseUrl) {
  const matchPattern = baseUrlToMatchPattern(baseUrl);
  if (!matchPattern) return false;
  let granted = await chrome.permissions.contains({ origins: [matchPattern] });
  if (!granted) {
    granted = await chrome.permissions.request({ origins: [matchPattern] });
  }
  if (!granted) return false;

  const stored = await chrome.storage.local.get(CONTENT_SCRIPT_ORIGIN_KEY);
  const previousPattern = String(
    stored?.[CONTENT_SCRIPT_ORIGIN_KEY] || "",
  );
  await unregisterCanvasContentScript();
  await chrome.scripting.registerContentScripts([
    {
      id: CONTENT_SCRIPT_ID,
      matches: [matchPattern],
      js: ["content.js"],
      css: ["content.css"],
      runAt: "document_start",
      persistAcrossSessions: true,
    },
  ]);
  if (previousPattern && previousPattern !== matchPattern) {
    await chrome.permissions.remove({ origins: [previousPattern] });
  }
  await chrome.storage.local.set({
    [CONTENT_SCRIPT_ORIGIN_KEY]: matchPattern,
  });
  return true;
}

async function clearCanvasSiteAccess() {
  const stored = await chrome.storage.local.get(CONTENT_SCRIPT_ORIGIN_KEY);
  const previousPattern = String(
    stored?.[CONTENT_SCRIPT_ORIGIN_KEY] || "",
  );
  await unregisterCanvasContentScript();
  if (previousPattern) {
    await chrome.permissions.remove({ origins: [previousPattern] });
  }
  await chrome.storage.local.set({ [CONTENT_SCRIPT_ORIGIN_KEY]: "" });
}

async function loadSettings() {
  const { canvasSettings } = await chrome.storage.sync.get("canvasSettings");
  if (!canvasSettings) return;
  baseUrlInput.value = canvasSettings.baseUrl || "";
  enabledInput.checked = canvasSettings.enabled ?? true;
}

async function saveSettings() {
  const baseUrl = normalizeBaseUrl(baseUrlInput.value);
  const enabled = enabledInput.checked;

  if (!baseUrl) {
    setStatus("Please enter a valid Canvas base URL.", true);
    return;
  }

  const accessGranted = await ensureCanvasSiteAccess(baseUrl);
  if (!accessGranted) {
    setStatus("Canvas site access was not granted.", true);
    return;
  }

  await chrome.storage.sync.set({
    canvasSettings: {
      baseUrl,
      authMode: "canvas_session",
      enabled,
    },
  });

  setStatus("Settings saved. Open or reload Canvas while signed in.");
}

async function clearSettings() {
  await clearCanvasSiteAccess();
  await chrome.storage.sync.remove("canvasSettings");
  baseUrlInput.value = "";
  enabledInput.checked = true;
  setStatus("Settings cleared.");
}

saveBtn.addEventListener("click", saveSettings);
clearBtn.addEventListener("click", clearSettings);
enabledInput.addEventListener("change", saveSettings);

loadSettings();
