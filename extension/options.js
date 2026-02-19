const baseUrlInput = document.getElementById("baseUrl");
const apiTokenInput = document.getElementById("apiToken");
const enabledInput = document.getElementById("enabled");
const saveBtn = document.getElementById("save");
const clearBtn = document.getElementById("clear");
const statusEl = document.getElementById("status");

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

async function loadSettings() {
  const { canvasSettings } = await chrome.storage.sync.get("canvasSettings");
  if (!canvasSettings) return;
  baseUrlInput.value = canvasSettings.baseUrl || "";
  apiTokenInput.value = canvasSettings.apiToken || "";
  enabledInput.checked = canvasSettings.enabled ?? true;
}

async function saveSettings() {
  const baseUrl = normalizeBaseUrl(baseUrlInput.value);
  const apiToken = apiTokenInput.value.trim();
  const enabled = enabledInput.checked;

  if (!baseUrl || !apiToken) {
    setStatus("Please enter a valid base URL and API token.", true);
    return;
  }

  await chrome.storage.sync.set({
    canvasSettings: {
      baseUrl,
      apiToken,
      enabled,
    },
  });

  setStatus("Settings saved.");
}

async function clearSettings() {
  await chrome.storage.sync.remove("canvasSettings");
  baseUrlInput.value = "";
  apiTokenInput.value = "";
  enabledInput.checked = true;
  setStatus("Settings cleared.");
}

saveBtn.addEventListener("click", saveSettings);
clearBtn.addEventListener("click", clearSettings);
enabledInput.addEventListener("change", saveSettings);

loadSettings();
