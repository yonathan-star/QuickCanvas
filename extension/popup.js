const baseUrlInput = document.getElementById("baseUrl");
const apiTokenInput = document.getElementById("apiToken");
const enabledInput = document.getElementById("enabled");
const calendarRemindersEnabledInput = document.getElementById(
  "calendarRemindersEnabled",
);
const saveBtn = document.getElementById("save");
const clearBtn = document.getElementById("clear");
const openCanvasBtn = document.getElementById("openCanvas");
const openOptionsBtn = document.getElementById("openOptions");
const statusEl = document.getElementById("status");
const globalStatusEl = document.getElementById("globalStatus");
const statusPill = document.getElementById("statusPill");

const themeModeSelect = document.getElementById("themeMode");
const accentColorInput = document.getElementById("accentColor");
const bgIntensityInput = document.getElementById("bgIntensity");
const surfaceContrastInput = document.getElementById("surfaceContrast");
const bgColorInput = document.getElementById("bgColor");
const surfaceColorInput = document.getElementById("surfaceColor");
const surfaceAltColorInput = document.getElementById("surfaceAltColor");
const borderColorInput = document.getElementById("borderColor");
const textColorInput = document.getElementById("textColor");
const mutedColorInput = document.getElementById("mutedColor");
const fontBodySelect = document.getElementById("fontBody");
const fontHeadSelect = document.getElementById("fontHead");
const radiusScaleInput = document.getElementById("radiusScale");
const shadowStrengthInput = document.getElementById("shadowStrength");
const themeNameInput = document.getElementById("themeName");
const customCssInput = document.getElementById("customCss");
const saveThemeBtn = document.getElementById("saveTheme");
const saveThemeCloudBtn = document.getElementById("saveThemeCloud");
const publishThemeBtn = document.getElementById("publishTheme");
const presetGrid = document.getElementById("presetGrid");
const communityListEl = document.getElementById("communityList");
const sortTrendingBtn = document.getElementById("sortTrending");
const sortLatestBtn = document.getElementById("sortLatest");

const authEmailInput = document.getElementById("authEmail");
const authPasswordInput = document.getElementById("authPassword");
const authUsernameInput = document.getElementById("authUsername");
const authUsernameField = document.getElementById("authUsernameField");
const authSchoolStartInput = document.getElementById("authSchoolStart");
const authSchoolStartField = document.getElementById("authSchoolStartField");
const authModeSwitch = document.getElementById("authModeSwitch");
const authModeSignInBtn = document.getElementById("authModeSignIn");
const authModeSignUpBtn = document.getElementById("authModeSignUp");
const signInBtn = document.getElementById("signIn");
const signUpBtn = document.getElementById("signUp");
const signOutBtn = document.getElementById("signOut");
const authStatusEl = document.getElementById("authStatus");
const authGateNoticeEl = document.getElementById("authGateNotice");
const cloudListEl = document.getElementById("cloudList");
const refreshCloudBtn = document.getElementById("refreshCloud");
const displayNameInput = document.getElementById("displayName");
const profileSchoolStartInput = document.getElementById("profileSchoolStart");
const profileEmailInput = document.getElementById("profileEmail");
const profileCurrentPasswordInput = document.getElementById(
  "profileCurrentPassword",
);
const profileNewPasswordInput = document.getElementById("profileNewPassword");
const profileNewPasswordConfirmInput = document.getElementById(
  "profileNewPasswordConfirm",
);
const saveProfileBtn = document.getElementById("saveProfile");
const profileStatusEl = document.getElementById("profileStatus");
const accountProfilePanel = document.getElementById("accountProfilePanel");
const accountCloudPanel = document.getElementById("accountCloudPanel");
const tokenHelpPanel = document.getElementById("tokenHelpPanel");

const adminTabBtn = document.getElementById("adminTab");
const reportListEl = document.getElementById("reportList");
const refreshReportsBtn = document.getElementById("refreshReports");
const adminThemesEl = document.getElementById("adminThemes");
const refreshAdminThemesBtn = document.getElementById("refreshAdminThemes");

const tabs = Array.from(document.querySelectorAll(".tab"));
const panes = Array.from(document.querySelectorAll("[data-pane]"));

const SUPABASE_URL = "https://oaecabcqpivoaycnmdwl.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9hZWNhYmNxcGl2b2F5Y25tZHdsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzA3NDk4MTQsImV4cCI6MjA4NjMyNTgxNH0.POx8O2L91qVmm0OVXg74EkzqdIGurSv0zRCNNLntxVs";
const ADMIN_EMAIL = "yonathangal12345@gmail.com";
const REPORT_FUNCTION = "report";
const REPORT_FUNCTION_ALIASES = [
  REPORT_FUNCTION,
  "report-theme",
  "report_theme",
  "reporttheme",
];
const REPORT_REVIEW_STATUS = {
  PENDING: "pending",
  VALID: "valid",
  INVALID: "invalid",
};
const AUTH_GATE_STATE_KEY = "cfeAuthState";
const LOCAL_USERNAMES_KEY = "cfeLocalUsernames";
const FORCE_SIGNED_OUT_KEY = "cfeForceSignedOut";
const AUTH_TOKENS_KEY = "cfeAuthTokens";
const POPUP_ACTIVE_TAB_KEY = "cfePopupActiveTab";
const CALENDAR_REMINDER_KEY = "cfeCalendarRemindersEnabled";
const SCHOOL_START_MINUTES_KEY = "cfeSchoolStartMinutes";
const SCHOOL_START_PENDING_SYNC_KEY = "cfeSchoolStartPendingSync";
const DEFAULT_SCHOOL_START_MINUTES = 8 * 60;
const CONTENT_SCRIPT_ID = "cfe-canvas-content";
const CONTENT_SCRIPT_ORIGIN_KEY = "cfeCanvasContentOrigin";
const POPUP_THEME_MIRROR_KEY = "cfePopupThemeMirror";
const SUPABASE_SINGLETON_KEY = "__cfeSupabaseClient";
const COMMUNITY_FETCH_LIMIT = 60;
const COMMUNITY_RENDER_LIMIT = 30;
const ADMIN_REPORT_LIMIT = 120;
const ADMIN_THEME_LIMIT = 140;
const COMMUNITY_LIST_CACHE_TTL_MS = 45_000;
const ADMIN_REPORT_CACHE_TTL_MS = 20_000;
const ADMIN_THEME_CACHE_TTL_MS = 20_000;
const COMMUNITY_LIST_SELECT =
  "id,user_id,name,username,visible,mode,accent,bg,surface,surface_alt,border,text,muted,bg_intensity,surface_contrast,font_body,font_head,radius,shadow,like_count,updated_at,created_at";
const COMMUNITY_DETAIL_SELECT =
  "id,user_id,name,username,visible,mode,accent,bg,surface,surface_alt,border,text,muted,bg_intensity,surface_contrast,font_body,font_head,radius,shadow,custom_css,like_count,updated_at,created_at";
const ADMIN_THEME_SELECT =
  "id,user_id,name,username,visible,mode,accent,like_count,created_at";

let supabaseClient = null;
if (window.supabase && typeof window.supabase.createClient === "function") {
  if (!window[SUPABASE_SINGLETON_KEY]) {
    window[SUPABASE_SINGLETON_KEY] = window.supabase.createClient(
      SUPABASE_URL,
      SUPABASE_ANON_KEY,
    );
  }
  supabaseClient = window[SUPABASE_SINGLETON_KEY];
}
let currentProfileUsername = "";
let currentSchoolStartMinutes = DEFAULT_SCHOOL_START_MINUTES;
let currentSchoolStartPendingSync = false;
let cloudSchoolStartMinutes = null;
let authMode = "signin";
let forceSignedOutState = false;
let isUiSignedIn = false;
const communityListCache = new Map();
const communityLoadInFlight = new Map();
const communityThemeDetailCache = new Map();
let adminReportCache = null;
let adminReportLoadInFlight = null;
let adminThemeCache = null;
let adminThemeLoadInFlight = null;
let isProfileSaveInProgress = false;
let isSessionBootstrapComplete = false;
let cachedActiveSession = null;
let cachedActiveSessionAt = 0;

const authEmailField = authEmailInput?.closest(".field") || null;
const authPasswordField = authPasswordInput?.closest(".field") || null;

const prefersDark = window.matchMedia("(prefers-color-scheme: dark)");

const PRESETS = [
  {
    name: "Canvas Classic",
    mode: "light",
    accent: "#1b6dff",
    bg: "#f5f6f7",
    surface: "#ffffff",
    surfaceAlt: "#f8fafc",
    border: "#e5e9ef",
    text: "#2d3b45",
    muted: "#6b7780",
    bgIntensity: 30,
    surfaceContrast: 58,
    fontBody: "Space Grotesk",
    fontHead: "Fraunces",
    radius: 10,
    shadow: 25,
  },
  {
    name: "Ocean Neon",
    mode: "dark",
    accent: "#36d7ff",
    bg: "#08111f",
    surface: "#0f1b2d",
    surfaceAlt: "#17263d",
    border: "#29405f",
    text: "#eef6ff",
    muted: "#a9bdd6",
    bgIntensity: 30,
    surfaceContrast: 66,
    fontBody: "Sora",
    fontHead: "Bebas Neue",
    radius: 10,
    shadow: 28,
  },
  {
    name: "Playful Pop",
    mode: "light",
    accent: "#ef3f7f",
    bg: "#fff7fb",
    surface: "#ffffff",
    surfaceAlt: "#ffeef7",
    border: "#edcade",
    text: "#2d1f2a",
    muted: "#725f6d",
    bgIntensity: 26,
    surfaceContrast: 58,
    fontBody: "Nunito",
    fontHead: "Fredoka",
    radius: 12,
    shadow: 22,
  },
  {
    name: "Sunset Paper",
    mode: "light",
    accent: "#f97316",
    bg: "#fff7ed",
    surface: "#fffdf9",
    surfaceAlt: "#ffedd5",
    border: "#fed7aa",
    text: "#3b2a1e",
    muted: "#8a6a52",
    bgIntensity: 34,
    surfaceContrast: 56,
    fontBody: "DM Sans",
    fontHead: "Cinzel",
    radius: 11,
    shadow: 24,
  },
  {
    name: "Sage Studio",
    mode: "light",
    accent: "#1b8a5a",
    bg: "#f3f7f3",
    surface: "#ffffff",
    surfaceAlt: "#f1f7f3",
    border: "#dfe9e1",
    text: "#1f2a1f",
    muted: "#6b7b6f",
    bgIntensity: 24,
    surfaceContrast: 54,
    fontBody: "Plus Jakarta Sans",
    fontHead: "Alegreya",
    radius: 10,
    shadow: 20,
  },
  {
    name: "Ink Mono",
    mode: "dark",
    accent: "#93c5fd",
    bg: "#0b0f14",
    surface: "#131923",
    surfaceAlt: "#1b2431",
    border: "#2b3646",
    text: "#e5e7eb",
    muted: "#9ca3af",
    bgIntensity: 18,
    surfaceContrast: 68,
    fontBody: "IBM Plex Sans",
    fontHead: "IBM Plex Serif",
    radius: 8,
    shadow: 18,
  },
  {
    name: "Royal Plum",
    mode: "dark",
    accent: "#b88cff",
    bg: "#130d1f",
    surface: "#1c1430",
    surfaceAlt: "#281d42",
    border: "#3f3060",
    text: "#f5efff",
    muted: "#c5b5e3",
    bgIntensity: 28,
    surfaceContrast: 65,
    fontBody: "Outfit",
    fontHead: "Cinzel",
    radius: 10,
    shadow: 26,
  },
  {
    name: "Mint Arcade",
    mode: "light",
    accent: "#0ea67c",
    bg: "#f3fff9",
    surface: "#ffffff",
    surfaceAlt: "#eafbf3",
    border: "#cfeee0",
    text: "#173328",
    muted: "#567968",
    bgIntensity: 24,
    surfaceContrast: 57,
    fontBody: "Quicksand",
    fontHead: "Righteous",
    radius: 11,
    shadow: 20,
  },
  {
    name: "Cherry Cola",
    mode: "dark",
    accent: "#ff6b86",
    bg: "#1a0d12",
    surface: "#241219",
    surfaceAlt: "#311824",
    border: "#4a2834",
    text: "#ffeef2",
    muted: "#d0a6b2",
    bgIntensity: 30,
    surfaceContrast: 63,
    fontBody: "DM Sans",
    fontHead: "Abril Fatface",
    radius: 10,
    shadow: 24,
  },
  {
    name: "Lemonade",
    mode: "light",
    accent: "#d98b00",
    bg: "#fffced",
    surface: "#ffffff",
    surfaceAlt: "#fff6d9",
    border: "#f2df9d",
    text: "#3c2d10",
    muted: "#7b673f",
    bgIntensity: 20,
    surfaceContrast: 60,
    fontBody: "Lexend",
    fontHead: "Baloo 2",
    radius: 10,
    shadow: 18,
  },
  {
    name: "Monochrome Pro",
    mode: "dark",
    accent: "#7dd3fc",
    bg: "#0b0c0f",
    surface: "#14161c",
    surfaceAlt: "#1c2029",
    border: "#2f3744",
    text: "#f3f4f6",
    muted: "#a5adba",
    bgIntensity: 16,
    surfaceContrast: 70,
    fontBody: "Plus Jakarta Sans",
    fontHead: "Oswald",
    radius: 8,
    shadow: 16,
  },
  {
    name: "Sky Paper",
    mode: "light",
    accent: "#2a76ff",
    bg: "#f4f8ff",
    surface: "#ffffff",
    surfaceAlt: "#ebf2ff",
    border: "#d2e1ff",
    text: "#1b2a45",
    muted: "#5f7298",
    bgIntensity: 22,
    surfaceContrast: 59,
    fontBody: "Figtree",
    fontHead: "Playfair Display",
    radius: 10,
    shadow: 20,
  },
];

const FONT_OPTIONS = [
  { label: "Space Grotesk", value: "Space Grotesk" },
  { label: "Manrope", value: "Manrope" },
  { label: "Sora", value: "Sora" },
  { label: "Outfit", value: "Outfit" },
  { label: "Urbanist", value: "Urbanist" },
  { label: "Plus Jakarta Sans", value: "Plus Jakarta Sans" },
  { label: "Lexend", value: "Lexend" },
  { label: "Rubik", value: "Rubik" },
  { label: "DM Sans", value: "DM Sans" },
  { label: "Figtree", value: "Figtree" },
  { label: "Nunito", value: "Nunito" },
  { label: "Quicksand", value: "Quicksand" },
  { label: "Fredoka", value: "Fredoka" },
  { label: "Comfortaa", value: "Comfortaa" },
  { label: "Baloo 2", value: "Baloo 2" },
  { label: "Poppins", value: "Poppins" },
  { label: "Montserrat", value: "Montserrat" },
  { label: "Oswald", value: "Oswald" },
  { label: "Anton", value: "Anton" },
  { label: "Work Sans", value: "Work Sans" },
  { label: "IBM Plex Sans", value: "IBM Plex Sans" },
  { label: "Fraunces", value: "Fraunces" },
  { label: "Bebas Neue", value: "Bebas Neue" },
  { label: "Righteous", value: "Righteous" },
  { label: "Lobster", value: "Lobster" },
  { label: "Pacifico", value: "Pacifico" },
  { label: "Bangers", value: "Bangers" },
  { label: "Permanent Marker", value: "Permanent Marker" },
  { label: "Caveat", value: "Caveat" },
  { label: "Amatic SC", value: "Amatic SC" },
  { label: "Patrick Hand", value: "Patrick Hand" },
  { label: "Cinzel", value: "Cinzel" },
  { label: "Cormorant Garamond", value: "Cormorant Garamond" },
  { label: "Abril Fatface", value: "Abril Fatface" },
  { label: "Alegreya", value: "Alegreya" },
  { label: "Playfair Display", value: "Playfair Display" },
  { label: "Merriweather", value: "Merriweather" },
  { label: "IBM Plex Serif", value: "IBM Plex Serif" },
];

const PROFANITY = [
  "ass",
  "bastard",
  "bitch",
  "cock",
  "cunt",
  "damn",
  "dick",
  "fuck",
  "motherfucker",
  "nigger",
  "nigga",
  "piss",
  "porn",
  "pussy",
  "shit",
  "slut",
  "whore",
  "rape",
  "rapist",
  "fag",
  "faggot",
  "kike",
  "spic",
  "retard",
  "retarded",
  "pedo",
];

const LEET_MAP = {
  a: "4@",
  b: "8",
  e: "3",
  g: "69",
  i: "1!|",
  l: "1|",
  o: "0",
  s: "5$",
  t: "7+",
  z: "2",
};

function normalizeText(text) {
  if (!text) return "";
  let cleaned = text.toLowerCase();
  cleaned = cleaned.replace(/[\s\-_.]/g, "");
  cleaned = cleaned.replace(/[@]/g, "a");
  cleaned = cleaned.replace(/[0]/g, "o");
  cleaned = cleaned.replace(/[1!|]/g, "i");
  cleaned = cleaned.replace(/[3]/g, "e");
  cleaned = cleaned.replace(/[5$]/g, "s");
  cleaned = cleaned.replace(/[7+]/g, "t");
  cleaned = cleaned.replace(/[8]/g, "b");
  cleaned = cleaned.replace(/(.)\1{2,}/g, "$1$1");
  return cleaned;
}

function containsProfanity(text) {
  if (!text) return false;
  const lower = text.toLowerCase();
  const normalized = normalizeText(text);
  const tokens = lower.split(/\s+/);
  return PROFANITY.some((word) => {
    const wordRegex = new RegExp(`\\b${word}\\b`, "i");
    if (wordRegex.test(lower)) return true;
    if (normalized.includes(word)) return true;
    const leetPattern = word
      .split("")
      .map((char) => {
        const map = LEET_MAP[char] || "";
        return map ? `[${char}${map}]` : char;
      })
      .join("\\W*");
    try {
      const leetRegex = new RegExp(leetPattern, "i");
      return leetRegex.test(lower);
    } catch (error) {
      return tokens.includes(word);
    }
  });
}

function setStatus(message, isError = false) {
  if (statusEl) {
    statusEl.textContent = message;
    statusEl.classList.toggle("error", isError);
  }
  if (globalStatusEl) {
    globalStatusEl.textContent = message;
    globalStatusEl.classList.toggle("error", isError);
  }
}

function setAuthStatus(message, isError = false) {
  if (!authStatusEl) {
    setStatus(message, isError);
    return;
  }
  authStatusEl.textContent = message;
  authStatusEl.classList.toggle("error", isError);
  if (statusEl) {
    setStatus(message, isError);
  }
  if (authGateNoticeEl && message) {
    authGateNoticeEl.textContent = message;
    authGateNoticeEl.classList.toggle("error", isError);
  }
  if (statusPill && message) {
    statusPill.textContent = isError ? "Auth error" : "Connected";
    statusPill.classList.toggle("connected", !isError);
  }
  if (authStatusEl && message) {
    authStatusEl.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }
}

function setProfileStatus(message, isError = false) {
  if (!profileStatusEl) return;
  profileStatusEl.textContent = message;
  profileStatusEl.classList.toggle("error", isError);
}

function getStoredActiveTab() {
  try {
    const stored = localStorage.getItem(POPUP_ACTIVE_TAB_KEY);
    return stored ? String(stored) : "";
  } catch (error) {
    return "";
  }
}

function errorMessage(error, fallback = "Unexpected error.") {
  if (!error) return fallback;
  if (typeof error === "string") return error;
  if (typeof error.message === "string" && error.message.trim()) {
    return error.message.trim();
  }
  try {
    return JSON.stringify(error);
  } catch (serializeError) {
    return fallback;
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeAttr(value) {
  return escapeHtml(value);
}

function withUiError(label, handler) {
  return async (...args) => {
    try {
      await handler(...args);
    } catch (error) {
      const msg = `${label}: ${errorMessage(error)}`;
      console.error("[QuickCanvas]", msg, error);
      setAuthStatus(msg, true);
    }
  };
}

function withTimeout(promise, timeoutMs = 8000, label = "operation") {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      reject(new Error(`${label} timed out after ${timeoutMs}ms`));
    }, timeoutMs);
    Promise.resolve(promise)
      .then((value) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        resolve(value);
      })
      .catch((error) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(error);
      });
  });
}

function usernameFromEmail(email) {
  const localPart = String(email || "").split("@")[0] || "";
  const normalized = localPart
    .toLowerCase()
    .replace(/[^a-z0-9_]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
  if (normalized.length >= 3 && normalized.length <= 20) return normalized;
  if (normalized.length > 20) return normalized.slice(0, 20);
  if (normalized.length > 0) return `${normalized}001`.slice(0, 20);
  return "";
}

function clampSchoolStartMinutes(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return DEFAULT_SCHOOL_START_MINUTES;
  const rounded = Math.round(numeric);
  if (rounded < 0 || rounded > 23 * 60 + 59) {
    return DEFAULT_SCHOOL_START_MINUTES;
  }
  return rounded;
}

function parseSchoolStartInputValue(value) {
  const match = String(value || "")
    .trim()
    .match(/^(\d{1,2}):(\d{2})(?::\d{2}(?:\.\d+)?)?\s*([AaPp][Mm])?$/);
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const meridiem = String(match[3] || "").toLowerCase();
  if (meridiem) {
    if (!Number.isInteger(hour) || hour < 1 || hour > 12) return null;
    if (meridiem === "am") {
      if (hour === 12) hour = 0;
    } else if (meridiem === "pm") {
      if (hour !== 12) hour += 12;
    } else {
      return null;
    }
  }
  if (
    !Number.isInteger(hour) ||
    !Number.isInteger(minute) ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59
  ) {
    return null;
  }
  return hour * 60 + minute;
}

function parseDbSchoolStartValue(value) {
  const match = String(value || "")
    .trim()
    .match(
      /^(\d{1,2}):(\d{2})(?::\d{2}(?:\.\d+)?)?(?:\s?(?:Z|[+-]\d{2}(?::?\d{2})?))?$/i,
    );
  if (!match) return null;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (
    !Number.isInteger(hour) ||
    !Number.isInteger(minute) ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59
  ) {
    return null;
  }
  return hour * 60 + minute;
}

function formatSchoolStartForInput(minutes) {
  const safe = clampSchoolStartMinutes(minutes);
  const hour = Math.floor(safe / 60);
  const minute = safe % 60;
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}

function formatSchoolStartForDb(minutes) {
  return `${formatSchoolStartForInput(minutes)}:00`;
}

function hasSchoolStartMinutesValue(value) {
  if (value === null || value === undefined) return false;
  if (typeof value === "string" && value.trim() === "") return false;
  const numeric = Number(value);
  return Number.isFinite(numeric);
}

function getMetadataSchoolStartMinutes(session) {
  const minutesRaw = session?.user?.user_metadata?.school_start_minutes;
  if (hasSchoolStartMinutesValue(minutesRaw)) {
    return clampSchoolStartMinutes(minutesRaw);
  }
  const direct = parseDbSchoolStartValue(
    session?.user?.user_metadata?.school_start_time,
  );
  if (direct !== null) return direct;
  const hhmm = parseSchoolStartInputValue(
    session?.user?.user_metadata?.school_start_hhmm,
  );
  if (hhmm !== null) return hhmm;
  const legacyCamel = parseDbSchoolStartValue(
    session?.user?.user_metadata?.schoolStartTime,
  );
  if (legacyCamel !== null) return legacyCamel;
  return parseSchoolStartInputValue(
    session?.user?.user_metadata?.schoolStartHhmm,
  );
}

async function loadStoredSchoolStartMinutes() {
  try {
    const stored = await chrome.storage.sync.get([
      SCHOOL_START_MINUTES_KEY,
      SCHOOL_START_PENDING_SYNC_KEY,
    ]);
    const minutes = clampSchoolStartMinutes(stored?.[SCHOOL_START_MINUTES_KEY]);
    currentSchoolStartMinutes = minutes;
    currentSchoolStartPendingSync = Boolean(
      stored?.[SCHOOL_START_PENDING_SYNC_KEY],
    );
  } catch (error) {
    currentSchoolStartMinutes = DEFAULT_SCHOOL_START_MINUTES;
    currentSchoolStartPendingSync = false;
  }
  if (authSchoolStartInput) {
    authSchoolStartInput.value = formatSchoolStartForInput(
      currentSchoolStartMinutes,
    );
  }
  if (profileSchoolStartInput) {
    profileSchoolStartInput.value = formatSchoolStartForInput(
      currentSchoolStartMinutes,
    );
  }
  return currentSchoolStartMinutes;
}

async function saveSchoolStartMinutes(
  minutes,
  { markPending = true, persist = true } = {},
) {
  const safe = clampSchoolStartMinutes(minutes);
  currentSchoolStartMinutes = safe;
  if (markPending === true || markPending === false) {
    currentSchoolStartPendingSync = Boolean(markPending);
  }
  if (authSchoolStartInput) {
    authSchoolStartInput.value = formatSchoolStartForInput(safe);
  }
  if (profileSchoolStartInput) {
    profileSchoolStartInput.value = formatSchoolStartForInput(safe);
  }
  if (persist) {
    const payload = { [SCHOOL_START_MINUTES_KEY]: safe };
    if (markPending === true || markPending === false) {
      payload[SCHOOL_START_PENDING_SYNC_KEY] = currentSchoolStartPendingSync;
    }
    try {
      await chrome.storage.sync.set(payload);
    } catch (error) {
      // ignore storage write errors in popup lifecycle
    }
  }
  return safe;
}

async function persistSchoolStartPreference(
  minutes,
  { showStatus = false } = {},
) {
  const safe = await saveSchoolStartMinutes(minutes);
  const session = await getActiveSession();
  if (!session?.user?.id) {
    if (showStatus) {
      setProfileStatus("School start time saved locally.");
      setStatus("School start time saved locally.");
    }
    return { ok: true, localOnly: true };
  }
  const usernameForSync = getFirstValidUsername(
    currentProfileUsername,
    getMetadataUsername(session),
    await getLocalUsername(session.user.id),
    displayNameInput?.value,
    authUsernameInput?.value,
    usernameFromEmail(session.user?.email || ""),
  );
  const profileResult = await persistProfileSettings(session, {
    username: usernameForSync,
    schoolStartMinutes: safe,
  });
  if (!profileResult.ok) {
    const metadataResult = await syncAuthProfileMetadata({
      username: usernameForSync,
      schoolStartMinutes: safe,
    });
    if (showStatus) {
      if (metadataResult.ok) {
        const warnText = `School start synced to auth metadata, but profile table sync failed: ${profileResult.error || "Unknown error."}`;
        setProfileStatus(warnText, true);
        setStatus(warnText, true);
      } else {
        const errorText = `School start saved locally. Cloud sync failed: ${profileResult.error || "Unknown error."}`;
        setProfileStatus(errorText, true);
        setStatus(errorText, true);
      }
    }
    if (metadataResult.ok) {
      return {
        ok: true,
        localOnly: false,
        profileTableOk: false,
        metadataOk: true,
        error: profileResult.error || "Cloud profile sync failed.",
      };
    }
    return {
      ok: false,
      localOnly: true,
      error: profileResult.error || "Cloud profile sync failed.",
    };
  }
  const metadataResult = await syncAuthProfileMetadata({
    username: usernameForSync,
    schoolStartMinutes: safe,
  });
  await saveSchoolStartMinutes(safe, { markPending: false });
  if (showStatus) {
    const cloudMinutes = hasSchoolStartMinutesValue(cloudSchoolStartMinutes)
      ? clampSchoolStartMinutes(cloudSchoolStartMinutes)
      : safe;
    const cloudLabel = formatSchoolStartForInput(cloudMinutes);
    if (metadataResult.ok || !metadataResult.error) {
      const successText = `School start saved to Supabase (${cloudLabel}).`;
      setProfileStatus(successText);
      setStatus(successText);
    } else {
      const warnText = `School start saved to Supabase (${cloudLabel}). Auth metadata sync failed: ${metadataResult.error}`;
      setProfileStatus(warnText);
      setStatus(warnText);
    }
  }
  return {
    ok: true,
    localOnly: false,
    metadataOk: metadataResult.ok,
    metadataError: metadataResult.error || "",
  };
}

function setAuthMode(nextMode) {
  authMode = nextMode === "signup" ? "signup" : "signin";
  const signUpMode = authMode === "signup";
  if (authModeSignInBtn) {
    authModeSignInBtn.classList.toggle("is-active", !signUpMode);
  }
  if (authModeSignUpBtn) {
    authModeSignUpBtn.classList.toggle("is-active", signUpMode);
  }
  if (authUsernameField) {
    authUsernameField.hidden = !signUpMode;
  }
  if (authSchoolStartField) {
    authSchoolStartField.hidden = !signUpMode;
  }
  if (authSchoolStartInput && signUpMode) {
    authSchoolStartInput.value = formatSchoolStartForInput(
      currentSchoolStartMinutes,
    );
  }
  if (signInBtn) {
    signInBtn.hidden = signUpMode;
  }
  if (signUpBtn) {
    signUpBtn.hidden = !signUpMode;
  }
}

async function setAuthGateState({
  authenticated = false,
  hasUsername = false,
  userId = "",
} = {}) {
  const safeUserId = userId ? String(userId) : "";
  try {
    await chrome.storage.local.set({
      [AUTH_GATE_STATE_KEY]: {
        authenticated: Boolean(authenticated),
        hasUsername: Boolean(hasUsername),
        userId: safeUserId,
        updatedAt: Date.now(),
      },
    });
    try {
      await chrome.storage.sync.set({
        cfeAuthGateMirror: {
          authenticated: Boolean(authenticated),
          userId: safeUserId,
          updatedAt: Date.now(),
        },
      });
    } catch (error) {
      // ignore mirror write failures
    }
  } catch (error) {
    // ignore storage write errors in popup lifecycle
  }
}

async function isUsernameTaken(username) {
  if (!supabaseClient || !username) return false;
  const normalized = username.trim();
  const { data, error } = await supabaseClient
    .from("cfe_profiles")
    .select("username")
    .ilike("username", normalized)
    .limit(1);
  if (error) return false;
  return Array.isArray(data) && data.length > 0;
}

async function isUsernameTakenByOtherUser(username, currentUserId) {
  if (!supabaseClient || !username) return false;
  const normalized = username.trim();
  const { data, error } = await executeWithAuthRetry(() =>
    supabaseClient
      .from("cfe_profiles")
      .select("user_id, username")
      .ilike("username", normalized)
      .limit(3),
  );
  if (error || !Array.isArray(data)) return false;
  return data.some((row) => {
    const rowUserId = String(row?.user_id || "");
    return rowUserId && rowUserId !== String(currentUserId || "");
  });
}

function isValidEmailAddress(email) {
  const value = String(email || "").trim();
  if (!value) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

async function verifyCurrentPasswordForSensitiveChange(session, password) {
  const email = String(session?.user?.email || "").trim();
  const candidate = String(password || "");
  if (!email || !candidate) {
    return { ok: false, error: "Current password is required." };
  }
  try {
    const response = await fetch(
      `${SUPABASE_URL}/auth/v1/token?grant_type=password`,
      {
        method: "POST",
        headers: {
          apikey: SUPABASE_ANON_KEY,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          email,
          password: candidate,
        }),
      },
    );
    if (response.ok) {
      return { ok: true, error: "" };
    }
    const raw = await response.text().catch(() => "");
    let detail = "";
    try {
      const parsed = JSON.parse(raw);
      detail = String(parsed?.msg || parsed?.error_description || "").trim();
    } catch (error) {
      detail = "";
    }
    return {
      ok: false,
      error: detail || "Current password verification failed.",
    };
  } catch (error) {
    return {
      ok: false,
      error: errorMessage(error, "Current password verification failed."),
    };
  }
}

function getMetadataUsername(session) {
  return String(session?.user?.user_metadata?.username || "").trim();
}

function getFirstValidUsername(...values) {
  for (const value of values) {
    const candidate = String(value || "").trim();
    if (!candidate) continue;
    if (!validateUsername(candidate)) return candidate;
  }
  return "";
}

function getRestErrorMessage(
  error,
  fallback = "Supabase REST request failed.",
) {
  if (!error) return fallback;
  return String(error.message || fallback).trim() || fallback;
}

async function restSyncProfileSettings(
  session,
  { username = "", schoolStartMinutes = null } = {},
) {
  const userId = String(session?.user?.id || "").trim();
  if (!userId) {
    return { ok: false, error: "Missing profile context." };
  }
  const restAuth = await getRestAuthContext(session);
  const accessToken = String(restAuth?.accessToken || "");
  if (!accessToken) {
    return { ok: false, error: "Missing auth token for REST profile sync." };
  }
  const trimmedUsername = String(username || "").trim();
  const hasSchoolStart = hasSchoolStartMinutesValue(schoolStartMinutes);
  const expectedSchoolStartMinutes = hasSchoolStart
    ? clampSchoolStartMinutes(schoolStartMinutes)
    : null;
  const expectedSchoolStartDb = hasSchoolStart
    ? formatSchoolStartForDb(expectedSchoolStartMinutes)
    : "";
  const mutationPayload = {};
  if (trimmedUsername) {
    mutationPayload.username = trimmedUsername;
  }
  if (hasSchoolStart) {
    mutationPayload.school_start_time = expectedSchoolStartDb;
  }
  if (Object.keys(mutationPayload).length === 0) {
    return { ok: false, error: "No profile changes to persist." };
  }

  let patchResult = await restTableRequest("cfe_profiles", {
    method: "PATCH",
    params: { user_id: `eq.${userId}` },
    body: mutationPayload,
    accessToken,
    prefer: "return=representation",
    timeoutMs: 10000,
  });
  if (patchResult.error) {
    patchResult = await restTableRequest("cfe_profiles", {
      method: "POST",
      params: { on_conflict: "user_id" },
      body: { user_id: userId, ...mutationPayload },
      accessToken,
      prefer: "resolution=merge-duplicates,return=representation",
      timeoutMs: 10000,
    });
  } else if (Array.isArray(patchResult.data) && patchResult.data.length === 0) {
    patchResult = await restTableRequest("cfe_profiles", {
      method: "POST",
      params: { on_conflict: "user_id" },
      body: { user_id: userId, ...mutationPayload },
      accessToken,
      prefer: "resolution=merge-duplicates,return=representation",
      timeoutMs: 10000,
    });
  }
  if (patchResult.error) {
    return {
      ok: false,
      error: getRestErrorMessage(
        patchResult.error,
        "REST profile sync failed.",
      ),
    };
  }

  if (!hasSchoolStart) {
    return { ok: true, error: "" };
  }
  const verifyResult = await restTableRequest("cfe_profiles", {
    method: "GET",
    params: { user_id: `eq.${userId}`, select: "school_start_time", limit: 1 },
    accessToken,
    timeoutMs: 10000,
  });
  if (verifyResult.error) {
    return {
      ok: false,
      error: getRestErrorMessage(
        verifyResult.error,
        "REST profile verification failed.",
      ),
    };
  }
  const rows = Array.isArray(verifyResult.data) ? verifyResult.data : [];
  const row = rows[0] || null;
  const storedMinutes = parseDbSchoolStartValue(row?.school_start_time);
  if (storedMinutes !== expectedSchoolStartMinutes) {
    const actualText =
      storedMinutes === null
        ? "unknown"
        : formatSchoolStartForInput(storedMinutes);
    return {
      ok: false,
      error: `REST profile verification mismatch (expected ${formatSchoolStartForInput(expectedSchoolStartMinutes)}, found ${actualText}).`,
    };
  }
  return { ok: true, error: "", storedMinutes };
}

async function persistProfileSettings(
  session,
  { username = "", schoolStartMinutes = null } = {},
) {
  if (!session?.user?.id) {
    return { ok: false, error: "Missing profile context." };
  }
  const payload = { user_id: session.user.id };
  const trimmedUsername = String(username || "").trim();
  if (trimmedUsername) {
    payload.username = trimmedUsername;
  }
  const hasSchoolStart = hasSchoolStartMinutesValue(schoolStartMinutes);
  const expectedSchoolStartMinutes = hasSchoolStart
    ? clampSchoolStartMinutes(schoolStartMinutes)
    : null;
  const expectedSchoolStartDb = hasSchoolStart
    ? formatSchoolStartForDb(expectedSchoolStartMinutes)
    : "";
  if (hasSchoolStart) {
    payload.school_start_time = expectedSchoolStartDb;
  }
  if (Object.keys(payload).length <= 1) {
    return { ok: false, error: "No profile changes to persist." };
  }
  const upsertResult = await resilientUpsert("cfe_profiles", payload, {
    onConflict: "user_id",
  });
  if (upsertResult.error) {
    const restFallback = await restSyncProfileSettings(session, {
      username: trimmedUsername,
      schoolStartMinutes: hasSchoolStart ? expectedSchoolStartMinutes : null,
    });
    if (restFallback.ok) {
      if (hasSchoolStart) {
        cloudSchoolStartMinutes = expectedSchoolStartMinutes;
      }
      return { ok: true, error: "" };
    }
    return {
      ok: false,
      error:
        upsertResult.error?.message ||
        restFallback.error ||
        "Cloud profile sync failed.",
    };
  }
  if (!hasSchoolStart) {
    return { ok: true, error: "" };
  }
  // Fast path: if upsert succeeds, trust the write and avoid extra verify reads.
  cloudSchoolStartMinutes = expectedSchoolStartMinutes;
  return { ok: true, error: "" };
}

async function persistProfileUsername(session, username) {
  return persistProfileSettings(session, { username });
}

async function syncAuthProfileMetadata({
  username = "",
  schoolStartMinutes = null,
} = {}) {
  if (!supabaseClient) {
    return { ok: false, error: "Missing auth client." };
  }
  const metadataPatch = {};
  const trimmedUsername = String(username || "").trim();
  if (trimmedUsername) {
    metadataPatch.username = trimmedUsername;
  }
  if (hasSchoolStartMinutesValue(schoolStartMinutes)) {
    const safeMinutes = clampSchoolStartMinutes(schoolStartMinutes);
    metadataPatch.school_start_time = formatSchoolStartForDb(safeMinutes);
    metadataPatch.school_start_hhmm = formatSchoolStartForInput(safeMinutes);
    metadataPatch.school_start_minutes = safeMinutes;
  }
  if (Object.keys(metadataPatch).length === 0) {
    return { ok: false, error: "No metadata changes to persist." };
  }
  const { error } = await withTimeout(
    executeWithAuthRetry(() =>
      supabaseClient.auth.updateUser({ data: metadataPatch }),
    ),
    7000,
    "auth metadata update",
  ).catch((timeoutError) => ({
    data: null,
    error: {
      message: errorMessage(timeoutError, "Auth metadata update timed out."),
      details: "",
      hint: "",
      code: "QUERY_TIMEOUT",
    },
  }));
  if (!error) {
    return { ok: true, error: "" };
  }

  const session = await getActiveSession();
  const restAuth = await getRestAuthContext(session);
  const accessToken = String(
    restAuth?.accessToken || session?.access_token || "",
  );
  if (!accessToken) {
    return { ok: false, error: error?.message || "Missing auth token." };
  }
  try {
    const response = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
      method: "PUT",
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ data: metadataPatch }),
    });
    const raw = await response.text().catch(() => "");
    let parsed = null;
    try {
      parsed = raw ? JSON.parse(raw) : null;
    } catch (parseError) {
      parsed = null;
    }
    if (!response.ok) {
      const fallbackMessage = error?.message || `HTTP ${response.status}`;
      const nextError =
        String(
          parsed?.error || parsed?.message || raw || fallbackMessage,
        ).trim() || fallbackMessage;
      return { ok: false, error: nextError };
    }
    return { ok: true, error: "" };
  } catch (fetchError) {
    return {
      ok: false,
      error:
        error?.message ||
        errorMessage(fetchError, "Auth metadata sync failed."),
    };
  }
}

function clearSupabaseStoredSession() {
  try {
    const keys = Object.keys(localStorage);
    keys.forEach((key) => {
      if (key.includes("-auth-token") && key.startsWith("sb-")) {
        localStorage.removeItem(key);
      }
    });
  } catch (error) {
    // ignore storage cleanup failures
  }
}

async function getLocalUsername(userId) {
  if (!userId) return "";
  try {
    const stored = await chrome.storage.local.get(LOCAL_USERNAMES_KEY);
    const map =
      stored?.[LOCAL_USERNAMES_KEY] &&
      typeof stored[LOCAL_USERNAMES_KEY] === "object"
        ? stored[LOCAL_USERNAMES_KEY]
        : {};
    return String(map[userId] || "").trim();
  } catch (error) {
    return "";
  }
}

async function setLocalUsername(userId, username) {
  if (!userId || !username) return;
  try {
    const stored = await chrome.storage.local.get(LOCAL_USERNAMES_KEY);
    const map =
      stored?.[LOCAL_USERNAMES_KEY] &&
      typeof stored[LOCAL_USERNAMES_KEY] === "object"
        ? { ...stored[LOCAL_USERNAMES_KEY] }
        : {};
    map[userId] = username;
    await chrome.storage.local.set({ [LOCAL_USERNAMES_KEY]: map });
  } catch (error) {
    // ignore local persistence issues
  }
}

async function clearLocalUsername(userId) {
  if (!userId) return;
  try {
    const stored = await chrome.storage.local.get(LOCAL_USERNAMES_KEY);
    const map =
      stored?.[LOCAL_USERNAMES_KEY] &&
      typeof stored[LOCAL_USERNAMES_KEY] === "object"
        ? { ...stored[LOCAL_USERNAMES_KEY] }
        : {};
    if (Object.prototype.hasOwnProperty.call(map, userId)) {
      delete map[userId];
      await chrome.storage.local.set({ [LOCAL_USERNAMES_KEY]: map });
    }
  } catch (error) {
    // ignore local persistence issues
  }
}

async function persistAuthTokens(session) {
  if (!session?.access_token || !session?.refresh_token) return;
  cachedActiveSession = session;
  cachedActiveSessionAt = Date.now();
  try {
    await chrome.storage.local.set({
      [AUTH_TOKENS_KEY]: {
        access_token: session.access_token,
        refresh_token: session.refresh_token,
        expires_at: session.expires_at || null,
        user_id: session.user?.id || "",
        email: session.user?.email || "",
        updatedAt: Date.now(),
      },
    });
  } catch (error) {
    // ignore token persistence errors
  }
}

async function clearAuthTokens() {
  cachedActiveSession = null;
  cachedActiveSessionAt = 0;
  try {
    await chrome.storage.local.remove(AUTH_TOKENS_KEY);
  } catch (error) {
    // ignore token cleanup errors
  }
}

async function getActiveSession() {
  if (!supabaseClient) return null;
  await loadForceSignedOutState();
  if (
    !forceSignedOutState &&
    cachedActiveSession?.access_token &&
    cachedActiveSession?.user?.id &&
    Date.now() - cachedActiveSessionAt < 120000
  ) {
    return cachedActiveSession;
  }
  let directSession = null;
  try {
    const { data } = await withTimeout(
      supabaseClient.auth.getSession(),
      1800,
      "auth.getSession",
    );
    directSession = await enforceForcedSignOut(data?.session || null);
  } catch (error) {
    directSession = null;
  }
  if (directSession) {
    await persistAuthTokens(directSession);
    return directSession;
  }
  if (forceSignedOutState) return null;

  try {
    const stored = await chrome.storage.local.get(AUTH_TOKENS_KEY);
    const tokens = stored?.[AUTH_TOKENS_KEY] || null;
    if (!tokens?.access_token || !tokens?.refresh_token) return null;
    try {
      const { data: restored, error } = await withTimeout(
        supabaseClient.auth.setSession({
          access_token: tokens.access_token,
          refresh_token: tokens.refresh_token,
        }),
        2200,
        "auth.setSession",
      );
      if (!error && restored?.session) {
        const restoredSession = await enforceForcedSignOut(restored.session);
        if (restoredSession) {
          await persistAuthTokens(restoredSession);
          return restoredSession;
        }
      }
    } catch (error) {
      // fall through to token-only fallback
    }
    const tokenOnlySession = await enforceForcedSignOut({
      access_token: String(tokens.access_token || ""),
      refresh_token: String(tokens.refresh_token || ""),
      expires_at: Number(tokens.expires_at || 0) || null,
      user: {
        id: String(tokens.user_id || ""),
        email: String(tokens.email || ""),
      },
    });
    if (tokenOnlySession?.access_token && tokenOnlySession?.user?.id) {
      return tokenOnlySession;
    }
    await clearAuthTokens();
    return null;
  } catch (error) {
    return null;
  }
}

async function getRestAuthContext(session = null) {
  let accessToken = session?.access_token || "";
  let userId = session?.user?.id || "";
  let email = session?.user?.email || "";
  try {
    const stored = await chrome.storage.local.get(AUTH_TOKENS_KEY);
    const tokens = stored?.[AUTH_TOKENS_KEY] || null;
    if (!accessToken) accessToken = String(tokens?.access_token || "");
    if (!userId) userId = String(tokens?.user_id || "");
    if (!email) email = String(tokens?.email || "");
  } catch (error) {
    // ignore local token read failures
  }
  return { accessToken, userId, email };
}

async function getFastAuthContext(sessionTimeoutMs = 3500) {
  let session = null;
  try {
    session = await withTimeout(
      getActiveSession(),
      sessionTimeoutMs,
      "session lookup",
    );
  } catch (error) {
    session = null;
  }
  const restAuth = await getRestAuthContext(session);
  const actorUserId = session?.user?.id || restAuth.userId || "";
  const actorEmail = session?.user?.email || restAuth.email || "";
  return { session, restAuth, actorUserId, actorEmail };
}

function buildRestError(status, payload) {
  if (!payload || typeof payload !== "object") {
    return { message: `HTTP ${status}`, details: "", hint: "", code: status };
  }
  return {
    message:
      String(payload.message || payload.error || `HTTP ${status}`) ||
      `HTTP ${status}`,
    details: String(payload.details || ""),
    hint: String(payload.hint || ""),
    code: payload.code || status,
    status,
  };
}

async function restTableRequest(
  table,
  {
    method = "GET",
    params = null,
    body = null,
    accessToken = "",
    prefer = "",
    timeoutMs = 12000,
  } = {},
) {
  const query = new URLSearchParams();
  if (params && typeof params === "object") {
    Object.entries(params).forEach(([key, value]) => {
      if (value === undefined || value === null) return;
      query.set(key, String(value));
    });
  }
  const queryText = query.toString();
  const url = `${SUPABASE_URL}/rest/v1/${table}${queryText ? `?${queryText}` : ""}`;
  const headers = {
    apikey: SUPABASE_ANON_KEY,
    Authorization: `Bearer ${accessToken || SUPABASE_ANON_KEY}`,
  };
  if (body !== null) {
    headers["Content-Type"] = "application/json";
  }
  if (prefer) {
    headers.Prefer = prefer;
  }
  let timeout = null;
  try {
    const controller = new AbortController();
    timeout = setTimeout(() => controller.abort(), timeoutMs);
    const response = await fetch(url, {
      method,
      headers,
      body: body !== null ? JSON.stringify(body) : undefined,
      cache: "no-store",
      signal: controller.signal,
    });
    clearTimeout(timeout);
    timeout = null;
    const raw = await response.text();
    let parsed = null;
    try {
      parsed = raw ? JSON.parse(raw) : null;
    } catch (error) {
      parsed = raw;
    }
    if (!response.ok) {
      return { data: null, error: buildRestError(response.status, parsed) };
    }
    return { data: parsed, error: null };
  } catch (error) {
    if (timeout) {
      clearTimeout(timeout);
      timeout = null;
    }
    return {
      data: null,
      error: {
        message: errorMessage(error, "Network request failed."),
        details: "",
        hint: "",
        code: "NETWORK_ERROR",
      },
    };
  }
}

function getFreshCacheEntry(cacheMap, key, ttlMs) {
  const entry = cacheMap.get(key);
  if (!entry) return null;
  if (Date.now() - Number(entry.timestamp || 0) > ttlMs) {
    return null;
  }
  return entry;
}

function setCacheEntry(cacheMap, key, value) {
  cacheMap.set(key, {
    timestamp: Date.now(),
    ...value,
  });
}

function invalidateCommunityThemeCaches() {
  communityListCache.clear();
  communityThemeDetailCache.clear();
}

function invalidateAdminReportCache() {
  adminReportCache = null;
}

function invalidateAdminThemeCache() {
  adminThemeCache = null;
}

async function restFindOne(table, accessToken, filters = {}, select = "*") {
  const params = { select, limit: 1, ...filters };
  const result = await restTableRequest(table, {
    method: "GET",
    params,
    accessToken,
  });
  if (result.error) return { row: null, error: result.error };
  const list = Array.isArray(result.data) ? result.data : [];
  return { row: list[0] || null, error: null };
}

async function restUpsertByUserAndName(
  table,
  accessToken,
  actorUserId,
  name,
  payloadVariants = [],
) {
  const existingLookup = await restFindOne(
    table,
    accessToken,
    {
      user_id: `eq.${actorUserId}`,
      name: `eq.${name}`,
    },
    "id,visible,user_id,name",
  );
  let existing = existingLookup.row || null;
  let lastError = existingLookup.error || null;

  for (const payload of payloadVariants) {
    if (!payload || typeof payload !== "object") continue;
    let result = null;
    if (existing?.id) {
      result = await restTableRequest(table, {
        method: "PATCH",
        params: { id: `eq.${existing.id}`, select: "id,visible,user_id,name" },
        body: payload,
        accessToken,
        prefer: "return=representation",
      });
    } else {
      result = await restTableRequest(table, {
        method: "POST",
        params: { select: "id,visible,user_id,name" },
        body: [payload],
        accessToken,
        prefer: "return=representation",
      });
    }
    if (!result.error) {
      const rows = Array.isArray(result.data) ? result.data : [];
      const row = rows[0] || existing || null;
      return { ok: true, row, error: null };
    }
    lastError = result.error;
    if (!isSchemaCompatibilityError(result.error)) {
      break;
    }
  }

  return { ok: false, row: null, error: lastError };
}

async function loadForceSignedOutState() {
  try {
    const stored = await chrome.storage.local.get(FORCE_SIGNED_OUT_KEY);
    forceSignedOutState = Boolean(stored?.[FORCE_SIGNED_OUT_KEY]);
  } catch (error) {
    forceSignedOutState = false;
  }
}

async function setForceSignedOutState(value) {
  forceSignedOutState = Boolean(value);
  try {
    await chrome.storage.local.set({
      [FORCE_SIGNED_OUT_KEY]: forceSignedOutState,
    });
  } catch (error) {
    // ignore persistence failures
  }
}

function getEffectiveSession(session) {
  if (forceSignedOutState) return null;
  return session || null;
}

function syncAuthGateFromSession(session) {
  const effective = getEffectiveSession(session);
  const signedIn = Boolean(effective);
  const userId = signedIn ? String(effective?.user?.id || "") : "";
  setAuthGateState({
    authenticated: signedIn,
    hasUsername: signedIn,
    userId,
  }).catch(() => {
    // ignore sync errors
  });
}

async function enforceForcedSignOut(rawSession) {
  if (!forceSignedOutState || !rawSession || !supabaseClient) {
    return getEffectiveSession(rawSession);
  }
  try {
    await supabaseClient.auth.signOut({ scope: "local" });
  } catch (error) {
    // ignore cleanup failure
  }
  clearSupabaseStoredSession();
  return null;
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
  if (!origin) return "";
  try {
    const parsed = new URL(origin);
    return `${parsed.origin}/*`;
  } catch (error) {
    return "";
  }
}

async function unregisterCanvasContentScript() {
  if (!chrome?.scripting?.unregisterContentScripts) return;
  try {
    await chrome.scripting.unregisterContentScripts({
      ids: [CONTENT_SCRIPT_ID],
    });
  } catch (error) {
    // ignore missing script registration
  }
}

async function registerCanvasContentScript(matchPattern) {
  if (!chrome?.scripting?.registerContentScripts) {
    return { ok: false, error: "Scripting API unavailable." };
  }
  try {
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
    return { ok: true, error: "" };
  } catch (error) {
    return {
      ok: false,
      error: errorMessage(error, "Failed to register Canvas content script."),
    };
  }
}

async function ensureCanvasSiteAccess(baseUrl, { prompt = false } = {}) {
  const matchPattern = baseUrlToMatchPattern(baseUrl);
  if (!matchPattern) {
    await unregisterCanvasContentScript();
    await chrome.storage.local.set({ [CONTENT_SCRIPT_ORIGIN_KEY]: "" });
    return { ok: false, error: "Invalid Canvas base URL.", matchPattern: "" };
  }

  let granted = false;
  try {
    granted = await chrome.permissions.contains({ origins: [matchPattern] });
  } catch (error) {
    granted = false;
  }

  if (!granted && prompt) {
    try {
      granted = await chrome.permissions.request({ origins: [matchPattern] });
    } catch (error) {
      granted = false;
    }
  }

  if (!granted) {
    await unregisterCanvasContentScript();
    return {
      ok: false,
      error:
        "Site access not granted. Click Save and allow access for your Canvas URL.",
      matchPattern,
    };
  }

  await unregisterCanvasContentScript();
  const registration = await registerCanvasContentScript(matchPattern);
  if (!registration.ok) {
    return { ok: false, error: registration.error, matchPattern };
  }

  let previousPattern = "";
  try {
    const stored = await chrome.storage.local.get(CONTENT_SCRIPT_ORIGIN_KEY);
    previousPattern = String(stored?.[CONTENT_SCRIPT_ORIGIN_KEY] || "");
  } catch (error) {
    previousPattern = "";
  }

  if (previousPattern && previousPattern !== matchPattern) {
    try {
      await chrome.permissions.remove({ origins: [previousPattern] });
    } catch (error) {
      // ignore permission removal issues
    }
  }

  await chrome.storage.local.set({ [CONTENT_SCRIPT_ORIGIN_KEY]: matchPattern });
  return { ok: true, error: "", matchPattern };
}

async function clearCanvasSiteAccess() {
  let previousPattern = "";
  try {
    const stored = await chrome.storage.local.get(CONTENT_SCRIPT_ORIGIN_KEY);
    previousPattern = String(stored?.[CONTENT_SCRIPT_ORIGIN_KEY] || "");
  } catch (error) {
    previousPattern = "";
  }

  await unregisterCanvasContentScript();
  if (previousPattern) {
    try {
      await chrome.permissions.remove({ origins: [previousPattern] });
    } catch (error) {
      // ignore permission removal issues
    }
  }
  await chrome.storage.local.set({ [CONTENT_SCRIPT_ORIGIN_KEY]: "" });
}

function setPill(connected) {
  if (connected) {
    statusPill.textContent = "Connected";
    statusPill.classList.add("connected");
  } else {
    statusPill.textContent = "Not connected";
    statusPill.classList.remove("connected");
  }
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function hexToRgb(hex) {
  const cleaned = hex.replace("#", "").trim();
  if (cleaned.length !== 6) return null;
  const r = parseInt(cleaned.slice(0, 2), 16);
  const g = parseInt(cleaned.slice(2, 4), 16);
  const b = parseInt(cleaned.slice(4, 6), 16);
  return { r, g, b };
}

function toHex(value) {
  return value.toString(16).padStart(2, "0");
}

function mixColors(colorA, colorB, amount) {
  const a = hexToRgb(colorA);
  const b = hexToRgb(colorB);
  if (!a || !b) return colorA;
  const t = clamp(amount, 0, 1);
  const r = Math.round(a.r + (b.r - a.r) * t);
  const g = Math.round(a.g + (b.g - a.g) * t);
  const bVal = Math.round(a.b + (b.b - a.b) * t);
  return `#${toHex(r)}${toHex(g)}${toHex(bVal)}`;
}

function rgbaFromHex(hex, alpha) {
  const rgb = hexToRgb(hex);
  if (!rgb) return `rgba(31, 95, 139, ${alpha})`;
  return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})`;
}

function getContrastColor(hex) {
  const rgb = hexToRgb(hex);
  if (!rgb) return "#ffffff";
  const luminance = (0.2126 * rgb.r + 0.7152 * rgb.g + 0.0722 * rgb.b) / 255;
  return luminance > 0.6 ? "#0f172a" : "#ffffff";
}

function channelToLinear(channel) {
  const value = channel / 255;
  if (value <= 0.03928) return value / 12.92;
  return Math.pow((value + 0.055) / 1.055, 2.4);
}

function getLuminance(hex) {
  const rgb = hexToRgb(hex);
  if (!rgb) return 0;
  const r = channelToLinear(rgb.r);
  const g = channelToLinear(rgb.g);
  const b = channelToLinear(rgb.b);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function getContrastRatio(hexA, hexB) {
  const l1 = getLuminance(hexA);
  const l2 = getLuminance(hexB);
  const lighter = Math.max(l1, l2);
  const darker = Math.min(l1, l2);
  return (lighter + 0.05) / (darker + 0.05);
}

function ensureMinContrast(foreground, background, minRatio, mode) {
  if (getContrastRatio(foreground, background) >= minRatio) return foreground;
  const darkTarget = "#111827";
  const lightTarget = "#f8fafc";
  const preferredTarget = mode === "dark" ? lightTarget : darkTarget;
  let candidate = foreground;
  for (let step = 1; step <= 10; step += 1) {
    const mixed = mixColors(candidate, preferredTarget, step * 0.1);
    if (getContrastRatio(mixed, background) >= minRatio) {
      return mixed;
    }
  }
  const darkRatio = getContrastRatio(darkTarget, background);
  const lightRatio = getContrastRatio(lightTarget, background);
  return darkRatio >= lightRatio ? darkTarget : lightTarget;
}

function ensureMinContrastAcrossSurfaces(
  foreground,
  backgrounds,
  minRatio,
  mode,
) {
  const activeBackgrounds = (backgrounds || []).filter(Boolean);
  if (!activeBackgrounds.length) return foreground;
  let candidate = foreground;
  for (let pass = 0; pass < 3; pass += 1) {
    let updated = false;
    for (const bg of activeBackgrounds) {
      const improved = ensureMinContrast(candidate, bg, minRatio, mode);
      if (improved !== candidate) {
        candidate = improved;
        updated = true;
      }
    }
    if (!updated) break;
  }
  return candidate;
}

function normalizeThemeReadability(theme, mode, palette) {
  const normalized = { ...theme };
  const bg = theme.bg || palette.bg;
  const surface = theme.surface || palette.surface;
  const surfaceAlt = theme.surfaceAlt || palette.surfaceAlt || surface;
  const surfaces = [bg, surface, surfaceAlt];
  normalized.bg = bg;
  normalized.surface = surface;
  normalized.accent = ensureMinContrastAcrossSurfaces(
    theme.accent || "#1f5f8b",
    [surface, surfaceAlt],
    3.2,
    mode,
  );
  normalized.text = ensureMinContrastAcrossSurfaces(
    theme.text || palette.text,
    surfaces,
    4.9,
    mode,
  );
  normalized.muted = ensureMinContrastAcrossSurfaces(
    theme.muted || palette.muted,
    surfaces,
    3.9,
    mode,
  );
  normalized.border = ensureMinContrastAcrossSurfaces(
    theme.border || palette.border,
    [surface, surfaceAlt],
    1.9,
    mode,
  );
  return normalized;
}

function resolveThemeMode(mode) {
  if (mode === "auto") {
    return prefersDark.matches ? "dark" : "light";
  }
  return mode === "dark" ? "dark" : "light";
}

function getDefaultPalette(mode) {
  if (mode === "dark") {
    return {
      bg: "#111519",
      surface: "#1d2227",
      surfaceAlt: "#242a31",
      border: "#2c343b",
      text: "#f1f3f5",
      muted: "#a6b0b8",
    };
  }
  return {
    bg: "#f7f5f0",
    surface: "#ffffff",
    surfaceAlt: "#fff7ee",
    border: "#e6e1d6",
    text: "#1f2a33",
    muted: "#6f7b83",
  };
}

function applyTheme(theme) {
  const root = document.documentElement;
  const mode = resolveThemeMode(theme.mode || "auto");
  const palette = getDefaultPalette(mode);
  const safeTheme = normalizeThemeReadability(theme, mode, palette);
  const accent = safeTheme.accent || "#1f5f8b";

  root.dataset.theme = mode;
  root.style.setProperty("--accent", accent);
  root.style.setProperty("--font-body", safeTheme.fontBody || "Space Grotesk");
  root.style.setProperty("--font-head", safeTheme.fontHead || "Fraunces");
  root.style.setProperty("--bg", safeTheme.bg || palette.bg);
  root.style.setProperty("--surface", safeTheme.surface || palette.surface);
  root.style.setProperty("--text", safeTheme.text || palette.text);
  root.style.setProperty("--muted", safeTheme.muted || palette.muted);
  root.style.setProperty("--border", safeTheme.border || palette.border);

  const radius = clamp(Number(safeTheme.radius || 12), 6, 22);
  root.style.setProperty("--radius", `${radius}px`);
  root.style.setProperty("--radius-lg", `${radius + 4}px`);

  const shadowStrength = clamp(Number(safeTheme.shadow || 45), 0, 100) / 100;
  const shadowAlpha =
    mode === "dark" ? 0.5 * shadowStrength : 0.28 * shadowStrength;
  root.style.setProperty(
    "--shadow",
    `0 18px 42px rgba(26, 35, 46, ${shadowAlpha})`,
  );

  const intensity = clamp(Number(safeTheme.bgIntensity || 40), 0, 100) / 100;
  root.style.setProperty("--bg-glow-1", rgbaFromHex(accent, 0.18 * intensity));
  root.style.setProperty("--bg-glow-2", rgbaFromHex(accent, 0.12 * intensity));

  const surfaceContrast =
    clamp(Number(safeTheme.surfaceContrast || 50), 0, 100) / 100;
  const surfaceBase = safeTheme.surface || palette.surface;
  const bgBase = safeTheme.bg || palette.bg;
  const surfaceAlt =
    safeTheme.surfaceAlt ||
    (mode === "dark"
      ? mixColors(surfaceBase, bgBase, 0.28 + 0.52 * (1 - surfaceContrast))
      : mixColors(surfaceBase, bgBase, 0.18 + 0.38 * (1 - surfaceContrast)));
  root.style.setProperty("--surface-2", surfaceAlt);
  root.style.setProperty("--bg-2", mixColors(surfaceAlt, bgBase, 0.28));

  const customStyle =
    document.getElementById("qc-custom-theme") ||
    document.createElement("style");
  customStyle.id = "qc-custom-theme";
  customStyle.textContent = safeTheme.customCss || "";
  if (!customStyle.parentNode) {
    document.head.appendChild(customStyle);
  }
}

function getThemeFromInputs() {
  const draft = {
    mode: themeModeSelect.value,
    accent: accentColorInput.value || "#1f5f8b",
    bgIntensity: Number(bgIntensityInput.value || 40),
    surfaceContrast: Number(surfaceContrastInput.value || 50),
    bg: bgColorInput.value || "",
    surface: surfaceColorInput.value || "",
    surfaceAlt: surfaceAltColorInput.value || "",
    border: borderColorInput.value || "",
    text: textColorInput.value || "",
    muted: mutedColorInput.value || "",
    fontBody: fontBodySelect.value || "Space Grotesk",
    fontHead: fontHeadSelect.value || "Fraunces",
    radius: Number(radiusScaleInput.value || 12),
    shadow: Number(shadowStrengthInput.value || 45),
    customCss: customCssInput.value || "",
  };
  const mode = resolveThemeMode(draft.mode || "auto");
  return normalizeThemeReadability(draft, mode, getDefaultPalette(mode));
}

function getDefaultPopupTheme() {
  const theme = {
    mode: "auto",
    accent: "#1f5f8b",
    bgIntensity: 40,
    surfaceContrast: 50,
    bg: "",
    surface: "",
    surfaceAlt: "",
    border: "",
    text: "",
    muted: "",
    fontBody: "Space Grotesk",
    fontHead: "Fraunces",
    radius: 12,
    shadow: 45,
    customCss: "",
  };
  const mode = resolveThemeMode(theme.mode);
  const palette = getDefaultPalette(mode);
  return normalizeThemeReadability(theme, mode, palette);
}

function applyThemeToInputs(theme) {
  const mode = resolveThemeMode(theme.mode);
  const palette = getDefaultPalette(mode);
  const safeTheme = normalizeThemeReadability(theme, mode, palette);

  themeModeSelect.value = safeTheme.mode;
  accentColorInput.value = safeTheme.accent;
  bgIntensityInput.value = safeTheme.bgIntensity;
  surfaceContrastInput.value = safeTheme.surfaceContrast;
  bgColorInput.value = safeTheme.bg || palette.bg;
  surfaceColorInput.value = safeTheme.surface || palette.surface;
  surfaceAltColorInput.value = safeTheme.surfaceAlt || palette.surfaceAlt;
  borderColorInput.value = safeTheme.border || palette.border;
  textColorInput.value = safeTheme.text || palette.text;
  mutedColorInput.value = safeTheme.muted || palette.muted;
  fontBodySelect.value = safeTheme.fontBody;
  fontHeadSelect.value = safeTheme.fontHead;
  radiusScaleInput.value = safeTheme.radius;
  shadowStrengthInput.value = safeTheme.shadow;
  customCssInput.value = safeTheme.customCss;

  applyTheme(safeTheme);
}

function ensureSignedInForThemeAction() {
  if (isUiSignedIn) return true;
  setAuthStatus("Sign in to use themes.", true);
  setActiveTab("account");
  applyThemeToInputs(getDefaultPopupTheme());
  return false;
}

async function loadTheme(allowSavedTheme = true) {
  const [syncStored, localStored] = await Promise.all([
    chrome.storage.sync.get("popupTheme").catch(() => ({})),
    chrome.storage.local.get(POPUP_THEME_MIRROR_KEY).catch(() => ({})),
  ]);
  const popupTheme = syncStored?.popupTheme || null;
  const mirrorPayload = localStored?.[POPUP_THEME_MIRROR_KEY] || null;
  const mirrorTheme =
    mirrorPayload && typeof mirrorPayload === "object"
      ? mirrorPayload.theme || null
      : null;
  const syncUpdatedAt = Number(popupTheme?.updatedAt || 0);
  const mirrorUpdatedAt = Number(mirrorPayload?.updatedAt || 0);
  const effectiveTheme =
    mirrorTheme && mirrorUpdatedAt > syncUpdatedAt ? mirrorTheme : popupTheme;
  const theme = allowSavedTheme
    ? {
        mode: effectiveTheme?.mode || "auto",
        accent: effectiveTheme?.accent || "#1f5f8b",
        bgIntensity: effectiveTheme?.bgIntensity ?? 40,
        surfaceContrast: effectiveTheme?.surfaceContrast ?? 50,
        bg: effectiveTheme?.bg || "",
        surface: effectiveTheme?.surface || "",
        surfaceAlt: effectiveTheme?.surfaceAlt || "",
        border: effectiveTheme?.border || "",
        text: effectiveTheme?.text || "",
        muted: effectiveTheme?.muted || "",
        fontBody: effectiveTheme?.fontBody || "Space Grotesk",
        fontHead: effectiveTheme?.fontHead || "Fraunces",
        radius: effectiveTheme?.radius ?? 12,
        shadow: effectiveTheme?.shadow ?? 45,
        customCss: effectiveTheme?.customCss || "",
        updatedAt: Number(effectiveTheme?.updatedAt || 0),
      }
    : getDefaultPopupTheme();
  applyThemeToInputs(theme);
}

async function saveTheme() {
  if (!ensureSignedInForThemeAction()) return;
  const theme = {
    ...getThemeFromInputs(),
    updatedAt: Date.now(),
  };
  applyTheme(theme);
  await Promise.allSettled([
    chrome.storage.sync.set({ popupTheme: theme }),
    chrome.storage.local.set({
      [POPUP_THEME_MIRROR_KEY]: {
        theme,
        updatedAt: Number(theme.updatedAt || Date.now()),
      },
    }),
  ]);
}

function renderPresets() {
  presetGrid.innerHTML = PRESETS.map((preset, index) => {
    return `
      <button class="preset" type="button" data-index="${index}" data-name="${preset.name}">
        <div class="preset-title">${preset.name}</div>
        <div class="preset-swatches">
          <span class="preset-swatch" style="background:${preset.accent}"></span>
          <span class="preset-swatch" style="background:${preset.surface || (preset.mode === "dark" ? "#111827" : "#ffffff")}"></span>
          <span class="preset-swatch" style="background:${preset.bg || (preset.mode === "dark" ? "#0f172a" : "#f8fafc")}"></span>
        </div>
      </button>
    `;
  }).join("");

  presetGrid.querySelectorAll(".preset").forEach((button) => {
    button.addEventListener("click", async () => {
      if (!ensureSignedInForThemeAction()) return;
      const preset = PRESETS[Number(button.dataset.index)];
      if (!preset) return;
      const palette = getDefaultPalette(
        resolveThemeMode(preset.mode || "auto"),
      );
      themeModeSelect.value = preset.mode || "auto";
      accentColorInput.value = preset.accent || "#1f5f8b";
      bgIntensityInput.value = Number(preset.bgIntensity ?? 40);
      surfaceContrastInput.value = Number(preset.surfaceContrast ?? 50);
      bgColorInput.value = preset.bg || palette.bg;
      surfaceColorInput.value = preset.surface || palette.surface;
      surfaceAltColorInput.value = preset.surfaceAlt || palette.surfaceAlt;
      borderColorInput.value = preset.border || palette.border;
      textColorInput.value = preset.text || palette.text;
      mutedColorInput.value = preset.muted || palette.muted;
      fontBodySelect.value = preset.fontBody || "Space Grotesk";
      fontHeadSelect.value = preset.fontHead || "Fraunces";
      radiusScaleInput.value = Number(preset.radius ?? 12);
      shadowStrengthInput.value = Number(preset.shadow ?? 45);
      themeNameInput.value = preset.name || button.dataset.name || "";
      await saveTheme();
    });
  });
}

function renderFontOptions() {
  const options = FONT_OPTIONS.map(
    (font) => `<option value="${font.value}">${font.label}</option>`,
  ).join("");
  fontBodySelect.innerHTML = options;
  fontHeadSelect.innerHTML = options;
}

async function loadSettings() {
  const { canvasSettings, [CALENDAR_REMINDER_KEY]: calendarRemindersEnabled } =
    await chrome.storage.sync.get(["canvasSettings", CALENDAR_REMINDER_KEY]);
  if (!canvasSettings) {
    await clearCanvasSiteAccess();
    setPill(false);
    enabledInput.checked = true;
    if (calendarRemindersEnabledInput) {
      calendarRemindersEnabledInput.checked = Boolean(calendarRemindersEnabled);
    }
    return;
  }

  baseUrlInput.value = canvasSettings.baseUrl || "";
  apiTokenInput.value = canvasSettings.apiToken || "";
  enabledInput.checked = canvasSettings.enabled ?? true;
  if (calendarRemindersEnabledInput) {
    calendarRemindersEnabledInput.checked = Boolean(calendarRemindersEnabled);
  }
  const accessState = await ensureCanvasSiteAccess(canvasSettings.baseUrl, {
    prompt: false,
  });
  if (!accessState.ok) {
    setStatus(accessState.error, true);
  }
  setPill(
    Boolean(
      canvasSettings.baseUrl && canvasSettings.apiToken && accessState.ok,
    ),
  );
}

async function saveSettings() {
  const baseUrl = normalizeBaseUrl(baseUrlInput.value);
  const apiToken = apiTokenInput.value.trim();
  const enabled = enabledInput.checked;
  const calendarRemindersEnabled = Boolean(
    calendarRemindersEnabledInput?.checked,
  );

  const saveReminderPreference = async () => {
    await chrome.storage.sync.set({
      [CALENDAR_REMINDER_KEY]: calendarRemindersEnabled,
    });
    try {
      await chrome.runtime.sendMessage({
        type: "cfe-reminders-settings-updated",
        enabled: calendarRemindersEnabled,
      });
    } catch (error) {
      // ignore background reminder update errors
    }
  };

  const ensureGoogleCalendarConnected = async () => {
    if (!calendarRemindersEnabled) return true;
    try {
      const current = await chrome.runtime.sendMessage({
        type: "cfe-google-calendar-status",
      });
      if (current?.ok && current?.connected) return true;
    } catch (error) {
      // continue to interactive connect
    }

    try {
      const connected = await chrome.runtime.sendMessage({
        type: "cfe-google-calendar-connect",
        interactive: true,
      });
      if (connected?.ok && connected?.connected) return true;
      setStatus(
        connected?.error ||
          "Google Calendar permission is required to sync assignments.",
        true,
      );
      return false;
    } catch (error) {
      setStatus(`Google Calendar connect failed: ${errorMessage(error)}`, true);
      return false;
    }
  };

  const runImmediateCalendarSync = async () => {
    if (!calendarRemindersEnabled) return;
    try {
      await chrome.runtime.sendMessage({
        type: "cfe-sync-reminders-from-canvas",
      });
    } catch (error) {
      // ignore immediate sync errors; background periodic/content sync will retry
    }
  };

  if (!baseUrl || !apiToken) {
    await saveReminderPreference();
    const calendarConnected = await ensureGoogleCalendarConnected();
    if (!calendarConnected && calendarRemindersEnabled) {
      setPill(false);
      return;
    }
    await runImmediateCalendarSync();
    setStatus(
      "Saved reminder preference. Add a valid base URL and API token to update connection settings.",
      true,
    );
    setPill(false);
    return;
  }

  await chrome.storage.sync.set({
    canvasSettings: {
      baseUrl,
      apiToken,
      enabled,
    },
  });
  await saveReminderPreference();
  const calendarConnected = await ensureGoogleCalendarConnected();

  const accessState = await ensureCanvasSiteAccess(baseUrl, { prompt: true });
  if (!accessState.ok) {
    setStatus(accessState.error, true);
    setPill(false);
    return;
  }

  if (calendarRemindersEnabled && !calendarConnected) {
    setStatus(
      "Saved Canvas settings, but Google Calendar is not connected yet.",
      true,
    );
    setPill(true);
    return;
  }

  await runImmediateCalendarSync();

  if (calendarRemindersEnabled) {
    setStatus("Saved. Canvas connected and Google Calendar sync is on.");
  } else {
    setStatus("Saved. Access limited to your Canvas URL.");
  }
  setPill(true);
}

async function clearSettings() {
  await clearCanvasSiteAccess();
  await chrome.storage.sync.remove(["canvasSettings", CALENDAR_REMINDER_KEY]);
  baseUrlInput.value = "";
  apiTokenInput.value = "";
  enabledInput.checked = true;
  if (calendarRemindersEnabledInput) {
    calendarRemindersEnabledInput.checked = false;
  }
  try {
    await chrome.runtime.sendMessage({
      type: "cfe-reminders-settings-updated",
      enabled: false,
    });
  } catch (error) {
    // ignore background reminder update errors
  }
  setStatus("Settings cleared. Canvas site access removed.");
  setPill(false);
}

async function openCanvas() {
  const baseUrl = normalizeBaseUrl(baseUrlInput.value);
  if (!baseUrl) {
    setStatus("Add your Canvas base URL first.", true);
    return;
  }
  chrome.tabs.create({ url: baseUrl });
}

function openOptions() {
  if (chrome.runtime.openOptionsPage) {
    chrome.runtime.openOptionsPage();
  }
}

function setActiveTab(tabName) {
  if (isProfileSaveInProgress && tabName !== "account") {
    tabName = "account";
  }
  if (!isUiSignedIn && tabName !== "account") {
    tabName = "account";
  }
  tabs.forEach((tab) => {
    const isActive = tab.dataset.tab === tabName;
    tab.classList.toggle("is-active", isActive);
    tab.setAttribute("aria-selected", isActive ? "true" : "false");
  });
  panes.forEach((pane) => {
    const shouldShow = pane.dataset.pane === tabName;
    pane.hidden = !shouldShow;
  });
  try {
    localStorage.setItem(POPUP_ACTIVE_TAB_KEY, tabName);
  } catch (error) {
    // ignore local tab persistence issues
  }
  // Keep account-only sections hidden unless signed in.
  if (tabName === "account") {
    if (accountProfilePanel) {
      accountProfilePanel.hidden = !isUiSignedIn;
    }
    if (accountCloudPanel) {
      accountCloudPanel.hidden = true;
    }
    if (tokenHelpPanel) {
      tokenHelpPanel.hidden = true;
    }
  }
  if (tabName === "themes" && communityListEl && supabaseClient) {
    const hasItems = communityListEl.querySelector(".community-item");
    if (!hasItems) {
      loadCommunityThemes(
        sortLatestBtn.classList.contains("is-active") ? "latest" : "trending",
        { force: true },
      ).catch((error) => {
        console.warn(
          "[QuickCanvas] themes-tab community refresh failed:",
          error,
        );
      });
    }
  }
  if (tabName === "admin" && supabaseClient) {
    const hasReports = reportListEl?.querySelector(".admin-card");
    const hasThemes = adminThemesEl?.querySelector(".admin-card");
    if (!hasReports || !hasThemes) {
      Promise.all([loadAdminReports(), loadAdminThemes()]).catch((error) => {
        console.warn("[QuickCanvas] admin-tab preload failed:", error);
      });
    }
  }
}

function validateUsername(name) {
  if (!name) return "Username is required.";
  if (containsProfanity(name)) return "Please choose a different username.";
  if (!/^[a-zA-Z0-9_]{3,20}$/.test(name)) {
    return "Use 3-20 letters, numbers, or underscores.";
  }
  return "";
}

function validateThemeName(name) {
  if (!name) return "Theme name is required.";
  if (containsProfanity(name)) return "Please choose a different theme name.";
  if (name.length < 3 || name.length > 40) {
    return "Theme name must be 3-40 characters.";
  }
  return "";
}

function isAdmin(session) {
  return Boolean(
    session?.user?.email &&
    session.user.email.toLowerCase() === ADMIN_EMAIL.toLowerCase(),
  );
}

function isModerationRestricted(state) {
  if (!state) return false;
  if (state.status === "banned") return true;
  if (state.status !== "restricted") return false;
  if (!state.restricted_until) return true;
  return new Date(state.restricted_until).getTime() > Date.now();
}

function formatRestriction(state) {
  if (!state) return "";
  if (state.status === "banned") {
    return "Publishing disabled: account is banned.";
  }
  if (
    state.status === "restricted" &&
    state.restricted_until &&
    new Date(state.restricted_until).getTime() > Date.now()
  ) {
    return `Publishing restricted until ${new Date(state.restricted_until).toLocaleString()}.`;
  }
  if (state.status === "restricted") {
    return "Publishing is temporarily restricted.";
  }
  return "";
}

async function loadModerationState(userId) {
  if (!supabaseClient || !userId) return null;
  const { data, error } = await supabaseClient
    .from("cfe_user_moderation")
    .select("status, restricted_until, ban_reason")
    .eq("user_id", userId)
    .maybeSingle();
  if (error) return null;
  return data || null;
}

async function callReportFunction(action, payload, session, options = {}) {
  const restAuth = await getRestAuthContext(session);
  const token = session?.access_token || restAuth.accessToken || "";
  if (!token) {
    return {
      ok: false,
      status: 401,
      error: "Sign in required.",
      fnErrors: ["No session token."],
    };
  }
  const requestTimeoutMs = Math.max(
    1500,
    Number(options?.timeoutMs || 7000) || 7000,
  );
  const skipRefresh = Boolean(options?.skipRefresh);
  const fnList = Array.isArray(options?.functionNames)
    ? options.functionNames.filter(Boolean).map((name) => String(name))
    : [];
  const functionNames = fnList.length
    ? fnList
    : options?.onlyPrimary
      ? [REPORT_FUNCTION]
      : REPORT_FUNCTION_ALIASES;
  const fnErrors = [];
  let firstNon404Status = 0;
  let anyNon404FunctionError = false;
  for (const fnName of functionNames) {
    let triedRefreshForFn = false;
    let bearer = token;
    let authMode = "user";
    while (true) {
      let response = null;
      let timer = null;
      try {
        const controller = new AbortController();
        timer = setTimeout(() => controller.abort(), requestTimeoutMs);
        const headers = {
          "Content-Type": "application/json",
          apikey: SUPABASE_ANON_KEY,
          Authorization:
            authMode === "gateway"
              ? `Bearer ${SUPABASE_ANON_KEY}`
              : `Bearer ${bearer}`,
          "X-User-Authorization": `Bearer ${bearer}`,
        };
        response = await fetch(`${SUPABASE_URL}/functions/v1/${fnName}`, {
          method: "POST",
          headers,
          body: JSON.stringify({ action, ...payload }),
          signal: controller.signal,
        });
      } catch (error) {
        anyNon404FunctionError = true;
        if (!firstNon404Status) firstNon404Status = 500;
        fnErrors.push(
          `${fnName}: ${errorMessage(error, "Network request failed.")}`,
        );
        break;
      } finally {
        if (timer) clearTimeout(timer);
      }
      if (response.ok) {
        const data = await response.json().catch(() => ({}));
        return { ok: true, status: response.status, data, fnName };
      }
      const raw = await response.text().catch(() => "");
      let detail = raw.trim();
      try {
        const parsed = JSON.parse(raw);
        detail = String(parsed?.error || parsed?.message || detail);
      } catch (error) {
        // Keep raw text as fallback.
      }
      if (response.status === 401 && authMode === "user") {
        // Some deployments validate only via gateway anon + user header.
        authMode = "gateway";
        continue;
      }
      if (
        response.status === 401 &&
        authMode === "gateway" &&
        !triedRefreshForFn &&
        !skipRefresh
      ) {
        triedRefreshForFn = true;
        authMode = "user";
        try {
          const refreshed = await withTimeout(
            refreshAuthSession(),
            2500,
            "function auth refresh",
          );
          if (refreshed?.access_token) {
            bearer = refreshed.access_token;
            continue;
          }
        } catch (error) {
          // fall through to error reporting below
        }
      }
      if (response.status !== 404) anyNon404FunctionError = true;
      if (response.status !== 404 && !firstNon404Status) {
        firstNon404Status = response.status;
      }
      fnErrors.push(
        `${fnName}: ${response.status}${detail ? ` ${detail}` : ""}`,
      );
      break;
    }
  }
  return {
    ok: false,
    status: firstNon404Status || (anyNon404FunctionError ? 500 : 404),
    error: anyNon404FunctionError
      ? fnErrors.join(" | ")
      : "Reporting service is not deployed yet.",
    fnErrors,
  };
}

function isDuplicateInsertError(errorText) {
  const text = String(errorText || "").toLowerCase();
  return (
    text.includes("duplicate key") ||
    text.includes("23505") ||
    text.includes("already have an active report")
  );
}

function getDbErrorText(error) {
  if (!error) return "";
  const parts = [
    error.message,
    error.details,
    error.hint,
    error.code,
    error.status,
  ]
    .filter(Boolean)
    .map((part) => String(part));
  return parts.join(" | ");
}

function isMissingColumnError(error) {
  const text = getDbErrorText(error).toLowerCase();
  return (
    (text.includes("column") && text.includes("does not exist")) ||
    (text.includes("could not find") && text.includes("column"))
  );
}

function isMissingRelationError(error) {
  const text = getDbErrorText(error).toLowerCase();
  return text.includes("relation") && text.includes("does not exist");
}

function isSchemaCompatibilityError(error) {
  return isMissingColumnError(error) || isMissingRelationError(error);
}

function extractMissingColumn(error) {
  const text = getDbErrorText(error);
  const patterns = [
    /column\s+"?([a-zA-Z0-9_]+)"?\s+does not exist/i,
    /could not find the ['"]([a-zA-Z0-9_]+)['"] column/i,
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) return match[1];
  }
  return "";
}

function isAuthErrorLike(error) {
  const text = getDbErrorText(error).toLowerCase();
  return (
    text.includes("jwt") ||
    text.includes("token") ||
    text.includes("401") ||
    text.includes("expired") ||
    text.includes("auth")
  );
}

async function refreshAuthSession() {
  if (!supabaseClient) return null;
  const persistAndReturn = async (session) => {
    if (!session) return null;
    await persistAuthTokens(session);
    return session;
  };
  try {
    const { data } = await supabaseClient.auth.refreshSession();
    const session = data?.session || null;
    if (session) return await persistAndReturn(session);
  } catch (error) {
    // fall through to storage-backed recovery
  }
  try {
    const stored = await chrome.storage.local.get(AUTH_TOKENS_KEY);
    const tokens = stored?.[AUTH_TOKENS_KEY] || null;
    const refreshToken = String(tokens?.refresh_token || "");
    const accessToken = String(tokens?.access_token || "");
    if (refreshToken) {
      const { data } = await supabaseClient.auth.refreshSession({
        refresh_token: refreshToken,
      });
      const refreshedSession = data?.session || null;
      if (refreshedSession) return await persistAndReturn(refreshedSession);
    }
    if (accessToken && refreshToken) {
      const { data, error } = await supabaseClient.auth.setSession({
        access_token: accessToken,
        refresh_token: refreshToken,
      });
      if (!error && data?.session) {
        return await persistAndReturn(data.session);
      }
    }
  } catch (error) {
    // ignore recovery errors
  }
  return null;
}

async function executeWithAuthRetry(runQuery) {
  let result = null;
  try {
    result = await runQuery();
  } catch (error) {
    result = {
      data: null,
      error: {
        message: errorMessage(error, "Supabase query failed."),
        details: "",
        hint: "",
        code: "QUERY_THROW",
      },
    };
  }
  if (!result?.error || !isAuthErrorLike(result.error)) {
    return result;
  }
  await refreshAuthSession();
  try {
    result = await runQuery();
  } catch (error) {
    result = {
      data: null,
      error: {
        message: errorMessage(error, "Supabase query failed after refresh."),
        details: "",
        hint: "",
        code: "QUERY_THROW",
      },
    };
  }
  return result;
}

async function executeWithAuthRetryTimeout(
  runQuery,
  timeoutMs = 9000,
  label = "Supabase query",
) {
  try {
    return await withTimeout(executeWithAuthRetry(runQuery), timeoutMs, label);
  } catch (error) {
    return {
      data: null,
      error: {
        message: errorMessage(error, `${label} timed out.`),
        details: "",
        hint: "",
        code: "QUERY_TIMEOUT",
      },
    };
  }
}

async function resilientInsert(table, payload) {
  if (!supabaseClient) {
    return {
      data: null,
      error: { message: "Supabase library failed to load." },
    };
  }
  const safePayload =
    payload && typeof payload === "object" ? { ...payload } : payload;
  let attempts = 0;
  while (attempts < 12) {
    const result = await executeWithAuthRetryTimeout(
      () => supabaseClient.from(table).insert(safePayload),
      7000,
      `${table} insert`,
    );
    if (!result?.error) return result;
    const missingColumn = extractMissingColumn(result.error);
    if (
      !missingColumn ||
      typeof safePayload !== "object" ||
      !Object.prototype.hasOwnProperty.call(safePayload, missingColumn)
    ) {
      return result;
    }
    delete safePayload[missingColumn];
    attempts += 1;
  }
  return {
    data: null,
    error: {
      message: `Insert failed after retrying missing columns for ${table}.`,
    },
  };
}

async function resilientUpsert(table, payload, options = undefined) {
  if (!supabaseClient) {
    return {
      data: null,
      error: { message: "Supabase library failed to load." },
    };
  }
  const safePayload =
    payload && typeof payload === "object" ? { ...payload } : payload;
  let attempts = 0;
  while (attempts < 12) {
    const result = await executeWithAuthRetryTimeout(
      () => {
        if (options) {
          return supabaseClient.from(table).upsert(safePayload, options);
        }
        return supabaseClient.from(table).upsert(safePayload);
      },
      7000,
      `${table} upsert`,
    );
    if (!result?.error) return result;
    const missingColumn = extractMissingColumn(result.error);
    const isProtectedProfileSchoolStartColumn =
      table === "cfe_profiles" && missingColumn === "school_start_time";
    if (
      isProtectedProfileSchoolStartColumn ||
      !missingColumn ||
      typeof safePayload !== "object" ||
      !Object.prototype.hasOwnProperty.call(safePayload, missingColumn)
    ) {
      return result;
    }
    delete safePayload[missingColumn];
    attempts += 1;
  }
  return {
    data: null,
    error: {
      message: `Upsert failed after retrying missing columns for ${table}.`,
    },
  };
}

async function verifyLatestReport(themeId, reporterId) {
  if (!supabaseClient || !themeId || !reporterId) {
    return { found: false, verifiable: false, error: "Missing context." };
  }
  let { data, error } = await executeWithAuthRetry(() =>
    supabaseClient
      .from("cfe_reports")
      .select("id, status, created_at")
      .eq("theme_id", themeId)
      .eq("reporter_id", reporterId)
      .order("created_at", { ascending: false })
      .limit(1),
  );
  if (error && isMissingColumnError(error)) {
    ({ data, error } = await executeWithAuthRetry(() =>
      supabaseClient
        .from("cfe_reports")
        .select("id")
        .eq("theme_id", themeId)
        .eq("reporter_id", reporterId)
        .limit(1),
    ));
  }
  if (error) {
    return { found: false, verifiable: false, error: error.message };
  }
  return { found: Boolean(data?.length), verifiable: true, error: "" };
}

async function fallbackInsertReport(theme, reason, session) {
  if (!supabaseClient || !theme?.id || !session?.user?.id) {
    return { ok: false, error: "Missing fallback report context." };
  }
  const basePayload = {
    theme_id: theme.id,
    reason: String(reason || "").trim(),
    status: "pending",
    reporter_id: session.user.id,
    reporter_email: session.user.email || "",
    theme_name: theme.name || "",
    reported_username: theme.username || "Anonymous",
    reported_user_id: theme.user_id || null,
  };
  const payloadAttempts = [
    basePayload,
    {
      theme_id: basePayload.theme_id,
      reason: basePayload.reason,
      status: "pending",
      reporter_id: basePayload.reporter_id,
      reporter_email: basePayload.reporter_email,
    },
    {
      theme_id: basePayload.theme_id,
      reason: basePayload.reason,
      reporter_id: basePayload.reporter_id,
      reporter_email: basePayload.reporter_email,
    },
    {
      theme_id: basePayload.theme_id,
      reason: basePayload.reason,
      reporter_id: basePayload.reporter_id,
    },
    {
      theme_id: basePayload.theme_id,
      reason: basePayload.reason,
    },
  ];

  let lastError = "";
  for (const payload of payloadAttempts) {
    const { error } = await resilientInsert("cfe_reports", payload);
    if (!error) {
      return { ok: true, error: "", duplicate: false };
    }
    if (isDuplicateInsertError(error.message)) {
      return { ok: true, error: "", duplicate: true };
    }
    lastError =
      getDbErrorText(error) || error.message || "Report insert failed.";
  }

  return {
    ok: false,
    error: lastError || "Report insert failed.",
    duplicate: false,
  };
}

async function signIn() {
  if (!supabaseClient) {
    setAuthStatus("Supabase library failed to load.", true);
    await setAuthGateState({ authenticated: false, hasUsername: false });
    return;
  }
  // Clear forced-signout before attempting login to prevent auth-state races.
  await setForceSignedOutState(false);
  const email = authEmailInput.value.trim();
  const password = authPasswordInput.value;
  if (!email || !password) {
    setAuthStatus("Enter email and password.", true);
    return;
  }
  const { data: signInData, error } =
    await supabaseClient.auth.signInWithPassword({
      email,
      password,
    });
  if (error) {
    const rawMessage = String(error.message || "");
    if (/email.*not.*confirm/i.test(rawMessage)) {
      setAuthStatus(
        "Check email to confirm account, then sign in. If a confirmation tab says localhost refused to connect, you can close it.",
        true,
      );
    } else {
      setAuthStatus(rawMessage || "Sign in failed.", true);
    }
    return;
  }
  let session = signInData?.session || null;
  if (!session) {
    session = await getActiveSession();
  }
  if (!session) {
    setAuthStatus(
      "Sign-in succeeded but no active session was found. Reopen popup.",
      true,
    );
    return;
  }
  await persistAuthTokens(session);
  let username = await loadProfile();
  if (session && !username) {
    let savedProfileUsername = "";
    try {
      savedProfileUsername = await withTimeout(
        getSavedUsername(session.user.id),
        1500,
        "saved username lookup",
      );
    } catch (error) {
      savedProfileUsername = "";
    }
    const fallbackUsername = getFirstValidUsername(
      savedProfileUsername,
      getMetadataUsername(session),
      await getLocalUsername(session.user.id),
      usernameFromEmail(session.user.email || ""),
    );
    if (fallbackUsername) {
      const saveResult = await persistProfileSettings(session, {
        username: fallbackUsername,
        schoolStartMinutes:
          getMetadataSchoolStartMinutes(session) ?? currentSchoolStartMinutes,
      });
      if (saveResult.ok) {
        username = fallbackUsername;
        currentProfileUsername = fallbackUsername;
        displayNameInput.value = fallbackUsername;
        await setLocalUsername(session.user.id, fallbackUsername);
      } else {
        setProfileStatus(
          "Could not sync profile username to cloud. Signed in with local fallback.",
          true,
        );
        username = fallbackUsername;
        currentProfileUsername = fallbackUsername;
        displayNameInput.value = fallbackUsername;
        await setLocalUsername(session.user.id, fallbackUsername);
      }
    }
  }
  setAuthStatus("Signed in.");
  await setAuthGateState({
    authenticated: Boolean(session),
    hasUsername: Boolean(session),
    userId: session?.user?.id || "",
  });
  await loadTheme(true);
  updateAuthUI(session);
  setActiveTab("themes");
  await loadCloudThemes();
  await loadCommunityThemes(
    sortLatestBtn.classList.contains("is-active") ? "latest" : "trending",
    { force: true },
  );
}

async function signUp() {
  if (!supabaseClient) {
    setAuthStatus("Supabase library failed to load.", true);
    await setAuthGateState({ authenticated: false, hasUsername: false });
    return;
  }
  const email = authEmailInput.value.trim();
  const password = authPasswordInput.value;
  const username = authUsernameInput.value.trim();
  const schoolStartRaw = authSchoolStartInput?.value || "";
  const parsedSchoolStartMinutes = parseSchoolStartInputValue(schoolStartRaw);
  if (!email || !password || !username || parsedSchoolStartMinutes === null) {
    setAuthStatus(
      "Enter email, password, username, and a valid school start time.",
      true,
    );
    return;
  }
  const usernameError = validateUsername(username);
  if (usernameError) {
    setAuthStatus(usernameError, true);
    return;
  }
  if (await isUsernameTaken(username)) {
    setAuthStatus("That username is already taken.", true);
    return;
  }
  setAuthStatus("Creating account...");
  setStatus("Creating account...");
  await setForceSignedOutState(false);
  const { data, error } = await supabaseClient.auth.signUp({
    email,
    password,
    options: {
      data: {
        username,
        school_start_time: formatSchoolStartForDb(parsedSchoolStartMinutes),
        school_start_hhmm: formatSchoolStartForInput(parsedSchoolStartMinutes),
        school_start_minutes: clampSchoolStartMinutes(parsedSchoolStartMinutes),
      },
    },
  });
  if (error) {
    setAuthStatus(error.message, true);
    return;
  }
  await saveSchoolStartMinutes(parsedSchoolStartMinutes);
  // If email confirmation is disabled, persist profile immediately.
  if (data?.session?.user?.id) {
    await persistAuthTokens(data.session);
    await setForceSignedOutState(false);
    await setLocalUsername(data.session.user.id, username);
    currentProfileUsername = username;
    displayNameInput.value = username;
    await setAuthGateState({
      authenticated: true,
      hasUsername: true,
      userId: data.session.user.id,
    });
    setAuthStatus("Account created and signed in.");
    updateAuthUI(data.session);
    setActiveTab("themes");
    void (async () => {
      const profilePersistResult = await persistProfileSettings(data.session, {
        username,
        schoolStartMinutes: parsedSchoolStartMinutes,
      });
      if (profilePersistResult.ok) {
        await saveSchoolStartMinutes(parsedSchoolStartMinutes, {
          markPending: false,
        });
        const metadataResult = await syncAuthProfileMetadata({
          username,
          schoolStartMinutes: parsedSchoolStartMinutes,
        });
        if (!metadataResult.ok && metadataResult.error) {
          setProfileStatus(
            `Account created. Auth metadata sync may still be updating: ${metadataResult.error}`,
            true,
          );
        }
      } else {
        setProfileStatus(
          `Account created. Cloud profile sync failed: ${profilePersistResult.error || "Unknown error."}`,
          true,
        );
      }
      await Promise.allSettled([
        loadProfile(),
        loadCloudThemes(),
        loadCommunityThemes(
          sortLatestBtn.classList.contains("is-active") ? "latest" : "trending",
          { force: true },
        ),
      ]);
    })();
    return;
  }
  await setForceSignedOutState(false);
  await setAuthGateState({ authenticated: false, hasUsername: false });
  setAuthMode("signin");
  setAuthStatus(
    "Check email to confirm account. After clicking the confirmation link, return here and sign in. If the page says localhost refused to connect, you can close it.",
  );
}

async function signOut() {
  if (!supabaseClient) {
    setAuthStatus("Supabase library failed to load.", true);
    await setAuthGateState({ authenticated: false, hasUsername: false });
    return;
  }
  await setForceSignedOutState(true);
  currentProfileUsername = "";
  cloudSchoolStartMinutes = null;
  await saveSchoolStartMinutes(DEFAULT_SCHOOL_START_MINUTES, {
    markPending: false,
  });
  await setAuthGateState({ authenticated: false, hasUsername: false });
  updateAuthUI(null);
  setAuthMode("signin");
  setActiveTab("account");
  setPill(false);
  await loadTheme(false);

  const {
    data: { session: existingSession },
  } = await supabaseClient.auth.getSession();
  let signOutError = null;
  try {
    const { error } = await supabaseClient.auth.signOut({ scope: "local" });
    signOutError = error || null;
  } catch (error) {
    signOutError = error;
  }
  if (signOutError) {
    try {
      const { error } = await supabaseClient.auth.signOut();
      signOutError = error || null;
    } catch (error) {
      signOutError = error;
    }
  }
  clearSupabaseStoredSession();
  await clearAuthTokens();
  if (existingSession?.user?.id) {
    await clearLocalUsername(existingSession.user.id);
  }
  setAuthStatus(
    signOutError
      ? "Signed out locally. Reopen popup if account still appears signed in."
      : "Signed out.",
    Boolean(signOutError),
  );
  cloudListEl.innerHTML = "";
  authEmailInput.value = "";
  authPasswordInput.value = "";
  authUsernameInput.value = "";
  if (authSchoolStartInput) {
    authSchoolStartInput.value = formatSchoolStartForInput(
      DEFAULT_SCHOOL_START_MINUTES,
    );
  }
  if (profileSchoolStartInput) {
    profileSchoolStartInput.value = formatSchoolStartForInput(
      DEFAULT_SCHOOL_START_MINUTES,
    );
  }
  displayNameInput.value = "";
  if (profileEmailInput) {
    profileEmailInput.value = "";
    profileEmailInput.placeholder = "new-email@example.com";
  }
  if (profileCurrentPasswordInput) {
    profileCurrentPasswordInput.value = "";
  }
  if (profileNewPasswordInput) {
    profileNewPasswordInput.value = "";
  }
  if (profileNewPasswordConfirmInput) {
    profileNewPasswordConfirmInput.value = "";
  }
  profileStatusEl.textContent = "";
  if (adminTabBtn) adminTabBtn.hidden = true;
  await loadCommunityThemes(
    sortLatestBtn.classList.contains("is-active") ? "latest" : "trending",
    { force: true },
  );
}

async function loadProfile() {
  if (!supabaseClient) return "";
  const session = await getActiveSession();
  if (!session) return "";
  const localUsername = await getLocalUsername(session.user.id);
  let keepPendingSchoolStart = Boolean(currentSchoolStartPendingSync);
  if (
    currentSchoolStartPendingSync &&
    hasSchoolStartMinutesValue(currentSchoolStartMinutes)
  ) {
    const pendingMinutes = clampSchoolStartMinutes(currentSchoolStartMinutes);
    const pendingUsername = getFirstValidUsername(
      currentProfileUsername,
      getMetadataUsername(session),
      localUsername,
      authUsernameInput?.value,
      displayNameInput?.value,
    );
    const pendingSaveResult = await persistProfileSettings(session, {
      username: pendingUsername,
      schoolStartMinutes: pendingMinutes,
    });
    if (pendingSaveResult.ok) {
      const pendingMetadataResult = await syncAuthProfileMetadata({
        username: pendingUsername,
        schoolStartMinutes: pendingMinutes,
      });
      await saveSchoolStartMinutes(pendingMinutes, { markPending: false });
      cloudSchoolStartMinutes = pendingMinutes;
      keepPendingSchoolStart = false;
      if (!pendingMetadataResult.ok && pendingMetadataResult.error) {
        setProfileStatus(
          `School start synced to Supabase (${formatSchoolStartForInput(pendingMinutes)}), but metadata sync failed: ${pendingMetadataResult.error}`,
          true,
        );
      }
    } else if (!profileStatusEl.textContent) {
      setProfileStatus(
        `Pending school start sync failed: ${pendingSaveResult.error || "Unknown error."}`,
        true,
      );
      keepPendingSchoolStart = true;
    }
  }
  let data = null;
  let error = null;
  ({ data, error } = await executeWithAuthRetry(() =>
    supabaseClient
      .from("cfe_profiles")
      .select("username,school_start_time")
      .eq("user_id", session.user.id)
      .maybeSingle(),
  ));
  if (error && !isMissingColumnError(error)) {
    const readErrorText = `Cloud profile read failed: ${error.message || "Unknown error."}`;
    setProfileStatus(readErrorText, true);
    setStatus(readErrorText, true);
  }
  if (error && isMissingColumnError(error)) {
    ({ data, error } = await executeWithAuthRetry(() =>
      supabaseClient
        .from("cfe_profiles")
        .select("username")
        .eq("user_id", session.user.id)
        .maybeSingle(),
    ));
  }
  let username = "";
  const dbUsername =
    !error && data?.username ? String(data.username).trim() : "";
  if (dbUsername) {
    username = dbUsername;
  }
  const dbSchoolStartMinutes =
    !error && data?.school_start_time
      ? parseDbSchoolStartValue(data.school_start_time)
      : null;
  cloudSchoolStartMinutes =
    dbSchoolStartMinutes === null
      ? null
      : clampSchoolStartMinutes(dbSchoolStartMinutes);
  const metadataSchoolStartMinutes = getMetadataSchoolStartMinutes(session);
  const shouldRepairDbSchoolStart =
    dbSchoolStartMinutes !== null &&
    metadataSchoolStartMinutes !== null &&
    dbSchoolStartMinutes !== metadataSchoolStartMinutes &&
    dbSchoolStartMinutes === DEFAULT_SCHOOL_START_MINUTES &&
    metadataSchoolStartMinutes !== DEFAULT_SCHOOL_START_MINUTES;
  const pendingMatchesCloud =
    keepPendingSchoolStart &&
    dbSchoolStartMinutes !== null &&
    clampSchoolStartMinutes(dbSchoolStartMinutes) ===
      clampSchoolStartMinutes(currentSchoolStartMinutes);
  if (pendingMatchesCloud) {
    keepPendingSchoolStart = false;
  }
  const shouldPreferPendingLocal =
    keepPendingSchoolStart &&
    hasSchoolStartMinutesValue(currentSchoolStartMinutes);
  const resolvedSchoolStartMinutes = clampSchoolStartMinutes(
    shouldPreferPendingLocal
      ? currentSchoolStartMinutes
      : shouldRepairDbSchoolStart
        ? metadataSchoolStartMinutes
        : (dbSchoolStartMinutes ??
          metadataSchoolStartMinutes ??
          currentSchoolStartMinutes ??
          DEFAULT_SCHOOL_START_MINUTES),
  );
  await saveSchoolStartMinutes(resolvedSchoolStartMinutes, {
    markPending: shouldPreferPendingLocal,
  });

  const fallbackCandidate = getFirstValidUsername(
    getMetadataUsername(session),
    localUsername,
    authUsernameInput.value,
    displayNameInput.value,
    usernameFromEmail(session.user?.email || ""),
  );
  let persistedFallbackProfile = false;

  // Backfill for older accounts created before profile enforcement.
  if (!username && fallbackCandidate) {
    const saveResult = await persistProfileSettings(session, {
      username: fallbackCandidate,
      schoolStartMinutes: resolvedSchoolStartMinutes,
    });
    persistedFallbackProfile = saveResult.ok;
    if (!saveResult.ok && !profileStatusEl.textContent) {
      setProfileStatus(
        "Could not save your username to cloud. Using local username for now.",
        true,
      );
    }
    username = fallbackCandidate;
  }

  if (username && validateUsername(username) === "") {
    await setLocalUsername(session.user.id, username);
  }
  if (
    username &&
    (dbSchoolStartMinutes === null || shouldRepairDbSchoolStart) &&
    !persistedFallbackProfile
  ) {
    const saveResult = await persistProfileSettings(session, {
      username,
      schoolStartMinutes: resolvedSchoolStartMinutes,
    });
    if (!saveResult.ok) {
      setProfileStatus(
        `Could not sync school start to cloud: ${saveResult.error || "Unknown error."}`,
        true,
      );
    }
  }

  currentProfileUsername = username;
  if (username) {
    displayNameInput.value = username;
  }
  if (profileEmailInput) {
    profileEmailInput.value = "";
    profileEmailInput.placeholder = session.user?.email
      ? `Current: ${session.user.email}`
      : "new-email@example.com";
  }
  if (profileCurrentPasswordInput) {
    profileCurrentPasswordInput.value = "";
  }
  if (profileNewPasswordInput) {
    profileNewPasswordInput.value = "";
  }
  if (profileNewPasswordConfirmInput) {
    profileNewPasswordConfirmInput.value = "";
  }
  if (!authUsernameInput.value) {
    authUsernameInput.value = username || getMetadataUsername(session);
  }
  if (authSchoolStartInput) {
    authSchoolStartInput.value = formatSchoolStartForInput(
      resolvedSchoolStartMinutes,
    );
  }
  if (profileSchoolStartInput) {
    profileSchoolStartInput.value = formatSchoolStartForInput(
      resolvedSchoolStartMinutes,
    );
  }
  if (adminTabBtn) {
    adminTabBtn.hidden = !isAdmin(session);
  }
  await setAuthGateState({
    authenticated: true,
    hasUsername: true,
    userId: session.user.id,
  });
  return username;
}

async function getSavedUsername(userId) {
  if (!supabaseClient || !userId) return "";
  try {
    const { data, error } = await supabaseClient
      .from("cfe_profiles")
      .select("username")
      .eq("user_id", userId)
      .single();
    if (error) return "";
    return (data?.username || "").trim();
  } catch (error) {
    return "";
  }
}

async function syncPublicThemeUsername(accessToken, userId, username) {
  const safeUserId = String(userId || "").trim();
  const safeUsername = String(username || "").trim();
  if (!accessToken || !safeUserId || !safeUsername) return { ok: false };
  const result = await restTableRequest("cfe_public_themes", {
    method: "PATCH",
    params: { user_id: `eq.${safeUserId}` },
    body: { username: safeUsername },
    accessToken,
    prefer: "return=minimal",
    timeoutMs: 9000,
  });
  return { ok: !result.error, error: result.error || null };
}

async function saveProfile() {
  isProfileSaveInProgress = true;
  setActiveTab("account");
  setProfileStatus("Saving profile...");
  setStatus("Saving profile...");
  if (!supabaseClient) {
    setProfileStatus("Supabase library failed to load.", true);
    isProfileSaveInProgress = false;
    return;
  }
  const session = await getActiveSession();
  if (!session) {
    setProfileStatus("Sign in to set a username.", true);
    isProfileSaveInProgress = false;
    return;
  }
  try {
    const username = getFirstValidUsername(
      displayNameInput.value,
      currentProfileUsername,
      getMetadataUsername(session),
      await getLocalUsername(session.user.id),
      authUsernameInput?.value,
      usernameFromEmail(session.user?.email || ""),
    );
    displayNameInput.value = username;
    const errorText = validateUsername(username);
    if (errorText) {
      setProfileStatus(errorText, true);
      return;
    }
    const parsedSchoolStart =
      parseSchoolStartInputValue(profileSchoolStartInput?.value || "") ??
      parseSchoolStartInputValue(authSchoolStartInput?.value || "");
    if (parsedSchoolStart === null) {
      setProfileStatus("Enter a valid school start time.", true);
      return;
    }
    const schoolStartMinutes = parsedSchoolStart;
    const currentUserId = String(session.user.id || "");
    const currentUsername = String(currentProfileUsername || "").trim();
    const usernameChanged =
      username.toLowerCase() !== currentUsername.toLowerCase();
    if (usernameChanged) {
      const taken = await isUsernameTakenByOtherUser(username, currentUserId);
      if (taken) {
        setProfileStatus("That username is already taken.", true);
        return;
      }
    }
    const currentEmail = String(session.user?.email || "")
      .trim()
      .toLowerCase();
    const newEmail = String(profileEmailInput?.value || "")
      .trim()
      .toLowerCase();
    if (newEmail && !isValidEmailAddress(newEmail)) {
      setProfileStatus("Enter a valid new email address.", true);
      return;
    }
    const wantsEmailChange = Boolean(newEmail) && newEmail !== currentEmail;
    const newPassword = String(profileNewPasswordInput?.value || "");
    const newPasswordConfirm = String(
      profileNewPasswordConfirmInput?.value || "",
    );
    const hasAnyNewPassword =
      Boolean(newPassword) || Boolean(newPasswordConfirm);
    if (hasAnyNewPassword) {
      if (!newPassword || !newPasswordConfirm) {
        setProfileStatus("Enter and confirm your new password.", true);
        return;
      }
      if (newPassword !== newPasswordConfirm) {
        setProfileStatus("New password confirmation does not match.", true);
        return;
      }
      if (newPassword.length < 6) {
        setProfileStatus("New password must be at least 6 characters.", true);
        return;
      }
    }
    const wantsPasswordChange = hasAnyNewPassword;
    const needsSensitiveVerification = wantsEmailChange || wantsPasswordChange;
    const currentPassword = String(profileCurrentPasswordInput?.value || "");
    if (needsSensitiveVerification && !currentPassword) {
      setProfileStatus(
        "Enter your current password to change email or password.",
        true,
      );
      return;
    }
    const schoolStartAlreadySynced =
      schoolStartMinutes === currentSchoolStartMinutes &&
      cloudSchoolStartMinutes === schoolStartMinutes;
    if (
      !usernameChanged &&
      schoolStartAlreadySynced &&
      !wantsEmailChange &&
      !wantsPasswordChange
    ) {
      setProfileStatus("No profile changes detected.");
      return;
    }
    if (needsSensitiveVerification) {
      const verified = await verifyCurrentPasswordForSensitiveChange(
        session,
        currentPassword,
      );
      if (!verified.ok) {
        setProfileStatus(
          verified.error || "Current password is incorrect.",
          true,
        );
        return;
      }
    }
    await setLocalUsername(session.user.id, username);
    await saveSchoolStartMinutes(schoolStartMinutes);
    const profileMessages = [];
    const backgroundTasks = [];
    let hadAnyError = false;
    const { ok, error } = await persistProfileSettings(session, {
      username,
      schoolStartMinutes,
    });
    if (!ok) {
      const metadataFallback = await syncAuthProfileMetadata({
        username,
        schoolStartMinutes,
      });
      if (metadataFallback.ok) {
        const partialOkText = `Profile table sync failed, but auth metadata was updated: ${error || "Unknown error."}`;
        profileMessages.push(partialOkText);
        setProfileStatus(partialOkText, true);
        setStatus(partialOkText, true);
      } else {
        hadAnyError = true;
        const saveErrorText = `Saved locally. Cloud profile sync failed: ${error || "Unknown error."}`;
        profileMessages.push(saveErrorText);
        setProfileStatus(saveErrorText, true);
        setStatus(saveErrorText, true);
      }
    } else {
      await saveSchoolStartMinutes(schoolStartMinutes, { markPending: false });
      profileMessages.push("Profile details saved.");
      setProfileStatus("Profile details saved.");
      setStatus("Profile details saved.");
      backgroundTasks.push(
        (async () => {
          const metadataSyncResult = await syncAuthProfileMetadata({
            username,
            schoolStartMinutes,
          });
          if (!metadataSyncResult.ok && metadataSyncResult.error) {
            return {
              error: false,
              message:
                "Profile saved, but auth metadata sync may still be updating.",
            };
          }
          return null;
        })(),
      );
    }
    if (usernameChanged) {
      backgroundTasks.push(
        (async () => {
          const restAuth = await getRestAuthContext(session);
          if (!restAuth?.accessToken) return null;
          const syncResult = await syncPublicThemeUsername(
            restAuth.accessToken,
            session.user.id,
            username,
          );
          if (!syncResult.ok) {
            return {
              error: false,
              message:
                "Profile saved, but community theme username sync may still be updating.",
            };
          }
          return null;
        })(),
      );
    }
    if (wantsEmailChange) {
      backgroundTasks.push(
        (async () => {
          const { error: emailError } = await supabaseClient.auth.updateUser({
            email: newEmail,
          });
          if (emailError) {
            return {
              error: true,
              message: `Email change failed: ${emailError.message || "Unknown error."}`,
            };
          }
          return {
            error: false,
            message:
              "Email change requested. Check your inbox to verify the new email.",
          };
        })(),
      );
    }
    if (wantsPasswordChange) {
      backgroundTasks.push(
        (async () => {
          const { error: passwordError } = await supabaseClient.auth.updateUser(
            {
              password: newPassword,
            },
          );
          if (passwordError) {
            return {
              error: true,
              message: `Password change failed: ${passwordError.message || "Unknown error."}`,
            };
          }
          return { error: false, message: "Password updated." };
        })(),
      );
    }
    currentProfileUsername = username;
    authUsernameInput.value = username;
    if (profileSchoolStartInput) {
      profileSchoolStartInput.value =
        formatSchoolStartForInput(schoolStartMinutes);
    }
    if (profileEmailInput) {
      profileEmailInput.value = "";
    }
    if (profileCurrentPasswordInput) {
      profileCurrentPasswordInput.value = "";
    }
    if (profileNewPasswordInput) {
      profileNewPasswordInput.value = "";
    }
    if (profileNewPasswordConfirmInput) {
      profileNewPasswordConfirmInput.value = "";
    }
    await setAuthGateState({
      authenticated: true,
      hasUsername: true,
      userId: session.user.id,
    });
    setProfileStatus(profileMessages.join(" "), hadAnyError);
    setAuthStatus(
      hadAnyError
        ? "Profile update completed with issues."
        : "Profile updated.",
      hadAnyError,
    );
    updateAuthUI(session);
    if (backgroundTasks.length > 0) {
      const baseMessage = profileMessages.join(" ").trim();
      const baseHadError = hadAnyError;
      void Promise.allSettled(backgroundTasks).then((results) => {
        const lateMessages = [];
        let lateError = false;
        for (const result of results) {
          if (result.status !== "fulfilled" || !result.value) continue;
          if (result.value.message) lateMessages.push(result.value.message);
          if (result.value.error) lateError = true;
        }
        if (lateMessages.length === 0) return;
        const combined = [baseMessage, ...lateMessages]
          .filter(Boolean)
          .join(" ");
        setProfileStatus(combined, baseHadError || lateError);
        setAuthStatus(
          baseHadError || lateError
            ? "Profile update completed with issues."
            : "Profile updated.",
          baseHadError || lateError,
        );
      });
    }
  } finally {
    isProfileSaveInProgress = false;
  }
}

async function loadCloudThemes() {
  if (!supabaseClient) {
    cloudListEl.innerHTML =
      '<div class="status error">Supabase library failed to load.</div>';
    return;
  }
  const hasRenderedRows = Boolean(cloudListEl.querySelector(".cloud-item"));
  if (!hasRenderedRows) {
    cloudListEl.innerHTML = '<div class="status">Loading cloud themes...</div>';
  }
  const { restAuth, actorUserId } = await getFastAuthContext(1200);
  if (!actorUserId) {
    cloudListEl.innerHTML =
      '<div class="status">Sign in to load cloud themes.</div>';
    return;
  }
  let data = [];
  let error = null;
  const restResult = await restTableRequest("cfe_themes", {
    method: "GET",
    params: {
      select: "*",
      user_id: `eq.${actorUserId}`,
      order: "updated_at.desc",
      limit: 100,
    },
    accessToken: restAuth.accessToken,
    timeoutMs: 12000,
  });
  if (!restResult.error) {
    data = Array.isArray(restResult.data) ? restResult.data : [];
  } else {
    error = restResult.error;
  }
  if (error) {
    let dbResult = await executeWithAuthRetry(() =>
      supabaseClient
        .from("cfe_themes")
        .select("*")
        .eq("user_id", actorUserId)
        .order("updated_at", { ascending: false }),
    );
    if (dbResult.error && isMissingColumnError(dbResult.error)) {
      dbResult = await executeWithAuthRetry(() =>
        supabaseClient
          .from("cfe_themes")
          .select("*")
          .eq("user_id", actorUserId),
      );
    }
    if (!dbResult.error) {
      data = Array.isArray(dbResult.data) ? dbResult.data : [];
      error = null;
    } else {
      error = dbResult.error;
    }
  }
  if (error) {
    cloudListEl.innerHTML = `<div class="status error">${error.message}</div>`;
    return;
  }
  const rows = Array.isArray(data)
    ? [...data].sort((a, b) => {
        const aTime = new Date(a?.updated_at || a?.created_at || 0).getTime();
        const bTime = new Date(b?.updated_at || b?.created_at || 0).getTime();
        return bTime - aTime;
      })
    : [];
  if (!rows.length) {
    cloudListEl.innerHTML = '<div class="status">No cloud themes yet.</div>';
    return;
  }
  cloudListEl.innerHTML = rows
    .map((theme) => {
      const id = escapeAttr(theme?.id || "");
      const name = escapeHtml(theme?.name || "Untitled Theme");
      const mode = escapeHtml(theme?.mode || "auto");
      const accent = escapeHtml(theme?.accent || "#1f5f8b");
      return `
    <div class="cloud-item" data-id="${id}">
      <div class="preset-title">${name}</div>
      <div class="cloud-meta">${mode} · ${accent}</div>
      <div class="actions">
        <button class="primary cloud-apply" type="button" data-id="${id}">Apply</button>
        <button class="ghost cloud-delete" type="button" data-id="${id}">Delete</button>
      </div>
    </div>
  `;
    })
    .join("");

  cloudListEl.querySelectorAll(".cloud-apply").forEach((btn) => {
    btn.addEventListener("click", async () => {
      await withUiError("Apply cloud theme failed", async () => {
        const theme = rows.find((item) => String(item.id) === btn.dataset.id);
        if (!theme) return;
        themeModeSelect.value = theme.mode || "auto";
        accentColorInput.value = theme.accent || "#1f5f8b";
        bgIntensityInput.value = theme.bg_intensity ?? 40;
        surfaceContrastInput.value = theme.surface_contrast ?? 50;
        const mode = theme.mode === "dark" ? "dark" : "light";
        const palette = getDefaultPalette(mode);
        bgColorInput.value = theme.bg || palette.bg;
        surfaceColorInput.value = theme.surface || palette.surface;
        surfaceAltColorInput.value = theme.surface_alt || palette.surfaceAlt;
        borderColorInput.value = theme.border || palette.border;
        textColorInput.value = theme.text || palette.text;
        mutedColorInput.value = theme.muted || palette.muted;
        fontBodySelect.value = theme.font_body || "Space Grotesk";
        fontHeadSelect.value = theme.font_head || "Fraunces";
        radiusScaleInput.value = theme.radius ?? 12;
        shadowStrengthInput.value = theme.shadow ?? 45;
        customCssInput.value = theme.custom_css || "";
        themeNameInput.value = theme.name || "My Theme";
        await saveTheme();
      })();
    });
  });

  cloudListEl.querySelectorAll(".cloud-delete").forEach((btn) => {
    btn.addEventListener("click", async () => {
      await withUiError("Delete cloud theme failed", async () => {
        let result = await restTableRequest("cfe_themes", {
          method: "DELETE",
          params: {
            id: `eq.${btn.dataset.id}`,
            user_id: `eq.${actorUserId}`,
          },
          accessToken: restAuth.accessToken,
          timeoutMs: 12000,
        });
        if (result.error) {
          result = await executeWithAuthRetry(() =>
            supabaseClient
              .from("cfe_themes")
              .delete()
              .eq("id", btn.dataset.id)
              .eq("user_id", actorUserId),
          );
        }
        if (result.error) {
          throw new Error(result.error.message || "Cloud delete failed.");
        }
        await loadCloudThemes();
      })();
    });
  });
}

async function saveThemeToCloud() {
  setAuthStatus("Saving theme to cloud...");
  let session = null;
  try {
    session = await withTimeout(getActiveSession(), 4000, "session lookup");
  } catch (error) {
    session = null;
  }
  const restAuth = await getRestAuthContext(session);
  const actorUserId = session?.user?.id || restAuth.userId || "";
  if (!session && actorUserId) {
    session = {
      access_token: restAuth.accessToken,
      user: {
        id: actorUserId,
        email: restAuth.email || "",
      },
    };
  }
  if (!actorUserId || !restAuth.accessToken) {
    setAuthStatus("Sign in to save themes.", true);
    setActiveTab("account");
    return;
  }
  let moderationState = null;
  try {
    moderationState = await withTimeout(
      loadModerationState(actorUserId),
      2500,
      "moderation lookup",
    );
  } catch (error) {
    moderationState = null;
  }
  if (isModerationRestricted(moderationState)) {
    setAuthStatus(formatRestriction(moderationState), true);
    return;
  }
  const name = themeNameInput.value.trim() || "My Theme";
  const errorText = validateThemeName(name);
  if (errorText) {
    setAuthStatus(errorText, true);
    return;
  }
  const theme = getThemeFromInputs();
  const payloadVariants = [
    {
      user_id: actorUserId,
      name,
      mode: theme.mode,
      accent: theme.accent,
      bg: theme.bg || null,
      surface: theme.surface || null,
      surface_alt: theme.surfaceAlt || null,
      border: theme.border || null,
      text: theme.text || null,
      muted: theme.muted || null,
      bg_intensity: theme.bgIntensity,
      surface_contrast: theme.surfaceContrast,
      font_body: theme.fontBody,
      font_head: theme.fontHead,
      radius: theme.radius,
      shadow: theme.shadow,
      custom_css: theme.customCss,
    },
    {
      user_id: actorUserId,
      name,
      mode: theme.mode,
      accent: theme.accent,
      bg_intensity: theme.bgIntensity,
      surface_contrast: theme.surfaceContrast,
      font_body: theme.fontBody,
      font_head: theme.fontHead,
      radius: theme.radius,
      shadow: theme.shadow,
      custom_css: theme.customCss,
    },
    {
      user_id: actorUserId,
      name,
      mode: theme.mode,
      accent: theme.accent,
      font_body: theme.fontBody,
      font_head: theme.fontHead,
      custom_css: theme.customCss,
    },
    {
      user_id: actorUserId,
      name,
      mode: theme.mode,
      accent: theme.accent,
    },
  ];

  let saved = false;
  let lastError = null;
  for (const payload of payloadVariants) {
    const result = await restTableRequest("cfe_themes", {
      method: "POST",
      params: { select: "id,name,user_id" },
      body: [payload],
      accessToken: restAuth.accessToken,
      prefer: "return=representation",
      timeoutMs: 9000,
    });
    if (!result.error) {
      saved = true;
      break;
    }
    lastError = result.error;
    if (!isSchemaCompatibilityError(result.error)) {
      break;
    }
  }

  if (!saved) {
    setAuthStatus(
      `Cloud save failed: ${getDbErrorText(lastError) || "Unknown error"}`,
      true,
    );
    return;
  }
  setAuthStatus("Theme saved to cloud.");
  loadCloudThemes().catch((error) => {
    console.warn("[QuickCanvas] loadCloudThemes after save failed:", error);
  });
}

async function publishTheme() {
  setAuthStatus("Publishing theme...");
  let session = null;
  try {
    session = await withTimeout(getActiveSession(), 4000, "session lookup");
  } catch (error) {
    session = null;
  }
  const restAuth = await getRestAuthContext(session);
  const actorUserId = session?.user?.id || restAuth.userId || "";
  if (!session && actorUserId) {
    session = {
      access_token: restAuth.accessToken,
      user: {
        id: actorUserId,
        email: restAuth.email || "",
      },
    };
  }
  if (!actorUserId || !restAuth.accessToken) {
    setAuthStatus("Sign in to publish themes.", true);
    setActiveTab("account");
    return;
  }
  let moderationState = null;
  try {
    moderationState = await withTimeout(
      loadModerationState(actorUserId),
      2500,
      "moderation lookup",
    );
  } catch (error) {
    moderationState = null;
  }
  if (isModerationRestricted(moderationState)) {
    setAuthStatus(formatRestriction(moderationState), true);
    return;
  }
  const accountEmail = session.user?.email || restAuth.email || "";
  let username = getFirstValidUsername(
    displayNameInput.value,
    currentProfileUsername,
    getMetadataUsername(session),
    await getLocalUsername(actorUserId),
    usernameFromEmail(accountEmail),
  );
  if (!username) {
    try {
      username = await withTimeout(
        getSavedUsername(actorUserId),
        1500,
        "saved username lookup",
      );
    } catch (error) {
      username = "";
    }
    if (username) {
      displayNameInput.value = username;
    }
  }
  const usernameError = validateUsername(username);
  if (usernameError) {
    setAuthStatus("No valid username was found for publishing.", true);
    setActiveTab("account");
    return;
  }
  currentProfileUsername = username;
  displayNameInput.value = username;
  await setLocalUsername(actorUserId, username);
  syncPublicThemeUsername(restAuth.accessToken, actorUserId, username).catch(
    () => {
      // best-effort only
    },
  );
  if (supabaseClient && session?.user?.id) {
    withTimeout(
      persistProfileUsername(session, username),
      2000,
      "profile sync",
    ).catch(() => {
      // best-effort only
    });
  }
  const name = themeNameInput.value.trim() || "My Theme";
  const nameError = validateThemeName(name);
  if (nameError) {
    setAuthStatus(nameError, true);
    return;
  }
  const theme = getThemeFromInputs();
  const fullPayload = {
    user_id: actorUserId,
    username,
    name,
    visible: true,
    mode: theme.mode,
    accent: theme.accent,
    bg: theme.bg || null,
    surface: theme.surface || null,
    surface_alt: theme.surfaceAlt || null,
    border: theme.border || null,
    text: theme.text || null,
    muted: theme.muted || null,
    bg_intensity: theme.bgIntensity,
    surface_contrast: theme.surfaceContrast,
    font_body: theme.fontBody,
    font_head: theme.fontHead,
    radius: theme.radius,
    shadow: theme.shadow,
    custom_css: theme.customCss,
  };
  const payloadVariants = [
    fullPayload,
    {
      user_id: actorUserId,
      username,
      name,
      visible: true,
      mode: theme.mode,
      accent: theme.accent,
      bg_intensity: theme.bgIntensity,
      surface_contrast: theme.surfaceContrast,
      font_body: theme.fontBody,
      font_head: theme.fontHead,
      radius: theme.radius,
      shadow: theme.shadow,
      custom_css: theme.customCss,
    },
    {
      user_id: actorUserId,
      username,
      name,
      visible: true,
      mode: theme.mode,
      accent: theme.accent,
      font_body: theme.fontBody,
      font_head: theme.fontHead,
    },
    {
      user_id: actorUserId,
      username,
      name,
      mode: theme.mode,
      accent: theme.accent,
    },
  ];
  const upsertResult = await restUpsertByUserAndName(
    "cfe_public_themes",
    restAuth.accessToken,
    actorUserId,
    name,
    payloadVariants,
  );

  if (!upsertResult.ok) {
    setAuthStatus(
      `Publish failed: ${getDbErrorText(upsertResult.error) || "Unknown error"}`,
      true,
    );
    return;
  }

  const savedRowId = upsertResult.row?.id || null;
  if (savedRowId) {
    await restTableRequest("cfe_public_themes", {
      method: "PATCH",
      params: { id: `eq.${savedRowId}` },
      body: { visible: true },
      accessToken: restAuth.accessToken,
    });
  }

  setAuthStatus("Theme published.");
  invalidateCommunityThemeCaches();
  loadCommunityThemes(
    sortLatestBtn.classList.contains("is-active") ? "latest" : "trending",
    { force: true },
  ).catch((error) => {
    console.warn(
      "[QuickCanvas] loadCommunityThemes after publish failed:",
      error,
    );
  });
}

async function reportTheme(theme) {
  if (!supabaseClient) {
    setAuthStatus("Supabase library failed to load.", true);
    return;
  }
  let session = await getActiveSession();
  if (!session?.access_token) {
    const { data: refreshed } = await supabaseClient.auth.refreshSession();
    session = refreshed?.session || null;
    if (session) {
      await persistAuthTokens(session);
    }
  }
  if (!session?.access_token) {
    setAuthStatus("Sign in to report themes.", true);
    setActiveTab("account");
    return;
  }
  const reason = prompt("Why are you reporting this theme?") || "";
  if (reason.trim().length < 5) {
    setAuthStatus("Report reason must be at least 5 characters.", true);
    return;
  }

  const payload = {
    theme_id: theme.id,
    reason: reason.trim(),
    theme_name: theme.name,
    reported_username: theme.username || "Anonymous",
  };
  const direct = await fallbackInsertReport(theme, reason, session);
  if (!direct.ok) {
    const result = await callReportFunction("create_report", payload, session);
    if (!result.ok) {
      setAuthStatus(
        `Report failed: ${result.error} | Direct insert failed: ${direct.error}`,
        true,
      );
      return;
    }
    const verification = await verifyLatestReport(theme.id, session.user.id);
    if (verification.verifiable && !verification.found) {
      setAuthStatus(
        "Report function responded, but the report row was not found in DB.",
        true,
      );
      return;
    }
    setAuthStatus("Report sent and queued for admin review.");
    invalidateAdminReportCache();
    return;
  }

  if (direct.duplicate) {
    setAuthStatus("You already reported this theme. It is in review.");
  } else {
    setAuthStatus("Report sent and queued for admin review.");
    invalidateAdminReportCache();
  }

  // Best-effort notification path only; direct DB insert is authoritative.
  const notify = await callReportFunction("create_report", payload, session);
  if (!notify.ok) {
    console.warn("[QuickCanvas] report function notify failed:", notify.error);
  }
}

function sortCommunityRows(rows, order) {
  const next = Array.isArray(rows) ? [...rows] : [];
  next.sort((a, b) => {
    if (order === "latest") {
      const aTime = new Date(a?.created_at || a?.updated_at || 0).getTime();
      const bTime = new Date(b?.created_at || b?.updated_at || 0).getTime();
      return bTime - aTime;
    }
    const likeDiff = Number(b?.like_count || 0) - Number(a?.like_count || 0);
    if (likeDiff !== 0) return likeDiff;
    const aTime = new Date(a?.updated_at || a?.created_at || 0).getTime();
    const bTime = new Date(b?.updated_at || b?.created_at || 0).getTime();
    return bTime - aTime;
  });
  return next.slice(0, COMMUNITY_RENDER_LIMIT);
}

async function fetchCommunityThemeDetails(themeId, accessToken = "") {
  const id = String(themeId || "");
  if (!id) return null;
  const cached = getFreshCacheEntry(communityThemeDetailCache, id, 300_000);
  if (cached?.row) return cached.row;

  let token = String(accessToken || "");
  let restResult = await restTableRequest("cfe_public_themes", {
    method: "GET",
    params: { select: "*", id: `eq.${id}`, limit: 1 },
    accessToken: token,
    timeoutMs: 8000,
  });
  if (restResult.error && isAuthErrorLike(restResult.error)) {
    let refreshed = null;
    try {
      refreshed = await withTimeout(refreshAuthSession(), 3000, "auth refresh");
    } catch (error) {
      refreshed = null;
    }
    const auth = await getRestAuthContext(refreshed);
    token = auth.accessToken || token;
    restResult = await restTableRequest("cfe_public_themes", {
      method: "GET",
      params: { select: "*", id: `eq.${id}`, limit: 1 },
      accessToken: token,
      timeoutMs: 8000,
    });
  }

  let row = null;
  if (!restResult.error && Array.isArray(restResult.data)) {
    row = restResult.data[0] || null;
  }
  if (row) {
    setCacheEntry(communityThemeDetailCache, id, { row });
  }
  return row;
}

function renderCommunityThemes(rows, likedIds = new Set(), order = "trending") {
  if (!communityListEl) return;
  if (!rows.length) {
    communityListEl.innerHTML =
      '<div class="status">No community themes yet.</div>';
    return;
  }
  const likedSet =
    likedIds instanceof Set ? likedIds : new Set(Array.from(likedIds || []));

  communityListEl.innerHTML = rows
    .map((theme) => {
      const id = escapeAttr(theme?.id || "");
      const name = escapeHtml(theme?.name || "Untitled Theme");
      const username = escapeHtml(theme?.username || "Anonymous");
      const mode = escapeHtml(theme?.mode || "auto");
      const accent = escapeHtml(theme?.accent || "#1f5f8b");
      const likes = Number(theme?.like_count || 0);
      const liked = likedSet.has(theme?.id);
      const hiddenLabel = theme?.visible === false ? " · hidden" : "";
      return `
    <div class="community-item" data-id="${id}">
      <div class="preset-title">${name}</div>
      <div class="community-meta">by ${username}</div>
      <div class="tag">${mode} · ${accent} · ${likes} likes${hiddenLabel}</div>
      <div class="actions">
        <button class="primary community-apply" type="button" data-id="${id}">Apply</button>
        <button class="ghost community-like" type="button" data-id="${id}" data-liked="${liked}">${liked ? "Liked" : "Like"}</button>
        <button class="ghost community-report" type="button" data-id="${id}">Report</button>
      </div>
    </div>
  `;
    })
    .join("");

  communityListEl.querySelectorAll(".community-apply").forEach((btn) => {
    btn.addEventListener("click", async () => {
      await withUiError("Apply community theme failed", async () => {
        const theme = rows.find((item) => String(item.id) === btn.dataset.id);
        if (!theme) return;
        const { restAuth } = await getFastAuthContext(900);
        const fullTheme =
          (await fetchCommunityThemeDetails(theme.id, restAuth.accessToken)) ||
          theme;
        themeModeSelect.value = fullTheme.mode || "auto";
        accentColorInput.value = fullTheme.accent || "#1f5f8b";
        bgIntensityInput.value = fullTheme.bg_intensity ?? 40;
        surfaceContrastInput.value = fullTheme.surface_contrast ?? 50;
        const mode = fullTheme.mode === "dark" ? "dark" : "light";
        const palette = getDefaultPalette(mode);
        bgColorInput.value = fullTheme.bg || palette.bg;
        surfaceColorInput.value = fullTheme.surface || palette.surface;
        surfaceAltColorInput.value =
          fullTheme.surface_alt || palette.surfaceAlt;
        borderColorInput.value = fullTheme.border || palette.border;
        textColorInput.value = fullTheme.text || palette.text;
        mutedColorInput.value = fullTheme.muted || palette.muted;
        fontBodySelect.value = fullTheme.font_body || "Space Grotesk";
        fontHeadSelect.value = fullTheme.font_head || "Fraunces";
        radiusScaleInput.value = fullTheme.radius ?? 12;
        shadowStrengthInput.value = fullTheme.shadow ?? 45;
        customCssInput.value = fullTheme.custom_css || "";
        themeNameInput.value = fullTheme.name || "My Theme";
        await saveTheme();
      })();
    });
  });

  communityListEl.querySelectorAll(".community-like").forEach((btn) => {
    btn.addEventListener("click", async () => {
      await withUiError("Like action failed", async () => {
        const themeId = btn.dataset.id;
        const { restAuth, actorUserId } = await getFastAuthContext(1200);
        if (!actorUserId) {
          setAuthStatus("Sign in to like themes.", true);
          setActiveTab("account");
          return;
        }
        const likePayload = { theme_id: themeId, user_id: actorUserId };
        if (btn.dataset.liked === "true") {
          let error = null;
          const restResult = await restTableRequest("cfe_theme_likes", {
            method: "DELETE",
            params: { theme_id: `eq.${themeId}`, user_id: `eq.${actorUserId}` },
            accessToken: restAuth.accessToken,
            timeoutMs: 9000,
          });
          if (restResult.error) {
            const dbResult = await executeWithAuthRetry(() =>
              supabaseClient
                .from("cfe_theme_likes")
                .delete()
                .eq("theme_id", themeId)
                .eq("user_id", actorUserId),
            );
            error = dbResult.error;
          }
          if (error) {
            if (isMissingRelationError(error) || isMissingColumnError(error)) {
              setAuthStatus(
                "Theme likes table is missing in Supabase. Run the latest SQL migration.",
                true,
              );
              return;
            }
            throw new Error(error.message || "Un-like failed.");
          }
        } else {
          let error = null;
          const restResult = await restTableRequest("cfe_theme_likes", {
            method: "POST",
            params: { select: "theme_id,user_id" },
            body: [likePayload],
            accessToken: restAuth.accessToken,
            prefer: "return=representation",
            timeoutMs: 9000,
          });
          if (restResult.error) {
            const dbResult = await executeWithAuthRetry(() =>
              supabaseClient.from("cfe_theme_likes").insert(likePayload),
            );
            error = dbResult.error;
          }
          if (error) {
            if (isMissingRelationError(error) || isMissingColumnError(error)) {
              setAuthStatus(
                "Theme likes table is missing in Supabase. Run the latest SQL migration.",
                true,
              );
              return;
            }
            throw new Error(error.message || "Like failed.");
          }
        }
        invalidateCommunityThemeCaches();
        await loadCommunityThemes(order, { force: true });
      })();
    });
  });

  communityListEl.querySelectorAll(".community-report").forEach((btn) => {
    btn.addEventListener("click", async () => {
      await withUiError("Report action failed", async () => {
        const theme = rows.find((item) => String(item.id) === btn.dataset.id);
        if (!theme) return;
        await reportTheme(theme);
      })();
    });
  });
}

async function loadCommunityThemes(order = "trending", options = {}) {
  if (!communityListEl) return;
  if (!supabaseClient) {
    communityListEl.innerHTML =
      '<div class="status error">Supabase library failed to load.</div>';
    return;
  }
  const force = Boolean(options?.force);
  const hasRenderedRows = Boolean(
    communityListEl.querySelector(".community-item"),
  );
  if (!hasRenderedRows) {
    communityListEl.innerHTML =
      '<div class="status">Loading community themes...</div>';
  }
  const { restAuth, actorUserId } = await getFastAuthContext(1200);
  let authToken = String(restAuth.accessToken || "");
  const cacheKey = `${order}:${actorUserId || "anon"}`;
  const cachedEntry = communityListCache.get(cacheKey);
  const freshEntry = getFreshCacheEntry(
    communityListCache,
    cacheKey,
    COMMUNITY_LIST_CACHE_TTL_MS,
  );

  if (cachedEntry?.rows?.length) {
    renderCommunityThemes(
      cachedEntry.rows,
      new Set(cachedEntry.likedIds || []),
      order,
    );
    if (freshEntry && !force) return;
  }

  if (!force && communityLoadInFlight.has(cacheKey)) {
    return communityLoadInFlight.get(cacheKey);
  }

  const loadPromise = withTimeout(
    (async () => {
      const restGet = async (table, params, timeoutMs = 9000) => {
        let result = await restTableRequest(table, {
          method: "GET",
          params,
          accessToken: authToken,
          timeoutMs,
        });
        if (result.error && isAuthErrorLike(result.error)) {
          let refreshed = null;
          try {
            refreshed = await withTimeout(
              refreshAuthSession(),
              3000,
              "auth refresh",
            );
          } catch (error) {
            refreshed = null;
          }
          const auth = await getRestAuthContext(refreshed);
          authToken = auth.accessToken || authToken;
          result = await restTableRequest(table, {
            method: "GET",
            params,
            accessToken: authToken,
            timeoutMs,
          });
        }
        return result;
      };

      const orderCandidates =
        order === "latest"
          ? ["created_at.desc", "updated_at.desc"]
          : ["updated_at.desc", "created_at.desc"];

      const runCommunityThemeQuery = async (paramsBase, label) => {
        const visibilityCandidates = Object.prototype.hasOwnProperty.call(
          paramsBase,
          "visible",
        )
          ? [paramsBase.visible, null]
          : [null];
        let lastError = null;

        for (const visibleValue of visibilityCandidates) {
          for (const orderValue of orderCandidates) {
            const params = {
              ...paramsBase,
              select: "*",
              order: orderValue,
              limit: COMMUNITY_FETCH_LIMIT,
            };
            if (visibleValue === null) {
              delete params.visible;
            } else {
              params.visible = visibleValue;
            }

            const result = await restGet("cfe_public_themes", params, 8000);
            if (!result.error) {
              return {
                rows: Array.isArray(result.data) ? result.data : [],
                error: null,
              };
            }

            lastError = result.error;
            const missingColumn = String(
              extractMissingColumn(result.error) || "",
            ).toLowerCase();
            if (missingColumn === "visible" && visibleValue !== null) {
              // Retry same order without visibility filter for older schemas.
              break;
            }
            if (!isSchemaCompatibilityError(result.error)) {
              return {
                rows: [],
                error: new Error(
                  getDbErrorText(result.error) || `Failed to load ${label}.`,
                ),
              };
            }
          }
        }

        return {
          rows: [],
          error: new Error(
            getDbErrorText(lastError) || `Failed to load ${label}.`,
          ),
        };
      };

      const fetchPublicThemes = async () =>
        runCommunityThemeQuery({ visible: "eq.true" }, "public themes");

      const fetchOwnThemes = async () => {
        if (!actorUserId) return { rows: [], error: null };
        return runCommunityThemeQuery(
          { user_id: `eq.${actorUserId}`, visible: "eq.true" },
          "own themes",
        );
      };

      const [publicResult, ownResult] = await Promise.all([
        fetchPublicThemes(),
        fetchOwnThemes(),
      ]);
      const publicData = Array.isArray(publicResult?.rows)
        ? publicResult.rows
        : [];
      const ownData = Array.isArray(ownResult?.rows) ? ownResult.rows : [];
      const publicError = publicResult?.error || null;
      const ownError = ownResult?.error || null;
      if (!publicData.length && !ownData.length && publicError && ownError) {
        throw new Error(
          `${errorMessage(publicError, "Public themes query failed.")} | ${errorMessage(
            ownError,
            "Own themes query failed.",
          )}`,
        );
      }

      const byId = new Map();
      for (const row of publicData || []) {
        if (row?.id) byId.set(row.id, row);
      }
      for (const row of ownData || []) {
        if (row?.id) byId.set(row.id, row);
      }
      const rows = sortCommunityRows(Array.from(byId.values()), order);
      if (!rows.length) {
        setCacheEntry(communityListCache, cacheKey, { rows: [], likedIds: [] });
        communityListEl.innerHTML =
          '<div class="status">No community themes yet.</div>';
        return;
      }

      renderCommunityThemes(rows, new Set(), order);

      let likedIds = new Set();
      if (actorUserId) {
        const ids = rows.map((item) => item.id).filter(Boolean);
        if (ids.length) {
          const inList = ids
            .map((id) => `"${String(id).replace(/"/g, '\\"')}"`)
            .join(",");
          const likeRest = await restGet(
            "cfe_theme_likes",
            {
              select: "theme_id",
              user_id: `eq.${actorUserId}`,
              theme_id: `in.(${inList})`,
              limit: Math.min(200, ids.length),
            },
            7000,
          );
          if (!likeRest.error && Array.isArray(likeRest.data)) {
            likedIds = new Set(
              (likeRest.data || []).map((like) => like.theme_id),
            );
          }
        }
      }

      setCacheEntry(communityListCache, cacheKey, {
        rows,
        likedIds: Array.from(likedIds),
      });
      renderCommunityThemes(rows, likedIds, order);
    })(),
    14_000,
    "community themes load",
  )
    .catch((error) => {
      if (!cachedEntry?.rows?.length) {
        communityListEl.innerHTML = `<div class="status error">${errorMessage(error, "Failed to load community themes.")}</div>`;
      } else {
        console.warn("[QuickCanvas] community refresh failed:", error);
      }
    })
    .finally(() => {
      if (communityLoadInFlight.get(cacheKey) === loadPromise) {
        communityLoadInFlight.delete(cacheKey);
      }
    });

  communityLoadInFlight.set(cacheKey, loadPromise);
  return loadPromise;
}

function getActiveCommunityOrder() {
  return sortLatestBtn.classList.contains("is-active") ? "latest" : "trending";
}

function refreshAdminAndCommunityViews() {
  Promise.all([
    loadAdminReports({ force: true }),
    loadAdminThemes({ force: true }),
  ]).catch((error) => {
    console.warn("[QuickCanvas] admin refresh failed:", error);
  });
  loadCommunityThemes(getActiveCommunityOrder(), { force: true }).catch(
    (error) => {
      console.warn(
        "[QuickCanvas] community refresh after admin action failed:",
        error,
      );
    },
  );
}

async function adminReviewReport(reportId, verdict) {
  if (!supabaseClient) return;
  const note =
    prompt(
      verdict === REPORT_REVIEW_STATUS.VALID
        ? "Optional note for valid report:"
        : "Why is this report invalid?",
    ) || "";
  const { session, restAuth, actorEmail } = await getFastAuthContext(3000);
  const adminAllowed =
    isAdmin(session) ||
    (actorEmail && actorEmail.toLowerCase() === ADMIN_EMAIL.toLowerCase());
  if (!adminAllowed) {
    setAuthStatus("Admin access required.", true);
    return;
  }
  const sessionLike = session?.access_token
    ? session
    : { ...(session || {}), access_token: restAuth.accessToken || "" };
  if (!sessionLike?.access_token) {
    setAuthStatus("Session expired. Please sign in again.", true);
    return;
  }
  const result = await callReportFunction(
    "review_report",
    {
      report_id: reportId,
      verdict,
      note: note.trim(),
    },
    sessionLike,
  );
  if (!result.ok) {
    setAuthStatus(`Review failed: ${result.error}`, true);
    return;
  }
  if (result.data?.auto_hidden) {
    setAuthStatus(
      "Report marked valid. Theme hit threshold and was auto-hidden.",
    );
  } else {
    setAuthStatus("Report review saved.");
  }
  invalidateAdminReportCache();
  invalidateAdminThemeCache();
  invalidateCommunityThemeCaches();
  refreshAdminAndCommunityViews();
}

async function adminModerateUser(userId, actionType) {
  if (!userId || !supabaseClient) return;
  const reason =
    prompt("Reason for this moderation action? (required)")?.trim() || "";
  if (!reason) {
    setAuthStatus("Moderation action cancelled: reason is required.", true);
    return;
  }
  let days = 0;
  if (actionType === "restrict") {
    const rawDays = prompt("Restrict for how many days?", "7") || "7";
    days = Math.max(1, Number(rawDays) || 7);
  }
  const { session, restAuth, actorEmail } = await getFastAuthContext(3000);
  const adminAllowed =
    isAdmin(session) ||
    (actorEmail && actorEmail.toLowerCase() === ADMIN_EMAIL.toLowerCase());
  if (!adminAllowed) {
    setAuthStatus("Admin access required.", true);
    return;
  }
  const sessionLike = session?.access_token
    ? session
    : { ...(session || {}), access_token: restAuth.accessToken || "" };
  if (!sessionLike?.access_token) {
    setAuthStatus("Session expired. Please sign in again.", true);
    return;
  }
  const result = await callReportFunction(
    "moderate_user",
    {
      target_user_id: userId,
      action_type: actionType,
      reason,
      days,
    },
    sessionLike,
  );
  if (!result.ok) {
    setAuthStatus(`Moderation failed: ${result.error}`, true);
    return;
  }
  setAuthStatus("Moderation action saved.");
  invalidateAdminReportCache();
  invalidateAdminThemeCache();
  invalidateCommunityThemeCaches();
  refreshAdminAndCommunityViews();
}

async function adminModerateTheme(themeId, actionType) {
  if (!themeId || !supabaseClient) return;
  const reason =
    prompt("Reason for this theme moderation action? (required)")?.trim() || "";
  if (!reason) {
    setAuthStatus("Theme moderation cancelled: reason is required.", true);
    return;
  }
  const { session, restAuth, actorEmail } = await getFastAuthContext(3000);
  const adminAllowed =
    isAdmin(session) ||
    (actorEmail && actorEmail.toLowerCase() === ADMIN_EMAIL.toLowerCase());
  if (!adminAllowed) {
    setAuthStatus("Admin access required.", true);
    return;
  }
  const sessionLike = session?.access_token
    ? session
    : { ...(session || {}), access_token: restAuth.accessToken || "" };
  if (!sessionLike?.access_token) {
    setAuthStatus("Session expired. Please sign in again.", true);
    return;
  }
  const result = await callReportFunction(
    "moderate_theme",
    {
      theme_id: themeId,
      action_type: actionType,
      reason,
    },
    sessionLike,
  );
  if (!result.ok) {
    setAuthStatus(`Theme moderation failed: ${result.error}`, true);
    return;
  }
  if (actionType === "remove") {
    setAuthStatus("Theme removed from community themes.");
  } else if (actionType === "hide") {
    setAuthStatus("Theme hidden from community themes.");
  } else if (actionType === "show") {
    setAuthStatus("Theme is visible in community themes.");
  } else {
    setAuthStatus("Theme moderation action saved.");
  }
  invalidateAdminThemeCache();
  invalidateAdminReportCache();
  invalidateCommunityThemeCaches();
  refreshAdminAndCommunityViews();
}

function renderAdminReports(data) {
  if (!reportListEl) return;
  if (!data?.length) {
    reportListEl.innerHTML = '<div class="status">No reports.</div>';
    return;
  }
  reportListEl.innerHTML = data
    .map((report) => {
      const reportId = escapeAttr(report?.id || "");
      const themeName = escapeHtml(report?.theme_name || "Reported theme");
      const reportedUsername = escapeHtml(
        report?.reported_username || "Unknown",
      );
      const createdLabel = escapeHtml(
        new Date(report?.created_at || Date.now()).toLocaleString(),
      );
      const statusLabel = escapeHtml(
        String(report?.status || REPORT_REVIEW_STATUS.PENDING).toUpperCase(),
      );
      const reasonLabel = escapeHtml(report?.reason || "");
      const userId = escapeAttr(report?.reported_user_id || "");
      return `
    <div class="admin-card" data-id="${reportId}">
      <div class="preset-title">${themeName}</div>
      <div class="community-meta">by ${reportedUsername}</div>
      <div class="tag">${createdLabel}</div>
      <div class="tag">Status: ${statusLabel}</div>
      <div class="cloud-meta">Reason: ${reasonLabel}</div>
      <div class="actions">
        ${
          (report?.status || REPORT_REVIEW_STATUS.PENDING) ===
          REPORT_REVIEW_STATUS.PENDING
            ? `
        <button class="secondary admin-report-valid" type="button" data-id="${reportId}">Mark Valid</button>
        <button class="ghost admin-report-invalid" type="button" data-id="${reportId}">Mark Invalid</button>
        `
            : ""
        }
        ${
          report?.reported_user_id
            ? `
        <button class="ghost admin-user-restrict" type="button" data-user="${userId}">Restrict</button>
        <button class="ghost admin-user-ban" type="button" data-user="${userId}">Ban</button>
        <button class="ghost admin-user-unban" type="button" data-user="${userId}">Unban</button>
        `
            : ""
        }
      </div>
    </div>
  `;
    })
    .join("");

  reportListEl.querySelectorAll(".admin-report-valid").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminReviewReport(btn.dataset.id, REPORT_REVIEW_STATUS.VALID),
    );
  });
  reportListEl.querySelectorAll(".admin-report-invalid").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminReviewReport(btn.dataset.id, REPORT_REVIEW_STATUS.INVALID),
    );
  });
  reportListEl.querySelectorAll(".admin-user-restrict").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminModerateUser(btn.dataset.user, "restrict"),
    );
  });
  reportListEl.querySelectorAll(".admin-user-ban").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminModerateUser(btn.dataset.user, "ban"),
    );
  });
  reportListEl.querySelectorAll(".admin-user-unban").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminModerateUser(btn.dataset.user, "unban"),
    );
  });
}

async function fetchAdminReportsFromRest(accessToken) {
  const selectCandidates = [
    "id,theme_id,theme_name,reported_username,reported_user_id,reason,status,created_at,updated_at",
    "*",
  ];
  const orderCandidates = ["created_at.desc", "updated_at.desc"];
  let lastError = null;
  for (const selectValue of selectCandidates) {
    for (const orderValue of orderCandidates) {
      const result = await restTableRequest("cfe_reports", {
        method: "GET",
        params: {
          select: selectValue,
          order: orderValue,
          limit: ADMIN_REPORT_LIMIT,
        },
        accessToken,
        timeoutMs: 4500,
      });
      if (!result.error) {
        return {
          data: Array.isArray(result.data) ? result.data : [],
          error: null,
        };
      }
      lastError = result.error;
      if (!isSchemaCompatibilityError(result.error)) {
        return {
          data: [],
          error: new Error(
            getDbErrorText(result.error) || "Failed to load reports.",
          ),
        };
      }
    }
  }
  return {
    data: [],
    error: new Error(getDbErrorText(lastError) || "Failed to load reports."),
  };
}

async function loadAdminReports(options = {}) {
  if (!supabaseClient || !reportListEl) return;
  const force = Boolean(options?.force);
  const { session, actorEmail } = await getFastAuthContext(1200);
  const adminAllowed =
    isAdmin(session) ||
    (actorEmail && actorEmail.toLowerCase() === ADMIN_EMAIL.toLowerCase());
  if (!adminAllowed) return;

  const actorKey = String(session?.user?.id || actorEmail || "admin");
  const cached =
    adminReportCache && adminReportCache.actorKey === actorKey
      ? adminReportCache
      : null;
  const isFresh =
    cached &&
    Date.now() - Number(cached.timestamp || 0) <= ADMIN_REPORT_CACHE_TTL_MS;

  if (cached && Array.isArray(cached.data)) {
    renderAdminReports(cached.data);
    if (isFresh && !force) return;
  } else if (!adminReportLoadInFlight) {
    reportListEl.innerHTML = '<div class="status">Loading reports...</div>';
  }

  if (adminReportLoadInFlight && !force) {
    return adminReportLoadInFlight;
  }

  const loadPromise = withTimeout(
    (async () => {
      let effectiveSession = session;
      if (!effectiveSession?.access_token) {
        try {
          effectiveSession = await withTimeout(
            getActiveSession(),
            2200,
            "admin reports session restore",
          );
        } catch (error) {
          // fall back to stored token below
        }
      }
      let authToken = String(effectiveSession?.access_token || "");
      if (!authToken) {
        const auth = await getRestAuthContext(effectiveSession || session);
        authToken = String(auth.accessToken || "");
      }
      const sessionLike = authToken
        ? { ...(effectiveSession || session || {}), access_token: authToken }
        : effectiveSession || session || null;

      const fnResult = await callReportFunction(
        "list_admin_reports",
        { limit: ADMIN_REPORT_LIMIT },
        sessionLike,
        { onlyPrimary: true, timeoutMs: 1800 },
      );
      if (fnResult.ok && Array.isArray(fnResult?.data?.reports)) {
        const data = fnResult.data.reports;
        adminReportCache = {
          actorKey,
          timestamp: Date.now(),
          data,
        };
        renderAdminReports(data);
        return;
      }
      const restFallback = await fetchAdminReportsFromRest(authToken);
      if (!restFallback.error) {
        const data = Array.isArray(restFallback.data) ? restFallback.data : [];
        adminReportCache = {
          actorKey,
          timestamp: Date.now(),
          data,
        };
        renderAdminReports(data);
        return;
      }
      throw new Error(
        `Admin reports function failed: ${fnResult.error || "Unknown error."} | REST fallback failed: ${errorMessage(
          restFallback.error,
          "Unknown error.",
        )}`,
      );
    })(),
    9000,
    "admin reports load",
  )
    .catch((error) => {
      if (!cached?.data?.length) {
        reportListEl.innerHTML = `<div class="status error">${errorMessage(error, "Failed to load reports.")}</div>`;
      } else {
        console.warn("[QuickCanvas] reports refresh failed:", error);
      }
    })
    .finally(() => {
      if (adminReportLoadInFlight === loadPromise) {
        adminReportLoadInFlight = null;
      }
    });

  adminReportLoadInFlight = loadPromise;
  return loadPromise;
}

function renderAdminThemes(data) {
  if (!adminThemesEl) return;
  if (!data?.length) {
    adminThemesEl.innerHTML = '<div class="status">No themes.</div>';
    return;
  }
  adminThemesEl.innerHTML = data
    .map((theme) => {
      const themeId = escapeAttr(theme?.id || "");
      const userId = escapeAttr(theme?.user_id || "");
      const name = escapeHtml(theme?.name || "Untitled Theme");
      const username = escapeHtml(theme?.username || "Anonymous");
      const mode = escapeHtml(theme?.mode || "auto");
      const accent = escapeHtml(theme?.accent || "#1f5f8b");
      const likeCount = Number(theme?.like_count || 0);
      const visible = Boolean(theme?.visible);
      return `
    <div class="admin-card" data-id="${themeId}">
      <div class="preset-title">${name}</div>
      <div class="community-meta">by ${username}</div>
      <div class="tag">${mode} · ${accent} · ${likeCount} likes</div>
      <div class="tag">Visible: ${visible ? "Yes" : "Hidden"}</div>
      <div class="actions">
        <button class="ghost admin-remove" type="button" data-id="${themeId}">Remove</button>
        <button class="ghost admin-toggle-visibility" type="button" data-id="${themeId}" data-visible="${visible}">
          ${visible ? "Hide" : "Show"}
        </button>
        <button class="ghost admin-user-restrict" type="button" data-user="${userId}">Restrict</button>
        <button class="ghost admin-user-ban" type="button" data-user="${userId}">Ban</button>
        <button class="ghost admin-user-unban" type="button" data-user="${userId}">Unban</button>
      </div>
    </div>
  `;
    })
    .join("");

  adminThemesEl.querySelectorAll(".admin-remove").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminModerateTheme(btn.dataset.id, "remove"),
    );
  });
  adminThemesEl.querySelectorAll(".admin-toggle-visibility").forEach((btn) => {
    btn.addEventListener("click", () => {
      const nextVisible = btn.dataset.visible !== "true";
      adminModerateTheme(btn.dataset.id, nextVisible ? "show" : "hide");
    });
  });
  adminThemesEl.querySelectorAll(".admin-user-restrict").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminModerateUser(btn.dataset.user, "restrict"),
    );
  });
  adminThemesEl.querySelectorAll(".admin-user-ban").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminModerateUser(btn.dataset.user, "ban"),
    );
  });
  adminThemesEl.querySelectorAll(".admin-user-unban").forEach((btn) => {
    btn.addEventListener("click", () =>
      adminModerateUser(btn.dataset.user, "unban"),
    );
  });
}

async function fetchAdminThemesFromRest(accessToken) {
  const selectCandidates = [ADMIN_THEME_SELECT, "*"];
  const orderCandidates = ["updated_at.desc", "created_at.desc"];
  let lastError = null;
  for (const selectValue of selectCandidates) {
    for (const orderValue of orderCandidates) {
      const result = await restTableRequest("cfe_public_themes", {
        method: "GET",
        params: {
          select: selectValue,
          order: orderValue,
          limit: ADMIN_THEME_LIMIT,
        },
        accessToken,
        timeoutMs: 4500,
      });
      if (!result.error) {
        return {
          data: Array.isArray(result.data) ? result.data : [],
          error: null,
        };
      }
      lastError = result.error;
      if (!isSchemaCompatibilityError(result.error)) {
        return {
          data: [],
          error: new Error(
            getDbErrorText(result.error) || "Failed to load admin themes.",
          ),
        };
      }
    }
  }
  return {
    data: [],
    error: new Error(
      getDbErrorText(lastError) || "Failed to load admin themes.",
    ),
  };
}

async function loadAdminThemes(options = {}) {
  if (!supabaseClient || !adminThemesEl) return;
  const force = Boolean(options?.force);
  const { session, actorEmail } = await getFastAuthContext(1200);
  const adminAllowed =
    isAdmin(session) ||
    (actorEmail && actorEmail.toLowerCase() === ADMIN_EMAIL.toLowerCase());
  if (!adminAllowed) return;

  const actorKey = String(session?.user?.id || actorEmail || "admin");
  const cached =
    adminThemeCache && adminThemeCache.actorKey === actorKey
      ? adminThemeCache
      : null;
  const isFresh =
    cached &&
    Date.now() - Number(cached.timestamp || 0) <= ADMIN_THEME_CACHE_TTL_MS;

  if (cached && Array.isArray(cached.data)) {
    renderAdminThemes(cached.data);
    if (isFresh && !force) return;
  } else if (!adminThemeLoadInFlight) {
    adminThemesEl.innerHTML = '<div class="status">Loading themes...</div>';
  }

  if (adminThemeLoadInFlight && !force) {
    return adminThemeLoadInFlight;
  }

  const loadPromise = withTimeout(
    (async () => {
      let effectiveSession = session;
      if (!effectiveSession?.access_token) {
        try {
          effectiveSession = await withTimeout(
            getActiveSession(),
            2200,
            "admin themes session restore",
          );
        } catch (error) {
          // fall back to stored token below
        }
      }
      let authToken = String(effectiveSession?.access_token || "");
      if (!authToken) {
        const auth = await getRestAuthContext(effectiveSession || session);
        authToken = String(auth.accessToken || "");
      }
      const sessionLike = authToken
        ? { ...(effectiveSession || session || {}), access_token: authToken }
        : effectiveSession || session || null;

      const fnResult = await callReportFunction(
        "list_admin_themes",
        { limit: ADMIN_THEME_LIMIT },
        sessionLike,
        { onlyPrimary: true, timeoutMs: 1800 },
      );
      if (fnResult.ok && Array.isArray(fnResult?.data?.themes)) {
        const data = fnResult.data.themes;
        adminThemeCache = {
          actorKey,
          timestamp: Date.now(),
          data,
        };
        renderAdminThemes(data);
        return;
      }
      const restFallback = await fetchAdminThemesFromRest(authToken);
      if (!restFallback.error) {
        const data = Array.isArray(restFallback.data) ? restFallback.data : [];
        adminThemeCache = {
          actorKey,
          timestamp: Date.now(),
          data,
        };
        renderAdminThemes(data);
        return;
      }
      throw new Error(
        `Admin themes function failed: ${fnResult.error || "Unknown error."} | REST fallback failed: ${errorMessage(
          restFallback.error,
          "Unknown error.",
        )}`,
      );
    })(),
    9000,
    "admin themes load",
  )
    .catch((error) => {
      if (!cached?.data?.length) {
        adminThemesEl.innerHTML = `<div class="status error">${errorMessage(error, "Failed to load admin themes.")}</div>`;
      } else {
        console.warn("[QuickCanvas] admin themes refresh failed:", error);
      }
    })
    .finally(() => {
      if (adminThemeLoadInFlight === loadPromise) {
        adminThemeLoadInFlight = null;
      }
    });

  adminThemeLoadInFlight = loadPromise;
  return loadPromise;
}

function updateAuthUI(session) {
  const signedIn = Boolean(getEffectiveSession(session));
  syncAuthGateFromSession(session);
  isUiSignedIn = signedIn;
  const unlocked = signedIn;
  signOutBtn.hidden = !signedIn;
  if (!signedIn) {
    setAuthMode(authMode);
  }
  if (authModeSwitch) {
    authModeSwitch.hidden = signedIn;
  }
  if (authEmailField) {
    authEmailField.hidden = signedIn;
  }
  if (authPasswordField) {
    authPasswordField.hidden = signedIn;
  }
  if (authUsernameField) {
    authUsernameField.hidden = signedIn || authMode !== "signup";
  }
  if (authSchoolStartField) {
    authSchoolStartField.hidden = signedIn || authMode !== "signup";
  }
  signInBtn.hidden = signedIn || authMode !== "signin";
  signUpBtn.hidden = signedIn || authMode !== "signup";
  authEmailInput.disabled = signedIn;
  authPasswordInput.disabled = signedIn;
  authUsernameInput.disabled = signedIn;
  if (authSchoolStartInput) {
    authSchoolStartInput.disabled = signedIn;
  }
  if (signedIn) {
    authEmailInput.value = session?.user?.email || "";
    authPasswordInput.value = "";
    if (authSchoolStartInput) {
      authSchoolStartInput.value = formatSchoolStartForInput(
        currentSchoolStartMinutes,
      );
    }
    if (profileEmailInput) {
      profileEmailInput.placeholder = session?.user?.email
        ? `Current: ${session.user.email}`
        : "new-email@example.com";
    }
  } else {
    if (profileEmailInput) {
      profileEmailInput.value = "";
      profileEmailInput.placeholder = "new-email@example.com";
    }
    if (profileCurrentPasswordInput) {
      profileCurrentPasswordInput.value = "";
    }
    if (profileNewPasswordInput) {
      profileNewPasswordInput.value = "";
    }
    if (profileNewPasswordConfirmInput) {
      profileNewPasswordConfirmInput.value = "";
    }
  }
  if (authGateNoticeEl) {
    authGateNoticeEl.textContent = signedIn
      ? "Account is active. Update your profile below or use Sign out to switch accounts."
      : authMode === "signup"
        ? 'Create your account with a username and school start time for the "Next Due Date" filter.'
        : "Sign in with your email and password.";
  }
  if (accountProfilePanel) {
    accountProfilePanel.hidden = !signedIn;
  }
  if (accountCloudPanel) {
    accountCloudPanel.hidden = true;
  }
  if (tokenHelpPanel) {
    tokenHelpPanel.hidden = true;
  }
  tabs.forEach((tab) => {
    if (tab.dataset.tab === "account") {
      tab.hidden = false;
      tab.disabled = false;
      return;
    }
    tab.hidden = !unlocked;
    tab.disabled = !unlocked;
  });
  if (!signedIn) {
    applyThemeToInputs(getDefaultPopupTheme());
    setActiveTab("account");
    setPill(false);
  } else {
    setPill(true);
  }
  if (adminTabBtn) {
    adminTabBtn.hidden = !signedIn || !isAdmin(session);
  }
}

saveBtn.addEventListener(
  "click",
  withUiError("Save settings failed", saveSettings),
);
clearBtn.addEventListener(
  "click",
  withUiError("Clear settings failed", clearSettings),
);
openCanvasBtn.addEventListener(
  "click",
  withUiError("Open Canvas failed", openCanvas),
);
openOptionsBtn.addEventListener(
  "click",
  withUiError("Open options failed", openOptions),
);
enabledInput.addEventListener(
  "change",
  withUiError("Toggle dashboard failed", saveSettings),
);
if (calendarRemindersEnabledInput) {
  calendarRemindersEnabledInput.addEventListener(
    "change",
    withUiError("Toggle reminders failed", saveSettings),
  );
}
saveThemeBtn.addEventListener(
  "click",
  withUiError("Apply theme failed", saveTheme),
);
saveThemeCloudBtn.addEventListener(
  "click",
  withUiError("Cloud save failed", saveThemeToCloud),
);
publishThemeBtn.addEventListener(
  "click",
  withUiError("Publish failed", publishTheme),
);
refreshCloudBtn.addEventListener(
  "click",
  withUiError("Cloud refresh failed", loadCloudThemes),
);
signInBtn.addEventListener("click", withUiError("Sign-in failed", signIn));
signUpBtn.addEventListener("click", withUiError("Sign-up failed", signUp));
signOutBtn.addEventListener("click", withUiError("Sign-out failed", signOut));
if (saveProfileBtn) {
  saveProfileBtn.addEventListener("click", async () => {
    try {
      await saveProfile();
    } catch (error) {
      const msg = `Save profile failed: ${errorMessage(error)}`;
      console.error("[QuickCanvas]", msg, error);
      setProfileStatus(msg, true);
      setStatus(msg, true);
    }
  });
}
if (authModeSignInBtn) {
  authModeSignInBtn.addEventListener("click", () => setAuthMode("signin"));
}
if (authModeSignUpBtn) {
  authModeSignUpBtn.addEventListener("click", () => setAuthMode("signup"));
}
if (authSchoolStartInput) {
  authSchoolStartInput.addEventListener("change", () => {
    const parsed = parseSchoolStartInputValue(authSchoolStartInput.value);
    if (parsed !== null) {
      void saveSchoolStartMinutes(parsed);
    }
  });
}
if (profileSchoolStartInput) {
  profileSchoolStartInput.addEventListener("input", () => {
    const parsed = parseSchoolStartInputValue(profileSchoolStartInput.value);
    if (parsed === null) return;
    void saveSchoolStartMinutes(parsed);
    void persistSchoolStartPreference(parsed, { showStatus: true });
  });
  profileSchoolStartInput.addEventListener("change", () => {
    const parsed = parseSchoolStartInputValue(profileSchoolStartInput.value);
    if (parsed === null) {
      setProfileStatus("Enter a valid school start time.", true);
      return;
    }
    void persistSchoolStartPreference(parsed, { showStatus: true });
  });
}
if (refreshReportsBtn)
  refreshReportsBtn.addEventListener(
    "click",
    withUiError("Reports refresh failed", () =>
      loadAdminReports({ force: true }),
    ),
  );
if (refreshAdminThemesBtn) {
  refreshAdminThemesBtn.addEventListener(
    "click",
    withUiError("Admin themes refresh failed", () =>
      loadAdminThemes({ force: true }),
    ),
  );
}

sortTrendingBtn.addEventListener("click", () => {
  sortTrendingBtn.classList.add("is-active");
  sortLatestBtn.classList.remove("is-active");
  withUiError("Community refresh failed", () =>
    loadCommunityThemes("trending"),
  )();
});
sortLatestBtn.addEventListener("click", () => {
  sortLatestBtn.classList.add("is-active");
  sortTrendingBtn.classList.remove("is-active");
  withUiError("Community refresh failed", () =>
    loadCommunityThemes("latest"),
  )();
});

tabs.forEach((tab) => {
  tab.addEventListener(
    "click",
    withUiError("Tab switch failed", async () => setActiveTab(tab.dataset.tab)),
  );
});

[
  themeModeSelect,
  accentColorInput,
  bgIntensityInput,
  surfaceContrastInput,
  bgColorInput,
  surfaceColorInput,
  surfaceAltColorInput,
  borderColorInput,
  textColorInput,
  mutedColorInput,
  fontBodySelect,
  fontHeadSelect,
  radiusScaleInput,
  shadowStrengthInput,
  customCssInput,
].forEach((input) => {
  input.addEventListener("input", () => {
    if (!isUiSignedIn) return;
    applyTheme(getThemeFromInputs());
  });
});

prefersDark.addEventListener("change", () => {
  if (isUiSignedIn && themeModeSelect.value === "auto") {
    applyTheme(getThemeFromInputs());
  }
});

window.addEventListener("error", (event) => {
  const message = errorMessage(
    event?.error || event?.message,
    "Unexpected script error.",
  );
  setAuthStatus(`Runtime error: ${message}`, true);
});

window.addEventListener("unhandledrejection", (event) => {
  const message = errorMessage(event?.reason, "Unhandled promise rejection.");
  setAuthStatus(`Unhandled error: ${message}`, true);
});

if (supabaseClient) {
  supabaseClient.auth.onAuthStateChange(async (event, session) => {
    try {
      await loadForceSignedOutState();
      let effectiveSession = await enforceForcedSignOut(session);
      if (!effectiveSession && !forceSignedOutState && event !== "SIGNED_OUT") {
        effectiveSession = await getActiveSession();
      }
      if (effectiveSession) {
        await persistAuthTokens(effectiveSession);
      } else if (event === "SIGNED_OUT" || forceSignedOutState) {
        await clearAuthTokens();
      }
      if (!isSessionBootstrapComplete && !effectiveSession) {
        return;
      }
      if (isProfileSaveInProgress) {
        return;
      }
      updateAuthUI(effectiveSession);
      if (effectiveSession) {
        await loadTheme(true);
        await loadProfile();
        updateAuthUI(effectiveSession);
        await setAuthGateState({
          authenticated: true,
          hasUsername: true,
          userId: effectiveSession.user.id,
        });
        await loadCloudThemes();
        await loadCommunityThemes(
          sortLatestBtn.classList.contains("is-active") ? "latest" : "trending",
          { force: true },
        );
        if (isAdmin(effectiveSession)) {
          await Promise.all([
            loadAdminReports({ force: true }),
            loadAdminThemes({ force: true }),
          ]);
        }
      } else {
        await loadTheme(false);
        currentProfileUsername = "";
        cloudSchoolStartMinutes = null;
        await setAuthGateState({ authenticated: false, hasUsername: false });
      }
    } catch (error) {
      setAuthStatus(
        `Auth state sync failed: ${errorMessage(error, "unknown error")}`,
        true,
      );
    }
  });
}

(async () => {
  try {
    await loadForceSignedOutState();
    await loadStoredSchoolStartMinutes();
    setAuthMode("signin");
    setActiveTab(getStoredActiveTab() || "account");
    renderFontOptions();
    renderPresets();
    await loadTheme(false);
    await loadSettings();
    await loadCommunityThemes("trending");

    if (supabaseClient) {
      const session = await getActiveSession();
      if (session) {
        await loadTheme(true);
        setActiveTab(getStoredActiveTab() || "themes");
        await loadProfile();
        await setAuthGateState({
          authenticated: true,
          hasUsername: true,
          userId: session.user.id,
        });
        updateAuthUI(session);
        await loadCloudThemes();
        await loadCommunityThemes("trending", { force: true });
        if (isAdmin(session)) {
          await Promise.all([
            loadAdminReports({ force: true }),
            loadAdminThemes({ force: true }),
          ]);
        }
      } else {
        await loadTheme(false);
        currentProfileUsername = "";
        cloudSchoolStartMinutes = null;
        await setAuthGateState({ authenticated: false, hasUsername: false });
        updateAuthUI(null);
      }
    } else {
      await loadTheme(false);
      currentProfileUsername = "";
      cloudSchoolStartMinutes = null;
      await setAuthGateState({ authenticated: false, hasUsername: false });
      updateAuthUI(null);
      setAuthStatus("Supabase library failed to load.", true);
    }
  } catch (error) {
    setAuthStatus(`Popup init failed: ${errorMessage(error)}`, true);
  } finally {
    isSessionBootstrapComplete = true;
  }
})();
