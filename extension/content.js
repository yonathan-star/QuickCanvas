(() => {
  const scriptInstanceId = `cfe_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  window.__cfeActiveInstanceId = scriptInstanceId;

  function isStaleInstance() {
    return window.__cfeActiveInstanceId !== scriptInstanceId;
  }

  const prefersDark = window.matchMedia("(prefers-color-scheme: dark)");
  let activeTheme = { mode: "auto", accent: "#1f5f8b" };
  let themeListenerBound = false;
  let routeWatcherBound = false;
  let routeRefreshTimer = null;
  let globalNavObserver = null;
  let nonDashboardThemeQueued = false;
  let passiveRouteSyncBound = false;
  let routeWatcherIntervalId = null;
  let passiveRouteSyncIntervalId = null;
  let dashboardLayoutPersistTimer = null;
  let dashboardStorageSyncListener = null;
  let dashboardPageHideListener = null;
  let dashboardVisibilityPersistListener = null;
  const AUTH_GATE_STATE_KEY = "cfeAuthState";
  const SCHOOL_START_MINUTES_KEY = "cfeSchoolStartMinutes";
  const DEFAULT_SCHOOL_START_MINUTES = 8 * 60;
  const POPUP_THEME_MIRROR_KEY = "cfePopupThemeMirror";
  const SNAP_GRID_STEP = 16;
  const SNAP_WIDGET_GAP = 4;
  const FREE_WIDGET_GAP = 6;
  const COURSE_WIDGET_GAP = 6;
  const API_RESPONSE_CACHE_TTL = {
    planner: 45_000,
    courses: 300_000,
    nicknames: 600_000,
    announcements: 120_000,
    discussions: 120_000,
    events: 90_000,
  };
  let schoolStartMinutes = DEFAULT_SCHOOL_START_MINUTES;

  function isDashboardPath(pathname) {
    return (
      pathname === "/" ||
      pathname === "/dashboard" ||
      pathname.startsWith("/dashboard/")
    );
  }

  function isQuizLikePath(pathname) {
    const path = String(pathname || "").toLowerCase();
    return (
      /\/courses\/\d+\/quizzes(\/|$)/.test(path) ||
      /\/courses\/\d+\/assignment_quizzes(\/|$)/.test(path) ||
      /\/assessment_questions(\/|$)/.test(path) ||
      /\/quiz_submission(\/|$)/.test(path) ||
      /\/quiz_submissions(\/|$)/.test(path)
    );
  }

  function syncPageTypeClasses(pathname = window.location.pathname || "/") {
    const isQuizPage = isQuizLikePath(pathname);
    document.documentElement.classList.toggle("cfe-quiz-page", isQuizPage);
    if (document.body) {
      document.body.classList.toggle("cfe-quiz-page", isQuizPage);
      return;
    }
    document.addEventListener(
      "DOMContentLoaded",
      () => {
        document.body?.classList.toggle("cfe-quiz-page", isQuizPage);
      },
      { once: true },
    );
  }

  function isExtensionContextValid() {
    return Boolean(chrome?.runtime?.id);
  }

  function normalizeSchoolStartMinutes(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return DEFAULT_SCHOOL_START_MINUTES;
    const rounded = Math.round(numeric);
    if (rounded < 0 || rounded > 23 * 60 + 59) {
      return DEFAULT_SCHOOL_START_MINUTES;
    }
    return rounded;
  }

  function formatSchoolStartLabel(minutes = schoolStartMinutes) {
    const safe = normalizeSchoolStartMinutes(minutes);
    const hour24 = Math.floor(safe / 60);
    const minute = safe % 60;
    const meridiem = hour24 >= 12 ? "PM" : "AM";
    const hour12 = hour24 % 12 || 12;
    return `${hour12}:${String(minute).padStart(2, "0")} ${meridiem}`;
  }

  async function loadSchoolStartMinutesSetting() {
    if (!isExtensionContextValid()) {
      schoolStartMinutes = DEFAULT_SCHOOL_START_MINUTES;
      return schoolStartMinutes;
    }
    try {
      const stored = await chrome.storage.sync.get(SCHOOL_START_MINUTES_KEY);
      schoolStartMinutes = normalizeSchoolStartMinutes(
        stored?.[SCHOOL_START_MINUTES_KEY],
      );
    } catch (error) {
      schoolStartMinutes = DEFAULT_SCHOOL_START_MINUTES;
    }
    return schoolStartMinutes;
  }

  function toBool(value) {
    if (typeof value === "boolean") return value;
    if (typeof value === "number") return value !== 0;
    if (typeof value === "string") {
      const normalized = value.trim().toLowerCase();
      return (
        normalized === "true" || normalized === "1" || normalized === "yes"
      );
    }
    return false;
  }

  function isAuthStateUnlocked(state) {
    if (!state || typeof state !== "object") return false;
    // New model: authenticated is the source of truth.
    if (Object.prototype.hasOwnProperty.call(state, "authenticated")) {
      return toBool(state.authenticated);
    }
    // Legacy fallback for older saved state shape.
    return (
      toBool(state.hasUsername) && Boolean(String(state.userId || "").trim())
    );
  }

  async function getAuthGateState() {
    if (!isExtensionContextValid()) {
      return { authenticated: false, hasUsername: false, userId: "" };
    }
    try {
      const stored = await chrome.storage.local.get(AUTH_GATE_STATE_KEY);
      const state = stored?.[AUTH_GATE_STATE_KEY] || {};
      if (!isAuthStateUnlocked(state)) {
        try {
          const mirror = await chrome.storage.sync.get("cfeAuthGateMirror");
          const mirroredState = mirror?.cfeAuthGateMirror || {};
          if (isAuthStateUnlocked(mirroredState)) {
            return {
              authenticated: toBool(mirroredState.authenticated),
              hasUsername: true,
              userId: String(mirroredState.userId || ""),
            };
          }
        } catch (error) {
          // ignore sync fallback failures
        }
      }
      return {
        authenticated: toBool(state.authenticated),
        hasUsername: toBool(state.hasUsername),
        userId: String(state.userId || ""),
      };
    } catch (error) {
      return { authenticated: false, hasUsername: false, userId: "" };
    }
  }

  async function isAuthUnlocked() {
    const state = await getAuthGateState();
    return isAuthStateUnlocked(state);
  }

  function bindRouteWatcher() {
    if (routeWatcherBound) return;
    routeWatcherBound = true;
    let lastHandledPath = window.location.pathname;
    let refreshInFlight = false;

    const refreshForRoute = async () => {
      if (isStaleInstance() || refreshInFlight) return;
      refreshInFlight = true;
      try {
        destroy();
        await init();
      } finally {
        refreshInFlight = false;
      }
    };

    const scheduleRouteRefresh = () => {
      if (isStaleInstance()) return;
      const currentPath = window.location.pathname;
      if (currentPath === lastHandledPath) return;
      lastHandledPath = currentPath;
      if (!isDashboardPath(currentPath)) {
        removeDashboardShell();
      }
      if (routeRefreshTimer) {
        clearTimeout(routeRefreshTimer);
      }
      routeRefreshTimer = setTimeout(refreshForRoute, 120);
    };

    window.addEventListener("popstate", scheduleRouteRefresh);
    window.addEventListener("hashchange", scheduleRouteRefresh);

    // Passive watcher for SPA route transitions.
    routeWatcherIntervalId = window.setInterval(() => {
      if (isStaleInstance()) return;
      const currentPath = window.location.pathname;
      if (currentPath !== lastHandledPath) {
        scheduleRouteRefresh();
      }
    }, 300);
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
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
    return escapeHtml(value).replace(/`/g, "&#96;");
  }

  function sanitizeHref(rawUrl) {
    const raw = String(rawUrl || "").trim();
    if (!raw || raw.startsWith("#")) return "#";
    try {
      const url = new URL(raw, window.location.origin);
      const protocol = String(url.protocol || "").toLowerCase();
      if (protocol === "http:" || protocol === "https:") {
        return url.toString();
      }
      return "#";
    } catch (error) {
      return "#";
    }
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

  function rgbToHsl(rgb) {
    if (!rgb) return { h: 210, s: 60, l: 45 };
    const r = rgb.r / 255;
    const g = rgb.g / 255;
    const b = rgb.b / 255;
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h = 0;
    let s = 0;
    const l = (max + min) / 2;
    const d = max - min;
    if (d !== 0) {
      s = d / (1 - Math.abs(2 * l - 1));
      switch (max) {
        case r:
          h = 60 * (((g - b) / d) % 6);
          break;
        case g:
          h = 60 * ((b - r) / d + 2);
          break;
        default:
          h = 60 * ((r - g) / d + 4);
          break;
      }
    }
    if (h < 0) h += 360;
    return { h, s: s * 100, l: l * 100 };
  }

  function hslToHex(h, s, l) {
    const sat = clamp(s, 0, 100) / 100;
    const lig = clamp(l, 0, 100) / 100;
    const c = (1 - Math.abs(2 * lig - 1)) * sat;
    const hh = (h % 360) / 60;
    const x = c * (1 - Math.abs((hh % 2) - 1));
    let r1 = 0;
    let g1 = 0;
    let b1 = 0;
    if (hh >= 0 && hh < 1) {
      r1 = c;
      g1 = x;
    } else if (hh >= 1 && hh < 2) {
      r1 = x;
      g1 = c;
    } else if (hh >= 2 && hh < 3) {
      g1 = c;
      b1 = x;
    } else if (hh >= 3 && hh < 4) {
      g1 = x;
      b1 = c;
    } else if (hh >= 4 && hh < 5) {
      r1 = x;
      b1 = c;
    } else {
      r1 = c;
      b1 = x;
    }
    const m = lig - c / 2;
    const r = Math.round((r1 + m) * 255);
    const g = Math.round((g1 + m) * 255);
    const b = Math.round((b1 + m) * 255);
    return `#${toHex(clamp(r, 0, 255))}${toHex(clamp(g, 0, 255))}${toHex(clamp(b, 0, 255))}`;
  }

  function getSeriesColor(accentHex, index, total, mode, bg, surface, text) {
    const accent = accentHex || "#1f5f8b";
    const darkMode = mode === "dark";
    const base = rgbToHsl(hexToRgb(accent));
    const lightnessSteps = darkMode
      ? [78, 67, 56, 45, 34, 26, 72, 50]
      : [22, 33, 44, 55, 66, 77, 28, 61];
    const saturationSteps = darkMode
      ? [92, 88, 84, 80, 76, 72, 90, 82]
      : [90, 86, 82, 78, 74, 70, 88, 76];
    const hueShift = [0, 6, -6, 12, -12, 18, -18, 24][index % 8];

    let color = hslToHex(
      (base.h + hueShift + 360) % 360,
      saturationSteps[index % saturationSteps.length],
      lightnessSteps[index % lightnessSteps.length],
    );

    // Pull slightly toward accent so it always feels tied to the chosen theme.
    color = mixColors(color, accent, darkMode ? 0.16 : 0.22);

    return ensureMinContrastAcrossSurfaces(
      color,
      [bg || "", surface || ""],
      3.0,
      darkMode ? "dark" : "light",
    );
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
    const bg = theme?.bg || palette.bg;
    const surface = theme?.surface || palette.surface;
    const surfaceAlt = theme?.surfaceAlt || palette.surfaceAlt || surface;
    const surfaces = [bg, surface, surfaceAlt];
    normalized.bg = bg;
    normalized.surface = surface;
    normalized.accent = ensureMinContrastAcrossSurfaces(
      theme?.accent || "#1f5f8b",
      [surface, surfaceAlt],
      3.2,
      mode,
    );
    normalized.text = ensureMinContrastAcrossSurfaces(
      theme?.text || palette.text,
      surfaces,
      4.9,
      mode,
    );
    normalized.muted = ensureMinContrastAcrossSurfaces(
      theme?.muted || palette.muted,
      surfaces,
      3.9,
      mode,
    );
    normalized.border = ensureMinContrastAcrossSurfaces(
      theme?.border || palette.border,
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
        bg: "#0f1318",
        surface: "#151a20",
        surfaceAlt: "#1c232b",
        border: "#2a323c",
        text: "#f1f3f5",
        muted: "#a8b2bb",
      };
    }
    return {
      bg: "#f7f5f0",
      surface: "#ffffff",
      surfaceAlt: "#f8faf9",
      border: "#e6e1d6",
      text: "#1f2a33",
      muted: "#6f7b83",
    };
  }

  function ensureBodyClass() {
    if (document.body) {
      document.body.classList.add("cfe-theme-applied");
      return true;
    }
    return false;
  }

  function setStyleImportant(el, prop, value) {
    if (!el || !value) return;
    el.style.setProperty(prop, value, "important");
  }

  function getRootVar(name, fallback) {
    const value = getComputedStyle(document.documentElement)
      .getPropertyValue(name)
      .trim();
    return value || fallback;
  }

  function parseCssColorToRgb(color) {
    if (!color) return null;
    const trimmed = color.trim();
    if (trimmed.startsWith("#")) return hexToRgb(trimmed);
    const match = trimmed.match(
      /rgba?\(\s*([0-9.]+)\s*,\s*([0-9.]+)\s*,\s*([0-9.]+)(?:\s*,\s*([0-9.]+))?\s*\)/i,
    );
    if (!match) return null;
    return {
      r: clamp(Math.round(Number(match[1])), 0, 255),
      g: clamp(Math.round(Number(match[2])), 0, 255),
      b: clamp(Math.round(Number(match[3])), 0, 255),
    };
  }

  function rgbToHexString(rgb) {
    if (!rgb) return null;
    return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`;
  }

  function parseCssColorToRgba(color) {
    if (!color) return null;
    const trimmed = color.trim();
    if (trimmed === "transparent") return { r: 0, g: 0, b: 0, a: 0 };
    if (trimmed.startsWith("#")) {
      const rgb = hexToRgb(trimmed);
      return rgb ? { ...rgb, a: 1 } : null;
    }
    const match = trimmed.match(
      /rgba?\(\s*([0-9.]+)\s*,\s*([0-9.]+)\s*,\s*([0-9.]+)(?:\s*,\s*([0-9.]+))?\s*\)/i,
    );
    if (!match) return null;
    return {
      r: clamp(Math.round(Number(match[1])), 0, 255),
      g: clamp(Math.round(Number(match[2])), 0, 255),
      b: clamp(Math.round(Number(match[3])), 0, 255),
      a: clamp(match[4] === undefined ? 1 : Number(match[4]), 0, 1),
    };
  }

  function getEffectiveBackgroundHex(element) {
    let current = element;
    while (current && current !== document.documentElement) {
      const bg = parseCssColorToRgba(getComputedStyle(current).backgroundColor);
      if (bg && bg.a > 0.05) return rgbToHexString(bg);
      current = current.parentElement;
    }
    const bodyBg = parseCssColorToRgba(
      getComputedStyle(document.body || document.documentElement)
        .backgroundColor,
    );
    if (bodyBg && bodyBg.a > 0.05) return rgbToHexString(bodyBg);
    return getRootVar("--cfe-bg", "#f7f5f0");
  }

  function paintGlobalNavColors() {
    const nav = document.getElementById("global_nav");
    if (!nav) return;

    const isDark = document.body?.classList.contains("cfe-theme-dark");
    const navText = isDark ? "#f8fafc" : "#1f2a33";
    const navActiveText = isDark ? "#ffffff" : "#111827";

    setStyleImportant(nav, "color", navText);
    setStyleImportant(nav, "-webkit-text-fill-color", navText);
    setStyleImportant(nav, "text-shadow", "none");
    setStyleImportant(nav, "opacity", "1");

    const links = nav.querySelectorAll(
      '.ic-app-header__menu-list-link, [id^="global_nav_"][id$="_link"], #primaryNavToggle',
    );

    const knownIds = [
      "global_nav_profile_link",
      "global_nav_dashboard_link",
      "global_nav_courses_link",
      "global_nav_calendar_link",
      "global_nav_conversations_link",
      "global_nav_history_link",
      "global_nav_help_link",
      "primaryNavToggle",
    ];

    knownIds.forEach((id) => {
      const item = document.getElementById(id);
      if (!item) return;
      const itemActive = Boolean(
        item.closest(
          ".ic-app-header__menu-list-item--active, .menu-item--active",
        ),
      );
      const itemColor = itemActive ? navActiveText : navText;
      setStyleImportant(item, "color", itemColor);
      setStyleImportant(item, "-webkit-text-fill-color", itemColor);
      setStyleImportant(item, "opacity", "1");
      item.querySelectorAll("*").forEach((node) => {
        setStyleImportant(node, "color", itemColor);
        setStyleImportant(node, "-webkit-text-fill-color", itemColor);
        setStyleImportant(node, "fill", "currentColor");
        setStyleImportant(node, "stroke", "currentColor");
        setStyleImportant(node, "opacity", "1");
        setStyleImportant(node, "text-shadow", "none");
        setStyleImportant(node, "filter", "none");
      });
    });

    links.forEach((link) => {
      const isActive = Boolean(
        link.closest(
          ".ic-app-header__menu-list-item--active, .menu-item--active",
        ),
      );
      const targetColor = isActive ? navActiveText : navText;
      const parentItem = link.closest(
        ".ic-app-header__menu-list-item, .menu-item",
      );
      if (parentItem) {
        setStyleImportant(parentItem, "opacity", "1");
        setStyleImportant(parentItem, "color", targetColor);
      }
      setStyleImportant(link, "color", targetColor);
      setStyleImportant(link, "-webkit-text-fill-color", targetColor);
      setStyleImportant(link, "text-shadow", "none");
      setStyleImportant(link, "opacity", "1");

      const descendants = link.querySelectorAll(
        ".menu-item__text, .menu-item-icon-container, .menu-item-icon-container span, .ic-icon-svg, svg, svg path, svg g, svg use, .menu-item__badge, .menu-item__badge *",
      );
      descendants.forEach((node) => {
        setStyleImportant(node, "color", targetColor);
        setStyleImportant(node, "-webkit-text-fill-color", targetColor);
        setStyleImportant(node, "fill", "currentColor");
        setStyleImportant(node, "stroke", "currentColor");
        setStyleImportant(node, "opacity", "1");
        setStyleImportant(node, "text-shadow", "none");
        setStyleImportant(node, "filter", "none");
      });
    });
  }

  function paintBreadcrumbPathColors() {
    const isDark = document.body?.classList.contains("cfe-theme-dark");
    const muted = getRootVar("--cfe-muted", isDark ? "#a8b2bb" : "#6f7b83");
    const text = getRootVar("--cfe-text", isDark ? "#f1f3f5" : "#1f2a33");

    const pathNodes = document.querySelectorAll(
      [
        ".ic-app-nav-toggle-and-crumbs li > a",
        ".ic-app-nav-toggle-and-crumbs li > a > span",
        ".ic-app-nav-toggle-and-crumbs li > a .ellipsible",
        ".ic-app-crumbs li > a",
        ".ic-app-crumbs li > a > span",
        ".ic-app-crumbs li > a .ellipsible",
        "#breadcrumbs li > a",
        "#breadcrumbs li > a > span",
        "#breadcrumbs li > a .ellipsible",
        ".breadcrumbs li > a",
        ".breadcrumbs li > a > span",
        ".breadcrumbs li > a .ellipsible",
        'nav[aria-label*="breadcrumb" i] li > a',
        'nav[aria-label*="breadcrumb" i] li > a > span',
        'nav[aria-label*="breadcrumb" i] li > a .ellipsible',
      ].join(", "),
    );

    pathNodes.forEach((node) => {
      setStyleImportant(node, "color", muted);
      setStyleImportant(node, "-webkit-text-fill-color", muted);
      setStyleImportant(node, "opacity", "1");
      setStyleImportant(node, "text-shadow", "none");
      setStyleImportant(node, "filter", "none");
    });

    const activePathNodes = document.querySelectorAll(
      [
        ".ic-app-crumbs__crumb--current",
        ".ic-app-crumbs__crumb--current a",
        ".ic-app-crumbs__crumb--current span",
        ".ic-app-nav-toggle-and-crumbs li[aria-current='page'] > a",
        ".ic-app-nav-toggle-and-crumbs li[aria-current='page'] > a > span",
        ".ic-app-nav-toggle-and-crumbs li[aria-current='page'] > a .ellipsible",
        "#breadcrumbs li[aria-current='page'] > a",
        "#breadcrumbs li[aria-current='page'] > a > span",
        "#breadcrumbs li[aria-current='page'] > a .ellipsible",
        ".breadcrumbs li[aria-current='page'] > a",
        ".breadcrumbs li[aria-current='page'] > a > span",
        ".breadcrumbs li[aria-current='page'] > a .ellipsible",
      ].join(", "),
    );

    activePathNodes.forEach((node) => {
      setStyleImportant(node, "color", text);
      setStyleImportant(node, "-webkit-text-fill-color", text);
      setStyleImportant(node, "opacity", "1");
      setStyleImportant(node, "text-shadow", "none");
      setStyleImportant(node, "filter", "none");
    });

    // Hard fallback for generic Canvas path items:
    // <li><a ...><span class="ellipsible">...</span></a></li>
    const genericPathSpans = document.querySelectorAll("li > a > .ellipsible");
    genericPathSpans.forEach((span) => {
      const li = span.closest("li");
      const isCurrent = Boolean(
        li?.matches("[aria-current='page'], .ic-app-crumbs__crumb--current"),
      );
      const color = isCurrent ? text : muted;
      if (li) {
        setStyleImportant(li, "color", color);
        setStyleImportant(li, "-webkit-text-fill-color", color);
        setStyleImportant(li, "opacity", "1");
      }
      setStyleImportant(span, "color", color);
      setStyleImportant(span, "-webkit-text-fill-color", color);
      setStyleImportant(span, "opacity", "1");
      setStyleImportant(span, "text-shadow", "none");
      setStyleImportant(span, "filter", "none");
      setStyleImportant(span, "mix-blend-mode", "normal");
      const a = span.closest("a");
      if (a) {
        setStyleImportant(a, "color", color);
        setStyleImportant(a, "-webkit-text-fill-color", color);
        setStyleImportant(a, "opacity", "1");
        setStyleImportant(a, "mix-blend-mode", "normal");
      }
    });
  }

  function bindGlobalNavColorObserver() {
    if (globalNavObserver) return;
    globalNavObserver = new MutationObserver(() => {
      paintGlobalNavColors();
      paintBreadcrumbPathColors();
    });
    globalNavObserver.observe(document.documentElement, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["class", "style", "aria-current"],
    });
  }

  async function waitForBody(timeoutMs = 5000) {
    if (document.body) return document.body;
    return new Promise((resolve) => {
      let settled = false;
      let observer = null;
      const settle = (value) => {
        if (settled) return;
        settled = true;
        if (observer) observer.disconnect();
        resolve(value);
      };

      const timer = setTimeout(() => {
        settle(document.body || null);
      }, timeoutMs);

      observer = new MutationObserver(() => {
        if (!document.body) return;
        clearTimeout(timer);
        settle(document.body);
      });

      observer.observe(document.documentElement, {
        childList: true,
        subtree: true,
      });
    });
  }

  function resetTheme() {
    document.documentElement.classList.remove("cfe-theme-applied");
    if (document.body) {
      document.body.classList.remove("cfe-theme-applied", "cfe-theme-dark");
    }
    const styleTag = document.getElementById("cfe-custom-theme");
    if (styleTag) {
      styleTag.remove();
    }
  }

  function applyPopupTheme(theme) {
    syncPageTypeClasses();
    const mode = resolveThemeMode(theme?.mode || "auto");
    const palette = getDefaultPalette(mode);
    const safeTheme = normalizeThemeReadability(theme || {}, mode, palette);
    const accent = safeTheme.accent || "#1f5f8b";
    const bgIntensity = Number(safeTheme.bgIntensity ?? 40);
    const surfaceContrast = Number(safeTheme.surfaceContrast ?? 50);
    const bg = safeTheme.bg || palette.bg;
    const surface = safeTheme.surface || palette.surface;
    const text = safeTheme.text || palette.text;
    const muted = safeTheme.muted || palette.muted;
    const border = safeTheme.border || palette.border;
    const fontBody = safeTheme.fontBody || "Space Grotesk";
    const fontHead = safeTheme.fontHead || "Fraunces";
    const radius = Number(safeTheme.radius ?? 12);
    const shadow = Number(safeTheme.shadow ?? 45);
    const customCss = safeTheme.customCss || "";

    activeTheme = {
      mode: safeTheme.mode || "auto",
      accent,
      bgIntensity,
      surfaceContrast,
      bg,
      surface,
      surfaceAlt: safeTheme.surfaceAlt || "",
      border,
      text,
      muted,
      fontBody,
      fontHead,
      radius,
      shadow,
      customCss,
    };

    document.documentElement.classList.add("cfe-theme-applied");
    if (document.body) {
      document.body.classList.toggle("cfe-theme-dark", mode === "dark");
      document.body.classList.add("cfe-theme-applied");
    } else {
      document.addEventListener(
        "DOMContentLoaded",
        () => {
          document.body?.classList.toggle("cfe-theme-dark", mode === "dark");
          ensureBodyClass();
        },
        { once: true },
      );
    }

    const rootStyle = document.documentElement.style;
    rootStyle.setProperty("--cfe-accent", accent);
    rootStyle.setProperty("--cfe-bg", bg);
    rootStyle.setProperty("--cfe-surface", surface);
    rootStyle.setProperty("--cfe-text", text);
    rootStyle.setProperty("--cfe-muted", muted);
    rootStyle.setProperty("--cfe-border", border);
    rootStyle.setProperty(
      "--cfe-accent-soft",
      mode === "dark"
        ? mixColors(accent, "#ffffff", 0.2)
        : mixColors(accent, "#ffffff", 0.85),
    );
    rootStyle.setProperty(
      "--cfe-accent-strong",
      mode === "dark"
        ? mixColors(accent, "#000000", 0.35)
        : mixColors(accent, "#000000", 0.25),
    );
    rootStyle.setProperty("--cfe-accent-contrast", getContrastColor(accent));
    rootStyle.setProperty("--cfe-divider", mixColors(border, surface, 0.35));
    rootStyle.setProperty("--cfe-muted-strong", mixColors(muted, text, 0.3));
    rootStyle.setProperty(
      "--cfe-hover",
      mode === "dark"
        ? mixColors(surface, "#ffffff", 0.06)
        : mixColors(surface, "#000000", 0.03),
    );
    const navBg =
      mode === "dark"
        ? mixColors(surface, bg, 0.16)
        : mixColors(surface, bg, 0.28);
    const navText = ensureMinContrast(text, navBg, 7, mode);
    const navMuted = ensureMinContrast(muted, navBg, 4.8, mode);
    const navHoverBg =
      mode === "dark"
        ? mixColors(navBg, "#ffffff", 0.1)
        : mixColors(navBg, "#000000", 0.06);
    const navActiveBg =
      mode === "dark"
        ? mixColors(accent, navBg, 0.72)
        : mixColors(accent, navBg, 0.84);
    const navActiveText = ensureMinContrast(navText, navActiveBg, 4.8, mode);
    const navDivider = ensureMinContrast(border, navBg, 1.7, mode);
    rootStyle.setProperty("--cfe-nav-bg", navBg);
    rootStyle.setProperty("--cfe-nav-text", navText);
    rootStyle.setProperty("--cfe-nav-muted", navMuted);
    rootStyle.setProperty("--cfe-nav-hover-bg", navHoverBg);
    rootStyle.setProperty("--cfe-nav-active-bg", navActiveBg);
    rootStyle.setProperty("--cfe-nav-active-text", navActiveText);
    rootStyle.setProperty("--cfe-nav-divider", navDivider);
    rootStyle.setProperty("--cfe-font-body", fontBody);
    rootStyle.setProperty("--cfe-font-head", fontHead);

    const radiusValue = clamp(radius, 6, 22);
    rootStyle.setProperty("--cfe-radius", `${radiusValue}px`);
    rootStyle.setProperty("--cfe-radius-lg", `${radiusValue + 6}px`);
    rootStyle.setProperty(
      "--cfe-radius-sm",
      `${Math.max(6, radiusValue - 2)}px`,
    );

    const intensity = clamp(bgIntensity, 0, 100) / 100;
    rootStyle.setProperty(
      "--cfe-bg-glow-1",
      rgbaFromHex(accent, 0.18 * intensity),
    );
    rootStyle.setProperty(
      "--cfe-bg-glow-2",
      rgbaFromHex(accent, 0.1 * intensity),
    );

    const contrast = clamp(surfaceContrast, 0, 100) / 100;
    const surfaceBase = surface;
    const surfaceAlt =
      safeTheme.surfaceAlt ||
      (mode === "dark"
        ? mixColors(surfaceBase, bg, 0.28 + 0.5 * (1 - contrast))
        : mixColors(surfaceBase, bg, 0.16 + 0.35 * (1 - contrast)));
    rootStyle.setProperty("--cfe-surface-alt", surfaceAlt);

    const shadowStrength = clamp(shadow, 0, 100) / 100;
    const baseAlpha = mode === "dark" ? 0.5 : 0.28;
    rootStyle.setProperty(
      "--cfe-shadow",
      `0 18px 42px rgba(26, 35, 46, ${baseAlpha * shadowStrength})`,
    );
    rootStyle.setProperty(
      "--cfe-shadow-side",
      `-10px 0 22px rgba(26, 35, 46, ${0.22 * shadowStrength})`,
    );

    let styleTag = document.getElementById("cfe-custom-theme");
    if (!styleTag) {
      styleTag = document.createElement("style");
      styleTag.id = "cfe-custom-theme";
      const target =
        document.head || document.documentElement || document.body || null;
      if (target) {
        target.appendChild(styleTag);
      } else {
        document.addEventListener(
          "DOMContentLoaded",
          () => {
            (
              document.head ||
              document.documentElement ||
              document.body
            )?.appendChild(styleTag);
          },
          { once: true },
        );
      }
    }
    styleTag.textContent = customCss;
    bindGlobalNavColorObserver();
    paintGlobalNavColors();
    paintBreadcrumbPathColors();
    requestAnimationFrame(() => paintGlobalNavColors());
    requestAnimationFrame(() => paintBreadcrumbPathColors());
    setTimeout(() => paintGlobalNavColors(), 250);
    setTimeout(() => paintBreadcrumbPathColors(), 250);
  }

  function applyPopupThemeLite(theme) {
    syncPageTypeClasses();
    const mode = resolveThemeMode(theme?.mode || "auto");
    const palette = getDefaultPalette(mode);
    const safeTheme = normalizeThemeReadability(theme || {}, mode, palette);
    const accent = safeTheme.accent || "#1f5f8b";
    const bgIntensity = Number(safeTheme.bgIntensity ?? 40);
    const surfaceContrast = Number(safeTheme.surfaceContrast ?? 50);
    const bg = safeTheme.bg || palette.bg;
    const surface = safeTheme.surface || palette.surface;
    const text = safeTheme.text || palette.text;
    const muted = safeTheme.muted || palette.muted;
    const border = safeTheme.border || palette.border;
    const fontBody = safeTheme.fontBody || "Space Grotesk";
    const fontHead = safeTheme.fontHead || "Fraunces";
    const radius = Number(safeTheme.radius ?? 12);
    const shadow = Number(safeTheme.shadow ?? 45);

    document.documentElement.classList.add("cfe-theme-applied");
    if (document.body) {
      document.body.classList.toggle("cfe-theme-dark", mode === "dark");
      document.body.classList.add("cfe-theme-applied");
    } else {
      document.addEventListener(
        "DOMContentLoaded",
        () => {
          document.body?.classList.toggle("cfe-theme-dark", mode === "dark");
          ensureBodyClass();
        },
        { once: true },
      );
    }

    const rootStyle = document.documentElement.style;
    rootStyle.setProperty("--cfe-accent", accent);
    rootStyle.setProperty("--cfe-bg", bg);
    rootStyle.setProperty("--cfe-surface", surface);
    rootStyle.setProperty("--cfe-text", text);
    rootStyle.setProperty("--cfe-muted", muted);
    rootStyle.setProperty("--cfe-border", border);
    rootStyle.setProperty(
      "--cfe-accent-soft",
      mode === "dark"
        ? mixColors(accent, "#ffffff", 0.2)
        : mixColors(accent, "#ffffff", 0.85),
    );
    rootStyle.setProperty(
      "--cfe-accent-strong",
      mode === "dark"
        ? mixColors(accent, "#000000", 0.35)
        : mixColors(accent, "#000000", 0.25),
    );
    rootStyle.setProperty("--cfe-accent-contrast", getContrastColor(accent));
    rootStyle.setProperty("--cfe-divider", mixColors(border, surface, 0.35));
    rootStyle.setProperty("--cfe-muted-strong", mixColors(muted, text, 0.3));
    rootStyle.setProperty(
      "--cfe-hover",
      mode === "dark"
        ? mixColors(surface, "#ffffff", 0.06)
        : mixColors(surface, "#000000", 0.03),
    );
    const navBg =
      mode === "dark"
        ? mixColors(surface, bg, 0.16)
        : mixColors(surface, bg, 0.28);
    const navText = ensureMinContrast(text, navBg, 7, mode);
    const navMuted = ensureMinContrast(muted, navBg, 4.8, mode);
    const navHoverBg =
      mode === "dark"
        ? mixColors(navBg, "#ffffff", 0.1)
        : mixColors(navBg, "#000000", 0.06);
    const navActiveBg =
      mode === "dark"
        ? mixColors(accent, navBg, 0.72)
        : mixColors(accent, navBg, 0.84);
    const navActiveText = ensureMinContrast(navText, navActiveBg, 4.8, mode);
    const navDivider = ensureMinContrast(border, navBg, 1.7, mode);
    rootStyle.setProperty("--cfe-nav-bg", navBg);
    rootStyle.setProperty("--cfe-nav-text", navText);
    rootStyle.setProperty("--cfe-nav-muted", navMuted);
    rootStyle.setProperty("--cfe-nav-hover-bg", navHoverBg);
    rootStyle.setProperty("--cfe-nav-active-bg", navActiveBg);
    rootStyle.setProperty("--cfe-nav-active-text", navActiveText);
    rootStyle.setProperty("--cfe-nav-divider", navDivider);
    rootStyle.setProperty("--cfe-font-body", fontBody);
    rootStyle.setProperty("--cfe-font-head", fontHead);

    const radiusValue = clamp(radius, 6, 22);
    rootStyle.setProperty("--cfe-radius", `${radiusValue}px`);
    rootStyle.setProperty("--cfe-radius-lg", `${radiusValue + 6}px`);
    rootStyle.setProperty(
      "--cfe-radius-sm",
      `${Math.max(6, radiusValue - 2)}px`,
    );

    const intensity = clamp(bgIntensity, 0, 100) / 100;
    rootStyle.setProperty(
      "--cfe-bg-glow-1",
      rgbaFromHex(accent, 0.18 * intensity),
    );
    rootStyle.setProperty(
      "--cfe-bg-glow-2",
      rgbaFromHex(accent, 0.1 * intensity),
    );

    const contrast = clamp(surfaceContrast, 0, 100) / 100;
    const surfaceBase = surface;
    const surfaceAlt =
      safeTheme.surfaceAlt ||
      (mode === "dark"
        ? mixColors(surfaceBase, bg, 0.28 + 0.5 * (1 - contrast))
        : mixColors(surfaceBase, bg, 0.16 + 0.35 * (1 - contrast)));
    rootStyle.setProperty("--cfe-surface-alt", surfaceAlt);

    const shadowStrength = clamp(shadow, 0, 100) / 100;
    const baseAlpha = mode === "dark" ? 0.5 : 0.28;
    rootStyle.setProperty(
      "--cfe-shadow",
      `0 18px 42px rgba(26, 35, 46, ${baseAlpha * shadowStrength})`,
    );
    rootStyle.setProperty(
      "--cfe-shadow-side",
      `-10px 0 22px rgba(26, 35, 46, ${0.22 * shadowStrength})`,
    );

    // Never apply custom CSS off-dashboard to avoid breaking Canvas interactions.
    const styleTag = document.getElementById("cfe-custom-theme");
    if (styleTag) {
      styleTag.remove();
    }
  }

  async function loadPopupTheme() {
    if (!isExtensionContextValid()) return;
    try {
      const { popupTheme } = await chrome.storage.sync.get("popupTheme");
      applyPopupTheme(popupTheme || {});
    } catch (error) {
      // ignore transient reload invalidation
    }
  }

  async function readThemeForCurrentPage() {
    if (!isExtensionContextValid()) {
      return { canvasSettings: null, popupTheme: null };
    }
    const [syncStored, localStored] = await Promise.all([
      chrome.storage.sync
        .get(["canvasSettings", "popupTheme"])
        .catch(() => ({ canvasSettings: null, popupTheme: null })),
      chrome.storage.local
        .get(POPUP_THEME_MIRROR_KEY)
        .catch(() => ({ [POPUP_THEME_MIRROR_KEY]: null })),
    ]);
    const syncTheme = syncStored?.popupTheme || null;
    const mirrorPayload = localStored?.[POPUP_THEME_MIRROR_KEY] || null;
    const mirrorTheme =
      mirrorPayload && typeof mirrorPayload === "object"
        ? mirrorPayload.theme || null
        : null;
    const syncUpdatedAt = Number(syncTheme?.updatedAt || 0);
    const mirrorUpdatedAt = Number(mirrorPayload?.updatedAt || 0);
    const popupTheme =
      mirrorTheme && mirrorUpdatedAt > syncUpdatedAt ? mirrorTheme : syncTheme;
    return {
      canvasSettings: syncStored?.canvasSettings || null,
      popupTheme: popupTheme || null,
    };
  }

  async function syncThemeForCurrentPage() {
    if (!isExtensionContextValid()) return;
    syncPageTypeClasses();
    if (!(await isAuthUnlocked())) {
      resetTheme();
      return;
    }
    const { canvasSettings, popupTheme } = await readThemeForCurrentPage();
    if (!(canvasSettings?.enabled ?? true)) {
      resetTheme();
      return;
    }
    const baseOrigin = (() => {
      try {
        return new URL(canvasSettings.baseUrl || "").origin;
      } catch (error) {
        return "";
      }
    })();
    if (!baseOrigin || baseOrigin !== window.location.origin) {
      resetTheme();
      return;
    }
    if (isDashboardPath(window.location.pathname || "/")) {
      applyPopupTheme(popupTheme || {});
      return;
    }
    const safeTheme = { ...(popupTheme || {}), customCss: "" };
    applyPopupThemeLite(safeTheme);
  }

  async function preloadTheme() {
    try {
      await syncThemeForCurrentPage();
    } catch (error) {
      // ignore preload failures
    }
  }

  async function applyThemeForCurrentOrigin() {
    await syncThemeForCurrentPage();
  }

  function queueNonDashboardThemeApply() {
    if (nonDashboardThemeQueued) return;
    nonDashboardThemeQueued = true;
    const run = () => {
      setTimeout(() => {
        applyThemeForCurrentOrigin().catch(() => {
          // ignore transient reload invalidation
        });
      }, 700);
    };
    if (document.readyState === "complete") {
      run();
      return;
    }
    window.addEventListener("load", run, { once: true });
  }

  function bindPassiveRouteSync() {
    if (passiveRouteSyncBound) return;
    passiveRouteSyncBound = true;
    let lastPath = window.location.pathname || "/";
    let lastThemeSyncAt = 0;

    const sync = () => {
      if (isStaleInstance()) return;
      const path = window.location.pathname || "/";
      syncPageTypeClasses(path);
      const now = Date.now();
      const shouldRefreshTheme =
        document.visibilityState === "visible" &&
        now - lastThemeSyncAt > 15_000;
      if (document.getElementById("cfe-auth-wall")) {
        ensureAuthWall();
      }
      const changed = path !== lastPath;
      if (changed) {
        lastPath = path;
      }

      if (!isDashboardPath(path)) {
        if (changed) {
          cleanupNonDashboardUi();
          lastThemeSyncAt = now;
          applyThemeForCurrentOrigin().catch(() => {
            // ignore transient reload invalidation
          });
          renderCourseDueWidgetForCurrentPage().catch(() => {
            // ignore non-dashboard widget failures
          });
          window.__cfeInjected = false;
          return;
        }
        if (!document.documentElement.classList.contains("cfe-theme-applied")) {
          lastThemeSyncAt = now;
          applyThemeForCurrentOrigin().catch(() => {
            // ignore transient reload invalidation
          });
        } else if (shouldRefreshTheme) {
          lastThemeSyncAt = now;
          applyThemeForCurrentOrigin().catch(() => {
            // ignore transient reload invalidation
          });
        }
        if (
          isCourseHomePath(path) &&
          !document.getElementById("cfe-course-due-widget")
        ) {
          renderCourseDueWidgetForCurrentPage().catch(() => {
            // ignore non-dashboard widget failures
          });
        }
        return;
      }

      if (shouldRefreshTheme) {
        lastThemeSyncAt = now;
        applyThemeForCurrentOrigin().catch(() => {
          // ignore transient reload invalidation
        });
      }

      if (changed && !document.getElementById("cfe-dashboard")) {
        window.__cfeInjected = false;
        init();
      }
    };

    window.addEventListener("popstate", () => setTimeout(sync, 60));
    window.addEventListener("hashchange", () => setTimeout(sync, 60));
    passiveRouteSyncIntervalId = window.setInterval(sync, 700);
  }

  function destroy() {
    teardownDashboardRuntimeBindings();
    if (routeWatcherIntervalId) {
      clearInterval(routeWatcherIntervalId);
      routeWatcherIntervalId = null;
    }
    if (passiveRouteSyncIntervalId) {
      clearInterval(passiveRouteSyncIntervalId);
      passiveRouteSyncIntervalId = null;
    }
    if (dashboardLayoutPersistTimer) {
      clearTimeout(dashboardLayoutPersistTimer);
      dashboardLayoutPersistTimer = null;
    }
    if (resumeThemeSyncTimer) {
      clearTimeout(resumeThemeSyncTimer);
      resumeThemeSyncTimer = null;
    }
    clearResumeThemeBurstTimers();
    if (courseCanvasBlocksObserver) {
      try {
        courseCanvasBlocksObserver.disconnect();
      } catch (error) {
        // ignore observer teardown failures
      }
      courseCanvasBlocksObserver = null;
    }
    document.querySelectorAll("#cfe-dashboard").forEach((el) => el.remove());
    document.querySelectorAll("#cfe-sidebar").forEach((el) => el.remove());

    if (document.body) {
      document.body.classList.remove(
        "cfe-dashboard-only",
        "cfe-has-sidebar",
        "cfe-sidebar-collapsed",
      );
    }
    resetTheme();
    if (routeRefreshTimer) {
      clearTimeout(routeRefreshTimer);
      routeRefreshTimer = null;
    }
    window.__cfeInjected = false;
  }

  function teardownDashboardRuntimeBindings() {
    if (dashboardStorageSyncListener && isExtensionContextValid()) {
      try {
        chrome.storage.onChanged.removeListener(dashboardStorageSyncListener);
      } catch (error) {
        // ignore listener cleanup failures
      }
    }
    dashboardStorageSyncListener = null;

    if (dashboardPageHideListener) {
      window.removeEventListener("pagehide", dashboardPageHideListener);
      dashboardPageHideListener = null;
    }
    if (dashboardVisibilityPersistListener) {
      document.removeEventListener(
        "visibilitychange",
        dashboardVisibilityPersistListener,
      );
      dashboardVisibilityPersistListener = null;
    }
  }

  function removeAllInjectedContainers() {
    document.querySelectorAll("#cfe-dashboard").forEach((el) => el.remove());
    document.querySelectorAll("#cfe-sidebar").forEach((el) => el.remove());
  }

  function removeDashboardShell() {
    teardownDashboardRuntimeBindings();
    document.querySelectorAll("#cfe-dashboard").forEach((el) => el.remove());
    if (document.body) {
      document.body.classList.remove("cfe-dashboard-only");
    }
  }

  function cleanupNonDashboardUi() {
    teardownDashboardRuntimeBindings();
    if (courseCanvasBlocksObserver) {
      try {
        courseCanvasBlocksObserver.disconnect();
      } catch (error) {
        // ignore disconnect errors
      }
      courseCanvasBlocksObserver = null;
    }
    removeAllInjectedContainers();
    document
      .querySelectorAll("#cfe-course-due-widget, #cfe-course-widget-board")
      .forEach((el) => el.remove());
    const strip = () => {
      if (!document.body) return;
      document.body.classList.remove(
        "cfe-dashboard-only",
        "cfe-has-sidebar",
        "cfe-sidebar-collapsed",
      );
    };
    strip();
    if (!document.body) {
      document.addEventListener(
        "DOMContentLoaded",
        () => {
          removeAllInjectedContainers();
          strip();
        },
        { once: true },
      );
    }
  }

  function removeAuthWall() {
    document.getElementById("cfe-auth-wall")?.remove();
  }

  function ensureAuthWall() {
    const path = window.location.pathname || "/";
    if (!isDashboardPath(path) && !isCourseHomePath(path)) {
      removeAuthWall();
      return;
    }
    const body = document.body;
    if (!body) return;
    if (document.getElementById("cfe-auth-wall")) return;
    const wall = document.createElement("div");
    wall.id = "cfe-auth-wall";
    wall.innerHTML = `
      <div class="cfe-auth-wall-card">
        <h2>QuickCanvas sign-in required</h2>
        <p>Create an account or sign in to use dashboard widgets and layout tools.</p>
        <button type="button" class="cfe-auth-wall-open">Open QuickCanvas Account</button>
      </div>
    `;
    const openBtn = wall.querySelector(".cfe-auth-wall-open");
    openBtn?.addEventListener("click", () => {
      try {
        window.open(chrome.runtime.getURL("popup.html"), "_blank", "noopener");
      } catch (error) {
        // ignore popup open failures
      }
    });
    body.appendChild(wall);
  }

  function getCourseIdFromPath(pathname) {
    const match = String(pathname || "").match(/^\/courses\/(\d+)(?:\/|$)/);
    return match ? match[1] : "";
  }

  function isCourseHomePath(pathname) {
    const path = String(pathname || "");
    return /^\/courses\/\d+\/?$/.test(path);
  }

  function getNextDueDateWindow(baseDate = new Date()) {
    const anchor = baseDate instanceof Date ? new Date(baseDate) : new Date();
    const start = new Date(anchor);
    start.setHours(0, 0, 0, 0);
    const end = new Date(start);
    end.setDate(end.getDate() + 1);
    const safeMinutes = normalizeSchoolStartMinutes(schoolStartMinutes);
    const hour = Math.floor(safeMinutes / 60);
    const minute = safeMinutes % 60;
    end.setHours(hour, minute, 0, 0);
    return { start, end };
  }

  function getNextDueDateWindowFromItems(
    items,
    referenceDate = new Date(),
    mapDueDate,
  ) {
    const ref =
      referenceDate instanceof Date ? new Date(referenceDate) : new Date();
    if (Number.isNaN(ref.getTime())) return null;
    let nextDue = null;
    (Array.isArray(items) ? items : []).forEach((item) => {
      const candidateRaw =
        typeof mapDueDate === "function" ? mapDueDate(item) : item?.due_at;
      const candidate =
        candidateRaw instanceof Date ? candidateRaw : new Date(candidateRaw);
      if (Number.isNaN(candidate.getTime())) return;
      if (candidate < ref) return;
      if (!nextDue || candidate < nextDue) {
        nextDue = candidate;
      }
    });
    if (!nextDue) return null;
    const { start, end } = getNextDueDateWindow(nextDue);
    return { start, end, anchor: nextDue };
  }

  function normalizeCourseHomeWidgetPrefs(raw) {
    const valid = new Set(["duecourse", "coursequicklinks", "coursecalendar"]);
    const fallbackPositions = {
      duecourse: { left: 0, top: 0 },
      coursequicklinks: { left: 0, top: 312 },
      coursecalendar: { left: 0, top: 520 },
    };
    const fallbackSizes = {
      duecourse: { width: 320, height: 286 },
      coursequicklinks: { width: 320, height: 182 },
      coursecalendar: { width: 320, height: 360 },
    };
    const toFiniteNumber = (value, fallback = 0) => {
      if (typeof value === "number" && Number.isFinite(value)) {
        return value;
      }
      const parsed = Number.parseFloat(String(value ?? ""));
      return Number.isFinite(parsed) ? parsed : fallback;
    };
    const normalizeLayout = (layoutRaw) => {
      let order = Array.isArray(layoutRaw?.order)
        ? layoutRaw.order.filter((id) => valid.has(id))
        : ["duecourse"];
      order = order.filter((id, index) => order.indexOf(id) === index);
      const positionsRaw =
        layoutRaw?.positions && typeof layoutRaw.positions === "object"
          ? layoutRaw.positions
          : {};
      const sizesRaw =
        layoutRaw?.sizes && typeof layoutRaw.sizes === "object"
          ? layoutRaw.sizes
          : {};
      const positions = {};
      const sizes = {};
      valid.forEach((id) => {
        const fallbackPos = fallbackPositions[id] || { left: 0, top: 0 };
        const posRaw =
          positionsRaw[id] && typeof positionsRaw[id] === "object"
            ? positionsRaw[id]
            : {};
        positions[id] = {
          left: Math.max(
            0,
            Math.round(toFiniteNumber(posRaw.left, fallbackPos.left)),
          ),
          top: Math.max(
            0,
            Math.round(toFiniteNumber(posRaw.top, fallbackPos.top)),
          ),
        };

        const fallbackSize = fallbackSizes[id] || { width: 260, height: 180 };
        const sizeRaw =
          sizesRaw[id] && typeof sizesRaw[id] === "object" ? sizesRaw[id] : {};
        sizes[id] = {
          width: Math.max(
            1,
            Math.round(toFiniteNumber(sizeRaw.width, fallbackSize.width)),
          ),
          height: Math.max(
            1,
            Math.round(toFiniteNumber(sizeRaw.height, fallbackSize.height)),
          ),
        };
      });
      return { order, positions, sizes };
    };
    const layoutsByCourse =
      raw?.layoutsByCourse && typeof raw.layoutsByCourse === "object"
        ? Object.fromEntries(
            Object.entries(raw.layoutsByCourse).map(([courseId, layout]) => [
              courseId,
              normalizeLayout(layout),
            ]),
          )
        : {};
    const globalLayout =
      raw?.layout && typeof raw.layout === "object"
        ? normalizeLayout(raw.layout)
        : null;
    const hiddenCanvasByCourse =
      raw?.hiddenCanvasByCourse && typeof raw.hiddenCanvasByCourse === "object"
        ? raw.hiddenCanvasByCourse
        : {};
    return { layout: globalLayout, layoutsByCourse, hiddenCanvasByCourse };
  }

  function getDefaultCourseHomeLayout() {
    return {
      order: ["duecourse"],
      positions: {
        duecourse: { left: 0, top: 0 },
        coursequicklinks: { left: 0, top: 312 },
        coursecalendar: { left: 0, top: 520 },
      },
      sizes: {
        duecourse: { width: 320, height: 286 },
        coursequicklinks: { width: 320, height: 182 },
        coursecalendar: { width: 320, height: 360 },
      },
    };
  }

  function getCourseHomeLayout(prefs, courseId) {
    const defaults = getDefaultCourseHomeLayout();
    const layout = prefs.layout || prefs.layoutsByCourse?.[courseId] || {};
    const merged = {
      order: Array.isArray(layout.order) ? layout.order : defaults.order,
      positions: {
        ...defaults.positions,
        ...(layout.positions || {}),
      },
      sizes: {
        ...defaults.sizes,
        ...(layout.sizes || {}),
      },
    };
    merged.order = merged.order.filter(
      (id, index) => merged.order.indexOf(id) === index,
    );
    return merged;
  }

  async function loadCourseHomeWidgetPrefs() {
    if (!isExtensionContextValid()) return normalizeCourseHomeWidgetPrefs({});
    try {
      const { cfeCourseHomeWidgets } = await chrome.storage.sync.get(
        "cfeCourseHomeWidgets",
      );
      return normalizeCourseHomeWidgetPrefs(cfeCourseHomeWidgets);
    } catch (error) {
      return normalizeCourseHomeWidgetPrefs({});
    }
  }

  async function saveCourseHomeWidgetPrefs(prefs) {
    if (!isExtensionContextValid()) return;
    try {
      await chrome.storage.sync.set({
        cfeCourseHomeWidgets: normalizeCourseHomeWidgetPrefs(prefs),
      });
    } catch (error) {
      // ignore transient reload invalidation
    }
  }

  let courseCanvasBlocksObserver = null;

  function getCanvasBlockLegacyKey(el) {
    const idPart = el.id ? `id:${el.id}` : "";
    const dataTestId = el.getAttribute("data-testid") || "";
    const dataContext = el.getAttribute("data-context-view") || "";
    const classPart = (el.className || "")
      .toString()
      .trim()
      .split(/\s+/)
      .slice(0, 3)
      .join(".");
    const rolePart =
      el.getAttribute("aria-label") || el.getAttribute("role") || "";
    const titlePart = (
      el.querySelector("h1,h2,h3,h4,strong,.header")?.textContent || ""
    )
      .trim()
      .slice(0, 40);
    return [
      idPart,
      `testid:${dataTestId}`,
      `ctx:${dataContext}`,
      classPart,
      rolePart,
      titlePart,
    ]
      .filter(Boolean)
      .join("|");
  }

  function getCanvasBlockKey(el) {
    const idPart = el.id ? `id:${el.id}` : "";
    const dataTestId = el.getAttribute("data-testid") || "";
    const dataContext = el.getAttribute("data-context-view") || "";
    const classPart = (el.className || "")
      .toString()
      .trim()
      .split(/\s+/)
      .slice(0, 3)
      .join(".");
    const rolePart =
      el.getAttribute("aria-label") || el.getAttribute("role") || "";
    const tagPart = (el.tagName || "").toLowerCase();
    return [
      idPart,
      `testid:${dataTestId}`,
      `ctx:${dataContext}`,
      classPart,
      rolePart,
      tagPart,
    ]
      .filter(Boolean)
      .join("|");
  }

  function collectCanvasBlocks(host) {
    const children = Array.from(host.children || []);
    return children
      .filter((el) => el instanceof HTMLElement)
      .filter((el) => el.id !== "cfe-course-widget-board")
      .map((el, index) => {
        const key = getCanvasBlockKey(el);
        const legacyKey = getCanvasBlockLegacyKey(el);
        const label =
          (el.querySelector("h1,h2,h3,h4,strong,.header")?.textContent || "")
            .trim()
            .slice(0, 48) ||
          el.id ||
          (el.className || "Canvas block").toString().split(" ")[0] ||
          `Block ${index + 1}`;
        return { el, key, legacyKey, label };
      });
  }

  function applyHiddenCanvasBlocks(host, hiddenSet) {
    collectCanvasBlocks(host).forEach((item) => {
      if (hiddenSet.has(item.key) || hiddenSet.has(item.legacyKey)) {
        item.el.style.setProperty("display", "none", "important");
        item.el.setAttribute("data-cfe-canvas-hidden", "1");
      } else if (item.el.getAttribute("data-cfe-canvas-hidden") === "1") {
        item.el.style.removeProperty("display");
        item.el.removeAttribute("data-cfe-canvas-hidden");
      }
    });
  }

  function renderCourseWidgetBoard(host, prefs, courseId) {
    if (!host) return null;
    if (courseCanvasBlocksObserver) {
      try {
        courseCanvasBlocksObserver.disconnect();
      } catch (error) {
        // ignore disconnect errors
      }
      courseCanvasBlocksObserver = null;
    }
    const existing = document.getElementById("cfe-course-widget-board");
    if (existing) existing.remove();
    const hiddenByCourse = prefs.hiddenCanvasByCourse || {};
    const hiddenSet = new Set(
      Array.isArray(hiddenByCourse[courseId]) ? hiddenByCourse[courseId] : [],
    );
    applyHiddenCanvasBlocks(host, hiddenSet);
    const layout = getCourseHomeLayout(prefs, courseId);
    const minSizes = {
      duecourse: { width: 180, height: 210 },
      coursequicklinks: { width: 240, height: 220 },
      coursecalendar: { width: 280, height: 380 },
    };

    const board = document.createElement("section");
    board.id = "cfe-course-widget-board";
    board.className = "cfe-card";
    board.innerHTML = `
      <div class="cfe-course-widget-head">
        <h3>Course Widgets</h3>
        <button class="cfe-course-edit-layout" type="button">Edit layout</button>
      </div>
      <div class="cfe-course-widget-dock" hidden>
        <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-course-widget-add="duecourse">+ Due For This Class</button>
        <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-course-widget-add="coursequicklinks">+ Course Quick Links</button>
        <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-course-widget-add="coursecalendar">+ Course Calendar</button>
      </div>
      <div class="cfe-course-canvas-blocks" hidden></div>
      <div class="cfe-course-widget-canvas">
        <section id="cfe-course-due-widget" class="cfe-card cfe-course-home-widget" data-course-widget-id="duecourse"${
          layout.order.includes("duecourse") ? "" : " hidden"
        }>
          <button class="cfe-widget-drag-handle" type="button" data-course-widget-drag="duecourse" aria-label="Drag widget">Drag</button>
          <button class="cfe-widget-trash" type="button" data-course-widget-remove="duecourse" aria-label="Delete widget">Delete</button>
          <span class="cfe-widget-resize cfe-widget-resize-e" data-course-widget-resize="duecourse" data-course-widget-resize-dir="e" aria-hidden="true"></span>
          <span class="cfe-widget-resize cfe-widget-resize-s" data-course-widget-resize="duecourse" data-course-widget-resize-dir="s" aria-hidden="true"></span>
          <span class="cfe-widget-resize cfe-widget-resize-se" data-course-widget-resize="duecourse" data-course-widget-resize-dir="se" aria-hidden="true"></span>
          <div class="cfe-course-due-head">
            <h3>Due For This Class</h3>
            <p>Shows the next upcoming assignment due for this class.</p>
          </div>
          <div class="cfe-course-due-list cfe-list" data-state="loading">
            <div class="cfe-loading">
              <span class="cfe-spinner"></span>
              <span>Loading class assignments...</span>
            </div>
          </div>
        </section>
        <section class="cfe-card cfe-course-home-widget" data-course-widget-id="coursequicklinks"${
          layout.order.includes("coursequicklinks") ? "" : " hidden"
        }>
          <button class="cfe-widget-drag-handle" type="button" data-course-widget-drag="coursequicklinks" aria-label="Drag widget">Drag</button>
          <button class="cfe-widget-trash" type="button" data-course-widget-remove="coursequicklinks" aria-label="Delete widget">Delete</button>
          <span class="cfe-widget-resize cfe-widget-resize-e" data-course-widget-resize="coursequicklinks" data-course-widget-resize-dir="e" aria-hidden="true"></span>
          <span class="cfe-widget-resize cfe-widget-resize-s" data-course-widget-resize="coursequicklinks" data-course-widget-resize-dir="s" aria-hidden="true"></span>
          <span class="cfe-widget-resize cfe-widget-resize-se" data-course-widget-resize="coursequicklinks" data-course-widget-resize-dir="se" aria-hidden="true"></span>
          <div class="cfe-course-due-head">
            <h3>Course Quick Links</h3>
          </div>
          <div class="cfe-quick-links">
            <a class="cfe-quick-link" href="/courses/${courseId}">Home</a>
            <a class="cfe-quick-link" href="/courses/${courseId}/modules">Modules</a>
            <a class="cfe-quick-link" href="/courses/${courseId}/assignments">Assignments</a>
            <a class="cfe-quick-link" href="/courses/${courseId}/grades">Grades</a>
            <a class="cfe-quick-link" href="/courses/${courseId}/syllabus">Syllabus</a>
            <a class="cfe-quick-link" href="/courses/${courseId}/discussion_topics">Discussions</a>
          </div>
        </section>
        <section class="cfe-card cfe-course-home-widget" data-course-widget-id="coursecalendar"${
          layout.order.includes("coursecalendar") ? "" : " hidden"
        }>
          <button class="cfe-widget-drag-handle" type="button" data-course-widget-drag="coursecalendar" aria-label="Drag widget">Drag</button>
          <button class="cfe-widget-trash" type="button" data-course-widget-remove="coursecalendar" aria-label="Delete widget">Delete</button>
          <span class="cfe-widget-resize cfe-widget-resize-e" data-course-widget-resize="coursecalendar" data-course-widget-resize-dir="e" aria-hidden="true"></span>
          <span class="cfe-widget-resize cfe-widget-resize-s" data-course-widget-resize="coursecalendar" data-course-widget-resize-dir="s" aria-hidden="true"></span>
          <span class="cfe-widget-resize cfe-widget-resize-se" data-course-widget-resize="coursecalendar" data-course-widget-resize-dir="se" aria-hidden="true"></span>
          <div class="cfe-course-due-head">
            <h3>Course Calendar</h3>
            <p>Click a day to view assignments due.</p>
          </div>
          <div class="cfe-course-calendar-root"></div>
        </section>
        <div class="cfe-course-widget-empty" hidden>
          No course widgets selected. Click <strong>Edit layout</strong> to add widgets.
        </div>
      </div>
    `;
    host.prepend(board);

    const editBtn = board.querySelector(".cfe-course-edit-layout");
    const dock = board.querySelector(".cfe-course-widget-dock");
    const canvasBlocksEl = board.querySelector(".cfe-course-canvas-blocks");
    const addButtons = Array.from(
      board.querySelectorAll("[data-course-widget-add]"),
    );
    const dueCard = board.querySelector("#cfe-course-due-widget");
    const canvasEl = board.querySelector(".cfe-course-widget-canvas");
    const emptyStateEl = board.querySelector(".cfe-course-widget-empty");
    const widgets = Array.from(
      board.querySelectorAll(".cfe-course-home-widget[data-course-widget-id]"),
    ).filter((el) => el instanceof HTMLElement);
    let editMode = false;
    let dragState = null;
    let resizeState = null;

    const saveLayout = async () => {
      prefs.layout = layout;
      await saveCourseHomeWidgetPrefs(prefs);
    };

    const getWidgetEl = (id) =>
      board.querySelector(
        `.cfe-course-home-widget[data-course-widget-id="${id}"]`,
      );

    const toFiniteNumber = (value, fallback = 0) => {
      if (typeof value === "number" && Number.isFinite(value)) {
        return value;
      }
      const parsed = Number.parseFloat(String(value ?? ""));
      return Number.isFinite(parsed) ? parsed : fallback;
    };

    const getWidgetContentMinimums = (widget, id) => {
      const base = minSizes[id] || { width: 220, height: 140 };
      return {
        minWidth: Math.round(base.width),
        minHeight: Math.round(base.height),
      };
    };

    const applyLayoutStyles = (activeWidgetId = "") => {
      if (!editMode) {
        widgets.forEach((widget) => {
          widget.style.position = "relative";
          widget.style.left = "";
          widget.style.top = "";
          widget.style.width = "100%";
          widget.style.height = "auto";
        });
        if (canvasEl instanceof HTMLElement) {
          canvasEl.style.minHeight = "0px";
        }
        return;
      }

      const canvasWidth = Math.max(220, canvasEl?.clientWidth || 340);
      const hardMaxWidth = Math.max(180, canvasWidth - 12);
      let maxBottom = 0;
      const placed = [];
      const orderedWidgets = [...widgets].sort((a, b) => {
        const aId = a.getAttribute("data-course-widget-id") || "";
        const bId = b.getAttribute("data-course-widget-id") || "";
        if (activeWidgetId) {
          if (aId === activeWidgetId && bId !== activeWidgetId) return -1;
          if (bId === activeWidgetId && aId !== activeWidgetId) return 1;
        }
        const aTop = toFiniteNumber(layout.positions[aId]?.top, 0);
        const bTop = toFiniteNumber(layout.positions[bId]?.top, 0);
        if (aTop !== bTop) return aTop - bTop;
        const aLeft = toFiniteNumber(layout.positions[aId]?.left, 0);
        const bLeft = toFiniteNumber(layout.positions[bId]?.left, 0);
        return aLeft - bLeft;
      });
      orderedWidgets.forEach((widget) => {
        const id = widget.getAttribute("data-course-widget-id") || "";
        if (!id) return;
        if (!layout.order.includes(id)) return;
        const pos = layout.positions[id] || { left: 0, top: 0 };
        const size = layout.sizes[id] ||
          minSizes[id] || { width: 260, height: 180 };
        const min = getWidgetContentMinimums(widget, id);
        const effectiveMinWidth = Math.min(min.minWidth, hardMaxWidth);
        const posLeft = toFiniteNumber(pos.left, 0);
        const posTop = toFiniteNumber(pos.top, 0);
        const sizeWidth = toFiniteNumber(size.width, effectiveMinWidth);
        const sizeHeight = toFiniteNumber(size.height, min.minHeight);
        const maxWidth = Math.min(
          hardMaxWidth,
          Math.max(effectiveMinWidth, canvasWidth - Math.max(0, posLeft)),
        );
        const width = Math.max(
          effectiveMinWidth,
          Math.min(sizeWidth, maxWidth),
        );
        const height = Math.max(min.minHeight, sizeHeight);
        const maxLeft = Math.max(0, canvasWidth - width);
        const left = Math.min(maxLeft, Math.max(0, posLeft));
        let top = Math.max(0, posTop);

        // Prevent widgets from overlapping: push current widget below prior ones.
        let guard = 0;
        while (guard < 50) {
          let overlapBottom = -1;
          placed.forEach((r) => {
            const xOverlap = left < r.right && left + width > r.left;
            const yOverlap = top < r.bottom && top + height > r.top;
            if (xOverlap && yOverlap) {
              overlapBottom = Math.max(overlapBottom, r.bottom);
            }
          });
          if (overlapBottom < 0) break;
          const nextTop = overlapBottom + COURSE_WIDGET_GAP;
          if (nextTop <= top) break;
          top = nextTop;
          guard += 1;
        }

        layout.positions[id] = { left, top };
        layout.sizes[id] = { width, height };
        widget.style.position = "absolute";
        widget.style.left = `${left}px`;
        widget.style.top = `${top}px`;
        widget.style.width = `${width}px`;
        widget.style.height = `${height}px`;
        maxBottom = Math.max(maxBottom, top + height);
        placed.push({
          left,
          top,
          right: left + width,
          bottom: top + height,
        });
      });
      if (canvasEl instanceof HTMLElement) {
        canvasEl.style.minHeight = `${Math.max(280, maxBottom + 16)}px`;
      }
    };

    const refreshBoardState = () => {
      let visibleCount = 0;
      widgets.forEach((widget) => {
        const id = widget.getAttribute("data-course-widget-id") || "";
        const visible = layout.order.includes(id);
        widget.hidden = !visible;
        if (visible) visibleCount += 1;
      });
      board.classList.toggle("cfe-layout-edit-mode", editMode);
      board.classList.toggle(
        "cfe-no-course-widgets",
        !editMode && visibleCount === 0,
      );
      if (dock instanceof HTMLElement) {
        dock.hidden = !editMode;
      }
      if (canvasBlocksEl instanceof HTMLElement) {
        canvasBlocksEl.hidden = !editMode;
      }
      if (emptyStateEl instanceof HTMLElement) {
        emptyStateEl.hidden = editMode || visibleCount > 0;
      }
      addButtons.forEach((button) => {
        const id = button.getAttribute("data-course-widget-add") || "";
        if (button instanceof HTMLButtonElement) {
          button.disabled = layout.order.includes(id);
        }
      });
      // Always apply layout styles first so UI cannot remain in stale absolute mode.
      applyLayoutStyles();
      if (canvasBlocksEl instanceof HTMLElement) {
        try {
          const blocks = collectCanvasBlocks(host);
          const rows = blocks
            .map((item) => {
              const hidden =
                hiddenSet.has(item.key) || hiddenSet.has(item.legacyKey);
              return `
                <div class="cfe-canvas-block-row">
                  <span>${escapeHtml(item.label)}</span>
                  <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-canvas-block-toggle="${encodeURIComponent(
                    item.key,
                  )}">
                    ${hidden ? "Restore" : "Trash"}
                  </button>
                </div>
              `;
            })
            .join("");
          canvasBlocksEl.innerHTML = `
            <div class="cfe-course-canvas-title">Canvas Blocks</div>
            ${rows || '<div class="cfe-widget-empty">No right-side blocks found.</div>'}
          `;
        } catch (error) {
          canvasBlocksEl.innerHTML = `
            <div class="cfe-course-canvas-title">Canvas Blocks</div>
            <div class="cfe-widget-empty">Canvas blocks unavailable on this page.</div>
          `;
        }
      }
    };

    const setEditMode = (enabled) => {
      editMode = Boolean(enabled);
      if (editBtn instanceof HTMLButtonElement) {
        editBtn.classList.toggle("is-active", editMode);
        editBtn.textContent = editMode ? "Done" : "Edit layout";
      }
      refreshBoardState();
    };

    editBtn?.addEventListener("click", () => {
      setEditMode(!editMode);
    });

    addButtons.forEach((button) => {
      button.addEventListener("click", async () => {
        const id = button.getAttribute("data-course-widget-add") || "";
        if (!id || layout.order.includes(id)) return;
        layout.order.push(id);
        if (!layout.positions[id]) {
          const defaults = getDefaultCourseHomeLayout();
          layout.positions[id] = {
            ...(defaults.positions[id] || { left: 0, top: 0 }),
          };
        }
        if (!layout.sizes[id]) {
          const defaults = getDefaultCourseHomeLayout();
          layout.sizes[id] = {
            ...(defaults.sizes[id] || { width: 280, height: 180 }),
          };
        }
        await saveLayout();
        refreshBoardState();
        renderCourseDueWidgetForCurrentPage({
          keepEditMode: editMode,
        }).catch(() => {
          // ignore widget refresh failures
        });
      });
    });

    board.addEventListener("pointerdown", (event) => {
      if (!editMode) return;
      const target = event.target;
      if (!(target instanceof Element)) return;
      const dragHandle = target.closest("[data-course-widget-drag]");
      if (dragHandle instanceof HTMLElement) {
        const id = dragHandle.getAttribute("data-course-widget-drag") || "";
        const widget = getWidgetEl(id);
        if (!(widget instanceof HTMLElement) || !canvasEl) return;
        const widgetRect = widget.getBoundingClientRect();
        dragState = {
          id,
          pointerId: event.pointerId,
          offsetX: event.clientX - widgetRect.left,
          offsetY: event.clientY - widgetRect.top,
        };
        widget.classList.add("is-dragging");
        dragHandle.setPointerCapture(event.pointerId);
        event.preventDefault();
        return;
      }
      const resizeHandle = target.closest("[data-course-widget-resize]");
      if (resizeHandle instanceof HTMLElement) {
        const id = resizeHandle.getAttribute("data-course-widget-resize") || "";
        const dir =
          resizeHandle.getAttribute("data-course-widget-resize-dir") || "se";
        const widget = getWidgetEl(id);
        if (!(widget instanceof HTMLElement)) return;
        const rect = widget.getBoundingClientRect();
        resizeState = {
          id,
          dir,
          pointerId: event.pointerId,
          startX: event.clientX,
          startY: event.clientY,
          startWidth: rect.width,
          startHeight: rect.height,
        };
        resizeHandle.setPointerCapture(event.pointerId);
        widget.classList.add("is-resizing");
        event.preventDefault();
      }
    });

    board.addEventListener("pointermove", (event) => {
      if (dragState && canvasEl) {
        const widget = getWidgetEl(dragState.id);
        if (!(widget instanceof HTMLElement)) return;
        const min = getWidgetContentMinimums(widget, dragState.id);
        const size = layout.sizes[dragState.id] || {
          width: min.minWidth,
          height: min.minHeight,
        };
        const canvasRect = canvasEl.getBoundingClientRect();
        const maxLeft = Math.max(0, canvasRect.width - size.width);
        const left = Math.min(
          maxLeft,
          Math.max(0, event.clientX - canvasRect.left - dragState.offsetX),
        );
        const top = Math.max(
          0,
          event.clientY - canvasRect.top - dragState.offsetY,
        );
        layout.positions[dragState.id] = {
          left: Math.round(left),
          top: Math.round(top),
        };
        applyLayoutStyles(dragState.id);
        return;
      }
      if (resizeState && canvasEl) {
        const id = resizeState.id;
        const dir = resizeState.dir || "se";
        const widget = getWidgetEl(id);
        const min = getWidgetContentMinimums(widget, id);
        const canvasWidth = Math.max(220, canvasEl.clientWidth || 340);
        const hardMaxWidth = Math.max(180, canvasWidth - 12);
        const pos = layout.positions[id] || { left: 0, top: 0 };
        const effectiveMinWidth = Math.min(min.minWidth, hardMaxWidth);
        const maxWidth = Math.min(
          hardMaxWidth,
          Math.max(effectiveMinWidth, canvasWidth - pos.left),
        );
        let width =
          resizeState.startWidth + (event.clientX - resizeState.startX);
        let height =
          resizeState.startHeight + (event.clientY - resizeState.startY);
        if (dir.includes("e")) {
          width = Math.max(effectiveMinWidth, Math.min(width, maxWidth));
        } else {
          width = resizeState.startWidth;
        }
        if (dir.includes("s")) {
          height = Math.max(min.minHeight, height);
        } else {
          height = resizeState.startHeight;
        }
        layout.sizes[id] = {
          width: Math.round(width),
          height: Math.round(height),
        };
        applyLayoutStyles(id);
      }
    });

    const endPointerEdit = async () => {
      if (dragState) {
        const widget = getWidgetEl(dragState.id);
        if (widget instanceof HTMLElement)
          widget.classList.remove("is-dragging");
        dragState = null;
        await saveLayout();
      }
      if (resizeState) {
        const widget = getWidgetEl(resizeState.id);
        if (widget instanceof HTMLElement)
          widget.classList.remove("is-resizing");
        resizeState = null;
        await saveLayout();
      }
    };

    board.addEventListener("pointerup", () => {
      endPointerEdit();
    });
    board.addEventListener("pointercancel", () => {
      endPointerEdit();
    });

    board.addEventListener("click", async (event) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const canvasToggle = target.closest("[data-canvas-block-toggle]");
      if (canvasToggle instanceof HTMLElement) {
        const rawKey = canvasToggle.getAttribute("data-canvas-block-toggle");
        if (!rawKey) return;
        let key = rawKey;
        try {
          key = decodeURIComponent(rawKey);
        } catch (error) {
          key = rawKey;
        }
        if (!key) return;
        const match = collectCanvasBlocks(host).find(
          (item) => item.key === key,
        );
        const isHiddenNow =
          hiddenSet.has(key) ||
          (match ? hiddenSet.has(match.legacyKey) : false);
        if (isHiddenNow) {
          hiddenSet.delete(key);
          if (match?.legacyKey) hiddenSet.delete(match.legacyKey);
        } else {
          hiddenSet.add(key);
        }
        prefs.hiddenCanvasByCourse = {
          ...(prefs.hiddenCanvasByCourse || {}),
          [courseId]: Array.from(hiddenSet),
        };
        await saveCourseHomeWidgetPrefs(prefs);
        applyHiddenCanvasBlocks(host, hiddenSet);
        refreshBoardState();
        return;
      }
      const removeBtn = target.closest("[data-course-widget-remove]");
      if (removeBtn instanceof HTMLElement) {
        const id = removeBtn.getAttribute("data-course-widget-remove");
        if (!id) return;
        layout.order = layout.order.filter((item) => item !== id);
        await saveLayout();
        refreshBoardState();
      }
    });

    refreshBoardState();
    courseCanvasBlocksObserver = new MutationObserver(() => {
      applyHiddenCanvasBlocks(host, hiddenSet);
    });
    courseCanvasBlocksObserver.observe(host, {
      childList: true,
      subtree: true,
    });
    return {
      board,
      dueCard: dueCard instanceof HTMLElement ? dueCard : null,
      dueList:
        dueCard instanceof HTMLElement
          ? dueCard.querySelector(".cfe-course-due-list")
          : null,
      calendarCard:
        board.querySelector(
          '.cfe-course-home-widget[data-course-widget-id="coursecalendar"]',
        ) || null,
      hasDueWidget: layout.order.includes("duecourse"),
      setEditMode,
    };
  }

  async function loadCourseDueChecks() {
    if (!isExtensionContextValid()) return {};
    try {
      const { cfeCourseDueChecks } =
        await chrome.storage.sync.get("cfeCourseDueChecks");
      return cfeCourseDueChecks && typeof cfeCourseDueChecks === "object"
        ? cfeCourseDueChecks
        : {};
    } catch (error) {
      return {};
    }
  }

  async function saveCourseDueChecks(checks) {
    if (!isExtensionContextValid()) return;
    try {
      await chrome.storage.sync.set({ cfeCourseDueChecks: checks || {} });
    } catch (error) {
      // ignore transient reload invalidation
    }
  }

  function getCourseDueCheckKey(courseId, item) {
    const itemId = Number(item?.id || 0);
    if (courseId && itemId) return `${courseId}:${itemId}`;
    const fallbackUrl = String(item?.html_url || item?.url || "");
    return `${courseId || "course"}:${fallbackUrl || item?.name || "item"}`;
  }

  function renderCourseDueWidgetItems(card, items, courseId, checksMap) {
    if (!card) return;
    const listEl = card.querySelector(".cfe-course-due-list");
    if (!listEl) return;
    if (!items.length) {
      listEl.innerHTML =
        '<div class="cfe-widget-empty">No upcoming assignments.</div>';
      return;
    }
    const rows = items
      .map((item) => {
        const title = String(item.name || "Untitled");
        const safeTitle = escapeHtml(title);
        const url = String(item.html_url || item.url || "#");
        const safeUrl = escapeAttr(sanitizeHref(url));
        const due = item.due_at ? new Date(item.due_at) : null;
        const dueText = due
          ? due.toLocaleString(undefined, {
              month: "short",
              day: "numeric",
              hour: "numeric",
              minute: "2-digit",
            })
          : "No due date";
        const safeDueText = escapeHtml(dueText);
        const checkKey = getCourseDueCheckKey(courseId, item);
        const checkKeyAttr = encodeURIComponent(checkKey);
        const checked = Boolean(checksMap?.[checkKey]);
        return `
          <div class="cfe-task-row${checked ? " is-complete" : ""}">
            <label class="cfe-task-check">
              <input type="checkbox" data-course-due-check="${checkKeyAttr}" ${
                checked ? "checked" : ""
              } />
              <span></span>
            </label>
            <div class="cfe-task-body">
              <a class="cfe-task-title" href="${safeUrl}">${safeTitle}</a>
              <div class="cfe-task-meta">Due ${safeDueText}</div>
            </div>
          </div>
        `;
      })
      .join("");
    listEl.innerHTML = rows;
  }

  function toLocalDayKey(dateValue) {
    const d = new Date(dateValue);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function renderCourseCalendarWidget(card, assignments) {
    if (!(card instanceof HTMLElement)) return;
    const root = card.querySelector(".cfe-course-calendar-root");
    if (!(root instanceof HTMLElement)) return;

    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    const gridStart = new Date(monthStart);
    gridStart.setDate(monthStart.getDate() - monthStart.getDay());

    const dueByDay = {};
    (assignments || [])
      .filter((item) => item && item.due_at)
      .forEach((item) => {
        const key = toLocalDayKey(item.due_at);
        if (!dueByDay[key]) dueByDay[key] = [];
        dueByDay[key].push(item);
      });

    let selectedKey = card.getAttribute("data-cfe-cal-selected") || "";
    if (!selectedKey) {
      selectedKey = toLocalDayKey(now);
    }
    card.setAttribute("data-cfe-cal-selected", selectedKey);

    const days = [];
    for (let i = 0; i < 42; i += 1) {
      const d = new Date(gridStart);
      d.setDate(gridStart.getDate() + i);
      const key = toLocalDayKey(d);
      const inMonth = d >= monthStart && d <= monthEnd;
      const isToday = key === toLocalDayKey(now);
      const isSelected = key === selectedKey;
      const dueCount = (dueByDay[key] || []).length;
      days.push(`
        <button class="cfe-course-cal-day${inMonth ? "" : " is-out"}${isToday ? " is-today" : ""}${isSelected ? " is-selected" : ""}" type="button" data-cfe-cal-day="${key}">
          <span class="cfe-course-cal-daynum">${d.getDate()}</span>
          ${dueCount ? `<span class="cfe-course-cal-dot">${dueCount}</span>` : ""}
        </button>
      `);
    }

    const selectedItems = (dueByDay[selectedKey] || [])
      .slice()
      .sort((a, b) => new Date(a.due_at || 0) - new Date(b.due_at || 0));
    const detailRows = selectedItems.length
      ? selectedItems
          .map((item) => {
            const due = new Date(item.due_at);
            const time = due.toLocaleTimeString([], {
              hour: "numeric",
              minute: "2-digit",
            });
            const title = escapeHtml(String(item.name || "Untitled"));
            const href = escapeAttr(
              sanitizeHref(String(item.html_url || item.url || "#")),
            );
            const safeTime = escapeHtml(time);
            return `<div class="cfe-course-cal-item"><a class="cfe-task-title" href="${href}">${title}</a><div class="cfe-task-meta">Due ${safeTime}</div></div>`;
          })
          .join("")
      : '<div class="cfe-widget-empty">No assignments due on this day.</div>';

    root.innerHTML = `
      <div class="cfe-course-cal-head">${monthStart.toLocaleDateString([], { month: "long", year: "numeric" })}</div>
      <div class="cfe-course-cal-week">
        <span>Sun</span><span>Mon</span><span>Tue</span><span>Wed</span><span>Thu</span><span>Fri</span><span>Sat</span>
      </div>
      <div class="cfe-course-cal-grid">${days.join("")}</div>
      <div class="cfe-course-cal-detail">${detailRows}</div>
    `;

    root.querySelectorAll("[data-cfe-cal-day]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const key = btn.getAttribute("data-cfe-cal-day") || "";
        if (!key) return;
        card.setAttribute("data-cfe-cal-selected", key);
        renderCourseCalendarWidget(card, assignments);
      });
    });
  }

  async function renderCourseDueWidgetForCurrentPage(options = {}) {
    if (!isExtensionContextValid()) return;
    if (!isCourseHomePath(window.location.pathname || "")) return;
    const courseId = getCourseIdFromPath(window.location.pathname || "");
    if (!courseId) return;

    const { canvasSettings } = await chrome.storage.sync.get("canvasSettings");
    if (!canvasSettings?.enabled) return;
    const baseOrigin = (() => {
      try {
        return new URL(canvasSettings.baseUrl || "").origin;
      } catch (error) {
        return "";
      }
    })();
    if (!baseOrigin || baseOrigin !== window.location.origin) return;
    const apiToken = String(canvasSettings.apiToken || "").trim();
    if (!apiToken) return;

    let host = null;
    for (let attempt = 0; attempt < 20; attempt += 1) {
      host =
        document.querySelector("#right-side") ||
        document.querySelector(".right-side") ||
        document.querySelector(".ic-Layout-aside") ||
        document.querySelector(".ic-Layout-aside-right") ||
        document.querySelector(".ic-Layout-sidebar") ||
        document.querySelector(".ic-Layout-aside__content");
      if (host instanceof Element) break;
      // Canvas often mounts right-side async after initial content script run.
      await new Promise((resolve) => setTimeout(resolve, 250));
    }
    if (!(host instanceof Element)) return;

    const priorBoard = document.getElementById("cfe-course-widget-board");
    const restoreEditMode =
      options?.keepEditMode === true ||
      Boolean(priorBoard?.classList.contains("cfe-layout-edit-mode"));
    const prefs = await loadCourseHomeWidgetPrefs();
    if (!prefs.layout && prefs.layoutsByCourse?.[courseId]) {
      // One-time migration: adopt the current course's legacy layout as the
      // shared layout so widget changes stay consistent across all courses.
      prefs.layout = prefs.layoutsByCourse[courseId];
      await saveCourseHomeWidgetPrefs(prefs);
    }
    const boardRefs = renderCourseWidgetBoard(host, prefs, courseId);
    if (restoreEditMode && typeof boardRefs?.setEditMode === "function") {
      boardRefs.setEditMode(true);
    }
    if (!boardRefs?.dueCard || !boardRefs.hasDueWidget) return;
    const card = boardRefs.dueCard;

    const url = new URL(`${baseOrigin}/api/v1/courses/${courseId}/assignments`);
    url.searchParams.set("per_page", "100");
    url.searchParams.append("include[]", "submission");
    url.searchParams.set("access_token", apiToken);

    try {
      const response = await fetch(url.toString(), {
        headers: { Accept: "application/json" },
      });
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }
      const all = await response.json();
      const allAssignments = (Array.isArray(all) ? all : []).filter(
        (item) => item && item.due_at,
      );
      const now = new Date();
      const upcoming = allAssignments
        .map((item) => ({
          item,
          due: new Date(item.due_at),
        }))
        .filter((entry) => !Number.isNaN(entry.due.getTime()))
        .filter((entry) => entry.due >= now)
        .sort((a, b) => {
          const diff = a.due.getTime() - b.due.getTime();
          if (diff !== 0) return diff;
          return String(a.item?.name || "").localeCompare(
            String(b.item?.name || ""),
          );
        });
      const filtered = upcoming.length ? [upcoming[0].item] : [];
      const checksMap = await loadCourseDueChecks();
      renderCourseDueWidgetItems(card, filtered, courseId, checksMap);
      if (boardRefs.calendarCard instanceof HTMLElement) {
        renderCourseCalendarWidget(boardRefs.calendarCard, allAssignments);
      }
      const listEl =
        boardRefs.dueList || card.querySelector(".cfe-course-due-list");
      listEl
        ?.querySelectorAll?.("input[data-course-due-check]")
        ?.forEach?.((input) => {
          input.addEventListener("change", async (event) => {
            const target = event.target;
            if (!(target instanceof HTMLInputElement)) return;
            const rawKey = target.getAttribute("data-course-due-check");
            if (!rawKey) return;
            let key = rawKey;
            try {
              key = decodeURIComponent(rawKey);
            } catch (error) {
              key = rawKey;
            }
            if (!key) return;
            const nextChecks = await loadCourseDueChecks();
            if (target.checked) {
              nextChecks[key] = true;
            } else {
              delete nextChecks[key];
            }
            await saveCourseDueChecks(nextChecks);
            const row = target.closest(".cfe-task-row");
            if (row) {
              row.classList.toggle("is-complete", target.checked);
            }
          });
        });
    } catch (error) {
      const listEl = card.querySelector(".cfe-course-due-list");
      if (listEl) {
        listEl.innerHTML =
          '<div class="cfe-widget-empty">Unable to load class assignments.</div>';
      }
    }
  }

  function bindHardNavigation(containerEl) {
    if (!containerEl || containerEl.dataset.cfeNavBound === "1") return;
    containerEl.dataset.cfeNavBound = "1";
    containerEl.addEventListener(
      "click",
      (event) => {
        if (event.defaultPrevented) return;
        if (event.button !== 0) return;
        if (event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) {
          return;
        }
        const target = event.target;
        if (!(target instanceof Element)) return;
        const anchor = target.closest("a.cfe-task-title, a.cfe-quick-link");
        if (!(anchor instanceof HTMLAnchorElement)) return;
        const href = anchor.getAttribute("href") || "";
        if (!href || href.startsWith("#")) return;
        if (anchor.target && anchor.target.toLowerCase() === "_blank") return;
        let resolved = "";
        try {
          resolved = new URL(href, window.location.origin).toString();
        } catch (error) {
          return;
        }
        event.preventDefault();
        event.stopPropagation();
        window.location.href = resolved;
      },
      true,
    );
  }

  async function init() {
    if (isStaleInstance()) return;
    if (window.__cfeInjected) return;
    window.__cfeInjected = true;

    const initialPath = window.location.pathname || "/";
    bindPassiveRouteSync();
    if (!(await isAuthUnlocked())) {
      cleanupNonDashboardUi();
      resetTheme();
      ensureAuthWall();
      window.__cfeInjected = false;
      return;
    }
    removeAuthWall();
    await loadSchoolStartMinutesSetting();
    if (!isDashboardPath(initialPath)) {
      cleanupNonDashboardUi();
      applyThemeForCurrentOrigin().catch(() => {
        // ignore transient reload invalidation
      });
      renderCourseDueWidgetForCurrentPage().catch(() => {
        // ignore non-dashboard widget failures
      });
      window.__cfeInjected = false;
      return;
    }

    const { canvasSettings } = await chrome.storage.sync.get("canvasSettings");
    if (!canvasSettings || !canvasSettings.enabled) {
      resetTheme();
      destroy();
      return;
    }

    const baseUrl = (canvasSettings.baseUrl || "").replace(/\/$/, "");
    const apiToken = canvasSettings.apiToken || "";
    if (!baseUrl || !apiToken) {
      resetTheme();
      removeDashboardShell();
      window.__cfeInjected = false;
      return;
    }

    const currentUrl = window.location.origin;
    const baseOrigin = (() => {
      try {
        return new URL(baseUrl).origin;
      } catch (error) {
        return "";
      }
    })();

    if (!baseOrigin || currentUrl !== baseOrigin) {
      resetTheme();
      removeDashboardShell();
      window.__cfeInjected = false;
      return;
    }

    // Disabled runtime route re-init to avoid Canvas SPA load loops.
    const path = window.location.pathname || "/";
    const isDashboard = isDashboardPath(path);
    await loadPopupTheme();
    if (!themeListenerBound) {
      themeListenerBound = true;
      prefersDark.addEventListener("change", () => {
        if (activeTheme.mode === "auto") {
          applyPopupTheme(activeTheme);
        }
      });
    }
    const body = await waitForBody();
    if (!body) {
      window.__cfeInjected = false;
      return;
    }
    // Clear stale containers from older content-script contexts after extension reloads.
    removeAllInjectedContainers();

    // Temporary safety: only inject custom UI on dashboard routes.
    // This prevents course/module render interference while preserving theming.
    if (!isDashboard) {
      window.__cfeInjected = false;
      return;
    }

    let container = null;
    let assignmentsEl = null;
    let completionWidgetEl = null;
    let taskListEl = null;
    let taskTabButtons = [];
    let widgetDockEl = null;
    let widgetCanvasEl = null;
    let editLayoutBtn = null;
    let sidebarListEl = null;
    let sidebarCourseEl = null;
    let refreshBtn = null;
    let personalListEl = null;
    let eventsMiniListEl = null;
    let announcementsMiniListEl = null;
    let personalTitleInput = null;
    let personalDateInput = null;
    let personalAddBtn = null;

    if (isDashboard) {
      container = document.createElement("section");
      container.id = "cfe-dashboard";
      container.classList.add("cfe-dashboard-only");
      container.innerHTML = `
        <div class="cfe-header">
          <div>
            <h2>Assignments</h2>
            <p>Due by next due date, this week, or this month.</p>
          </div>
          <div class="cfe-actions">
            <button class="cfe-edit-layout" type="button">Edit layout</button>
            <button class="cfe-filter is-active" data-filter="nextdue" type="button">Next Due Date</button>
            <button class="cfe-filter" data-filter="week" type="button">Week</button>
            <button class="cfe-filter" data-filter="month" type="button">Month</button>
            <div class="cfe-date-range-tools">
              <div class="cfe-date-range-label">Custom date range filter: select a start and end date, then press Go.</div>
              <div class="cfe-date-range-inputs">
                <input type="date" id="cfe-start-date" class="cfe-date-filter">
                <input type="date" id="cfe-end-date" class="cfe-date-filter">
                <button id="cfe-date-filter-btn" class="cfe-filter" data-filter="custom" type="button">Go</button>
              </div>
            </div>
            <button class="cfe-refresh" type="button">Refresh</button>
          </div>
        </div>
        <div class="cfe-widget-dock" id="cfe-widget-dock"></div>
        <div class="cfe-grid cfe-widget-canvas" id="cfe-widget-canvas">
          <div class="cfe-card cfe-dashboard-widget" id="cfe-assignments" data-widget-id="assignments" draggable="false">
            <button class="cfe-widget-drag-handle" type="button" data-widget-drag-handle="assignments" aria-label="Drag widget">Drag</button>
            <button class="cfe-widget-trash" type="button" data-widget-remove="assignments" aria-label="Delete widget">Trash</button>
            <h3>Assignments</h3>
            <div class="cfe-list" data-state="loading">
              <div class="cfe-loading">
                <span class="cfe-spinner"></span>
                <span>Loading assignments...</span>
              </div>
            </div>
          </div>
          <div class="cfe-card cfe-progress-card cfe-dashboard-widget" id="cfe-assignment-progress" data-widget-id="progress" draggable="false">
            <button class="cfe-widget-drag-handle" type="button" data-widget-drag-handle="progress" aria-label="Drag widget">Drag</button>
            <button class="cfe-widget-trash" type="button" data-widget-remove="progress" aria-label="Delete widget">Trash</button>
            <div class="cfe-progress-content" id="cfe-assignment-progress-content"></div>
          </div>
            <div class="cfe-widget cfe-tasks-widget cfe-dashboard-widget" data-widget-id="tasks" draggable="false">
              <button class="cfe-widget-drag-handle" type="button" data-widget-drag-handle="tasks" aria-label="Drag widget">Drag</button>
              <button class="cfe-widget-trash" type="button" data-widget-remove="tasks" aria-label="Delete widget">Trash</button>
              <div class="cfe-widget-head">
                <h3>Tasks</h3>
                <div class="cfe-tabs">
                  <button class="cfe-tab" data-tab="announcements" type="button" aria-label="Announcements">
                    <svg class="cfe-tab-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                      <path d="M3 11v2l8 2V9l-8 2z" />
                      <path d="M11 9l8-3v12l-8-3" />
                      <path d="M6 13l1 4h3l-2-4" />
                    </svg>
                  </button>
                  <button class="cfe-tab is-active" data-tab="assignments" type="button" aria-label="Assignments">
                    <svg class="cfe-tab-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                      <path d="M4 5h10a2 2 0 0 1 2 2v12H6a2 2 0 0 0-2 2V5z" />
                      <path d="M6 5v14" />
                      <path d="M14 8l4 4" />
                      <path d="M13 13l5-5" />
                    </svg>
                  </button>
                  <button class="cfe-tab" data-tab="quizzes" type="button" aria-label="Quizzes">
                    <svg class="cfe-tab-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                      <rect x="3" y="3" width="18" height="18" rx="2" />
                      <path d="M7 12l3 3 7-7" />
                    </svg>
                  </button>
                  <button class="cfe-tab" data-tab="discussions" type="button" aria-label="Discussions">
                    <svg class="cfe-tab-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                      <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
                    </svg>
                  </button>
                  <button class="cfe-tab" data-tab="events" type="button" aria-label="Calendar Events">
                    <svg class="cfe-tab-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                      <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
                      <line x1="16" y1="2" x2="16" y2="6"></line>
                      <line x1="8" y1="2" x2="8" y2="6"></line>
                      <line x1="3" y1="10" x2="21" y2="10"></line>
                    </svg>
                  </button>
                </div>
              </div>
              <div class="cfe-task-list" id="cfe-task-list"></div>
            </div>
            <div class="cfe-widget cfe-personal-widget cfe-dashboard-widget" data-widget-id="personal" draggable="false">
              <button class="cfe-widget-drag-handle" type="button" data-widget-drag-handle="personal" aria-label="Drag widget">Drag</button>
              <button class="cfe-widget-trash" type="button" data-widget-remove="personal" aria-label="Delete widget">Trash</button>
              <div class="cfe-widget-head">
                <h3>Personal To-Do</h3>
              </div>
              <div class="cfe-personal-controls">
                <input class="cfe-input" id="cfe-personal-title" type="text" placeholder="Add a personal task" />
                <input class="cfe-input" id="cfe-personal-date" type="date" />
                <button class="cfe-button" id="cfe-personal-add" type="button">Add</button>
              </div>
              <div class="cfe-task-list" id="cfe-personal-list"></div>
            </div>
            <div class="cfe-widget cfe-dashboard-widget cfe-filterbar-widget" data-widget-id="filterbar" data-filter-layout="horizontal" draggable="false">
              <button class="cfe-widget-drag-handle" type="button" data-widget-drag-handle="filterbar" aria-label="Drag widget">Drag</button>
              <button class="cfe-widget-trash" type="button" data-widget-remove="filterbar" aria-label="Delete widget">Trash</button>
              <div class="cfe-widget-head">
                <h3>Filter Bar</h3>
              </div>
              <div class="cfe-filterbar-buttons">
                <button class="cfe-filter cfe-filterbar-btn" data-filter="nextdue" type="button"><span class="cfe-filterbar-label">Next Due Date</span></button>
                <button class="cfe-filter cfe-filterbar-btn" data-filter="week" type="button"><span class="cfe-filterbar-label">Week</span></button>
                <button class="cfe-filter cfe-filterbar-btn" data-filter="month" type="button"><span class="cfe-filterbar-label">Month</span></button>
                <button class="cfe-filter cfe-filterbar-btn" data-filter="next3" type="button"><span class="cfe-filterbar-label">Next 3 Days</span></button>
                <button class="cfe-filter cfe-filterbar-btn" data-filter="overdue" type="button"><span class="cfe-filterbar-label">Overdue</span></button>
                <button class="cfe-filter cfe-filterbar-btn" data-filter="all" type="button"><span class="cfe-filterbar-label">All</span></button>
              </div>
            </div>
            <div class="cfe-widget cfe-dashboard-widget" data-widget-id="quicklinks" draggable="false">
              <button class="cfe-widget-drag-handle" type="button" data-widget-drag-handle="quicklinks" aria-label="Drag widget">Drag</button>
              <button class="cfe-widget-trash" type="button" data-widget-remove="quicklinks" aria-label="Delete widget">Trash</button>
              <div class="cfe-widget-head">
                <h3>Quick Links</h3>
              </div>
              <div class="cfe-quick-links">
                <a class="cfe-quick-link" href="/dashboard">Dashboard</a>
                <a class="cfe-quick-link" href="/calendar">Calendar</a>
                <a class="cfe-quick-link" href="/courses">Courses</a>
                <a class="cfe-quick-link" href="/grades">Grades</a>
              </div>
            </div>
            <div class="cfe-widget cfe-dashboard-widget" data-widget-id="eventsmini" draggable="false">
              <button class="cfe-widget-drag-handle" type="button" data-widget-drag-handle="eventsmini" aria-label="Drag widget">Drag</button>
              <button class="cfe-widget-trash" type="button" data-widget-remove="eventsmini" aria-label="Delete widget">Trash</button>
              <div class="cfe-widget-head">
                <h3>Upcoming Events</h3>
              </div>
              <div class="cfe-task-list" id="cfe-events-mini-list"></div>
            </div>
            <div class="cfe-widget cfe-dashboard-widget" data-widget-id="announcementsmini" draggable="false">
              <button class="cfe-widget-drag-handle" type="button" data-widget-drag-handle="announcementsmini" aria-label="Drag widget">Drag</button>
              <button class="cfe-widget-trash" type="button" data-widget-remove="announcementsmini" aria-label="Delete widget">Trash</button>
              <div class="cfe-widget-head">
                <h3>Recent Announcements</h3>
              </div>
              <div class="cfe-task-list" id="cfe-announcements-mini-list"></div>
            </div>
        </div>
      `;

      body.classList.add("cfe-dashboard-only");
      body.prepend(container);

      assignmentsEl = container.querySelector("#cfe-assignments .cfe-list");
      completionWidgetEl = container.querySelector(
        "#cfe-assignment-progress-content",
      );
      taskListEl = container.querySelector("#cfe-task-list");
      taskTabButtons = Array.from(container.querySelectorAll(".cfe-tab"));
      widgetDockEl = container.querySelector("#cfe-widget-dock");
      widgetCanvasEl = container.querySelector("#cfe-widget-canvas");
      editLayoutBtn = container.querySelector(".cfe-edit-layout");
      refreshBtn = container.querySelector(".cfe-refresh");
      personalListEl = container.querySelector("#cfe-personal-list");
      personalTitleInput = container.querySelector("#cfe-personal-title");
      personalDateInput = container.querySelector("#cfe-personal-date");
      personalAddBtn = container.querySelector("#cfe-personal-add");
      eventsMiniListEl = container.querySelector("#cfe-events-mini-list");
      announcementsMiniListEl = container.querySelector(
        "#cfe-announcements-mini-list",
      );
    } else {
      const findRightSide = () =>
        document.querySelector("#right-side") ||
        document.querySelector(".right-side") ||
        document.querySelector(".ic-Layout-aside") ||
        document.querySelector(".ic-Layout-aside-right") ||
        document.querySelector(".ic-Layout-sidebar") ||
        document.querySelector(".ic-Layout-aside__content");

      container = document.createElement("aside");
      container.id = "cfe-sidebar";
      container.innerHTML = `
        <div class="cfe-sidebar-header">
          <div class="cfe-sidebar-title">Tasks</div>
          <div class="cfe-sidebar-course" id="cfe-sidebar-course"></div>
          <div class="cfe-sidebar-actions">
            <button class="cfe-sidebar-refresh" type="button">Refresh</button>
            <button class="cfe-sidebar-toggle" type="button" aria-label="Toggle sidebar">⟨</button>
          </div>
        </div>
        <div class="cfe-task-list" id="cfe-sidebar-list"></div>
      `;
      let rightSide = findRightSide();
      let useFloating = !rightSide;
      container.classList.add(useFloating ? "cfe-floating" : "cfe-embedded");
      if (useFloating) {
        body.appendChild(container);
        body.classList.add("cfe-has-sidebar");
      } else {
        rightSide.prepend(container);
      }

      refreshBtn = container.querySelector(".cfe-sidebar-refresh");
      const toggleBtn = container.querySelector(".cfe-sidebar-toggle");
      sidebarListEl = container.querySelector("#cfe-sidebar-list");
      sidebarCourseEl = container.querySelector("#cfe-sidebar-course");

      const applySidebarState = (collapsed) => {
        if (useFloating) {
          body.classList.toggle("cfe-sidebar-collapsed", collapsed);
          toggleBtn.textContent = collapsed ? "⟩" : "⟨";
        } else {
          toggleBtn.style.display = "none";
        }
      };

      chrome.storage.sync
        .get("cfeSidebarCollapsed")
        .then(({ cfeSidebarCollapsed }) => {
          applySidebarState(Boolean(cfeSidebarCollapsed));
        });

      toggleBtn.addEventListener("click", async () => {
        if (!useFloating) return;
        const isCollapsed = body.classList.toggle("cfe-sidebar-collapsed");
        toggleBtn.textContent = isCollapsed ? "⟩" : "⟨";
        await chrome.storage.sync.set({ cfeSidebarCollapsed: isCollapsed });
      });

      // If the right-side container appears later, embed the sidebar there.
      if (useFloating) {
        const reattach = () => {
          const host = findRightSide();
          if (!host) return false;
          useFloating = false;
          body.classList.remove("cfe-has-sidebar", "cfe-sidebar-collapsed");
          container.classList.remove("cfe-floating");
          container.classList.add("cfe-embedded");
          host.prepend(container);
          toggleBtn.style.display = "none";
          return true;
        };

        const observer = new MutationObserver(() => {
          if (reattach()) {
            observer.disconnect();
          }
        });
        observer.observe(document.documentElement, {
          childList: true,
          subtree: true,
        });
      }

      // Handle SPA navigation
      let lastPath = window.location.pathname;
      setInterval(() => {
        if (isStaleInstance()) return;
        const newPath = window.location.pathname;
        if (newPath !== lastPath) {
          lastPath = newPath;
          renderSidebarTasks();
        }
      }, 500);
    }

    if (!container) return;
    bindHardNavigation(container);

    let activeFilter = "nextdue";
    let activeTaskTab = "assignments";
    let assignmentsCache = [];
    let announcementsCache = [];
    let discussionsCache = [];
    let eventsCache = [];
    let nicknamesCache = {};
    let coursesCache = {};
    let latestCourseIds = [];
    let auxiliaryDataLoaded = false;
    let auxiliaryDataPromise = null;
    let loadCycleId = 0;
    let manualCompletionMap = {};
    let personalTodos = [];
    let ringSlotByCourseKey = {};
    let ringSlotCounter = 0;
    const apiResponseCache = new Map();
    const apiInFlightRequests = new Map();
    const dashboardWidgetDefaults = {
      order: ["assignments", "progress", "tasks", "personal"],
    };
    const dashboardLayoutDefaults = {
      order: ["assignments", "progress", "tasks", "personal", "filterbar"],
      snapToFit: false,
      filterBarLayout: "vertical",
      positions: {
        assignments: { left: 408, top: 720 },
        progress: { left: 400, top: 20 },
        tasks: { left: 24, top: 24 },
        personal: { left: 408, top: 504 },
        filterbar: { left: 744, top: 24 },
      },
      sizes: {
        assignments: { width: 528, height: 252 },
        progress: { width: 320, height: 454 },
        tasks: { width: 320, height: 420 },
        personal: { width: 528, height: 185 },
        filterbar: { width: 264, height: 456 },
      },
    };
    const DASHBOARD_LAYOUT_VERSION = 6;
    const dashboardWidgetCatalog = [
      { id: "assignments", label: "Assignments" },
      { id: "progress", label: "Progress Ring" },
      { id: "tasks", label: "Tasks" },
      { id: "personal", label: "Personal To-Do" },
      { id: "filterbar", label: "Filter Bar" },
      { id: "quicklinks", label: "Quick Links" },
      { id: "eventsmini", label: "Upcoming Events" },
      { id: "announcementsmini", label: "Recent Announcements" },
    ];
    const dashboardWidgetIdSet = new Set(
      dashboardWidgetCatalog.map((item) => item.id),
    );
    let dashboardWidgetPrefs = {
      order: [...dashboardWidgetDefaults.order],
    };
    let draggingWidgetId = "";
    let layoutEditMode = false;
    let snapToFit = true;
    let filterBarLayout = "horizontal";
    let suppressAutoFit = false;
    let freeWidgetPositions = {};
    let lastFreeWidgetPositions = {};
    let widgetSizes = {};
    let freeDragState = null;
    let suppressManualCompletionSync = false;
    let suppressManualCompletionSyncTimer = null;
    const canvasTimeZone =
      window.ENV?.TIMEZONE ||
      window.ENV?.TIME_ZONE ||
      window.ENV?.timezone ||
      Intl.DateTimeFormat().resolvedOptions().timeZone;

    function normalizeDashboardWidgetPrefs(raw) {
      const hasSavedOrder = Array.isArray(raw?.order);
      let order = hasSavedOrder
        ? raw.order.filter((id) => dashboardWidgetIdSet.has(id))
        : [...dashboardWidgetDefaults.order];
      order = order.filter((id, index) => order.indexOf(id) === index);
      return { order };
    }

    function clonePositions(map) {
      const next = {};
      Object.entries(map || {}).forEach(([key, value]) => {
        if (!value || typeof value !== "object") return;
        next[key] = {
          left: Number(value.left || 0),
          top: Number(value.top || 0),
        };
      });
      return next;
    }

    function cloneWidgetSizes(map) {
      const next = {};
      Object.entries(map || {}).forEach(([key, value]) => {
        if (!value || typeof value !== "object") return;
        next[key] = {
          width: Number(value.width || 0),
          height: Number(value.height || 0),
        };
      });
      return next;
    }

    function stableSerialize(value) {
      const normalize = (input) => {
        if (Array.isArray(input)) {
          return input.map((entry) => normalize(entry));
        }
        if (input && typeof input === "object") {
          return Object.keys(input)
            .sort()
            .reduce((acc, key) => {
              acc[key] = normalize(input[key]);
              return acc;
            }, {});
        }
        return input;
      };
      try {
        return JSON.stringify(normalize(value));
      } catch (error) {
        return "";
      }
    }

    function ensureProgressWidgetState(defaultLayout) {
      let changed = false;
      if (!dashboardWidgetPrefs.order.includes("progress")) {
        dashboardWidgetPrefs.order = [
          ...dashboardWidgetPrefs.order,
          "progress",
        ];
        changed = true;
      }
      if (
        !freeWidgetPositions.progress ||
        typeof freeWidgetPositions.progress !== "object"
      ) {
        freeWidgetPositions.progress = { ...defaultLayout.positions.progress };
        changed = true;
      }
      if (
        !widgetSizes.progress ||
        typeof widgetSizes.progress !== "object" ||
        Number(widgetSizes.progress.width || 0) < 240 ||
        Number(widgetSizes.progress.height || 0) < 260
      ) {
        widgetSizes.progress = { ...defaultLayout.sizes.progress };
        changed = true;
      }
      return changed;
    }

    function getDefaultDashboardLayoutState() {
      return {
        order: [...dashboardLayoutDefaults.order],
        snapToFit: Boolean(dashboardLayoutDefaults.snapToFit),
        filterBarLayout: dashboardLayoutDefaults.filterBarLayout,
        positions: clonePositions(dashboardLayoutDefaults.positions),
        sizes: cloneWidgetSizes(dashboardLayoutDefaults.sizes),
      };
    }

    async function loadDashboardWidgetPrefs() {
      if (!isExtensionContextValid()) return;
      let stored;
      try {
        stored = await chrome.storage.sync.get([
          "cfeDashboardWidgets",
          "cfeDashboardSnapToFit",
          "cfeFilterBarLayout",
          "cfeDashboardLayoutCustomized",
          "cfeDashboardWidgetPositions",
          "cfeDashboardLastFreeWidgetPositions",
          "cfeDashboardWidgetSizes",
          "cfeDashboardLayoutVersion",
        ]);
      } catch (error) {
        return;
      }
      const {
        cfeDashboardWidgets,
        cfeDashboardSnapToFit,
        cfeFilterBarLayout,
        cfeDashboardLayoutCustomized,
        cfeDashboardWidgetPositions,
        cfeDashboardLastFreeWidgetPositions,
        cfeDashboardWidgetSizes,
        cfeDashboardLayoutVersion,
      } = stored || {};
      const defaultLayout = getDefaultDashboardLayoutState();
      if (Number(cfeDashboardLayoutVersion || 0) < DASHBOARD_LAYOUT_VERSION) {
        const hasCustomized = Boolean(cfeDashboardLayoutCustomized);
        dashboardWidgetPrefs = hasCustomized
          ? normalizeDashboardWidgetPrefs(cfeDashboardWidgets)
          : { order: [...defaultLayout.order] };
        snapToFit =
          typeof cfeDashboardSnapToFit === "boolean"
            ? cfeDashboardSnapToFit
            : defaultLayout.snapToFit;
        filterBarLayout =
          typeof cfeFilterBarLayout === "string" &&
          ["horizontal", "vertical"].includes(cfeFilterBarLayout)
            ? cfeFilterBarLayout
            : defaultLayout.filterBarLayout;
        freeWidgetPositions =
          cfeDashboardWidgetPositions &&
          typeof cfeDashboardWidgetPositions === "object"
            ? cfeDashboardWidgetPositions
            : clonePositions(defaultLayout.positions);
        const storedLastFree =
          cfeDashboardLastFreeWidgetPositions &&
          typeof cfeDashboardLastFreeWidgetPositions === "object"
            ? cfeDashboardLastFreeWidgetPositions
            : {};
        lastFreeWidgetPositions = Object.keys(storedLastFree).length
          ? clonePositions(storedLastFree)
          : clonePositions(freeWidgetPositions);
        widgetSizes =
          cfeDashboardWidgetSizes && typeof cfeDashboardWidgetSizes === "object"
            ? cfeDashboardWidgetSizes
            : cloneWidgetSizes(defaultLayout.sizes);
        try {
          await chrome.storage.sync.set({
            cfeDashboardWidgets: dashboardWidgetPrefs,
            cfeDashboardWidgetPositions: freeWidgetPositions,
            cfeDashboardLastFreeWidgetPositions: lastFreeWidgetPositions,
            cfeDashboardWidgetSizes: widgetSizes,
            cfeFilterBarLayout: filterBarLayout,
            cfeDashboardSnapToFit: snapToFit,
            cfeDashboardLayoutCustomized: hasCustomized,
            cfeDashboardLayoutVersion: DASHBOARD_LAYOUT_VERSION,
          });
        } catch (error) {
          // ignore transient reload invalidation
        }
        return;
      }
      const hasCustomized = Boolean(cfeDashboardLayoutCustomized);
      dashboardWidgetPrefs = hasCustomized
        ? normalizeDashboardWidgetPrefs(cfeDashboardWidgets)
        : { order: [...defaultLayout.order] };
      snapToFit =
        typeof cfeDashboardSnapToFit === "boolean"
          ? cfeDashboardSnapToFit
          : defaultLayout.snapToFit;
      filterBarLayout =
        typeof cfeFilterBarLayout === "string" &&
        ["horizontal", "vertical"].includes(cfeFilterBarLayout)
          ? cfeFilterBarLayout
          : defaultLayout.filterBarLayout;
      freeWidgetPositions =
        cfeDashboardWidgetPositions &&
        typeof cfeDashboardWidgetPositions === "object"
          ? cfeDashboardWidgetPositions
          : hasCustomized
            ? {}
            : clonePositions(defaultLayout.positions);
      const storedLastFree =
        cfeDashboardLastFreeWidgetPositions &&
        typeof cfeDashboardLastFreeWidgetPositions === "object"
          ? cfeDashboardLastFreeWidgetPositions
          : {};
      lastFreeWidgetPositions = Object.keys(storedLastFree).length
        ? clonePositions(storedLastFree)
        : clonePositions(freeWidgetPositions);
      widgetSizes =
        cfeDashboardWidgetSizes && typeof cfeDashboardWidgetSizes === "object"
          ? cfeDashboardWidgetSizes
          : hasCustomized
            ? {}
            : cloneWidgetSizes(defaultLayout.sizes);
      const recoveredProgress = ensureProgressWidgetState(defaultLayout);
      if (recoveredProgress && isExtensionContextValid()) {
        try {
          await chrome.storage.sync.set({
            cfeDashboardWidgets: dashboardWidgetPrefs,
            cfeDashboardWidgetPositions: freeWidgetPositions,
            cfeDashboardWidgetSizes: widgetSizes,
            cfeDashboardLayoutCustomized: true,
          });
        } catch (error) {
          // ignore transient reload invalidation
        }
      }
    }

    async function saveDashboardWidgetPrefs() {
      if (!isExtensionContextValid()) return;
      try {
        await chrome.storage.sync.set({
          cfeDashboardWidgets: dashboardWidgetPrefs,
          cfeDashboardLayoutCustomized: true,
        });
      } catch (error) {
        // ignore transient reload invalidation
      }
    }

    async function saveSnapToFitPref() {
      if (!isExtensionContextValid()) return;
      try {
        await chrome.storage.sync.set({ cfeDashboardSnapToFit: snapToFit });
      } catch (error) {
        // ignore transient reload invalidation
      }
    }

    async function saveFilterBarLayoutPref() {
      if (!isExtensionContextValid()) return;
      try {
        await chrome.storage.sync.set({ cfeFilterBarLayout: filterBarLayout });
      } catch (error) {
        // ignore transient reload invalidation
      }
    }

    async function saveWidgetPositions() {
      if (!isExtensionContextValid()) return;
      try {
        await chrome.storage.sync.set({
          cfeDashboardWidgetPositions: freeWidgetPositions,
          cfeDashboardLayoutCustomized: true,
        });
      } catch (error) {
        // ignore transient reload invalidation
      }
    }

    async function saveLastFreeWidgetPositions() {
      if (!isExtensionContextValid()) return;
      try {
        await chrome.storage.sync.set({
          cfeDashboardLastFreeWidgetPositions: lastFreeWidgetPositions,
        });
      } catch (error) {
        // ignore transient reload invalidation
      }
    }

    async function saveWidgetSizes() {
      if (!isExtensionContextValid()) return;
      try {
        await chrome.storage.sync.set({
          cfeDashboardWidgetSizes: widgetSizes,
          cfeDashboardLayoutCustomized: true,
        });
      } catch (error) {
        // ignore transient reload invalidation
      }
    }

    function scheduleLayoutPersist(delayMs = 700) {
      if (!isExtensionContextValid()) return;
      if (dashboardLayoutPersistTimer) {
        clearTimeout(dashboardLayoutPersistTimer);
      }
      dashboardLayoutPersistTimer = setTimeout(
        () => {
          dashboardLayoutPersistTimer = null;
          persistLayoutState();
        },
        Math.max(120, Number(delayMs) || 700),
      );
    }

    async function persistLayoutState() {
      await saveDashboardWidgetPrefs();
      await saveSnapToFitPref();
      await saveWidgetPositions();
      await saveLastFreeWidgetPositions();
      await saveWidgetSizes();
    }

    function getDashboardWidgetEl(widgetId) {
      if (!container) return null;
      return container.querySelector(`[data-widget-id="${widgetId}"]`);
    }

    function applyFilterBarLayout() {
      const filterWidget = getDashboardWidgetEl("filterbar");
      if (!filterWidget) return;
      filterWidget.setAttribute("data-filter-layout", filterBarLayout);
    }

    function applyWidgetSizes() {
      if (!widgetCanvasEl) return;
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]")
        .forEach((widget) => {
          const widgetId = widget.getAttribute("data-widget-id") || "";
          const size = widgetSizes[widgetId];
          if (!size || typeof size !== "object") {
            widget.style.width = "";
            widget.style.height = "";
            return;
          }
          const width = Number(size.width || 0);
          const height = Number(size.height || 0);
          const minimums = getWidgetResizeMinimums(
            widget,
            widgetId,
            width > 0 ? width : widget.getBoundingClientRect().width,
          );
          const safeWidth =
            width > 0 ? Math.max(minimums.minWidth, Math.round(width)) : 0;
          const rawSafeHeight =
            height > 0 ? Math.max(minimums.minHeight, Math.round(height)) : 0;
          const widgetMaxHeight =
            widgetId === "tasks"
              ? getTasksWidgetHeightCap()
              : Number.POSITIVE_INFINITY;
          const safeHeight =
            rawSafeHeight > 0
              ? Math.min(rawSafeHeight, widgetMaxHeight)
              : rawSafeHeight;
          widget.style.width = safeWidth > 0 ? `${safeWidth}px` : "";
          widget.style.height = safeHeight > 0 ? `${safeHeight}px` : "";
          if (safeWidth !== width || safeHeight !== height) {
            widgetSizes[widgetId] = {
              width: safeWidth || minimums.minWidth,
              height: safeHeight || minimums.minHeight,
            };
          }
        });
    }

    function getFilterBarMinWidth(widget) {
      const labels = Array.from(
        widget.querySelectorAll(".cfe-filterbar-btn .cfe-filterbar-label"),
      );
      if (!labels.length) return 220;
      const widest = labels.reduce((max, el) => {
        const node = el;
        return Math.max(max, node.scrollWidth || node.clientWidth || 0);
      }, 0);
      const base = Math.ceil(widest + 68);
      if (filterBarLayout === "vertical") {
        return Math.max(170, base);
      }
      return Math.max(220, base);
    }

    function getTasksWidgetHeightCap() {
      return Math.max(260, Math.min(620, window.innerHeight - 180));
    }

    function getWidgetResizeMinimums(widget, widgetId, proposedWidth) {
      const legendWidths = Array.from(
        widget.querySelectorAll(".cfe-legend-row"),
      ).map((el) => Math.ceil(el.clientWidth || 0));
      const legendMinWidth = legendWidths.length
        ? Math.max(...legendWidths) + 20
        : 0;
      const baseMinWidth =
        widgetId === "filterbar"
          ? getFilterBarMinWidth(widget)
          : widgetId === "progress"
            ? 180
            : 220;
      const effectiveLegendMinWidth =
        widgetId === "progress" ? 0 : legendMinWidth;
      const progressMinWidth = widgetId === "progress" ? 180 : 0;
      const minWidthBase = Math.max(
        baseMinWidth,
        effectiveLegendMinWidth,
        progressMinWidth,
      );
      const widthToMeasure = Math.max(minWidthBase, Number(proposedWidth || 0));
      const prevWidth = widget.style.width;
      const prevHeight = widget.style.height;
      widget.style.width = `${Math.round(widthToMeasure)}px`;
      widget.style.height = "auto";
      const naturalHeight = Math.ceil(
        widget.scrollHeight || widget.offsetHeight || 0,
      );
      widget.style.width = prevWidth;
      widget.style.height = prevHeight;
      const extraHeightPadding = widgetId === "progress" ? 24 : 14;
      return {
        minWidth: minWidthBase,
        minHeight: Math.max(140, naturalHeight + extraHeightPadding),
      };
    }

    function measureWidgetContentHeight(widget, width, widgetId = "") {
      const prevWidth = widget.style.width;
      const prevHeight = widget.style.height;
      widget.style.width = `${Math.round(width)}px`;
      widget.style.height = "auto";
      const height = Math.ceil(widget.scrollHeight || widget.offsetHeight || 0);
      widget.style.width = prevWidth;
      widget.style.height = prevHeight;
      const extraHeightPadding = widgetId === "progress" ? 24 : 14;
      return Math.max(140, height + extraHeightPadding);
    }

    function fitWidgetSizeToContent(widgetId) {
      const widget = getDashboardWidgetEl(widgetId);
      if (!(widget instanceof HTMLElement) || widget.hidden) return false;
      const currentRect = widget.getBoundingClientRect();
      const currentWidth = Math.round(currentRect.width || 0);
      const currentHeight = Math.round(currentRect.height || 0);
      const minimums = getWidgetResizeMinimums(widget, widgetId, currentWidth);
      const nextWidth = Math.max(currentWidth, minimums.minWidth);
      // Keep width user-driven, but make height follow current content size.
      const nextHeight =
        widgetId === "tasks"
          ? Math.min(minimums.minHeight, getTasksWidgetHeightCap())
          : minimums.minHeight;
      const shrank = nextWidth < currentWidth || nextHeight < currentHeight;
      if (nextWidth === currentWidth && nextHeight === currentHeight) {
        return false;
      }
      widget.style.width = `${nextWidth}px`;
      widget.style.height = `${nextHeight}px`;
      widgetSizes[widgetId] = { width: nextWidth, height: nextHeight };
      if (!snapToFit) {
        const left = Number(widget.style.left?.replace("px", "") || 0);
        const top = Number(widget.style.top?.replace("px", "") || 0);
        freeWidgetPositions[widgetId] = { left, top };
        lastFreeWidgetPositions[widgetId] = { left, top };
        pushOverlappingWidgetsDown(widgetId);
        if (shrank) {
          compactFreeWidgetsUpward();
        }
        refreshFreeCanvasHeight();
      }
      return true;
    }

    function fitAllWidgetsToContent() {
      if (!widgetCanvasEl) return;
      let changed = false;
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]:not([hidden])")
        .forEach((el) => {
          const widgetId = el.getAttribute("data-widget-id") || "";
          if (!widgetId) return;
          if (fitWidgetSizeToContent(widgetId)) {
            changed = true;
          }
        });
      if (!changed) return;
      saveWidgetSizes();
      if (!snapToFit) {
        saveWidgetPositions();
        saveLastFreeWidgetPositions();
      }
    }

    function ensureWidgetResizeHandles() {
      if (!widgetCanvasEl) return;
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]")
        .forEach((widget) => {
          const host = widget;
          if (host.querySelector("[data-widget-resize]")) return;
          const handles = [
            { dir: "e", label: "Resize width" },
            { dir: "s", label: "Resize height" },
            { dir: "se", label: "Resize width and height" },
          ];
          handles.forEach(({ dir, label }) => {
            const handle = document.createElement("button");
            handle.type = "button";
            handle.className = `cfe-widget-resize cfe-widget-resize-${dir}`;
            handle.setAttribute("data-widget-resize", dir);
            handle.setAttribute("aria-label", label);
            host.appendChild(handle);
          });
        });
    }

    function autoScrollViewport(clientY) {
      const edge = 96;
      const maxStep = 20;
      let delta = 0;
      if (clientY < edge) {
        const ratio = Math.max(0, (edge - clientY) / edge);
        delta = -Math.ceil(ratio * maxStep);
      } else if (clientY > window.innerHeight - edge) {
        const ratio = Math.max(
          0,
          (clientY - (window.innerHeight - edge)) / edge,
        );
        delta = Math.ceil(ratio * maxStep);
      }
      if (!delta) return;
      window.scrollBy(0, delta);
    }

    function setLayoutEditMode(enabled) {
      if (!widgetCanvasEl) return;
      layoutEditMode = Boolean(enabled);
      container.classList.toggle("cfe-layout-edit-mode", layoutEditMode);
      // Keep dashboard in absolute layout for persistent point-based placement.
      container.classList.add("cfe-snap-off");
      container.classList.toggle(
        "cfe-snap-grid",
        Boolean(layoutEditMode && snapToFit),
      );
      if (widgetDockEl) {
        widgetDockEl.hidden = !layoutEditMode;
      }
      if (editLayoutBtn) {
        editLayoutBtn.textContent = layoutEditMode
          ? "Done editing"
          : "Edit layout";
        editLayoutBtn.classList.toggle("is-active", layoutEditMode);
      }
      if (widgetCanvasEl) {
        widgetCanvasEl
          .querySelectorAll(".cfe-dashboard-widget")
          .forEach((el) => {
            el.setAttribute("draggable", "false");
          });
      }
      seedMissingFreePositionsFromCurrentDom();
      applyFreeLayoutPositions(true);
      applyWidgetSizes();
      ensureWidgetResizeHandles();
      applyFilterBarLayout();
      renderWidgetDock();
    }

    function applyDashboardWidgetLayout() {
      if (!widgetCanvasEl) return;
      const seen = new Set();
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]")
        .forEach((widget) => {
          const widgetId = widget.getAttribute("data-widget-id") || "";
          if (!widgetId) return;
          if (seen.has(widgetId)) {
            widget.remove();
            return;
          }
          seen.add(widgetId);
        });
      const activeIds = new Set(dashboardWidgetPrefs.order);
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]")
        .forEach((widget) => {
          const widgetId = widget.getAttribute("data-widget-id") || "";
          widget.hidden = !activeIds.has(widgetId);
        });

      dashboardWidgetPrefs.order.forEach((id) => {
        const widget = getDashboardWidgetEl(id);
        if (widget && widget.parentElement === widgetCanvasEl) {
          widgetCanvasEl.appendChild(widget);
        }
      });
      seedMissingFreePositionsFromCurrentDom();
      applyFreeLayoutPositions(true);
      applyWidgetSizes();
      ensureWidgetResizeHandles();
      applyFilterBarLayout();
    }

    function ensureFreePositionFor(widgetId, widgetIndex) {
      if (!widgetCanvasEl || !widgetId) return;
      if (freeWidgetPositions[widgetId]) return;
      const colWidth = 380;
      const rowHeight = 300;
      freeWidgetPositions[widgetId] = {
        left: 20 + (widgetIndex % 2) * colWidth,
        top: 20 + Math.floor(widgetIndex / 2) * rowHeight,
      };
    }

    function distanceBetweenRects(a, b) {
      const dx = Math.max(0, Math.max(a.left - b.right, b.left - a.right));
      const dy = Math.max(0, Math.max(a.top - b.bottom, b.top - a.bottom));
      if (dx === 0 && dy === 0) return 0;
      return Math.sqrt(dx * dx + dy * dy);
    }

    function findBestFreePositionForWidget(widget) {
      if (!widgetCanvasEl || !(widget instanceof HTMLElement)) {
        return { left: 20, top: 20 };
      }
      const canvasWidth = Math.max(320, widgetCanvasEl.clientWidth - 24);
      const width = widget.offsetWidth || 320;
      const height = widget.offsetHeight || 260;
      const maxLeft = Math.max(0, canvasWidth - width);
      const currentCanvasHeight = Math.max(
        520,
        widgetCanvasEl.offsetHeight || 520,
      );
      const maxTopInCurrent = Math.max(0, currentCanvasHeight - height - 16);
      const step = 18;
      const others = Array.from(
        widgetCanvasEl.querySelectorAll(
          ".cfe-dashboard-widget[data-widget-id]:not([hidden])",
        ),
      )
        .filter((el) => el !== widget)
        .map((el) => {
          const otherEl = el;
          const left = Number(otherEl.style.left?.replace("px", "") || 0);
          const top = Number(otherEl.style.top?.replace("px", "") || 0);
          const w = otherEl.offsetWidth || 320;
          const h = otherEl.offsetHeight || 260;
          return getWidgetRect(left, top, w, h);
        });

      let best = null;
      for (let top = 0; top <= maxTopInCurrent; top += step) {
        for (let left = 0; left <= maxLeft; left += step) {
          const probe = getWidgetRect(left, top, width, height);
          if (others.some((rect) => rectsOverlap(probe, rect))) continue;
          const edgeDistance = Math.min(
            probe.left,
            probe.top,
            Math.max(0, canvasWidth - probe.right),
            Math.max(0, currentCanvasHeight - probe.bottom),
          );
          let nearest = Number.POSITIVE_INFINITY;
          others.forEach((rect) => {
            nearest = Math.min(nearest, distanceBetweenRects(probe, rect));
          });
          const nearestScore =
            nearest === Number.POSITIVE_INFINITY ? 9999 : nearest;
          const score = Math.min(nearestScore, edgeDistance + 16);
          if (!best || score > best.score) {
            best = { left, top, score };
          }
        }
      }

      if (best) {
        return { left: best.left, top: best.top };
      }

      let maxBottom = 0;
      others.forEach((rect) => {
        maxBottom = Math.max(maxBottom, rect.bottom);
      });
      const top = maxBottom + 20;
      const nextMinHeight = Math.max(currentCanvasHeight, top + height + 24);
      widgetCanvasEl.style.minHeight = `${nextMinHeight}px`;
      return { left: 20, top };
    }

    function refreshFreeCanvasHeight() {
      if (!widgetCanvasEl) return;
      let maxBottom = 0;
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]:not([hidden])")
        .forEach((widget) => {
          const top = Number(widget.style.top?.replace("px", "") || 0);
          const height = widget.offsetHeight || 260;
          maxBottom = Math.max(maxBottom, top + height);
        });
      widgetCanvasEl.style.minHeight = `${Math.max(520, maxBottom + 24)}px`;
    }

    function seedMissingFreePositionsFromCurrentDom() {
      if (!widgetCanvasEl) return;
      const canvasRect = widgetCanvasEl.getBoundingClientRect();
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]:not([hidden])")
        .forEach((widget) => {
          const widgetId = widget.getAttribute("data-widget-id") || "";
          if (!widgetId) return;
          if (freeWidgetPositions[widgetId]) return;
          const rect = widget.getBoundingClientRect();
          freeWidgetPositions[widgetId] = {
            left: Math.max(0, rect.left - canvasRect.left),
            top: Math.max(0, rect.top - canvasRect.top),
          };
        });
    }

    function captureCurrentFreeLayoutFromDom() {
      if (!widgetCanvasEl) return clonePositions(freeWidgetPositions);
      const next = clonePositions(freeWidgetPositions);
      const canvasRect = widgetCanvasEl.getBoundingClientRect();
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]:not([hidden])")
        .forEach((widget) => {
          const widgetId = widget.getAttribute("data-widget-id") || "";
          if (!widgetId) return;
          const inlineLeft = Number(widget.style.left?.replace("px", ""));
          const inlineTop = Number(widget.style.top?.replace("px", ""));
          if (Number.isFinite(inlineLeft) && Number.isFinite(inlineTop)) {
            next[widgetId] = {
              left: Math.max(0, inlineLeft),
              top: Math.max(0, inlineTop),
            };
            return;
          }
          const rect = widget.getBoundingClientRect();
          next[widgetId] = {
            left: Math.max(0, rect.left - canvasRect.left),
            top: Math.max(0, rect.top - canvasRect.top),
          };
        });
      return next;
    }

    function writeFreePositionDataset(positions) {
      if (!widgetCanvasEl) return;
      const safePositions = positions || {};
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]")
        .forEach((widget) => {
          const widgetId = widget.getAttribute("data-widget-id") || "";
          if (!widgetId) return;
          const pos = safePositions[widgetId];
          if (!pos || typeof pos !== "object") return;
          widget.setAttribute("data-free-left", String(Number(pos.left || 0)));
          widget.setAttribute("data-free-top", String(Number(pos.top || 0)));
        });
    }

    function readFreePositionDataset() {
      if (!widgetCanvasEl) return {};
      const next = {};
      widgetCanvasEl
        .querySelectorAll(".cfe-dashboard-widget[data-widget-id]")
        .forEach((widget) => {
          const widgetId = widget.getAttribute("data-widget-id") || "";
          if (!widgetId) return;
          const left = Number(widget.getAttribute("data-free-left"));
          const top = Number(widget.getAttribute("data-free-top"));
          if (!Number.isFinite(left) || !Number.isFinite(top)) return;
          next[widgetId] = {
            left: Math.max(0, left),
            top: Math.max(0, top),
          };
        });
      return next;
    }

    function rectsOverlap(a, b, padding = 4) {
      return !(
        a.right + padding <= b.left ||
        a.left >= b.right + padding ||
        a.bottom + padding <= b.top ||
        a.top >= b.bottom + padding
      );
    }

    function getWidgetRect(left, top, width, height) {
      return {
        left,
        top,
        right: left + width,
        bottom: top + height,
      };
    }

    function findCollisionFreePosition(
      widget,
      desiredLeft,
      desiredTop,
      widgetId,
    ) {
      if (!widgetCanvasEl) {
        return { left: desiredLeft, top: desiredTop };
      }
      const canvasWidth = Math.max(320, widgetCanvasEl.clientWidth - 24);
      const width = widget.offsetWidth || 320;
      const height = widget.offsetHeight || 260;
      const maxLeft = Math.max(0, canvasWidth - width);
      const maxTop = Math.max(
        0,
        (widgetCanvasEl.offsetHeight || 900) + 900 - height,
      );
      const snapStep = snapToFit ? SNAP_GRID_STEP : 1;
      const quantize = (value) =>
        snapStep > 1 ? Math.round(value / snapStep) * snapStep : value;
      const clampLeft = Math.min(maxLeft, Math.max(0, quantize(desiredLeft)));
      const clampTop = Math.min(maxTop, Math.max(0, quantize(desiredTop)));
      const others = Array.from(
        widgetCanvasEl.querySelectorAll(
          ".cfe-dashboard-widget[data-widget-id]:not([hidden])",
        ),
      )
        .filter((el) => el.getAttribute("data-widget-id") !== widgetId)
        .map((el) => {
          const otherEl = el;
          const otherLeft = Number(otherEl.style.left?.replace("px", "") || 0);
          const otherTop = Number(otherEl.style.top?.replace("px", "") || 0);
          const otherWidth = otherEl.offsetWidth || 320;
          const otherHeight = otherEl.offsetHeight || 260;
          return getWidgetRect(otherLeft, otherTop, otherWidth, otherHeight);
        });

      const candidate = getWidgetRect(clampLeft, clampTop, width, height);
      const collides = others.some((rect) => rectsOverlap(candidate, rect));
      if (!collides) {
        return { left: clampLeft, top: clampTop };
      }

      let best = {
        left: clampLeft,
        top: clampTop,
        score: Number.POSITIVE_INFINITY,
      };
      const step = snapToFit ? SNAP_GRID_STEP : 12;
      for (let top = 0; top <= maxTop; top += step) {
        for (let left = 0; left <= maxLeft; left += step) {
          const probe = getWidgetRect(left, top, width, height);
          if (others.some((rect) => rectsOverlap(probe, rect))) continue;
          const score = Math.abs(left - clampLeft) + Math.abs(top - clampTop);
          if (score < best.score) {
            best = { left, top, score };
          }
        }
      }
      if (best.score < Number.POSITIVE_INFINITY) {
        return { left: best.left, top: best.top };
      }
      return { left: clampLeft, top: clampTop };
    }

    function collidesWithOtherWidgets(widgetId, left, top, width, height) {
      if (!widgetCanvasEl) return false;
      const probe = getWidgetRect(left, top, width, height);
      const others = Array.from(
        widgetCanvasEl.querySelectorAll(
          ".cfe-dashboard-widget[data-widget-id]:not([hidden])",
        ),
      )
        .filter((el) => el.getAttribute("data-widget-id") !== widgetId)
        .map((el) => {
          const otherEl = el;
          const otherLeft = Number(otherEl.style.left?.replace("px", "") || 0);
          const otherTop = Number(otherEl.style.top?.replace("px", "") || 0);
          const otherWidth = otherEl.offsetWidth || 320;
          const otherHeight = otherEl.offsetHeight || 260;
          return getWidgetRect(otherLeft, otherTop, otherWidth, otherHeight);
        });
      return others.some((rect) => rectsOverlap(probe, rect));
    }

    function pushOverlappingWidgetsDown(sourceWidgetId, options = {}) {
      const force = Boolean(options?.force);
      if (!widgetCanvasEl || (snapToFit && !force)) return false;
      const queue = [sourceWidgetId];
      const visitedPairs = new Set();
      const gap = snapToFit ? SNAP_WIDGET_GAP : FREE_WIDGET_GAP;
      const step = snapToFit ? SNAP_GRID_STEP : 1;
      let movedAny = false;
      let guard = 0;

      while (queue.length && guard < 220) {
        guard += 1;
        const currentId = queue.shift();
        if (!currentId) continue;
        const currentEl = getDashboardWidgetEl(currentId);
        if (!(currentEl instanceof HTMLElement) || currentEl.hidden) continue;
        const currentLeft = Number(
          currentEl.style.left?.replace("px", "") || 0,
        );
        const currentTop = Number(currentEl.style.top?.replace("px", "") || 0);
        const currentRect = getWidgetRect(
          currentLeft,
          currentTop,
          currentEl.offsetWidth || 320,
          currentEl.offsetHeight || 260,
        );

        const others = Array.from(
          widgetCanvasEl.querySelectorAll(
            ".cfe-dashboard-widget[data-widget-id]:not([hidden])",
          ),
        )
          .filter((el) => el.getAttribute("data-widget-id") !== currentId)
          .map((el) => ({
            id: el.getAttribute("data-widget-id") || "",
            el,
            left: Number(el.style.left?.replace("px", "") || 0),
            top: Number(el.style.top?.replace("px", "") || 0),
            width: el.offsetWidth || 320,
            height: el.offsetHeight || 260,
          }));

        others.forEach((other) => {
          if (!other.id) return;
          const pairKey = `${currentId}->${other.id}`;
          if (visitedPairs.has(pairKey)) return;
          visitedPairs.add(pairKey);
          const otherRect = getWidgetRect(
            other.left,
            other.top,
            other.width,
            other.height,
          );
          if (!rectsOverlap(currentRect, otherRect, 0)) return;
          const rawRequiredTop = Math.ceil(currentRect.bottom + gap);
          const requiredTop =
            step > 1 ? Math.ceil(rawRequiredTop / step) * step : rawRequiredTop;
          if (other.top >= requiredTop) return;
          const newTop = requiredTop;
          other.el.style.top = `${newTop}px`;
          freeWidgetPositions[other.id] = { left: other.left, top: newTop };
          lastFreeWidgetPositions[other.id] = { left: other.left, top: newTop };
          queue.push(other.id);
          movedAny = true;
        });
      }

      if (movedAny) {
        refreshFreeCanvasHeight();
      }
      return movedAny;
    }

    function compactFreeWidgetsUpward(options = {}) {
      const force = Boolean(options?.force);
      if (!widgetCanvasEl || (snapToFit && !force)) return false;
      const gap = snapToFit ? SNAP_WIDGET_GAP : FREE_WIDGET_GAP;
      const step = snapToFit ? SNAP_GRID_STEP : 1;
      const widgets = Array.from(
        widgetCanvasEl.querySelectorAll(
          ".cfe-dashboard-widget[data-widget-id]:not([hidden])",
        ),
      )
        .map((el) => {
          const id = el.getAttribute("data-widget-id") || "";
          if (!id) return null;
          const left = Number(el.style.left?.replace("px", "") || 0);
          const top = Number(el.style.top?.replace("px", "") || 0);
          const width = el.offsetWidth || 320;
          const height = el.offsetHeight || 260;
          return { id, el, left, top, width, height };
        })
        .filter(Boolean)
        .sort((a, b) => a.top - b.top || a.left - b.left);

      let movedAny = false;
      const placed = [];
      widgets.forEach((item) => {
        let minTop = 0;
        placed.forEach((prev) => {
          const xOverlap =
            item.left < prev.right && item.left + item.width > prev.left;
          if (xOverlap) {
            minTop = Math.max(minTop, prev.bottom + gap);
          }
        });
        const quantized =
          step > 1 ? Math.ceil(minTop / step) * step : Math.round(minTop);
        const newTop = Math.max(0, quantized);
        if (newTop !== item.top) {
          item.el.style.top = `${newTop}px`;
          freeWidgetPositions[item.id] = { left: item.left, top: newTop };
          lastFreeWidgetPositions[item.id] = { left: item.left, top: newTop };
          movedAny = true;
        }
        placed.push(getWidgetRect(item.left, newTop, item.width, item.height));
      });

      if (movedAny) {
        refreshFreeCanvasHeight();
      }
      return movedAny;
    }

    function applyFreeLayoutPositions(useExactSnapshot = false) {
      if (!widgetCanvasEl) return;
      const visibleWidgets = Array.from(
        widgetCanvasEl.querySelectorAll(
          ".cfe-dashboard-widget[data-widget-id]:not([hidden])",
        ),
      );
      visibleWidgets.forEach((widget, index) => {
        const widgetId = widget.getAttribute("data-widget-id") || "";
        if (!widgetId) return;
        ensureFreePositionFor(widgetId, index);
        const pos = freeWidgetPositions[widgetId];
        let safe;
        if (useExactSnapshot) {
          safe = {
            left: Math.max(0, Number(pos?.left || 0)),
            top: Math.max(0, Number(pos?.top || 0)),
          };
        } else {
          safe = findCollisionFreePosition(
            widget,
            Math.max(0, Number(pos?.left || 0)),
            Math.max(0, Number(pos?.top || 0)),
            widgetId,
          );
        }
        widget.style.left = `${safe.left}px`;
        widget.style.top = `${safe.top}px`;
        widget.setAttribute("data-free-left", String(safe.left));
        widget.setAttribute("data-free-top", String(safe.top));
        freeWidgetPositions[widgetId] = safe;
      });
      lastFreeWidgetPositions = clonePositions(freeWidgetPositions);
      refreshFreeCanvasHeight();
    }

    function renderWidgetDock() {
      if (!widgetDockEl) return;
      if (!layoutEditMode) {
        widgetDockEl.hidden = true;
        widgetDockEl.innerHTML = "";
        return;
      }
      const activeIds = new Set(dashboardWidgetPrefs.order);
      const addable = dashboardWidgetCatalog.filter(
        (item) => !activeIds.has(item.id),
      );
      const filterBarAddable = addable.find((item) => item.id === "filterbar");
      const otherAddable = addable.filter((item) => item.id !== "filterbar");
      widgetDockEl.innerHTML = `
        <div class="cfe-widget-taskbar-label">Add widget</div>
        <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-widget-snap-toggle="true">
          Snap to fit: ${snapToFit ? "On" : "Off"}
        </button>
        <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-filter-layout-toggle="true">
          Filter Bar Layout: ${
            filterBarLayout === "horizontal" ? "Horizontal" : "Vertical"
          }
        </button>
        ${
          filterBarAddable
            ? `<button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-widget-add="${filterBarAddable.id}">
          + ${filterBarAddable.label}
        </button>`
            : ""
        }
        <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-widget-action="delete-all">
          Delete all widgets
        </button>
        <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-widget-action="restore-original">
          Restore original widgets
        </button>
        <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-widget-action="default-dashboard">
          Use default dashboard
        </button>
        ${
          otherAddable.length
            ? otherAddable
                .map(
                  (item) => `
          <button class="cfe-widget-chip cfe-widget-chip-action" type="button" data-widget-add="${item.id}">
            + ${item.label}
          </button>
        `,
                )
                .join("")
            : !filterBarAddable
              ? '<span class="cfe-widget-taskbar-empty">All available widgets are already on the dashboard.</span>'
              : ""
        }
      `;

      widgetDockEl.querySelectorAll("[data-widget-add]").forEach((button) => {
        button.addEventListener("click", async () => {
          const widgetId = button.getAttribute("data-widget-add");
          if (!widgetId || !dashboardWidgetIdSet.has(widgetId)) return;
          if (!dashboardWidgetPrefs.order.includes(widgetId)) {
            dashboardWidgetPrefs.order.push(widgetId);
            applyDashboardWidgetLayout();
            if (!snapToFit) {
              const newWidget = getDashboardWidgetEl(widgetId);
              if (newWidget) {
                freeWidgetPositions[widgetId] =
                  findBestFreePositionForWidget(newWidget);
              } else {
                ensureFreePositionFor(
                  widgetId,
                  dashboardWidgetPrefs.order.length - 1,
                );
              }
              applyFreeLayoutPositions();
            }
            lastFreeWidgetPositions = clonePositions(freeWidgetPositions);
            renderWidgetDock();
            await saveDashboardWidgetPrefs();
            await saveWidgetPositions();
            await saveLastFreeWidgetPositions();
            updateWidgets();
          }
        });
      });

      widgetDockEl
        .querySelectorAll("[data-widget-snap-toggle]")
        .forEach((button) => {
          button.addEventListener("click", async () => {
            const switchingToSnapOff = snapToFit === true;
            const switchingToSnapOn = snapToFit === false;
            if (switchingToSnapOn) {
              // Preserve exact free layout before moving into snap mode.
              freeWidgetPositions = captureCurrentFreeLayoutFromDom();
              lastFreeWidgetPositions = clonePositions(freeWidgetPositions);
              writeFreePositionDataset(lastFreeWidgetPositions);
            }
            snapToFit = !snapToFit;
            container.classList.add("cfe-snap-off");
            if (switchingToSnapOff) {
              // Restore previous free-layout coordinates if they exist.
              if (!Object.keys(lastFreeWidgetPositions).length) {
                const domSnapshot = readFreePositionDataset();
                if (Object.keys(domSnapshot).length) {
                  lastFreeWidgetPositions = clonePositions(domSnapshot);
                }
              }
              if (!Object.keys(lastFreeWidgetPositions).length) {
                try {
                  const { cfeDashboardLastFreeWidgetPositions } =
                    await chrome.storage.sync.get(
                      "cfeDashboardLastFreeWidgetPositions",
                    );
                  if (
                    cfeDashboardLastFreeWidgetPositions &&
                    typeof cfeDashboardLastFreeWidgetPositions === "object"
                  ) {
                    lastFreeWidgetPositions = clonePositions(
                      cfeDashboardLastFreeWidgetPositions,
                    );
                  }
                } catch (error) {
                  // ignore transient reload invalidation
                }
              }
              if (Object.keys(lastFreeWidgetPositions).length) {
                freeWidgetPositions = clonePositions(lastFreeWidgetPositions);
              } else {
                seedMissingFreePositionsFromCurrentDom();
              }
              applyFreeLayoutPositions(true);
            } else {
              // Entering snap mode: move each widget to nearest available snap point.
              seedMissingFreePositionsFromCurrentDom();
              applyFreeLayoutPositions(false);
            }
            setLayoutEditMode(layoutEditMode);
            renderWidgetDock();
            await saveSnapToFitPref();
            await saveWidgetPositions();
            await saveLastFreeWidgetPositions();
          });
        });

      widgetDockEl
        .querySelectorAll("[data-filter-layout-toggle]")
        .forEach((button) => {
          button.addEventListener("click", async () => {
            filterBarLayout =
              filterBarLayout === "horizontal" ? "vertical" : "horizontal";
            applyFilterBarLayout();
            renderWidgetDock();
            await saveFilterBarLayoutPref();
          });
        });

      widgetDockEl
        .querySelectorAll("[data-widget-action]")
        .forEach((button) => {
          button.addEventListener("click", async () => {
            const action = button.getAttribute("data-widget-action");
            if (action === "delete-all") {
              dashboardWidgetPrefs.order = [];
              freeWidgetPositions = {};
              lastFreeWidgetPositions = {};
              widgetSizes = {};
              widgetCanvasEl
                ?.querySelectorAll(".cfe-dashboard-widget[data-widget-id]")
                .forEach((widget) => {
                  widget.hidden = true;
                });
              applyDashboardWidgetLayout();
              renderWidgetDock();
              await saveDashboardWidgetPrefs();
              await saveWidgetPositions();
              await saveLastFreeWidgetPositions();
              await saveWidgetSizes();
              updateWidgets();
              return;
            }
            if (action === "restore-original") {
              dashboardWidgetPrefs.order = [...dashboardWidgetDefaults.order];
              freeWidgetPositions = {};
              lastFreeWidgetPositions = {};
              widgetSizes = {};
              applyDashboardWidgetLayout();
              renderWidgetDock();
              await saveDashboardWidgetPrefs();
              await saveWidgetPositions();
              await saveLastFreeWidgetPositions();
              await saveWidgetSizes();
              updateWidgets();
              return;
            }
            if (action === "default-dashboard") {
              const defaultLayout = getDefaultDashboardLayoutState();
              dashboardWidgetPrefs.order = [...defaultLayout.order];
              snapToFit = defaultLayout.snapToFit;
              filterBarLayout = defaultLayout.filterBarLayout;
              freeWidgetPositions = clonePositions(defaultLayout.positions);
              lastFreeWidgetPositions = clonePositions(defaultLayout.positions);
              widgetSizes = cloneWidgetSizes(defaultLayout.sizes);
              container.classList.add("cfe-snap-off");
              applyDashboardWidgetLayout();
              applyFilterBarLayout();
              setLayoutEditMode(false);
              renderWidgetDock();
              if (isExtensionContextValid()) {
                try {
                  await chrome.storage.sync.set({
                    cfeDashboardWidgets: dashboardWidgetPrefs,
                    cfeDashboardSnapToFit: defaultLayout.snapToFit,
                    cfeDashboardWidgetPositions: freeWidgetPositions,
                    cfeDashboardLastFreeWidgetPositions:
                      lastFreeWidgetPositions,
                    cfeDashboardWidgetSizes: widgetSizes,
                    cfeFilterBarLayout: defaultLayout.filterBarLayout,
                    cfeDashboardLayoutCustomized: false,
                    cfeDashboardLayoutVersion: DASHBOARD_LAYOUT_VERSION,
                  });
                } catch (error) {
                  // ignore transient reload invalidation
                }
              }
              updateWidgets();
            }
          });
        });
    }

    function bindWidgetTrash() {
      if (!widgetCanvasEl) return;
      widgetCanvasEl.addEventListener("pointerdown", (event) => {
        const target = event.target;
        if (!(target instanceof Element)) return;
        const button = target.closest("[data-widget-remove]");
        if (!button) return;
        event.stopPropagation();
      });

      widgetCanvasEl.addEventListener("click", async (event) => {
        const target = event.target;
        if (!(target instanceof Element)) return;
        const button = target.closest("[data-widget-remove]");
        if (!button) return;
        if (!layoutEditMode) return;
        const widgetId = button.getAttribute("data-widget-remove");
        if (!widgetId || !dashboardWidgetIdSet.has(widgetId)) return;
        dashboardWidgetPrefs.order = dashboardWidgetPrefs.order.filter(
          (id) => id !== widgetId,
        );
        delete freeWidgetPositions[widgetId];
        delete lastFreeWidgetPositions[widgetId];
        delete widgetSizes[widgetId];
        applyDashboardWidgetLayout();
        renderWidgetDock();
        await saveDashboardWidgetPrefs();
        await saveWidgetPositions();
        await saveLastFreeWidgetPositions();
        await saveWidgetSizes();
        updateWidgets();
      });
    }

    function bindWidgetDragDrop() {
      if (!widgetCanvasEl || !assignmentsEl) return;
      const widgets = Array.from(
        widgetCanvasEl.querySelectorAll(
          ".cfe-dashboard-widget[data-widget-id]",
        ),
      );
      widgets.forEach((widget) => {
        widget.addEventListener("dragstart", (event) => {
          if (!layoutEditMode || !snapToFit) {
            event.preventDefault();
            return;
          }
          const widgetId = widget.getAttribute("data-widget-id") || "";
          draggingWidgetId = widgetId;
          widget.classList.add("is-dragging");
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", widgetId);
        });
        widget.addEventListener("dragend", () => {
          draggingWidgetId = "";
          widget.classList.remove("is-dragging");
          widgetCanvasEl
            .querySelectorAll(".cfe-widget-drop-target")
            .forEach((el) => el.classList.remove("cfe-widget-drop-target"));
        });
        widget.addEventListener("dragover", (event) => {
          if (!layoutEditMode || !snapToFit) return;
          event.preventDefault();
          autoScrollViewport(event.clientY);
          if (
            !draggingWidgetId ||
            draggingWidgetId === widget.dataset.widgetId
          ) {
            return;
          }
          widget.classList.add("cfe-widget-drop-target");
        });
        widget.addEventListener("dragleave", () => {
          widget.classList.remove("cfe-widget-drop-target");
        });
        widget.addEventListener("drop", async (event) => {
          if (!layoutEditMode || !snapToFit) return;
          event.preventDefault();
          widget.classList.remove("cfe-widget-drop-target");
          const targetId = widget.getAttribute("data-widget-id");
          if (!draggingWidgetId || !targetId || draggingWidgetId === targetId) {
            return;
          }
          const draggingEl = getDashboardWidgetEl(draggingWidgetId);
          const targetEl = getDashboardWidgetEl(targetId);
          if (!draggingEl || !targetEl || !widgetCanvasEl) return;

          const all = Array.from(
            widgetCanvasEl.querySelectorAll(
              ".cfe-dashboard-widget[data-widget-id]:not([hidden])",
            ),
          );
          const draggingIndex = all.indexOf(draggingEl);
          const targetIndex = all.indexOf(targetEl);
          if (draggingIndex === -1 || targetIndex === -1) return;

          if (draggingIndex < targetIndex) {
            widgetCanvasEl.insertBefore(draggingEl, targetEl.nextSibling);
          } else {
            widgetCanvasEl.insertBefore(draggingEl, targetEl);
          }
          dashboardWidgetPrefs.order = Array.from(
            widgetCanvasEl.querySelectorAll(
              ".cfe-dashboard-widget[data-widget-id]:not([hidden])",
            ),
          ).map((el) => el.getAttribute("data-widget-id"));
          await saveDashboardWidgetPrefs();
        });
      });
    }

    function bindFreeDrag() {
      if (!widgetCanvasEl) return;
      widgetCanvasEl.addEventListener("pointerdown", (event) => {
        if (!layoutEditMode) return;
        const target = event.target;
        if (!(target instanceof Element)) return;
        const handle = target.closest("[data-widget-drag-handle]");
        if (!handle) return;
        const widget = handle.closest(".cfe-dashboard-widget[data-widget-id]");
        if (!(widget instanceof HTMLElement) || widget.hidden) return;
        const widgetId = widget.getAttribute("data-widget-id") || "";
        if (!widgetId) return;

        const canvasRect = widgetCanvasEl.getBoundingClientRect();
        const widgetRect = widget.getBoundingClientRect();
        freeDragState = {
          widget,
          widgetId,
          offsetX: event.clientX - widgetRect.left,
          offsetY: event.clientY - widgetRect.top,
          canvasRect,
          pendingLeft: Number(widget.style.left?.replace("px", "") || 0),
          pendingTop: Number(widget.style.top?.replace("px", "") || 0),
        };
        widget.classList.add("is-dragging");
        widget.setPointerCapture(event.pointerId);
        event.preventDefault();
      });

      widgetCanvasEl.addEventListener("pointermove", (event) => {
        if (!freeDragState) return;
        const { widget, offsetX, offsetY } = freeDragState;
        autoScrollViewport(event.clientY);
        const canvasRect = widgetCanvasEl.getBoundingClientRect();
        const maxLeft = Math.max(0, canvasRect.width - widget.offsetWidth);
        const maxTop = Math.max(
          0,
          (widgetCanvasEl.offsetHeight || 600) + 900 - widget.offsetHeight,
        );
        const left = Math.min(
          maxLeft,
          Math.max(0, event.clientX - canvasRect.left - offsetX),
        );
        const top = Math.min(
          maxTop,
          Math.max(0, event.clientY - canvasRect.top - offsetY),
        );
        const step = snapToFit ? SNAP_GRID_STEP : 1;
        const snappedLeft = step > 1 ? Math.round(left / step) * step : left;
        const snappedTop = step > 1 ? Math.round(top / step) * step : top;
        // Passthrough mode: allow moving through occupied space while dragging.
        widget.style.left = `${snappedLeft}px`;
        widget.style.top = `${snappedTop}px`;
        freeDragState.pendingLeft = snappedLeft;
        freeDragState.pendingTop = snappedTop;
        refreshFreeCanvasHeight();
      });

      const endFreeDrag = async () => {
        if (!freeDragState) return;
        const { widget, widgetId, pendingLeft, pendingTop } = freeDragState;
        // Finalize to the nearest legal (non-overlapping) location.
        const safe = findCollisionFreePosition(
          widget,
          Math.max(0, Number(pendingLeft || 0)),
          Math.max(0, Number(pendingTop || 0)),
          widgetId,
        );
        widget.style.left = `${safe.left}px`;
        widget.style.top = `${safe.top}px`;
        freeWidgetPositions[widgetId] = safe;
        lastFreeWidgetPositions[widgetId] = { ...safe };
        widget.setAttribute("data-free-left", String(safe.left));
        widget.setAttribute("data-free-top", String(safe.top));
        widget.classList.remove("is-dragging");
        refreshFreeCanvasHeight();
        freeDragState = null;
        await saveWidgetPositions();
        await saveLastFreeWidgetPositions();
      };

      widgetCanvasEl.addEventListener("pointerup", () => {
        endFreeDrag();
      });
      widgetCanvasEl.addEventListener("pointercancel", () => {
        endFreeDrag();
      });
    }

    function bindWidgetResize() {
      if (!widgetCanvasEl) return;
      let resizeState = null;

      widgetCanvasEl.addEventListener("pointerdown", (event) => {
        if (!layoutEditMode) return;
        const target = event.target;
        if (!(target instanceof Element)) return;
        const handle = target.closest("[data-widget-resize]");
        if (!(handle instanceof HTMLElement)) return;
        const widget = handle.closest(".cfe-dashboard-widget[data-widget-id]");
        if (!(widget instanceof HTMLElement) || widget.hidden) return;
        const dir = handle.getAttribute("data-widget-resize") || "";
        if (!["e", "s", "se"].includes(dir)) return;
        const widgetId = widget.getAttribute("data-widget-id") || "";
        if (!widgetId) return;
        const rect = widget.getBoundingClientRect();
        const left = Number(widget.style.left?.replace("px", "") || 0);
        const top = Number(widget.style.top?.replace("px", "") || 0);
        resizeState = {
          pointerId: event.pointerId,
          handle,
          widget,
          widgetId,
          dir,
          startX: event.clientX,
          startY: event.clientY,
          startWidth: rect.width,
          startHeight: rect.height,
          left,
          top,
        };
        handle.setPointerCapture(event.pointerId);
        widget.classList.add("is-resizing");
        event.preventDefault();
        event.stopPropagation();
      });

      widgetCanvasEl.addEventListener("pointermove", (event) => {
        if (!resizeState) return;
        if (event.pointerId !== resizeState.pointerId) return;
        autoScrollViewport(event.clientY);
        const {
          widget,
          widgetId,
          dir,
          startX,
          startY,
          startWidth,
          startHeight,
          left,
          top,
        } = resizeState;
        const dx = event.clientX - startX;
        const dy = event.clientY - startY;
        const initialMinimums = getWidgetResizeMinimums(
          widget,
          widgetId,
          startWidth,
        );
        let minWidth = initialMinimums.minWidth;
        let minHeight = initialMinimums.minHeight;
        const canvasWidth = Math.max(280, widgetCanvasEl.clientWidth - 4);
        const snapMaxWidthAtPosition = Math.max(
          120,
          Math.floor(canvasWidth - Math.max(0, left)),
        );
        let maxWidth = snapToFit
          ? snapMaxWidthAtPosition
          : Math.max(minWidth, canvasWidth * 3);
        const maxHeight = Math.max(
          minHeight,
          (widgetCanvasEl.offsetHeight || 900) + 900 - top,
        );

        let nextWidth = startWidth;
        let nextHeight = startHeight;
        if (dir.includes("e")) {
          nextWidth = Math.min(maxWidth, Math.max(minWidth, startWidth + dx));
        }
        if (dir.includes("s")) {
          nextHeight = Math.min(
            maxHeight,
            Math.max(minHeight, startHeight + dy),
          );
        }

        const dynamicMinimums = getWidgetResizeMinimums(
          widget,
          widgetId,
          nextWidth,
        );
        minWidth = dynamicMinimums.minWidth;
        minHeight = dynamicMinimums.minHeight;
        maxWidth = snapToFit
          ? snapMaxWidthAtPosition
          : Math.max(minWidth, canvasWidth * 3);
        nextWidth = Math.min(maxWidth, Math.max(minWidth, nextWidth));
        if (snapToFit) {
          const step = SNAP_GRID_STEP;
          nextWidth = Math.round(nextWidth / step) * step;
          nextWidth = Math.min(maxWidth, Math.max(minWidth, nextWidth));
        }
        const contentMinHeight = measureWidgetContentHeight(
          widget,
          nextWidth,
          widgetId,
        );
        minHeight = Math.max(minHeight, contentMinHeight);
        if (dir === "e") {
          // Horizontal drag should keep content naturally fitted, not ratio-locked.
          nextHeight = minHeight;
        }
        nextHeight = Math.min(maxHeight, Math.max(minHeight, nextHeight));
        if (snapToFit) {
          const step = SNAP_GRID_STEP;
          nextHeight = Math.round(nextHeight / step) * step;
          nextHeight = Math.min(maxHeight, Math.max(minHeight, nextHeight));
        }

        if (
          !snapToFit &&
          collidesWithOtherWidgets(widgetId, left, top, nextWidth, nextHeight)
        ) {
          // In free mode, allow resize and push neighbors down minimally.
        }

        widget.style.width = `${Math.round(nextWidth)}px`;
        widget.style.height = `${Math.round(nextHeight)}px`;
        const currentLeft = Number(
          widget.style.left?.replace("px", "") || left,
        );
        const currentTop = Number(widget.style.top?.replace("px", "") || top);
        freeWidgetPositions[widgetId] = { left: currentLeft, top: currentTop };
        lastFreeWidgetPositions[widgetId] = {
          left: currentLeft,
          top: currentTop,
        };
        if (!snapToFit) {
          pushOverlappingWidgetsDown(widgetId);
        }
        widgetSizes[widgetId] = {
          width: Math.round(nextWidth),
          height: Math.round(nextHeight),
        };
        refreshFreeCanvasHeight();
      });

      const endResize = async (event) => {
        if (!resizeState) return;
        if (event && event.pointerId !== resizeState.pointerId) return;
        const { widget, widgetId } = resizeState;
        widget.classList.remove("is-resizing");
        const rect = widget.getBoundingClientRect();
        widgetSizes[widgetId] = {
          width: Math.round(rect.width),
          height: Math.round(rect.height),
        };
        if (snapToFit) {
          const currentLeft = Number(widget.style.left?.replace("px", "") || 0);
          const currentTop = Number(widget.style.top?.replace("px", "") || 0);
          const safe = findCollisionFreePosition(
            widget,
            Math.max(0, currentLeft),
            Math.max(0, currentTop),
            widgetId,
          );
          widget.style.left = `${safe.left}px`;
          widget.style.top = `${safe.top}px`;
          freeWidgetPositions[widgetId] = safe;
          lastFreeWidgetPositions[widgetId] = { ...safe };
          widget.setAttribute("data-free-left", String(safe.left));
          widget.setAttribute("data-free-top", String(safe.top));
          refreshFreeCanvasHeight();
        }
        resizeState = null;
        await saveWidgetSizes();
        await saveWidgetPositions();
        await saveLastFreeWidgetPositions();
      };

      widgetCanvasEl.addEventListener("pointerup", (event) => {
        endResize(event);
      });
      widgetCanvasEl.addEventListener("pointercancel", (event) => {
        endResize(event);
      });
      window.addEventListener("pointerup", (event) => {
        endResize(event);
      });
      window.addEventListener("pointercancel", (event) => {
        endResize(event);
      });
    }

    if (refreshBtn) {
      refreshBtn.addEventListener("click", () => {
        suppressAutoFit = true;
        Promise.resolve(loadData(true)).finally(() => {
          suppressAutoFit = false;
        });
      });
    }

    if (editLayoutBtn) {
      editLayoutBtn.addEventListener("click", async () => {
        const closingEditMode = layoutEditMode;
        setLayoutEditMode(!layoutEditMode);
        if (!closingEditMode) return;
        await persistLayoutState();
      });
    }

    if (assignmentsEl) {
      if (dashboardPageHideListener) {
        window.removeEventListener("pagehide", dashboardPageHideListener);
      }
      if (dashboardVisibilityPersistListener) {
        document.removeEventListener(
          "visibilitychange",
          dashboardVisibilityPersistListener,
        );
      }
      dashboardPageHideListener = () => {
        persistLayoutState();
      };
      dashboardVisibilityPersistListener = () => {
        if (document.visibilityState === "hidden") {
          persistLayoutState();
        }
      };
      window.addEventListener("pagehide", dashboardPageHideListener);
      document.addEventListener(
        "visibilitychange",
        dashboardVisibilityPersistListener,
      );
    }

    taskTabButtons.forEach((button) => {
      button.addEventListener("click", async () => {
        activeTaskTab = button.dataset.tab;
        taskTabButtons.forEach((btn) => {
          btn.classList.toggle("is-active", btn.dataset.tab === activeTaskTab);
        });
        if (
          ["announcements", "discussions", "events"].includes(activeTaskTab) &&
          !auxiliaryDataLoaded
        ) {
          renderLoadingState(taskListEl, "Loading tasks...");
          await ensureAuxiliaryData({ courseIds: collectCourseIds() });
        }
        renderDashboardTasks();
      });
    });

    const filterButtons = Array.from(
      container.querySelectorAll(".cfe-filter") || [],
    );
    const startDateInput = container.querySelector("#cfe-start-date");
    const endDateInput = container.querySelector("#cfe-end-date");

    filterButtons.forEach((button) => {
      button.addEventListener("click", () => {
        applyFilter(button.dataset.filter);
      });
    });

    if (dashboardStorageSyncListener && isExtensionContextValid()) {
      try {
        chrome.storage.onChanged.removeListener(dashboardStorageSyncListener);
      } catch (error) {
        // ignore listener reset failures
      }
      dashboardStorageSyncListener = null;
    }

    dashboardStorageSyncListener = (changes, area) => {
      if (isStaleInstance()) return;
      if (area !== "sync") return;
      if (changes.popupTheme) {
        applyPopupTheme(changes.popupTheme.newValue || {});
        if (assignmentsEl) {
          applyFilter(activeFilter);
        }
      }
      if (changes[SCHOOL_START_MINUTES_KEY]) {
        schoolStartMinutes = normalizeSchoolStartMinutes(
          changes[SCHOOL_START_MINUTES_KEY].newValue,
        );
        if (assignmentsEl) {
          applyFilter(activeFilter);
        }
        if (isCourseHomePath(window.location.pathname || "")) {
          renderCourseDueWidgetForCurrentPage().catch(() => {
            // ignore transient refresh failures
          });
        }
      }
      if (changes.cfeManualCompletions) {
        if (suppressManualCompletionSync) {
          return;
        }
        manualCompletionMap = changes.cfeManualCompletions.newValue || {};
        if (assignmentsEl) {
          applyFilter(activeFilter);
        }
        updateWidgets();
      }
      if (changes.cfePersonalTodos) {
        personalTodos = changes.cfePersonalTodos.newValue || [];
        renderPersonalTodos();
      }
      if (changes.cfeDashboardWidgets && assignmentsEl) {
        const nextPrefs = normalizeDashboardWidgetPrefs(
          changes.cfeDashboardWidgets.newValue,
        );
        if (
          stableSerialize(nextPrefs) !== stableSerialize(dashboardWidgetPrefs)
        ) {
          dashboardWidgetPrefs = nextPrefs;
          applyDashboardWidgetLayout();
          renderWidgetDock();
          updateWidgets();
        }
      }
      if (changes.cfeDashboardSnapToFit && assignmentsEl) {
        const nextSnap = changes.cfeDashboardSnapToFit.newValue !== false;
        if (nextSnap !== snapToFit) {
          snapToFit = nextSnap;
          container.classList.add("cfe-snap-off");
          setLayoutEditMode(layoutEditMode);
          applyFreeLayoutPositions(true);
          applyFilterBarLayout();
          renderWidgetDock();
        }
      }
      if (changes.cfeFilterBarLayout && assignmentsEl) {
        const nextRaw = changes.cfeFilterBarLayout.newValue;
        const next =
          typeof nextRaw === "string" &&
          ["horizontal", "vertical"].includes(nextRaw)
            ? nextRaw
            : "horizontal";
        if (next !== filterBarLayout) {
          filterBarLayout = next;
          applyFilterBarLayout();
          renderWidgetDock();
        }
      }
      if (changes.cfeDashboardWidgetPositions && assignmentsEl) {
        const nextPositions = clonePositions(
          changes.cfeDashboardWidgetPositions.newValue || {},
        );
        if (
          stableSerialize(nextPositions) !==
          stableSerialize(freeWidgetPositions)
        ) {
          freeWidgetPositions = nextPositions;
          lastFreeWidgetPositions = clonePositions(freeWidgetPositions);
          applyFreeLayoutPositions(true);
        }
      }
      if (changes.cfeDashboardLastFreeWidgetPositions && assignmentsEl) {
        const nextLast = clonePositions(
          changes.cfeDashboardLastFreeWidgetPositions.newValue || {},
        );
        if (
          stableSerialize(nextLast) !== stableSerialize(lastFreeWidgetPositions)
        ) {
          lastFreeWidgetPositions = nextLast;
        }
      }
      if (changes.cfeDashboardWidgetSizes && assignmentsEl) {
        const nextSizes = cloneWidgetSizes(
          changes.cfeDashboardWidgetSizes.newValue || {},
        );
        if (stableSerialize(nextSizes) !== stableSerialize(widgetSizes)) {
          widgetSizes = nextSizes;
          applyWidgetSizes();
          if (!snapToFit) {
            refreshFreeCanvasHeight();
          }
        }
      }
      if (changes.cfeSidebarCollapsed && !assignmentsEl) {
        body.classList.toggle(
          "cfe-sidebar-collapsed",
          Boolean(changes.cfeSidebarCollapsed.newValue),
        );
      }
    };

    if (isExtensionContextValid()) {
      chrome.storage.onChanged.addListener(dashboardStorageSyncListener);
    }

    async function loadManualCompletions() {
      try {
        const { cfeManualCompletions } = await chrome.storage.sync.get(
          "cfeManualCompletions",
        );
        manualCompletionMap = cfeManualCompletions || {};
      } catch (error) {
        manualCompletionMap = {};
      }
    }

    async function loadPersonalTodos() {
      try {
        const { cfePersonalTodos } =
          await chrome.storage.sync.get("cfePersonalTodos");
        personalTodos = Array.isArray(cfePersonalTodos) ? cfePersonalTodos : [];
      } catch (error) {
        personalTodos = [];
      }
    }

    async function savePersonalTodos() {
      if (!isExtensionContextValid()) return;
      try {
        await chrome.storage.sync.set({ cfePersonalTodos: personalTodos });
      } catch (error) {
        // ignore transient reload invalidation
      }
    }

    async function setManualCompletion(key, value) {
      if (!key) return;
      if (value) {
        manualCompletionMap[key] = true;
      } else {
        delete manualCompletionMap[key];
      }
      suppressManualCompletionSync = true;
      if (suppressManualCompletionSyncTimer) {
        clearTimeout(suppressManualCompletionSyncTimer);
      }
      suppressManualCompletionSyncTimer = setTimeout(() => {
        suppressManualCompletionSync = false;
        suppressManualCompletionSyncTimer = null;
      }, 900);
      try {
        await chrome.storage.sync.set({
          cfeManualCompletions: manualCompletionMap,
        });
      } catch (error) {
        suppressManualCompletionSync = false;
        if (suppressManualCompletionSyncTimer) {
          clearTimeout(suppressManualCompletionSyncTimer);
          suppressManualCompletionSyncTimer = null;
        }
      }
      if (assignmentsEl) {
        applyFilter(activeFilter);
      }
      updateWidgets();
    }

    function makeCacheKey(path, params = {}) {
      const query = new URLSearchParams();
      const keys = Object.keys(params).sort();
      keys.forEach((key) => {
        const value = params[key];
        if (Array.isArray(value)) {
          value.forEach((item) => {
            if (item !== undefined && item !== null) {
              query.append(key, String(item));
            }
          });
          return;
        }
        if (value !== undefined && value !== null) {
          query.append(key, String(value));
        }
      });
      const text = query.toString();
      return text ? `${path}?${text}` : path;
    }

    async function canvasFetch(path, params = {}) {
      const url = new URL(`${baseOrigin}${path}`);
      Object.entries(params).forEach(([key, value]) => {
        if (Array.isArray(value)) {
          value.forEach((item) => url.searchParams.append(key, item));
        } else if (value !== undefined && value !== null) {
          url.searchParams.set(key, value);
        }
      });
      url.searchParams.set("access_token", apiToken);

      const response = await fetch(url.toString(), {
        headers: {
          Accept: "application/json",
        },
      });

      if (!response.ok) {
        if (response.status === 401 || response.status === 403) {
          throw new Error(
            "Unauthorized. Check your API token and its permissions.",
          );
        }
        if (response.status === 404) {
          throw new Error(
            "API endpoint not found. Check your Canvas base URL.",
          );
        }

        const contentType = response.headers.get("content-type") || "";
        let message = "";
        if (contentType.includes("application/json")) {
          const data = await response.json().catch(() => null);
          message = data?.message || data?.errors?.[0]?.message || "";
        } else {
          message = await response.text().catch(() => "");
        }
        const cleaned = message.replace(/<[^>]+>/g, "").trim();
        throw new Error(cleaned || `Request failed: ${response.status}`);
      }

      const contentType = response.headers.get("content-type") || "";
      if (!contentType.includes("application/json")) {
        throw new Error(
          "Unexpected response. Check your Canvas base URL and API token.",
        );
      }

      return response.json();
    }

    async function cachedCanvasFetch(
      path,
      params = {},
      { ttlMs = 0, force = false } = {},
    ) {
      if (!ttlMs || force) {
        return canvasFetch(path, params);
      }
      const key = makeCacheKey(path, params);
      const now = Date.now();
      const cached = apiResponseCache.get(key);
      if (cached && now - cached.timestamp < ttlMs) {
        return cached.data;
      }
      const existingRequest = apiInFlightRequests.get(key);
      if (existingRequest) {
        return existingRequest;
      }
      const request = canvasFetch(path, params)
        .then((data) => {
          apiResponseCache.set(key, {
            timestamp: Date.now(),
            data,
          });
          return data;
        })
        .finally(() => {
          apiInFlightRequests.delete(key);
        });
      apiInFlightRequests.set(key, request);
      return request;
    }

    function renderLoadingState(targetEl, label) {
      if (!targetEl) return;
      targetEl.innerHTML = `
        <div class="cfe-loading">
          <span class="cfe-spinner"></span>
          <span>${label}</span>
        </div>
      `;
    }

    function toZonedDate(date) {
      const formatter = new Intl.DateTimeFormat("en-US", {
        timeZone: canvasTimeZone,
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: false,
      });
      const parts = formatter.formatToParts(date).reduce((acc, part) => {
        acc[part.type] = part.value;
        return acc;
      }, {});
      return new Date(
        `${parts.year}-${parts.month}-${parts.day}T${parts.hour}:${parts.minute}:${parts.second}`,
      );
    }

    function formatDate(value) {
      if (!value) return "No due date";
      const date = new Date(value);
      return date.toLocaleString(undefined, {
        timeZone: canvasTimeZone,
        month: "short",
        day: "numeric",
        hour: "numeric",
        minute: "2-digit",
      });
    }

    function formatShortDate(value) {
      const date = new Date(value);
      return date.toLocaleDateString(undefined, {
        timeZone: canvasTimeZone,
        month: "short",
        day: "numeric",
      });
    }

    function resolveCourseName(courseId, fallbackName) {
      if (!courseId) return fallbackName || "";
      return (
        nicknamesCache[courseId] || coursesCache[courseId] || fallbackName || ""
      );
    }

    function normalizeCanvasLink(rawUrl) {
      if (!rawUrl) return "#";
      try {
        const url = new URL(rawUrl, window.location.origin);
        if (url.origin !== window.location.origin) {
          return url.toString();
        }
        let path = url.pathname || "/";
        if (path.startsWith("/api/v1/")) {
          path = path.replace(/^\/api\/v1/, "");
        }
        // Normalize known plannable API-style paths to user-facing routes.
        const assignmentMatch = path.match(
          /^\/courses\/(\d+)\/assignments\/(\d+)/,
        );
        if (assignmentMatch) {
          path = `/courses/${assignmentMatch[1]}/assignments/${assignmentMatch[2]}`;
        }
        const quizMatch = path.match(/^\/courses\/(\d+)\/quizzes\/(\d+)/);
        if (quizMatch) {
          path = `/courses/${quizMatch[1]}/quizzes/${quizMatch[2]}`;
        }
        const discussionMatch = path.match(
          /^\/courses\/(\d+)\/discussion_topics\/(\d+)/,
        );
        if (discussionMatch) {
          path = `/courses/${discussionMatch[1]}/discussion_topics/${discussionMatch[2]}`;
        }
        return `${path}${url.search || ""}${url.hash || ""}`;
      } catch (error) {
        return rawUrl;
      }
    }

    function isItemComplete(item) {
      return item.is_submitted || Boolean(manualCompletionMap[item.item_key]);
    }

    function buildTaskRow(task, options = {}) {
      const isComplete = task.isComplete;
      const safeItemKey = escapeAttr(task.itemKey || "");
      const showCheckbox = options.showCheckbox && !task.isSubmitted;
      const checkboxMarkup = showCheckbox
        ? `
          <label class="cfe-task-check">
            <input type="checkbox" data-key="${safeItemKey}" ${
              task.isManual ? "checked" : ""
            } />
            <span></span>
          </label>
        `
        : "";
      const metaParts = [];
      if (options.showCourse && task.courseName) {
        metaParts.push(task.courseName);
      }
      if (task.dateValue) {
        metaParts.push(`${task.dateLabel} ${task.dateValue}`);
      }
      const metaLine = escapeHtml(metaParts.join(" - "));

      const safeUrl = escapeAttr(
        sanitizeHref(normalizeCanvasLink(task.url || "#")),
      );
      const safeTitle = escapeHtml(task.title || "Untitled");
      return `
        <div class="cfe-task-row${isComplete ? " is-complete" : ""}">
          ${checkboxMarkup}
          <div class="cfe-task-body">
            <a class="cfe-task-title" href="${safeUrl}">
              ${safeTitle}
            </a>
            <div class="cfe-task-meta">${metaLine || ""}</div>
          </div>
        </div>
      `;
    }

    function renderTaskList(targetEl, tasks, options = {}) {
      if (!targetEl) return;
      if (!tasks.length) {
        targetEl.innerHTML = `<div class="cfe-widget-empty">${
          options.emptyText || "No items."
        }</div>`;
        return;
      }

      targetEl.innerHTML = tasks
        .map((task) => buildTaskRow(task, options))
        .join("");

      if (options.showCheckbox) {
        targetEl.querySelectorAll(".cfe-task-check input").forEach((input) => {
          input.addEventListener("change", (event) => {
            setManualCompletion(event.target.dataset.key, event.target.checked);
          });
        });
      }
    }

    function renderPersonalTodos() {
      const syncPersonalWidgetHeightToContent = () => {
        const personalWidget = getDashboardWidgetEl("personal");
        if (!(personalWidget instanceof HTMLElement)) return;
        const currentRect = personalWidget.getBoundingClientRect();
        const currentWidth = Math.round(currentRect.width || 0);
        const currentHeight = Math.round(currentRect.height || 0);
        const minimums = getWidgetResizeMinimums(
          personalWidget,
          "personal",
          currentWidth,
        );
        const nextHeight = Math.round(minimums.minHeight);
        if (nextHeight === currentHeight) return;
        personalWidget.style.height = `${nextHeight}px`;
        widgetSizes.personal = {
          width: Math.max(currentWidth, minimums.minWidth),
          height: nextHeight,
        };
        refreshFreeCanvasHeight();
        if (layoutEditMode) {
          scheduleLayoutPersist(650);
        }
      };

      if (!personalListEl) return;
      if (!personalTodos.length) {
        personalListEl.innerHTML =
          '<div class="cfe-widget-empty">No personal tasks yet.</div>';
        syncPersonalWidgetHeightToContent();
        return;
      }

      const sorted = personalTodos.slice().sort((a, b) => {
        if (a.completed !== b.completed) {
          return a.completed ? 1 : -1;
        }
        if (a.due_on && b.due_on) {
          return new Date(a.due_on) - new Date(b.due_on);
        }
        if (a.due_on) return -1;
        if (b.due_on) return 1;
        return (a.created_at || 0) - (b.created_at || 0);
      });

      personalListEl.innerHTML = sorted
        .map((item) => {
          const dueLabel = item.due_on
            ? `Due ${escapeHtml(formatShortDate(item.due_on))}`
            : "";
          const safeId = escapeAttr(item.id || "");
          const safeTitle = escapeHtml(item.title || "");
          return `
            <div class="cfe-personal-row${item.completed ? " is-complete" : ""}">
              <label class="cfe-task-check">
                <input type="checkbox" data-id="${safeId}" ${
                  item.completed ? "checked" : ""
                } />
                <span></span>
              </label>
              <div class="cfe-personal-body">
                <div class="cfe-personal-title">${safeTitle}</div>
                <div class="cfe-personal-meta">${dueLabel}</div>
              </div>
              <button class="cfe-personal-delete" type="button" data-id="${safeId}">Remove</button>
            </div>
          `;
        })
        .join("");

      personalListEl
        .querySelectorAll(".cfe-task-check input")
        .forEach((input) => {
          input.addEventListener("change", (event) => {
            togglePersonalTodo(event.target.dataset.id, event.target.checked);
          });
        });

      personalListEl
        .querySelectorAll(".cfe-personal-delete")
        .forEach((button) => {
          button.addEventListener("click", (event) => {
            removePersonalTodo(event.target.dataset.id);
          });
        });
      syncPersonalWidgetHeightToContent();
    }

    async function addPersonalTodo() {
      if (!personalTitleInput) return;
      const title = personalTitleInput.value.trim();
      if (!title) return;
      const dueOn = personalDateInput?.value || "";
      const newItem = {
        id: `p_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        title,
        due_on: dueOn,
        completed: false,
        created_at: Date.now(),
      };
      personalTodos.unshift(newItem);
      personalTitleInput.value = "";
      if (personalDateInput) personalDateInput.value = "";
      await savePersonalTodos();
      renderPersonalTodos();
    }

    async function togglePersonalTodo(id, completed) {
      const index = personalTodos.findIndex((item) => item.id === id);
      if (index === -1) return;
      personalTodos[index].completed = completed;
      await savePersonalTodos();
      renderPersonalTodos();
    }

    async function removePersonalTodo(id) {
      personalTodos = personalTodos.filter((item) => item.id !== id);
      await savePersonalTodos();
      renderPersonalTodos();
    }

    function animateNumber(from, to, durationMs, onUpdate) {
      const start = performance.now();
      const delta = to - from;
      const duration = Math.max(150, durationMs || 520);
      const easeOutCubic = (t) => 1 - Math.pow(1 - t, 3);

      const step = (now) => {
        const progress = clamp((now - start) / duration, 0, 1);
        const eased = easeOutCubic(progress);
        const value = from + delta * eased;
        onUpdate(value, progress >= 1);
        if (progress < 1) requestAnimationFrame(step);
      };

      requestAnimationFrame(step);
    }

    function renderCompletionWidget(items) {
      if (!completionWidgetEl) return;

      if (!items.length) {
        ringSlotByCourseKey = {};
        ringSlotCounter = 0;
        completionWidgetEl.dataset.lastRatio = "0";
        completionWidgetEl.dataset.lastDone = "0";
        completionWidgetEl.dataset.lastTotal = "0";
        completionWidgetEl.dataset.lastRingRatios = "{}";
        completionWidgetEl.innerHTML =
          '<div class="cfe-widget-empty">No assignments in this range.</div>';
        return;
      }

      const byCourse = {};
      items.forEach((item) => {
        const courseName = resolveCourseName(
          item.course_id,
          item.course?.name || "Unknown course",
        );
        const courseKey =
          String(item.course_id || courseName).trim() || courseName;
        if (!byCourse[courseKey]) {
          byCourse[courseKey] = {
            key: courseKey,
            name: courseName,
            done: 0,
            total: 0,
          };
        }
        byCourse[courseKey].total += 1;
        if (isItemComplete(item)) {
          byCourse[courseKey].done += 1;
        }
      });

      const rows = Object.values(byCourse);
      const rowsBySlot = [...rows].sort((a, b) =>
        String(a.name || "").localeCompare(String(b.name || "")),
      );

      const totalCount = rowsBySlot.reduce((sum, row) => sum + row.total, 0);
      const completedCount = rowsBySlot.reduce((sum, row) => sum + row.done, 0);
      const completedRatio = totalCount
        ? Math.round((completedCount / totalCount) * 100)
        : 0;

      const maxRings = 6;
      const ringRows = rowsBySlot.slice(0, maxRings);
      const ringCount = Math.max(1, ringRows.length);
      const accent = activeTheme.accent || "#1f5f8b";
      const mode = resolveThemeMode(activeTheme.mode || "auto");
      const modePalette = getDefaultPalette(mode);
      const bg = activeTheme.bg || modePalette.bg;
      const surface = activeTheme.surface || modePalette.surface;
      const text = activeTheme.text || modePalette.text;
      let previousRingRatios = {};
      try {
        previousRingRatios = JSON.parse(
          completionWidgetEl.dataset.lastRingRatios || "{}",
        );
      } catch (error) {
        previousRingRatios = {};
      }
      const previousRatio = Number(completionWidgetEl.dataset.lastRatio || 0);
      const previousDone = Number(completionWidgetEl.dataset.lastDone || 0);
      const previousTotal = Number(completionWidgetEl.dataset.lastTotal || 0);
      const previousTotalDisplay = previousTotal || totalCount;
      const minInnerRadius = 30;
      const ringSize = clamp(152 + ringCount * 10, 162, 214);
      const ringOuterRadius = Math.round(ringSize * 0.4);
      let ringThickness = clamp(Math.round(10 - ringCount * 0.85), 4, 8);
      let ringGap = clamp(Math.round(9 - ringCount * 0.95), 2, 7);

      const computeInnerRadius = () =>
        ringOuterRadius -
        (ringCount - 1) * (ringThickness + ringGap) -
        ringThickness;

      let innerRadius = computeInnerRadius();
      while (innerRadius < minInnerRadius && ringGap > 2) {
        ringGap -= 1;
        innerRadius = computeInnerRadius();
      }
      while (innerRadius < minInnerRadius && ringThickness > 3) {
        ringThickness -= 1;
        innerRadius = computeInnerRadius();
      }
      const centerY = ringSize / 2;
      const showSubLabel = innerRadius >= 26;
      const ratioTextY = showSubLabel
        ? Math.round(centerY - Math.max(4, innerRadius * 0.12))
        : Math.round(centerY);
      const subTextY = Math.round(centerY + Math.max(13, innerRadius * 0.34));
      const centerFontSize = clamp(Math.round(innerRadius * 0.48), 13, 22);
      const subFontSize = clamp(Math.round(innerRadius * 0.22), 9, 12);
      const ringSvg = ringRows
        .map((row, ringIndex) => {
          const safeCourseName = String(row.name || "Unknown course")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
          const radius =
            ringOuterRadius - ringIndex * (ringThickness + ringGap);
          const circumference = 2 * Math.PI * radius;
          const pct = row.total ? row.done / row.total : 0;
          const length = pct * circumference;
          const previousPct = Number(previousRingRatios[row.key] || 0);
          const previousLength = clamp(previousPct, 0, 1) * circumference;
          const color = getSeriesColor(
            accent,
            ringIndex,
            maxRings,
            mode,
            bg,
            surface,
            text,
          );
          return `
            <circle class="cfe-ring-bg" cx="${ringSize / 2}" cy="${ringSize / 2}" r="${radius}" stroke-width="${ringThickness}" data-key="${row.key}" data-course-label="${safeCourseName}"></circle>
            <circle class="cfe-ring-progress" cx="${ringSize / 2}" cy="${ringSize / 2}" r="${radius}" stroke="${color}" stroke-width="${ringThickness}" stroke-linecap="${pct > 0.001 ? "round" : "butt"}" data-key="${row.key}" data-course-label="${safeCourseName}" data-length="${length}" data-offset="0" stroke-dasharray="${previousLength} ${circumference}" stroke-dashoffset="0" transform="rotate(-90 ${ringSize / 2} ${ringSize / 2})"></circle>
          `;
        })
        .join("");

      completionWidgetEl.innerHTML = `
        <div class="cfe-progress-title">Completion by class</div>
        <div class="cfe-multi-ring">
          <svg class="cfe-ring-chart" viewBox="0 0 ${ringSize} ${ringSize}" aria-label="Assignment completion by class">
            ${ringSvg}
            <text class="cfe-ring-center" x="50%" y="${ratioTextY}" text-anchor="middle" dominant-baseline="middle" style="font-size:${centerFontSize}px;">${previousRatio}%</text>
            ${
              showSubLabel
                ? `<text class="cfe-ring-sub" x="50%" y="${subTextY}" text-anchor="middle" dominant-baseline="middle" style="font-size:${subFontSize}px;">${previousDone}/${previousTotalDisplay} done</text>`
                : ""
            }
          </svg>
        </div>
        <div class="cfe-legend">
          ${rowsBySlot
            .map((row, rowIndex) => {
              const color = getSeriesColor(
                accent,
                rowIndex,
                rowsBySlot.length,
                mode,
                bg,
                surface,
                text,
              );
              const rowPct = row.total
                ? Math.round((row.done / row.total) * 100)
                : 0;
              const safeRowName = escapeHtml(row.name || "Unknown course");
              return `<div class="cfe-legend-row">
                <span class="cfe-legend-swatch" style="background:${color}"></span>
                <span class="cfe-legend-name">${safeRowName}</span>
                <span class="cfe-legend-meta">${row.done}/${row.total} (${rowPct}%)</span>
              </div>`;
            })
            .join("")}
        </div>
        <div class="cfe-ring-hover-tip" hidden></div>
      `;

      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          completionWidgetEl
            .querySelectorAll(".cfe-ring-progress")
            .forEach((segment) => {
              const length = Number(segment.getAttribute("data-length") || 0);
              const segmentOffset = Number(
                segment.getAttribute("data-offset") || 0,
              );
              const radius = Number(segment.getAttribute("r") || 0);
              const circumference = 2 * Math.PI * radius;
              segment.setAttribute(
                "stroke-dasharray",
                `${length} ${circumference}`,
              );
              segment.setAttribute("stroke-dashoffset", String(-segmentOffset));
            });
        });
      });

      const centerLabel = completionWidgetEl.querySelector(".cfe-ring-center");
      if (centerLabel) {
        animateNumber(previousRatio, completedRatio, 560, (value, done) => {
          centerLabel.textContent = `${Math.round(value)}%`;
          if (done) centerLabel.textContent = `${completedRatio}%`;
        });
      }

      const subLabel = completionWidgetEl.querySelector(".cfe-ring-sub");
      if (subLabel) {
        animateNumber(previousDone, completedCount, 560, (value, done) => {
          subLabel.textContent = `${Math.round(value)}/${totalCount} done`;
          if (done)
            subLabel.textContent = `${completedCount}/${totalCount} done`;
        });
      }

      const hoverTip = completionWidgetEl.querySelector(".cfe-ring-hover-tip");
      const ringSegments = completionWidgetEl.querySelectorAll(
        ".cfe-ring-progress, .cfe-ring-bg[data-course-label]",
      );
      if (hoverTip instanceof HTMLElement && ringSegments.length) {
        const hideTip = () => {
          hoverTip.hidden = true;
        };
        ringSegments.forEach((segment) => {
          segment.addEventListener("pointerenter", () => {
            const label = segment.getAttribute("data-course-label") || "";
            if (!label) return;
            hoverTip.textContent = label;
            hoverTip.hidden = false;
          });
          segment.addEventListener("pointermove", (event) => {
            if (hoverTip.hidden) return;
            const hostRect = completionWidgetEl.getBoundingClientRect();
            const x = event.clientX - hostRect.left + 10;
            const y = event.clientY - hostRect.top + 10;
            hoverTip.style.left = `${Math.max(6, x)}px`;
            hoverTip.style.top = `${Math.max(6, y)}px`;
          });
          segment.addEventListener("pointerleave", hideTip);
          segment.addEventListener("blur", hideTip);
        });
      }

      completionWidgetEl.dataset.lastRatio = String(completedRatio);
      completionWidgetEl.dataset.lastDone = String(completedCount);
      completionWidgetEl.dataset.lastTotal = String(totalCount);
      const nextRingRatios = { ...previousRingRatios };
      rowsBySlot.forEach((row) => {
        nextRingRatios[row.key] = row.total ? row.done / row.total : 0;
      });
      completionWidgetEl.dataset.lastRingRatios =
        JSON.stringify(nextRingRatios);
    }

    function syncAssignmentsWidgetHeightToContent() {
      const assignmentsWidget = getDashboardWidgetEl("assignments");
      if (!(assignmentsWidget instanceof HTMLElement)) return;
      const currentRect = assignmentsWidget.getBoundingClientRect();
      const currentWidth = Math.round(currentRect.width || 0);
      const currentHeight = Math.round(currentRect.height || 0);
      const minimums = getWidgetResizeMinimums(
        assignmentsWidget,
        "assignments",
        currentWidth,
      );
      const nextHeight = Math.round(minimums.minHeight);
      if (nextHeight === currentHeight) return;
      const grew = nextHeight > currentHeight;
      const shrank = nextHeight < currentHeight;
      assignmentsWidget.style.height = `${nextHeight}px`;
      widgetSizes.assignments = {
        width: Math.max(currentWidth, minimums.minWidth),
        height: nextHeight,
      };
      if (grew) {
        pushOverlappingWidgetsDown("assignments", { force: true });
      }
      if (shrank) {
        compactFreeWidgetsUpward({ force: true });
      }
      refreshFreeCanvasHeight();
      if (layoutEditMode) {
        scheduleLayoutPersist(650);
      }
    }

    function renderAssignments(items) {
      if (!assignmentsEl) return;
      if (!items.length) {
        assignmentsEl.textContent = "No assignments due in this range.";
        renderCompletionWidget(items);
        syncAssignmentsWidgetHeightToContent();
        return;
      }

      const tasks = items.slice(0, 12).map((item) => ({
        title: item.name,
        url: item.url,
        courseName: resolveCourseName(
          item.course_id,
          item.course?.name || "Unknown course",
        ),
        dateLabel: "Due",
        dateValue: formatDate(item.due_at),
        isSubmitted: item.is_submitted,
        isComplete: isItemComplete(item),
        isManual: Boolean(manualCompletionMap[item.item_key]),
        itemKey: item.item_key,
      }));

      assignmentsEl.innerHTML = tasks
        .map((task) =>
          buildTaskRow(task, {
            showCourse: true,
            showCheckbox: true,
          }),
        )
        .join("");

      assignmentsEl
        .querySelectorAll(".cfe-task-check input")
        .forEach((input) => {
          input.addEventListener("change", (event) => {
            setManualCompletion(event.target.dataset.key, event.target.checked);
          });
        });

      renderCompletionWidget(items);
      refreshSubmissionStates(items.slice(0, 12));
      syncAssignmentsWidgetHeightToContent();
    }

    function filterAssignments(items, filter) {
      const now = toZonedDate(new Date());
      const parseDueInCanvasZone = (item) => {
        if (!item?.due_at) return null;
        try {
          const due = toZonedDate(new Date(item.due_at));
          return Number.isNaN(due.getTime()) ? null : due;
        } catch (error) {
          return null;
        }
      };
      const { start: startOfToday, end: endOfToday } =
        getNextDueDateWindow(now);

      if (filter === "custom") {
        const startDate = startDateInput.valueAsDate;
        const endDate = endDateInput.valueAsDate;
        if (!startDate || !endDate) return [];

        // Set end date to end of day
        endDate.setHours(23, 59, 59, 999);

        return items.filter((item) => {
          if (!item.due_at) return false;
          const due = toZonedDate(new Date(item.due_at));
          return due >= startDate && due <= endDate;
        });
      }

      if (filter === "all") {
        return items.filter((item) => item.due_at);
      }

      if (filter === "overdue") {
        return items.filter((item) => {
          if (!item.due_at) return false;
          if (isItemComplete(item)) return false;
          const due = toZonedDate(new Date(item.due_at));
          return due < now;
        });
      }

      if (filter === "next3") {
        const end = new Date(now);
        end.setDate(end.getDate() + 3);
        return items.filter((item) => {
          if (!item.due_at) return false;
          const due = toZonedDate(new Date(item.due_at));
          return due >= now && due <= end;
        });
      }

      const endOfWeek = new Date(startOfToday);
      endOfWeek.setDate(endOfWeek.getDate() + 7);
      const endOfMonth = new Date(startOfToday);
      endOfMonth.setDate(endOfMonth.getDate() + 31);
      const nextDueWindow = getNextDueDateWindowFromItems(
        items,
        now,
        parseDueInCanvasZone,
      );

      return items.filter((item) => {
        if (!item.due_at) return false;
        const due = parseDueInCanvasZone(item);
        if (!due) return false;
        if (filter === "nextdue" || filter === "today") {
          if (filter === "nextdue") {
            if (!nextDueWindow) return false;
            return due >= nextDueWindow.start && due <= nextDueWindow.end;
          }
          return due >= startOfToday && due <= endOfToday;
        }
        if (filter === "week") {
          return due >= startOfToday && due < endOfWeek;
        }
        return due >= startOfToday && due < endOfMonth;
      });
    }

    function applyFilter(filter) {
      activeFilter = filter;
      filterButtons.forEach((btn) => {
        btn.classList.toggle("is-active", btn.dataset.filter === filter);
      });
      const filtered = filterAssignments(
        assignmentsCache.filter((item) => item.type === "assignment"),
        filter,
      );
      renderAssignments(filtered);
      renderDashboardTasks();
      if (!suppressAutoFit && layoutEditMode) {
        fitAllWidgetsToContent();
      }
    }

    function syncTasksWidgetHeightToContent() {
      const resized = fitWidgetSizeToContent("tasks");
      if (!resized) return;
      refreshFreeCanvasHeight();
      if (layoutEditMode) {
        scheduleLayoutPersist(650);
      }
    }

    function renderDashboardTasks() {
      if (!taskListEl) return;

      if (activeTaskTab === "announcements") {
        if (!auxiliaryDataLoaded && auxiliaryDataPromise) {
          renderLoadingState(taskListEl, "Loading tasks...");
          return;
        }
        const items = announcementsCache
          .slice()
          .sort((a, b) => new Date(b.posted_at) - new Date(a.posted_at))
          .slice(0, 8)
          .map((item) => ({
            title: item.title,
            url: item.url,
            courseName: resolveCourseName(item.course_id, item.course_name),
            dateLabel: "Posted",
            dateValue: formatShortDate(item.posted_at),
            isSubmitted: false,
            isComplete: false,
            isManual: false,
            itemKey: "",
          }));

        renderTaskList(taskListEl, items, {
          showCourse: true,
          showCheckbox: false,
          emptyText: "No announcements found.",
        });
        syncTasksWidgetHeightToContent();
        return;
      }

      if (activeTaskTab === "discussions") {
        if (!auxiliaryDataLoaded && auxiliaryDataPromise) {
          renderLoadingState(taskListEl, "Loading tasks...");
          return;
        }
        const items = discussionsCache
          .slice()
          .sort((a, b) => new Date(b.posted_at) - new Date(a.posted_at))
          .slice(0, 8)
          .map((item) => ({
            title: item.title,
            url: item.url,
            courseName: resolveCourseName(item.course_id, item.course_name),
            dateLabel: "Posted",
            dateValue: formatShortDate(item.posted_at),
            isSubmitted: false,
            isComplete: false,
            isManual: false,
            itemKey: "",
          }));

        renderTaskList(taskListEl, items, {
          showCourse: true,
          showCheckbox: false,
          emptyText: "No discussions found.",
        });
        syncTasksWidgetHeightToContent();
        return;
      }

      if (activeTaskTab === "events") {
        if (!auxiliaryDataLoaded && auxiliaryDataPromise) {
          renderLoadingState(taskListEl, "Loading tasks...");
          return;
        }
        const items = eventsCache
          .slice()
          .sort((a, b) => new Date(a.start_at) - new Date(b.start_at))
          .slice(0, 8)
          .map((item) => ({
            title: item.title,
            url: item.url,
            courseName: resolveCourseName(item.course_id, item.context_name),
            dateLabel: "Starts",
            dateValue: formatDate(item.start_at),
            isSubmitted: false,
            isComplete: false,
            isManual: false,
            itemKey: "",
          }));

        renderTaskList(taskListEl, items, {
          showCourse: true,
          showCheckbox: false,
          emptyText: "No calendar events found.",
        });
        syncTasksWidgetHeightToContent();
        return;
      }

      const isQuiz = activeTaskTab === "quizzes";
      const sourceItems = assignmentsCache.filter((item) => {
        if (isQuiz) {
          return item.type.includes("quiz");
        }
        return item.type === "assignment";
      });
      const filteredSource = filterAssignments(sourceItems, activeFilter);
      const items = filteredSource
        .slice()
        .sort((a, b) => new Date(a.due_at || 0) - new Date(b.due_at || 0))
        .slice(0, 10)
        .map((item) => ({
          title: item.name,
          url: item.url,
          courseName: resolveCourseName(
            item.course_id,
            item.course?.name || "Unknown course",
          ),
          dateLabel: "Due",
          dateValue: formatDate(item.due_at),
          isSubmitted: item.is_submitted,
          isComplete: isItemComplete(item),
          isManual: Boolean(manualCompletionMap[item.item_key]),
          itemKey: item.item_key,
        }));

      renderTaskList(taskListEl, items, {
        showCourse: true,
        showCheckbox: true,
        emptyText:
          activeFilter === "all"
            ? isQuiz
              ? "No quizzes found."
              : "No assignments found."
            : isQuiz
              ? "No quizzes due in this range."
              : "No assignments due in this range.",
      });
      syncTasksWidgetHeightToContent();
    }

    function renderExtraWidgets() {
      if (eventsMiniListEl) {
        if (!auxiliaryDataLoaded && auxiliaryDataPromise) {
          renderLoadingState(eventsMiniListEl, "Loading events...");
        } else {
          const eventItems = eventsCache
            .slice()
            .sort((a, b) => new Date(a.start_at) - new Date(b.start_at))
            .slice(0, 6)
            .map((item) => ({
              title: item.title,
              url: item.url,
              courseName: resolveCourseName(item.course_id, item.context_name),
              dateLabel: "Starts",
              dateValue: formatDate(item.start_at),
              isSubmitted: false,
              isComplete: false,
              isManual: false,
              itemKey: "",
            }));
          renderTaskList(eventsMiniListEl, eventItems, {
            showCourse: true,
            showCheckbox: false,
            emptyText: "No upcoming events.",
          });
        }
      }

      if (announcementsMiniListEl) {
        if (!auxiliaryDataLoaded && auxiliaryDataPromise) {
          renderLoadingState(
            announcementsMiniListEl,
            "Loading announcements...",
          );
        } else {
          const announcementItems = announcementsCache
            .slice()
            .sort((a, b) => new Date(b.posted_at) - new Date(a.posted_at))
            .slice(0, 6)
            .map((item) => ({
              title: item.title,
              url: item.url,
              courseName: resolveCourseName(item.course_id, item.course_name),
              dateLabel: "Posted",
              dateValue: formatShortDate(item.posted_at),
              isSubmitted: false,
              isComplete: false,
              isManual: false,
              itemKey: "",
            }));
          renderTaskList(announcementsMiniListEl, announcementItems, {
            showCourse: true,
            showCheckbox: false,
            emptyText: "No recent announcements.",
          });
        }
      }
    }

    function getCurrentCourseId() {
      const match = window.location.pathname.match(/\/courses\/(\d+)/);
      return match ? match[1] : "";
    }

    function renderSidebarTasks() {
      if (!sidebarListEl) return;
      const courseId = getCurrentCourseId();
      let items = assignmentsCache.filter(
        (item) => item.type === "assignment" || item.type.includes("quiz"),
      );

      if (courseId) {
        items = items.filter((item) => item.course_id === courseId);
        if (sidebarCourseEl) {
          sidebarCourseEl.textContent = resolveCourseName(courseId, "");
        }
      } else if (sidebarCourseEl) {
        sidebarCourseEl.textContent = "";
      }

      if (!courseId) {
        sidebarListEl.innerHTML =
          '<div class="cfe-widget-empty">Open a course to see tasks.</div>';
        return;
      }

      const tasks = items
        .slice()
        .sort((a, b) => new Date(a.due_at || 0) - new Date(b.due_at || 0))
        .slice(0, 8)
        .map((item) => ({
          title: item.name,
          url: item.url,
          courseName: resolveCourseName(
            item.course_id,
            item.course?.name || "",
          ),
          dateLabel: "Due",
          dateValue: formatDate(item.due_at),
          isSubmitted: item.is_submitted,
          isComplete: isItemComplete(item),
          isManual: Boolean(manualCompletionMap[item.item_key]),
          itemKey: item.item_key,
        }));

      renderTaskList(sidebarListEl, tasks, {
        showCourse: false,
        showCheckbox: true,
        emptyText: "No upcoming tasks.",
      });
    }

    function updateWidgets() {
      if (taskListEl) {
        renderDashboardTasks();
      }
      renderExtraWidgets();
      if (sidebarListEl) {
        renderSidebarTasks();
      }
      if (personalListEl) {
        renderPersonalTodos();
      }
      if (!suppressAutoFit && layoutEditMode) {
        fitAllWidgetsToContent();
      }
    }

    function normalizePlannerItems(items) {
      return (items || [])
        .filter((item) => item?.plannable || item?.title)
        .map((item) => {
          const plannable = item.plannable || {};
          const submission =
            plannable.submission ||
            item.submission ||
            plannable.submissions?.[0] ||
            item.submissions?.[0] ||
            {};
          const submissionState =
            submission.workflow_state ||
            plannable.submission_state ||
            item.submission_state ||
            "";
          const isSubmitted =
            Boolean(submission.submitted_at) ||
            ["submitted", "graded", "complete"].includes(submissionState) ||
            Boolean(item.submitted) ||
            Boolean(plannable.submitted) ||
            Boolean(item.completion_date) ||
            Boolean(plannable.completion_date);
          const contextCode = item.context_code || plannable.context_code || "";
          const courseIdMatch = contextCode.match(/course_(\d+)/);
          const courseIdFromContext = courseIdMatch ? courseIdMatch[1] : "";
          const url =
            plannable.html_url ||
            plannable.url ||
            item.html_url ||
            item.plannable?.html_url ||
            "";
          const urlMatch = url.match(/\/courses\/(\d+)\/assignments\/(\d+)/);
          const courseIdFromUrl = urlMatch ? urlMatch[1] : "";
          const assignmentIdFromUrl = urlMatch ? urlMatch[2] : "";
          const courseId = courseIdFromUrl || courseIdFromContext || "";
          const assignmentId =
            assignmentIdFromUrl ||
            plannable.assignment_id ||
            item.assignment_id ||
            plannable.id ||
            item.plannable_id ||
            "";
          const rawType = (
            item.plannable_type ||
            plannable.plannable_type ||
            ""
          ).toLowerCase();
          const type = rawType || (assignmentId ? "assignment" : "other");
          const itemKey =
            courseId && assignmentId
              ? `${courseId}:${assignmentId}`
              : url ||
                `${item.title || "item"}:${
                  item.plannable_date || item.due_at || ""
                }`;
          return {
            name: plannable.title || plannable.name || item.title || "Untitled",
            due_at:
              plannable.due_at || item.plannable_date || item.due_at || null,
            course: {
              name: resolveCourseName(
                courseId,
                item.context_name || plannable.course?.name || "",
              ),
            },
            url,
            is_submitted: isSubmitted,
            submission_state: submissionState,
            course_id: courseId,
            assignment_id: assignmentId,
            item_key: itemKey,
            type,
          };
        });
    }

    const submissionCache = new Map();

    async function fetchSubmission(courseId, assignmentId) {
      if (!courseId || !assignmentId) return null;
      const cacheKey = `${courseId}:${assignmentId}`;
      if (submissionCache.has(cacheKey)) {
        return submissionCache.get(cacheKey);
      }
      try {
        const submission = await canvasFetch(
          `/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions/self`,
        );
        submissionCache.set(cacheKey, submission);
        return submission;
      } catch (error) {
        submissionCache.set(cacheKey, null);
        return null;
      }
    }

    async function refreshSubmissionStates(items) {
      const updates = await Promise.all(
        items.map(async (item) => {
          if (item.is_submitted) return null;
          if (!item.course_id || !item.assignment_id) return null;
          const submission = await fetchSubmission(
            item.course_id,
            item.assignment_id,
          );
          if (!submission) return null;
          const state = submission.workflow_state || "";
          const isSubmitted =
            Boolean(submission.submitted_at) ||
            ["submitted", "graded", "complete"].includes(state);
          if (!isSubmitted) return null;
          return { item };
        }),
      );

      let updated = false;
      updates.filter(Boolean).forEach(({ item }) => {
        item.is_submitted = true;
        updated = true;
      });

      if (updated) {
        if (assignmentsEl) {
          applyFilter(activeFilter);
        }
        updateWidgets();
      }
    }

    async function loadAnnouncements(courseIds, force = false) {
      if (!courseIds.length) return [];
      const today = new Date();
      const start = new Date(
        today.getTime() - 1000 * 60 * 60 * 24 * 7,
      ).toISOString();
      const end = new Date(
        today.getTime() + 1000 * 60 * 60 * 24 * 14,
      ).toISOString();
      try {
        const announcements = await cachedCanvasFetch(
          "/api/v1/announcements",
          {
            "context_codes[]": courseIds.map((id) => `course_${id}`),
            start_date: start,
            end_date: end,
            per_page: 30,
          },
          {
            ttlMs: API_RESPONSE_CACHE_TTL.announcements,
            force,
          },
        );
        return (announcements || []).map((item) => {
          const contextCode = item.context_code || "";
          const courseIdMatch = contextCode.match(/course_(\d+)/);
          const courseId = courseIdMatch ? courseIdMatch[1] : "";
          return {
            title: item.title || item.subject || "Announcement",
            posted_at: item.posted_at || item.created_at,
            course_name: resolveCourseName(courseId, item.context_name || ""),
            course_id: courseId,
            url: item.html_url || "",
          };
        });
      } catch (error) {
        return [];
      }
    }

    async function loadDiscussions(courseIds, force = false) {
      if (!courseIds.length) return [];
      try {
        const results = await Promise.all(
          courseIds.map((courseId) =>
            cachedCanvasFetch(
              `/api/v1/courses/${courseId}/discussion_topics`,
              {
                per_page: 15,
                order_by: "recent_activity",
                "include[]": ["all_dates"],
              },
              {
                ttlMs: API_RESPONSE_CACHE_TTL.discussions,
                force,
              },
            ).catch(() => []),
          ),
        );
        return results.flat().map((item) => {
          const contextCode = item.context_code || "";
          const courseIdMatch = contextCode.match(/course_(\d+)/);
          const courseId = courseIdMatch
            ? courseIdMatch[1]
            : String(item.course_id || "");
          return {
            title: item.title,
            posted_at: item.last_reply_at || item.posted_at || item.created_at,
            course_name: resolveCourseName(courseId, item.context_name || ""),
            course_id: courseId,
            url: item.html_url || "",
          };
        });
      } catch (error) {
        return [];
      }
    }

    async function loadCalendarEvents(courseIds, force = false) {
      try {
        const today = new Date();
        const start = today.toISOString();
        const end = new Date(
          today.getTime() + 1000 * 60 * 60 * 24 * 31,
        ).toISOString();
        const events = await cachedCanvasFetch(
          "/api/v1/calendar_events",
          {
            type: "event",
            all_events: true,
            start_date: start,
            end_date: end,
            per_page: 50,
            "context_codes[]": courseIds.map((id) => `course_${id}`),
          },
          {
            ttlMs: API_RESPONSE_CACHE_TTL.events,
            force,
          },
        );
        return (events || []).map((item) => {
          const contextCode = item.context_code || "";
          const courseIdMatch = contextCode.match(/course_(\d+)/);
          const courseId = courseIdMatch ? courseIdMatch[1] : "";
          return {
            title: item.title,
            start_at: item.start_at,
            context_name: resolveCourseName(courseId, item.context_name || ""),
            course_id: courseId,
            url: item.html_url || "",
          };
        });
      } catch (error) {
        return [];
      }
    }

    async function loadCourses(force = false) {
      try {
        const courses = await cachedCanvasFetch(
          "/api/v1/courses",
          {
            per_page: 50,
            enrollment_state: "active",
          },
          {
            ttlMs: API_RESPONSE_CACHE_TTL.courses,
            force,
          },
        );
        const list = (courses || [])
          .map((course) => String(course.id))
          .filter(Boolean);
        const map = {};
        (courses || []).forEach((course) => {
          if (!course?.id) return;
          map[String(course.id)] = course.name || "";
        });
        return { list, map };
      } catch (error) {
        return { list: [], map: {} };
      }
    }

    async function loadCourseNicknames(force = false) {
      try {
        const nicknames = await cachedCanvasFetch(
          "/api/v1/users/self/course_nicknames",
          {},
          {
            ttlMs: API_RESPONSE_CACHE_TTL.nicknames,
            force,
          },
        );
        const nicknameMap = {};
        (nicknames || []).forEach((item) => {
          if (item?.course_id && item?.nickname) {
            nicknameMap[item.course_id] = item.nickname;
          }
        });
        return nicknameMap;
      } catch (error) {
        return {};
      }
    }

    function collectCourseIds(coursesData = null) {
      const courseIdsFromAssignments = assignmentsCache
        .map((item) => item.course_id)
        .filter(Boolean);
      const courseIdsFromCourses = coursesData?.list || latestCourseIds || [];
      return Array.from(
        new Set([...courseIdsFromAssignments, ...courseIdsFromCourses]),
      );
    }

    async function ensureAuxiliaryData({
      force = false,
      cycleId = loadCycleId,
      courseIds = null,
    } = {}) {
      if (!force && auxiliaryDataLoaded) return;
      if (!force && auxiliaryDataPromise) return auxiliaryDataPromise;
      const ids = Array.isArray(courseIds) ? courseIds : collectCourseIds();
      if (!ids.length) {
        announcementsCache = [];
        discussionsCache = [];
        eventsCache = [];
        auxiliaryDataLoaded = true;
        updateWidgets();
        return;
      }
      auxiliaryDataPromise = Promise.all([
        loadAnnouncements(ids, force),
        loadDiscussions(ids, force),
        loadCalendarEvents(ids, force),
      ])
        .then(([announcements, discussions, events]) => {
          if (cycleId !== loadCycleId) return;
          announcementsCache = announcements;
          discussionsCache = discussions;
          eventsCache = events;
          auxiliaryDataLoaded = true;
          updateWidgets();
        })
        .catch(() => {
          if (cycleId !== loadCycleId) return;
          announcementsCache = [];
          discussionsCache = [];
          eventsCache = [];
          auxiliaryDataLoaded = true;
          updateWidgets();
        })
        .finally(() => {
          if (cycleId === loadCycleId) {
            auxiliaryDataPromise = null;
          }
        });
      return auxiliaryDataPromise;
    }

    async function loadData(force = false) {
      const cycleId = ++loadCycleId;
      auxiliaryDataLoaded = false;
      auxiliaryDataPromise = null;
      announcementsCache = [];
      discussionsCache = [];
      eventsCache = [];
      if (assignmentsEl) {
        assignmentsEl.dataset.state = "loading";
        renderLoadingState(assignmentsEl, "Loading assignments...");
      }
      if (taskListEl) {
        renderLoadingState(taskListEl, "Loading tasks...");
      }
      if (sidebarListEl) {
        renderLoadingState(sidebarListEl, "Loading tasks...");
      }
      if (eventsMiniListEl) {
        renderLoadingState(eventsMiniListEl, "Loading events...");
      }
      if (announcementsMiniListEl) {
        renderLoadingState(announcementsMiniListEl, "Loading announcements...");
      }

      try {
        const today = new Date();
        const start = today.toISOString();
        const end = new Date(
          today.getTime() + 1000 * 60 * 60 * 24 * 31,
        ).toISOString();

        const [plannerItems, nicknameMap, coursesData] = await Promise.all([
          cachedCanvasFetch(
            "/api/v1/planner/items",
            {
              start_date: start,
              end_date: end,
              per_page: 50,
              "include[]": ["submission", "submissions"],
            },
            {
              ttlMs: API_RESPONSE_CACHE_TTL.planner,
              force,
            },
          ),
          loadCourseNicknames(force),
          loadCourses(force),
        ]);
        if (cycleId !== loadCycleId) return;
        nicknamesCache = nicknameMap || {};
        coursesCache = coursesData?.map || {};
        latestCourseIds = coursesData?.list || [];

        assignmentsCache = normalizePlannerItems(plannerItems);
        try {
          await chrome.runtime.sendMessage({
            type: "cfe-sync-reminder-assignments",
            assignments: assignmentsCache
              .filter((item) => item && item.due_at && !item.is_submitted)
              .slice(0, 120)
              .map((item) => ({
                item_key: item.item_key,
                name: item.name,
                due_at: item.due_at,
                url: item.url,
                course: { name: item?.course?.name || "" },
              })),
          });
        } catch (error) {
          // ignore reminder sync errors
        }

        if (assignmentsEl) {
          applyFilter(activeFilter);
        }

        updateWidgets();
        ensureAuxiliaryData({
          force,
          cycleId,
          courseIds: collectCourseIds(coursesData),
        });

        const toCheck = assignmentsCache
          .filter((item) => item.type === "assignment" && !item.is_submitted)
          .slice(0, 20);
        refreshSubmissionStates(toCheck);
      } catch (error) {
        if (cycleId !== loadCycleId) return;
        const message = error?.message || "Failed to load assignments.";
        if (assignmentsEl) {
          assignmentsEl.innerHTML = `
            <div class="cfe-widget-empty">
              <strong>We couldn't load your Canvas data.</strong>
              <div>${message}</div>
            </div>
          `;
        }
        if (sidebarListEl) {
          sidebarListEl.innerHTML = `
            <div class="cfe-widget-empty">
              <strong>Tasks unavailable.</strong>
              <div>${message}</div>
            </div>
          `;
        }
        if (taskListEl) {
          taskListEl.innerHTML = `
            <div class="cfe-widget-empty">
              <strong>Tasks unavailable.</strong>
              <div>${message}</div>
            </div>
          `;
        }
      }
    }

    await Promise.all([loadManualCompletions(), loadPersonalTodos()]);
    if (assignmentsEl) {
      await loadDashboardWidgetPrefs();
      applyDashboardWidgetLayout();
      applyFilterBarLayout();
      renderWidgetDock();
      bindWidgetTrash();
      bindWidgetDragDrop();
      bindFreeDrag();
      bindWidgetResize();
      setLayoutEditMode(false);
    }
    loadData();
    renderPersonalTodos();

    if (personalAddBtn) {
      personalAddBtn.addEventListener("click", () => {
        addPersonalTodo();
      });
    }
    if (personalTitleInput) {
      personalTitleInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          addPersonalTodo();
        }
      });
    }
  }

  chrome.storage.onChanged.addListener((changes, area) => {
    if (isStaleInstance()) return;

    if (area === "local" && changes[AUTH_GATE_STATE_KEY]) {
      const next = changes[AUTH_GATE_STATE_KEY].newValue || {};
      const unlocked = isAuthStateUnlocked(next);
      if (unlocked) {
        removeAuthWall();
        init();
      } else {
        resetTheme();
        destroy();
        ensureAuthWall();
      }
      return;
    }

    if (area !== "sync") return;

    if (changes.popupTheme) {
      syncThemeForCurrentPage().catch(() => {
        // ignore transient theme sync failures
      });
    }

    if (changes[SCHOOL_START_MINUTES_KEY]) {
      schoolStartMinutes = normalizeSchoolStartMinutes(
        changes[SCHOOL_START_MINUTES_KEY].newValue,
      );
      if (isCourseHomePath(window.location.pathname || "")) {
        renderCourseDueWidgetForCurrentPage().catch(() => {
          // ignore transient widget refresh failures
        });
      }
    }

    if (!changes.canvasSettings) return;

    syncThemeForCurrentPage().catch(() => {
      // ignore transient theme sync failures
    });
    const wasEnabled = changes.canvasSettings.oldValue?.enabled ?? true;
    const isEnabled = changes.canvasSettings.newValue?.enabled ?? true;

    if (wasEnabled !== isEnabled) {
      if (isEnabled) {
        init();
      } else {
        resetTheme();
        destroy();
      }
    }
  });

  let resumeThemeSyncTimer = null;
  let resumeThemeBurstTimers = [];

  function clearResumeThemeBurstTimers() {
    if (!resumeThemeBurstTimers.length) return;
    resumeThemeBurstTimers.forEach((timerId) => {
      clearTimeout(timerId);
    });
    resumeThemeBurstTimers = [];
  }

  function queueResumeThemeBurst() {
    clearResumeThemeBurstTimers();
    const delays = [280, 760, 1600];
    resumeThemeBurstTimers = delays.map((delay) =>
      setTimeout(() => {
        if (isStaleInstance() || !isExtensionContextValid()) return;
        syncThemeForCurrentPage().catch(() => {
          // ignore transient wake sync failures
        });
      }, delay),
    );
  }

  function scheduleResumeThemeSync(delayMs = 120) {
    if (isStaleInstance() || !isExtensionContextValid()) return;
    if (resumeThemeSyncTimer) {
      clearTimeout(resumeThemeSyncTimer);
    }
    clearResumeThemeBurstTimers();
    resumeThemeSyncTimer = setTimeout(
      () => {
        resumeThemeSyncTimer = null;
        syncThemeForCurrentPage().catch(() => {
          // ignore transient wake sync failures
        });
        queueResumeThemeBurst();
      },
      Math.max(0, Number(delayMs) || 0),
    );
  }

  window.addEventListener("focus", () => {
    scheduleResumeThemeSync(120);
  });
  window.addEventListener("pageshow", () => {
    scheduleResumeThemeSync(60);
  });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      scheduleResumeThemeSync(80);
    }
  });

  preloadTheme();

  (async () => {
    const { canvasSettings } = await chrome.storage.sync.get("canvasSettings");
    if (canvasSettings?.enabled ?? true) {
      init();
    }
  })();
})();
