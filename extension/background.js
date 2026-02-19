"use strict";

const DEFAULT_SETTINGS = {
  baseUrl: "",
  apiToken: "",
  enabled: true,
};

const CALENDAR_REMINDER_KEY = "cfeCalendarRemindersEnabled";
const ALARM_MAP_KEY = "cfeReminderAlarmMap";
const GOOGLE_EVENT_MAP_KEY = "cfeGoogleCalendarEventMap";
const ALARM_PREFIX = "cfe-due-";
const REMINDER_LEAD_MS = 60 * 60 * 1000; // 1 hour before due
const MAX_DUE_LOOKAHEAD_MS = 45 * 24 * 60 * 60 * 1000; // 45 days
const REFRESH_ALARM = "cfe-calendar-refresh";
const REFRESH_INTERVAL_MINUTES = 120;

function alarmNameFor(itemKey, dueAtMs) {
  const raw = `${itemKey || "assignment"}:${dueAtMs || 0}`;
  const safe = raw
    .toLowerCase()
    .replace(/[^a-z0-9:_-]/g, "_")
    .replace(/_+/g, "_")
    .slice(0, 180);
  return `${ALARM_PREFIX}${safe}`;
}

function eventSignatureFor(item) {
  return [
    item?.dueAtMs || 0,
    item?.name || "",
    item?.courseName || "",
    item?.url || "",
  ].join("|");
}

function getAuthToken(interactive = false) {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive }, (token) => {
      const err = chrome.runtime.lastError;
      if (err) {
        reject(new Error(err.message || "Failed to get Google auth token."));
        return;
      }
      resolve(token || "");
    });
  });
}

function removeCachedAuthToken(token) {
  return new Promise((resolve) => {
    if (!token) {
      resolve();
      return;
    }
    chrome.identity.removeCachedAuthToken({ token }, () => resolve());
  });
}

async function googleCalendarRequest(path, token, options = {}) {
  if (!token) {
    const error = new Error("Google Calendar is not connected.");
    error.status = 401;
    throw error;
  }
  const method = options.method || "GET";
  const headers = {
    Authorization: `Bearer ${token}`,
  };
  if (options.body) {
    headers["Content-Type"] = "application/json";
  }
  const response = await fetch(
    `https://www.googleapis.com/calendar/v3${path}`,
    {
      method,
      headers,
      body: options.body ? JSON.stringify(options.body) : undefined,
    },
  );
  if (response.ok) {
    if (response.status === 204) return null;
    const text = await response.text();
    return text ? JSON.parse(text) : null;
  }
  const raw = await response.text();
  const error = new Error(
    `Google Calendar API ${response.status}: ${(raw || "unknown error").slice(0, 240)}`,
  );
  error.status = response.status;
  error.body = raw;
  throw error;
}

async function verifyGoogleCalendarConnection(interactive = false) {
  let token = "";
  try {
    token = await getAuthToken(interactive);
  } catch (error) {
    if (!interactive) {
      return {
        connected: false,
        token: "",
        error: String(error?.message || error),
      };
    }
    throw error;
  }
  if (!token) return { connected: false, token: "" };

  // A valid token is enough for connectivity. Scope checks happen naturally on
  // actual Calendar requests (create/update/delete events).
  return { connected: true, token };
}

function normalizeAssignments(rawItems = []) {
  const now = Date.now();
  const latest = now + MAX_DUE_LOOKAHEAD_MS;
  const byItemKey = new Map();
  for (const item of rawItems) {
    const dueAtMs = Date.parse(item?.due_at || "");
    if (!Number.isFinite(dueAtMs) || dueAtMs <= now || dueAtMs > latest)
      continue;
    const itemKey = String(item?.item_key || "").trim();
    if (!itemKey) continue;
    byItemKey.set(itemKey, {
      itemKey,
      dueAtMs,
      name: String(item?.name || "Assignment").trim() || "Assignment",
      courseName: String(item?.course?.name || "").trim(),
      url: String(item?.url || "").trim(),
    });
  }
  return Array.from(byItemKey.values());
}

function buildGoogleEventBody(item) {
  const dueAtIso = new Date(item.dueAtMs).toISOString();
  const endIso = new Date(item.dueAtMs + 30 * 60 * 1000).toISOString();
  const details = [];
  if (item.courseName) details.push(`Course: ${item.courseName}`);
  if (item.url) details.push(`Canvas: ${item.url}`);
  details.push("Created by QuickCanvas");

  const body = {
    summary: `${item.name}`,
    description: details.join("\n"),
    start: { dateTime: dueAtIso },
    end: { dateTime: endIso },
    reminders: {
      useDefault: false,
      overrides: [{ method: "popup", minutes: 60 }],
    },
    extendedProperties: {
      private: {
        quickcanvas_item_key: item.itemKey,
      },
    },
  };

  if (item.url) {
    body.source = { title: "Open in Canvas", url: item.url };
  }
  return body;
}

async function syncGoogleCalendarEvents(assignments = []) {
  const settings = await chrome.storage.sync.get(CALENDAR_REMINDER_KEY);
  if (!settings?.[CALENDAR_REMINDER_KEY]) {
    return { synced: 0, connected: false, enabled: false };
  }

  const auth = await verifyGoogleCalendarConnection(false);
  if (!auth.connected || !auth.token) {
    return { synced: 0, connected: false, enabled: true };
  }

  const normalized = normalizeAssignments(assignments);
  const desiredByKey = new Map(normalized.map((item) => [item.itemKey, item]));
  const stored = await chrome.storage.local.get(GOOGLE_EVENT_MAP_KEY);
  const prevMap =
    stored?.[GOOGLE_EVENT_MAP_KEY] &&
    typeof stored[GOOGLE_EVENT_MAP_KEY] === "object"
      ? stored[GOOGLE_EVENT_MAP_KEY]
      : {};
  const nextMap = {};
  let synced = 0;

  for (const item of normalized) {
    const signature = eventSignatureFor(item);
    const existing = prevMap[item.itemKey];

    if (existing?.eventId && existing?.signature === signature) {
      nextMap[item.itemKey] = existing;
      continue;
    }

    const payload = buildGoogleEventBody(item);
    let eventId = existing?.eventId || "";

    if (eventId) {
      try {
        await googleCalendarRequest(
          `/calendars/primary/events/${encodeURIComponent(eventId)}`,
          auth.token,
          { method: "PATCH", body: payload },
        );
      } catch (error) {
        if (error?.status !== 404) throw error;
        eventId = "";
      }
    }

    if (!eventId) {
      const created = await googleCalendarRequest(
        "/calendars/primary/events",
        auth.token,
        {
          method: "POST",
          body: payload,
        },
      );
      eventId = String(created?.id || "").trim();
    }

    if (eventId) {
      nextMap[item.itemKey] = {
        eventId,
        signature,
        dueAtMs: item.dueAtMs,
      };
      synced += 1;
    }
  }

  for (const [itemKey, entry] of Object.entries(prevMap)) {
    if (desiredByKey.has(itemKey)) continue;
    const staleEventId = String(entry?.eventId || "").trim();
    if (!staleEventId) continue;
    try {
      await googleCalendarRequest(
        `/calendars/primary/events/${encodeURIComponent(staleEventId)}`,
        auth.token,
        { method: "DELETE" },
      );
    } catch (error) {
      // If an event was deleted manually, ignore that cleanup failure.
    }
  }

  await chrome.storage.local.set({ [GOOGLE_EVENT_MAP_KEY]: nextMap });
  return { synced, connected: true, enabled: true };
}

async function ensureDefaults() {
  try {
    const stored = await chrome.storage.sync.get([
      "canvasSettings",
      CALENDAR_REMINDER_KEY,
    ]);
    const canvasSettings =
      stored?.canvasSettings && typeof stored.canvasSettings === "object"
        ? stored.canvasSettings
        : null;
    const payload = {};
    if (!canvasSettings) {
      payload.canvasSettings = DEFAULT_SETTINGS;
    } else {
      payload.canvasSettings = { ...DEFAULT_SETTINGS, ...canvasSettings };
    }
    if (typeof stored?.[CALENDAR_REMINDER_KEY] !== "boolean") {
      payload[CALENDAR_REMINDER_KEY] = false;
    }
    if (Object.keys(payload).length > 0) {
      await chrome.storage.sync.set(payload);
    }
  } catch (error) {
    // Ignore initialization issues to avoid hard-failing service worker startup.
  }
}

function resolveNextLink(linkHeader, baseUrl) {
  const raw = String(linkHeader || "");
  if (!raw) return "";
  const parts = raw.split(",");
  for (const part of parts) {
    if (!/rel\s*=\s*"next"/i.test(part)) continue;
    const match = part.match(/<([^>]+)>/);
    if (!match?.[1]) continue;
    const href = match[1].trim();
    if (/^https?:\/\//i.test(href)) return href;
    if (href.startsWith("/")) return `${baseUrl}${href}`;
    return `${baseUrl}/${href}`;
  }
  return "";
}

function toAbsoluteCanvasUrl(rawUrl, baseUrl) {
  const value = String(rawUrl || "").trim();
  if (!value) return "";
  if (/^https?:\/\//i.test(value)) return value;
  if (!baseUrl) return "";
  if (value.startsWith("/")) return `${baseUrl}${value}`;
  return `${baseUrl}/${value}`;
}

function plannerItemToReminderAssignment(item, baseUrl) {
  const dueAt =
    item?.plannable?.due_at ||
    item?.assignment?.due_at ||
    item?.quiz?.due_at ||
    item?.due_at ||
    "";
  if (!dueAt) return null;

  const plannableId =
    item?.plannable_id ||
    item?.plannable?.id ||
    item?.assignment?.id ||
    item?.id;
  const courseId =
    item?.course_id ||
    item?.context_id ||
    item?.plannable?.course_id ||
    item?.assignment?.course_id ||
    "";
  const itemKey = `${courseId || "course"}:${plannableId || dueAt}`;
  const name =
    item?.plannable?.title ||
    item?.plannable?.name ||
    item?.assignment?.name ||
    item?.title ||
    item?.name ||
    "Assignment";
  const rawUrl =
    item?.html_url ||
    item?.plannable?.html_url ||
    item?.assignment?.html_url ||
    "";
  const url = toAbsoluteCanvasUrl(rawUrl, baseUrl);
  const courseName =
    item?.context_name ||
    item?.course_name ||
    item?.plannable?.context_name ||
    "";

  return {
    item_key: itemKey,
    name: String(name || "Assignment"),
    due_at: String(dueAt),
    url: String(url || ""),
    course: { name: String(courseName || "") },
  };
}

async function loadCanvasAssignmentsFromApi() {
  const stored = await chrome.storage.sync.get([
    "canvasSettings",
    CALENDAR_REMINDER_KEY,
  ]);
  const enabled = Boolean(stored?.[CALENDAR_REMINDER_KEY]);
  if (!enabled) return [];

  const settings =
    stored?.canvasSettings && typeof stored.canvasSettings === "object"
      ? stored.canvasSettings
      : null;
  const baseUrl = String(settings?.baseUrl || "")
    .trim()
    .replace(/\/+$/, "");
  const apiToken = String(settings?.apiToken || "").trim();
  const dashboardEnabled = settings?.enabled !== false;

  if (!dashboardEnabled || !baseUrl || !apiToken) {
    return [];
  }

  const start = new Date().toISOString();
  const end = new Date(Date.now() + MAX_DUE_LOOKAHEAD_MS).toISOString();
  const params = new URLSearchParams();
  params.set("start_date", start);
  params.set("end_date", end);
  params.set("per_page", "100");
  params.append("include[]", "submission");
  params.append("include[]", "submissions");

  let nextUrl = `${baseUrl}/api/v1/planner/items?${params.toString()}`;
  let pages = 0;
  const allItems = [];

  while (nextUrl && pages < 6) {
    const response = await fetch(nextUrl, {
      headers: {
        Authorization: `Bearer ${apiToken}`,
      },
    });
    if (!response.ok) {
      const body = await response.text();
      throw new Error(
        `Canvas API ${response.status}: ${(body || "request failed").slice(0, 220)}`,
      );
    }
    const pageItems = await response.json();
    if (Array.isArray(pageItems)) {
      allItems.push(...pageItems);
    }
    nextUrl = resolveNextLink(response.headers.get("link"), baseUrl);
    pages += 1;
  }

  return allItems
    .map((item) => plannerItemToReminderAssignment(item, baseUrl))
    .filter((item) => item && item.due_at);
}

async function syncFromCanvasApi() {
  const assignments = await loadCanvasAssignmentsFromApi();
  const result = await syncReminderAlarms(assignments);
  return {
    ok: true,
    source: "canvas_api",
    fetchedAssignments: assignments.length,
    ...result,
  };
}

async function ensureRefreshAlarm() {
  try {
    await chrome.alarms.create(REFRESH_ALARM, {
      delayInMinutes: 2,
      periodInMinutes: REFRESH_INTERVAL_MINUTES,
    });
  } catch (error) {
    // ignore refresh alarm setup failures
  }
}

async function clearReminderAlarms() {
  try {
    const alarms = await chrome.alarms.getAll();
    const toRemove = alarms.filter((alarm) =>
      alarm.name.startsWith(ALARM_PREFIX),
    );
    await Promise.allSettled(
      toRemove.map((alarm) => chrome.alarms.clear(alarm.name)),
    );
    await chrome.storage.local.remove(ALARM_MAP_KEY);
  } catch (error) {
    // ignore cleanup errors
  }
}

async function syncReminderAlarms(assignments = []) {
  const settings = await chrome.storage.sync.get(CALENDAR_REMINDER_KEY);
  if (!settings?.[CALENDAR_REMINDER_KEY]) {
    await clearReminderAlarms();
    return {
      synced: 0,
      enabled: false,
      calendarSynced: 0,
      calendarConnected: false,
    };
  }

  const normalized = normalizeAssignments(assignments);
  const now = Date.now();
  const nextMap = {};

  for (const item of normalized) {
    const alarmAt = item.dueAtMs - REMINDER_LEAD_MS;
    if (alarmAt <= now) continue;
    const name = alarmNameFor(item.itemKey, item.dueAtMs);
    nextMap[name] = {
      ...item,
      alarmAt,
    };
    await chrome.alarms.create(name, { when: alarmAt });
  }

  const existing = await chrome.alarms.getAll();
  const stale = existing.filter(
    (alarm) => alarm.name.startsWith(ALARM_PREFIX) && !nextMap[alarm.name],
  );
  await Promise.allSettled(
    stale.map((alarm) => chrome.alarms.clear(alarm.name)),
  );
  await chrome.storage.local.set({ [ALARM_MAP_KEY]: nextMap });

  let calendarResult = {
    synced: 0,
    connected: false,
    enabled: false,
    error: "",
  };
  try {
    calendarResult = await syncGoogleCalendarEvents(assignments);
  } catch (error) {
    calendarResult = {
      synced: 0,
      connected: false,
      enabled: true,
      error: String(error?.message || error),
    };
  }

  return {
    synced: Object.keys(nextMap).length,
    enabled: true,
    calendarSynced: calendarResult.synced || 0,
    calendarConnected: Boolean(calendarResult.connected),
  };
}

function formatDueLabel(dueAtMs) {
  try {
    const due = new Date(dueAtMs);
    return due.toLocaleString([], {
      month: "short",
      day: "numeric",
      hour: "numeric",
      minute: "2-digit",
    });
  } catch (error) {
    return "soon";
  }
}

chrome.runtime.onInstalled.addListener(() => {
  void ensureDefaults();
  void ensureRefreshAlarm();
  void syncFromCanvasApi();
});

chrome.runtime.onStartup.addListener(() => {
  void ensureDefaults();
  void ensureRefreshAlarm();
  void syncFromCanvasApi();
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message !== "object") return;

  if (message.type === "cfe-sync-reminder-assignments") {
    void (async () => {
      try {
        const result = await syncReminderAlarms(
          Array.isArray(message.assignments) ? message.assignments : [],
        );
        sendResponse({ ok: true, ...result });
      } catch (error) {
        sendResponse({ ok: false, error: String(error?.message || error) });
      }
    })();
    return true;
  }

  if (message.type === "cfe-google-calendar-status") {
    void (async () => {
      try {
        const status = await verifyGoogleCalendarConnection(false);
        sendResponse({ ok: true, connected: Boolean(status.connected) });
      } catch (error) {
        sendResponse({
          ok: false,
          connected: false,
          error: String(error?.message || error),
        });
      }
    })();
    return true;
  }

  if (message.type === "cfe-google-calendar-connect") {
    void (async () => {
      try {
        const status = await verifyGoogleCalendarConnection(
          Boolean(message.interactive),
        );
        sendResponse({ ok: true, connected: Boolean(status.connected) });
      } catch (error) {
        sendResponse({
          ok: false,
          connected: false,
          error: String(error?.message || error),
        });
      }
    })();
    return true;
  }

  if (message.type === "cfe-sync-reminders-from-canvas") {
    void (async () => {
      try {
        const result = await syncFromCanvasApi();
        sendResponse(result);
      } catch (error) {
        sendResponse({ ok: false, error: String(error?.message || error) });
      }
    })();
    return true;
  }

  if (message.type === "cfe-reminders-settings-updated") {
    void (async () => {
      try {
        if (message.enabled === false) {
          await clearReminderAlarms();
        } else if (message.enabled === true) {
          await syncFromCanvasApi();
        }
        sendResponse({ ok: true });
      } catch (error) {
        sendResponse({ ok: false, error: String(error?.message || error) });
      }
    })();
    return true;
  }
});

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== "sync") return;
  if (!Object.prototype.hasOwnProperty.call(changes, CALENDAR_REMINDER_KEY))
    return;
  const next = Boolean(changes[CALENDAR_REMINDER_KEY]?.newValue);
  if (!next) {
    void clearReminderAlarms();
  } else {
    void syncFromCanvasApi();
  }
});

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm?.name === REFRESH_ALARM) {
    try {
      await syncFromCanvasApi();
    } catch (error) {
      // ignore periodic sync failures
    }
    return;
  }
  if (!alarm?.name?.startsWith(ALARM_PREFIX)) return;
  try {
    const stored = await chrome.storage.local.get(ALARM_MAP_KEY);
    const map =
      stored?.[ALARM_MAP_KEY] && typeof stored[ALARM_MAP_KEY] === "object"
        ? stored[ALARM_MAP_KEY]
        : {};
    const item = map?.[alarm.name];
    if (!item) return;
    const message = item.courseName
      ? `${item.courseName} - due ${formatDueLabel(item.dueAtMs)}`
      : `Due ${formatDueLabel(item.dueAtMs)}`;
    await chrome.notifications.create(alarm.name, {
      type: "basic",
      iconUrl: "icons/icon128.png",
      title: item.name || "Assignment due soon",
      message,
      priority: 2,
    });
  } catch (error) {
    // ignore notification errors
  }
});

chrome.notifications.onClicked.addListener(async (notificationId) => {
  if (!notificationId?.startsWith(ALARM_PREFIX)) return;
  try {
    const stored = await chrome.storage.local.get(ALARM_MAP_KEY);
    const map =
      stored?.[ALARM_MAP_KEY] && typeof stored[ALARM_MAP_KEY] === "object"
        ? stored[ALARM_MAP_KEY]
        : {};
    const url = String(map?.[notificationId]?.url || "").trim();
    if (url) {
      await chrome.tabs.create({ url });
    }
    await chrome.notifications.clear(notificationId);
  } catch (error) {
    // ignore click/open errors
  }
});
