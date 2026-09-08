const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(
  path.resolve(__dirname, "..", "extension", "background.js"),
  "utf8",
);

const listeners = { installed: [], startup: [], message: [], storage: [], alarm: [], notification: [] };
const stores = {
  sync: {
    canvasSettings: { enabled: true, baseUrl: "https://canvas.test", apiToken: "legacy-secret" },
    cfeCalendarRemindersEnabled: true,
  },
  local: {},
};
const alarms = new Map();
const notifications = new Map();
const openedTabs = [];

const select = (store, keys) => {
  if (keys == null) return { ...store };
  if (typeof keys === "string") return { [keys]: store[keys] };
  if (Array.isArray(keys)) return Object.fromEntries(keys.map((key) => [key, store[key]]));
  return Object.fromEntries(Object.entries(keys).map(([key, fallback]) => [key, store[key] ?? fallback]));
};
const area = (name) => ({
  get: async (keys) => select(stores[name], keys),
  set: async (values) => Object.assign(stores[name], values),
  remove: async (keys) => [].concat(keys || []).forEach((key) => delete stores[name][key]),
});

const chrome = {
  runtime: {
    lastError: null,
    onInstalled: { addListener: (fn) => listeners.installed.push(fn) },
    onStartup: { addListener: (fn) => listeners.startup.push(fn) },
    onMessage: { addListener: (fn) => listeners.message.push(fn) },
  },
  storage: {
    sync: area("sync"),
    local: area("local"),
    onChanged: { addListener: (fn) => listeners.storage.push(fn) },
  },
  alarms: {
    create: async (name, details) => alarms.set(name, { name, ...details }),
    clear: async (name) => alarms.delete(name),
    getAll: async () => Array.from(alarms.values()),
    onAlarm: { addListener: (fn) => listeners.alarm.push(fn) },
  },
  identity: {
    getAuthToken: (_options, callback) => {
      chrome.runtime.lastError = { message: "Not connected" };
      callback("");
      chrome.runtime.lastError = null;
    },
    removeCachedAuthToken: (_details, callback) => callback(),
  },
  notifications: {
    create: async (id, details) => notifications.set(id, details),
    clear: async (id) => notifications.delete(id),
    onClicked: { addListener: (fn) => listeners.notification.push(fn) },
  },
  tabs: {
    query: async () => [],
    sendMessage: async () => ({ ok: false }),
    create: async (details) => openedTabs.push(details),
  },
};

vm.runInNewContext(source, {
  chrome,
  fetch: async () => ({ ok: true, status: 200, text: async () => "{}" }),
  console,
  Date,
  Intl,
  URL,
  encodeURIComponent,
  setTimeout,
  clearTimeout,
});

async function send(message) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`Message timed out: ${message.type}`)), 2000);
    const keepAlive = listeners.message[0](message, {}, (value) => {
      clearTimeout(timer);
      resolve(value);
    });
    if (keepAlive !== true) reject(new Error(`Message was not kept alive: ${message.type}`));
  });
}

(async () => {
  listeners.installed[0]();
  await new Promise((resolve) => setTimeout(resolve, 20));
  assert.equal("apiToken" in stores.sync.canvasSettings, false, "legacy Canvas token was retained");
  assert.ok(alarms.has("cfe-calendar-refresh"), "periodic refresh alarm was not created");

  const dueAt = new Date(Date.now() + 3 * 60 * 60 * 1000).toISOString();
  const result = await send({
    type: "cfe-sync-reminder-assignments",
    assignments: [{
      item_key: "1:77",
      name: "Lab report",
      due_at: dueAt,
      url: "https://canvas.test/courses/1/assignments/77",
      course: { name: "Biology" },
    }],
  });
  assert.equal(result.ok, true);
  assert.equal(result.synced, 1);
  const dueAlarm = Array.from(alarms.values()).find((alarm) => alarm.name.startsWith("cfe-due-"));
  assert.ok(dueAlarm, "assignment alarm was not created");

  await listeners.alarm[0](dueAlarm);
  assert.equal(notifications.get(dueAlarm.name).title, "Lab report");
  await listeners.notification[0](dueAlarm.name);
  assert.equal(openedTabs.length, 1);
  assert.equal(openedTabs[0].url, "https://canvas.test/courses/1/assignments/77");
  assert.equal(notifications.has(dueAlarm.name), false);

  stores.sync.cfeCalendarRemindersEnabled = false;
  const disabled = await send({ type: "cfe-reminders-settings-updated", enabled: false });
  assert.equal(disabled.ok, true);
  assert.equal(Array.from(alarms.keys()).some((name) => name.startsWith("cfe-due-")), false);
  console.log("Background reminder regression checks passed.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
