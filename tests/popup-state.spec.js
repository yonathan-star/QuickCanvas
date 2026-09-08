const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");
const source = fs.readFileSync(
  path.join(root, "extension", "popup.js"),
  "utf8",
);

function extractFunction(name, nextName) {
  const asyncStart = source.indexOf(`async function ${name}`);
  const start =
    asyncStart >= 0 ? asyncStart : source.indexOf(`function ${name}`);
  const end = source.indexOf(`function ${nextName}`, start);
  assert.ok(start >= 0 && end > start, `Could not extract ${name}`);
  return source.slice(start, end);
}

(async () => {
  const durableStorage = { cfePopupActiveTab: "themes" };
  const localStorageValues = new Map([["cfePopupActiveTab", "account"]]);
  const context = {
    POPUP_ACTIVE_TAB_KEY: "cfePopupActiveTab",
    chrome: {
      storage: {
        local: {
          async get(key) {
            return { [key]: durableStorage[key] };
          },
          async set(values) {
            Object.assign(durableStorage, values);
          },
        },
      },
    },
    localStorage: {
      getItem(key) {
        return localStorageValues.get(key) || null;
      },
      setItem(key, value) {
        localStorageValues.set(key, value);
      },
    },
  };
  vm.createContext(context);
  vm.runInContext(
    extractFunction("getStoredActiveTab", "storeActiveTab") +
      extractFunction("storeActiveTab", "errorMessage"),
    context,
  );

  assert.equal(await context.getStoredActiveTab(), "themes");
  context.storeActiveTab("admin");
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(durableStorage.cfePopupActiveTab, "admin");
  assert.equal(localStorageValues.get("cfePopupActiveTab"), "admin");

  assert.match(
    source,
    /const storedActiveTab = await getStoredActiveTab\(\);[\s\S]*setActiveTab\("account", \{ persist: false \}\)/,
  );
  assert.doesNotMatch(
    source,
    /communityPanel\?\.classList\.toggle\("is-expanded"/,
  );
  assert.match(source, /accountCloudPanel\.hidden = !isUiSignedIn/);
  console.log("Popup persistence and theme-access checks passed.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
