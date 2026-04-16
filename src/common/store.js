/*
 * Roaming-settings store. The case list, user settings, and MRU live
 * here once seeded. Roaming settings sync per-mailbox so the same data
 * appears on every Outlook the user signs into. ~32 KB cap is more
 * than enough for thousands of cases.
 *
 * On first launch, when the cases key is empty, we seed from the
 * static cases.json shipped alongside the add-in.
 */
(function (global) {
  "use strict";

  const CONFIG = global.CaseTaggerConfig;

  function settings() {
    return Office.context.roamingSettings;
  }

  function save() {
    return new Promise((resolve, reject) => {
      settings().saveAsync((r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(r.error ? r.error.message : "saveAsync failed"));
      });
    });
  }

  async function loadCases({ seedIfEmpty = true } = {}) {
    let cases = settings().get(CONFIG.keys.cases);
    if (!Array.isArray(cases) || cases.length === 0) {
      if (!seedIfEmpty) return [];
      cases = await seedFromFile();
      settings().set(CONFIG.keys.cases, cases);
      await save();
    }
    return cases;
  }

  async function seedFromFile() {
    try {
      const res = await fetch(CONFIG.seedCasesUrl, { cache: "no-store" });
      if (!res.ok) return [];
      const data = await res.json();
      return Array.isArray(data.cases) ? data.cases.map(normalizeCase) : [];
    } catch (e) {
      console.warn("Could not seed cases:", e);
      return [];
    }
  }

  function normalizeCase(c) {
    return {
      id: c.id || cryptoRandomId(),
      name: c.name || "(unnamed)",
      client: c.client || "",
      status: c.status === "archived" ? "archived" : "active",
      pinned: !!c.pinned,
      lastUsed: c.lastUsed || null,
      useCount: typeof c.useCount === "number" ? c.useCount : 0
    };
  }

  async function saveCases(cases) {
    settings().set(CONFIG.keys.cases, cases.map(normalizeCase));
    await save();
  }

  async function upsertCase(updated) {
    const cases = await loadCases({ seedIfEmpty: false });
    const idx = cases.findIndex((c) => c.id === updated.id);
    if (idx === -1) cases.push(normalizeCase(updated));
    else cases[idx] = normalizeCase({ ...cases[idx], ...updated });
    await saveCases(cases);
    return cases;
  }

  async function archiveCase(id) {
    const cases = await loadCases({ seedIfEmpty: false });
    const idx = cases.findIndex((c) => c.id === id);
    if (idx === -1) return cases;
    cases[idx].status = cases[idx].status === "archived" ? "active" : "archived";
    await saveCases(cases);
    return cases;
  }

  async function deleteCase(id) {
    const cases = (await loadCases({ seedIfEmpty: false })).filter((c) => c.id !== id);
    await saveCases(cases);
    return cases;
  }

  async function recordUse(id) {
    const cases = await loadCases({ seedIfEmpty: false });
    const idx = cases.findIndex((c) => c.id === id);
    if (idx !== -1) {
      cases[idx].lastUsed = new Date().toISOString();
      cases[idx].useCount = (cases[idx].useCount || 0) + 1;
      await saveCases(cases);
    }
    const mru = (settings().get(CONFIG.keys.mru) || []).filter((x) => x !== id);
    mru.unshift(id);
    if (mru.length > CONFIG.maxMruEntries) mru.length = CONFIG.maxMruEntries;
    settings().set(CONFIG.keys.mru, mru);
    await save();
  }

  function getMru() {
    return settings().get(CONFIG.keys.mru) || [];
  }

  async function loadSettings() {
    const raw = settings().get(CONFIG.keys.settings) || {};
    return CONFIG.resolveSettings(raw);
  }

  async function saveSettings(patch) {
    const current = settings().get(CONFIG.keys.settings) || {};
    settings().set(CONFIG.keys.settings, { ...current, ...patch });
    await save();
    return CONFIG.resolveSettings(settings().get(CONFIG.keys.settings));
  }

  async function resetCases() {
    settings().set(CONFIG.keys.cases, null);
    settings().set(CONFIG.keys.mru, []);
    await save();
    return loadCases();
  }

  function cryptoRandomId() {
    if (global.crypto && global.crypto.randomUUID) return global.crypto.randomUUID();
    return "c_" + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
  }

  global.CaseTaggerStore = {
    loadCases,
    saveCases,
    upsertCase,
    archiveCase,
    deleteCase,
    recordUse,
    getMru,
    loadSettings,
    saveSettings,
    resetCases,
    cryptoRandomId
  };
})(typeof window !== "undefined" ? window : globalThis);
