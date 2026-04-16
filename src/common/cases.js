/*
 * Shared helpers for loading cases, remembering MRU, and computing the
 * sorted list the dropdown displays. No DOM access here - reused by both
 * the task pane UI and the send-time validator.
 */
(function (global) {
  "use strict";

  const CONFIG = global.CaseTaggerConfig;

  async function loadCases() {
    const res = await fetch(CONFIG.casesUrl, { cache: "no-store" });
    if (!res.ok) {
      throw new Error("Failed to load cases: " + res.status + " " + res.statusText);
    }
    const data = await res.json();
    const cases = Array.isArray(data.cases) ? data.cases : [];
    return cases.filter((c) => c && c.id && c.name && c.active !== false);
  }

  function readMruState() {
    const settings = Office.context.roamingSettings;
    const mru = settings.get(CONFIG.roamingKeys.mru) || [];
    const lastUsed = settings.get(CONFIG.roamingKeys.lastUsed) || {};
    return { mru: Array.isArray(mru) ? mru : [], lastUsed };
  }

  function recordUse(caseId) {
    const settings = Office.context.roamingSettings;
    const state = readMruState();
    const mru = [caseId].concat(state.mru.filter((id) => id !== caseId));
    if (mru.length > CONFIG.maxMruEntries) mru.length = CONFIG.maxMruEntries;
    state.lastUsed[caseId] = new Date().toISOString();
    settings.set(CONFIG.roamingKeys.mru, mru);
    settings.set(CONFIG.roamingKeys.lastUsed, state.lastUsed);
    return new Promise((resolve) => settings.saveAsync(() => resolve()));
  }

  function sortCases(cases, mode) {
    const { mru } = readMruState();
    const mruRank = new Map(mru.map((id, i) => [id, i]));
    const byAlpha = (a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    const byMru = (a, b) => {
      const ra = mruRank.has(a.id) ? mruRank.get(a.id) : Infinity;
      const rb = mruRank.has(b.id) ? mruRank.get(b.id) : Infinity;
      if (ra !== rb) return ra - rb;
      return byAlpha(a, b);
    };
    const byPinnedThen = (next) => (a, b) => {
      const pa = a.pinned ? 0 : 1;
      const pb = b.pinned ? 0 : 1;
      if (pa !== pb) return pa - pb;
      return next(a, b);
    };

    const copy = cases.slice();
    switch (mode) {
      case "alpha":
        return copy.sort(byAlpha);
      case "mru":
        return copy.sort(byMru);
      case "pinned-alpha":
        return copy.sort(byPinnedThen(byAlpha));
      case "pinned-mru-alpha":
      default:
        return copy.sort(byPinnedThen(byMru));
    }
  }

  function tagRegex() {
    return new RegExp(CONFIG.tagPattern);
  }

  function stripExistingTag(subject) {
    if (!subject) return "";
    return subject.replace(tagRegex(), "");
  }

  function buildTag(caseName) {
    return CONFIG.tagFormat.replace("{case}", caseName);
  }

  function subjectHasTag(subject, knownCaseNames) {
    if (!subject) return false;
    const m = subject.match(tagRegex());
    if (!m) return false;
    if (!knownCaseNames || !knownCaseNames.length) return true;
    const candidate = (m[1] || "").trim();
    return knownCaseNames.some((n) => n.trim() === candidate);
  }

  global.CaseTagger = {
    loadCases,
    readMruState,
    recordUse,
    sortCases,
    stripExistingTag,
    buildTag,
    subjectHasTag,
    tagRegex
  };
})(typeof window !== "undefined" ? window : globalThis);
