/*
 * Pure helpers: sort, filter, tag building. No Office.js calls so this
 * module is trivially testable and reusable from both UI and event
 * handlers.
 */
(function (global) {
  "use strict";

  function activeOnly(cases) {
    return cases.filter((c) => c.status !== "archived");
  }

  function filterByQuery(cases, query) {
    if (!query) return cases;
    const q = query.trim().toLowerCase();
    if (!q) return cases;
    return cases.filter((c) =>
      (c.name || "").toLowerCase().includes(q) ||
      (c.client || "").toLowerCase().includes(q)
    );
  }

  function sortCases(cases, mode, mru) {
    const mruRank = new Map((mru || []).map((id, i) => [id, i]));
    const byAlpha = (a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    const byMru = (a, b) => {
      const ra = mruRank.has(a.id) ? mruRank.get(a.id) : Infinity;
      const rb = mruRank.has(b.id) ? mruRank.get(b.id) : Infinity;
      if (ra !== rb) return ra - rb;
      return byAlpha(a, b);
    };
    const byFreq = (a, b) => {
      const fa = a.useCount || 0;
      const fb = b.useCount || 0;
      if (fa !== fb) return fb - fa;
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
      case "alpha":           return copy.sort(byAlpha);
      case "mru":             return copy.sort(byMru);
      case "freq":            return copy.sort(byFreq);
      case "pinned-alpha":    return copy.sort(byPinnedThen(byAlpha));
      case "pinned-mru-alpha":
      default:                return copy.sort(byPinnedThen(byMru));
    }
  }

  function buildTag(template, caseName) {
    return (template || "{caseName}: ").replace("{caseName}", caseName);
  }

  function tagRegex(pattern) {
    return new RegExp(pattern || "^([^:\\n]{1,120}):\\s+");
  }

  function stripExistingTag(subject, pattern) {
    if (!subject) return "";
    return subject.replace(tagRegex(pattern), "");
  }

  function subjectHasTag(subject, pattern, knownNames) {
    if (!subject) return false;
    const m = subject.match(tagRegex(pattern));
    if (!m) return false;
    if (!knownNames || !knownNames.length) return true;
    const candidate = (m[1] || "").trim();
    return knownNames.some((n) => n.trim() === candidate);
  }

  global.CaseTaggerHelpers = {
    activeOnly, filterByQuery, sortCases,
    buildTag, tagRegex, stripExistingTag, subjectHasTag
  };
})(typeof window !== "undefined" ? window : globalThis);
