/* global Office */
/*
 * Case Tagger task pane.
 *
 * Always-visible (when pinned) compose-time UI that lives next to the subject
 * line. Picks a case from an externally-editable JSON file, prepends the case
 * name to the subject in a configurable format, and remembers the choice on
 * the mail item via custom properties so the OnMessageSend handler can verify
 * the user picked something before letting the message go out.
 */

(function () {
    "use strict";

    // ---- Storage keys (custom properties on the mail item) -----------------
    var CP_CASE_ID = "caseTagger.caseId";
    var CP_CASE_NAME = "caseTagger.caseName";
    var CP_NOT_CASE = "caseTagger.notCaseRelated";
    var CP_ORIGINAL_SUBJECT = "caseTagger.originalSubject";

    // ---- Roaming settings keys (per-mailbox prefs) -------------------------
    var SETTING_MRU = "caseTagger.mru";          // [{id, ts}]
    var SETTING_PINNED = "caseTagger.pinned";    // [id, id, ...]
    var SETTING_SORT = "caseTagger.sort";        // "pinned-mru" | "mru" | "alpha" | "manual"

    var state = {
        config: null,
        cases: [],
        filtered: [],
        mru: [],
        pinned: [],
        sort: "pinned-mru",
        selectedId: null,
        notCaseRelated: false,
        item: null,
        customProps: null,
        searchTerm: ""
    };

    Office.onReady(function (info) {
        if (info.host !== Office.HostType.Outlook) {
            showStatus("Case Tagger only runs in Outlook.", "error");
            return;
        }
        state.item = Office.context.mailbox.item;

        bindUi();
        loadAll().catch(function (err) {
            console.error(err);
            showStatus("Failed to initialize: " + (err && err.message ? err.message : err), "error");
        });
    });

    // ------------------------------------------------------------------
    // Bootstrap
    // ------------------------------------------------------------------

    function loadAll() {
        return Promise.all([
            loadConfigAndCases(),
            loadCustomProperties(),
            loadRoamingSettings()
        ]).then(function () {
            // Hide the loading shim, show the app.
            document.getElementById("loading").hidden = true;
            document.getElementById("app").hidden = false;

            renderSortControl();
            applySortAndRender();
            refreshCurrentTagDisplay();
            refreshSubjectPreview();
        });
    }

    function loadConfigAndCases() {
        // Config first (so we know where cases.json lives).
        return fetchJson("../../config/config.json").then(function (cfg) {
            state.config = normalizeConfig(cfg);
            state.sort = state.sort || state.config.defaultSort;
            return fetchJson(state.config.casesFile);
        }).then(function (raw) {
            state.cases = normalizeCases(raw);
        });
    }

    function loadCustomProperties() {
        return new Promise(function (resolve, reject) {
            state.item.loadCustomPropertiesAsync(function (asyncResult) {
                if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(asyncResult.error);
                    return;
                }
                state.customProps = asyncResult.value;
                state.selectedId = state.customProps.get(CP_CASE_ID) || null;
                state.notCaseRelated = state.customProps.get(CP_NOT_CASE) === true;
                resolve();
            });
        });
    }

    function loadRoamingSettings() {
        var settings = Office.context.roamingSettings;
        state.mru = settings.get(SETTING_MRU) || [];
        state.pinned = settings.get(SETTING_PINNED) || [];
        state.sort = settings.get(SETTING_SORT) || state.sort || "pinned-mru";
        return Promise.resolve();
    }

    // ------------------------------------------------------------------
    // Config & cases parsing
    // ------------------------------------------------------------------

    function normalizeConfig(cfg) {
        cfg = cfg || {};
        return {
            casesFile: cfg.casesFile || "../../config/cases.json",
            tagFormat: typeof cfg.tagFormat === "string" ? cfg.tagFormat : "{case}: {subject}",
            tagDetectRegex: typeof cfg.tagDetectRegex === "string" ? cfg.tagDetectRegex : "",
            defaultSort: cfg.defaultSort || "pinned-mru",
            allowNotCaseRelated: cfg.allowNotCaseRelated !== false,
            mruLimit: typeof cfg.mruLimit === "number" ? cfg.mruLimit : 25
        };
    }

    function normalizeCases(raw) {
        // Accept either a bare array or {cases: [...]}.
        var arr = Array.isArray(raw) ? raw : (raw && Array.isArray(raw.cases) ? raw.cases : []);
        return arr.map(function (c, idx) {
            if (typeof c === "string") {
                return { id: slug(c), name: c, order: idx };
            }
            return {
                id: c.id || slug(c.name || ("case-" + idx)),
                name: c.name || c.id || ("Case " + (idx + 1)),
                aliases: Array.isArray(c.aliases) ? c.aliases : [],
                order: typeof c.order === "number" ? c.order : idx
            };
        });
    }

    function slug(s) {
        return String(s).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
    }

    // ------------------------------------------------------------------
    // Rendering
    // ------------------------------------------------------------------

    function bindUi() {
        document.getElementById("caseSearch").addEventListener("input", function (e) {
            state.searchTerm = e.target.value || "";
            applySortAndRender();
        });

        document.getElementById("caseSearch").addEventListener("keydown", function (e) {
            if (e.key === "Enter" && state.filtered.length > 0) {
                e.preventDefault();
                selectCase(state.filtered[0].id);
            }
        });

        document.getElementById("sortMode").addEventListener("change", function (e) {
            state.sort = e.target.value;
            Office.context.roamingSettings.set(SETTING_SORT, state.sort);
            Office.context.roamingSettings.saveAsync(function () {});
            applySortAndRender();
        });

        document.getElementById("clearTagBtn").addEventListener("click", function () {
            clearSelection();
        });

        document.getElementById("markNotCaseBtn").addEventListener("click", function () {
            toggleNotCaseRelated();
        });
    }

    function renderSortControl() {
        document.getElementById("sortMode").value = state.sort;
    }

    function applySortAndRender() {
        var term = state.searchTerm.trim().toLowerCase();
        var list = state.cases.slice();

        if (term) {
            list = list.filter(function (c) {
                if (c.name.toLowerCase().indexOf(term) !== -1) return true;
                if (c.aliases) {
                    for (var i = 0; i < c.aliases.length; i++) {
                        if (String(c.aliases[i]).toLowerCase().indexOf(term) !== -1) return true;
                    }
                }
                return false;
            });
        }

        list = sortCases(list, state.sort);
        state.filtered = list;
        renderCases(list);
    }

    function sortCases(list, mode) {
        var mruIndex = {};
        state.mru.forEach(function (entry, idx) { mruIndex[entry.id] = idx; });
        var pinnedSet = {};
        state.pinned.forEach(function (id) { pinnedSet[id] = true; });

        function alpha(a, b) { return a.name.localeCompare(b.name); }
        function mruCmp(a, b) {
            var ai = mruIndex[a.id], bi = mruIndex[b.id];
            if (ai === undefined && bi === undefined) return alpha(a, b);
            if (ai === undefined) return 1;
            if (bi === undefined) return -1;
            return ai - bi;
        }

        switch (mode) {
            case "alpha":
                return list.sort(alpha);
            case "mru":
                return list.sort(mruCmp);
            case "manual":
                return list.sort(function (a, b) { return a.order - b.order; });
            case "pinned-mru":
            default:
                return list.sort(function (a, b) {
                    var ap = pinnedSet[a.id] ? 0 : 1;
                    var bp = pinnedSet[b.id] ? 0 : 1;
                    if (ap !== bp) return ap - bp;
                    return mruCmp(a, b);
                });
        }
    }

    function renderCases(list) {
        var ul = document.getElementById("caseList");
        ul.innerHTML = "";

        if (list.length === 0) {
            var empty = document.createElement("li");
            empty.className = "ct-list-item ct-meta";
            empty.textContent = "No cases match.";
            ul.appendChild(empty);
            return;
        }

        var pinnedSet = {};
        state.pinned.forEach(function (id) { pinnedSet[id] = true; });
        var mruIndex = {};
        state.mru.forEach(function (e, i) { mruIndex[e.id] = i; });

        list.forEach(function (c) {
            var li = document.createElement("li");
            li.className = "ct-list-item";
            if (c.id === state.selectedId) li.classList.add("ct-selected");
            li.setAttribute("role", "option");
            li.dataset.caseId = c.id;

            var name = document.createElement("span");
            name.className = "ct-name";
            name.textContent = c.name;
            li.appendChild(name);

            if (mruIndex[c.id] !== undefined) {
                var meta = document.createElement("span");
                meta.className = "ct-meta";
                meta.textContent = mruIndex[c.id] === 0 ? "last used" : "#" + (mruIndex[c.id] + 1) + " recent";
                li.appendChild(meta);
            }

            var pin = document.createElement("button");
            pin.type = "button";
            pin.className = "ct-pin-btn";
            pin.title = pinnedSet[c.id] ? "Unpin" : "Pin to top";
            pin.setAttribute("aria-pressed", pinnedSet[c.id] ? "true" : "false");
            pin.textContent = pinnedSet[c.id] ? "★" : "☆";
            pin.addEventListener("click", function (e) {
                e.stopPropagation();
                togglePin(c.id);
            });
            li.appendChild(pin);

            li.addEventListener("click", function () { selectCase(c.id); });

            ul.appendChild(li);
        });
    }

    function refreshCurrentTagDisplay() {
        var tag = document.getElementById("currentTag");
        var clearBtn = document.getElementById("clearTagBtn");
        var notCaseBtn = document.getElementById("markNotCaseBtn");

        if (state.notCaseRelated) {
            tag.textContent = "(marked: not case related)";
            clearBtn.hidden = false;
            notCaseBtn.classList.add("ct-active-toggle");
            notCaseBtn.textContent = "Not case related ✓";
            return;
        }

        notCaseBtn.classList.remove("ct-active-toggle");
        notCaseBtn.textContent = "Not case related";

        if (state.selectedId) {
            var c = findCaseById(state.selectedId);
            tag.textContent = c ? c.name : state.selectedId;
            clearBtn.hidden = false;
        } else {
            tag.textContent = "— none —";
            clearBtn.hidden = true;
        }
    }

    // ------------------------------------------------------------------
    // Subject management
    // ------------------------------------------------------------------

    function refreshSubjectPreview() {
        getSubject().then(function (subject) {
            document.getElementById("subjectPreview").value = subject || "";
        }).catch(function () { /* ignore */ });
    }

    function getSubject() {
        return new Promise(function (resolve, reject) {
            state.item.subject.getAsync(function (r) {
                if (r.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(r.error);
                    return;
                }
                resolve(r.value || "");
            });
        });
    }

    function setSubject(value) {
        return new Promise(function (resolve, reject) {
            state.item.subject.setAsync(value, function (r) {
                if (r.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(r.error);
                    return;
                }
                resolve();
            });
        });
    }

    function buildTaggedSubject(caseName, baseSubject) {
        return state.config.tagFormat
            .replace("{case}", caseName)
            .replace("{subject}", baseSubject || "");
    }

    function stripExistingTagFor(caseName, subject) {
        // Strip a tag that matches the configured format for this exact case.
        if (!subject) return "";
        var prefix = buildTaggedSubject(caseName, "");
        if (subject.indexOf(prefix) === 0) {
            return subject.slice(prefix.length);
        }
        return subject;
    }

    function stripAnyKnownTag(subject) {
        if (!subject) return "";
        // 1) Format-derived prefix for any known case.
        for (var i = 0; i < state.cases.length; i++) {
            var pre = buildTaggedSubject(state.cases[i].name, "");
            if (pre && subject.indexOf(pre) === 0) {
                return subject.slice(pre.length);
            }
        }
        // 2) Optional regex fallback (config.tagDetectRegex).
        if (state.config.tagDetectRegex) {
            try {
                var rx = new RegExp(state.config.tagDetectRegex);
                return subject.replace(rx, "");
            } catch (e) { /* ignore bad regex */ }
        }
        return subject;
    }

    // ------------------------------------------------------------------
    // Selection / state changes
    // ------------------------------------------------------------------

    function selectCase(caseId) {
        var c = findCaseById(caseId);
        if (!c) return;

        getSubject().then(function (current) {
            var originalStored = state.customProps.get(CP_ORIGINAL_SUBJECT);
            var base;

            if (state.selectedId) {
                // Replace existing tag.
                var prevCase = findCaseById(state.selectedId);
                base = prevCase ? stripExistingTagFor(prevCase.name, current) : stripAnyKnownTag(current);
            } else {
                base = stripAnyKnownTag(current);
                if (originalStored === undefined || originalStored === null) {
                    state.customProps.set(CP_ORIGINAL_SUBJECT, base);
                }
            }

            var newSubject = buildTaggedSubject(c.name, base);
            return setSubject(newSubject).then(function () {
                state.selectedId = c.id;
                state.notCaseRelated = false;
                state.customProps.set(CP_CASE_ID, c.id);
                state.customProps.set(CP_CASE_NAME, c.name);
                state.customProps.set(CP_NOT_CASE, false);
                saveCustomProps();
                bumpMru(c.id);
                refreshCurrentTagDisplay();
                refreshSubjectPreview();
                applySortAndRender();
                showStatus("Tagged as " + c.name + ".", "ok");
            });
        }).catch(function (err) {
            console.error(err);
            showStatus("Could not update subject: " + (err && err.message ? err.message : err), "error");
        });
    }

    function clearSelection() {
        getSubject().then(function (current) {
            var stripped = current;
            if (state.selectedId) {
                var prev = findCaseById(state.selectedId);
                stripped = prev ? stripExistingTagFor(prev.name, current) : stripAnyKnownTag(current);
            } else {
                stripped = stripAnyKnownTag(current);
            }
            return setSubject(stripped).then(function () {
                state.selectedId = null;
                state.notCaseRelated = false;
                state.customProps.set(CP_CASE_ID, null);
                state.customProps.set(CP_CASE_NAME, null);
                state.customProps.set(CP_NOT_CASE, false);
                saveCustomProps();
                refreshCurrentTagDisplay();
                refreshSubjectPreview();
                applySortAndRender();
                showStatus("Cleared.", "ok");
            });
        }).catch(function (err) {
            console.error(err);
            showStatus("Could not clear: " + (err && err.message ? err.message : err), "error");
        });
    }

    function toggleNotCaseRelated() {
        if (!state.config.allowNotCaseRelated) {
            showStatus("'Not case related' is disabled in config.", "error");
            return;
        }
        if (state.notCaseRelated) {
            state.notCaseRelated = false;
            state.customProps.set(CP_NOT_CASE, false);
            saveCustomProps();
            refreshCurrentTagDisplay();
            showStatus("Unmarked.", "ok");
            return;
        }

        // If a case is currently selected, clear it (strip from subject) first.
        var maybeClear = state.selectedId ? clearSelectionSilently() : Promise.resolve();
        maybeClear.then(function () {
            state.notCaseRelated = true;
            state.customProps.set(CP_NOT_CASE, true);
            saveCustomProps();
            refreshCurrentTagDisplay();
            showStatus("Marked as not case related. Send won't be blocked.", "ok");
        });
    }

    function clearSelectionSilently() {
        return getSubject().then(function (current) {
            var prev = findCaseById(state.selectedId);
            var stripped = prev ? stripExistingTagFor(prev.name, current) : stripAnyKnownTag(current);
            return setSubject(stripped);
        }).then(function () {
            state.selectedId = null;
            state.customProps.set(CP_CASE_ID, null);
            state.customProps.set(CP_CASE_NAME, null);
            saveCustomProps();
        });
    }

    function togglePin(caseId) {
        var idx = state.pinned.indexOf(caseId);
        if (idx === -1) {
            state.pinned.unshift(caseId);
        } else {
            state.pinned.splice(idx, 1);
        }
        Office.context.roamingSettings.set(SETTING_PINNED, state.pinned);
        Office.context.roamingSettings.saveAsync(function () {});
        applySortAndRender();
    }

    function bumpMru(caseId) {
        var now = Date.now();
        state.mru = state.mru.filter(function (e) { return e.id !== caseId; });
        state.mru.unshift({ id: caseId, ts: now });
        if (state.mru.length > state.config.mruLimit) {
            state.mru.length = state.config.mruLimit;
        }
        Office.context.roamingSettings.set(SETTING_MRU, state.mru);
        Office.context.roamingSettings.saveAsync(function () {});
    }

    function saveCustomProps() {
        try {
            state.customProps.saveAsync(function () { /* fire and forget */ });
        } catch (e) { /* ignore */ }
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    function findCaseById(id) {
        for (var i = 0; i < state.cases.length; i++) {
            if (state.cases[i].id === id) return state.cases[i];
        }
        return null;
    }

    function fetchJson(relUrl) {
        var url = new URL(relUrl, window.location.href).toString();
        // Cache-bust so that edits to cases.json show up immediately.
        var bust = (url.indexOf("?") === -1 ? "?" : "&") + "_t=" + Date.now();
        return fetch(url + bust, { cache: "no-store" }).then(function (r) {
            if (!r.ok) throw new Error("HTTP " + r.status + " loading " + url);
            return r.json();
        });
    }

    function showStatus(msg, kind) {
        var el = document.getElementById("statusMsg");
        if (!el) return;
        el.textContent = msg;
        el.className = "ct-status" + (kind === "error" ? " ct-status-error" : kind === "ok" ? " ct-status-ok" : "");
    }
})();
