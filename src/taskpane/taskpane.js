/*
 * Task pane controller. Renders the case dropdown, updates the subject
 * when the user picks a case, and records the choice in the per-item
 * custom properties so the send-time validator can see it.
 */
(function () {
  "use strict";

  const CONFIG = window.CaseTaggerConfig;
  const CT = window.CaseTagger;

  let state = {
    cases: [],
    sort: CONFIG.defaultSort,
    selectedId: null,
    customProps: null
  };

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) return;

    document.getElementById("sort-select").value = state.sort;
    document.getElementById("sort-select").addEventListener("change", onSortChange);
    document.getElementById("case-select").addEventListener("change", onCaseChange);
    document.getElementById("refresh-btn").addEventListener("click", () => init(true));

    init(false);
  });

  async function init(forceReload) {
    setStatus("Loading cases…");
    try {
      const [cases, props] = await Promise.all([
        CT.loadCases(),
        getCustomProps()
      ]);
      state.cases = cases;
      state.customProps = props;

      const stored = props.get(CONFIG.customPropertyKey);
      state.selectedId = stored || null;

      renderOptions();
      if (state.selectedId) {
        document.getElementById("case-select").value = state.selectedId;
      }
      await updatePreview();
      setStatus(forceReload ? "Case list reloaded." : "", forceReload ? "success" : "");
    } catch (err) {
      console.error(err);
      setStatus("Could not load cases: " + err.message, "error");
    }
  }

  function getCustomProps() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error ? result.error.message : "loadCustomPropertiesAsync failed"));
      });
    });
  }

  function saveCustomProps() {
    return new Promise((resolve, reject) => {
      state.customProps.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error ? result.error.message : "saveAsync failed"));
      });
    });
  }

  function renderOptions() {
    const sel = document.getElementById("case-select");
    sel.innerHTML = "";

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Select a case…";
    placeholder.disabled = true;
    placeholder.selected = !state.selectedId;
    sel.appendChild(placeholder);

    const notRelated = document.createElement("option");
    notRelated.value = CONFIG.notCaseRelatedId;
    notRelated.textContent = CONFIG.notCaseRelatedLabel;
    sel.appendChild(notRelated);

    const sorted = CT.sortCases(state.cases, state.sort);
    const { mru } = CT.readMruState();
    const mruSet = new Set(mru);

    const pinnedGroup = document.createElement("optgroup");
    pinnedGroup.label = "Pinned";
    const recentGroup = document.createElement("optgroup");
    recentGroup.label = "Recent";
    const otherGroup = document.createElement("optgroup");
    otherGroup.label = "All cases";

    for (const c of sorted) {
      const opt = document.createElement("option");
      opt.value = c.id;
      opt.textContent = c.name;
      if (c.pinned && (state.sort === "pinned-mru-alpha" || state.sort === "pinned-alpha")) {
        pinnedGroup.appendChild(opt);
      } else if (mruSet.has(c.id) && (state.sort === "pinned-mru-alpha" || state.sort === "mru")) {
        recentGroup.appendChild(opt);
      } else {
        otherGroup.appendChild(opt);
      }
    }

    if (pinnedGroup.childElementCount) sel.appendChild(pinnedGroup);
    if (recentGroup.childElementCount) sel.appendChild(recentGroup);
    if (otherGroup.childElementCount) sel.appendChild(otherGroup);
  }

  function onSortChange(e) {
    state.sort = e.target.value;
    renderOptions();
    if (state.selectedId) {
      document.getElementById("case-select").value = state.selectedId;
    }
  }

  async function onCaseChange(e) {
    const value = e.target.value;
    if (!value) return;
    state.selectedId = value;
    state.customProps.set(CONFIG.customPropertyKey, value);
    try {
      await saveCustomProps();
    } catch (err) {
      console.error(err);
      setStatus("Couldn't save selection: " + err.message, "error");
      return;
    }

    if (value === CONFIG.notCaseRelatedId) {
      await stripTagFromSubject();
      setStatus("Marked as not case related. Subject left alone.", "success");
    } else {
      const match = state.cases.find((c) => c.id === value);
      if (!match) return;
      await applyTagToSubject(match.name);
      await CT.recordUse(match.id);
      setStatus("Tagged as '" + match.name + "'.", "success");
    }
    await updatePreview();
  }

  function getSubject() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value || "");
        else reject(new Error(result.error ? result.error.message : "subject.getAsync failed"));
      });
    });
  }

  function setSubject(value) {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.subject.setAsync(value, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error ? result.error.message : "subject.setAsync failed"));
      });
    });
  }

  async function applyTagToSubject(caseName) {
    const current = await getSubject();
    const base = CT.stripExistingTag(current);
    const next = CT.buildTag(caseName) + base;
    await setSubject(next);
  }

  async function stripTagFromSubject() {
    const current = await getSubject();
    const base = CT.stripExistingTag(current);
    if (base !== current) await setSubject(base);
  }

  async function updatePreview() {
    const el = document.getElementById("preview-value");
    try {
      const subject = await getSubject();
      el.textContent = subject || "(empty)";
    } catch (err) {
      el.textContent = "(unable to read subject)";
    }
  }

  function setStatus(text, cls) {
    const el = document.getElementById("status");
    el.textContent = text || "";
    el.className = "status" + (cls ? " " + cls : "");
  }
})();
