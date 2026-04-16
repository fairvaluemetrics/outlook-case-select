/*
 * Defaults. Anything in here can be overridden by the user via the
 * Settings tab; the override is persisted in Office roaming settings
 * under SETTINGS_KEY. Keep this file framework-free - it is loaded by
 * both the task pane and the command function file.
 */
(function (global) {
  "use strict";

  const CONFIG = {
    seedCasesUrl: "/config/cases.json",

    // Default tag template. {caseName} is replaced with the case name.
    tagFormat: "{caseName}: ",

    // Regex used to recognize an existing tag at the start of the
    // subject so it can be replaced when the user picks a different
    // case. Single capture group = the case name candidate.
    tagPattern: "^([^:\\n]{1,120}):\\s+",

    // "A" = inline / compact (one-row task pane styled to feel like
    //       an extension of the subject row).
    // "B" = dedicated row (full pane with preview + manage shortcut).
    // True inline DOM injection into Outlook's subject row is not
    // possible - this is the closest the platform allows.
    defaultPlacement: "B",

    // Default sort. One of: pinned-mru-alpha | mru | freq | alpha | pinned-alpha
    defaultSort: "pinned-mru-alpha",

    // If true, untagged sends are blocked. If false, the OnMessageSend
    // handler always allows the send.
    defaultRequired: true,

    notCaseRelatedLabel: "— Not case related —",
    notCaseRelatedId: "__not_case_related__",

    // Roaming-settings keys (per-mailbox; survives reinstall).
    keys: {
      cases: "caseTagger.cases",            // array of case objects
      settings: "caseTagger.settings",      // user settings overrides
      mru: "caseTagger.mru"                 // [id, ...] most-recent first
    },

    // Per-message custom property recording the user's choice so the
    // send-time validator can see it.
    customPropertyKey: "caseTagger.selection",

    maxMruEntries: 10,

    sendBlockedMessage:
      "Pick a case for this email in the Case Tagger pane, or mark it 'Not case related' before sending."
  };

  // Merge defaults with persisted overrides. Returns a flat object.
  CONFIG.resolveSettings = function (overrides) {
    const o = overrides || {};
    return {
      placement: o.placement || CONFIG.defaultPlacement,
      tagFormat: o.tagFormat || CONFIG.tagFormat,
      tagPattern: o.tagPattern || CONFIG.tagPattern,
      sort: o.sort || CONFIG.defaultSort,
      required: typeof o.required === "boolean" ? o.required : CONFIG.defaultRequired
    };
  };

  global.CaseTaggerConfig = CONFIG;
})(typeof window !== "undefined" ? window : globalThis);
