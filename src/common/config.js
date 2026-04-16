/*
 * Central configuration. Change values here to customize without touching
 * the rest of the code. Loaded by both the task pane and the command
 * function file, so keep it framework-free.
 */
(function (global) {
  "use strict";

  const CONFIG = {
    // Where the editable case list lives. Relative to the add-in root.
    casesUrl: "/config/cases.json",

    // Template used to prepend the tag. {case} is replaced with the case name.
    // Example: "{case}: " produces "Kandel v Claude: Re: meeting tomorrow"
    tagFormat: "{case}: ",

    // Regex used to recognize an existing tag at the start of the subject
    // so it can be replaced when the user picks a different case. Must
    // have a single capture group for the case name.
    tagPattern: "^([^:\\n]{1,120}):\\s+",

    // Default sort for the dropdown. One of:
    //   "pinned-mru-alpha" - pinned first, then recently used, then alphabetical
    //   "mru"              - most recently used first, rest alphabetical
    //   "alpha"            - plain alphabetical
    //   "pinned-alpha"     - pinned first, then alphabetical
    defaultSort: "pinned-mru-alpha",

    // How many MRU entries to remember (per user, roamed via Office settings).
    maxMruEntries: 10,

    // Label shown for the explicit "this email is not about a case" option.
    notCaseRelatedLabel: "— Not case related —",

    // Key used for the not-case-related sentinel. Must not collide with real ids.
    notCaseRelatedId: "__not_case_related__",

    // Key names for persisted state.
    roamingKeys: {
      mru: "caseTagger.mru",      // array of case ids, most recent first
      lastUsed: "caseTagger.lastUsed" // map of id -> ISO timestamp
    },

    // Key for the per-message custom property that records the user's choice.
    customPropertyKey: "caseTagger.selection",

    // Message shown when the user tries to send without tagging.
    sendBlockedMessage:
      "Please choose a case for this email, or mark it as 'Not case related' in the Case Tagger pane."
  };

  global.CaseTaggerConfig = CONFIG;
})(typeof window !== "undefined" ? window : globalThis);
