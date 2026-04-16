/*
 * Event-based command handlers. Wired up in manifest.xml:
 *   - onNewMessageComposeHandler: auto-opens the pinnable task pane
 *     whenever the user starts a new message, so the case dropdown is
 *     visible from the first keystroke.
 *   - onMessageSendHandler: runs before the message leaves the client,
 *     blocks the send if the subject has not been tagged and the user
 *     has not explicitly marked the email as not case related.
 */
(function () {
  "use strict";

  const CONFIG = window.CaseTaggerConfig;
  const CT = window.CaseTagger;

  Office.onReady(() => {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  });

  function onNewMessageComposeHandler(event) {
    // Programmatically open the task pane so the case dropdown appears
    // as soon as the compose window opens. SupportsPinning in the
    // manifest keeps it open across subsequent messages once the user
    // clicks the pin.
    try {
      if (Office.addin && typeof Office.addin.showAsTaskpane === "function") {
        Office.addin.showAsTaskpane().finally(() => event.completed());
        return;
      }
    } catch (e) {
      // fall through
    }
    event.completed();
  }

  async function onMessageSendHandler(event) {
    try {
      const [subject, selection, cases] = await Promise.all([
        getSubject(),
        getStoredSelection(),
        CT.loadCases().catch(() => [])
      ]);

      if (selection === CONFIG.notCaseRelatedId) {
        event.completed({ allowEvent: true });
        return;
      }

      const knownNames = cases.map((c) => c.name);
      const tagged = CT.subjectHasTag(subject, knownNames);

      if (tagged && selection) {
        event.completed({ allowEvent: true });
        return;
      }

      // Subject already manually tagged with a known case name - accept.
      if (tagged) {
        event.completed({ allowEvent: true });
        return;
      }

      event.completed({
        allowEvent: false,
        cancelLabel: "Choose a case",
        errorMessage: CONFIG.sendBlockedMessage,
        errorMessageMarkdown: CONFIG.sendBlockedMessage,
        commandId: "msgComposeOpenPaneButton",
        contextData: "{}"
      });
    } catch (err) {
      // Never silently accept on error - that would defeat the blocker.
      event.completed({
        allowEvent: false,
        errorMessage:
          "Case Tagger couldn't verify this message (" + (err && err.message ? err.message : err) +
          "). Open the Case Tagger pane and pick a case, then try again."
      });
    }
  }

  function getSubject() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value || "");
        else reject(new Error(result.error ? result.error.message : "subject.getAsync failed"));
      });
    });
  }

  function getStoredSelection() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error ? result.error.message : "loadCustomPropertiesAsync failed"));
          return;
        }
        resolve(result.value.get(CONFIG.customPropertyKey) || null);
      });
    });
  }
})();
