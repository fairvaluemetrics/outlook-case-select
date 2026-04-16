/* global Office */
/*
 * Headless event handlers for the Case Tagger add-in.
 *
 *   onNewMessageComposeHandler  - fires when a user starts composing a new
 *                                 message; we use it as a no-op hook today,
 *                                 but it gives us a place to seed defaults
 *                                 in the future.
 *
 *   onMessageSendHandler        - fires before the message is sent. If the
 *                                 user hasn't picked a case AND hasn't
 *                                 explicitly marked the message "not case
 *                                 related", we block the send with a Smart
 *                                 Alert that points them back at the pane.
 */

var CP_CASE_ID = "caseTagger.caseId";
var CP_NOT_CASE = "caseTagger.notCaseRelated";

Office.onReady(function () {
    // Nothing to do at module load. Functions are registered below.
});

function onNewMessageComposeHandler(event) {
    // Reserved for future seeding (e.g., infer case from recipient).
    event.completed();
}

function onMessageSendHandler(event) {
    var item = Office.context.mailbox.item;
    if (!item || typeof item.loadCustomPropertiesAsync !== "function") {
        // Be permissive if we can't introspect; never silently break send.
        event.completed({ allowEvent: true });
        return;
    }

    item.loadCustomPropertiesAsync(function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
        }

        var props = asyncResult.value;
        var caseId = props.get(CP_CASE_ID);
        var notCase = props.get(CP_NOT_CASE) === true;

        if (caseId || notCase) {
            event.completed({ allowEvent: true });
            return;
        }

        event.completed({
            allowEvent: false,
            errorMessage: "Pick a case from the Case Tagger pane, or mark this message as 'Not case related', then send again.",
            errorMessageMarkdown: "**No case selected.**\n\nOpen the **Case Tagger** pane (Message tab → *Tag case*) and either:\n\n- Pick the matching case, or\n- Click **Not case related**.\n\nThen press Send again.",
            cancelLabel: "Open Case Tagger",
            commandId: "caseTaggerOpenPaneButton"
        });
    });
}

// Office requires these to be globally addressable so the runtime can find
// them by the names declared in the manifest.
if (typeof globalThis !== "undefined") {
    globalThis.onNewMessageComposeHandler = onNewMessageComposeHandler;
    globalThis.onMessageSendHandler = onMessageSendHandler;
} else {
    // Fallback for older runtimes.
    /* eslint-disable no-undef */
    this.onNewMessageComposeHandler = onNewMessageComposeHandler;
    this.onMessageSendHandler = onMessageSendHandler;
    /* eslint-enable no-undef */
}
