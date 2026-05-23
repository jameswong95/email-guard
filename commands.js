/* ============================================================
   Email Guard — commands.js
   Hosted on GitHub Pages. Referenced by manifest.xml.
   ============================================================ */

Office.onReady(() => {});

// ── CONFIGURATION ─────────────────────────────────────────────
const INTERNAL_DOMAINS = [
  "stengg.com"       // ← EDIT: add your org domain(s)
];

// Fallback keyword check (catches unclassified emails with sensitive words)
const SENSITIVE_KEYWORDS = [
  "internal", "confidential", "do not forward", "not for external", "restricted"
];
// ─────────────────────────────────────────────────────────────


// ── RIBBON BUTTONS ───────────────────────────────────────────

function setInternal(event) {
  const item = Office.context.mailbox.item;
  item.subject.getAsync((result) => {
    let subject = (result.value || "").replace(/^\[(INTERNAL|EXTERNAL)\]\s*/i, "").trim();
    item.subject.setAsync("[INTERNAL] " + subject, () => event.completed());
  });
}

function setExternal(event) {
  const item = Office.context.mailbox.item;
  item.subject.getAsync((result) => {
    let subject = (result.value || "").replace(/^\[(INTERNAL|EXTERNAL)\]\s*/i, "").trim();
    item.subject.setAsync("[EXTERNAL] " + subject, () => event.completed());
  });
}


// ── SEND INTERCEPT ───────────────────────────────────────────

function onItemSend(event) {
  const item = Office.context.mailbox.item;

  item.to.getAsync((toResult) => {
    item.cc.getAsync((ccResult) => {
      item.subject.getAsync((subjectResult) => {
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {

          const recipients = [
            ...(toResult.value || []),
            ...(ccResult.value || [])
          ];
          const subject   = subjectResult.value || "";
          const body      = bodyResult.value || "";
          const fullText  = (subject + " " + body).toLowerCase();

          // Classify by subject tag
          const isInternal     = subject.toUpperCase().includes("[INTERNAL]");
          const isExternal     = subject.toUpperCase().includes("[EXTERNAL]");
          const isUnclassified = !isInternal && !isExternal;

          // Find external recipients
          const externalRecipients = recipients.filter(r => {
            const email = r.emailAddress.toLowerCase();
            return !INTERNAL_DOMAINS.some(d => email.endsWith("@" + d));
          });

          // Keyword match for unclassified emails
          const matchedKeyword = SENSITIVE_KEYWORDS.find(kw =>
            fullText.includes(kw.toLowerCase())
          );

          const shouldWarn =
            externalRecipients.length > 0 &&
            (isInternal || (isUnclassified && matchedKeyword));

          if (!shouldWarn) {
            event.completed({ allowEvent: true });
            return;
          }

          // Build warning reason
          const reason = isInternal
            ? "Subject is tagged [INTERNAL]"
            : `Contains keyword: "${matchedKeyword}"`;

          const externalList = externalRecipients.map(r => r.emailAddress).join(",");

          const dialogUrl =
            "https://jameswong95.github.io/email-guard/warning.html" +  // ← EDIT
            "?reason=" + encodeURIComponent(reason) +
            "&recipients=" + encodeURIComponent(externalList);

          Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 35, width: 45, displayInIframe: true },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // If dialog fails to open, block by default (safe failure)
                event.completed({ allowEvent: false });
                return;
              }
              const dialog = asyncResult.value;
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
                dialog.close();
                event.completed({ allowEvent: msg.message === "send" });
              });
              dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                // User closed dialog without choosing — block
                event.completed({ allowEvent: false });
              });
            }
          );

        });
      });
    });
  });
}
