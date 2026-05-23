Office.onReady(() => {});

const INTERNAL_DOMAINS = ["stengg.com"];

const SENSITIVE_KEYWORDS = [
  "internal", "confidential", "do not forward", "not for external", "restricted"
];

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
          const subject  = subjectResult.value || "";
          const body     = bodyResult.value || "";
          const fullText = (subject + " " + body).toLowerCase();

          const isInternal     = subject.toUpperCase().includes("[INTERNAL]");
          const isExternal     = subject.toUpperCase().includes("[EXTERNAL]");
          const isUnclassified = !isInternal && !isExternal;

          const externalRecipients = recipients.filter(r => {
            const email = r.emailAddress.toLowerCase();
            return !INTERNAL_DOMAINS.some(d => email.endsWith("@" + d));
          });

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

          const reason = isInternal
            ? "Subject is tagged [INTERNAL]"
            : `Contains keyword: "${matchedKeyword}"`;

          const externalList = externalRecipients.map(r => r.emailAddress).join(",");

          const dialogUrl =
            "https://jameswong95.github.io/email-guard/warning.html" +
            "?reason=" + encodeURIComponent(reason) +
            "&recipients=" + encodeURIComponent(externalList);

          Office.context.ui.displayDialogAsync(
            dialogUrl,
            { height: 35, width: 45, displayInIframe: true },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                event.completed({ allowEvent: false });
                return;
              }
              const dialog = asyncResult.value;
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
                dialog.close();
                event.completed({ allowEvent: msg.message === "send" });
              });
              dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                event.completed({ allowEvent: false });
              });
            }
          );

        });
      });
    });
  });
}
