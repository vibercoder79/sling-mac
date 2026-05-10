/* global Office */

Office.onReady(() => {});

(globalThis as unknown as Record<string, unknown>)["slingMail"] = slingMail;

export async function slingMail(event: Office.AddinCommands.Event): Promise<void> {
  const rawItem = Office.context.mailbox.item;
  if (!rawItem) {
    event.completed({ allowEvent: false });
    return;
  }
  const item: Office.Item & Office.MessageRead = rawItem as Office.Item & Office.MessageRead;

  const subject = item.subject ?? "(kein Betreff)";
  const from = (item as Office.MessageRead).from;
  const toRecipients = (item as Office.MessageRead).to ?? [];

  const body = await new Promise<string>((resolve) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      resolve(result.value ?? "");
    });
  });

  // Anhänge sammeln (Mailbox 1.8 — graceful skip bei Fehler)
  type AttachmentPayload = { name: string; isInline: boolean; contentBase64: string };
  const attachments: AttachmentPayload[] = [];
  for (const att of (item as Office.MessageRead).attachments ?? []) {
    if (att.attachmentType !== Office.MailboxEnums.AttachmentType.File) continue;
    try {
      const content = await new Promise<string>((resolve, reject) => {
        item.getAttachmentContentAsync(att.id, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value.content);
          } else {
            reject(new Error(result.error.message));
          }
        });
      });
      attachments.push({ name: att.name, isInline: att.isInline, contentBase64: content });
    } catch {
      // nicht lesbar — überspringen
    }
  }

  const basePayload = {
    subject,
    from: from ? { displayName: from.displayName, emailAddress: from.emailAddress } : null,
    to: toRecipients.map((r) => ({ displayName: r.displayName, emailAddress: r.emailAddress })),
    body,
    date: new Date().toISOString(),
    conversationId: (item as Office.MessageRead).conversationId ?? "",
    accountEmail: Office.context.mailbox.userProfile.emailAddress,
    attachments,
  };

  async function doSling(targetFolder: string): Promise<void> {
    try {
      const response = await fetch("https://localhost:7331/sling", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...basePayload, targetFolder }),
      });
      if (!response.ok) throw new Error(`Helper: ${response.status} ${response.statusText}`);
      const result: { path: string } = await response.json();
      item.notificationMessages.addAsync("sling-ok", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Geslingt → ${result.path}`,
        icon: "Icon.16x16",
        persistent: false,
      });
    } catch (err) {
      item.notificationMessages.addAsync("sling-err", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: `Fehler: ${(err as Error).message}`,
      });
    }
    event.completed();
  }

  // Ordner-Picker Dialog öffnen
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/picker.html",
    { height: 55, width: 35, displayInIframe: false },
    (dialogResult) => {
      if (dialogResult.status !== Office.AsyncResultStatus.Succeeded) {
        void doSling("");
        return;
      }
      const dialog = dialogResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
        const msg = (args as { message: string }).message;
        dialog.close();
        if (msg === "__cancel__") {
          event.completed();
        } else {
          void doSling(msg);
        }
      });
      // Dialog vom User geschlossen (X-Button)
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        event.completed();
      });
    }
  );
}
