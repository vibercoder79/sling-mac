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

  // Anhänge via EWS GetAttachment holen.
  // getAttachmentContentAsync und getCallbackTokenAsync({isRest:true}) schlagen im
  // Function-Command-Context fehl. makeEwsRequestAsync funktioniert ohne Azure-AD-
  // Registrierung und läuft auch in Function Commands — braucht ReadWriteMailbox.
  type AttachmentPayload = { name: string; contentBase64: string };
  const attachments: AttachmentPayload[] = [];

  const fileAtts = ((item as Office.MessageRead).attachments ?? []).filter(
    (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline
  );

  const attErrors: string[] = [];
  for (const att of fileAtts) {
    const result = await new Promise<Office.AsyncResult<Office.AttachmentContent>>((resolve) => {
      item.getAttachmentContentAsync(att.id, resolve);
    });
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      attachments.push({ name: att.name, contentBase64: result.value.content });
    } else {
      attErrors.push(`${att.name}: ${result.error?.message ?? "unbekannt"}`);
    }
  }

  if (attErrors.length > 0) {
    item.notificationMessages.addAsync("sling-att-err", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: `Anhang-Fehler: ${attErrors.join("; ")}`,
    });
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
      const result: { path: string; attachments: number } = await response.json();
      const attNote = result.attachments > 0 ? ` | ${result.attachments} Anhang/Anhänge` : "";
      item.notificationMessages.addAsync("sling-ok", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Geslingt → ${result.path}${attNote}`,
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

  void doSling("");
}
