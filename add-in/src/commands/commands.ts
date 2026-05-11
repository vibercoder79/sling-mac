/* global Office */

Office.onReady(() => {});

(globalThis as unknown as Record<string, unknown>)["slingMail"] = slingMail;

export async function slingMail(event: Office.AddinCommands.Event): Promise<void> {
  const rawItem = Office.context.mailbox.item;
  if (!rawItem) {
    event.completed({ allowEvent: false });
    return;
  }
  const item = rawItem as Office.Item & Office.MessageRead;

  const subject = item.subject ?? "(kein Betreff)";
  const from = item.from;
  const toRecipients = item.to ?? [];

  const body = await new Promise<string>((resolve) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      resolve(result.value ?? "");
    });
  });

  type AttachmentPayload = { name: string; contentBase64: string };
  const attachments: AttachmentPayload[] = [];

  const fileAtts = (item.attachments ?? []).filter(
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

  const accountEmail = Office.context.mailbox.userProfile.emailAddress;

  try {
    const response = await fetch("https://localhost:7331/sling", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        subject,
        from: from ? { displayName: from.displayName, emailAddress: from.emailAddress } : null,
        to: toRecipients.map((r) => ({ displayName: r.displayName, emailAddress: r.emailAddress })),
        body,
        date: new Date().toISOString(),
        conversationId: item.conversationId ?? "",
        accountEmail,
        targetFolder: "",
        attachments,
      }),
    });
    if (!response.ok) throw new Error(`Helper: ${response.status} ${response.statusText}`);
    const result = (await response.json()) as { path: string; attachments: number };
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
