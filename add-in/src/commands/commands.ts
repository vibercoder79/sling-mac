/* global Office */

Office.onReady(() => {});

(globalThis as unknown as Record<string, unknown>)["slingMail"] = slingMail;

export async function slingMail(event: Office.AddinCommands.Event): Promise<void> {
  const item = Office.context.mailbox.item;
  if (!item) {
    event.completed({ allowEvent: false });
    return;
  }

  const subject = item.subject ?? "(kein Betreff)";
  const from = (item as Office.MessageRead).from;
  const toRecipients = (item as Office.MessageRead).to ?? [];

  const body = await new Promise<string>((resolve) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      resolve(result.value ?? "");
    });
  });

  const payload = {
    subject,
    from: from ? { displayName: from.displayName, emailAddress: from.emailAddress } : null,
    to: toRecipients.map((r) => ({ displayName: r.displayName, emailAddress: r.emailAddress })),
    body,
    date: new Date().toISOString(),
    conversationId: (item as Office.MessageRead).conversationId ?? "",
  };

  try {
    const response = await fetch("https://localhost:7331/sling", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error(`Helper: ${response.status} ${response.statusText}`);
    }

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
