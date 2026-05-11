/* global Office */

Office.onReady(async () => {
  const item = Office.context.mailbox.item as Office.MessageRead;
  const accountEmail = Office.context.mailbox.userProfile.emailAddress;

  const subjectEl = document.getElementById("subject") as HTMLElement;
  subjectEl.textContent = item.subject ?? "(kein Betreff)";

  let allFolders: { label: string; path: string }[] = [];
  let selectedFolder: string | null = null;

  function renderFolders(folders: { label: string; path: string }[]): void {
    const list = document.getElementById("list") as HTMLElement;
    list.innerHTML = "";
    folders.forEach((f) => {
      const div = document.createElement("div");
      div.className = "folder" + (f.path === selectedFolder ? " selected" : "");
      div.textContent = f.label;
      div.addEventListener("click", () => {
        selectedFolder = f.path;
        document.querySelectorAll(".folder").forEach((el) => el.classList.remove("selected"));
        div.classList.add("selected");
        (document.getElementById("btn-sling") as HTMLButtonElement).disabled = false;
      });
      list.appendChild(div);
    });
    if (selectedFolder) {
      (document.getElementById("btn-sling") as HTMLButtonElement).disabled = false;
    }
  }

  function showStatus(msg: string, isError = false): void {
    const el = document.getElementById("status") as HTMLElement;
    el.textContent = msg;
    el.className = isError ? "error" : "ok";
    el.style.display = "block";
  }

  try {
    const resp = await fetch(
      `https://localhost:7331/folders?email=${encodeURIComponent(accountEmail)}`
    );
    const data = (await resp.json()) as { folders: { label: string; path: string }[]; default: string };
    allFolders = data.folders;
    selectedFolder = data.default ?? null;
    renderFolders(allFolders);
  } catch (err) {
    showStatus(`Ordner laden fehlgeschlagen: ${(err as Error).message}`, true);
  }

  (document.getElementById("search") as HTMLInputElement).addEventListener("input", (e) => {
    const q = (e.target as HTMLInputElement).value.toLowerCase();
    renderFolders(allFolders.filter((f) => f.label.toLowerCase().includes(q)));
  });

  (document.getElementById("btn-sling") as HTMLButtonElement).addEventListener("click", async () => {
    const btn = document.getElementById("btn-sling") as HTMLButtonElement;
    btn.disabled = true;
    showStatus("Wird geslingt…");

    try {
      const body = await new Promise<string>((resolve) => {
        item.body.getAsync(Office.CoercionType.Html, (result) => resolve(result.value ?? ""));
      });

      type AttachmentPayload = { name: string; contentBase64: string };
      const attachments: AttachmentPayload[] = [];
      const fileAtts = (item.attachments ?? []).filter(
        (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline
      );
      for (const att of fileAtts) {
        const result = await new Promise<Office.AsyncResult<Office.AttachmentContent>>((resolve) => {
          item.getAttachmentContentAsync(att.id, resolve);
        });
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          attachments.push({ name: att.name, contentBase64: result.value.content });
        }
      }

      const from = item.from;
      const response = await fetch("https://localhost:7331/sling", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          subject: item.subject ?? "(kein Betreff)",
          from: from ? { displayName: from.displayName, emailAddress: from.emailAddress } : null,
          to: (item.to ?? []).map((r) => ({ displayName: r.displayName, emailAddress: r.emailAddress })),
          body,
          date: new Date().toISOString(),
          conversationId: item.conversationId ?? "",
          accountEmail,
          targetFolder: selectedFolder ?? "",
          attachments,
        }),
      });

      if (!response.ok) throw new Error(`Helper: ${response.status} ${response.statusText}`);
      const result = (await response.json()) as { path: string; attachments: number };
      const attNote = result.attachments > 0 ? ` + ${result.attachments} Anhang/Anhänge` : "";
      showStatus(`Geslingt ✓\n${result.path}${attNote}`);
    } catch (err) {
      showStatus(`Fehler: ${(err as Error).message}`, true);
      btn.disabled = false;
    }
  });
});
