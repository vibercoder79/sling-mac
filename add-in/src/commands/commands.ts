/* global Office */

Office.onReady(() => {});

(globalThis as unknown as Record<string, unknown>)["slingMail"] = slingMail;

export function slingMail(event: Office.AddinCommands.Event): void {
  event.completed();
}
