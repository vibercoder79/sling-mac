import * as fs from "fs";
import * as path from "path";
import * as https from "https";
import express, { Request, Response, Router } from "express";
import { loadConfig, getAccountConfig, SlingMacConfig } from "./config";
import TurndownService from "turndown";

const turndown = new TurndownService({ headingStyle: "atx", codeBlockStyle: "fenced" });

let config: SlingMacConfig;
try {
  config = loadConfig();
} catch (err) {
  console.error((err as Error).message);
  process.exit(1);
}

const HIDDEN = new Set([".obsidian", ".git", ".trash", "node_modules"]);

async function readFolders(vaultPath: string): Promise<{ label: string; path: string }[]> {
  const folders: { label: string; path: string }[] = [];
  const level1 = await fs.promises.readdir(vaultPath, { withFileTypes: true });
  for (const e1 of level1) {
    if (!e1.isDirectory() || HIDDEN.has(e1.name) || e1.name.startsWith(".")) continue;
    folders.push({ label: e1.name, path: e1.name });
    const level2 = await fs.promises.readdir(path.join(vaultPath, e1.name), { withFileTypes: true });
    for (const e2 of level2) {
      if (!e2.isDirectory() || HIDDEN.has(e2.name) || e2.name.startsWith(".")) continue;
      folders.push({ label: `${e1.name}/${e2.name}`, path: `${e1.name}/${e2.name}` });
    }
  }
  return folders;
}

let foldersCache: { folders: { label: string; path: string }[]; at: number } | null = null;
let pendingFolder: string | null = null;

// Picker-Endpoints werden auf Port 7331 UND Port 3000 gemountet.
// Port 3000 = same-origin fuer picker.html im Dialog-WebView (Sandbox-Fix).
const pickerRouter = Router();
pickerRouter.use(express.json({ limit: "10mb" }));

pickerRouter.get("/folders", async (req: Request, res: Response) => {
  try {
    const email = req.query["email"] as string | undefined;
    const accountConfig = getAccountConfig(config, email ?? "");
    const now = Date.now();
    if (!foldersCache || now - foldersCache.at > 30_000) {
      foldersCache = { folders: await readFolders(accountConfig.vaultPath), at: now };
    }
    res.json({ folders: foldersCache.folders, default: accountConfig.defaultFolder });
  } catch (err) {
    res.status(500).json({ error: (err as Error).message });
  }
});

pickerRouter.post("/set-pending-folder", (req: Request, res: Response) => {
  pendingFolder = (req.body as { folder: string }).folder ?? null;
  res.json({ ok: true });
});

// API-App (Port 7331)
const app = express();
app.use(express.json({ limit: "10mb" }));

// CORS-Header fuer cross-origin Requests aus dem Add-in-WebView
app.use((_req: Request, res: Response, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  next();
});

app.options("*", (_req: Request, res: Response) => {
  res.sendStatus(200);
});

app.use(pickerRouter);

// GET /health
app.get("/health", (_req: Request, res: Response) => {
  res.json({ status: "ok", version: "0.1.0" });
});

// GET /get-pending-folder
app.get("/get-pending-folder", (_req: Request, res: Response) => {
  const folder = pendingFolder;
  pendingFolder = null;
  res.json({ folder });
});

// POST /sling
app.post("/sling", async (req: Request, res: Response) => {
  try {
    const { subject, from, to, body, date, accountEmail, targetFolder, attachments } = req.body as {
      subject: string;
      from: { displayName: string; emailAddress: string } | null;
      to: { displayName: string; emailAddress: string }[];
      body: string;
      date: string;
      conversationId: string;
      accountEmail: string;
      targetFolder?: string;
      attachments?: { name: string; contentBase64: string }[];
    };

    console.log(`[sling] subject="${subject.slice(0, 50)}" attachments=${attachments?.length ?? 0}`);

    const accountConfig = getAccountConfig(config, accountEmail);
    const slingFolder = targetFolder || accountConfig.defaultFolder;

    // Datum im Format YYYY-MM-DD
    const dateStr = new Date(date).toISOString().slice(0, 10);

    // Betreff bereinigen fuer Dateinamen
    const cleanSubject = subject
      .replace(/[^A-Za-z0-9äöüÄÖÜß\- ]/g, "")
      .slice(0, 80)
      .trim();

    const folderName = `${dateStr} ${cleanSubject}`;
    const fileName = `${folderName}.md`;
    const targetPath = path.join(accountConfig.vaultPath, slingFolder, folderName);
    await fs.promises.mkdir(targetPath, { recursive: true });
    const filePath = path.join(targetPath, fileName);

    // Empfaenger formatieren
    const toFormatted = Array.isArray(to)
      ? to.map((r) => `${r.displayName} <${r.emailAddress}>`).join(", ")
      : "";

    const fromEmail = from?.emailAddress ?? "";
    const fromName = from?.displayName ?? "";

    const bodyMarkdown = turndown.turndown(body || "");

    // Markdown-Inhalt erstellen
    const markdown = `---
tags: [inbox, mail]
source: sling-mac
subject: "${subject}"
from: ${fromEmail}
date: ${dateStr}
---

# ${subject}

**Von:** ${fromName} <${fromEmail}>
**An:** ${toFormatted}
**Datum:** ${dateStr}

---

${bodyMarkdown}
`;

    // Anhänge als Base64 vom Add-in empfangen (via EWS GetAttachment) und speichern
    const savedAttachments: string[] = [];
    if (attachments && attachments.length > 0) {
      for (const att of attachments) {
        try {
          await fs.promises.writeFile(
            path.join(targetPath, att.name),
            Buffer.from(att.contentBase64, "base64")
          );
          savedAttachments.push(att.name);
        } catch (err) {
          console.error(`Anhang-Fehler (${att.name}):`, (err as Error).message);
        }
      }
    }

    // Wikilinks an Markdown anhängen
    let finalMarkdown = markdown;
    if (savedAttachments.length > 0) {
      const wikilinks = savedAttachments
        .map((name) => {
          const isImage = /\.(png|jpg|jpeg|gif|webp|svg)$/i.test(name);
          return isImage ? `![[${name}]]` : `[[${name}]]`;
        })
        .join("\n");
      finalMarkdown += `\n## Anhänge\n\n${wikilinks}\n`;
    }

    await fs.promises.writeFile(filePath, finalMarkdown, "utf-8");

    res.json({ path: fileName, attachments: savedAttachments.length });
  } catch (err) {
    console.error("Fehler beim Slingen:", err);
    res.status(500).json({ error: (err as Error).message });
  }
});

// TLS-Zertifikate: dasselbe trusted Cert wie der Add-in Dev Server
const certPath = path.join(process.env.HOME!, ".office-addin-dev-certs/localhost.crt");
const keyPath = path.join(process.env.HOME!, ".office-addin-dev-certs/localhost.key");

const httpsOptions = {
  cert: fs.readFileSync(certPath),
  key: fs.readFileSync(keyPath),
};

// API-Server (Port 7331)
const PORT = 7331;
https.createServer(httpsOptions, app).listen(PORT, "localhost", () => {
  console.log(`Sling-Mac Helper laeuft auf https://localhost:${PORT}`);
});

// Static File Server + Picker-API (Port 3000)
// pickerRouter hier gemountet → /folders und /set-pending-folder sind same-origin
// fuer picker.html im Dialog-WebView → kein Cross-Origin-Block mehr.
const staticApp = express();
staticApp.use((_req, res, next) => {
  res.setHeader("Cache-Control", "no-store");
  next();
});
staticApp.use(pickerRouter);
staticApp.use(express.static(path.resolve(__dirname, "../../add-in/dist")));
const STATIC_PORT = 3000;
https.createServer(httpsOptions, staticApp).listen(STATIC_PORT, "localhost", () => {
  console.log(`Sling-Mac Static Server auf https://localhost:${STATIC_PORT}`);
});
