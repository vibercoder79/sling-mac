import * as fs from "fs";
import * as path from "path";
import * as https from "https";
import express, { Request, Response } from "express";

const app = express();
app.use(express.json());

// CORS-Header fuer alle Requests
app.use((_req: Request, res: Response, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  next();
});

// OPTIONS pre-flight
app.options("*", (_req: Request, res: Response) => {
  res.sendStatus(200);
});

// GET /health
app.get("/health", (_req: Request, res: Response) => {
  res.json({ status: "ok", version: "0.1.0" });
});

// POST /sling
app.post("/sling", async (req: Request, res: Response) => {
  try {
    const { subject, from, to, body, date } = req.body as {
      subject: string;
      from: { displayName: string; emailAddress: string } | null;
      to: { displayName: string; emailAddress: string }[];
      body: string;
      date: string;
      conversationId: string;
    };

    // Datum im Format YYYY-MM-DD
    const dateStr = new Date(date).toISOString().slice(0, 10);

    // Betreff bereinigen fuer Dateinamen
    const cleanSubject = subject
      .replace(/[^A-Za-z0-9äöüÄÖÜß\- ]/g, "")
      .slice(0, 80)
      .trim();

    const fileName = `${dateStr} ${cleanSubject}.md`;
    const vaultInbox = "/Users/tobiasschmidt/Obsidian/SecondBrain/01 Inbox/";
    const filePath = path.join(vaultInbox, fileName);

    // Empfaenger formatieren
    const toFormatted = Array.isArray(to)
      ? to.map((r) => `${r.displayName} <${r.emailAddress}>`).join(", ")
      : "";

    const fromEmail = from?.emailAddress ?? "";
    const fromName = from?.displayName ?? "";

    // Markdown-Inhalt erstellen
    const markdown = `---
tags: [inbox, mail]
source: sling-mac
subject: ${subject}
from: ${fromEmail}
date: ${dateStr}
---

# ${subject}

**Von:** ${fromName} <${fromEmail}>
**An:** ${toFormatted}
**Datum:** ${dateStr}

---

<!-- Body-Konvertierung in Sprint 1 -->
${body}
`;

    await fs.promises.writeFile(filePath, markdown, "utf-8");

    res.json({ path: fileName });
  } catch (err) {
    console.error("Fehler beim Slingen:", err);
    res.status(500).json({ error: (err as Error).message });
  }
});

// TLS-Zertifikate laden
const certPath = path.resolve(__dirname, "../certs/cert.pem");
const keyPath = path.resolve(__dirname, "../certs/key.pem");

const httpsOptions = {
  cert: fs.readFileSync(certPath),
  key: fs.readFileSync(keyPath),
};

const PORT = 7331;

https.createServer(httpsOptions, app).listen(PORT, "127.0.0.1", () => {
  console.log(`Sling-Mac Helper laeuft auf https://127.0.0.1:${PORT}`);
});
