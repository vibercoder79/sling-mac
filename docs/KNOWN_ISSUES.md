# Sling-Mac — Known Issues

Schnellnachschlagewerk für wiederkehrende Probleme. Jeder Eintrag enthält Symptom, Ursache, Diagnose-Befehl und Fix. Für ausführliche Troubleshooting-Pfade siehe [`TROUBLESHOOTING.md`](TROUBLESHOOTING.md).

## 1. «Inhalt gesperrt oder nicht mit einem gültigen Sicherheitszertifikat signiert»

**Häufigster Fehler.** Tritt in zwei verschiedenen Konstellationen auf.

### 1a. Cert ist abgelaufen (häufigster Auslöser)

**Symptom:** Outlook-Add-in lieferte gestern noch, heute nicht mehr. Helper läuft (`launchctl list | grep sling`), curl liefert 200 OK. Outlook trotzdem mit Fehlermeldung.

**Ursache:** `office-addin-dev-certs install` erzeugt im Default nur ein **30-Tage-Cert**. Nach Ablauf lehnt Outlooks WKWebView die TLS-Verbindung ab.

**Diagnose:**

```bash
openssl x509 -in ~/.office-addin-dev-certs/localhost.crt -noout -dates
# notAfter in der Vergangenheit? Dann abgelaufen.
```

**Fix:**

```bash
cd ~/Developer/sling-mac/helper
npx office-addin-dev-certs uninstall --machine
npx office-addin-dev-certs install --machine --days 365
launchctl kickstart -k "gui/$(id -u)/ch.owlist.sling-mac-helper"
osascript -e 'quit app "Microsoft Outlook"' && sleep 2 && rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/com.microsoft.Outlook/WebKitCache/ && open -a "Microsoft Outlook"
```

`--days 365` verlängert die Gültigkeit auf ein Jahr. Sonst tritt das Problem in 30 Tagen wieder auf.

### 1b. Helper bindet nur an IPv6 (gefixt in v0.1.3)

**Symptom:** Add-in liefert «Inhalt gesperrt». Helper läuft, `curl localhost` liefert 200 OK, `curl 127.0.0.1` liefert «Connection refused».

**Ursache:** Node 17+ bindet bei `listen(port, "localhost")` praktisch nur an IPv6 (`::1`). Outlook for Mac WKWebView löst `localhost` aber zu IPv4 (`127.0.0.1`) auf.

**Diagnose:**

```bash
lsof -nP -iTCP:7331 -sTCP:LISTEN
lsof -nP -iTCP:3000 -sTCP:LISTEN
# Wenn nur "[::1]:PORT" auftaucht → IPv4-Binding fehlt.

curl -ks -o /dev/null -w "%{http_code}\n" https://127.0.0.1:7331/health
curl -ks -o /dev/null -w "%{http_code}\n" https://localhost:7331/health
# Wenn 127.0.0.1 failed und localhost 200 liefert → bestätigt.
```

**Fix:** Auf v0.1.3+ updaten. Dort bindet der Helper dual-stack (`127.0.0.1` + `::1`). Wenn nicht möglich: in `helper/src/server.ts` die `listen`-Aufrufe von `"localhost"` auf `"127.0.0.1"` umstellen, neu bauen, Daemon neu starten.

## 2. Sling-Button reagiert nicht, keine Fehlermeldung

**Symptom:** Klick auf den Button passiert nichts. Keine Mail im Vault, keine Konsole-Ausgabe.

**Ursache:** `commands.js` wird vom Static-Server nicht ausgeliefert (HTTP 500). Tritt typischerweise auf, wenn das Repo im iCloud-Ordner liegt (`errno -11 EDEADLK` auf synchrone Datei-Reads).

**Diagnose:**

```bash
curl -k -I https://localhost:3000/commands.js
# HTTP 500 oder Connection refused? Problem bestätigt.
```

**Fix:** Repo aus iCloud-Ordner herausnehmen. Seit v0.1.1 liegt das Repo unter `~/Developer/sling-mac`, nicht in `~/Documents/GitHub/`.

## 3. Picker öffnet sich nie, alle Mails landen im Default-Ordner

**Symptom:** Es gibt keinen Zielordner-Dialog. Alle geslingten Mails landen in `01 Inbox/`.

**Ursache:** Designed so seit v0.1.2. Picker-UI und Helper-Endpoint `/folders` sind komplett implementiert, aber das Manifest exponiert nur `ExecuteFunction → slingMail` mit `targetFolder: ""` hardcodiert. Es gibt keinen Button, der das `ShowTaskpane` aufruft.

**Fix:** Keiner. Bewusste Scope-Reduktion. Reaktivierung verlangt Manifest-Anpassung auf `ShowTaskpane` + neue Add-in-ID, weil Outlook for Mac Manifest-Änderungen hartnäckig cached. Siehe PMO HUB im Vault.

## Daemon läuft nicht

**Symptom:** `launchctl list | grep sling` liefert leeres Ergebnis oder Exit-Code ungleich 0.

**Diagnose und Fix:**

```bash
# plist prüfen — zeigt der ProgramArguments-Pfad auf das aktuelle Repo (~/Developer/sling-mac)?
cat ~/Library/LaunchAgents/ch.owlist.sling-mac-helper.plist | grep -A 2 ProgramArguments

# Daemon neu laden
launchctl bootout "gui/$(id -u)/ch.owlist.sling-mac-helper" 2>/dev/null
launchctl bootstrap "gui/$(id -u)" ~/Library/LaunchAgents/ch.owlist.sling-mac-helper.plist
launchctl kickstart -k "gui/$(id -u)/ch.owlist.sling-mac-helper"
```

Wenn der plist-Pfad auf das alte iCloud-Repo zeigt, plist anpassen oder Setup-Skript neu laufen lassen (`npm run setup` im Helper-Ordner).
