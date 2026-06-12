# Sling-Mac Troubleshooting

Kurze Sammlung bekannter Fehlerbilder mit Diagnose-Befehlen und Fix.

## Add-in-Fehler «Der Inhalt ist gesperrt oder nicht mit einem gültigen Sicherheitszertifikat signiert»

**Symptom:** Klick auf den Sling-Button in Outlook for Mac liefert die generische Fehlermeldung. Helper-Daemon läuft laut `launchctl list | grep sling`, Add-in lädt aber keine Daten.

**Wahrscheinliche Ursachen, in Reihenfolge:**

1. **IPv4 vs. IPv6-Binding** (häufigster Fall, gefixt in v0.1.3)
2. Cert-Trust für `localhost` nicht gesetzt
3. Outlook-Add-in-Cache hängt
4. Helper-Daemon läuft nicht

### Diagnose

```bash
# Läuft der Helper?
launchctl list | grep -i sling
lsof -nP -iTCP:7331 -sTCP:LISTEN
lsof -nP -iTCP:3000 -sTCP:LISTEN

# Erreichbar via IPv4 UND IPv6?
curl -ks -o /dev/null -w "127.0.0.1:7331 → %{http_code}\n" https://127.0.0.1:7331/health
curl -ks -o /dev/null -w "127.0.0.1:3000 → %{http_code}\n" https://127.0.0.1:3000/taskpane.html
curl -ks -o /dev/null -w "localhost:7331  → %{http_code}\n" https://localhost:7331/health
curl -ks -o /dev/null -w "localhost:3000  → %{http_code}\n" https://localhost:3000/taskpane.html
```

Erwartet: vier mal `200`. Wenn `127.0.0.1` failed und `localhost` 200 liefert → IPv4-Binding fehlt. Wenn beides failed → Helper nicht erreichbar (Daemon, Cert oder Cache).

### Fix 1 — IPv4-Binding fehlt (vor v0.1.3)

Behoben in v0.1.3 durch Dual-Stack-Binding im Helper. Wenn auf älterer Version: auf v0.1.3+ updaten oder in `helper/src/server.ts` die `listen`-Aufrufe von `"localhost"` auf `"127.0.0.1"` umstellen, neu bauen und Daemon neu starten:

```bash
cd ~/Developer/sling-mac/helper
npm run build
launchctl kickstart -k "gui/$(id -u)/ch.owlist.sling-mac-helper"
```

### Fix 2 — Cert abgelaufen oder nicht im Keychain (häufigster Auslöser)

`office-addin-dev-certs install` erzeugt im Default nur ein **30-Tage-Cert**. Nach Ablauf lehnt Outlooks WKWebView die TLS-Verbindung ab und liefert die «Inhalt gesperrt»-Meldung.

**Diagnose:**

```bash
openssl x509 -in ~/.office-addin-dev-certs/localhost.crt -noout -dates
# notAfter prüfen. In der Vergangenheit? → Cert ist abgelaufen.

security find-certificate -c "Developer CA for Microsoft Office Add-ins" /Library/Keychains/System.keychain
# Wenn nichts gefunden: Cert nicht als trusted im System-Keychain.
```

**Fix:**

```bash
cd ~/Developer/sling-mac/helper
npx office-addin-dev-certs uninstall --machine
npx office-addin-dev-certs install --machine --days 365
launchctl kickstart -k "gui/$(id -u)/ch.owlist.sling-mac-helper"
```

Wichtig: `--days 365` setzen, sonst läuft das Cert in 30 Tagen wieder ab. `--machine` braucht das sudo-Passwort und installiert die CA in den System-Keychain.

Anschliessend Outlook beenden und neu starten.

### Fix 3 — Outlook-Cache leeren

```bash
osascript -e 'quit app "Microsoft Outlook"'
sleep 2
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/com.microsoft.Outlook/WebKitCache/
open -a "Microsoft Outlook"
```

### Fix 4 — Daemon manuell neu starten

```bash
launchctl kickstart -k "gui/$(id -u)/ch.owlist.sling-mac-helper"
```

Wenn der Daemon dauerhaft nicht startet, plist prüfen:

```bash
cat ~/Library/LaunchAgents/ch.owlist.sling-mac-helper.plist
```

Insbesondere die Pfade auf den aktuellen Repo-Ordner (`~/Developer/sling-mac` seit v0.1.1).

## Sling-Button reagiert nicht / kein Slingen, keine Fehlermeldung

**Häufige Ursache:** `commands.js` wird aus dem Static-Server nicht ausgeliefert. Tritt typischerweise auf, wenn das Repo im iCloud-Ordner liegt (`errno -11 EDEADLK` auf synchrone Datei-Reads). Fix in v0.1.1: Repo nach `~/Developer/sling-mac` verschoben.

Wenn das Problem wieder auftritt:

```bash
# Prüfen ob commands.js HTTP 200 liefert
curl -k -I https://localhost:3000/commands.js
```

Wenn 500: Repo-Lage prüfen, ggf. aus iCloud-synchronisierten Ordnern raushalten.

## Picker öffnet sich nicht

**Bekanntes Verhalten:** Der Vault-Ordner-Picker ist als Code vollständig vorhanden (Helper-Endpoint `/folders` und Picker-UI in `taskpane.ts`), aber der Manifest-Button ist auf `ExecuteFunction → slingMail` gemappt und ruft `targetFolder: ""` hardcodiert. Sling landet deshalb immer im `defaultFolder`.

Reaktivierung erfordert Manifest-Anpassung auf `ShowTaskpane`. Historisch scheiterte das am Outlook-Mac-Caching der Manifest-Änderungen. Siehe PMO HUB.
