# Outlook-Kategorien mit Xurrent abgleichen

[English](README.md) · Deutsch · inoffiziell, nicht mit Microsoft oder Xurrent verbunden

Vergleicht die Mails im Postfach, die noch **„Nicht erledigt"** oder **„In Bearbeitung"**
tragen, mit dem Status der zugehörigen Xurrent-Requests, setzt die Mails
abgeschlossener Requests auf **„Erledigt"** und die Mails von Requests, die auf jemand
anderen warten, von „Nicht erledigt" auf **„In Bearbeitung"**. Läuft lokal, auf Abruf.
Kein Server, kein Dienst, keine Cloud-Komponente.

## Garantien

| | |
|---|---|
| **Dry-run ist Default** | Ohne `--apply` wird in Outlook nichts geschrieben. Lokaler Cache, Logs und Tagesprotokoll werden trotzdem geführt. |
| **Xurrent ist read-only** | `XurrentClient._call` kann strukturell nur `GET`. Kein Create, Update oder Kommentar. |
| **Outlook: nur `categories`** | Einzige Schreiboperation ist `PATCH /me/messages/{id}` mit ausschließlich dem Feld `categories`. Kein Verschieben, Löschen, Markieren, Antworten, Weiterleiten. |
| **Fremde Kategorien bleiben** | Andere Kategorien derselben Mail bleiben in ihrer Reihenfolge erhalten — nur die konfigurierten „offen"-Kategorien werden entfernt. |
| **Keine Request-ID → keine Änderung** | Mails ohne erkannte ID werden nie angefasst, sondern zur manuellen Prüfung aufgelistet. |
| **Notiz nach Abschluss bleibt offen** | Eine Mail, die deutlich *nach* dem Abschluss des Requests eintraf, wird nie auf erledigt gesetzt. |
| **Idempotent** | Ein zweiter Lauf direkt danach ändert null Mails. |
| **Eigenes Postfach** | Delegierter Zugriff als angemeldeter Benutzer. Keine Application-Permissions, kein Zugriff auf fremde Postfächer. |

## Status

Im Einsatz seit 2026-10-05.

- **Xurrent:** PAT, Scope `Request:Read`, Account-Header und Region verifiziert.
- **Microsoft Graph:** freigegeben, Anmeldung und Kategorie-Validierung laufen.
- **Tests:** `python -m pytest -q`, ohne Netzwerkzugriff.
- **Erster Produktivlauf:** 283 offene Mails geprüft, 237 auf „Erledigt" gesetzt, 0 Fehler.
  266 von 283 Mails ließen sich einer Request zuordnen; die 17 übrigen waren tatsächlich
  ticketlos (Newsletter, Phishing-Warnungen, Hausmitteilungen). Erkennungsquellen über 371
  gecachte Mails: Body-Link 184, Betreff 169.

## Welche Mails es überhaupt gibt

Xurrent verschickt eine Mail, wenn ein Request **zugeteilt** wird, und wenn ein Kunde oder
Kollege eine **Notiz** schreibt. Nimmt man einen Request selbst, kommt keine Mail. Daraus
folgen drei Dinge für die Logik:

1. **Ein Request hat oft mehrere Mails** — Zuteilung plus jede Notiz. Ist der Request
   erledigt, werden alle zugehörigen Mails umkategorisiert, nicht nur eine.
2. **Selbst genommene Requests haben keine Mail.** Dafür gibt es in Outlook nichts zu
   kategorisieren; die Outlook-Liste ist nicht die vollständige Ticketliste und soll es
   nicht sein.
3. **Eine Notiz kann auf einem bereits abgeschlossenen Request landen.** Dann ist der
   Status `completed`, obwohl die Mail neu ist und womöglich noch eine Antwort braucht.
   Solche Mails würden sonst sofort als erledigt markiert und aus der To-do-Liste
   verschwinden.

Gegen Punkt 3 gilt: eine Mail wird nur dann als erledigt behandelt, wenn `completed_at`
des Requests nicht mehr als `AFTER_CLOSE_GRACE_SECONDS` vor dem `receivedDateTime` der
Mail liegt. Alles andere erscheint im Report als `AFTER-CLOSE` und bleibt unangetastet.
Fehlt einer der beiden Zeitstempel, entscheidet der Status allein.

Die Kulanzfrist ist nötig, weil Xurrent die Abschluss-Benachrichtigung **nach** dem
Abschluss verschickt — ohne sie wäre genau die Mail, die „fertig" meldet, die eine, die
offen bleibt. Der Default von zwei Stunden stammt aus einer Messung im Postfach: von 30
Mails, die nach dem Abschluss ihres Requests eintrafen, kamen 18 innerhalb von 96 Minuten
und die nächste erst 28,5 Stunden später.

```
     12 s   Request #80000011 Kann sich nicht anmelden     <- Benachrichtigung
     75 s   ✔️Request #80000012 - Router Standort B completed!
     23 min  Note added to request: Firewall Standort B               <- Abschlussnotiz
     96 min  👉You were mentioned in Request #80000013
   ---------------- Grenze bei 2 h, 17-fache Lücke ----------------
   28.5 h    Note added to request: Teams Anmeldung                   <- echte Nachfrage
   88.3 Tage AW: NEUER USER - Name: Erika Beispiel
```

`python sync.py --explain "..."` zeigt den Abstand pro Mail an, etwa
`mail 22s after the closure`.

## Wartende Requests → „In Bearbeitung"

Steht ein Request auf `waiting_for` oder `waiting_for_customer`, liegt der nächste
Schritt bei jemand anderem — Kunde, Lieferant, Kollege. Mit `WAITING_CATEGORY=In
Bearbeitung` wandern seine Mails von „Nicht erledigt" nach „In Bearbeitung". So bleibt
in „Nicht erledigt" nur, was gerade auf einen selbst wartet.

- **Alle Requests der Mail müssen warten oder erledigt sein.** Braucht auch nur einer
  noch eine Aktion (`assigned`, `in_progress`, …), bleibt die Mail, wo sie ist. Ebenso
  bei einer nicht lesbaren Request und bei einer Notiz auf einer bereits geschlossenen
  Request (`AFTER-CLOSE`).
- **„In Bearbeitung" bleibt eine offene Kategorie.** Die Mail ist damit nicht aus dem
  Blick: sobald die Request abgeschlossen ist, setzt der nächste Lauf sie auf
  „Erledigt". Deshalb muss `WAITING_CATEGORY` auch in `OPEN_CATEGORY` stehen, sonst
  bricht das Skript mit einem Konfigurationsfehler ab.
- **Nur in eine Richtung.** Springt die Request zurück auf `in_progress`, wird „In
  Bearbeitung" nicht wieder zu „Nicht erledigt". Die Kategorie wird auch von Hand
  vergeben, für Tickets, an denen man gerade arbeitet — ein automatischer Rückweg würde
  diese Markierungen überschreiben.
- **Erwähnungen bleiben in „Nicht erledigt".** `👉You were mentioned in Request #…`
  heißt, jemand hat einem eine Frage gestellt — auch wenn die Request selbst auf den
  Kunden wartet. Solche Mails erscheinen im Report als `MENTION` und werden nicht
  verschoben. Erkannt am Betreff über `MENTION_REGEX`. Auf das Erledigen hat das keinen
  Einfluss: ist die Request abgeschlossen, wird auch die Erwähnung „Erledigt".
- **Mails, die schon „In Bearbeitung" tragen, werden nicht angefasst**, auch nicht
  umsortiert. Fremde Kategorien bleiben wie beim Erledigen in ihrer Reihenfolge erhalten.

Stand 2026-10-08: von 98 offenen Mails gehören 21 zu wartenden Requests
(19 `waiting_for_customer`, 2 `waiting_for`). 18 davon liegen in „In Bearbeitung", die 3
übrigen sind Erwähnungen und bleiben in „Nicht erledigt".

## Graph-Freigabe

Erledigt — die delegierten Berechtigungen sind im Tenant zugestimmt, es läuft über die
Microsoft-App „Microsoft Graph Command Line Tools" (Default in `GRAPH_CLIENT_ID`). Der
folgende Abschnitt bleibt als Referenz, falls die Zustimmung einmal zurückgezogen wird
oder eine eigene Registrierung bevorzugt wird.

| Einstellung | Wert |
|---|---|
| Name | z. B. `Outlook-Xurrent-Sync` |
| Supported account types | nur dieses Verzeichnis |
| Authentication | *Mobile and desktop applications*, Redirect URI `http://localhost`, „Allow public client flows" = Yes |
| API permissions | Microsoft Graph → **Delegated** → `Mail.ReadWrite` (optional `MailboxSettings.Read`) |
| Application permissions | **keine** |
| Optional | „Assignment required" aktivieren und nur den Benutzer zuweisen |

`Mail.ReadWrite` ist die geringste Berechtigung, die Mails lesen *und* das Feld
`categories` schreiben kann — `Mail.Read` reicht zum Schreiben nicht.
`MailboxSettings.Read` wird ausschließlich benutzt, um beim Start zu prüfen, ob die zwei
Kategorienamen im Postfach existieren; ohne die Berechtigung läuft es mit
`VALIDATE_CATEGORIES=false` genauso.

Aus der Registrierung werden **Application (client) ID** und **Directory (tenant) ID**
gebraucht, sie kommen als `GRAPH_CLIENT_ID` und `GRAPH_TENANT_ID` in die `.env`.

## Setup

Voraussetzung: Python 3.10 oder neuer.

```bash
python -m pip install -r requirements.txt
```

Xurrent-Token erzeugen: Avatar → **My Profile** → **Personal Access Tokens** →
*Generate new token*, Scope-Tabelle:

| Spalte | Wert |
|---|---|
| Effect | `Allow` |
| Actions | Record type **Request**, Aktion **Read** — sonst nichts |
| Conditions | **Account** = `example-it` |

Ein PAT hat nie mehr Rechte als die Person, die ihn erzeugt hat.

Dann `.env` im Projektordner anlegen — sechs Zeilen genügen, alles andere hat brauchbare
Defaults:

```ini
XURRENT_TOKEN=<der Token von oben>
XURRENT_ACCOUNT=example-it
XURRENT_INSTANCE_HOST=example-it.xurrent.com
OPEN_CATEGORY=Nicht erledigt,In Bearbeitung
DONE_CATEGORY=Erledigt
WAITING_CATEGORY=In Bearbeitung
```

### Alle Einstellungen

Echte Umgebungsvariablen haben Vorrang vor der `.env`. Operative Werte (Timeouts,
Retries, Seiten- und Batchgrößen) sind bewusst keine Einstellung, sondern Konstanten in
[config.py](outlook_xurrent/config.py).

| Variable | Default | Bedeutung |
|---|---|---|
| `XURRENT_TOKEN` | — | **Pflicht.** Personal Access Token, Scope `Request:Read`. |
| `XURRENT_ACCOUNT` | — | **Pflicht.** Wert des `X-Xurrent-Account`-Headers, meist die Subdomain der Instanz. |
| `DONE_CATEGORY` | — | **Pflicht.** Exakter Name der „erledigt"-Kategorie. |
| `OPEN_CATEGORY` | leer | Exakte Namen der „offen"-Kategorien, mehrere durch Komma getrennt. Jede davon wird beim Erledigen entfernt. Leer = jede Mail ohne `DONE_CATEGORY` gilt als offen. |
| `WAITING_CATEGORY` | leer | Kategorie für Mails, deren Requests auf jemand anderen warten, z. B. `In Bearbeitung`. Muss auch in `OPEN_CATEGORY` stehen. Leer = aus. |
| `WAITING_STATUSES` | `waiting_for,waiting_for_customer` | Status, die als „wartet" gelten. Nur wirksam mit `WAITING_CATEGORY`, darf sich nicht mit `DONE_STATUSES` überschneiden. |
| `XURRENT_INSTANCE_HOST` | leer | Positivliste der Hosts für die Link-Erkennung, mehrere durch Komma getrennt (Agenten-Instanz und Kundenportal). Leer = jede `*.xurrent.com` wird akzeptiert. |
| `XURRENT_API_ROOT` | `https://api.xurrent.com/v1` | Regions-Root. Instanz ohne Regionssegment → global; sonst `api.au` / `api.uk` / `api.ch` / `api.us`. |
| `LOOKBACK_DAYS` | `0` | Obergrenze in Tagen. `0` = **keine Zeitgrenze**, der gesamte Altbestand wird geprüft. Pro Lauf via `--days` überschreibbar. |
| `DONE_STATUSES` | `completed` | Status, die als erledigt gelten. Gültig: `declined`, `on_backlog`, `assigned`, `accepted`, `in_progress`, `waiting_for`, `waiting_for_customer`, `reservation_pending`, `workflow_pending`, `project_pending`, `completed`. |
| `MULTI_REQUEST_POLICY` | `all` | Mail mit mehreren Requests: `all` = alle müssen erledigt sein, `any` = einer genügt. |
| `AFTER_CLOSE_GRACE_SECONDS` | `7200` | Wie lange nach `completed_at` eine Mail noch zum Abschluss zählt. `0` = strikt, jede Mail nach dem Abschluss bleibt offen. |
| `VALIDATE_CATEGORIES` | `true` | Prüft beim Start, ob alle konfigurierten Kategorienamen im Postfach existieren. Braucht `MailboxSettings.Read`. |
| `GRAPH_CLIENT_ID` | Microsoft CLI-App | Eigene App-Registrierung hier eintragen, wenn der Tenant User-Consent sperrt. |
| `GRAPH_TENANT_ID` | `organizations` | Eigene Tenant-ID ist etwas schneller. |
| `GRAPH_SCOPES` | `Mail.ReadWrite,MailboxSettings.Read` | Delegierte Scopes. `Mail.ReadWrite` ist Pflicht, der zweite nur für die Kategorie-Validierung. |
| `SUBJECT_REGEX` | `#\s*(\d{4,})` | Schritt 1 und 3 der ID-Erkennung. |
| `HEADER_REGEX` | `requests?[._/-](\d{4,})` | Schritt 4 der ID-Erkennung. |
| `MENTION_REGEX` | `\byou were mentioned\b` | Betreff einer Erwähnungs-Mail (Groß-/Kleinschreibung egal). Solche Mails wandern nie nach `WAITING_CATEGORY`. Leer = Erwähnungen werden wie jede andere Mail behandelt. |
| `GRAPH_TOKEN_CACHE` | `.state/msal_cache.json` | Refresh-Token. Wie ein Passwort behandeln. |
| `STATE_FILE` | `.state/state.json` | ID-Cache, Audit-Trail, letzter Lauf. |
| `LOG_DIR` | `logs` | Zielverzeichnis für `sync.log`. |
| `DAILY_LOG_DIR` | `logs/daily` | Ordner für das Tagesprotokoll, eine Datei je Tag. |

## Benutzung

```bash
python sync.py --check
```

Prüft Xurrent-Zugriff, Login und Kategorienamen. Schreibt nichts.

```bash
python sync.py
```

Dry-run über alle noch offenen Mails. **Kein Zeitfenster:** eine unbearbeitete Mail bleibt
beliebig lange liegen, ein kurzes Fenster würde den Altbestand stillschweigend verstecken.
Möglich wird das durch einen serverseitigen Filter auf die offene Kategorie
(`categories/any(c:c eq '…')`) — in der Graph-Doku nicht zugesichert, in diesem Tenant
aber verifiziert und dreimal schneller als das ungefilterte Listing des Ordners. Wird er
abgelehnt, listet das Skript den Ordner und filtert client-seitig; die Entscheidung fällt
ohnehin lokal, weil der Serverfilter „und nicht Erledigt" nicht ausdrücken kann.

```
REQUEST     STATUS                ACTION       SUBJECT
----------  --------------------  -----------  ----------------------------
80000001    completed             WOULD-APPLY  WG: Zugangsdaten Emailadresse
80000003    waiting_for_customer  WOULD-APPLY  Neuer Mitarbeiter - Onboarding
80000002    in_progress           STILL-OPEN   Drucker im 2. Stock offline
70000001    completed             AFTER-CLOSE  Notiz nach Abschluss
-           -                     NO-ID        Kannst du bitte den Druck...

Mails with no request id (1) - review these by hand:
  2026-09-18T09:20:00  Kannst du bitte den Drucker im 2. Stock anschauen?
```

Erst vorsichtig anwenden, dann regulär:

```bash
python sync.py --apply --limit 5
```

| Flag | Wirkung |
|---|---|
| `--apply` | schreibt tatsächlich (ohne das Flag: Dry-run) |
| `--days N` | nur Mails der letzten N Tage (Default: keine Zeitgrenze) |
| `--limit N` | mit `--apply`: maximal N Mails ändern |
| `--check` | nur Konfiguration und Zugriffe prüfen |
| `--explain "text"` | für die Mails, deren Betreff `text` enthält: warum wurden sie angefasst oder nicht (rein lesend) |
| `--no-login` | öffnet nie ein Anmeldefenster, sondern bricht mit Exit-Code 2 ab (für den Hintergrund-Task) |
| `--no-cache` | ID-Cache ignorieren, alle Kandidaten neu parsen |
| `--verbose` | Debug-Logging auf stdout |

Exit-Codes: `0` Erfolg · `1` Teilfehler · `2` Konfigurations- oder Auth-Problem.

## Wie die Request-ID gefunden wird

Vier Schritte, erster Treffer gewinnt ([extract.py](outlook_xurrent/extract.py)):

| # | Quelle | Muster |
|---|---|---|
| 1 | Betreff | `#\s*(\d{4,})` |
| 2 | Body-Link, rohes HTML (greift auch in `href`) | `<host>/(?:self-service/)?requests/(\d+)` |
| 3 | Body-Text, HTML entfernt | `#\s*(\d{4,})` |
| 4 | `X-Xurrent-*`, `Message-ID`, `References` | `requests?[._/-](\d{4,})` |

Schritt 2 trägt in der Praxis am meisten: von 266 erkannten Mails kamen 184 über den
Body-Link und nur 169 über den Betreff — Notiz-Mails (`Note added to request: …`) haben
keine Nummer im Titel. Erkannt werden beide Link-Formen, weil Agenten-Oberfläche und
Kundenportal dieselbe Request unterschiedlich verlinken:

```
https://example-it.xurrent.com/requests/80000008
https://example-customers.xurrent.com/self-service/requests/80000008
```

Die Hosts in `XURRENT_INSTANCE_HOST` sind dabei eine Positivliste: ein weitergeleiteter
Link auf eine fremde Xurrent-Instanz wird ignoriert.

Jedes Muster verlangt `#` oder `/requests/` vor den Ziffern. Der Grund steht im Body der
echten Benachrichtigung: `654321-BEISPIELKUNDE SÜDTIROL` — eine 6-stellige
Kundennummer direkt neben der 8-stelligen Request-ID. Ein naives `\d{6,}` würde die
Kundennummer erwischen; dagegen gibt es einen eigenen Test.

Body und Header werden nur für Kandidaten geholt (Einzel-GET pro Mail), weil Graph
`internetMessageHeaders` ausschließlich beim GET einer einzelnen Nachricht mit `$select`
liefert. Das Ergebnis landet im Cache, jede Mail wird also genau einmal geparst.

## Warum wurde diese Mail nicht angefasst?

```bash
python sync.py --explain "aggiornamento account ospite"
```

```
  subject     richiesta di aggiornamento account ospite
  received    2026-09-23T14:27:09Z
  categories  ['Erledigt']
  in scope    no - does not carry 'Nicht erledigt' without 'Erledigt'
  request ids 80000009   via body-link
    80000009  completed  completed_at=2026-09-24T09:04:31+02:00
  verdict     request done -> the open category would be swapped for the done one
              categories are already correct
  open        https://outlook.office365.com/owa/?ItemID=...
```

Zeigt für jede Mail, deren Betreff den Suchtext enthält: Empfangszeit, aktuelle
Kategorien, ob sie überhaupt im Blickfeld liegt, welche Request-IDs gefunden wurden und
über welchen der vier Schritte, den Status jeder Request samt Abschlusszeitpunkt, und das
daraus folgende Urteil. Rein lesend, schreibt auch mit `--apply` nichts.

## Tagesprotokoll

`logs/daily/JJJJ-MM-TT.log` hält pro Tag fest, was tatsächlich passiert ist — das ist das
Protokoll, das man nach einem Fehler liest. Tabulatorgetrennt, damit es greppbar bleibt:

```
2026-10-08T15:21:45+02:00  APPLIED      80000007  completed             ['Nicht erledigt'] -> ['Erledigt']       ...
2026-10-08T15:21:45+02:00  APPLIED      80000006  waiting_for_customer  ['Nicht erledigt'] -> ['In Bearbeitung']  ...
2026-10-08T15:21:45+02:00  ERROR        80000004  Srv01 Memory       HTTP 412 ...                              <webLink>
2026-10-08T15:21:45+02:00  RUN          mode=APPLY scanned=95 candidates=94 done=1 waiting=17 ... duration=1.8s
```

| Satzart | Wann |
|---|---|
| `APPLIED` | eine Mail wurde umkategorisiert, mit Vorher und Nachher |
| `WOULD-APPLY` | Dry-run: so hätte er entschieden |
| `ERROR` | Mail konnte nicht gelesen oder nicht geschrieben werden, mit Outlook-Link |
| `RUN` | eine Zeile pro Lauf mit allen Zählern — zugleich der Nachweis, dass der Task überhaupt lief |

Zeitstempel und Dateiname sind **lokale** Zeit, weil das hier ein Mensch liest. Unterschied
zu den anderen beiden Ablagen: `logs/sync.log` ist das technische Log und rotiert weg,
`.state/state.json` enthält denselben Audit-Trail maschinenlesbar, und diese Dateien sind
nach Datum sortiert und bleiben liegen.

Ein Lauf bricht nie ab, weil das Protokoll nicht geschrieben werden kann — im Fehlerfall
gibt es nur eine Warnung in `sync.log`.

## Starter im Startmenü

Zwei Doppelklick-Starter liegen im Projektordner:

| Datei | Was sie tut |
|---|---|
| `sync-vorschau.cmd` | Dry-run, zeigt nur an was geändert würde |
| `sync-anwenden.cmd` | führt `--apply` aus |

Beide setzen die Konsole auf UTF-8 (sonst zerlegt es Umlaute und Emoji in den Betreffen),
wechseln über `%~dp0` in den Projektordner — das Projekt darf also umziehen, ohne dass ein
Pfad angepasst werden muss — und halten das Fenster mit `pause` offen, damit man die
Ergebnistabelle lesen kann. Am Ende stehen Exit-Code und der Hinweis auf `logs/daily/`.

Verknüpfungen ins Startmenü legt man so an:

```powershell
$proj = "$env:USERPROFILE\source\repos\outlook-xurrent-sync"
$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Outlook-Xurrent-Sync (anwenden).lnk")
$lnk.TargetPath = "$proj\sync-anwenden.cmd"
$lnk.WorkingDirectory = $proj
$lnk.IconLocation = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe,0"
$lnk.Save()
```

Die Verknüpfung erscheint danach unter **Start → Alle Apps**. Das Anheften an das
Kachelraster von Windows 11 lässt sich nicht skripten, das ist ein Rechtsklick →
**An „Start" anheften**.

## Alle X Minuten im Hintergrund

```bash
schtasks /Create /TN "Outlook-Xurrent-Sync" /SC MINUTE /MO 15 /IT /RL LIMITED /F /TR "\"%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe\" \"%USERPROFILE%\source\repos\outlook-xurrent-sync\sync.py\" --apply --no-login"
```

Alle 15 Minuten, `/MO` ist das Intervall. Warum genau diese Flags:

| Flag | Grund |
|---|---|
| `/SC MINUTE /MO 15` | Intervall in Minuten |
| `/IT` | läuft mit dem interaktiven Token des angemeldeten Benutzers — **kein Passwort wird gespeichert**, und der MSAL-Token-Cache im Benutzerprofil ist erreichbar |
| `/RL LIMITED` | keine erhöhten Rechte, das Skript braucht keine |
| `pythonw.exe` | kein Konsolenfenster, das alle 15 Minuten aufblitzt |
| `--no-login` | läuft der Anmelde-Token einmal ab, öffnet sich kein unsichtbares Anmeldefenster, das den Task für immer blockiert. Der Lauf endet mit Exit-Code 2 und einer Zeile in `logs/sync.log`; dann einmal `python sync.py --check` im Terminal ausführen |
| kein `cmd /c`, keine Umleitung | das Skript protokolliert selbst nach `logs/sync.log` und `logs/daily/` |

Ein Arbeitsverzeichnis ist nicht nötig: alle Pfade werden vom Projektordner aus aufgelöst,
nicht vom aktuellen Verzeichnis.

Prüfen, von Hand starten, wieder entfernen:

```bash
schtasks /Query /TN "Outlook-Xurrent-Sync" /V /FO LIST
```

```bash
schtasks /Run /TN "Outlook-Xurrent-Sync"
```

```bash
schtasks /Delete /TN "Outlook-Xurrent-Sync" /F
```

Hinweise:

- **Erst im Dry-run laufen lassen.** Derselbe Befehl ohne `--apply` schreibt nichts, füllt
  aber `logs/daily/` mit `WOULD-APPLY`-Zeilen. Nach einem Tag sieht man, was der Task
  getan hätte; der Umstieg auf `--apply` ist dasselbe Kommando nochmal, `/F`
  überschreibt den Task.
- Läuft nur, **solange du angemeldet bist** — so gewollt, siehe `/IT`.
- Überläufe sind kein Problem: die Aufgabenplanung startet keine zweite Instanz, solange
  die erste noch läuft. Ein Lauf dauert normalerweise 2–3 Sekunden.
- Der Refresh-Token verfällt bei Nichtbenutzung nach 90 Tagen; ein Intervalllauf hält ihn
  dauerhaft frisch.
- 15 Minuten sind für die Rate-Limits beider APIs unkritisch. Deutlich kürzer als 5
  Minuten bringt nichts, weil sich die Daten nur ändern, wenn du ein Ticket schließt.

## Dateien

```
sync.py                  Einstiegspunkt
conftest.py              macht das Projekt für pytest importierbar
outlook_xurrent/
  config.py              Konstanten, .env laden, validieren
  extract.py             Request-ID-Extraktion (rein, getestet)
  decide.py              Erledigt-Entscheidung + Kategorie-Umschreibung (rein, getestet)
  graph.py               Microsoft Graph: Login, Ordner, Mails, PATCH categories
  xurrent.py             Xurrent REST, nur GET, Batch-Lookup über ?id=1,2,3
  httputil.py            Retries, 429/retry-after, Fehlertypen
  state.py               JSON-State: ID-Cache, Audit-Trail, letzter Lauf
  daily.py               Tagesprotokoll, eine Datei je Tag
  report.py              Tabelle, Zähler, Liste der Mails ohne ID
  cli.py                 Ablauf und Exit-Codes
tests/                   pytest + Fixtures, ohne Netzwerk (Entscheidungslogik, ID-Erkennung,
                         Konfiguration, HTTP-Retries, State, Graph-Abfragen, ganze Läufe mit Attrappen)
requirements.txt         Laufzeit-Abhängigkeiten
requirements-dev.txt     zusätzlich pytest
.env.example             Vorlage für die .env
```

Zur Laufzeit entstehen `.state/` (Refresh-Token, ID-Cache, Audit-Trail jeder Änderung mit
vorher/nachher), `logs/sync.log` (technisch, rotierend, 1 MB × 5) und `logs/daily/` (ein
Protokoll je Tag, siehe oben). Alles drei plus `.env` enthält Geheimnisse oder Postfachdaten
und ist in `.gitignore`.

Das Projekt liegt absichtlich **nicht** in OneDrive: ein synchronisiertes `.git` kann bei
gleichzeitigem Zugriff Objektdateien beschädigen, und `.env` wäre dann in der Cloud.

## Troubleshooting

| Symptom | Ursache | Lösung |
|---|---|---|
| `Need admin approval` beim Login | Tenant sperrt User-Consent | App-Registrierung, siehe oben |
| `AUTH ERROR: Graph returned 401` | Token abgelaufen/widerrufen | `.state/msal_cache.json` löschen, neu anmelden |
| `AUTH ERROR: Graph returned 403` | delegierte Berechtigung fehlt | `GRAPH_SCOPES` prüfen, Cache löschen |
| Warnung `cannot read masterCategories` | `MailboxSettings.Read` fehlt | `VALIDATE_CATEGORIES=false` |
| `CONFIG ERROR: These Outlook categories do not exist` | Tippfehler im Kategorienamen | Schreibweise aus der Fehlermeldung übernehmen, die listet alle vorhandenen auf |
| `AUTH ERROR: Xurrent returned 401/403` | PAT falsch, oder Scope ohne `Request:Read`, oder `XURRENT_ACCOUNT` passt nicht zur Account-Condition | `python sync.py --check` |
| `HTTP 404` von Xurrent | falscher Regions-Root | Instanz ohne Regionssegment → `https://api.xurrent.com/v1` |
| viele `UNRESOLVED` | Requests für den Token nicht sichtbar | IDs stehen im Log; Mails bleiben unangetastet |
| viele `NO-ID` | Benachrichtigungen sehen anders aus als erwartet | Mail öffnen, Link/Betreff prüfen, `SUBJECT_REGEX` anpassen, dann `--no-cache` |
| `http 429 … sleep=Xs` im Log | Rate Limit | passiert automatisch, `retry-after` wird respektiert |
