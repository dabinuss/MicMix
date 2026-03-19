# MicMix VST Host Plan (Stability + Security First)

## 1) Zweck und Scope

Dieses Dokument definiert den vollumfassenden Projekt- und Technikplan fuer einen VST-Host in MicMix mit folgenden festen Zielen:

1. VST3-Effektketten fuer `Mic` und `Music` (Audioquelle).
2. Eigenes separates Effekt-UI.
3. Dynamische Effektliste je Kette:
   - Plugin hinzufuegen
   - Plugin entfernen
   - Reihenfolge aendern
   - Bypass / Aktiv-Status
4. Maximale Stabilitaet und Sicherheit bei Drittanbieter-Plugins.
5. Keine zusaetzliche manuelle Installation fuer Nutzer: alles ueber ein `.ts3_plugin` Paket.

Nicht-Ziele fuer v1:

1. Kein VST2.
2. Kein In-Process-Hosting im TeamSpeak-Plugin-Prozess.
3. Kein frei gestaltetes neues Design.

## 2) Harte Architekturentscheidung

VST-Plugins laufen in einem separaten Host-Prozess:

1. `micmix_win64.dll` bleibt TS3-Plugin und Echtzeit-Orchestrator.
2. `micmix_vst_host.exe` hostet VST3-Plugins isoliert.
3. Kommunikation via IPC (Shared Memory + Control Channel).

Begruendung:

1. Ein abstuerzendes VST darf TS3 nicht mitreissen.
2. Echtzeit-Callback in TS3 darf nicht durch VST-Aufrufe blockieren.
3. Sicherheits- und Trust-Logik ist in separatem Prozess deutlich kontrollierbarer.

## 3) Packaging und Installation

Die Verteilung bleibt fuer Nutzer "ein Schritt":

1. `.ts3_plugin` enthaelt:
   - `micmix_win64.dll`
   - `micmix_vst_host.exe`
   - benoetigte Ressourcen/Manifeste
2. Nutzer installiert wie gewohnt nur das `.ts3_plugin`.
3. Kein separater Installer, kein manuelles Nachinstallieren.

## 4) UI-Vorgabe (verbindlich): Exakt gleiches Design

Die neue VST-UI muss exakt die gleichen Designregeln wie das bestehende MicMix-UI einhalten.

Das bedeutet verbindlich:

1. Gleiche Farbwelt, gleiche Theme-Werte.
2. Gleiche Schriftfamilien, Schriftgroessen, Gewichte.
3. Gleiche Spacing-/Margin-/Card-Systematik.
4. Gleicher Kontrollstil (Buttons, Combo/Listen, Owner-Draw Verhalten).
5. Gleiche Hover/Focus/Disabled-Logik.
6. Gleiche Header-/Status-Badge Sprache und visuelle Grammatik.
7. Keine abweichende "neue" Designrichtung.

Technische Umsetzungsregel:

1. Die VST-UI wird auf denselben UI-Primitive-Regeln aufgebaut wie `settings_window.cpp`.
2. Vor Produktivfreigabe ist eine Pixel-/Style-Checkliste Pflicht.
3. "Aehnlich" ist nicht ausreichend, nur "gleiches System" ist akzeptiert.

## 5) Funktionsmodell

Es gibt zwei Effektketten, die gleichzeitig im UI sichtbar sind:

1. `Mic Chain`
2. `Music Chain`

Es gibt bewusst keine Tabs und keine Segmentumschaltung.

Jede Kette ist eine geordnete Liste von Effekt-Slots.

Slot-Eigenschaften:

1. Plugin-Referenz (VST3 Modul + Component/Class ID)
2. Aktiv/Bypass
3. Reihenfolgeindex
4. Persistenter Parameterzustand
5. Laufzeitstatus (running, bypassed, timed_out, crashed, blocked)

Pflichtfunktionen:

1. Add Plugin
2. Remove Plugin
3. Move Up / Move Down (oder Drag+Drop, wenn robust umgesetzt)
4. Bypass Toggle
5. Ketten-Reihenfolge wird sofort wirksam und persistent gespeichert

## 6) Sicherheitsmodell (verbindlich)

### 6.1 Trust und Zulassung

1. Nur `VST3` und `x64`.
2. Nur lokale Dateipfade.
3. Pfadvalidierung gegen Traversal/Reparse/Symlink-Missbrauch.
4. Vor Aktivierung: Scan + Kompatibilitaetspruefung.
5. Optionaler Signaturcheck (falls Signatur vorhanden).

### 6.2 Laufzeitschutz

1. Host-Prozess in Job Object mit klaren Ressourcen-/Lifetime-Regeln.
2. Heartbeat zwischen Plugin und Host.
3. Block-Timeout pro Audioauftrag.
4. Bei Timeout/Hang: sofortiger Bypass fuer betroffene Chain, kein Blockieren des TS3-Callbacks.
5. Bei Host-Crash: automatischer kontrollierter Neustart mit Backoff.
6. Crash-Haeufung fuehrt zu temporaerer Blockierung des verursachenden Plugins.

### 6.3 Fallback-Strategie

1. Host nicht verfuegbar -> MicMix laeuft ohne VST weiter.
2. Einzelnes Plugin defekt -> nur dieses Plugin/Chain wird bypassed.
3. Niemals "hard fail" fuer gesamte Sprachuebertragung.

## 7) Echtzeit- und Stabilitaetsregeln

Im TS3-Audio-Callback gelten weiterhin strikt:

1. Keine Locks.
2. Keine Heap-Allokation.
3. Keine blockierenden IPC-Calls.
4. Keine Dateisystem-/Netzwerk-Operationen.

Operationales Ziel:

1. Deadline-Miss im Callback darf nicht auftreten.
2. Wenn Host zu langsam ist: sofortige degradierende Strategie (Bypass) statt Blockieren.

## 8) Technischer Zielaufbau

### 8.1 Komponenten

1. `VstHostClient` (im Plugin):
   - Host-Prozess starten/ueberwachen
   - IPC aufbauen
   - Befehle fuer Kettenkonfiguration senden
2. `VstHostService` (in `micmix_vst_host.exe`):
   - Plugin-Scan
   - Plugin-Instanz-Lifecycle
   - Audio-Blockverarbeitung pro Chain
3. `EffectChainModel`:
   - Reihenfolge, Bypass, State
4. `EffectConfigStore`:
   - Persistenz + Migration + lastgood fallback

### 8.2 Audiofluss

1. Music:
   - `Source -> Host Music Chain -> AudioEngine PushMusicSamples`
2. Mic:
   - `Captured Mic -> Host Mic Chain -> Mix/Limit -> TS3 Out`

### 8.3 IPC-Design

1. Shared Memory Ringbuffer fuer Audiodaten.
2. Control-Channel fuer:
   - Host Status
   - Chain Updates
   - Scan Resultate
   - Fehlercodes
3. Sequenznummern und Zeitstempel fuer deterministische Verarbeitung.

### 8.4 Zustandsautomat (vereinfacht)

1. `Stopped`
2. `Starting`
3. `Running`
4. `DegradedBypass`
5. `Restarting`
6. `Blocked`

## 9) Persistenzmodell

Neue Konfigurationsfelder:

1. `vst.enabled`
2. `vst.host.autostart`
3. `vst.chain.mic[]`
4. `vst.chain.music[]`
5. Pro Slot:
   - `plugin.path`
   - `plugin.uid`
   - `slot.enabled`
   - `slot.order`
   - `slot.state_blob`
   - `slot.last_status`

Regeln:

1. Atomischer Write (`tmp -> replace`), wie bestehendes Config-Verhalten.
2. `config.lastgood` fuer Rollback.
3. Bei defekter Chain-Konfig: saubere Deaktivierung der fehlerhaften Slots statt globalem Ausfall.

## 10) UI-Spezifikation (funktional)

### 10.1 Fensterstruktur

1. Neues Fenster "VST Effects" (separat).
2. Ganz oben exakt gleicher Header wie MicMix-UI:
   - gleicher Titelstil
   - gleiche Header-Metadatenstruktur
   - gleiches Badge-Prinzip
   - Badge-Zustand fuer Effects: `Effects On` / `Effects Off`
3. Bereich 1 (oben, unter Header): Session/Control-Bereich
   - `Enable Effects`
   - `Disable Effects`
   - `Monitor Mix` (gleicher Control-Stil und gleiches Verhalten wie bestehendes MicMix-Monitoring)
4. Bereich 2 (Mitte): Liste aller aktiven `Music`-Plugins.
5. Bereich 3 (darunter): Liste aller aktiven `Mic`-Plugins.
6. Ganz unten: Log-/Statusbereich mit wichtigen Informationen und Fehlercodes.

### 10.2 Listenverhalten

1. Jede Zeile in beiden Listen zeigt:
   - Plugin-Name
   - Aktiv/Bypass
   - Laufzeitstatus
2. Aktionen pro Liste:
   - Add Plugin
   - Remove Plugin
   - Move Up / Move Down
   - Bypass Toggle
3. Entfernen ist sofort sichtbar und persistiert.
4. Reihenfolgeaenderung ist sofort sichtbar und wirksam.
5. Reihenfolge ist audioseitig deterministisch.
6. Beide Listen sind gleichzeitig sichtbar, ohne Umschalt-Tabs.

### 10.3 Monitoring-Vorgabe

1. `Monitor Mix` soll bevorzugt direkt den bestehenden MicMix-Monitorpfad wiederverwenden.
2. Das minimiert Risiko, vermeidet doppelte Audiowege und behaelt bekannte Stabilitaet.
3. Neue Monitoring-Implementierung ist nur zulaessig, wenn technische Gruende die Wiederverwendung verhindern.

### 10.4 Design-Paritaetsabnahme

Pflicht-Checklist vor Freigabe:

1. Farben exakt aus bestehendem Theme.
2. Schriftarten/Groessen exakt aus bestehender UI.
3. Gleiche Card-/Spacing-Raster.
4. Gleiche Owner-Draw Optik.
5. Gleiche Header-/Badge-Logik inklusive Effects-Badge.
6. Gleiche Bereichshierarchie (Header -> Controls -> Music-Liste -> Mic-Liste -> Log).
7. Keine Tabs und keine Segmentumschaltung.
8. Keine neue visuelle Sprache.

## 11) Phasenplan

### Phase 0: Architektur-Freeze

1. Entscheidungen finalisieren: VST3-only, Out-of-Process-only.
2. Sicherheitsrichtlinien final dokumentieren.
3. DoD und Performancebudgets finalisieren.

### Phase 1: Host-Grundfunktion

1. `micmix_vst_host.exe` Prozessstart/Stop.
2. Heartbeat/Watchdog.
3. Basale IPC fuer Status und Control.
4. Crash-Recovery mit Backoff.

### Phase 2: Audio-MVP Music Chain

1. Music-Chain in Host lauffaehig.
2. Einfache Add/Remove/Bypass/Order Operations.
3. Sicherer Fallback bei Hostproblemen.

### Phase 3: Audio-MVP Mic Chain

1. Mic-Chain Pfad anbinden.
2. RT-Schutz mit harten Timeouts.
3. Stabilitaet unter Last pruefen.

### Phase 4: UI v1

1. Separates Effektfenster implementieren.
2. Exakte Design-Paritaet zum bestehenden UI sicherstellen.
3. Dynamische Listenoperationen + Persistenz verbinden.

### Phase 5: Security Hardening

1. Plugin-Scan und Validierung.
2. Pfadhygiene und Trust-Levels.
3. Blacklist/Blocklist Logik.

### Phase 6: Test + Rollout

1. Soak/Chaos/Recovery Tests.
2. Beta-Flag opt-in.
3. Stufenweise Freigabe.

## 12) Teststrategie (stabilitaetszentriert)

### 12.1 Funktionale Tests

1. Add/Remove/Reorder auf beiden Chains.
2. Reihenfolge auditiv verifizierbar.
3. Bypass per Slot wirkt sofort.
4. Persistenz ueber TS3-Neustart korrekt.

### 12.2 Stabilitaetstests

1. 8h Soak mit aktiver Mic+Music Chain.
2. 100x Host-Neustart/Plugin-Reconnect.
3. 100x Reihenfolgewechsel unter Last.
4. Kein TS3-Absturz durch defekte Plugins.

### 12.3 Sicherheitstests

1. Ungueltige Pfade / Symlink-Fallen.
2. Defekte/inkompatible VST3-Dateien.
3. Host-Hang-Simulation (Timeout-Verhalten).
4. Crash-Flood und Blocklist-Verhalten.

### 12.4 Performanceziele

1. Keine Callback-Blockierung im TS3-Hotpath.
2. Keine unkontrollierte CPU-Spitze bei normalem Betrieb.
3. Glitch-freies Verhalten innerhalb definierter Zielhardware.

## 13) Release-Kriterien (Go/No-Go)

Go nur wenn alle Punkte erfuellt:

1. Kein reproduzierbarer TS3-Crash durch VST-Fehler.
2. Host-Ausfall fuehrt zu kontrolliertem Bypass, nicht Totalausfall.
3. UI-Design-Paritaet ist abgenommen ("exakt gleiches UI-System").
4. Add/Remove/Reorder/Bypass robust auf beiden Chains.
5. Persistenz + Migration + Recovery nachweislich stabil.
6. Security-Checks fuer Drittplugins aktiv.

## 14) Backlog-Startpaket (erste Tickets)

1. Architektur-ADR: Out-of-Process VST Host.
2. Host-Prozessgeruest + Watchdog.
3. IPC-Protokoll v1 (Control + Audio).
4. Music-Chain Verarbeitung v1.
5. Mic-Chain Verarbeitung v1.
6. Effekt-Configmodell + Save/Load Migration.
7. Neues VST-UI-Fenster mit exakter Design-Paritaet.
8. Security-Scanner + Trust/Blocklist.
9. Soak- und Chaos-Testsuite.
10. Packaging-Erweiterung fuer `.ts3_plugin` inkl. Host-Binary.

## 15) Verbindliche Leitplanken fuer Umsetzung

1. Stabilitaet vor Funktionsumfang.
2. Sicherheit vor Komfort.
3. Kein UI-Design-Drift.
4. Keine In-Process-VST-Ausnahmen.
5. Bei Unsicherheit: Fail-safe und expliziter Bypass statt riskanter Verarbeitung.
