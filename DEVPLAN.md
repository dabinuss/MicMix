# TS3 "MicMix" Plugin (Windows-only) — Vollständiger Umsetzungsplan (TS3 3.6.2 / Plugin API 26)

## 0) Ziel & Anforderungen

### Ziel
Ein TeamSpeak 3 **Client-Plugin** (Windows-only), das eine **zweite Audioquelle** (“Soundkanal”, z. B. Spotify) **zusätzlich** in den **Mikrofon-Stream** mischt, ohne die bestehenden Mic-Einstellungen im TS3 zu zerstören.

### Muss-Kriterien
- **UI im Plugin** (Settings-Fenster)
- Auswahl einer Soundquelle, die auf den Mic-Kanal **dazu** gemischt wird:
  - **Modus A:** WASAPI **Device Loopback** (Output-Gerät)
  - **Modus B:** **Spotify Session/App-only** (Audio-Session Capture)
  - **Nur einer aktiv** (A oder B)
- “Soundboard-like”: **Soundkanal sendet auch ohne Mic-Aktivität**
- Mic-Chain bleibt wie eingestellt (Noisegate/VAD/Preprocessor soll *so weit TS3 das erlaubt* unverändert bleiben)
- **Ducking** optional (aktivierbar)
- **Mute/Unmute nur für den Soundkanal** + **Hotkey** (TS3 Hotkey-System)
- **Kompatibel mit TS3 3.6.2**, **Plugin API 26**
- **Spotify Session Source robust inkl. Reacquire** (kein Experimental-Label im Release)
- **Resampling mit speexdsp + Drift-Kompensation** als Release-Pflicht
- **Machbarkeits-Spike bestanden** (Go/No-Go Kriterien aus 1.1)

---

## 1) Quellen & technische Grundlage (TS3 Plugin SDK API 26)

### TS3 Plugin Hooks (API 26)
- Audio: `ts3plugin_onEditCapturedVoiceDataEvent(...)`
- Menüs: `ts3plugin_initMenus(...)`, `ts3plugin_onMenuItemEvent(...)`
- Hotkeys: `ts3plugin_initHotkeys(...)`, `ts3plugin_onHotkeyEvent(...)`
- Config path: `ts3Functions.getConfigPath(...)` / `getPluginPath(...)`

### 1.1 Machbarkeits-Spike (Go/No-Go, 2-3 Tage)
Ziel: Kritische Annahmen vor Architektur-Feinschliff validieren.

- Spike A: **Inject ohne Mic-Aktivität**
  - Testmatrix: TS3 VAD an/aus, TS3 PTT, TS3 Continuous.
  - Kriterium: Remote-Client hört Soundkanal stabil bei Mic-Stille.
  - Ergebnis: Falls nicht robust möglich, Feature als “best effort” markieren und UI-Warnung verpflichtend.
- Spike B: **Spotify App-/Session-Capture**
  - Kriterium: Spotify lässt sich zuverlässig selektieren und nach Neustart reacquiren.
  - Ergebnis: Falls nicht robust möglich, Release ist blockiert bis Ursache behoben ist.
- Spike C: **PTT-/Talkstate-Erkennung für Ducking**
  - Kriterium: Prüfen, ob TS3-Status direkt nutzbar ist; sonst Plugin-Hotkey als alleinige Quelle.
  - Ergebnis: UI-Optionen nur für tatsächlich verfügbare Signale anzeigen.

Go/No-Go:
- Wenn Spike A fehlschlägt: kein “garantiertes Soundboard-like” in Muss-Kriterien.
- Wenn Spike B fehlschlägt: **kein v1.0 Release** (kritischer Blocker).
- Wenn Spike C fehlschlägt: Ducking-Mode “PTT detect” nur über Plugin-Hotkey.

### 1.2 Support-Matrix (verbindlich)
- TeamSpeak: 3.6.2, Plugin API 26
- Architektur: x64 als Primärziel, x86 optional mit separatem Build
- Windows: dokumentierte Zielversionen (mindestens Windows 10; Windows 11 empfohlen)
- Capture-Fähigkeiten:
  - Loopback: auf allen Zielversionen
  - Spotify Session: nur wenn OS/API-Pfad verfügbar und im Spike validiert

### 1.3 Festgelegte Audio-Dependency (v1.0 direkt)
- Resampler-Library: **speexdsp** (feste Entscheidung)
- Grund:
  - gute Echtzeit-Performance bei niedriger Latenz
  - stabile Qualität für Voice+Music Mixing
  - einfache Integration in native C/C++ Toolchains
- `libsamplerate` wird für dieses Projekt nicht als Primärpfad eingeplant.

---

## 2) System-Architektur (Module)

### 2.1 Module
1. **TS3Integration**
   - Implementiert alle TS3-Plugin-Exports (init/shutdown, menus, hotkeys, audio callback)
   - Reicht Audio-Frames an `AudioEngine` weiter
   - Öffnet UI über Menüpunkt “Mic Mixer Settings…”

2. **AudioEngine (Realtime)**
   - Läuft im TS3 Audio-Callback (sehr performance-kritisch)
   - Liest Music-Frames aus Ringbuffer
   - Resample/Channelmix (falls nötig)
   - Ducking optional
   - Mix + Limiter/Softclip
   - Schreibt zurück in `short* samples` und setzt `*edited = 1`

3. **AudioSourceManager**
   - Verwaltet **genau eine aktive Quelle**:
     - `LoopbackSource` (Output Device Loopback)
     - `SpotifySessionSource` (App-/Session-Capture)
   - Start/Stop/Restart bei Moduswechsel
   - Auto-Reacquire (für Spotify Session)

4. **UI (Settings Window)**
   - Moduswahl (Loopback vs Spotify Session)
   - Device/Session-Auswahl
   - Gain/Volume
   - Ducking Einstellungen
   - Mute Button (nur Soundkanal)
   - Statusanzeigen (Running/Stopped/Reacquiring/Error)

5. **ConfigStore**
   - Speichert/liest Settings (JSON oder INI)
   - Pfad über TS3 config path
   - Optional: “Autostart Source” bei Plugin-Init

### 2.2 Realtime Constraints (verbindlich)
- Im TS3 Audio-Callback:
  - **keine Locks/Mutexe**
  - **keine Heap-Allokationen**
  - **keine I/O-Aufrufe** (Datei, Netzwerk, COM-Initialisierung)
- Parameterübergabe UI -> AudioEngine nur über atomare Werte oder lock-free Snapshot.
- Ringbuffer ist Single-Producer/Single-Consumer (SPSC), lock-free.
- Telemetrie-Zähler:
  - underruns
  - overruns
  - clipped samples
  - source reconnect count
- Overload-Strategie (degradierend statt abstürzend):
  - bei wiederholter Callback-Überlast: optionale Metering-Funktionen deaktivieren
  - Resampler-Qualität stufenweise reduzieren (innerhalb definierter Mindestqualität)
  - bei anhaltender Überlast: Music temporär stumm + sichtbarer Warnstatus, Mic bleibt aktiv

### 2.3 Persistente Settings (verbindlich)
Beim Klick auf `Apply/Save` und beim geordneten Plugin-Shutdown werden folgende Werte persistiert:

- `config.version`
- `source.mode` (`loopback` | `spotify_session`)
- `source.loopback.device_id` (stabile Device-ID, nicht nur Anzeigename)
- `source.spotify.process_name` (Default `Spotify.exe`)
- `source.spotify.session_id` (wenn verfügbar; sonst Reacquire per Prozessname)
- `source.autostart_enabled` (bool)
- `mix.music_gain_db`
- `mix.force_tx_enabled` (bool)
- `mix.buffer_target_ms`
- `mix.music_muted` (bool)
- `ducking.enabled` (bool)
- `ducking.mode` (`mic_rms` | `plugin_hotkey` | `ts3_talkstate`)
- `ducking.threshold_dbfs`
- `ducking.attack_ms`
- `ducking.release_ms`
- `ducking.amount_db`
- `ui.last_open_tab`

Ladeverhalten:
- In `ts3plugin_init()` wird die komplette Config geladen, validiert und auf Defaults gefallbackt.
- Ungültige oder veraltete Keys blockieren den Start nicht; sie werden geloggt und auf Default gesetzt.
- Nach Start zeigt die UI den tatsächlich aktiven Zustand (inkl. Reacquire/Error) und nicht nur gespeicherte Wunschwerte.

Write-Sicherheit:
- Config-Write ist atomar (`config.tmp` schreiben, flushen, dann replace/rename auf `config.ini`).
- Letzte gültige Config wird als `config.lastgood.ini` vorgehalten.
- Bei defekter/inkompletter Config: Fallback auf `lastgood`, sonst Defaults + Warnung im Log/UI.

Migration:
- Jede Config hat `config.version`.
- Beim Laden wird eine deterministische Migration auf die aktuelle Version ausgeführt.
- Nicht migrierbare Felder werden verworfen, aber Start bleibt möglich.

### 2.4 Diagnostik & Logging (verbindlich)
- Strukturierte Fehlercodes für Source/Config/Threading (z. B. `device_not_found`, `session_lost`, `config_parse_failed`).
- Rotierende lokale Logdatei im Plugin-Config-Pfad (z. B. `micmix.log`, max. Größe + Anzahl Backups).
- Log-Level: `error`, `warn`, `info`, `debug` (default: `info`).
- In Release-Builds keine sensitiven Daten loggen (keine personenbezogenen Inhalte, keine vollständigen Gerätedumps ohne Opt-in).

---

## 3) Audio-Pipeline (Detail)

### 3.1 Interne Audioformate
- **Intern:** `float` Samples in `[-1, +1]`, **mono**, Ziel: **48kHz**
- **TS3 Callback:** `short* samples` mit `sampleCount` und `channels`

### 3.2 Konvertierung
- `short -> float`: `f = s / 32768.0f`
- `float -> short`: clamp/softclip, `s = (short)(f * 32767)`

### 3.3 Ringbuffer / Jitterbuffer (Pflicht)
- Producer: Source-Thread schreibt Music-Frames
- Consumer: TS3 Callback zieht exakt die Anzahl Samples pro Call
- Buffer-Target: 40–80ms (configurierbar)
- Wenn Buffer leer: Zero-Fill + underrun counter inkrementieren
- Wenn Buffer voll: älteste Samples verwerfen oder write skip (fest definieren) + overrun counter

### 3.4 Resampling / Channelmix
- Quellen können 44.1kHz/48kHz und stereo/mono sein
- Stereo -> Mono: `(L + R) * 0.5`
- Resampler:
  - **v1.0 Direktpfad:** `speexdsp` (`speex_resampler`)
  - Qualitätsstufe per Config (Default: Qualität 5-7, je nach CPU-Budget)
- Drift-Kompensation:
  - Quellen- und TS3-Clock können auseinanderlaufen
  - Ziel: Buffer-Level im Zielkorridor halten (z. B. 60ms)
  - adaptive Ratio-Korrektur (ppm-Bereich) statt statischer Ratio

### 3.5 Mixing
- `out = mic + music * musicGainEffective`
- `musicGainEffective = musicGain * duckFactor`
- Limiter/Softclip danach
- Gain-Safety:
  - harte Grenzen für `musicGain` (z. B. min/max dB)
  - kein vollständiger Limiter-Bypass im Release
  - Clip-Events zählen und in Diagnostik ausweisen

### 3.6 “Soundboard-like” (sendet ohne Mic)
Problem: VAD/Noisegate kann Musik blocken, wenn TS3 nach dem Callback gated.

**Plan:**
- Option: **Force TX while music playing**
  - Wenn aktiv & Music nicht gemutet & Source liefert Signal:
    - Im Callback immer “nicht-leeren” Frame erzeugen (via Music)
    - `*edited = 1` setzen
- Default:
  - Wenn User VAD nutzt: Force TX standardmäßig an
  - Bei PTT: Force TX optional
- Wichtige Einschränkung:
  - Verhalten ist von TS3-Talkstate/Client-Einstellungen abhängig.
  - Ergebnis aus Spike A bestimmt, ob das Feature als “garantiert” oder “best effort” dokumentiert wird.

---

## 4) Audio-Quellen: Modus A und Modus B (nur eine aktiv)

### 4.1 Modus A: WASAPI Device Loopback
- Capture vom ausgewählten Output-Gerät (Loopback)
- UI:
  - Dropdown: Output Device
  - Hinweis: Loopback nimmt alles (inkl. TS3 Output) → Feedback-Risiko
- Empfehlung in UI:
  - “TS3 Output getrennt von Spotify Output” oder “Spotify Session Mode nutzen”
- Wenn gespeicherte Device-ID beim Start fehlt:
  - Status `Error` mit klarer Ursache
  - kein Crash, keine Blockade des Plugin-Starts
  - User kann im UI sofort ein neues Device wählen und `Apply/Save` ausführen

### 4.2 Modus B: Spotify Session/App-only Capture
- Ziel: Nur Spotify Audio (ohne TS3 / Games / Browser)
- UI:
  - Dropdown: Session/App (Spotify.exe)
  - Button: Rescan Sessions
  - Status: Connected / Waiting for Spotify / Reacquiring
- Auto-Reacquire:
  - Timer/Thread prüft alle X Sekunden, ob Session wieder vorhanden ist
  - Bei Spotify restart → Source verbindet neu
- Fallback-Verhalten (verbindlich):
  - Wenn Spotify nicht erreichbar ist, läuft Plugin im Zustand `Reacquiring` oder `Error` weiter, ohne Absturz.
  - Während `Reacquiring`: Music-Anteil wird als Stille behandelt (Mic bleibt unverändert nutzbar).
  - Optionaler Auto-Switch auf Loopback bleibt User-Option (default: aus), aber Fehlerbehandlung gilt immer.

### 4.3 Moduswechsel
- “Switch Mode”:
  - Stop aktuelle Source
  - Flush Ringbuffer
  - Start neue Source
  - UI Status update

### 4.4 Source-State-Machine (verbindlich)
Zustände:
- `Stopped`
- `Starting`
- `Running`
- `Reacquiring`
- `Error`

Transitions:
- `Stopped -> Starting`: User Start/Autostart
- `Starting -> Running`: Capture erfolgreich
- `Starting -> Error`: Initialisierung fehlgeschlagen
- `Running -> Reacquiring`: Session verloren (z. B. Spotify restart)
- `Reacquiring -> Running`: Session wieder verfügbar
- `Reacquiring -> Error`: Timeout überschritten
- `* -> Stopped`: User Stop oder Plugin Shutdown

Regeln:
- Exponential Backoff für Reacquire (z. B. 1s, 2s, 4s, max 15s)
- Nach `N` Fehlversuchen Error-Status mit konkreter Ursache
- Jeder Source-Fehler ist recoverable oder führt deterministisch in `Error`, niemals in undefinierten Zustand.
- Fehlergrund wird für UI/Logs als Code + Text geführt (z. B. `device_not_found`, `session_lost`, `open_failed`, `timeout`).

### 4.5 Fehlerbehandlung bei fehlenden/instabilen Quellen (verbindlich)
Beim Laden:
- Gespeicherte Source nicht verfügbar -> Plugin startet trotzdem, Source bleibt `Error`, UI fordert Re-Select.
- Config wird nicht verworfen; nur ungültige Source-spezifische Werte werden auf sichere Defaults gesetzt.

Zur Laufzeit:
- App schließt sich / Session verschwindet -> `Running -> Reacquiring` mit Backoff, ohne Audio-Thread-Crash.
- Device invalidiert (Treiberwechsel, Device entfernt) -> `Running -> Reacquiring` bzw. `Error` nach Timeout.
- Source liefert fehlerhafte Frames/Timeouts -> Frame verwerfen, Stille einmischen, Counter erhöhen, State-Übergang auslösen.
- Device Hotplug / Default-Device-Wechsel:
  - erkannte Device-Änderungen triggern Re-Resolve per gespeicherter Device-ID
  - falls ID nicht mehr auflösbar: `Error` + UI-Recovery-Pfad
- Dynamischer Formatwechsel der Quelle:
  - Wechsel von Sample-Rate/Kanalzahl im Betrieb löst kontrollierte Reinitialisierung der Source-Chain aus
  - während Reinit kein Crash/Block; Music ggf. kurz stumm

Nutzerwirkung:
- Mic bleibt durchgehend nutzbar.
- Music-Kanal kann temporär stumm werden, aber darf nie TS3-Client oder Plugin instabil machen.
- UI zeigt sofort den realen Zustand und konkrete Recovery-Hinweise.

---

## 5) Ducking (optional)

### 5.1 UI Optionen
- [ ] Ducking enabled
- Mode:
  - ( ) Mic RMS detect
  - ( ) Plugin Hotkey detect
  - ( ) TS3 Talkstate detect (nur wenn verfügbar; sonst ausgegraut)
- Amount: 0 … -30 dB
- Attack / Release (ms)
- Threshold (nur Mic RMS)

### 5.2 Implementierung
- Mic RMS pro Frame im Callback messen
- DuckFactor rampen:
  - Attack: schnell Richtung Ziel (z. B. -12 dB)
  - Release: langsamer zurück
- Nur MusicGain wird beeinflusst, Mic bleibt unverändert
- Falls nur Plugin-Hotkey verfügbar: eindeutige UI-Beschriftung “nicht identisch mit TS3 PTT”

---

## 6) Mute/Unmute & Hotkeys (nur Soundkanal)

### 6.1 UI
- Button: “Music Mute/Unmute”
- Anzeige: “MUTED” / “LIVE”
- Optional: “Remember state” (persist in config)

### 6.2 Hotkeys (TS3 System)
Hotkey Keywords:
- `music_toggle_mute`
- optional: `music_push_to_play` (hold = unmuted)

Flow:
- `ts3plugin_initHotkeys(...)` registriert Keywords + Beschreibung
- User bindet Hotkey in TS3 Hotkey-Dialog
- `ts3plugin_onHotkeyEvent(keyword)` togglet plugin state

Randfall-Regeln:
- Debounce gegen Repeat-Events (insb. bei Toggle-Hotkeys).
- `music_push_to_play` ist idempotent für Key-Down/Key-Up Sequenzen.
- Konflikte mit anderen TS3-Hotkeys werden im UI-Hinweis dokumentiert (Priorität/Erwartung).

---

## 7) UI (Windows-only, Standard-Plugin-UX)

### 7.1 Öffnen
- Plugin-Menü → “Mic Mixer Settings…”

### 7.2 Layout (konkret)
#### Tab: Source
- Radio:
  - ( ) Loopback (Output Device)
  - ( ) Spotify Session (App-only)
- Loopback:
  - Output Device Dropdown
- Spotify Session:
  - Session Dropdown (Spotify.exe)
  - Rescan Button
  - Status line
  - Bei fehlender Verfügbarkeit: klare Hinweise + deaktivierte Controls

#### Tab: Mix
- Music Volume Slider
- Force TX while music playing (Checkbox)
- Buffer latency slider (ms)
- Optional: Music Level Meter
- Hinweisfeld zu TS3-Preprocessing:
  - Empfehlungen für VAD/Noisegate/AGC je nach Nutzungsmodus (Music-only, Voice+Music, PTT)

#### Tab: Ducking
- Enable toggle
- Mode selector
- Threshold / Attack / Release / Amount

#### Tab: Hotkeys
- List of keywords + how to bind in TS3
- Optional: “Copy keyword” buttons

Bottom:
- Apply / Save
- Start / Stop
- Status (Running / Stopped / Error)
- Diagnostic Counters (optional sichtbar): underrun/overrun/reconnect

---

## 8) TS3 Lifecycle & Threading

### 8.1 Init
`ts3plugin_init()`:
- Load config
- Initialize AudioEngine
- Initialize UI components (lazy ok)
- Optional autostart: start source if enabled

### 8.2 Source Threads
- `AudioSource` läuft in eigenem Thread:
  - capture -> convert -> resample -> push into ringbuffer
- Stop:
  - signal stop -> join thread -> flush buffer
- UI-Thread darf niemals direkt Audioobjekte mutieren; nur Commands/atomare Parameter setzen.
- Shutdown-Reihenfolge:
  1. Source stop signal
  2. Thread join (mit Timeout + Fallback logging)
  3. Engine teardown
  4. UI teardown

### 8.3 Realtime Callback
`ts3plugin_onEditCapturedVoiceDataEvent(...)`:
1. Mic `short -> float`
2. Music aus Ringbuffer lesen (oder Zero-Fill bei underrun)
3. Ducking update (zustandslos pro Frame + geglättete Gain-Rampe)
4. mix + limit + float -> short
5. `*edited = 1`, wenn Musik hörbar beigemischt ist oder Force TX aktiv ist und Quelle Signal liefert; sonst `*edited = 0` für transparentes Pass-through

### 8.4 Power-/Session-Lifecycle (verbindlich)
- Windows Sleep/Resume:
  - nach Resume werden Quellen neu validiert und bei Bedarf kontrolliert reinitialisiert
  - kein hängender Capture-Thread, kein Deadlock im Callback
- TS3 ServerConnection-Scopes:
  - Plugin-Zustand ist pro `serverConnectionHandlerID` konsistent definiert
  - kein unbeabsichtigtes Cross-Mixing zwischen parallelen TS3-Verbindungen
- TS3 Restart/Plugin Reload:
  - vollständiger Cleanup aller Worker-Threads und Handles garantiert

---

## 9) Build & Packaging (TS3 3.6.2 / API 26)

### 9.1 SDK
- Teamspeak 3 Client Plugin SDK **API 26**
- Build: DLL (x64, ggf. auch x86 wenn TS3 x86 genutzt wird)
- Toolchain fixieren (MSVC Version + Windows SDK Version dokumentieren)
- Buildsystem festlegen (z. B. CMake + Presets) für reproduzierbare Builds
- Exportliste/ABI prüfen (alle geforderten TS3-Plugin-Exports vorhanden)
- Runtime-Strategie dokumentieren (statisch vs. benötigte VC++ Runtime)
- Third-Party Dependencies fixieren:
  - `third_party/ts3_sdk` (lokal, nicht versioniert)
  - `third_party/speexdsp` (lokal/vendor; klare Lizenzdokumentation)
  - `speexdsp` Version/Commit wird im Build fixiert (reproduzierbar)

### 9.2 Deployment
- `plugins/yourplugin.dll`
- optional: `plugins/yourplugin/` assets
- config files in TS3 config directory
- Release-Artefakte:
  - `x64` und optional `x86` getrennt
  - Versionsnummer + Changelog
  - SHA256 Checksums
- Optional aber empfohlen: Code Signing der DLL
- Lizenz-/Third-Party-Hinweise:
  - `THIRD_PARTY_NOTICES.md` mit `speexdsp`/weiteren Dependencies

---

## 10) Testplan (Definition of Done)

### 10.1 Core Tests
- Mic-only (plugin enabled, source off): identisch wie vorher
- Music-only: mic silent, music on -> remote hears music
- VAD on:
  - verify music transmits
  - if blocked -> enable Force TX and retest
- Ducking:
  - speaking lowers music; stop speaking raises music
- Music mute:
  - button and hotkey toggle only music, mic unchanged
- Persistenz:
  - Quelle/Device/Session, Mix- und Ducking-Parameter speichern -> TS3 Neustart -> identische Werte geladen
  - letzter Mute- und Autostart-Zustand wird korrekt wiederhergestellt
  - bei nicht mehr verfügbarem Device/Session: sauberer Fallback + klare UI-Fehlermeldung statt Crash
  - defekte Config-Datei simulieren -> Fallback auf `lastgood` oder Defaults ohne Startabbruch
  - ältere Config-Version laden -> Migration läuft deterministisch und verlustarm

Messbare Akzeptanzkriterien:
- Zusätzliche End-to-End Latenz durch Plugin: Ziel `< 10 ms`
- Callback-CPU-Budget: keine hörbaren Glitches unter Zielhardware bei 60 Minuten Dauerlast
- Clipping-Rate nach Limiter: nahe 0 bei nominalen Einstellungen
- Underruns/Overruns: im Normalbetrieb 0; bei Stress reproduzierbar geloggt
- Spotify Reacquire: nach Spotify-Neustart automatische Wiederverbindung innerhalb definierter Timeout-Grenze
- Sleep/Resume: nach Resume innerhalb definierter Zeit wieder stabiler `Running`-Status oder klarer `Error`-Status
- Multi-Server: kein Cross-Mixing zwischen parallelen Verbindungen

### 10.2 Risk Tests
- Loopback feedback:
  - verify TS3 output not re-injected (or warn user)
- Spotify session reacquire:
  - restart spotify -> plugin reconnects automatically
- Source weg beim Laden:
  - gespeicherte Device/Session existiert nicht -> Plugin startet, zeigt `Error`, UI-Recover möglich
- Runtime-Source-Ausfall:
  - Spotify/App schließen während Betrieb -> Übergang `Running -> Reacquiring -> Running/Error` ohne Crash
- Loopback-Device-Ausfall:
  - Device deaktivieren/entfernen während Betrieb -> sauberer State-Übergang, Mic bleibt aktiv
- Device Hotplug / Default-Wechsel:
  - neues Default-Gerät setzen/alte entfernen -> Source-Resolver verhält sich deterministisch
- Dynamischer Formatwechsel:
  - Quelle mit wechselnder Sample-Rate/Kanalzahl -> kontrollierte Reinitialisierung ohne Crash
- Fehlerhafte Source-Frames:
  - timeouts/invalid frames -> Stille statt Absturz, Fehlerzähler steigt, Zustand wird korrekt aktualisiert
- Sleep/Resume:
  - Windows Standby und Resume während aktiver Source -> kein Deadlock, sauberes Reacquire
- Hotkey-Randfälle:
  - schneller Repeat auf Toggle -> kein Flattern/Fehlzustand
  - Push-to-play Key-Down/Key-Up Sequenzen bleiben konsistent
- Overload-Fall:
  - künstliche CPU-Last -> degradierende Qualitätsstrategie greift, Mic bleibt nutzbar
- Config-Schreibabbruch:
  - erzwungener Abbruch während Save -> beim nächsten Start konsistente Fallback-Config
- Mehrere TS3-Verbindungen parallel:
  - kein State-Leak oder ungewolltes Routing zwischen Verbindungen
- Stability:
  - run 1h streaming audio -> no drift/dropouts
- Drift-Test:
  - 4h Dauerbetrieb, Buffer-Level bleibt im Zielkorridor
- Mode-Switch Robustheit:
  - 100x Wechsel Loopback <-> Spotify ohne Crash/Deadlock
- Shutdown/Restart:
  - TS3 beenden/starten während Source aktiv -> sauberer Cleanup

---

## 11) Roadmap

### Phase 0 (Gate, blockierend)
- Machbarkeits-Spike aus 1.1
- Go/No-Go Entscheidung
- Nur bei bestandenem Gate: Implementierung starten

### Phase 1 (Core v1.0)
- Loopback Source + Spotify Session Source (beide produktionsreif)
- Robuster Source-Manager mit Reacquire-State-Machine
- UI komplett (Source/Mix/Ducking/Hotkeys/Status)
- Realtime Constraints aus 2.2 vollständig umgesetzt

### Phase 2 (Audioqualität v1.0)
- `speexdsp` Resampler integriert (kein linearer Release-Pfad)
- Drift-Kompensation finalisiert
- Limiter/Softclip final abgestimmt
- Basis-Telemetrie (underrun/overrun/clipping/reconnect) produktionsreif

### Phase 3 (Release-Hardening)
- Vollständiger Testplan aus Kapitel 10 bestanden
- Packaging, Versionierung, Checksums, Dokumentation final
- Release Candidate -> v1.0

---

## 12) Abschluss: Deliverables

- Plugin DLL + folder
- Settings UI
- Config storage
- Rotierende Diagnose-Logs + Fehlercode-Referenz
- Dokumentation:
  - How to bind hotkeys
  - Recommended TS3 output separation for Loopback mode
  - Known limitations + troubleshooting
  - Support-Matrix (OS/Arch/Features)
  - Go/No-Go Ergebnis des Machbarkeits-Spikes
  - Third-Party/Lizenzhinweise (`THIRD_PARTY_NOTICES.md`)
  - Kurz-Hinweis zu rechtlicher Nutzung von Audioinhalten (Urheberrecht/Plattformregeln)
