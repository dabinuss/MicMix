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
  - Ergebnis: Falls nicht robust möglich, Spotify-Modus auf “experimental” setzen und Loopback als Standard.
- Spike C: **PTT-/Talkstate-Erkennung für Ducking**
  - Kriterium: Prüfen, ob TS3-Status direkt nutzbar ist; sonst Plugin-Hotkey als alleinige Quelle.
  - Ergebnis: UI-Optionen nur für tatsächlich verfügbare Signale anzeigen.

Go/No-Go:
- Wenn Spike A fehlschlägt: kein “garantiertes Soundboard-like” in Muss-Kriterien.
- Wenn Spike B fehlschlägt: Spotify in v1.0 optional/experimental.
- Wenn Spike C fehlschlägt: Ducking-Mode “PTT detect” nur über Plugin-Hotkey.

### 1.2 Support-Matrix (verbindlich)
- TeamSpeak: 3.6.2, Plugin API 26
- Architektur: x64 als Primärziel, x86 optional mit separatem Build
- Windows: dokumentierte Zielversionen (mindestens Windows 10; Windows 11 empfohlen)
- Capture-Fähigkeiten:
  - Loopback: auf allen Zielversionen
  - Spotify Session: nur wenn OS/API-Pfad verfügbar und im Spike validiert

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
  - MVP: linear resample
  - v1.0: speexdsp oder libsamplerate
- Drift-Kompensation:
  - Quellen- und TS3-Clock können auseinanderlaufen
  - Ziel: Buffer-Level im Zielkorridor halten (z. B. 60ms)
  - MVP: leichte adaptive Ratio-Korrektur (ppm-Bereich) statt statischer Ratio

### 3.5 Mixing
- `out = mic + music * musicGainEffective`
- `musicGainEffective = musicGain * duckFactor`
- Limiter/Softclip danach

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

### 4.2 Modus B: Spotify Session/App-only Capture
- Ziel: Nur Spotify Audio (ohne TS3 / Games / Browser)
- UI:
  - Dropdown: Session/App (Spotify.exe)
  - Button: Rescan Sessions
  - Status: Connected / Waiting for Spotify / Reacquiring
- Auto-Reacquire:
  - Timer/Thread prüft alle X Sekunden, ob Session wieder vorhanden ist
  - Bei Spotify restart → Source verbindet neu
- Fallback (optional):
  - Wenn Spotify nicht erreichbar → optional auf Loopback wechseln (nur wenn User aktiv erlaubt)

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

---

## 9) Build & Packaging (TS3 3.6.2 / API 26)

### 9.1 SDK
- Teamspeak 3 Client Plugin SDK **API 26**
- Build: DLL (x64, ggf. auch x86 wenn TS3 x86 genutzt wird)
- Toolchain fixieren (MSVC Version + Windows SDK Version dokumentieren)
- Buildsystem festlegen (z. B. CMake + Presets) für reproduzierbare Builds
- Exportliste/ABI prüfen (alle geforderten TS3-Plugin-Exports vorhanden)
- Runtime-Strategie dokumentieren (statisch vs. benötigte VC++ Runtime)

### 9.2 Deployment
- `plugins/yourplugin.dll`
- optional: `plugins/yourplugin/` assets
- config files in TS3 config directory
- Release-Artefakte:
  - `x64` und optional `x86` getrennt
  - Versionsnummer + Changelog
  - SHA256 Checksums
- Optional aber empfohlen: Code Signing der DLL

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

Messbare Akzeptanzkriterien:
- Zusätzliche End-to-End Latenz durch Plugin: Ziel `< 20 ms` (MVP), `< 10 ms` (v1.0)
- Callback-CPU-Budget: keine hörbaren Glitches unter Zielhardware bei 60 Minuten Dauerlast
- Clipping-Rate nach Limiter: nahe 0 bei nominalen Einstellungen
- Underruns/Overruns: im Normalbetrieb 0; bei Stress reproduzierbar geloggt

### 10.2 Risk Tests
- Loopback feedback:
  - verify TS3 output not re-injected (or warn user)
- Spotify session reacquire:
  - restart spotify -> plugin reconnects automatically
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

### Phase 0 (Gate)
- Machbarkeits-Spike aus 1.1
- Go/No-Go Entscheidung und ggf. Scope-Anpassung

### MVP
- Loopback Source
- UI with Start/Stop, Gain, Mute, Force TX
- Ringbuffer, basic linear resample
- Basic ducking (mic RMS)
- Realtime Constraints aus 2.2 erfüllt
- Basis-Telemetrie (underrun/overrun/clipping) im Debug-Log

### v1.0
- Spotify Session Source (app-only)
- robust reacquire + UI status
- improved resampler (speexdsp)
- improved limiter
- Drift-Kompensation finalisiert
- Messkriterien aus Testplan vollständig erreicht

---

## 12) Abschluss: Deliverables

- Plugin DLL + folder
- Settings UI
- Config storage
- Dokumentation:
  - How to bind hotkeys
  - Recommended TS3 output separation for Loopback mode
  - Known limitations + troubleshooting
  - Support-Matrix (OS/Arch/Features)
  - Go/No-Go Ergebnis des Machbarkeits-Spikes
