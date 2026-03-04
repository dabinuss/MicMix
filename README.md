# MicMix

Windows-only TeamSpeak 3 Client-Plugin, das eine zweite Audioquelle (z. B. Spotify) in den Mikrofon-Stream mischt.

## Status

Aktuell in Planung/Setup.  
Die technische Spezifikation liegt in [`DEVPLAN.md`](./DEVPLAN.md).

## Ziel

MicMix soll neben dem Mikrofon einen separaten "Soundkanal" senden, ohne bestehende TS3-Mikrofon-Einstellungen unnötig zu beeinträchtigen.

## Geplanter Funktionsumfang

### MVP
- WASAPI Loopback als zusätzliche Audioquelle
- Plugin-UI mit Start/Stop, Gain, Mute, Force TX
- Ringbuffer + Resampling + grundlegendes Ducking

### v1.0
- Spotify Session/App-only Capture
- Robustes Reacquire bei Session-Verlust
- Verbesserter Resampler/Limiter

## Voraussetzungen (Nutzung)

- Windows 10/11 (x64 empfohlen)
- TeamSpeak 3 Client `3.6.2`
- Plugin-Kompatibilität: TeamSpeak Plugin API `26`
- Für Loopback-Modus: korrektes Audio-Routing empfohlen (um Feedback zu vermeiden)

## Repository-Struktur

- [`DEVPLAN.md`](./DEVPLAN.md): vollständiger Umsetzungsplan, Architektur, Teststrategie, Roadmap

## Hinweise

- Das Projekt ist explizit auf TeamSpeak 3 und Windows ausgerichtet.
- Feature-Verfügbarkeit (insbesondere Spotify Session Capture) hängt von den im DEVPLAN definierten Machbarkeits-Spikes ab.
