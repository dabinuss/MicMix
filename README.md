# MicMix

> **Notice:** This project was developed with the help of Codex. Please use it with appropriate caution. Everyone is encouraged to review the publicly available code and report bugs or issues.

![MicMix Logo](assets/branding/MicMix%20Logo.png)

MicMix is a Windows plugin for TeamSpeak 3 that mixes a second audio source (music, browser audio, media player audio, etc.) into your microphone stream.

It is built for people who want to speak and play audio in the same voice channel without setting up complex virtual cable chains.

## Screenshot

Add your settings screenshot here (recommended path: `assets/screenshots/micmix-settings.png`).

<img width="701" height="691" alt="grafik" src="https://github.com/user-attachments/assets/d76d98e7-b9bb-424a-96ca-c830b10a1fc8" />


## What MicMix Can Do

- Mix your mic + one selected audio source into TeamSpeak voice transmission.
- Capture audio either from `Loopback` (system playback device) or `App Session` (a specific running app process, for example Spotify, browser, VLC).
- Control music level with a gain slider.
- Mute/unmute music quickly with a hotkey.
- Keep sending music even while you are not talking (`Force TX`).
- Show live level meters for music and microphone.
- Auto-start source capture on plugin startup (optional).
- Save settings and restore them on next launch.

## Typical Use Cases

- DJ in TeamSpeak: play music for your channel while still talking normally.
- Roleplay or gaming immersion: use it as an in-car radio replacement in RP servers.
- Community events: stream intro music, ambience, or short jingles.
- Watch parties: share browser or media app audio while coordinating by voice.

## How It Works (Simple + Technical)

MicMix hooks into TeamSpeak's captured-voice callback and edits outgoing voice frames before they are sent.

Technical basics:

- TeamSpeak plugin API `26`.
- Windows audio capture via WASAPI.
- Internal mix target at `48 kHz` mono.
- Ring buffer between capture thread and voice callback.
- Resampling done with `speexdsp` when source sample rate differs.
- Source state handling includes start, stop, running, and automatic reacquire after errors.

## Requirements

For users:

- Windows 10 or Windows 11 (64-bit recommended).
- TeamSpeak 3 Client `3.6.2`.
- TeamSpeak plugin support (API `26` compatible).

Recommended setup:

- Use headphones to avoid feedback/echo.
- Choose the correct source device or app process in MicMix settings.
- Keep source volume moderate and fine-tune with MicMix gain.

## Quick Start (User)

1. Install the `.ts3_plugin` package (double-click it, then confirm in TeamSpeak).
2. Open TeamSpeak and enable `MicMix` in Plugin settings.
3. Open `MicMix Settings...` from the plugin menu.
4. Select exactly one audio source: output device (`Loopback`) or running app (`App Session`).
5. Click `Enable MicMix`.
6. Set music gain and test levels in a channel with a friend.
7. Optional: enable `Force TX` if you want music to continue while you are silent.

## Build From Source (Developer)

Prerequisites:

- Windows + Visual Studio 2022 (C++ toolchain).
- CMake `>= 3.21`.
- TeamSpeak 3 Plugin SDK in `third_party/ts3_sdk`.
- SpeexDSP sources in `third_party/speexdsp`.

Build:

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Release
```

Package:

```powershell
.\scripts\package_ts3_plugin.ps1 -Configuration Release
```

Output package:

- `dist/MicMix-<version>-win64.ts3_plugin`

## Status

MicMix is in active development. Core mixing, source selection, UI controls, hotkey support, and config persistence are already implemented.

See [`DEVPLAN.md`](./DEVPLAN.md) for architecture details, roadmap, and test strategy.
