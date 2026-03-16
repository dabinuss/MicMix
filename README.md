# MicMix

## Important Disclaimer

**This project was developed with help from Codex, an AI coding agent by OpenAI.**
AI-assisted development can introduce mistakes, edge-case bugs, or unintended behavior. Use MicMix with caution, especially in production-like or critical setups.

The full source code is public. **Please review the code**, verify behavior, and report bugs, regressions, or security concerns.

![MicMix Logo](assets/branding/MicMix%20Logo.png)

MicMix is a Windows plugin for TeamSpeak 3 that mixes a second audio source (music, browser audio, media player audio, etc.) into your microphone stream.

It is built for people who want to speak and play audio in the same voice channel without setting up complex virtual cable chains.

## Screenshot

<img width="682" height="723" alt="grafik" src="https://github.com/user-attachments/assets/cc130399-1bfd-406e-801d-6bd609cdf14c" />

## What MicMix Can Do

- Mix your mic + one selected audio source into TeamSpeak voice transmission.
- Capture audio either from `Loopback` (system playback device) or `App Session` (a specific running app process, for example Spotify, browser, VLC).
- Control music level with a gain slider.
- Mute/unmute music quickly with a hotkey.
- Mute/unmute mic input separately (button + hotkey).
- Keep sending music even while you are not talking (`Force TX`).
- Show live level meters for music and microphone.
- Show clip strips and clip event counters for music and mic paths.
- Monitor the mixed output locally (`Monitor Mix`) while connected.
- Auto-start source capture on plugin startup (optional).
- Reacquire the selected source automatically after runtime errors.
- Save settings and restore them on next launch.
- Publish runtime activity tag `micmix_active=<0|1>` into TeamSpeak client metadata.

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

## Server Operator Tag (`micmix_active`)

MicMix writes a runtime marker into TeamSpeak client metadata:

- `micmix_active=1`: MicMix currently detects active music signal being sent.
- `micmix_active=0`: no current active music signal.

Implementation details:

- The value is written to `CLIENT_META_DATA` and merged into existing metadata as a `;`-separated key-value pair.
- Existing metadata tokens are kept, only the `micmix_active` token is updated/replaced.
- Updates are throttled to avoid metadata spam, so there can be short delay on state transitions.

How server operators can use it:

1. Read `client_meta_data` via ServerQuery or a bot.
2. Parse for token `micmix_active=1`.
3. Use it as a signal for UI/badges, logging, moderation hints, or music-bot/channel automation.

Example logic (pseudo):

```text
if "micmix_active=1" in client_meta_data:
    mark_client_as_music_source()
else:
    clear_music_source_marker()
```

Important: this tag is client-published metadata and should be treated as a helpful signal, not a hard security control.

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
