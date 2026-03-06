# Changelog

All notable changes to this project are documented in this file.

## [1.2.1] - 2026-03-06

### Changed

- Refined settings UI structure and labels for clarity (`MICMIX SETTINGS`, `AUDIO ROUTING`, `MIX BEHAVIOR`).
- Moved and rebalanced controls in the settings window with consistent spacing and cleaner meter/status alignment.
- Adjusted control styling toward a more native TeamSpeak-like feel (neutral themed controls, improved slider behavior, safer meter text rendering).

### Fixed

- Fixed regression where already-running music sometimes started sending only after first microphone activity.
- Improved force-TX transport stability by keeping the send path reliably active while MicMix is engaged.
- Reduced periodic TS reapply pressure and softened capture watchdog behavior to lower risk of intermittent robotic voice artifacts.
- Fixed meter/status text overflow outside card bounds in the settings UI.

### Removed

- Removed unused config/UI leftovers introduced during prior iterations (dead helper wrappers and stale control state fields).

## [1.2.0] - 2026-03-05

### Added

- Local mix monitoring playback so users can monitor the mixed signal while connected.
- Improved microphone telemetry path and mic level meter integration in the settings UI.
- Additional UX controls around live monitoring and source behavior.

### Changed

- Refined settings window behavior and presentation for source selection, level feedback, and hotkey handling.
- Expanded and professionalized README documentation with clearer user guidance, requirements, and use cases.

### Fixed

- Stronger exception safety around plugin lifecycle and callback entry points.
- Improved error handling and stability in audio/core initialization and runtime paths.

