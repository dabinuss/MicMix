# Changelog

All notable changes to this project are documented in this file.

## [1.2.5] - 2026-03-12

### Fixed

- Hotfix: prevented `CLIENT_META_DATA` updates during plugin shutdown to reduce TeamSpeak client teardown crashes caused by late metadata sync calls.

## [1.2.4] - 2026-03-12

### Added

- Added application icons for audio source entries with robust fallback behavior when icons are unavailable.
- Added MicMix activity publication via TeamSpeak client metadata with anti-spam throttling and synchronization guards.

### Changed

- Refined settings UI layout for better consistency, compactness, and safer control bounds.
- Improved app-session source selection to prefer processes with active audio sessions.
- Improved music meter behavior with a telemetry-based fallback so signal activity remains visible in more scenarios.

### Fixed

- Hardened plugin lifecycle and shutdown paths to reduce close-time crash risk.
- Improved thread safety around monitor enable/disable transitions and UI shutdown sequencing.
- Hardened core thread-safety and shutdown race handling in MicMix runtime paths.
- Improved shutdown cleanup and refresh-completion handling robustness.
- Improved safety/logging when posting source-refresh completion messages fails.
- Ensured `VERSION` updates trigger CMake reconfigure so generated version metadata stays in sync.

## [1.2.3] - 2026-03-08

### Added

- Added configurable resampler quality in settings with related runtime integration.
- Added inline UI hints for gain and force-transmission controls to improve discoverability.

### Changed

- Improved audio robustness and automatic resampler policy behavior for varying source rates.
- Refined settings UX and validation/sanitization paths for more predictable configuration handling.
- Streamlined README wording in selected sections for clearer guidance.

### Fixed

- Fixed app session ID candidate parsing to avoid invalid entries in edge cases.
- Improved directory creation error handling in configuration paths.

## [1.2.2] - 2026-03-06

### Changed

- Refined settings UI structure and labels for clarity (`MICMIX SETTINGS`, `AUDIO ROUTING`, `MIX BEHAVIOR`).
- Moved and rebalanced controls in the settings window with consistent spacing and cleaner meter/status alignment.
- Adjusted control styling toward a more native TeamSpeak-like feel (neutral themed controls, improved slider behavior, safer meter text rendering).

### Fixed

- Fixed regression where already-running music sometimes started sending only after first microphone activity.
- Improved force-TX transport stability by keeping the send path reliably active while MicMix is engaged.
- Stabilized force-send behavior for already-running music streams so transport activation no longer depends on a first mic trigger.
- Reduced periodic TS reapply pressure and softened capture watchdog behavior/timing to lower risk of intermittent robotic voice artifacts in long sessions.
- Fixed meter/status text overflow outside card bounds in the settings UI.
- Fixed release automation by fetching required third-party dependencies in CI/release workflows before CMake configure.

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

