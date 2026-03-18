# Changelog

All notable changes to this project are documented in this file.

## [1.2.9] - 2026-03-18

### Changed

- Added `VERSIONING.md` as the canonical SemVer policy for release bump and reset rules.
- Updated the README screenshot asset.

### Fixed

- Fixed a regression where enabling mic-input mute could unintentionally stop expected music pass-through.

## [1.2.8] - 2026-03-16

### Added

- Added a dedicated microphone-input mute hotkey in the settings UI, including persisted binding support.

### Changed

- Moved MicMix mute control fully to internal/global hotkey handling and removed TeamSpeak plugin hotkey registration paths.
- Hardened CI/release workflow quality gates by running Release tests during both CI and tag-release pipelines.
- Updated pinned GitHub Action references and selected README sections.

### Fixed

- Fixed cases where mic-input mute did not reliably mute transmitted microphone audio, including transport-level mute fallback and safe state restore.
- Fixed false continuous-transmit behavior and improved disconnect/shutdown handling around forced TX and talk-mode restoration.
- Fixed race conditions and duplicate-toggle behavior in music/mic mute hotkey handling.
- Fixed audio format fallback handling for app-session capture to prevent wrong sample interpretation after fallback.
- Hardened release packaging path handling and config loading size enforcement to prevent unsafe file/path processing.

## [1.2.7] - 2026-03-14

### Added

- Added calibrated dB scale markers to the main music and microphone meters.
- Added runtime `tag.micmix_active=<0|1>` visibility in the bottom status log.

### Changed

- Improved meter responsiveness tuning while keeping UI overhead low.
- Refined settings UI status visibility and interaction polish, including slider thumb-only resize cursor behavior and gain slider alignment.
- Reduced UI flicker by introducing buffered rendering and repaint-on-change behavior for meter/clip visuals and status log rendering.
- Updated README wording in selected sections for clarity.

### Fixed

- Hardened hotkey handling against local thread-queue message injection attempts.
- Hardened plugin input validation and config parsing/loading behavior, with additional regression test coverage.

## [1.2.6] - 2026-03-13

### Added

- Added clip telemetry event counters and dedicated clip strip visuals in the settings UI for both source and microphone paths.
- Added source entry icons with fallback behavior to improve source selection clarity.

### Changed

- Polished settings window layout, spacing, control styling, and interaction feedback (including cursor behavior for interactive controls).
- Improved running-app source selection heuristics to prioritize more relevant app sessions.
- Improved release metadata/version handling by deriving build version from tag workflow and keeping local version files in sync.

### Fixed

- Fixed plugin ID lifetime handling and settings window refresh/save flow regressions that could cause instability.
- Hardened TeamSpeak lifecycle/shutdown paths to reduce teardown race risks (capture/meta/logging related paths).
- Fixed a potential source-combo text buffer overflow by switching to length-aware dynamic retrieval.
- Fixed a potential deadlock path when source thread startup fails.
- Fixed a potential hotkey thread shutdown hang by using explicit stop signaling and non-blocking message wait behavior.
- Fixed transient talk-gate dropouts on live settings apply by preserving runtime talk state.
- Fixed muted-buffer stale audio backlog by clearing/skipping queued music while muted.

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

