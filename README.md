# CSV Viewer

A cross-platform desktop app for browsing, searching, filtering, and summarizing CSV files — built with Flutter. Runs natively on Windows, macOS, and Linux with a single shared codebase.

## Features

- **Browse & load** any `.csv` file via a native file picker
- **Search** across every field at once
- **Filters** — build multiple conditions (`contains`, `=`, `≠`, `>`, `<`) that combine together, with auto-suggested values pulled from the actual data in each column
- **Group By** — summarize the current (filtered) results into unique values and counts for any column
- **Export** — save the current view back out to CSV, whether that's the filtered rows or the grouped summary, via a native save dialog
- Dark-themed UI, built for large files (virtualized table rendering so big CSVs stay smooth)

## Getting the app

Prebuilt binaries for Windows, macOS, and Linux are built automatically via GitHub Actions on every push to `main`.

1. Go to the **Actions** tab of this repo
2. Open the latest successful workflow run
3. Scroll to **Artifacts** and download the zip for your platform:
   - `windows-build` → unzip, run the `.exe` inside
   - `macos-build` → unzip, run the `.app`
   - `linux-build` → unzip, run the executable inside the `bundle/` folder

> **Note:** these builds are unsigned. Windows may show a SmartScreen warning ("Windows protected your PC") and macOS may block the app on first launch ("Apple cannot check it for malicious software"). This is expected for an internal tool without a code-signing certificate — right-click → Open on macOS, or click "More info" → "Run anyway" on Windows, to proceed.

## Development setup

This project uses a `shell.nix` for a reproducible dev environment (NixOS-friendly, but works anywhere Nix is installed).

```bash
nix-shell
flutter pub get
flutter run -d linux   # or -d windows / -d macos, depending on your platform
```

VSCode users: install the **Flutter** extension (publisher: Dart Code), and launch `code .` from inside `nix-shell` so the editor inherits the Flutter SDK on PATH.

## Building manually

```bash
flutter build linux --release     # or windows / macos
```

Note: Flutter cannot cross-compile — each platform's build must run on that platform (or via the CI workflow in `.github/workflows/build.yml`, which handles this automatically using GitHub's hosted Windows/macOS/Linux runners).

## Tech stack

- [Flutter](https://flutter.dev) (Dart) — UI framework
- [`file_picker`](https://pub.dev/packages/file_picker) — native open/save dialogs
- [`csv`](https://pub.dev/packages/csv) — CSV parsing and encoding