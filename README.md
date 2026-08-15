# RowMaster

A cross-platform app for browsing, searching, filtering, and editing CSV and Excel files — built with Flutter. Runs natively on Windows, macOS, and Linux, and also as a web app, all from a single shared codebase.

**Try it now, no install:** [ryanveroniwooff.github.io/csv-viewer](https://ryanveroniwooff.github.io/csv-viewer/)

## Features

- **Browse & load** `.csv` or `.xlsx` files via a native file picker
- **Merge** — append additional CSV/Excel files into the currently loaded dataset, validating that headers match before combining
- **Search** across every field at once
- **Filters** — build multiple conditions (`contains`, `=`, `≠`, `>`, `<`) that combine together, with auto-suggested values pulled from the actual data in each column
- **Group By** — summarize the current (filtered) results into unique values and counts for any column
- **Inline cell editing** — click any cell to edit it directly; changes persist through search, filters, group-by, and export
- **Hide/show columns** — declutter the view via a small icon in each column header; hidden columns are excluded from exports too
- **Export as CSV or Excel** — save the current view (filtered rows or grouped summary) back out in either format, via a native save dialog
- Dark navy UI with a custom icon, built for large files (virtualized table rendering so big datasets stay smooth)

## Getting the app

### Web (easiest)

Just open **[ryanveroniwooff.github.io/csv-viewer](https://ryanveroniwooff.github.io/csv-viewer/)** — nothing to install. Runs entirely in your browser; files never leave your machine.

### Desktop (Windows / macOS / Linux)

Prebuilt binaries are built automatically via GitHub Actions on every push to `main`.

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
flutter run -d linux   # or -d windows / -d macos / -d chrome
```

VSCode users: install the **Flutter** extension (publisher: Dart Code), and launch `code .` from inside `nix-shell` so the editor inherits the Flutter SDK on PATH.

## Building manually

```bash
flutter build linux --release     # or windows / macos
flutter build web --release --base-href /csv-viewer/
```

Note: Flutter cannot cross-compile — each platform's build must run on that platform (or via the CI workflow in `.github/workflows/build.yml`, which handles this automatically using GitHub's hosted Windows/macOS/Linux runners plus a web build that auto-deploys to GitHub Pages).

## Tech stack

- [Flutter](https://flutter.dev) (Dart) — UI framework
- [`file_picker`](https://pub.dev/packages/file_picker) — native open/save dialogs
- [`csv`](https://pub.dev/packages/csv) — CSV parsing and encoding
- [`excel`](https://pub.dev/packages/excel) — Excel (.xlsx) parsing and encoding
- [`archive`](https://pub.dev/packages/archive) — used to sanitize malformed real-world xlsx files before parsing