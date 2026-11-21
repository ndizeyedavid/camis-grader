# Changelog

All notable changes to CAMIS Marks are listed here.
Adding a new version heading on `main` publishes a Windows release automatically.

## [2.0.0] - 2026-11-21

### Added
- New desktop app with a three-step flow: choose the file, open CAMIS, fill marks.
- Review screen before typing, with assessment column choice and mark conversion preview.
- Optional replacement of marks that already differ in CAMIS.
- Native Windows notifications and taskbar flashing when the page is found, marks are done, or something needs attention.
- Recent Excel files shortcut and drag and drop.
- Reports saved to `Documents\CAMIS Reports`.
- Windows installer built by GitHub Actions.

### Changed
- Engine planning logic shared between the app and the command line tool.

## [1.0.0] - 2026-10-21

### Added
- Command line tool that reads a CAMIS / Urubuto export and types the marks into CAMIS.
