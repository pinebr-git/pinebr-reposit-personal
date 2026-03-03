# Session Save (AutoHotkey v2)

A layout/session manager for Windows built with AutoHotkey v2.
It saves and restores window position, size, state, and basic context so your workspace can be rebuilt quickly.

## Features

- Save and restore current session to JSONL
- Quick Save / Quick Restore
- Manual save / manual restore
- Selective restore by process
- Settings UI for:
  - Shortcut Rules
  - Wait Rules
  - File Extensions
  - Exclusion Rules
- Window info dump
- DPI-safe UI (`-DPIScale`)

## Folder Structure (created on first run)

- `layouts/`
  - `quicksave.jsonl` (latest quick save)
  - `quicksave_1.jsonl`, `quicksave_2.jsonl` (rolling backups)
  - `restore_log.txt` (restore summary log)
- `settings/`
  - `settings.ini` (all settings)
  - `window_dump.txt` (saved from settings menu)

## Requirements

- Windows
- AutoHotkey v2.0
- Script is designed to request and run with administrator rights automatically

## How to Run

1. Run `session save.ahk`
2. Allow UAC prompt if requested
3. The main menu UI will open

## Main Menu

### 1) Quick Restore
- Restores from `layouts\quicksave.jsonl` if it exists.

### 2) Quick Save
- Saves current layout to `quicksave.jsonl`.
- Automatically keeps previous quick saves in `quicksave_1.jsonl` and `quicksave_2.jsonl`.

### 3) Manual Restore
- Choose a saved layout file from the list and restore it.
- Rename/Delete are available in the file list.

### 4) Manual Save
- Save current windows to a custom filename.
- Rename/Delete existing files from the same dialog.

### 5) Selective Restore
- Choose which process entries to restore from a saved file and restore only those.

### 6) Settings

#### Shortcut Rules
- Map `(executable + title keyword)` to a target path used during restore.
- Target can be a shortcut, document, or executable path.
- This is useful for apps that require special launch context.
- for example, you can map word(executable) and 'doc1234'(title keyword) to "C:\User\document\doc1234.docx"(target path), so that doc1234 can be re-opened when your save includes word process opening doc1234. 
- You can also match PWA ink file to the browser-title pair.

#### Wait Rules
- `Short Wait EXEs` and `Long Wait EXEs` define startup wait profile.
- `Timeout Overrides` allow per-executable timeout in ms.

#### File Extensions
- Configure extensions used to detect file paths from window titles.

#### Exclusion Rules
- Add executable names to skip when saving current layout.

#### Dump Windows Info
- Writes current window states to `settings/window_dump.txt`.

## Data Format

- Layout files: JSONL (`.jsonl`)
- Settings: INI (`settings/settings.ini`)

## Notes

- Some applications may not expose file paths or may restore slowly depending on app behavior.
- Accuracy strongly depends on proper wait/exclusion/shortcut rule tuning for your environment.
- Websites on your browser cannot be read by this script, there fore they cannot be restored. this script will only open empty browser, so saving browser session should be done in the browser.
- Usually does not save sessions in other desktops.

## Packaging / Distribution

- This project is single-file script based.
- If packaging with `Ahk2Exe`, keep `session save.ahk` as the entry file.
- Ensure write permission for generated folders/files on first run.