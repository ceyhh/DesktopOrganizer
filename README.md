# Desktop Organizer (Windows GUI + Release EXE)

Desktop Organizer is a Windows file organization tool that groups files into extension-based folders.

## Release

[Go to Release v1.0.0](https://github.com/ceyhh/DesktopOrganizer/releases/tag/v1.0.0)

[Download DesktopOrganizer-1.0.0.exe](https://github.com/ceyhh/DesktopOrganizer/releases/download/v1.0.0/DesktopOrganizer-1.0.0.exe)

## What This Project Includes

- GUI application with a `Start` button, progress bar, and completion message.
- Manual folder picker so users can organize any folder.
- Automatic fallback to Desktop if no folder is selected.
- Folder naming template using `@` placeholder for extension names.
- Ignore-extension input (with reminder to include dot format like `.tmp`).
- Special handling for unknown files:
    - `.lnk` and `.url` are moved to `unknowns`.
- Optional checkbox support:
    - `Put files without extension into unknowns (default: off)`
    - If enabled, files with no extension are also moved to `unknowns`.
- Duplicate filename collision handling (`_1`, `_2`, ...).
- Robust Desktop detection with multiple methods:
    - Windows Shell API
    - Windows Registry
    - Environment variables
    - Cloud folder roots (OneDrive, Dropbox, Google Drive, etc.)
    - 100+ localized Desktop folder names

## GUI Fields and Behavior

1. Folder Path
- You can choose a folder manually with `Browse`.
- If left empty, the app automatically uses Desktop.

2. Folder Name Template
- Use `@` where the extension name should be inserted.
- Example: `files_@`
- `.txt` files go to `files_txt`
- `.png` files go to `files_png`

3. Ignore Extensions
- Enter extensions as comma-separated values.
- You should include the dot (`.`), for example: `.tmp,.log,.bak`
- If you forget the dot, the app auto-normalizes and informs you.

4. Start + Progress
- Click `Start` to begin.
- Progress bar updates while processing.
- At the end, the app shows `Completed` with summary counts.

5. No-extension to unknowns (checkbox)
- On the start screen, you can enable `Put files without extension into unknowns`.
- Default is off.

## Build Release EXE

### Requirements

- Windows 10/11
- Python 3.10+
- Internet access for first-time dependency install

### Build Steps

1. Double-click `build_exe.bat`.
2. The script installs/updates build tools and creates a one-file executable.
3. Output files:
- `dist/DesktopOrganizer.exe`
- `release/DesktopOrganizer.exe`
- `release/README.txt`

## Run

1. Double-click `release/DesktopOrganizer.exe`.
2. Select folder (or leave empty for Desktop).
3. Enter template with `@`.
4. Enter ignored extensions.
5. Click `Start`.

## Safety Notes

- The tool only moves files, not folders.
- It does not modify file content.
- It keeps running even if some files cannot be moved.

## Files in This Repository

- `desktop_extension_organizer.py`: Main application logic and GUI.
- `build_exe.bat`: Release build script.
- `README.md`: Documentation.

---

Small note: this project was built with AI assistance.
