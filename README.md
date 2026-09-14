# OneNote Replace

OneNote Replace is a small, source-auditable Windows utility for finding and replacing visible text on the current Microsoft OneNote desktop page.

## Features

- Runs locally through the OneNote desktop COM API.
- Searches the currently viewed page only.
- Includes visible hyperlink names while preserving hyperlink destinations.
- Supports case-sensitive and whole-word matching.
- Shows a dry-run preview before changing the page.
- Requires explicit confirmation before writing.
- Detects page changes between analysis and update.
- Creates a local XML backup before applying replacements.
- Contains no network or telemetry code.

## Formatting safety

OneNote Replace preserves the original rich-text markup and submits only page-level outlines containing actual matches. It does not send the complete page back to OneNote. This avoids a OneNote behavior that can reapply stale embedded QuickStyle definitions and change manually formatted fonts, such as reverting Century Schoolbook to Calibri.

The utility changes visible text between markup tags without rewriting surrounding font, size, emphasis, color, or hyperlink attributes. Hyperlink display names can change, but their destination addresses remain unchanged.

Use OneNote's white page canvas while applying replacements. OneNote dark mode normally changes only the display, but programmatic updates made while the canvas is dark can normalize automatic text colors unexpectedly. You may switch back to the dark canvas after the update.

## Requirements

- Windows 10 or Windows 11.
- Microsoft OneNote desktop.
- Windows PowerShell 5.1.

OneNote for the web, macOS, iOS, and Android are not supported.

## Run from source

1. Download `src/OneNoteReplace.ps1`.
2. Open desktop OneNote and select the page to edit.
3. Close additional OneNote windows.
4. Use **View > Switch Background** if necessary so the page canvas is white.
5. Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\OneNoteReplace.ps1
```

6. Select **Preview (dry run)**.
7. Review the match count, changed text blocks, and page-level outlines to update.
8. Select **Apply replacement** only if the preview is correct.

Test with a duplicated page before using the utility on important notes.

## Backups

Backups are written to:

```text
%LOCALAPPDATA%\OneNoteLocalFindReplace\Backups
```

A backup contains the page XML and may contain confidential information. Protect and delete backups according to your own retention requirements.

## Build an executable

The optional build script uses the community `ps2exe` PowerShell module:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

The resulting executable is placed in `dist`. PS2EXE packages the script; it does not make the source secret or create a security boundary. Unsigned builds may trigger Microsoft Defender or SmartScreen warnings.

## Current limitations

- Current page only.
- Use a white page canvas when applying replacements.
- Page titles, attachments, and OCR metadata are excluded.
- A phrase split across separate formatting nodes may not match.
- Text fragments that cannot be parsed safely are skipped and reported.
- OneNote updates complete page-level outlines containing matches; unrelated outlines are not submitted.
- The utility depends on the locally installed OneNote desktop COM interface.

## Privacy

See [`docs/PRIVACY.md`](docs/PRIVACY.md). Review the source before using it with confidential material.

## License

MIT License. See [`LICENSE`](LICENSE).
