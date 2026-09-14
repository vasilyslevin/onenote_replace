## Privacy and security

OneNote Replace is designed to operate locally through the Microsoft OneNote desktop COM API. The source contains no networking, analytics, telemetry, advertising, or update-checking code.

### Information processed

The utility retrieves the XML representation of the currently viewed OneNote page, searches eligible visible text nodes, and returns modified XML to the locally installed OneNote desktop application.

It includes visible hyperlink names in replacements but does not modify hyperlink destination attributes. Page titles, attachments, and OCR metadata are excluded.

### Local backups

Before applying a replacement, the utility writes the original page XML to:

```text
%LOCALAPPDATA%\OneNoteLocalFindReplace\Backups
```

These files may contain the complete textual contents and metadata of the affected page. Anyone with access to the Windows account or backup folder may be able to read them. Users are responsible for securing and deleting backups according to their confidentiality and retention requirements.

### OneNote storage

This utility does not change where OneNote stores notebooks. A notebook stored in OneDrive, SharePoint, or another synchronized location remains subject to that service's storage, synchronization, access-control, and retention behavior.

### Executable builds

The PowerShell source is authoritative. Executable releases may be packaged with PS2EXE for convenience. Packaging does not make the source secret or provide a sandbox. Users should verify release hashes and may run the reviewed PowerShell source directly instead.

### Sensitive information

Test on a duplicated page. Confirm organizational policy before using third-party utilities with privileged, regulated, employment-related, medical, financial, or other confidential information.
