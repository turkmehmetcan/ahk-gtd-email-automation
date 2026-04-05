# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.1.0] - 2026-04-06

### Added

**Navigation:**

- New `Alt+Up` and `Alt+Down` hotkeys navigate to newer/older emails in both Gmail and Outlook without leaving the current view.

**Gmail:**

- Archiving now detects whether you are in list view or detailed view and uses the correct shortcut for each, preventing accidental returns to the inbox root.
- Email read/unread state is no longer altered when moving messages to GTD buckets; state is fully preserved.

**Configuration:**

- `config.ini` is auto-generated with defaults on first run; `config.example.ini` has been removed.
- `PrimaryEmail` setting now accepts `Off` to create tasks in each account's own ToDo list, or an email address to funnel all tasks into one account.
- New `RunAtStartup` setting controls whether the script registers itself in the Windows Startup folder automatically.

### Fixed

**Outlook:**

- GTD folder and category setup now runs reliably; a scoping bug previously caused setup to silently fail.
- Setup summary now reports accounts, categories, and folders created along with any errors encountered.

---

## [2.0.0] - 2025-10-07

### Breaking Changes

**The standalone `Gmail_GTD.ahk` and `Outlook_GTD.ahk` scripts have been removed.**

You must now use the unified `GTD_Automation.ahk` script for both platforms.

**What you need to do:**

1. Delete old shortcuts to `Gmail_GTD.ahk` and `Outlook_GTD.ahk`
2. Create a new shortcut to `GTD_Automation.ahk`
3. Use the same hotkeys—they work identically in both Gmail and Outlook

### Added

- **Unified script** - Run `GTD_Automation.ahk` for both Gmail and Outlook. The script automatically detects which email client is active and applies the correct automation.

- **Automatic platform detection** - The script switches between Gmail and Outlook modes based on your active window. No manual configuration needed.

- **Modular architecture** - Code reorganized into separate library modules (`lib/GmailGTD.ahk` and `lib/OutlookGTD.ahk`) for easier maintenance and future extensibility.

### What Stayed the Same

- All hotkeys (Alt+A, Alt+W, Alt+R, etc.) work exactly as before
- All features and functionality remain unchanged
- Configuration files (`config.ini` for Outlook) work the same way

---

## [1.0.0] - 2025-10-02

### Initial Release

Automate your GTD email workflow with keyboard shortcuts for both Outlook and Gmail.

### What You Get

**Two standalone scripts:**

- `Gmail_GTD.ahk` - For Gmail in any browser
- `Outlook_GTD.ahk` - For Outlook desktop

### Core Features

**Keyboard shortcuts for GTD workflow:**

- `Alt+A` - Move to `@Action` (creates Outlook task automatically)
- `Alt+W` - Move to `@Waiting For` (creates Outlook task automatically)
- `Alt+R` - Move to `@Reference`
- `Alt+E` - Archive email
- `Alt+Z` - Return to Inbox

**Outlook-specific:**

- `Alt+0` - First-time setup (creates folders and color-coded categories)
- Automatic task creation with email attachments
- Works with multiple email accounts via `config.ini`

**Gmail-specific:**

- `Alt+Delete` - Delete email
- `Alt+End` - Mark as spam
- `Alt+Space` - Refresh inbox
- Works in 25+ browsers (Chrome, Firefox, Edge, Brave, etc.)

### Setup

1. Install [AutoHotkey v2.0](https://www.autohotkey.com/v2/)
2. Run `Gmail_GTD.ahk` or `Outlook_GTD.ahk` (or both)
3. For Outlook: Create `config.ini` with your email address
4. For Gmail: Enable keyboard shortcuts in Gmail settings

---
