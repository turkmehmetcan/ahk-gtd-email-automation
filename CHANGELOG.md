# Changelog

All notable changes to this project will be documented in this file.


## Documentation

- Create comprehensive README for v1.0.0 release


- Add professional project overview with badges

- Document all hotkeys for Outlook and Gmail automation

- Include installation instructions with AutoHotkey v2.0 requirement

- Add hotkey reference table for quick comparison

- Provide first-time setup instructions for both platforms

- Include configuration guide for Outlook config.ini

- Add customization examples for labels/folders and hotkeys

- Document browser compatibility (25+ browsers for Gmail)

- Add author info, acknowledgments, and support links

- Prepare for first public release




## Features

- Add Outlook GTD automation with Quick Step hotkeys


- Implemented AutoHotkey v2 script to streamline GTD workflow in Outlook

- Mapped intuitive Alt-key combinations to existing Quick Steps in Outlook

- Alt+E: Archive emails

- Alt+A: Move to `@Action` folder

- Alt+W: Move to `@Waiting For` folder

- Alt+R: Move to `@Reference` folder

- Alt+Z: Move back to Inbox

- Hotkeys only active when Outlook window is focused for safety



- Add automatic task creation and external config support


- Added CreateTaskInPrimaryMail() function to create Outlook tasks from selected emails

- Enhanced Alt+A and Alt+W hotkeys to trigger Quick Steps AND create tasks automatically

- Added Alt+T hotkey to create tasks without moving emails (direct task creation)

- Implemented config.ini support for user-specific email configuration

- Tasks are created in primary work account specified in config file

- Tasks include original email as attachment and preserve subject/body

- Added option to auto-open task windows (OPEN_TASK setting)

- Script validates config file exists and shows helpful setup message if missing



- Replace Quick Steps with native GTD workflow and folder management


- Removed dependency on Outlook Quick Steps - now manages folders/categories directly via COM API

- Implemented ProcessGTDAction() to move emails, apply categories, and optionally create tasks

- Added ProcessArchive() and ProcessInbox() for complete GTD email workflow

- Implemented Alt+0 hotkey with SetupGTDFolders() to auto-create all required folders and categories

- Added intelligent folder search (FindOrCreateFolder, FindExistingFolder, FindFolderRecursive)

- Implemented ApplyCategory() with duplicate detection for proper category management

- Added SetupCategory() to create/update Outlook categories with custom colors

- Configured color-coded categories: Action (Red), Waiting For (Yellow), Reference (Blue)

- Enhanced Alt+A and Alt+W to move emails, mark unread, apply categories, and create tasks

- Alt+R now moves to `@Reference` folder without task creation

- Alt+E archives emails and marks as read

- Alt+Z returns emails to Inbox, clears categories, and marks unread

- All functionality now works independently of Quick Steps configuration

- Preserved external config.ini support for PRIMARY_SMTP email address



- Add Gmail GTD automation using keyboard shortcuts


- Implemented AutoHotkey v2 script for GTD workflow in Gmail web interface

- Configured customizable Gmail labels for GTD buckets (Archive, Action, Waiting, Reference)

- Mapped intuitive Alt-key combinations to Gmail keyboard shortcuts:

  - Alt+Enter: Archive email (mark as read)

  - Alt+A: Move to `@Action` (mark as unread)

  - Alt+W: Move to `@Waiting For` (mark as unread)

  - Alt+R: Move to `@Reference` (mark as unread)

  - Alt+E: Quick archive

  - Alt+Z: Return to Inbox

  - Alt+U: Mark as unread

  - Alt+I: Mark as read

- Added ProcessGTDBucket() to apply labels, archive from inbox, and control read status

- Implemented browser detection supporting 25+ browsers (Chrome, Firefox, Edge, Brave, etc.)

- IsGmailActive() validates both browser process and Gmail page title

- Uses Gmail native keyboard shortcuts (l, e, v, gi, +u, +i) for reliable operation

- Configurable delays for UI synchronization (DELAY_SHORT, DELAY_LONG)

- Hotkeys only active when Gmail is open in a supported browser



- Enhance Gmail GTD with spam handling, delete function, and improved timing


- Renamed Gmail shortcut constants from GM_* to GMS_* for clarity (Gmail Shortcuts)

- Increased DELAY_LONG from 150ms to 250ms for better UI synchronization reliability

- Added LABEL_GMAIL_SPAM constant for spam folder operations

- Implemented DeleteMail() function using Shift+3 Gmail shortcut (Alt+Delete)

- Implemented MarkSpam() function with multi-step spam marking workflow (Alt+End)

- Added RefreshInbox() function to refresh inbox view without page reload (Alt+Space)

- Renamed GmailArchive() to MoveToArchive() for naming consistency

- Improved MoveToInbox() to mark unread before moving for better GTD workflow

- Simplified ProcessGTDBucket() timing by removing redundant inbox refresh

- Enhanced MoveToArchive() as dedicated function replacing inline Send calls

- Added GMS_TAB constant for Tab key navigation in dialogs

- Replaced unused constants (GM_SELECT_ALL, GM_DELETE, GM_ESCAPE) with functional ones

- All hotkeys now have descriptive inline comments with Alt+key notation




## Miscellaneous

- Add git-cliff configuration for changelog generation


- Merge development into main for v1.0.0 release


This is the first public release featuring:

- Outlook GTD automation with COM API integration

- Gmail GTD automation with keyboard shortcuts for 25+ browsers

- External configuration support for Outlook (config.ini)

- Comprehensive README with installation and usage guide

- Git-cliff configuration for automated changelog generation




## Refactoring

- Improve code consistency and setup feedback with better diagnostics


- Renamed functions to camelCase for consistency (ProcessGTDAction -> MoveToGtdBucket, etc.)

- Added Outlook constants section (OL_CLASS_MAIL, OL_FOLDER_INBOX, etc.) for better code readability

- Enhanced CreateGtdElements() with detailed TrayTip notifications for setup progress tracking

- Implemented setupCategoryForStore() nested function for per-account category management

- Added reportChange parameter to findOrCreateFolder() to track what was created vs existing

- Improved setup completion messages to distinguish between new setup and no changes

- Removed recursive folder search in findExistingFolder() to avoid performance issues

- Added debug TrayTips in setupCategory() for troubleshooting category creation

- Enhanced error messages with more context (root folder name, account display name)

- Setup now reports success only when changes are made, or confirms nothing needed updating

- Preserved external config.ini support for PRIMARY_SMTP email address



- Major Gmail GTD rewrite with comprehensive documentation and modular architecture


- Added comprehensive script header with description, features, and usage guide

- Implemented #Requires directive for AutoHotkey v2.0 compatibility enforcement

- Added #SingleInstance Force and SendMode(Input) for reliability

- Restructured timing configuration with three-tier delay system:

  - DELAY_SHORT (200ms) for quick operations

  - DELAY_MEDIUM (400ms) for standard UI operations

  - DELAY_LONG (1600ms) for complex multi-step operations

- Renamed all Gmail shortcut constants from GMS_* to KEY_GMAIL_* for clarity

- Renamed label constants to LABEL_GTD_* and LABEL_GMAIL_* namespaces

- Renamed ProcessGTDBucket() to MoveToGtdBucket() for consistency

- Added KEY_GMAIL_NEWER_CONVERSATION constant for navigation (k key)

- Implemented helper function architecture for code reuse:

  - WaitDelay() - Centralized delay management

  - SendShortcut() - Send keys with configurable delays

  - SendText() - Send text with configurable delays

  - PressEnter() - Standardized Enter key handling

- Enhanced MarkSpam() to automatically move to next conversation after spam marking

- Improved code documentation with detailed function comments

- All core functions now use helper functions for consistency and maintainability

- Grouped constants by category (Actions, Status, Navigation, UI Elements)


