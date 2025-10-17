# GTD Email Automation Scripts

[![AutoHotkey](https://img.shields.io/badge/Language-AutoHotkey_v2.0-green.svg)](https://www.autohotkey.com/)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Automate your GTD (Getting Things Done) email workflow for **Outlook** and **Gmail** using keyboard shortcuts.

## Features

### Unified GTD Automation (`GTD_Automation.ahk`)

**New in v2.0:** Single entry point for both platforms with automatic detection

- **One script** for both Outlook and Gmail
- **Automatic platform detection** - hotkeys work based on active window
- **Modular architecture** - clean separation of platform logic
- **Consistent shortcuts** - same Alt+key combinations across platforms

### Outlook GTD Features (via `lib/OutlookGTD.ahk`)

- **Alt+A**: Move email to `@ACTION` folder + create task
- **Alt+W**: Move email to `@WAITING FOR` folder + create task
- **Alt+R**: Move email to `@REFERENCE` folder
- **Alt+E**: Archive email (mark as read)
- **Alt+Z**: Move email back to Inbox (mark as unread)
- **Alt+0**: Create GTD folders, categories, and colors in all accounts

### Gmail GTD Features (via `lib/GmailGTD.ahk`)

- **Alt+Enter**: Move email to `[0] @GTD ARCHIVE` (mark as read)
- **Alt+A**: Move email to `[1] @ACTION` label (mark as unread)
- **Alt+W**: Move email to `[2] @WAITING FOR` label (mark as unread)
- **Alt+R**: Move email to `[3] @REFERENCE` label (mark as unread)
- **Alt+E**: Archive email
- **Alt+Z**: Move email back to Inbox
- **Alt+Delete**: Delete email
- **Alt+End**: Mark as spam
- **Alt+I**: Mark as read
- **Alt+U**: Mark as unread
- **Alt+Space**: Refresh inbox

Works with 25+ browsers including Chrome, Edge, Firefox, Brave, Opera, and more.

## Requirements

- [AutoHotkey v2.0](https://www.autohotkey.com/) or later
- Windows OS
- **Outlook**: Desktop application with COM API access
- **Gmail**: Web interface with keyboard shortcuts enabled

## Installation

1. **Install AutoHotkey v2.0**
   - Download from [autohotkey.com](https://www.autohotkey.com/)

2. **Clone or download this repository**

   ```bash
   git clone https://github.com/turkmehmetcan/ahk-gtd-email-automation.git
   ```

3. **Configure Outlook (if using)**
   - Copy `config.example.ini` to `config.ini`
   - Edit `config.ini` and set your primary email address:

     ```ini
     [Settings]
     PrimaryEmail=your.email@domain.com
     ```

   - Or set `PrimaryEmail=Off` to create tasks in each email's own account instead of a single primary account

4. **Enable Gmail keyboard shortcuts (if using Gmail)**
   - Open Gmail → Settings (gear icon) → See all settings
   - Go to "General" tab → Keyboard shortcuts → Enable
   - Save changes

## Usage

### Running the Script

Simply run `GTD_Automation.ahk` - it automatically works with both Outlook and Gmail:

```bash
# Double-click or run from command line
GTD_Automation.ahk
```

The script detects which application is active and applies the appropriate hotkeys automatically.

### First-Time Setup

**Outlook Users:**

1. Run `GTD_Automation.ahk`
2. Open Outlook
3. Press **Alt+0** to automatically create:
   - GTD folders (@ACTION, @WAITING FOR, @REFERENCE, Archive)
   - GTD categories with colors
   - Works across all your Outlook accounts

**Gmail Users:**

1. Run `GTD_Automation.ahk`
2. Open Gmail in your browser
3. Manually create labels (the script will apply them automatically):
   - `[0] @GTD ARCHIVE`
   - `[1] @ACTION`
   - `[2] @WAITING FOR`
   - `[3] @REFERENCE`

### Hotkey Reference

| Hotkey | Outlook Action | Gmail Action |
|--------|----------------|--------------|
| **Alt+A** | Move to @ACTION + Create Task | Move to [1] @ACTION (unread) |
| **Alt+W** | Move to @WAITING FOR + Create Task | Move to [2] @WAITING FOR (unread) |
| **Alt+R** | Move to @REFERENCE | Move to [3] @REFERENCE (unread) |
| **Alt+E** | Archive (mark read) | Archive |
| **Alt+Z** | Back to Inbox (unread) | Back to Inbox |
| **Alt+0** | Setup GTD elements | *(not applicable)* |
| **Alt+Enter** | *(not applicable)* | Move to [0] @GTD ARCHIVE (read) |
| **Alt+Delete** | *(not applicable)* | Delete Email |
| **Alt+End** | *(not applicable)* | Mark as Spam |
| **Alt+I** | *(not applicable)* | Mark as Read |
| **Alt+U** | *(not applicable)* | Mark as Unread |
| **Alt+Space** | *(not applicable)* | Refresh Inbox |

## Quick Setup

1. Copy `config.example.ini` to `config.ini`.
2. Set your preferred mode:
   - `PrimaryEmail=Off` — tasks go to each account's own list.
   - `PrimaryEmail=your.email@domain.com` — all tasks go to one account.
   - `RunAtStartup=On` or `Off` — run script at startup or not.
3. Save and run the script.

## Project Structure

```text
ahk-gtd-email-automation/
├── GTD_Automation.ahk      # Main entry point (run this)
├── lib/
│   ├── GmailGTD.ahk        # Gmail automation module
│   └── OutlookGTD.ahk      # Outlook automation module
├── config.ini              # Your Outlook email configuration
├── config.example.ini      # Configuration template
├── CHANGELOG.md            # Version history
└── README.md               # This file
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License.

## Author

### Mehmet Can Türk

- GitHub: [@turkmehmetcan](https://github.com/turkmehmetcan)

## Acknowledgments

- Inspired by the [GTD (Getting Things Done)](https://gettingthingsdone.com/) methodology by David Allen
- Built with [AutoHotkey v2.0](https://www.autohotkey.com/)

## Support

If you encounter any issues or have questions:

- Open an [issue](https://github.com/turkmehmetcan/ahk-gtd-email-automation/issues)
- Check the [AutoHotkey documentation](https://www.autohotkey.com/docs/v2/)
