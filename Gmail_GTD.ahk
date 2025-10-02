#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode("Input")

/*
================================================================================
Gmail GTD Automation Script
================================================================================

Description:
Automates Getting Things Done (GTD) workflow for Gmail using keyboard shortcuts.
Provides quick email processing with Alt+key combinations to categorize,
archive, and manage emails efficiently across multiple browsers.

Features:
- GTD bucket categorization (Archive, Action, Waiting, Reference)
- Quick archive, delete, and spam management
- Cross-browser compatibility (Chrome, Firefox, Edge, etc.)
- Configurable delays for reliable automation across different systems
- Modular functions for easy customization and extension

Usage:
Alt+Enter - Move to GTD Archive as read
Alt+A     - Move to GTD Action as unread
Alt+W     - Move to GTD Waiting For as unread
Alt+R     - Move to GTD Reference as unread
Alt+E     - Archive email
Alt+Delete- Delete email
Alt+End   - Mark as spam
Alt+I     - Mark as read
Alt+U     - Mark as unread
Alt+Z     - Move back to inbox
Alt+Space - Refresh inbox without reloading page

Requirements:
- AutoHotkey v2.0+
- Gmail keyboard shortcuts enabled
- Supported browser with Gmail open

Author: Mehmet Can Türk
Version: 1.2
License: MIT
================================================================================
*/

; ========== CONFIGURATION ==========

; Timing Configuration (milliseconds)
DELAY_SHORT := 200     ; Quick operations (mark read/unread)
DELAY_MEDIUM := 400    ; Standard operations (label dialogs)
DELAY_LONG := 1600      ; Complex operations (archive/move)

; Gmail Default Labels
LABEL_GMAIL_INBOX := "inbox"
LABEL_GMAIL_SPAM := "spam"

; GTD Bucket Labels
LABEL_GTD_ARCHIVE := "[0] @GTD ARCHIVE"
LABEL_GTD_ACTION := "[1] @ACTION"
LABEL_GTD_WAITING := "[2] @WAITING FOR"
LABEL_GTD_REFERENCE := "[3] @REFERENCE"

; ========== GMAIL KEYBOARD SHORTCUTS ==========

; Email Actions
KEY_GMAIL_ARCHIVE := "e"            ; Archive current email
KEY_GMAIL_DELETE := "+3"            ; Delete current email (Shift+3)
KEY_GMAIL_LABEL := "l"              ; Open label menu
KEY_GMAIL_MOVE := "v"               ; Move to folder/label

; Email Status
KEY_GMAIL_MARK_READ := "+i"         ; Mark as read (Shift+I)
KEY_GMAIL_MARK_UNREAD := "+u"       ; Mark as unread (Shift+U)

; Navigation
KEY_GMAIL_GO_TO_INBOX := "gi"       ; Go to inbox
KEY_GMAIL_NEWER_CONVERSATION := "k" ; Next conversation

; UI Elements
KEY_GMAIL_ENTER := "{Enter}"        ; Confirm action
KEY_GMAIL_TAB := "{Tab}"            ; Navigate elements

; ========== SUPPORTED BROWSERS ==========
BROWSERS := [
    "chrome.exe", "firefox.exe", "msedge.exe", "opera.exe", "brave.exe",
    "vivaldi.exe", "arc.exe", "safari.exe", "yandex.exe", "whale.exe",
    "waterfox.exe", "librewolf.exe", "palemoon.exe", "basilisk.exe",
    "tor.exe", "seamonkey.exe", "ghostbrowser.exe", "maxthon.exe",
    "sidekick.exe", "cent.exe", "uc.exe", "slimjet.exe", "comodo.exe",
    "chromium.exe", "ungoogled-chromium.exe"
]

; ========== BROWSER DETECTION ==========

; Returns true if Gmail is active in any supported browser
IsGmailActive() {
    for browser in BROWSERS {
        if (WinActive("ahk_exe " . browser)) {
            title := WinGetTitle("A")
            return (InStr(title, "Gmail") || InStr(title, "mail.google.com") || InStr(title, "Inbox"))
        }
    }
    return false
}

; ========== HOTKEYS ==========

#HotIf IsGmailActive()

; GTD Workflow Hotkeys
!Enter:: MoveToGtdBucket(LABEL_GTD_ARCHIVE, false) ; Alt+Enter: Move to GTD Archive as read
!a:: MoveToGtdBucket(LABEL_GTD_ACTION, true)       ; Alt+a: Move to GTD Action as unread
!w:: MoveToGtdBucket(LABEL_GTD_WAITING, true)      ; Alt+w: Move to GTD Waiting For as unread
!r:: MoveToGtdBucket(LABEL_GTD_REFERENCE, true)    ; Alt+r: Move to GTD Reference as unread

; Email Management
!e:: MoveToArchive()     ; Alt+E: Archive email
!Delete:: DeleteMail()   ; Alt+Delete: Delete email
!End:: MarkSpam()        ; Alt+End: Mark as spam
!z:: MoveToInbox()       ; Alt+Z: Move back to inbox

; Email Status
!i:: MarkRead()          ; Alt+I: Mark as read
!u:: MarkUnread()        ; Alt+U: Mark as unread

; Navigation
!Space:: RefreshInbox()  ; Alt+Space: Refresh inbox without reloading page

#HotIf

; ========== SUPPORT FUNCTIONS ==========

; Waits for specified duration in milliseconds
WaitDelay(delayMs := DELAY_LONG) {
    Sleep(delayMs)
}

; Sends keys and waits for specified delay
SendShortcut(keys, delayMs := DELAY_LONG) {
    Send(keys)
    WaitDelay(delayMs)
}

; Sends text and waits for specified delay
SendText(text, delayMs := DELAY_LONG) {
    Send(text)
    WaitDelay(delayMs)
}

; Sends Enter key and waits for specified delay
PressEnter(delayMs := DELAY_LONG) {
    Send(KEY_GMAIL_ENTER)
    WaitDelay(delayMs)
}

; ========== CORE FUNCTIONS ==========

; Main GTD workflow - applies label, archives, and optionally marks unread
MoveToGtdBucket(labelName, markAsUnread := true) {
    SendShortcut(KEY_GMAIL_LABEL, DELAY_MEDIUM)
    SendText(labelName, DELAY_LONG)
    PressEnter(DELAY_MEDIUM)
    ; SendShortcut(KEY_GMAIL_MOVE, DELAY_MEDIUM)
    ; SendText(labelName, DELAY_LONG)
    ; PressEnter(DELAY_MEDIUM)
    MoveToArchive()
    if (markAsUnread)
        MarkUnread()
}

; Archives current email
MoveToArchive() {
    SendShortcut(KEY_GMAIL_ARCHIVE, DELAY_MEDIUM)
}

; Deletes current email
DeleteMail() {
    SendShortcut(KEY_GMAIL_DELETE, DELAY_MEDIUM)
}

; Marks email as spam and moves to next conversation
MarkSpam() {
    SendShortcut(KEY_GMAIL_MOVE, DELAY_MEDIUM)
    SendText(LABEL_GMAIL_SPAM, DELAY_LONG)
    PressEnter(DELAY_LONG * 2)
    SendShortcut(KEY_GMAIL_TAB, DELAY_MEDIUM)
    PressEnter(DELAY_MEDIUM)
    MarkRead()
    SendShortcut(KEY_GMAIL_NEWER_CONVERSATION, DELAY_SHORT)
}

; Moves email back to inbox as unread
MoveToInbox() {
    SendShortcut(KEY_GMAIL_MOVE, DELAY_LONG)
    SendText(LABEL_GMAIL_INBOX, DELAY_LONG)
    PressEnter(DELAY_LONG)
    MarkUnread()
    ; SendShortcut(KEY_GMAIL_NEWER_CONVERSATION, DELAY_SHORT)
}

; Marks current email as read
MarkRead() {
    SendShortcut(KEY_GMAIL_MARK_READ, DELAY_LONG)
}

; Marks current email as unread
MarkUnread() {
    SendShortcut(KEY_GMAIL_MARK_UNREAD, DELAY_LONG)
}

; Refreshes Gmail inbox view
RefreshInbox() {
    SendShortcut(KEY_GMAIL_GO_TO_INBOX, DELAY_MEDIUM)
}
