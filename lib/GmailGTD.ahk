#Requires AutoHotkey v2.0

/*
================================================================================
Gmail GTD Module
================================================================================
Library module for Gmail GTD automation.
Provides GTD workflow functions for Gmail in any browser.
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
