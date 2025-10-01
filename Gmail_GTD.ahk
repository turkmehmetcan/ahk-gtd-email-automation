; Gmail GTD - AutoHotkey v2 Script

; ========== CONFIGURATION ==========
LABEL_ARCHIVE := "[0] GTD ARCHIVE"
LABEL_ACTION := "[1] @ACTION"
LABEL_WAITING := "[2] @WAITING FOR"
LABEL_REFERENCE := "[3] @REFERENCE"
LABEL_GMAIL_INBOX := "inbox"

DELAY_SHORT := 50     ; Short delay for key sequences
DELAY_LONG := 150     ; Long delay for UI operations

; ========== DEFAULT GMAIL(GM) KEYBOARD SHORTCUTS ==========
GM_LABEL := "l"             ; Open label menu
GM_ARCHIVE := "e"           ; Archive email
GM_MOVE := "v"              ; Move to folder/label
GM_GO_TO_INBOX := "gi"      ; Move to inbox
GM_MARK_UNREAD := "+u"      ; Mark as unread (Shift+U)
GM_MARK_READ := "+i"        ; Mark as read (Shift+I)
GM_SELECT_ALL := "^a"       ; Select all text (Ctrl+A)
GM_DELETE := "{Delete}"     ; Delete selected text
GM_ESCAPE := "{Escape}"     ; Close dialog/menu
GM_ENTER := "{Enter}"       ; Confirm action

; ========== SUPPORTED BROWSERS ==========
BROWSERS := [
    "firefox.exe", "chrome.exe", "msedge.exe", "opera.exe", "brave.exe",
    "vivaldi.exe", "waterfox.exe", "librewolf.exe", "tor.exe", "seamonkey.exe",
    "palemoon.exe", "basilisk.exe", "safari.exe", "yandex.exe", "whale.exe",
    "sidekick.exe", "arc.exe", "ghostbrowser.exe", "maxthon.exe", "cent.exe",
    "uc.exe", "slimjet.exe", "comodo.exe", "chromium.exe", "ungoogled-chromium.exe"
]

; ========== HOTKEYS ==========
#HotIf IsGmailActive()

!Enter:: ProcessGTDBucket(LABEL_ARCHIVE, false) ; Archive (read)
!a:: ProcessGTDBucket(LABEL_ACTION, true)       ; Action (unread)
!w:: ProcessGTDBucket(LABEL_WAITING, true)      ; Waiting (unread)
!r:: ProcessGTDBucket(LABEL_REFERENCE, true)    ; Reference (unread)

!e:: GmailArchive()     ; Archive email
!z:: MoveToInbox()      ; Remove labels & move to inbox
!u:: MarkUnread()       ; Mark as unread
!i:: MarkRead()         ; Mark as read

#HotIf

; ========== CORE FUNCTIONS ==========
ProcessGTDBucket(labelName, isUnread := true) {
    ApplyLabel(labelName)
    Sleep(DELAY_LONG * 3)

    ; Move to label folder (archive from inbox)
    Send(GM_ARCHIVE)
    Sleep(DELAY_LONG * 3)
    Send(GM_GO_TO_INBOX)  ; Go to inbox view to refresh
    Sleep(DELAY_LONG)

    if (isUnread)
        MarkUnread()
}

ApplyLabel(labelName) {
    Send(GM_LABEL)
    Sleep(DELAY_LONG)
    Send(labelName)
    Sleep(DELAY_LONG)
    Send(GM_ENTER)
}

GmailArchive() {
    Send(GM_ARCHIVE)
}

MoveToInbox() {
    ; Clear all labels
    ; Send(GM_LABEL)
    ; Sleep(DELAY_LONG)
    ; Send(GM_SELECT_ALL)
    ; Sleep(DELAY_SHORT)
    ; Send(GM_DELETE)
    ; Sleep(DELAY_SHORT)
    ; Send(GM_ESCAPE)
    ; Sleep(DELAY_LONG)

    ; MarkUnread()

    ; Move to inbox and mark unread
    Send(GM_MOVE)
    Sleep(DELAY_LONG)
    Send(LABEL_GMAIL_INBOX)
    Sleep(DELAY_LONG)
    Send(GM_ENTER)

}

MarkUnread() {
    Send(GM_MARK_UNREAD)
}

MarkRead() {
    Send(GM_MARK_READ)
}

; ========== BROWSER DETECTION ==========
IsGmailActive() {
    ; Check browser
    for browser in BROWSERS {
        if (WinActive("ahk_exe " . browser)) {
            ; Check if tab title contains Gmail indicators
            title := WinGetTitle("A")
            return (InStr(title, "Gmail") || InStr(title, "mail.google.com") || InStr(title, "Inbox"))
        }
    }
    return false
}
