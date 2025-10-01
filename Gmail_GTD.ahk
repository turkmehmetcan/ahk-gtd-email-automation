; Gmail GTD - AutoHotkey v2 Script

; ========== CONFIGURATION ==========
LABEL_ARCHIVE := "[0] @GTD ARCHIVE"
LABEL_ACTION := "[1] @ACTION"
LABEL_WAITING := "[2] @WAITING FOR"
LABEL_REFERENCE := "[3] @REFERENCE"
LABEL_GMAIL_INBOX := "inbox"
LABEL_GMAIL_SPAM := "spam"

DELAY_SHORT := 50     ; Short delay for key sequences
DELAY_LONG := 250     ; Long delay for UI operations

; ========== DEFAULT GMAIL(GM) KEYBOARD SHORTCUTS (GMS_*) ==========
GMS_LABEL := "l"            ; Open label menu
GMS_ARCHIVE := "e"          ; Archive email
GMS_MOVE := "v"             ; Move to folder/label
GMS_GO_TO_INBOX := "gi"     ; Go to inbox (refresh view without page reload)
GMS_MARK_UNREAD := "+u"     ; Mark as unread (Shift+U)
GMS_MARK_READ := "+i"       ; Mark as read (Shift+I)
GMS_DELETE := "+3"          ; Delete current email (Shift+3)
GMS_ENTER := "{Enter}"      ; Apply action (Enter)
GMS_TAB := "{Tab}"          ; Select next element (Tab)

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

!Enter:: ProcessGTDBucket(LABEL_ARCHIVE, false) ; Alt+Enter: GTD Archive (read)
!a:: ProcessGTDBucket(LABEL_ACTION, true)       ; Alt+a: Action (unread)
!w:: ProcessGTDBucket(LABEL_WAITING, true)      ; Alt+w: Waiting For (unread)
!r:: ProcessGTDBucket(LABEL_REFERENCE, true)    ; Alt+r: Reference (unread)

!Delete:: DeleteMail()  ; Alt+Delete: Delete email
!End:: MarkSpam()       ; Alt+End: Mark as spam
!i:: MarkRead()         ; Alt+i: Mark as read
!u:: MarkUnread()       ; Alt+u: Mark as unread
!e:: MoveToArchive()    ; Alt+e: Archive email
!z:: MoveToInbox()      ; Alt+z: Move back to inbox
!Space:: RefreshInbox() ; Alt+Space: Refresh inbox view

#HotIf

; ========== CORE FUNCTIONS ==========
ProcessGTDBucket(labelName, isUnread := true) {

    ApplyLabel(labelName)
    Sleep(DELAY_LONG)

    ; Move to label folder (archive from inbox)
    MoveToArchive()
    Sleep(DELAY_LONG)
    ; RefreshInbox()
    ; Sleep(DELAY_LONG)

    if (isUnread)
        MarkUnread()

}

ApplyLabel(labelName) {
    Send(GMS_LABEL)
    Sleep(DELAY_LONG)
    Send(labelName)
    Sleep(DELAY_LONG)
    Send(GMS_ENTER)
}

DeleteMail() {
    Send(GMS_DELETE)
}

MoveToInbox() {
    ; Mark unread and move to inbox
    MarkUnread()
    Sleep(DELAY_LONG)
    Send(GMS_MOVE)
    Sleep(DELAY_LONG)
    Send(LABEL_GMAIL_INBOX)
    Sleep(DELAY_LONG * 2)
    Send(GMS_ENTER)

}

MarkSpam() {
    Send(GMS_MOVE)
    Sleep(DELAY_LONG)
    Send(LABEL_GMAIL_SPAM)
    Sleep(DELAY_LONG * 2)
    Send(GMS_ENTER)
    Sleep(DELAY_LONG * 3)
    Send(GMS_TAB)
    Sleep(DELAY_LONG)
    Send(GMS_ENTER)
    Sleep(DELAY_LONG)
    MarkRead()
    Sleep(DELAY_LONG)
    MoveToArchive()
}

MarkRead() {
    Send(GMS_MARK_READ)
}

MarkUnread() {
    Send(GMS_MARK_UNREAD)
}

MoveToArchive() {
    Send(GMS_ARCHIVE)
}

RefreshInbox() {
    Send(GMS_GO_TO_INBOX)
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
