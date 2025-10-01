; AutoHotkey v2 script to trigger your existing Quick Steps
#HotIf WinActive("ahk_exe OUTLOOK.EXE")

; Replace CTRL+SHIFT+1 with Alt+E which triggers the quickstep "Archive" in Outlook
!e:: {
    Send "^+1"
}

; Replace CTRL+SHIFT+2 with Alt+A which triggers the quickstep "@Action" in Outlook
!a:: {
    Send "^+2"
}

; Replace CTRL+SHIFT+3 with Alt+W which triggers the quickstep "@Waiting For" in Outlook
!w:: {
    Send "^+3"
}

; Replace CTRL+SHIFT+4 with Alt+R which triggers the quickstep "@Reference" in Outlook
!r:: {
    Send "^+4"
}

; Replace CTRL+SHIFT+9 with Alt+Z which triggers the quickstep "Inbox" in Outlook
!z:: {
    Send "^+9"
}

#HotIf