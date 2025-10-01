; Outlook GTD – Always create tasks in Work A (AutoHotkey v2)

; Load email from config file
CONFIG_FILE := A_ScriptDir "\config.ini"
TARGET_SMTP := IniRead(CONFIG_FILE, "Settings", "TargetEmail", "")
if (TARGET_SMTP = "") {
    MsgBox "Please create config.ini with your email address.`n`nExample:`n[Settings]`nTargetEmail=your.email@domain.com"
    ExitApp
}

ACTION_CAT := "@Action"
WAIT_CAT := "@Waiting For"
OPEN_TASK := true                 ; set true to open each new task window

#HotIf WinActive("ahk_exe OUTLOOK.EXE")

!e:: {
    Send "^+1"
}	; Replace CTRL+SHIFT+1 with Alt+E which triggers the quickstep "Archive" in Outlook
!r:: {
    Send "^+4"
}	; Replace CTRL+SHIFT+4 with Alt+R which triggers the quickstep "@Reference" in Outlook
!z:: {
    Send "^+9"
}	; Replace CTRL+SHIFT+9 with Alt+Z which triggers the quickstep "Inbox" in Outlook

; Alt+A: run @Action QS, then create task in Work A
!a:: {
    Send "^+2"
    Sleep 200
    CreateTaskInPrimaryMail(ACTION_CAT)
}

; Alt+W: run @Waiting For QS, then create task in Work A
!w:: {
    Send "^+3"
    Sleep 200
    CreateTaskInPrimaryMail(WAIT_CAT)
}

; Alt+T: create a task in Work A without running a Quick Step
!t:: {
    CreateTaskInPrimaryMail(ACTION_CAT)
}

#HotIf

CreateTaskInPrimaryMail(category := "") {
    global TARGET_SMTP, OPEN_TASK
    try {
        ol := ComObjActive("Outlook.Application")
        exp := ol.ActiveExplorer
        if !exp
            return
        sel := exp.Selection
        if !sel || sel.Count = 0
            return

        ; Find Primary Account's store
        session := ol.Session
        accounts := session.Accounts
        store := 0
        loop accounts.Count {
            acc := accounts.Item(A_Index)
            if StrLower(acc.SmtpAddress) = StrLower(TARGET_SMTP) {
                store := acc.DeliveryStore
                break
            }
        }
        if !store {
            MsgBox "Mailbox for " TARGET_SMTP " not found."
            return
        }

        tFolder := store.GetDefaultFolder(13)  ; 13 = olFolderTasks

        ; Create one task per selected email
        loop sel.Count {
            itm := sel.Item(A_Index)
            if itm.Class != 43                  ; 43 = olMail
                continue

            task := tFolder.Items.Add(3)        ; 3 = olTaskItem
            task.Subject := itm.Subject
            if (category != "")
                task.Categories := category
            task.Body := itm.Body
            task.Attachments.Add(itm)           ; embed the original email
            if OPEN_TASK
                task.Display()
            else
                task.Save()
        }
    } catch as e {
        TrayTip "Task creation failed: " e.Message, "Outlook", 3000
    }
}
