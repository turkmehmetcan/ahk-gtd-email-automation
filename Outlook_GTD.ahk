; Outlook GTD - AutoHotkey v2 Script
; Hotkeys: Alt+A(Action) Alt+W(Waiting) Alt+R(Reference) Alt+E(Archive) Alt+Z(Inbox) Alt+0(Setup)
; All folder/category names and colors are defined at the top for easy configuration.

; ========== CONFIGURATION ==========
; Load email from config file
CONFIG_FILE := A_ScriptDir "\config.ini"
PRIMARY_SMTP := IniRead(CONFIG_FILE, "Settings", "PrimaryEmail", "")
if (PRIMARY_SMTP = "") {
    MsgBox "Please create config.ini with your email address.`n`nExample:`n[Settings]`nPrimaryEmail=your.email@domain.com"
    ExitApp
}

OPEN_TASK := true

FOLDER_ACTION := "@ACTION"
FOLDER_WAITING := "@WAITING FOR"
FOLDER_REFERENCE := "@REFERENCE"

CATEGORY_ACTION := "@Action"
CATEGORY_WAITING := "@Waiting For"
CATEGORY_REFERENCE := "@Reference"

COLOR_ACTION := 1   ; Red
COLOR_WAITING := 3   ; Yellow
COLOR_REFERENCE := 4   ; Blue

; ========== HOTKEYS ==========
#HotIf WinActive("ahk_exe OUTLOOK.EXE")

!a:: ProcessGTDAction(FOLDER_ACTION, CATEGORY_ACTION, true, true)
!w:: ProcessGTDAction(FOLDER_WAITING, CATEGORY_WAITING, true, true)
!r:: ProcessGTDAction(FOLDER_REFERENCE, CATEGORY_REFERENCE, true, false)
!e:: ProcessArchive()
!z:: ProcessInbox()
!0:: SetupGTDFolders()

#HotIf

; ========== CORE HELPERS ==========
ProcessGTDAction(folderName, categoryName, markUnread, createTask := false) {
    try {
        ol := ComObjActive("Outlook.Application")
        exp := ol.ActiveExplorer
        if !exp || !exp.Selection || exp.Selection.Count = 0
            return

        items := []
        loop exp.Selection.Count {
            itm := exp.Selection.Item(A_Index)
            if itm && itm.Class = 43  ; olMail
                items.Push(itm)
        }

        for itm in items {
            store := itm.Parent.Store
            target := FindOrCreateFolder(store, folderName)
            ApplyCategory(itm, categoryName)
            itm.UnRead := markUnread
            itm.Save()
            moved := itm.Move(target)
            if createTask {
                CreateTaskInPrimary(moved, categoryName)
            }
        }
    } catch as e {
        TrayTip("GTD Action failed: " e.Message, "Outlook", 3000)
    }
}

ProcessArchive() {
    try {
        ol := ComObjActive("Outlook.Application")
        exp := ol.ActiveExplorer
        if !exp || !exp.Selection || exp.Selection.Count = 0
            return

        loop exp.Selection.Count {
            itm := exp.Selection.Item(A_Index)
            if itm && itm.Class = 43 {  ; olMail
                store := itm.Parent.Store
                try {
                    archive := store.GetDefaultFolder(32) ; olFolderArchive
                } catch {
                    archive := FindOrCreateFolder(store, "Archive")
                }
                ; Mark as read (per your Quick Step)
                itm.UnRead := false
                itm.Save()
                itm.Move(archive)
            }
        }
    } catch as e {
        TrayTip("Archive failed: " e.Message, "Outlook", 3000)
    }
}

ProcessInbox() {
    try {
        ol := ComObjActive("Outlook.Application")
        exp := ol.ActiveExplorer
        if !exp || !exp.Selection || exp.Selection.Count = 0
            return

        loop exp.Selection.Count {
            itm := exp.Selection.Item(A_Index)
            if itm && itm.Class = 43 {  ; olMail
                store := itm.Parent.Store

                inbox := store.GetDefaultFolder(6)  ; olFolderInbox

                ; Clear categories and mark unread (per your Quick Step)
                itm.Categories := ""
                itm.UnRead := true
                itm.Save()
                itm.Move(inbox)
            }
        }
    } catch as e {
        TrayTip("Move to Inbox failed: " e.Message, "Outlook", 3000)
    }
}

SetupGTDFolders() {
    try {
        ol := ComObjActive("Outlook.Application")
        accounts := ol.Session.Accounts

        ; GTD Categories with colors (olCategoryColor constants)
        categories := [
            [CATEGORY_ACTION, COLOR_ACTION],
            [CATEGORY_WAITING, COLOR_WAITING],
            [CATEGORY_REFERENCE, COLOR_REFERENCE]
        ]

        ; Setup categories first
        for pair in categories {
            SetupCategory(ol, pair[1], pair[2])
        }

        ; Setup folders for each account
        loop accounts.Count {
            acc := accounts.Item(A_Index)
            store := acc.DeliveryStore
            if !store
                continue
            FindOrCreateFolder(store, FOLDER_ACTION)
            FindOrCreateFolder(store, FOLDER_WAITING)
            FindOrCreateFolder(store, FOLDER_REFERENCE)
            try {
                store.GetDefaultFolder(32)
            } catch {
                FindOrCreateFolder(store, "Archive")
            }
        }

        MsgBox("GTD setup completed successfully!", "Setup", 0x40)

    } catch as e {
        TrayTip("Setup failed: " e.Message, "Outlook", 3000)
    }
}

CreateTaskInPrimary(mailItem, category := "") {
    global PRIMARY_SMTP, OPEN_TASK
    try {
        if !mailItem || mailItem.Class != 43
            return

        ol := ComObjActive("Outlook.Application")
        accounts := ol.Session.Accounts

        loop accounts.Count {
            acc := accounts.Item(A_Index)
            if StrLower(acc.SmtpAddress) == StrLower(PRIMARY_SMTP) {
                tFolder := acc.DeliveryStore.GetDefaultFolder(13)  ; olFolderTasks
                task := tFolder.Items.Add(3)  ; olTaskItem
                task.Subject := mailItem.Subject
                if category
                    task.Categories := category
                task.Body := mailItem.Body
                task.Attachments.Add(mailItem)

                if OPEN_TASK
                    task.Display()
                else
                    task.Save()
                return
            }
        }
        MsgBox("Mailbox for " PRIMARY_SMTP " not found.")
    } catch as e {
        TrayTip("Task creation failed: " e.Message, "Outlook", 3000)
    }
}

ApplyCategory(mailItem, category) {
    try {
        cats := Trim(mailItem.Categories)
        if !cats {
            mailItem.Categories := category
        } else {
            ; Check if category already exists
            for part in StrSplit(cats, ",") {
                if StrLower(Trim(part)) == StrLower(category)
                    return  ; already has this category
            }
            mailItem.Categories := cats ", " category
        }
        mailItem.Save()
    } catch as e {
        TrayTip("Apply category failed: " e.Message, "Outlook", 3000)
    }
}

FindOrCreateFolder(store, name) {
    ; First try to find existing folder
    folder := FindExistingFolder(store, name)
    if folder
        return folder

    ; Create folder at root level if not found
    try {
        root := store.GetRootFolder()
        newFolder := root.Folders.Add(name)
        return newFolder
    } catch as e {
        TrayTip("Failed to create folder '" name "': " e.Message, "Outlook", 3000)
        return 0
    }
}

FindExistingFolder(store, name) {
    ; Search top-level folders first
    try {
        root := store.GetRootFolder()
        loop root.Folders.Count {
            f := root.Folders.Item(A_Index)
            if StrLower(f.Name) == StrLower(name)
                return f
        }
        ; Recursive search if not found at top level
        return FindFolderRecursive(root, name)
    } catch {
        return 0
    }
}

FindFolderRecursive(folder, name) {
    try {
        if StrLower(folder.Name) == StrLower(name)
            return folder
        loop folder.Folders.Count {
            sub := folder.Folders.Item(A_Index)
            result := FindFolderRecursive(sub, name)
            if result
                return result
        }
    } catch {
        ; Silent fail for inaccessible folders
    }
    return 0
}

SetupCategory(ol, categoryName, colorIndex) {
    try {
        categories := ol.Session.Categories

        ; Check if category exists
        existingCat := 0
        loop categories.Count {
            cat := categories.Item(A_Index)
            if StrLower(cat.Name) == StrLower(categoryName) {
                existingCat := cat
                break
            }
        }

        ; Create or update category
        if existingCat {
            ; Update color if different
            if existingCat.Color != colorIndex {
                existingCat.Color := colorIndex
            }
        } else {
            ; Create new category
            categories.Add(categoryName, colorIndex)
        }
    } catch as e {
        TrayTip("Failed to setup category '" categoryName "': " e.Message, "Outlook", 3000)
    }
}
