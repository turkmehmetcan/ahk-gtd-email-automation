#Requires AutoHotkey v2.0

/*
================================================================================
Outlook GTD Module
================================================================================
Library module for Outlook GTD automation.
Provides GTD workflow functions for Outlook using COM API.
================================================================================
*/

; ========== CONFIGURATION ==========

; Load email from config file (config.ini is auto-created by main script if missing)
CONFIG_FILE := A_ScriptDir "\config.ini"
PRIMARY_SMTP := IniRead(CONFIG_FILE, "Settings", "PrimaryEmail", "Off")

OPEN_TASK := true

FOLDER_ACTION := "@ACTION"
FOLDER_WAITING := "@WAITING FOR"
FOLDER_REFERENCE := "@REFERENCE"

CATEGORY_ACTION := "@Action"
CATEGORY_WAITING := "@Waiting For"
CATEGORY_REFERENCE := "@Reference"

COLOR_ACTION := 1   ; Red
COLOR_WAITING := 4   ; Yellow
COLOR_REFERENCE := 8   ; Blue

; ========== OUTLOOK CONSTANTS ==========
OL_CLASS_MAIL := 43
OL_FOLDER_INBOX := 6
OL_FOLDER_TASKS := 13
OL_FOLDER_ARCHIVE := 32
OL_ITEM_TASK := 3

; ========== DETECTION ==========

IsOutlookActive() {
    return WinActive("ahk_exe OUTLOOK.EXE")
}

; ========== CORE HELPERS ==========

OutlookMoveToGtdBucket(folderName, categoryName, markUnread, createTask := false) {
    try {
        ol := ComObjActive("Outlook.Application")
        exp := ol.ActiveExplorer
        if !exp || !exp.Selection || exp.Selection.Count = 0
            return

        items := []
        loop exp.Selection.Count {
            itm := exp.Selection.Item(A_Index)
            if itm && itm.Class = OL_CLASS_MAIL
                items.Push(itm)
        }

        for itm in items {
            store := itm.Parent.Store
            target := findOrCreateFolder(store, folderName)
            applyCategory(itm, categoryName)
            itm.UnRead := markUnread
            itm.Save()
            moved := itm.Move(target)
            if createTask {
                createTaskFromEmail(moved, categoryName)
            }
        }
    } catch as e {
        TrayTip("GTD Action failed: " e.Message, "Outlook", 3000)
    }
}

OutlookMoveToArchive() {
    try {
        ol := ComObjActive("Outlook.Application")
        exp := ol.ActiveExplorer
        if !exp || !exp.Selection || exp.Selection.Count = 0
            return

        loop exp.Selection.Count {
            itm := exp.Selection.Item(A_Index)
            if itm && itm.Class = OL_CLASS_MAIL {
                store := itm.Parent.Store
                try {
                    archive := store.GetDefaultFolder(OL_FOLDER_ARCHIVE)
                } catch {
                    archive := findOrCreateFolder(store, "Archive")
                }
                itm.UnRead := false
                itm.Save()
                itm.Move(archive)
            }
        }
    } catch as e {
        TrayTip("Archive failed: " e.Message, "Outlook", 3000)
    }
}

OutlookMoveToInbox() {
    try {
        ol := ComObjActive("Outlook.Application")
        exp := ol.ActiveExplorer
        if !exp || !exp.Selection || exp.Selection.Count = 0
            return

        loop exp.Selection.Count {
            itm := exp.Selection.Item(A_Index)
            if itm && itm.Class = OL_CLASS_MAIL {
                store := itm.Parent.Store
                inbox := store.GetDefaultFolder(OL_FOLDER_INBOX)
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

OutlookCreateGtdElements() {
    try {
        ol := ComObjActive("Outlook.Application")
        accounts := ol.Session.Accounts

        ; GTD Categories with colors (olCategoryColor constants)
        categories := [
            [CATEGORY_ACTION, COLOR_ACTION],
            [CATEGORY_WAITING, COLOR_WAITING],
            [CATEGORY_REFERENCE, COLOR_REFERENCE]
        ]

        ; Track changes for summary
        accountsProcessed := 0
        categoriesCreated := 0
        categoriesUpdated := 0
        foldersCreated := 0
        errors := []

        ; Setup categories and folders for each account
        loop accounts.Count {
            acc := accounts.Item(A_Index)
            store := acc.DeliveryStore
            if !store {
                errors.Push("No DeliveryStore for account: " acc.DisplayName)
                continue
            }

            try {
                root := store.GetRootFolder()
            } catch as e {
                errors.Push("Failed to get root for " acc.DisplayName ": " e.Message)
                continue
            }

            accountsProcessed++

            ; Setup categories for this account's store
            for pair in categories {
                try {
                    result := setupCategoryForStore(store, pair[1], pair[2])
                    if result = 1
                        categoriesCreated++
                    else if result = 2
                        categoriesUpdated++
                } catch as e {
                    errors.Push("Failed to setup category " pair[1] " in " acc.DisplayName ": " e.Message)
                }
            }

            ; Setup folders for this account
            if findOrCreateFolder(store, FOLDER_ACTION, true)
                foldersCreated++
            if findOrCreateFolder(store, FOLDER_WAITING, true)
                foldersCreated++
            if findOrCreateFolder(store, FOLDER_REFERENCE, true)
                foldersCreated++

            try {
                store.GetDefaultFolder(OL_FOLDER_ARCHIVE)
            } catch {
                if findOrCreateFolder(store, "Archive", true)
                    foldersCreated++
            }
        }

        ; Build summary message
        summary := "GTD Setup Complete`n`n"
        summary .= "Accounts processed: " accountsProcessed "`n"
        summary .= "Categories created: " categoriesCreated "`n"
        summary .= "Categories updated: " categoriesUpdated "`n"
        summary .= "Folders created: " foldersCreated

        if errors.Length > 0 {
            summary .= "`n`nErrors encountered: " errors.Length
            if errors.Length <= 3 {
                summary .= "`n"
                for err in errors
                    summary .= "`n- " err
            }
        }

        ; Show summary
        if categoriesCreated > 0 || categoriesUpdated > 0 || foldersCreated > 0 {
            MsgBox(summary, "GTD Setup", 0x40)
        } else {
            MsgBox("GTD setup: Everything already exists, no changes made.`n`nAccounts processed: " accountsProcessed,
                "GTD Setup", 0x40)
        }

    } catch as e {
        MsgBox("Setup failed: " e.Message, "GTD Setup Error", 0x10)
    }
}

; Create category in a specific store/account
; Returns: 0 = already exists, 1 = newly created, 2 = color updated
setupCategoryForStore(store, categoryName, colorIndex) {
    try {
        categories := store.Application.Session.Categories
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
            if existingCat.Color != colorIndex {
                existingCat.Color := colorIndex
                return 2 ; color updated
            }
            return 0 ; already exists, no change
        } else {
            categories.Add(categoryName, colorIndex)
            return 1 ; new category created
        }
    } catch as e {
        throw Error("Failed to setup category '" categoryName "': " e.Message)
    }
}

createTaskFromEmail(mailItem, category := "") {
    global PRIMARY_SMTP, OPEN_TASK
    try {
        if !mailItem || mailItem.Class != OL_CLASS_MAIL
            return

        ol := ComObjActive("Outlook.Application")

        ; Determine target task folder based on PRIMARY_SMTP setting
        usePrimaryAccount := (StrLower(Trim(PRIMARY_SMTP)) != "off")

        if usePrimaryAccount {
            ; Create task in specified primary email account
            tFolder := getTaskFolderForPrimaryAccount(ol, PRIMARY_SMTP)
            if !tFolder {
                MsgBox("Mailbox for " PRIMARY_SMTP " not found.")
                return
            }
        } else {
            ; Create task in the same account as the email
            tFolder := mailItem.Parent.Store.GetDefaultFolder(OL_FOLDER_TASKS)
        }

        ; Create and configure the task
        task := tFolder.Items.Add(OL_ITEM_TASK)
        task.Subject := mailItem.Subject
        if category
            task.Categories := category
        task.Body := mailItem.Body
        task.Attachments.Add(mailItem)

        if OPEN_TASK
            task.Display()
        else
            task.Save()

    } catch as e {
        TrayTip("Task creation failed: " e.Message, "Outlook", 3000)
    }
}

getTaskFolderForPrimaryAccount(ol, primaryEmail) {
    try {
        accounts := ol.Session.Accounts
        loop accounts.Count {
            acc := accounts.Item(A_Index)
            if StrLower(acc.SmtpAddress) == StrLower(primaryEmail) {
                return acc.DeliveryStore.GetDefaultFolder(OL_FOLDER_TASKS)
            }
        }
    }
    return 0
}

applyCategory(mailItem, category) {
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

findOrCreateFolder(store, name, reportChange := false) {
    folder := findExistingFolder(store, name)
    if folder {
        return reportChange ? false : folder
    }
    try {
        root := store.GetRootFolder()
        newFolder := root.Folders.Add(name)
        return reportChange ? true : newFolder
    } catch as e {
        throw Error("Failed to create folder '" name "': " e.Message)
    }
}

findExistingFolder(store, name) {
    try {
        root := store.GetRootFolder()
        loop root.Folders.Count {
            f := root.Folders.Item(A_Index)
            if StrLower(f.Name) == StrLower(name)
                return f
        }
        ; Skip recursive search to avoid performance issues
        return 0
    } catch {
        return 0
    }
}

findFolderRecursive(folder, name) {
    try {
        if StrLower(folder.Name) == StrLower(name)
            return folder
        loop folder.Folders.Count {
            sub := folder.Folders.Item(A_Index)
            result := findFolderRecursive(sub, name)
            if result
                return result
        }
    } catch {
        ; Silent fail for inaccessible folders
    }
    return 0
}

setupCategory(ol, categoryName, colorIndex) {
    try {
        TrayTip("Debug", "Accessing categories for: " categoryName, 1500)
        categories := ol.Session.Categories

        TrayTip("Debug", "Categories count: " categories.Count, 1000)

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
            if existingCat.Color != colorIndex {
                existingCat.Color := colorIndex
                TrayTip("Setup", "Updated category color: " categoryName, 2000)
                return true ; color updated
            }
            TrayTip("Setup", "Category exists: " categoryName, 1000)
            return false ; already exists, no change
        } else {
            TrayTip("Debug", "Creating new category: " categoryName " with color: " colorIndex, 2000)
            categories.Add(categoryName, colorIndex)
            TrayTip("Setup", "Created category: " categoryName, 2000)
            return true ; new category created
        }
    } catch as e {
        TrayTip("Failed to setup category '" categoryName "': " e.Message, "Outlook", 3000)
        return false
    }
}
