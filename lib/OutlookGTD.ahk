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
                createTaskInPrimary(moved, categoryName)
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

        madeChange := false
        TrayTip("Setup", "Starting GTD setup...", 2000)

        ; Setup categories and folders for each account
        TrayTip("Setup", "Setting up categories and folders for all accounts...", 2000)
        loop accounts.Count {
            acc := accounts.Item(A_Index)
            TrayTip("Setup", "Account " A_Index ": " acc.DisplayName, 2000)
            store := acc.DeliveryStore
            if !store {
                TrayTip("Setup", "No DeliveryStore for account: " acc.DisplayName, 3000)
                continue
            }
            try {
                root := store.GetRootFolder()
                TrayTip("Setup", "Root for " acc.DisplayName ": " root.Name " (" root.DefaultItemType ")", 2000)
            } catch as e {
                TrayTip("Setup", "Failed to get root for " acc.DisplayName ": " e.Message, 3000)
                continue
            }
            ; Setup categories for this account's store
            for pair in categories {
                try {
                    if setupCategoryForStore(store, pair[1], pair[2]) {
                        TrayTip("Setup", "Created or updated category: " pair[1] " in " acc.DisplayName, 2000)
                        madeChange := true
                    } else {
                        TrayTip("Setup", "Category already exists: " pair[1] " in " acc.DisplayName, 1000)
                    }
                } catch as e {
                    TrayTip("Setup", "Failed to setup category " pair[1] " in " acc.DisplayName ": " e.Message, 3000)
                }
            }
            ; Setup folders for this account
            if findOrCreateFolder(store, FOLDER_ACTION, true) {
                TrayTip("Setup", "Created folder: " FOLDER_ACTION " in " acc.DisplayName, 2000)
                madeChange := true
            }
            if findOrCreateFolder(store, FOLDER_WAITING, true) {
                TrayTip("Setup", "Created folder: " FOLDER_WAITING " in " acc.DisplayName, 2000)
                madeChange := true
            }
            if findOrCreateFolder(store, FOLDER_REFERENCE, true) {
                TrayTip("Setup", "Created folder: " FOLDER_REFERENCE " in " acc.DisplayName, 2000)
                madeChange := true
            }
            try {
                store.GetDefaultFolder(OL_FOLDER_ARCHIVE)
            } catch {
                if findOrCreateFolder(store, "Archive", true) {
                    TrayTip("Setup", "Created folder: Archive in " acc.DisplayName, 2000)
                    madeChange := true
                }
            }
        }
        ; Create category in a specific store/account
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
                        return true ; color updated
                    }
                    return false ; already exists, no change
                } else {
                    categories.Add(categoryName, colorIndex)
                    return true ; new category created
                }
            } catch as e {
                TrayTip("Failed to setup category '" categoryName "' in store: " e.Message, "Outlook", 3000)
                return false
            }
        }

        if madeChange {
            MsgBox("GTD setup completed or updated folders/categories!", "Setup", 0x40)
        } else {
            MsgBox("GTD setup: Everything already exists, no changes made.", "Setup", 0x40)
        }

    } catch as e {
        TrayTip("Setup failed: " e.Message, "Outlook", 3000)
    }
}

createTaskInPrimary(mailItem, category := "") {
    global PRIMARY_SMTP, OPEN_TASK
    try {
        if !mailItem || mailItem.Class != OL_CLASS_MAIL
            return

        ol := ComObjActive("Outlook.Application")
        accounts := ol.Session.Accounts

        loop accounts.Count {
            acc := accounts.Item(A_Index)
            if StrLower(acc.SmtpAddress) == StrLower(PRIMARY_SMTP) {
                tFolder := acc.DeliveryStore.GetDefaultFolder(OL_FOLDER_TASKS)
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
                return
            }
        }
        MsgBox("Mailbox for " PRIMARY_SMTP " not found.")
    } catch as e {
        TrayTip("Task creation failed: " e.Message, "Outlook", 3000)
    }
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
        if reportChange
            TrayTip("Setup", "Folder exists: " name, 1000)
        return reportChange ? false : folder
    }
    try {
        root := store.GetRootFolder()
        TrayTip("Setup", "Attempting to create folder '" name "' under root: " root.Name, 2000)
        newFolder := root.Folders.Add(name)
        TrayTip("Setup", "Created folder: " name, 2000)
        return reportChange ? true : newFolder
    } catch as e {
        TrayTip("Failed to create folder '" name "' under root: " root.Name ". Error: " e.Message, "Outlook", 3000)
        return reportChange ? false : 0
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
