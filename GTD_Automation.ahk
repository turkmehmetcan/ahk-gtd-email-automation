#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode("Input")

/*
================================================================================
GTD Email Automation - Main Wrapper
================================================================================

Description:
Unified GTD (Getting Things Done) automation for both Gmail and Outlook.
Automatically loads the appropriate module based on active application.

Features:
- Seamless switching between Gmail and Outlook workflows
- Modular architecture with separate libraries for each platform
- Consistent Alt+key hotkeys across both applications
- Easy to extend and maintain

Supported Applications:
- Gmail (web interface in 25+ browsers)
- Outlook (desktop application)

Hotkeys (Common across both platforms):
Alt+Enter - Move to GTD Archive
Alt+A     - Move to Action folder/label
Alt+W     - Move to Waiting For folder/label
Alt+R     - Move to Reference folder/label
Alt+E     - Archive email
Alt+Z     - Move back to inbox
Alt+Up    - Move to next (newer) email
Alt+Down  - Move to previous (older) email
Alt+I     - Mark as read (Gmail only)
Alt+U     - Mark as unread (Gmail only)
Alt+Delete- Delete email (Gmail only)
Alt+End   - Mark as spam (Gmail only)
Alt+Space - Refresh inbox (Gmail only)
Alt+0     - Setup GTD folders/categories (Outlook only)

Requirements:
- AutoHotkey v2.0+
- Gmail: Keyboard shortcuts enabled in Gmail settings
- Outlook: Desktop application installed
- config.ini file (auto-generated on first run with default settings)

Author: Mehmet Can Türk
Version: 2.0
License: MIT
================================================================================
*/

; ========== LOAD MODULES ==========
#Include lib\GmailGTD.ahk
#Include lib\OutlookGTD.ahk

; ========== CONFIG FILE CREATION ==========
CONFIG_FILE := A_ScriptDir "\config.ini"

; Create config.ini with default values if it doesn't exist
if !FileExist(CONFIG_FILE) {
    defaultConfig := "
    (
[Settings]
; PrimaryEmail=Off --> Create tasks in each email's own Tasks/ToDo list
; PrimaryEmail=your.email@domain.com --> Create all tasks in a single email's Tasks/ToDo list
PrimaryEmail=Off

; RunAtStartup=On --> Run script at Windows startup
; RunAtStartup=Off --> Do not run script at Windows startup
RunAtStartup=On
    )"

    try {
        FileAppend(defaultConfig, CONFIG_FILE)
    } catch as err {
        MsgBox "Failed to create config.ini: " err.Message
        ExitApp
    }
}

; ========== STARTUP SHORTCUT LOGIC ==========
RUN_AT_STARTUP := IniRead(CONFIG_FILE, "Settings", "RunAtStartup", "Off")
STARTUP_FOLDER := A_AppData "\Microsoft\Windows\Start Menu\Programs\Startup"
SCRIPT_PATH := A_ScriptFullPath
SHORTCUT_PATH := STARTUP_FOLDER "\GTD_Automation.lnk"

if (RUN_AT_STARTUP = "On") {
    if !FileExist(SHORTCUT_PATH) {
        try FileCreateShortcut(SCRIPT_PATH, SHORTCUT_PATH)
    }
} else {
    if FileExist(SHORTCUT_PATH) {
        try FileDelete(SHORTCUT_PATH)
    }
}

; ========== GMAIL HOTKEYS ==========
#HotIf IsGmailActive()

; GTD Workflow Hotkeys
!Enter:: MoveToGtdBucket(LABEL_GTD_ARCHIVE) ; Alt+Enter: Move to GTD Archive
!a:: MoveToGtdBucket(LABEL_GTD_ACTION)       ; Alt+a: Move to GTD Action
!w:: MoveToGtdBucket(LABEL_GTD_WAITING)      ; Alt+w: Move to GTD Waiting For
!r:: MoveToGtdBucket(LABEL_GTD_REFERENCE)    ; Alt+r: Move to GTD Reference

; Email Management
!e:: MoveToArchive()     ; Alt+E: Archive email
!Delete:: DeleteMail()   ; Alt+Delete: Delete email
!End:: MarkSpam()        ; Alt+End: Mark as spam
!z:: MoveToInbox()       ; Alt+Z: Move back to inbox

; Email Status
!i:: MarkRead()          ; Alt+I: Mark as read
!u:: MarkUnread()        ; Alt+U: Mark as unread

; Navigation
!Up:: MoveToNewerEmail()    ; Alt+Up: Move to next (newer) email
!Down:: MoveToOlderEmail() ; Alt+Down: Move to previous (older) email
!Space:: RefreshInbox()  ; Alt+Space: Refresh inbox without reloading page

#HotIf

; ========== OUTLOOK HOTKEYS ==========
#HotIf IsOutlookActive()

; GTD Workflow Hotkeys
!a:: OutlookMoveToGtdBucket(FOLDER_ACTION, CATEGORY_ACTION, true, true)
!w:: OutlookMoveToGtdBucket(FOLDER_WAITING, CATEGORY_WAITING, true, true)
!r:: OutlookMoveToGtdBucket(FOLDER_REFERENCE, CATEGORY_REFERENCE, true, false)

; Navigation
!Up:: OutlookMoveToNewerEmail()    ; Alt+Up: Move to next (newer) email
!Down:: OutlookMoveToOlderEmail()  ; Alt+Down: Move to previous (older) email

; Email Management
!e:: OutlookMoveToArchive()     ; Alt+E: Archive email
!z:: OutlookMoveToInbox()       ; Alt+Z: Move back to inbox

; Setup
!0:: OutlookCreateGtdElements() ; Alt+0: Setup GTD folders and categories

#HotIf